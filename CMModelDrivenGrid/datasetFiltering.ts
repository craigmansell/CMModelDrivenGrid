import { FilterExpression, publicDatasetFilter } from './filtering';
import { appendFetchFilter, childGuid, childLogicalName } from './childRecords';
import { extraDateOperatorIds } from './dateConditions';

type Column = ComponentFramework.PropertyHelper.DataSetApi.Column;
export const FILTER_CANDIDATE_LIMIT = 5000;

/** The SDK enum is larger than the supported PCF dataset subset. Resolve these on the server. */
function needsServerCondition(condition: ComponentFramework.PropertyHelper.DataSetApi.ConditionExpression): boolean {
    const operator = Number(condition.conditionOperator);
    return [7,13].includes(operator) || extraDateOperatorIds.has(operator)
        || (operator === 6 && typeof condition.value === 'string' && condition.value.includes('['));
}

export function needsServerFilter(filter: FilterExpression | undefined): boolean {
    return !!filter && ((filter.conditions ?? []).some(needsServerCondition) || (filter.filters ?? []).some(needsServerFilter));
}

export function needsFilterCompilation(filter: FilterExpression | undefined): boolean {
    return needsServerFilter(filter);
}

/**
 * Replace unsupported predicates with TRUE to obtain a safe candidate superset.
 * TRUE under OR makes that whole group TRUE, never an accidentally narrower query.
 * Presence also requires server resolution: this host includes blanks in a
 * dataset EQ/NE disjunction even though the corresponding FetchXML excludes them.
 */
export function supportedFilterSuperset(filter: FilterExpression | undefined, columns: Column[]): FilterExpression | undefined {
    if (!filter) return undefined;
    if (!needsFilterCompilation(filter)) return publicDatasetFilter(filter);
    const parts: (FilterExpression | undefined)[] = [
        ...(filter.conditions ?? []).map((condition): FilterExpression | undefined => needsServerCondition(condition) ? undefined
            : publicDatasetFilter({filterOperator:0,conditions:[{...condition}]})),
        ...(filter.filters ?? []).map(child => supportedFilterSuperset(child,columns)),
    ];
    if (!parts.length || (filter.filterOperator === 1 && parts.some(part => !part))) return undefined;
    const retained = parts.filter((part): part is FilterExpression => !!part);
    if (!retained.length) return undefined;
    if (retained.length === 1) return retained[0];
    return {filterOperator:filter.filterOperator,conditions:[],filters:retained};
}

export function buildFilterResolutionFetch(view: string, entityName: string, idField: string, filter: FilterExpression, candidateIds: string[]): string {
    childLogicalName(entityName); childLogicalName(idField);
    const ids = [...new Set(candidateIds.map(childGuid))];
    if (!ids.length || ids.length > 200) throw new Error('Invalid filter-resolution batch.');
    const xml = new DOMParser().parseFromString(view,'text/xml');
    const fetch = xml.documentElement;
    const entity = fetch.querySelector(':scope > entity');
    if (xml.querySelector('parsererror') || fetch.tagName !== 'fetch' || entity?.getAttribute('name') !== entityName || fetch.getAttribute('aggregate') === 'true') {
        throw new Error('Advanced filters require a non-aggregate view of this table.');
    }
    for (const key of ['top','savedqueryid','paging-cookie','returntotalrecordcount']) fetch.removeAttribute(key);
    fetch.setAttribute('distinct','true'); fetch.setAttribute('count','200'); fetch.setAttribute('page','1');
    for (const node of Array.from(entity.querySelectorAll('attribute,all-attributes,order'))) node.remove();
    const attribute = xml.createElement('attribute'); attribute.setAttribute('name',idField); entity.appendChild(attribute);
    const order = xml.createElement('order'); order.setAttribute('attribute',idField); entity.appendChild(order);
    appendFetchFilter(xml,entity,filter);
    appendFetchFilter(xml,entity,{filterOperator:0,conditions:[{attributeName:idField,conditionOperator:8,value:ids}]});
    return new XMLSerializer().serializeToString(xml);
}

export async function resolveFilterIds(webAPI: ComponentFramework.WebApi, view: string, entity: string, idField: string, filter: FilterExpression, candidates: string[], assertCurrent: () => void): Promise<string[]> {
    const ids = [...new Set(candidates.map(childGuid))];
    if (ids.length > FILTER_CANDIDATE_LIMIT) throw new Error('Advanced filters can check up to 5,000 rows. Narrow another column filter first.');
    const matches = new Set<string>();
    for (let offset=0;offset<ids.length;offset+=200) {
        assertCurrent();
        const batch=ids.slice(offset,offset+200);
        const query=buildFilterResolutionFetch(view,entity,idField,filter,batch);
        const result=await webAPI.retrieveMultipleRecords(entity,`?fetchXml=${encodeURIComponent(query)}`,200);
        assertCurrent();
        // A distinct ID projection bounded to 200 IDs must be complete in one page.
        if (result.nextLink) throw new Error('The text-filter check was incomplete. No partial filter has been applied.');
        for (const row of result.entities) {
            const id=childGuid(String(row[idField] ?? ''));
            if (!batch.includes(id)) throw new Error('The text-filter check returned a record outside this grid.');
            matches.add(id);
        }
    }
    return ids.filter(id => matches.has(id));
}

export function filterForIds(idField: string, ids: string[]): FilterExpression {
    childLogicalName(idField);
    if (!ids.length) return {filterOperator:0,conditions:[{attributeName:idField,conditionOperator:0,value:'00000000-0000-0000-0000-000000000000'}]};
    const filters: FilterExpression[] = [];
    for (let offset=0;offset<ids.length;offset+=200) filters.push({filterOperator:0,conditions:[{attributeName:idField,conditionOperator:8,value:ids.slice(offset,offset+200).map(childGuid)}]});
    return {filterOperator:1,conditions:[],filters};
}

interface CandidateHost {
    page: () => {ids: string[]; hasNextPage: boolean; error?: boolean};
    next: () => Promise<void>;
    assertCurrent: () => void;
}

/** Enumerate the public bound dataset, preserving its actual subgrid relationship scope. */
export async function collectFilterCandidates(host: CandidateHost): Promise<string[]> {
    const ids = new Set<string>();
    let previous = '';
    for (let pages=0;pages<500;pages++) {
        host.assertCurrent();
        const page=host.page();
        if (page.error) throw new Error('Power Apps could not load rows for this filter.');
        const signature=page.ids.join('|');
        if (pages && signature === previous) throw new Error('Power Apps repeated a page. No partial filter has been applied.');
        previous=signature;
        for (const id of page.ids) ids.add(childGuid(id));
        if (ids.size > FILTER_CANDIDATE_LIMIT) throw new Error('Advanced filters can check up to 5,000 rows. Narrow another column filter first.');
        if (!page.hasNextPage) return [...ids];
        await host.next();
    }
    throw new Error('The text-filter paging limit was reached. Narrow another column filter first.');
}
