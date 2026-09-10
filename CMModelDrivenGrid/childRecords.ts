import type { FilterExpression } from './filtering';
import { extraDateEntries } from './dateConditions';

type RecordApi = ComponentFramework.PropertyHelper.DataSetApi.EntityRecord;
type Sort = ComponentFramework.PropertyHelper.DataSetApi.SortStatus;
export interface ChildContinuation { query: string; nextPage: number; nextLink?: string; pagingCookie?: string }
export interface ChildFetchResult { entities: ComponentFramework.WebApi.Entity[]; nextLink?: string; pagingCookie?: string; moreRecords?: boolean }
export interface ChildRequest { parentId: string; page: number; sorting: Sort[]; filtering?: FilterExpression; continuation?: ChildContinuation }
export interface ChildPage { records: RecordApi[]; hasNextPage: boolean; continuation?: ChildContinuation }
export interface ChildRowsConfig {
    loadPage: (request: ChildRequest) => Promise<ChildPage>;
    /** Check the child view in batches without downloading every child row. */
    hasChildren?: (parentIds: string[]) => Promise<Record<string, boolean>>;
    revision?: number;
    readOnly?: boolean;
    onAdd?: (parentId: string) => Promise<boolean>;
    onDelete?: (parentId: string, record: RecordApi) => Promise<boolean>;
}

const operators: Record<number, string> = {
	...Object.fromEntries(extraDateEntries.map(([, definition]) => [definition[1], definition[2]])),
    0:'eq',1:'ne',2:'gt',3:'lt',4:'ge',5:'le',6:'like',7:'not-like',8:'in',12:'null',13:'not-null',14:'yesterday',15:'today',16:'tomorrow',
    17:'last-seven-days',18:'next-seven-days',19:'last-week',20:'this-week',22:'last-month',23:'this-month',
    25:'on',26:'on-or-before',27:'on-or-after',28:'last-year',29:'this-year',33:'last-x-days',34:'next-x-days',37:'last-x-months',38:'next-x-months',
};
function logicalName(value: string): string {
    if (!/^[a-zA-Z_][a-zA-Z0-9_]*$/.test(value)) throw new Error('Invalid child-grid field configuration.');
    return value;
}
function guid(value: string): string {
    const normalized = value.replace(/[{}]/g, '').toLowerCase();
    if (!/^[0-9a-f]{8}(-[0-9a-f]{4}){3}-[0-9a-f]{12}$/.test(normalized)) throw new Error('Invalid child-grid record or view ID.');
    return normalized;
}
export { guid as childGuid, logicalName as childLogicalName };
function parseXml(value: string): XMLDocument {
    const xml = new DOMParser().parseFromString(value, 'text/xml');
    if (xml.querySelector('parsererror')) throw new Error('The child view contains invalid FetchXML.');
    return xml;
}
export function appendFetchFilter(xml: XMLDocument, container: Element, filter: FilterExpression): void {
    const group = xml.createElement('filter');
    group.setAttribute('type', filter.filterOperator === 1 ? 'or' : 'and');
    for (const condition of filter.conditions ?? []) {
        const operator = operators[condition.conditionOperator];
        if (!operator) throw new Error('This filter operator is not supported in child rows.');
        const node = xml.createElement('condition');
        node.setAttribute('attribute', logicalName(condition.attributeName));
        node.setAttribute('operator', operator);
        if (condition.entityAliasName) node.setAttribute('entityname', logicalName(condition.entityAliasName));
        if (operator === 'in' || Array.isArray(condition.value)) {
            const values = Array.isArray(condition.value) ? condition.value : [condition.value];
            for (const value of values) { const child = xml.createElement('value'); child.textContent = String(value); node.appendChild(child); }
        } else if (condition.value !== undefined && condition.value !== null && condition.value !== '' && operator !== 'null' && operator !== 'not-null') {
            node.setAttribute('value', String(condition.value));
        }
        group.appendChild(node);
    }
    for (const child of filter.filters ?? []) appendFetchFilter(xml, group, child);
    container.appendChild(group);
}

/** Preserve the child view's filters; always AND the parent scope outside them. */
export function buildChildFetch(view: string, entityName: string, idField: string, parentField: string, request: ChildRequest): string {
    logicalName(entityName); logicalName(idField); logicalName(parentField);
    if (!Number.isInteger(request.page) || request.page < 1 || request.page > 1000) throw new Error('Child paging limit reached. Narrow the child filters.');
    const xml = parseXml(view);
    const fetch = xml.documentElement;
    const entity = fetch.querySelector(':scope > entity');
    if (fetch.tagName !== 'fetch' || entity?.getAttribute('name') !== entityName || fetch.getAttribute('aggregate') === 'true') throw new Error('Child view must be a non-aggregate view of the same table.');
    fetch.removeAttribute('top'); fetch.removeAttribute('savedqueryid'); fetch.removeAttribute('paging-cookie');
    fetch.setAttribute('count', '50'); fetch.setAttribute('page', String(request.page));
    if (!entity.querySelector(`:scope > attribute[name="${idField}"]`)) {
        const id = xml.createElement('attribute'); id.setAttribute('name', idField); entity.appendChild(id);
    }
    const scope = xml.createElement('filter'); scope.setAttribute('type','and');
    const condition = xml.createElement('condition'); condition.setAttribute('attribute',parentField); condition.setAttribute('operator','eq'); condition.setAttribute('value',guid(request.parentId)); scope.appendChild(condition); entity.appendChild(scope);
    if (request.filtering) appendFetchFilter(xml, entity, request.filtering);
    if (request.sorting.length) {
        for (const order of Array.from(entity.querySelectorAll(':scope > order'))) order.remove();
        for (const sort of request.sorting) { const order = xml.createElement('order'); order.setAttribute('attribute',logicalName(sort.name)); order.setAttribute('descending', String(sort.sortDirection === 1)); entity.appendChild(order); }
    }
    if (!entity.querySelector(`:scope > order[attribute="${idField}"]`)) { const order = xml.createElement('order'); order.setAttribute('attribute',idField); entity.appendChild(order); }
    return new XMLSerializer().serializeToString(xml);
}

export function childRecord(row: ComponentFramework.WebApi.Entity, entityName: string, idField: string, nameField: string): RecordApi {
    const id = guid(String(row[idField] ?? ''));
    const field = (name: string) => Object.prototype.hasOwnProperty.call(row, `_${name}_value`) ? `_${name}_value` : name;
    return {
        getRecordId: () => id,
        getNamedReference: () => ({ id: {guid:id}, etn:entityName, name:String(row[nameField] ?? '') }),
        getValue: (name: string) => {
            const key = field(name); const value: unknown = row[key];
            if (key !== name && typeof value === 'string' && value) return { id:{guid:value}, etn: String(row[`${key}@Microsoft.Dynamics.CRM.lookuplogicalname`] ?? ''), name:String(row[`${key}@OData.Community.Display.V1.FormattedValue`] ?? '') };
            return value ?? null;
        },
        getFormattedValue: (name: string) => String(row[`${field(name)}@OData.Community.Display.V1.FormattedValue`] ?? row[field(name)] ?? ''),
    } as unknown as RecordApi;
}

export function createChildLoader(webAPI: ComponentFramework.WebApi, entityName: string, idField: string, nameField: string, parentField: string, viewId: string,
    retrieve?: (fetchXml: string) => Promise<ChildFetchResult>): ChildRowsConfig['loadPage'] {
    let viewPromise: Promise<string> | undefined;
    return async (request) => {
        viewPromise ??= webAPI.retrieveRecord('savedquery', guid(viewId), '?$select=fetchxml,returnedtypecode').then((view) => {
            if (view.returnedtypecode !== entityName) throw new Error('The child view belongs to a different table.');
            return String(view.fetchxml);
        }).catch((error: unknown) => { viewPromise = undefined; throw error; });
        const view = await viewPromise;
        const query = buildChildFetch(view, entityName, idField, parentField, {...request,page:1});
        let fetchXml = buildChildFetch(view, entityName, idField, parentField, request);
        // The client API accepts raw FetchXML, and encodes it itself. Preserve a
        // server continuation verbatim when supplied, scoped to this exact query.
        if (request.continuation) {
            if (request.continuation.nextPage !== request.page || request.page === 1 || request.continuation.query !== query) throw new Error('Child paging belongs to a different filter, parent or page. Refresh the grid.');
            if (request.continuation.pagingCookie) {
                const xml = parseXml(fetchXml);
                const cookie = parseXml(request.continuation.pagingCookie).documentElement;
                if (cookie.tagName !== 'cookie') throw new Error('Invalid child paging cookie.');
                const encoded = cookie.getAttribute('pagingcookie');
                const inner = encoded ? decodeURIComponent(decodeURIComponent(encoded)) : request.continuation.pagingCookie;
                if (parseXml(inner).documentElement.tagName !== 'cookie') throw new Error('Invalid child paging cookie.');
                xml.documentElement.setAttribute('paging-cookie',inner);
                fetchXml = new XMLSerializer().serializeToString(xml);
            }
        }
        let options = `?fetchXml=${fetchXml}`;
        if (!retrieve && request.continuation?.nextLink) {
            const start = request.continuation.nextLink.indexOf('?');
            if (start < 0) throw new Error('The child paging link is invalid. Refresh the grid.');
            options = request.continuation.nextLink.slice(start);
        }
        const result: ChildFetchResult = retrieve ? await retrieve(fetchXml) : await webAPI.retrieveMultipleRecords(entityName, options, 50);
        return { records:result.entities.map(row => childRecord(row,entityName,idField,nameField)), hasNextPage:result.moreRecords ?? (!!result.nextLink || result.entities.length === 50),
            continuation: result.nextLink || result.pagingCookie ? {query,nextPage:request.page+1,nextLink:result.nextLink,pagingCookie:result.pagingCookie} : undefined };
    };
}

/** One distinct parent projection, retaining all child-view security/filter joins. */
export function buildChildAvailabilityFetch(view: string, entityName: string, parentField: string, parentIds: string[]): string {
    logicalName(entityName); logicalName(parentField);
    const ids = [...new Set(parentIds.map(guid))];
    if (!ids.length || ids.length > 50) throw new Error('Check between 1 and 50 parent rows at a time.');
    const xml = parseXml(view);
    const fetch = xml.documentElement;
    const entity = fetch.querySelector(':scope > entity');
    if (fetch.tagName !== 'fetch' || entity?.getAttribute('name') !== entityName || fetch.getAttribute('aggregate') === 'true') throw new Error('Child view must be a non-aggregate view of the same table.');
    for (const name of ['top','savedqueryid','paging-cookie','returntotalrecordcount']) fetch.removeAttribute(name);
    fetch.setAttribute('distinct','true'); fetch.setAttribute('count','50'); fetch.setAttribute('page','1');
    for (const node of Array.from(entity.querySelectorAll('attribute,all-attributes,order'))) node.remove();
    const attribute = xml.createElement('attribute'); attribute.setAttribute('name',parentField); attribute.setAttribute('alias','drivengrid_parent'); entity.appendChild(attribute);
    appendFetchFilter(xml,entity,{filterOperator:0,conditions:[{attributeName:parentField,conditionOperator:8,value:ids}]});
    return new XMLSerializer().serializeToString(xml);
}

export function createChildAvailabilityLoader(webAPI: ComponentFramework.WebApi, entityName: string, parentField: string, viewId: string): NonNullable<ChildRowsConfig['hasChildren']> {
    let viewPromise: Promise<string> | undefined;
    return async parentIds => {
        const ids = [...new Set(parentIds.map(guid))];
        if (!ids.length) return {};
        viewPromise ??= webAPI.retrieveRecord('savedquery',guid(viewId),'?$select=fetchxml,returnedtypecode').then(view => {
            if (view.returnedtypecode !== entityName) throw new Error('The child view belongs to a different table.');
            return String(view.fetchxml);
        }).catch((error: unknown) => {viewPromise=undefined;throw error;});
        const found = Object.fromEntries(ids.map(id => [id,false]));
        const view = await viewPromise;
        for (let offset=0;offset<ids.length;offset+=50) {
            const batch=ids.slice(offset,offset+50);
            const query=buildChildAvailabilityFetch(view,entityName,parentField,batch);
            const result=await webAPI.retrieveMultipleRecords(entityName,`?fetchXml=${encodeURIComponent(query)}`,50);
            for (const row of result.entities) {
                const id=guid(String(row.drivengrid_parent ?? row[`_${parentField}_value`] ?? row[parentField] ?? ''));
                if (!batch.includes(id)) throw new Error('The child check returned a parent outside the requested rows.');
                found[id]=true;
            }
            if (result.nextLink && batch.some(id => !found[id])) throw new Error('The child check was incomplete. Refresh the subgrid.');
        }
        return found;
    };
}
