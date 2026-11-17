import {r200sSnapshot} from './r200s-demo-data';

export interface SetupComponent {id: string; code: string; name: string; available: boolean; isDefault?: boolean}
export interface SetupGroup {id: string; name: string; order: number; hidden?: boolean; freeText?: boolean; components: SetupComponent[]}
export interface Catalogue {
    version: 1; blindType: {id: number; code: string; name: string}; groups: SetupGroup[];
    legacyResults?: {conditionId: string; resultId: string; component: SetupComponent; groupId: string}[];
}
export interface OptionSetup {groupId: string; name?: string; sourceGroupId?: string; order: number; enabled: boolean; hidden: boolean; freeText: boolean; componentIds: string[]; defaultId: string | null}
export interface SetupResult {id: string; componentId: string; order: number}
export interface SetupCondition {
    id: string; order: number; mode: 'all' | 'any'; tests: {groupId: string; componentId: string}[];
    minWidth: number; maxWidth: number | null; minDrop: number; maxDrop: number | null;
    maxArea: number; minChainLength: number; maxChainLength: number | null;
    minPanels: number; maxPanels: number | null; formula: string | null; results: SetupResult[];
}
export interface LinkedSetup {id: string; groupId: string; order: number; conditions: SetupCondition[]}
export interface SetupDocument {version: 1 | 2; blindTypeCode: string; options: OptionSetup[]; links: LinkedSetup[]}
export interface SetupRow {
    id: string; parentId?: string; name: string; kind: 'Option' | 'Component' | 'Linked option' | 'Condition' | 'Result';
    code: string; order: number; details: string; count: number | null; groupId: string;
    linkId?: string; conditionId?: string; resultId?: string; componentId?: string;
}
export const clone = <T,>(value: T): T => JSON.parse(JSON.stringify(value)) as T;
export const ordered = <T extends {order: number}>(values: T[]): T[] => [...values].sort((a, b) => a.order - b.order);
export function groupFor(catalogue: Catalogue, id: string): SetupGroup {
    const group = catalogue.groups.find(item => item.id === id);
    if (!group) throw new Error('This option is not available for the selected blind type.');
    return group;
}
export const isCustomOption = (option: OptionSetup): boolean => option.groupId.startsWith('custom-');
export function optionGroup(catalogue: Catalogue, document: SetupDocument, id: string): SetupGroup {
    const option = document.options.find(item => item.groupId === id);
    if (!option) throw new Error('This option is not configured for the selected blind type.');
    const source = groupFor(catalogue, option.sourceGroupId ?? option.groupId);
    return {...source, id: option.groupId, name: option.name ?? source.name, order: option.order};
}
export function withOption(catalogue: Catalogue, document: SetupDocument, option: OptionSetup): SetupDocument {
    const next = clone(document);
    const value = clone(option);
    if (isCustomOption(value)) {value.name = value.name?.trim(); next.version = 2;}
    const index = next.options.findIndex(item => item.groupId === value.groupId);
    if (index < 0) next.options.push(value); else next.options[index] = value;
    validateSetup(catalogue, next);
    return next;
}
export function withoutOption(catalogue: Catalogue, document: SetupDocument, groupId: string): SetupDocument {
    const option = document.options.find(item => item.groupId === groupId);
    if (!option || !isCustomOption(option)) throw new Error('Catalogue options can be disabled, rather than removed.');
    if (document.links.some(link => link.groupId === groupId || link.conditions.some(condition => condition.tests.some(test => test.groupId === groupId)))) {
        throw new Error('This option is used by linked rules. Remove those references before removing the option.');
    }
    const next = clone(document);
    next.options = next.options.filter(item => item.groupId !== groupId);
    if (!next.options.some(isCustomOption)) next.version = 1;
    validateSetup(catalogue, next);
    return next;
}
export function availableComponents(catalogue: Catalogue, document: SetupDocument, groupId: string): SetupComponent[] {
    const option = document.options.find(item => item.groupId === groupId);
    if (!option?.enabled) return [];
    return optionGroup(catalogue, document, groupId).components.filter(component => component.available && option.componentIds.includes(component.id));
}
export function componentChoices(catalogue: Catalogue, document: SetupDocument, groupId: string, replacingId?: string): SetupComponent[] {
    const option = document.options.find(item => item.groupId === groupId);
    if (!option?.enabled) return [];
    return optionGroup(catalogue, document, groupId).components.filter(component => component.available && (component.id === replacingId || !option.componentIds.includes(component.id)));
}
export function withComponent(catalogue: Catalogue, document: SetupDocument, groupId: string, componentId: string, makeDefault: boolean, replacingId?: string): SetupDocument {
    if (!componentChoices(catalogue, document, groupId, replacingId).some(item => item.id === componentId)) throw new Error('Choose an available component that is not already in this option.');
    const next = clone(document); const option = next.options.find(item => item.groupId === groupId)!;
    if (replacingId && !option.componentIds.includes(replacingId)) throw new Error('The component row is no longer available.');
    option.componentIds = replacingId ? option.componentIds.map(id => id === replacingId ? componentId : id) : [...option.componentIds, componentId];
    if (makeDefault) option.defaultId = componentId;
    else if (option.defaultId === replacingId) option.defaultId = null;
    validateSetup(catalogue, next);
    return next;
}
export function withoutSetupRow(catalogue: Catalogue, document: SetupDocument, row: SetupRow): SetupDocument {
    const next = clone(document);
    if (row.kind === 'Linked option') next.links = next.links.filter(item => item.id !== row.linkId);
    else if (row.kind === 'Condition') {const target = next.links.find(item => item.id === row.linkId)!; target.conditions = target.conditions.filter(item => item.id !== row.conditionId);}
    else if (row.kind === 'Result') {const target = next.links.find(item => item.id === row.linkId)!.conditions.find(item => item.id === row.conditionId)!; target.results = target.results.filter(item => item.id !== row.resultId);}
    else if (row.kind === 'Option') {const option = next.options.find(item => item.groupId === row.groupId)!; if (isCustomOption(option)) return withoutOption(catalogue, next, row.groupId); option.enabled = false;}
    else {const option = next.options.find(item => item.groupId === row.groupId)!; option.componentIds = option.componentIds.filter(item => item !== row.componentId); if (option.defaultId === row.componentId) option.defaultId = null;}
    validateSetup(catalogue, next);
    return next;
}
export function resultComponent(catalogue: Catalogue, document: SetupDocument, groupId: string, conditionId: string, result: SetupResult): SetupComponent {
    return optionGroup(catalogue, document, groupId).components.find(component => component.id === result.componentId)
        ?? catalogue.legacyResults?.find(item => item.groupId === groupId && item.conditionId === conditionId && item.resultId === result.id && item.component.id === result.componentId)?.component
        ?? {id: result.componentId, code: '', name: 'Unavailable component', available: false};
}
const upper = (value: number): number | null => value >= 9999999 ? null : value;
export function initialR200sSetup(): {catalogue: Catalogue; document: SetupDocument} {
    const catalogue = clone(require('./fixtures/r200s-component-catalogue.json') as Catalogue);
    const groupId = (name: string): string => catalogue.groups.find(group => group.name === name)!.id;
    const document: SetupDocument = {version: 1, blindTypeCode: catalogue.blindType.code,
        options: catalogue.groups.map(group => ({groupId: group.id, order: group.order, enabled: true,
            hidden: !!group.hidden, freeText: !!group.freeText, componentIds: group.components.map(component => component.id),
            defaultId: group.components.find(component => component.isDefault)?.id ?? null})),
        links: r200sSnapshot.linkedOptions.map(link => ({id: String(link.id), groupId: groupId(link.name), order: link.sequence,
            conditions: link.conditions.map(condition => {
                const extra = condition as typeof condition & {maxArea: number; minChainLength: number; maxChainLength: number; minPanels: number; maxPanels: number};
                return {id: String(condition.id), order: condition.sequence, mode: /OR/.test(condition.rule) ? 'any' : 'all',
                    tests: condition.choices.map(choice => ({groupId: groupId(choice.group), componentId: catalogue.groups.flatMap(group => group.components).find(component => component.code === choice.code)!.id})),
                    minWidth: condition.minWidth, maxWidth: upper(condition.maxWidth), minDrop: condition.minDrop, maxDrop: upper(condition.maxDrop),
                    maxArea: extra.maxArea, minChainLength: extra.minChainLength, maxChainLength: upper(extra.maxChainLength),
                    minPanels: extra.minPanels, maxPanels: upper(extra.maxPanels), formula: condition.formula,
                    results: condition.results.map(result => ({id: String(result.id), componentId: String(result.choiceId), order: result.sequence}))};
            })})),
    };
    catalogue.legacyResults = [];
    for (const sourceLink of r200sSnapshot.linkedOptions) for (const sourceCondition of sourceLink.conditions) for (const result of sourceCondition.results) {
        const targetId = groupId(sourceLink.name);
        if (!groupFor(catalogue, targetId).components.some(component => component.id === String(result.choiceId))) {
            catalogue.legacyResults.push({groupId: targetId, conditionId: String(sourceCondition.id), resultId: String(result.id),
                component: {id: String(result.choiceId), code: result.code, name: result.description, available: false}});
        }
    }
    validateSetup(catalogue, document);
    return {catalogue, document};
}

export function validateSetup(catalogue: Catalogue, document: SetupDocument): void {
    if (catalogue.version !== 1 || ![1, 2].includes(document.version) || catalogue.blindType.code !== document.blindTypeCode) throw new Error('The setup belongs to a different blind type or format.');
    if (!Array.isArray(document.options) || !Array.isArray(document.links) || catalogue.groups.some(group => !document.options.some(option => option.groupId === group.id))) throw new Error('The blind-type option list is incomplete.');
    if (document.options.length > 500) throw new Error('A blind type can contain at most 500 options.');
    const seen = new Set<string>();
    const unique = (id: string): void => {if (!id || typeof id !== 'string' || id.length > 100 || seen.has(id)) throw new Error('The setup contains a missing or duplicate identifier.'); seen.add(id);};
    const position = (value: number): void => {if (!Number.isInteger(value) || value < 0 || value > 100000) throw new Error('Order must be a whole number between 0 and 100,000.');};
    const optionMap = new Map<string, OptionSetup>();
    const names = new Set<string>();
    for (const option of document.options) {
        if (typeof option.groupId !== 'string') throw new Error('The option identifier is invalid.');
        unique(`option:${option.groupId}`); position(option.order);
        const original = catalogue.groups.find(group => group.id === option.groupId);
        if (!original) {
            if (document.version !== 2 || !/^custom-[a-z0-9-]{1,80}$/i.test(option.groupId) || typeof option.sourceGroupId !== 'string') throw new Error('A new option must use a component type available for this blind type.');
            if (typeof option.name !== 'string' || !option.name.trim() || option.name.length > 100 || option.name !== option.name.trim()) throw new Error('Enter an option name of 1–100 characters.');
        } else if (option.sourceGroupId !== undefined || option.name !== undefined) {
            throw new Error('The original catalogue option identity cannot be changed.');
        }
        const group = optionGroup(catalogue, document, option.groupId);
        const key = group.name.normalize('NFKC').toLocaleLowerCase('en-GB');
        if (names.has(key)) throw new Error(`An option named ${group.name} already exists. Choose a different name.`);
        names.add(key);
        if (typeof option.enabled !== 'boolean' || typeof option.hidden !== 'boolean' || typeof option.freeText !== 'boolean' || !Array.isArray(option.componentIds)) throw new Error('Invalid option settings.');
        if (!original && !option.componentIds.length) throw new Error(`Choose at least one component available for ${catalogue.blindType.code}.`);
        if (new Set(option.componentIds).size !== option.componentIds.length || option.componentIds.some(id => !group.components.some(component => component.id === id && component.available))) throw new Error(`${group.name}: select only components available for ${catalogue.blindType.code}.`);
        if (option.defaultId !== null && !option.componentIds.includes(option.defaultId)) throw new Error(`${group.name}: the default component must be in its selected components.`);
        optionMap.set(option.groupId, option);
    }
    const targets = new Set<string>();
    const edges = new Map<string, Set<string>>();
    const usable = (groupId: string, id: string): boolean => !!optionMap.get(groupId)?.enabled && availableComponents(catalogue, document, groupId).some(component => component.id === id);
    for (const link of document.links) {
        unique(`link:${link.id}`); position(link.order); const group = optionGroup(catalogue, document, link.groupId);
        if (targets.has(link.groupId)) throw new Error(`There is already a linked option for ${group.name}.`);
        targets.add(link.groupId);
        if (!optionMap.get(link.groupId)?.enabled || !Array.isArray(link.conditions)) throw new Error(`${group.name}: enable the option before linking it.`);
        const dependencies = new Set<string>(); edges.set(link.groupId, dependencies);
        for (const condition of link.conditions) {
            unique(`condition:${condition.id}`); position(condition.order);
            if (!['all', 'any'].includes(condition.mode) || !Array.isArray(condition.tests) || condition.tests.length > 20 || !Array.isArray(condition.results)) throw new Error('Invalid linked-option condition.');
            const testGroups = new Set<string>();
            for (const test of condition.tests) {
                if (test.groupId === link.groupId || testGroups.has(test.groupId) || !usable(test.groupId, test.componentId)) throw new Error(`${group.name}: each condition must use an available component from a different enabled option.`);
                testGroups.add(test.groupId); dependencies.add(test.groupId);
            }
            for (const [min, max] of [[condition.minWidth, condition.maxWidth], [condition.minDrop, condition.maxDrop], [condition.minChainLength, condition.maxChainLength], [condition.minPanels, condition.maxPanels]] as [number, number | null][]) {
                if (!Number.isFinite(min) || min < 0 || min > 9999999 || (max !== null && (!Number.isFinite(max) || max < min || max > 9999999))) throw new Error('Size limits must be valid non-negative numbers, with maximum at least minimum.');
            }
            if (!Number.isInteger(condition.minPanels) || (condition.maxPanels !== null && !Number.isInteger(condition.maxPanels)) || !Number.isFinite(condition.maxArea) || condition.maxArea < 0) throw new Error('Panel counts must be whole numbers and maximum area cannot be negative.');
            if (condition.formula !== null && (typeof condition.formula !== 'string' || condition.formula.length > 4000)) throw new Error('The formula is too long.');
            const resultIds = new Set<string>();
            for (const result of condition.results) {
                unique(`result:${result.id}`); position(result.order);
                if (resultIds.has(result.componentId)) throw new Error(`${group.name}: a component can appear only once in a condition's results.`);
                resultIds.add(result.componentId);
                const legacy = catalogue.legacyResults?.some(item => item.groupId === link.groupId && item.conditionId === condition.id && item.resultId === result.id && item.component.id === result.componentId);
                if (!usable(link.groupId, result.componentId) && !legacy) throw new Error(`${group.name}: the result component is not available for this blind type and option.`);
            }
        }
    }
    const complete = new Set<string>(); const visiting = new Set<string>();
    const visit = (id: string): void => {if (visiting.has(id)) throw new Error('Linked options cannot depend on each other in a loop.'); if (complete.has(id)) return; visiting.add(id); for (const next of edges.get(id) ?? []) visit(next); visiting.delete(id); complete.add(id);};
    for (const id of edges.keys()) visit(id);
    if (JSON.stringify(document).length > 900000) throw new Error('This setup is too large to save.');
}

export function setupRows(catalogue: Catalogue, document: SetupDocument, mode: 'options' | 'links'): SetupRow[] {
    const rows: SetupRow[] = [];
    if (mode === 'options') for (const option of ordered(document.options)) {
        const group = optionGroup(catalogue, document, option.groupId); const id = `option-${group.id}`;
        rows.push({id, groupId: group.id, name: group.name, kind: 'Option', code: '', order: option.order,
            details: [option.enabled ? 'Enabled' : 'Disabled', option.hidden ? 'Hidden' : '', option.freeText ? 'Free text allowed' : '', isCustomOption(option) ? `Component type: ${groupFor(catalogue, option.sourceGroupId!).name}` : ''].filter(Boolean).join(' · '), count: option.componentIds.length});
        for (const component of group.components.filter(item => option.componentIds.includes(item.id))) rows.push({id: `component-${group.id}-${component.id}`, parentId: id, groupId: group.id, componentId: component.id,
            name: component.name, kind: 'Component', code: component.code, order: 0, details: component.id === option.defaultId ? 'Default component' : '', count: null});
    }
    else for (const link of ordered(document.links)) {
        const group = optionGroup(catalogue, document, link.groupId); const id = `link-${link.id}`;
        rows.push({id, linkId: link.id, groupId: group.id, name: group.name, kind: 'Linked option', code: '', order: link.order, details: `${link.conditions.length} conditions`, count: link.conditions.reduce((sum, item) => sum + item.results.length, 0)});
        for (const condition of ordered(link.conditions)) {
            const conditionId = `condition-${condition.id}`;
            const tests = condition.tests.map(test => optionGroup(catalogue, document, test.groupId).components.find(component => component.id === test.componentId)!.name);
            const width = `Width ${condition.minWidth.toLocaleString('en-GB')}–${condition.maxWidth?.toLocaleString('en-GB') ?? 'any'} mm`;
            const limits = [condition.minWidth || condition.maxWidth !== null ? width : '', condition.minDrop || condition.maxDrop !== null ? `Drop ${condition.minDrop}–${condition.maxDrop ?? 'any'} mm` : ''].filter(Boolean).join(' · ');
            rows.push({id: conditionId, parentId: id, linkId: link.id, conditionId: condition.id, groupId: group.id, name: tests.join(condition.mode === 'all' ? ' + ' : ' OR ') || width,
                kind: 'Condition', code: '', order: condition.order, details: [condition.results.length ? '' : 'No results configured', limits].filter(Boolean).join(' · '), count: condition.results.length});
            for (const result of ordered(condition.results)) {
                const component = resultComponent(catalogue, document, group.id, condition.id, result);
                rows.push({id: `result-${result.id}`, parentId: conditionId, linkId: link.id, conditionId: condition.id, resultId: result.id, groupId: group.id, componentId: component.id,
                    name: component.name, kind: 'Result', code: component.code, order: result.order, details: component.available ? '' : 'Existing result: unavailable for new selections', count: null});
            }
        }
    }
    return rows;
}

export const newCondition = (id: string, order: number): SetupCondition => ({id, order, mode: 'all', tests: [], minWidth: 0, maxWidth: null, minDrop: 0, maxDrop: null,
    maxArea: 0, minChainLength: 0, maxChainLength: null, minPanels: 0, maxPanels: null, formula: null, results: []});
export const setupColumns = [
    ['name', 'Option / component', 'SingleLine.Text', 430], ['kind', 'Type', 'SingleLine.Text', 110], ['code', 'Component code', 'SingleLine.Text', 150],
    ['order', 'Order', 'Whole.None', 65], ['details', 'Settings', 'SingleLine.Text', 260], ['count', 'Choices / results', 'Whole.None', 100],
].map(([name, displayName, dataType, visualSizeFactor], order) => ({name, displayName, dataType, visualSizeFactor, order, alias: name, isHidden: false})) as ComponentFramework.PropertyHelper.DataSetApi.Column[];
export function setupRecord(row: SetupRow): ComponentFramework.PropertyHelper.DataSetApi.EntityRecord {
    return {getRecordId: () => row.id, getNamedReference: () => ({id: {guid: row.id}, name: row.name, etn: 'blind_type_setup'}),
        getValue: (name: string) => row[name as keyof SetupRow], getFormattedValue: (name: string) => String(row[name as keyof SetupRow] ?? '')} as ComponentFramework.PropertyHelper.DataSetApi.EntityRecord;
}
