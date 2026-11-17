/** Read-only product configuration captured from Blindata on 17 November 2026. */
export interface R200sChoice { group: string; code: string; description: string }
export interface R200sResult { id: number; sequence: number; choiceId: number; code: string; description: string }
export interface R200sCondition {
    id: number; sequence: number; choices: R200sChoice[]; rule: string;
    minWidth: number; maxWidth: number; minDrop: number; maxDrop: number;
    formula: string | null; results: R200sResult[];
}
export interface R200sSnapshot {
    capturedOn: string;
    blindType: { id: number; code: string; description: string };
    linkedOptions: { id: number; sequence: number; name: string; conditions: R200sCondition[] }[];
}
export const r200sSnapshot = require('./fixtures/r200s-linked-options.json') as R200sSnapshot;
export interface R200sDemoRow {
    id: string; parentId?: string; name: string; kind: 'Linked option' | 'Condition' | 'Result';
    code: string | null; sequence: number; details: string; resultCount: number | null;
}
const number = (value: number): string => value.toLocaleString('en-GB');
const inSequence = <T extends {sequence: number; id: number}>(items: T[]): T[] =>
    [...items].sort((a, b) => a.sequence - b.sequence || a.id - b.id);
const limits = (condition: R200sCondition): string => [
    condition.maxWidth < 9999999 || condition.minWidth > 0 ? `Width ${number(condition.minWidth)}–${number(condition.maxWidth)} mm` : '',
    condition.maxDrop < 9999999 || condition.minDrop > 0 ? `Drop ${number(condition.minDrop)}–${number(condition.maxDrop)} mm` : '',
    condition.formula,
].filter(Boolean).join(' · ');
const ruleDescription = (condition: R200sCondition): string => condition.rule.replace(/\bChoice([123])\b/g,
    (token, ordinal: string) => condition.choices[Number(ordinal) - 1]?.group ?? token);

export function createR200sDemoRows(snapshot = r200sSnapshot): R200sDemoRow[] {
    return inSequence(snapshot.linkedOptions).flatMap(group => {
        const groupId = `r200s-link-${group.id}`;
        const dependencies = [...new Set(group.conditions.flatMap(condition => condition.choices.map(choice => choice.group)))];
        return [
            {id: groupId, name: group.name, kind: 'Linked option' as const, code: null, sequence: group.sequence,
                details: `${group.conditions.length} conditions · ${dependencies.join(' + ') || 'Blind width'}`,
                resultCount: group.conditions.reduce((sum, condition) => sum + condition.results.length, 0)},
            ...inSequence(group.conditions).flatMap(condition => {
                const conditionId = `r200s-condition-${condition.id}`;
                const constraint = limits(condition);
                return [
                    {id: conditionId, parentId: groupId, name: condition.choices.map(choice => choice.description).join(' + ') || constraint,
                        kind: 'Condition' as const, code: `Rule ${condition.id}`, sequence: condition.sequence,
                        details: [condition.results.length ? '' : 'No results configured', ruleDescription(condition), condition.choices.length ? constraint : ''].filter(Boolean).join(' · '),
                        resultCount: condition.results.length},
                    ...inSequence(condition.results).map(result => ({id: `r200s-result-${result.id}`, parentId: conditionId,
                        name: result.description, kind: 'Result' as const, code: result.code,
                        sequence: result.sequence, details: `Choice ID ${result.choiceId}`, resultCount: null})),
                ];
            }),
        ];
    });
}

export const r200sColumns = [
    ['name', 'Name', 'SingleLine.Text', 460], ['kind', 'Type', 'SingleLine.Text', 110],
    ['code', 'Code', 'SingleLine.Text', 160], ['sequence', 'Order', 'Whole.None', 65],
    ['details', 'Details', 'SingleLine.Text', 260], ['resultCount', 'Results', 'Whole.None', 70],
].map(([name, displayName, dataType, visualSizeFactor], order) => ({
    name, displayName, dataType, visualSizeFactor, order, alias: name, isHidden: false,
})) as ComponentFramework.PropertyHelper.DataSetApi.Column[];

export function r200sRecord(row: R200sDemoRow): ComponentFramework.PropertyHelper.DataSetApi.EntityRecord {
    return {
        getRecordId: () => row.id,
        getNamedReference: () => ({id: {guid: row.id}, name: row.name, etn: 'r200s_demo'}),
        getValue: (name: string) => row[name as keyof R200sDemoRow],
        getFormattedValue: (name: string) => String(row[name as keyof R200sDemoRow] ?? ''),
    } as ComponentFramework.PropertyHelper.DataSetApi.EntityRecord;
}
