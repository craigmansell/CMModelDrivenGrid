import * as React from 'react';
import {Dialog, DialogFooter} from '@fluentui/react/lib/Dialog';
import {DefaultButton} from '@fluentui/react/lib/Button';
import {ResponsiveGrid} from '../CMModelDrivenGrid/ResponsiveGrid';
import {ChildRequest} from '../CMModelDrivenGrid/childRecords';
import {FilterExpression, updateColumnFilter} from '../CMModelDrivenGrid/filtering';
import {checkSelectionCancelled, selectionResult} from '../CMModelDrivenGrid/recordSelection';
import {DemoFilterExpression, evaluateDemoFilter, sortDemoRows} from './demo-filter-evaluator';
import {pageDemoRows} from './demo-data';
import {createR200sDemoRows, r200sColumns, r200sRecord, r200sSnapshot, R200sDemoRow} from './r200s-demo-data';

const rows = createR200sDemoRows();
const byId = new Map(rows.map(row => [row.id, row]));
const roots = rows.filter(row => !row.parentId);
const branches = new Map<string, R200sDemoRow[]>();
rows.forEach(row => { if (row.parentId) branches.set(row.parentId, [...branches.get(row.parentId) ?? [], row]); });
const columnKinds = {name: 'text', kind: 'text', code: 'text', sequence: 'number', details: 'text', resultCount: 'number'} as const;
const raw = (row: R200sDemoRow, name: string): unknown => row[name as keyof R200sDemoRow];
const queryRows = (candidates: R200sDemoRow[], filtering: FilterExpression | undefined, sorting: ChildRequest['sorting']): R200sDemoRow[] =>
    sortDemoRows(candidates.filter(row => evaluateDemoFilter(row, filtering as unknown as DemoFilterExpression | undefined, {getValue: raw, columnKinds})),
        sorting.map(sort => ({columnName: sort.name, direction: sort.sortDirection === 1 ? 'desc' as const : 'asc' as const})),
        {getValue: raw, getDisplayValue: (row, name) => String(raw(row, name) ?? ''), columnKinds});
const resources = {
    getString: (key: string) => ({Label_Grid_Footer_RecordCount: '{0} records ({1} Selected)', Label_Grid_Footer: 'Page {0}',
        Label_NoRecords: 'No matching records', Label_SortAZ: 'A to Z', Label_SortZA: 'Z to A', Label_DoesNotContainData: 'Does not contain data'}[key] ?? key),
    getResource: () => undefined,
} as unknown as ComponentFramework.Resources;

export function R200sDemo(): React.ReactElement {
    React.useEffect(() => {document.title = 'R200S linked options · DrivenGrid';}, []);
    const [generation, setGeneration] = React.useState(0);
    const [filtering, setFiltering] = React.useState<FilterExpression>();
    const [sorting, setSorting] = React.useState<ChildRequest['sorting']>([]);
    const [selectedIds, setSelectedIds] = React.useState<string[]>([]);
    const [openRow, setOpenRow] = React.useState<R200sDemoRow>();
    const visible = React.useMemo(() => queryRows(roots, filtering, sorting), [filtering, sorting]);
    const records = React.useMemo(() => Object.fromEntries(visible.map(row => [row.id, r200sRecord(row)])), [visible]);
    const ids = React.useMemo(() => visible.map(row => row.id), [visible]);
    const loadChildren = React.useCallback(async (request: ChildRequest) => {
        const matches = queryRows(branches.get(request.parentId) ?? [], request.filtering, request.sorting);
        const page = pageDemoRows(matches, request.page);
        return {records: page.rows.map(r200sRecord), hasNextPage: page.hasNextPage};
    }, []);
    const checkChildren = React.useCallback(async (parentIds: string[]) =>
        Object.fromEntries(parentIds.map(id => [id, !!branches.get(id)?.length])), []);
    const reset = (): void => {setFiltering(undefined); setSorting([]); setSelectedIds([]); setOpenRow(undefined); setGeneration(value => value + 1);};
    const fullTab = new URLSearchParams(window.location.search).get('fulltab') !== '0';
    return <>
        <header className="demo-header">{r200sSnapshot.blindType.description} · Linked options demo</header>
        <main className={`demo-main${fullTab ? ' demo-main-full-tab' : ''}`}>
            <p style={{margin: '4px 0 12px', flexShrink: 0}}>Expand the arrows to view conditions and result choices. Click a column header again to close its menu.</p>
            <div className="cm-child-toolbar" role="toolbar" aria-label="R200S demo commands">
                <span role="status">{selectedIds.length ? `${selectedIds.length} selected` : '4 linked options · 32 conditions · 68 results'} · Read-only demo · 17 Nov 2026</span>
                <div className="cm-child-commands">
                    <a href="?demo=r200s&setup=1">Set up blind type</a>
                    <button type="button" disabled={selectedIds.length !== 1} onClick={() => setOpenRow(byId.get(selectedIds[0]))}>Open selected</button>
                    <button type="button" onClick={reset}>Reset demo</button>
                </div>
            </div>
            <div className="demo-grid"><ResponsiveGrid key={generation} columns={r200sColumns} records={records} sortedRecordIds={ids}
                layoutKey="local-demo:r200s:v1" features={{showTypeIcons: true, locale: 'en-GB', accessibleLabel: 'R200S linked options'}}
                childRows={{loadPage: loadChildren, hasChildren: checkChildren, readOnly: true}}
                hasNextPage={false} hasPreviousPage={false} totalResultCount={visible.length} currentPage={1}
                sorting={sorting} filtering={filtering!} resources={resources} itemsLoading={false}
                highlightValue={null} highlightColor={null} enableLookupLinks={false} enableInlineEdit={false}
                setSelectedRecords={setSelectedIds} onNavigate={record => setOpenRow(byId.get(record?.getRecordId() ?? ''))} onOpenLookup={() => undefined}
                onSelectAllRecords={request => {checkSelectionCancelled(request?.signal); return selectionResult(ids);}}
                onFilter={(name, mode, value) => setFiltering(current => updateColumnFilter(current, name, mode, value))}
                onClearAllFilters={() => setFiltering(undefined)}
                onSort={(name, direction, additive) => setSorting(current => {
                    const next = direction === 'none' || additive ? current.filter(sort => sort.name !== name) : [];
                    if (direction !== 'none') next.splice(additive && current.some(sort => sort.name === name) ? current.findIndex(sort => sort.name === name) : next.length, 0, {name, sortDirection: direction === 'desc' ? 1 : 0});
                    return next;
                })}
                loadFirstPage={() => undefined} loadNextPage={() => undefined} loadPreviousPage={() => undefined}
                onUpdateCell={() => {throw new Error('This R200S snapshot is read only.');}}
            /></div>
        </main>
        <Dialog hidden={!openRow} onDismiss={() => setOpenRow(undefined)} dialogContentProps={{title: openRow?.name, subText: openRow?.kind}}>
            {openRow && <dl><dt>Code</dt><dd>{openRow.code ?? '(None)'}</dd><dt>Order</dt><dd>{openRow.sequence}</dd><dt>Details</dt><dd>{openRow.details || '(None)'}</dd></dl>}
            <DialogFooter><DefaultButton text="Close" onClick={() => setOpenRow(undefined)}/></DialogFooter>
        </Dialog>
    </>;
}
