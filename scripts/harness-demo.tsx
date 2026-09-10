import * as React from "react";
import * as ReactDOM from "react-dom";
import { initializeIcons } from "@fluentui/react/lib/Icons";
import { Dialog, DialogFooter } from '@fluentui/react/lib/Dialog';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { HierarchyGrid as Grid } from "../CMModelDrivenGrid/HierarchyGrid";
import { ChildRequest } from "../CMModelDrivenGrid/childRecords";
import { FilterExpression, updateColumnFilter } from "../CMModelDrivenGrid/filtering";
import { DemoColumnKind, DemoFilterExpression, evaluateDemoFilter, getLocalDateOrdinal, sortDemoRows } from "./demo-filter-evaluator";
import { DemoRow as Row, DemoScenario, demoScenario, seedDemoRows, indexDemoRows, pageDemoRows } from './demo-data';
import { checkSelectionCancelled, selectionResult } from '../CMModelDrivenGrid/recordSelection';
import "../CMModelDrivenGrid/css/CMModelDrivenGrid.css";

initializeIcons(undefined, { disableWarnings: true });
const owners = ["Alex Morgan", "Taylor Reed", "Jordan Patel", "Sam Wilson"];
const ownerId = (index: number) => `00000000-0000-4000-8000-${String(index + 1).padStart(12, "0")}`;
const statuses = ["Draft", "In progress", "Complete"];
const money = new Intl.NumberFormat("en-GB", { style: "currency", currency: "GBP" });
const columnKinds: Record<string, DemoColumnKind> = {
    name: "text", amount: "number", status: "option", owner: "lookup", start: "dateOnly", active: "boolean",
};
const columns = [
    ["name", "Project", "SingleLine.Text", 180], ["amount", "Budget", "Currency", 110],
    ["status", "Status", "OptionSet", 120], ["owner", "Owner", "Lookup.Simple", 130],
    ["start", "Start date", "DateAndTime.DateOnly", 110], ["active", "Active", "TwoOptions", 90],
].map(([name, displayName, dataType, factor], order) => ({ name, displayName, dataType, visualSizeFactor: factor, order, alias: name, isHidden: false })) as ComponentFramework.PropertyHelper.DataSetApi.Column[];
const resources = {
    getString: (key: string) => ({
        Label_Grid_Footer_RecordCount: "{0} records ({1} Selected)", Label_Grid_Footer: "Page {0}",
        Label_NoRecords: "No matching records", Label_SortAZ: "A to Z", Label_SortZA: "Z to A",
        Label_DoesNotContainData: "Does not contain data",
    }[key] ?? key),
    getResource: () => undefined,
} as unknown as ComponentFramework.Resources;
const raw = (row: Row, column: string): unknown => column === "owner"
    ? row.owner === null ? null : ownerId(row.owner)
    : row[column as keyof Row];
const displayValue = (row: Row, column: string): string => {
    if (raw(row, column) === null || raw(row, column) === undefined) return "";
    if (column === "amount") return money.format(row.amount!);
    if (column === "status") return statuses[row.status!];
    if (column === "owner") return owners[row.owner!];
    if (column === "active") return row.active ? "Yes" : "No";
    return String(row[column as keyof Row] ?? "");
};
function record(row: Row): ComponentFramework.PropertyHelper.DataSetApi.EntityRecord {
    return {
        getRecordId: () => row.id,
        getNamedReference: () => ({ id: { guid: row.id }, name: row.name, etn: "sample" }),
        getValue: (name: string) => name === "owner" && row.owner !== null
            ? { id: { guid: ownerId(row.owner) }, name: owners[row.owner], etn: "systemuser" }
            : row[name as keyof Row],
        getFormattedValue: (name: string) => {
            if (raw(row, name) === null) return "";
            if (name === "amount") return money.format(row.amount!);
            if (name === "status") return statuses[row.status!];
            if (name === "owner") return owners[row.owner!];
            if (name === "active") return row.active ? "Yes" : "No";
            return String(row[name as keyof Row] ?? "");
        },
    } as ComponentFramework.PropertyHelper.DataSetApi.EntityRecord;
}
function Demo(): React.ReactElement {
    const [scenario, setScenario] = React.useState(() => demoScenario(new URLSearchParams(window.location.search).get('scenario')));
    const [rows, setRows] = React.useState(() => seedDemoRows(new Date(), scenario));
    const dataIndex = React.useMemo(() => indexDemoRows(rows), [rows]);
    const [childRequests, setChildRequests] = React.useState({pages: 0, rows: 0});
    const [filter, setFilter] = React.useState<FilterExpression>();
    const [sorting, setSorting] = React.useState<ComponentFramework.PropertyHelper.DataSetApi.SortStatus[]>([]);
    const [limit, setLimit] = React.useState(25);
    const [loading, setLoading] = React.useState(false);
    const [generation, setGeneration] = React.useState(0);
    const [childReadOnly, setChildReadOnly] = React.useState(false);
    const [selectedIds, setSelectedIds] = React.useState<string[]>([]);
    const selectedRows = selectedIds.map(id => dataIndex.byId.get(id)).filter((row): row is Row => !!row);
    const [commandError, setCommandError] = React.useState('');
    const [openRow, setOpenRow] = React.useState<Row>();
    const [demoDialog,setDemoDialog] = React.useState<{name:string;resolve:(value:boolean)=>void}>();
    const askDemo = (name:string): Promise<boolean> => new Promise(resolve => setDemoDialog({name,resolve}));
    const finishDialog = (value:boolean) => {demoDialog?.resolve(value);setDemoDialog(undefined);};
    const [childRefresh, setChildRefresh] = React.useState(0);
    const [revision, setRevision] = React.useState(0);
    const timer = React.useRef<ReturnType<typeof setTimeout>>();
    const filteredRows = React.useMemo(() => {
        const next = dataIndex.parents.filter((row) => evaluateDemoFilter(row, filter as unknown as DemoFilterExpression | undefined, {
            getValue: raw,
            columnKinds,
            now: new Date(),
        }));
        const sortLevels = sorting.map((sort) => ({
            columnName: sort.name,
            direction: sort.sortDirection === 1 ? "desc" as const : "asc" as const,
        }));
        return sortLevels.length ? sortDemoRows(next, sortLevels, {
            getValue: raw,
            getDisplayValue: displayValue,
            columnKinds,
        }) : next;
    }, [dataIndex, filter, sorting]);
    const visible = React.useMemo(() => filteredRows.slice(0, limit), [filteredRows, limit, revision]);
    const records = React.useMemo(() => Object.fromEntries(visible.map((row) => [row.id, record(row)])), [visible]);
    const ids = React.useMemo(() => visible.map((row) => row.id), [visible]);
    const resetPage = () => { clearTimeout(timer.current); setLoading(false); setLimit(25); setRevision((value) => value + 1); };
    const resetScenario = (next: DemoScenario) => {
        resetPage(); setScenario(next); setRows(seedDemoRows(new Date(), next)); setFilter(undefined); setSorting([]);
        setSelectedIds([]); setCommandError(''); setOpenRow(undefined); setChildRequests({pages: 0, rows: 0}); setGeneration(value => value + 1);
        const url = new URL(window.location.href);
        if (next === 'standard') url.searchParams.delete('scenario'); else url.searchParams.set('scenario', next);
        window.history.replaceState(null, '', url);
    };
    const loadChildren = React.useCallback(async (request: ChildRequest) => {
        const matches = (dataIndex.children.get(request.parentId) ?? []).filter(row => evaluateDemoFilter(row,request.filtering as unknown as DemoFilterExpression | undefined,{getValue:raw,columnKinds,now:new Date()}));
        const sorted = sortDemoRows(matches,request.sorting.map(sort => ({columnName:sort.name,direction:sort.sortDirection === 1 ? 'desc' as const : 'asc' as const})),{getValue:raw,getDisplayValue:displayValue,columnKinds});
        const page = pageDemoRows(sorted, request.page);
        setChildRequests(current => ({pages: current.pages + 1, rows: current.rows + page.rows.length}));
        return {records:page.rows.map(record),hasNextPage:page.hasNextPage};
    },[dataIndex]);
    const checkChildren = React.useCallback(async (parentIds: string[]) =>
        Object.fromEntries(parentIds.map(id => [id,!!dataIndex.children.get(id)?.length])),[dataIndex]);
    React.useEffect(() => () => clearTimeout(timer.current), []);
    const loadMore = () => {
        if (loading || limit >= filteredRows.length) return;
        setLoading(true);
        timer.current = setTimeout(() => { setLimit((value) => value + 25); setLoading(false); }, 250);
    };
    const deleteSelected = async () => {
        setCommandError('');
        const targets = [...selectedIds];
        if (rows.some(row => row.parentId && targets.includes(row.parentId))) {
            setCommandError('This sample selection includes an option with children. Remove its children first.');
            return;
        }
        if (!targets.length || !await askDemo(selectedRows.map(row => row.name).join(', '))) return;
        setSelectedIds([]);
        setRows(current => current.filter(row => !targets.includes(row.id)));
    };
    return <>
        <header className="demo-header">Dynamics 365 style · DrivenGrid test harness</header>
        <main className="demo-main">
            <div className="demo-title"><div><h1>Projects</h1><p className="demo-caption">{filteredRows.length} matching sample records · {visible.length} loaded</p></div>
                <button className="demo-reset" onClick={() => resetScenario('standard')}>Reset sample data</button></div>
            <div className="demo-help"><p className="demo-dataset-summary" data-testid="demo-dataset-summary">{dataIndex.parents.length.toLocaleString('en-GB')} parents · {(rows.length - dataIndex.parents.length).toLocaleString('en-GB')} children · {rows.length.toLocaleString('en-GB')} total records</p>
                Expand a row to see child options. {scenario === 'parents-2000-children-20' ? 'Every parent has exactly 20 children, with no grandchildren.' : `The first project has ${scenario === 'large' ? '2,000' : '85'} children and a nested option.`} Select child checkboxes, then use the single command bar above the grid. Child panels contain no action buttons. Only one grid level is selected at a time. Collapsing clears hidden selection but keeps filters and scroll position.<br/>Sample commands only — no Dynamics connection. Native child commands require EnableNativeChildSelection and a bound view that includes children; verify these in your signed-in app.<br/>
                Fiscal filters in this local fixture use January–December with monthly periods. Dataverse uses your organisation's fiscal calendar.<br/>
                <button type="button" onClick={() => resetScenario('parents-2000-children-20')}>Use 2,000 parents × 20 children</button>{' '}
                <button type="button" onClick={() => resetScenario('large')}>Use 10,000-row test data</button>{' '}
                {scenario !== 'standard' && <button type="button" onClick={() => setLimit(filteredRows.length)}>Preload local rows for rendering test</button>}<br/>
                <span data-testid="demo-child-requests">Child page requests: {childRequests.pages.toLocaleString('en-GB')} · Child rows returned: {childRequests.rows.toLocaleString('en-GB')}</span><br/>
                <label><input type="checkbox" checked={childReadOnly} onChange={event => setChildReadOnly(event.target.checked)}/> Read-only child cells (parents remain editable)</label>
            </div>
            <div className="cm-child-toolbar" role="toolbar" aria-label="Sample record commands">
                <span role="status">{selectedRows.length} selected{selectedRows.length === 1 ? ` · ${selectedRows[0].name}` : ''}</span>
                <div className="cm-child-commands">
                    <button type="button" disabled={selectedRows.length !== 1} onClick={() => setOpenRow(selectedRows[0])}>Open selected</button>
                    <button type="button" disabled={!selectedRows.length} onClick={() => setRows(current => current.map(row => selectedIds.includes(row.id) ? {...row,active:true} : row))}>Activate</button>
                    <button type="button" disabled={!selectedRows.length} onClick={() => setRows(current => current.map(row => selectedIds.includes(row.id) ? {...row,active:false} : row))}>Deactivate</button>
                    <button type="button" disabled={!selectedRows.length} onClick={() => void deleteSelected()}>Delete selected</button>
                    <button type="button" onClick={() => {setSelectedIds([]);setCommandError('');setChildRefresh(value => value + 1);resetPage();}}>Refresh</button>
                </div>
            </div>
            {commandError && <div role="alert">{commandError}</div>}
            <div className="demo-grid"><Grid key={generation} columns={columns} records={records} sortedRecordIds={ids} selectionResetVersion={childRefresh}
                features={{reflow:new URLSearchParams(window.location.search).get('layout') === 'list' ? 'list' : 'auto', showAvatar:true, showTypeIcons:true, showChoiceColors:true}}
                childRows={{loadPage:loadChildren,hasChildren:checkChildren,revision:childRefresh,readOnly:childReadOnly}}
                hasNextPage={limit < filteredRows.length} hasPreviousPage={false} totalResultCount={filteredRows.length} currentPage={1}
                sorting={sorting} filtering={filter!} resources={resources} itemsLoading={loading}
                highlightValue={null} highlightColor={null} enableLookupLinks={false} enableInlineEdit={true} infiniteScroll={true}
                columnOptions={{ status: statuses.map((label, value) => ({ label, value: String(value), color: ['#f3f2f1','#fff4ce','#dff6dd'][value] })) }} lookupTargetEntityMap={{ owner: "systemuser" }}
                setSelectedRecords={setSelectedIds} onNavigate={item => setOpenRow(rows.find(row => row.id === item?.getRecordId()))} onOpenLookup={() => undefined}
                onSelectAllRecords={options => {
                    checkSelectionCancelled(options?.signal);
                    const result = selectionResult(filteredRows.map(row => row.id));
                    clearTimeout(timer.current); setLoading(false);
                    setLimit(current => Math.max(current, result.ids.length));
                    return result;
                }}
                onFilter={(name, mode, value) => { resetPage(); setFilter((current) => updateColumnFilter(current, name, mode, value)); }}
                onClearAllFilters={() => { resetPage(); setFilter(undefined); }}
                onSort={(name, direction, additive) => {
                    resetPage();
                    setSorting((current) => {
                        if (direction === "none") return current.filter((sort) => sort.name !== name);
                        const nextLevel: ComponentFramework.PropertyHelper.DataSetApi.SortStatus = {
                            name,
                            sortDirection: direction === "desc" ? 1 : 0,
                        };
                        if (!additive) return [nextLevel];
                        const existingIndex = current.findIndex((sort) => sort.name === name);
                        if (existingIndex < 0) return [...current, nextLevel];
                        const next = [...current];
                        next[existingIndex] = nextLevel;
                        return next;
                    });
                }}
                loadFirstPage={resetPage} loadPreviousPage={resetPage} loadNextPage={loadMore}
                onUpdateCell={async (id, name, value) => {
                    const trimmed = value.trim();
                    let parsed: string | number | boolean | null = trimmed === "" ? null : trimmed;
                    if (name === "amount" && trimmed !== "") {
                        parsed = Number(trimmed);
                        if (!Number.isFinite(parsed)) throw new Error("Enter a valid number.");
                    } else if (name === "status" && trimmed !== "") {
                        if (!/^\d+$/.test(trimmed)) throw new Error("Enter a valid status.");
                        parsed = Number(trimmed);
                        if (!Number.isInteger(parsed) || parsed < 0 || parsed >= statuses.length) throw new Error("Enter a valid status.");
                    } else if (name === "active" && trimmed !== "") {
                        const lowered = trimmed.toLocaleLowerCase();
                        if (lowered !== "true" && lowered !== "false") throw new Error("Enter true or false.");
                        parsed = lowered === "true";
                    } else if (name === "start" && trimmed !== "") {
                        if (!/^\d{4}-\d{2}-\d{2}$/.test(trimmed) || getLocalDateOrdinal(trimmed) === null) {
                            throw new Error("Enter a valid date (YYYY-MM-DD).");
                        }
                        parsed = trimmed;
                    }
                    setRows((current) => current.map((row) => row.id === id ? { ...row, [name]: parsed } : row));
                }}
                onSearchLookup={async (_entity, term) => owners.map((name, index) => ({ name, id: ownerId(index) })).filter((owner) => owner.name.toLocaleLowerCase().includes(term.toLocaleLowerCase()))}
                onUpdateLookupCell={async (id, targetColumn, targetId) => { setRows((current) => current.map((row) => row.id === id ? { ...row, [targetColumn]: targetId ? owners.findIndex((_name, index) => ownerId(index) === targetId) : null } : row)); }}
            /></div>
        </main>
        <Dialog hidden={!openRow} onDismiss={() => setOpenRow(undefined)} dialogContentProps={{title:openRow?.name,subText:`${openRow?.parentId ? 'Child option' : 'Parent record'} · ${openRow?.id ?? ''}`}}>
            {openRow && <dl><dt>Parent</dt><dd>{rows.find(row => row.id === openRow.parentId)?.name ?? '(None)'}</dd><dt>Budget</dt><dd>{displayValue(openRow,'amount')}</dd><dt>Status</dt><dd>{displayValue(openRow,'status')}</dd><dt>Active</dt><dd>{displayValue(openRow,'active') || '(Blank)'}</dd></dl>}
            <DialogFooter><DefaultButton text="Close" onClick={() => setOpenRow(undefined)}/></DialogFooter>
        </Dialog>
        <Dialog hidden={!demoDialog} onDismiss={() => finishDialog(false)} modalProps={{isBlocking:true}}
            dialogContentProps={{title:'Delete selected sample records?',subText:`Delete ${demoDialog?.name ?? ''}? This only removes disposable sample records.`}}>
            <DialogFooter><PrimaryButton text="Delete" onClick={() => finishDialog(true)}/><DefaultButton text="Cancel" onClick={() => finishDialog(false)}/></DialogFooter>
        </Dialog>
    </>;
}
ReactDOM.render(<Demo />, document.getElementById("demo-root"));
