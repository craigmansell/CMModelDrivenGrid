import * as React from 'react';
import { Grid, GridProps } from './Grid';
import { ChildRowsConfig, ChildRequest, ChildPage } from './childRecords';
import { FilterExpression, updateColumnFilter } from './filtering';
import { GridStateCache, useGridState } from './gridState';
import { MAX_SELECTED_RECORDS, SelectAllRequest, selectionResult, checkSelectionCancelled } from './recordSelection';

export interface HierarchyGridProps extends GridProps { childRows?: ChildRowsConfig; ancestors?: string[]; hierarchyVisible?: boolean; selectionResetVersion?: number }

interface HierarchySelection {
    activeScope: string;
    selectedCount: number;
    selectedIds: string[];
    states: Map<string, GridStateCache>;
    select: (scope: string, ids: string[], activateEmpty?: boolean) => void;
    clearBranch: (scope: string) => void;
}
const SelectionContext = React.createContext<HierarchySelection | undefined>(undefined);
const selectionScope = (ancestors: string[]) => `root${ancestors.map(id => `/${encodeURIComponent(id)}`).join('')}`;
const availabilityId = (id: string) => id.replace(/[{}]/g,'').toLowerCase();

export function HierarchyGrid(props: HierarchyGridProps): React.ReactElement {
    const inherited = React.useContext(SelectionContext);
    return inherited ? <HierarchyLevel {...props}/> : <HierarchyRoot {...props}/>;
}

function HierarchyRoot(props: HierarchyGridProps): React.ReactElement {
    const [activeScope, setActiveScope] = React.useState('root');
    const [selectedCount, setSelectedCount] = React.useState(0);
    const [selectedIds, setSelectedIds] = React.useState<string[]>([]);
    const states = React.useRef(new Map<string, GridStateCache>()).current;
    const active = React.useRef('root');
    const publish = React.useRef(props.setSelectedRecords);
    publish.current = props.setSelectedRecords;
    const select = React.useCallback((scope: string, ids: string[], activateEmpty = false) => {
        // Inactive lists can emit an empty selection when their items refresh.
        if (!ids.length && active.current !== scope && !activateEmpty) return;
        active.current = scope;
        setActiveScope(scope);
        setSelectedCount(ids.length);
        setSelectedIds(ids);
        publish.current(ids);
    }, []);
    const clearBranch = React.useCallback((scope: string) => {
        if (active.current !== scope && !active.current.startsWith(`${scope}/`)) return;
        active.current = 'root';
        setActiveScope('root');
        setSelectedCount(0);
        setSelectedIds([]);
        publish.current([]);
    }, []);
    const resetVersion = React.useRef(props.selectionResetVersion);
    React.useEffect(() => {
        if (resetVersion.current === props.selectionResetVersion) return;
        resetVersion.current = props.selectionResetVersion;
        active.current = 'root'; setActiveScope('root'); setSelectedCount(0);
        setSelectedIds([]);
    }, [props.selectionResetVersion]);
    const value = React.useMemo(() => ({activeScope, selectedCount, selectedIds, states, select, clearBranch}), [activeScope, selectedCount, selectedIds, states, select, clearBranch]);
    return <SelectionContext.Provider value={value}><HierarchyLevel {...props}/></SelectionContext.Provider>;
}

function HierarchyLevel(props: HierarchyGridProps): React.ReactElement {
    const selection = React.useContext(SelectionContext)!;
    const {states, clearBranch} = selection;
    const ancestors = props.ancestors ?? [];
    const scope = selectionScope(ancestors);
    if (!selection.states.has(scope)) selection.states.set(scope, new Map());
    const stateCache = selection.states.get(scope)!;
    const checker = props.childRows?.hasChildren;
    const revision = props.childRows?.revision;
    const idsKey = JSON.stringify([...new Set(props.sortedRecordIds.map(availabilityId))].sort());
    const [availability,setAvailability] = React.useState<Record<string,boolean>>({});
    const [availabilityError,setAvailabilityError] = React.useState('');
    const checked = React.useRef<Record<string,boolean>>({});
    const cacheVersion = React.useRef({checker,revision});
    React.useEffect(() => {
        if (props.itemsLoading) return;
        const prefix = `${scope}/`;
        const ids = new Set(props.sortedRecordIds);
        for (const key of states.keys()) {
            if (!key.startsWith(prefix)) continue;
            const parent = decodeURIComponent(key.slice(prefix.length).split('/')[0]);
            if (!ids.has(parent)) { clearBranch(`${prefix}${encodeURIComponent(parent)}`); states.delete(key); }
        }
    }, [props.sortedRecordIds, props.itemsLoading, scope, states, clearBranch]);
    React.useEffect(() => {
        if (!checker || props.hierarchyVisible === false) return;
        if (cacheVersion.current.checker !== checker || cacheVersion.current.revision !== revision) {
            checked.current = {};
            cacheVersion.current = {checker,revision};
        }
        const ids = JSON.parse(idsKey) as string[];
        const missing = ids.filter(id => checked.current[id] === undefined);
        if (!missing.length) return;
        let cancelled = false;
        setAvailabilityError('');
        void checker(missing).then(result => {
            if (cancelled) return;
            for (const id of missing) {
                if (typeof result[id] !== 'boolean') throw new Error('The child check was incomplete.');
            }
            checked.current = {...checked.current,...result};
            setAvailability({...checked.current});
            return undefined;
        }).catch((error: unknown) => {
            if (!cancelled) setAvailabilityError(error instanceof Error ? error.message : 'Unable to check child options.');
        });
        return () => {cancelled = true;};
    }, [checker,revision,idsKey,props.hierarchyVisible]);
    return <>{availabilityError && <div role="alert">{availabilityError} Use Refresh on the subgrid command bar to retry.</div>}
    <Grid {...props} stateCache={stateCache} selectedRecordIds={selection.activeScope === scope ? selection.selectedIds : undefined}
        onCollapseChild={id => selection.clearBranch(selectionScope([...ancestors,id]))}
        hasChildRows={checker ? record => availability[availabilityId(record.getRecordId())] : undefined} selectionActive={selection.activeScope === scope}
        selectionCountOverride={ancestors.length ? undefined : props.selectionCountOverride ?? selection.selectedCount}
        selectionResetVersion={props.selectionResetVersion}
        setSelectedRecords={ids => selection.select(scope, ids)}
        onSelectAllRecords={request => {
            selection.select(scope, [], true);
            return props.onSelectAllRecords(request);
        }} renderChildRows={props.childRows ? (record, expanded) => {
        const id = record.getRecordId();
        if (ancestors.includes(id) || ancestors.length >= 20) return <div role="alert">Child hierarchy contains a cycle or exceeds 20 levels.</div>;
        return <ChildPanel key={id} parentId={id} root={props} ancestors={[...ancestors,id]} visible={props.hierarchyVisible !== false && expanded}/>;
    } : undefined} /></>;
}

function ChildPanel({parentId, root, ancestors, visible}: {parentId:string; root:HierarchyGridProps; ancestors:string[]; visible:boolean}): React.ReactElement {
    const {clearBranch, states} = React.useContext(SelectionContext)!;
    const scope = selectionScope(ancestors);
    if (!states.has(scope)) states.set(scope, new Map());
    const cache = states.get(scope)!;
    React.useEffect(() => {
        if (!visible) clearBranch(scope);
    }, [clearBranch, scope, visible]);
    const [sorting,setSorting] = useGridState<ChildRequest['sorting']>(cache, 'sorting', () => []);
    const [filtering,setFiltering] = useGridState<FilterExpression | undefined>(cache, 'filtering', () => undefined);
    const [reload,setReload] = React.useState(0);
    const [page,setPage] = useGridState(cache, 'page', () => 0);
    const [data,setData] = useGridState<ChildPage>(cache, 'data', () => ({records:[],hasNextPage:false}));
    const [loading,setLoading] = React.useState(true);
    const [error,setError] = React.useState('');
    const readOnly = !!root.childRows?.readOnly || !root.enableInlineEdit;
    const serial = React.useRef(0);
    const busy = React.useRef(false);
    const pendingPage = React.useRef<Promise<void>>();
    const loader = root.childRows!.loadPage;
    const revision = root.childRows?.revision;
    const request = React.useCallback((pageNumber:number, replace:boolean, generation:number): Promise<void> => {
        if (busy.current) return pendingPage.current ?? Promise.resolve();
        busy.current = true; setLoading(true); setError('');
        const work = (async () => {
        try {
            const previous = cache.get('data') as ChildPage | undefined;
            const result = await loader({parentId,page:pageNumber,sorting,filtering,continuation:replace ? undefined : previous?.continuation});
            if (generation !== serial.current) return;
            const previousIds = new Set(previous?.records.map(record => record.getRecordId()));
            if (!replace && result.records.length && result.records.every(record => previousIds.has(record.getRecordId()))) {
                throw new Error(`The service repeated the previous child records for page ${pageNumber}.`);
            }
            setData(previous => {
                const merged = new Map((replace ? [] : previous.records).map(record => [record.getRecordId(),record]));
                result.records.forEach(record => merged.set(record.getRecordId(),record));
                return {records:[...merged.values()],hasNextPage:result.hasNextPage,continuation:result.continuation};
            });
            setPage(pageNumber);
        } catch (failure) {
            if (generation === serial.current) setError(failure instanceof Error ? failure.message : 'Unable to load child options.');
        } finally {
            if (generation === serial.current) { busy.current = false; setLoading(false); }
        }
        })();
        pendingPage.current = work;
        return work;
    }, [loader,parentId,sorting,filtering,setData,setPage,cache]);
    React.useEffect(() => {
        const generation = ++serial.current;
        const prior = cache.get('query') as {loader: typeof loader; revision: typeof revision; sorting: typeof sorting; filtering: typeof filtering; reload: number} | undefined;
        if (cache.get('page') && prior?.loader === loader && prior.revision === revision && prior.sorting === sorting && prior.filtering === filtering && prior.reload === reload) {
            setLoading(false);
        } else {
            clearBranch(scope);
            cache.set('query', {loader,revision,sorting,filtering,reload});
            busy.current = false; setPage(0); setData({records:[],hasNextPage:false});
            void request(1,true,generation);
        }
        return () => { serial.current = generation + 1; };
    }, [request,revision,reload,cache,loader,sorting,filtering,setData,setPage,clearBranch,scope]);
    const refresh = () => setReload(value => value + 1);
    const records = React.useMemo(() => Object.fromEntries(data.records.map(record => [record.getRecordId(),record])), [data.records]);
    const ids = React.useMemo(() => data.records.map(record => record.getRecordId()), [data.records]);
    const selectAllChildren = async (options?: SelectAllRequest) => {
        const generation = serial.current;
        await pendingPage.current;
        checkSelectionCancelled(options?.signal);
        if (generation !== serial.current) throw new Error('The child view changed. Select the records again.');
        if (error) throw new Error(error);
        busy.current = true; setLoading(true);
        try {
            let current = cache.get('data') as ChildPage;
            let currentPage = cache.get('page') as number;
            let batches = 0;
            while (current.records.length < MAX_SELECTED_RECORDS && current.hasNextPage) {
                checkSelectionCancelled(options?.signal);
                if (++batches > 500) throw new Error('Unable to finish selecting children. Narrow the filters and try again.');
                const result = await loader({parentId,page:currentPage + 1,sorting,filtering,continuation:current.continuation});
                checkSelectionCancelled(options?.signal);
                if (generation !== serial.current) throw new Error('The child view changed. Select the records again.');
                const merged = new Map(current.records.map(row => [row.getRecordId(), row]));
                result.records.forEach(row => merged.set(row.getRecordId(), row));
                if (merged.size === current.records.length && result.hasNextPage) throw new Error('No additional child records were returned. Refresh and try again.');
                current = {...result, records: [...merged.values()]};
                currentPage++;
                setData(current); setPage(currentPage);
            }
            return selectionResult(current.records.map(row => row.getRecordId()), current.hasNextPage);
        } finally {
            if (generation === serial.current) {busy.current = false; setLoading(false);}
        }
    };
    return <section className="cm-child-panel" aria-label={readOnly ? 'Child options (read only)' : 'Child options'} data-selection-disabled={true}
        onClick={event => event.stopPropagation()}>
        {error && <div role="alert">{error} Use Refresh on the subgrid command bar to try again.</div>}
        <HierarchyGrid {...root} width={undefined} height={360} ancestors={ancestors} hierarchyVisible={visible} records={records} sortedRecordIds={ids}
            commandStatus={undefined} commandError={undefined} dataError={error}
            itemsLoading={loading} hasNextPage={!error && data.hasNextPage} hasPreviousPage={false} currentPage={1} totalResultCount={ids.length}
            sorting={sorting} filtering={filtering!} infiniteScroll={true} disableSelection={root.disableSelection}
            enableInlineEdit={!readOnly}
            renderRowActions={undefined} onSelectAllRecords={selectAllChildren}
            onSort={(name,direction) => setSorting(direction === 'none' ? [] : [{name,sortDirection:direction === 'desc' ? 1 : 0}])}
            onFilter={(name,mode,value) => setFiltering(current => updateColumnFilter(current,name,mode,value))}
            onClearAllFilters={() => {setFiltering(undefined); refresh();}}
            loadNextPage={() => request(page + 1,false,serial.current)} loadFirstPage={refresh} loadPreviousPage={refresh}
            onUpdateCell={async (...args) => {await root.onUpdateCell(...args); refresh();}}
            onUpdateLookupCell={root.onUpdateLookupCell ? async (...args) => {await root.onUpdateLookupCell!(...args); refresh();} : undefined}
        />
    </section>;
}
