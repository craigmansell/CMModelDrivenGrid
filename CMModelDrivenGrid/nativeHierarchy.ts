import { FilterExpression } from './filtering';
import { MAX_SELECTED_RECORDS, SelectAllRequest, SelectAllResult, checkSelectionCancelled, selectionResult } from './recordSelection';

type Dataset = ComponentFramework.PropertyTypes.DataSet;
type Row = ComponentFramework.PropertyHelper.DataSetApi.EntityRecord;
type Sort = ComponentFramework.PropertyHelper.DataSetApi.SortStatus;
const normalize = (id: string) => id.replace(/[{}]/g, '').toLowerCase();

export function isHierarchyRoot(row: Row, parentField: string): boolean {
    const parent = row.getValue(parentField);
    return parent == null || parent === '' || (Array.isArray(parent) && parent.length === 0);
}

/** Keep root filters intact, admitting only explicitly selected children to the host dataset. */
export function hierarchyFilter(parentField: string, idField: string, roots: FilterExpression | undefined, children: string[]): FilterExpression {
    const root: FilterExpression = {filterOperator: 0, conditions: [{attributeName: parentField, conditionOperator: 12, value: ''}]};
    if (roots) root.filters = [roots];
    if (!children.length) return root;
    return {filterOperator: 1, conditions: [], filters: [root, {
        filterOperator: 0, conditions: [{attributeName: idField, conditionOperator: 8, value: children.map(normalize)}],
    }]};
}

interface NativeHierarchyHost {
    dataset: () => Dataset;
    load: (change: () => void, firstPage: boolean) => Promise<void>;
    basePageSize: () => number;
    changed: () => void;
    resetRows: () => void;
    resetSelection: () => void;
    error: (message: string) => void;
}

/**
 * Uses only public dataset filtering, sorting, paging and selection APIs. The
 * bound view must include children and the parent lookup. No host records are
 * inserted or modified: Power Apps must retrieve and recognise every target.
 */
export class NativeHierarchySelection {
    rootFilter: FilterExpression | undefined;
    rootSorting: Sort[];
    childIds: string[] = [];
    busy = false;
    selectedCount = 0;
    acceptRows = true;
    private generation = 0;
    private disposed = false;
    private work: Promise<void> = Promise.resolve();

    constructor(readonly parentField: string, readonly idField: string, private host: NativeHierarchyHost) {
        const dataset = host.dataset();
        this.rootFilter = dataset.filtering.getFilter();
        this.rootSorting = dataset.sorting.filter(sort => sort.name !== parentField).map(sort => ({...sort}));
    }

    private current(generation: number): boolean { return !this.disposed && generation === this.generation; }

    private async query(children: string[]): Promise<void> {
        const dataset = this.host.dataset();
        this.childIds = children;
        // Non-null parent lookups sort ahead of roots, so selected children
        // arrive first without changing the relative ordering of root rows.
        await this.host.load(() => {
            dataset.sorting.splice(0, dataset.sorting.length, {name: this.parentField, sortDirection: 1}, ...this.rootSorting);
            dataset.filtering.setFilter(hierarchyFilter(this.parentField, this.idField, this.rootFilter, children));
            dataset.paging.setPageSize(Math.max(this.host.basePageSize(), children.length));
            this.acceptRows = true;
            dataset.refresh();
        }, true);
        if (dataset.error) throw new Error('Power Apps could not load the selected options. Use Refresh on the subgrid command bar.');
    }

    private enqueue(action: (generation: number) => Promise<void>): Promise<void> {
        const generation = ++this.generation;
        this.host.dataset().setSelectedRecordIds([]);
        this.selectedCount = 0;
        this.busy = true;
        this.acceptRows = false;
        this.host.changed();
        this.work = this.work.catch(() => undefined).then(() => this.run(action, generation));
        return this.work;
    }

    private async run(action: (generation: number) => Promise<void>, generation: number): Promise<void> {
            if (!this.current(generation)) return;
            try {
                await action(generation);
            } catch (error) {
                if (!this.current(generation)) return;
                this.host.dataset().setSelectedRecordIds([]);
                this.host.resetSelection();
                this.host.error(error instanceof Error ? error.message : 'Unable to prepare the selected record commands.');
                // Do not keep a failed child target in the host query.
                if (this.childIds.length) {
                    try { await this.query([]); } catch { /* Leave commands safely unselected until Refresh. */ }
                }
            } finally {
                if (this.current(generation)) {
                    this.busy = false;
                    this.acceptRows = true;
                    this.host.changed();
                }
            }
    }

    refreshRoots(filter: FilterExpression | undefined, sorting = this.rootSorting): Promise<void> {
        this.rootFilter = filter;
        this.rootSorting = sorting.map(sort => ({...sort}));
        this.host.resetSelection();
        return this.enqueue(async () => {
            this.host.resetRows();
            await this.query([]);
        });
    }

    select(ids: string[], childIds: string[]): Promise<void> {
        const unique = [...new Set(ids.map(normalize))];
        const children = [...new Set(childIds.map(normalize))];
        if (!unique.length && !this.childIds.length && !this.busy) {
            this.host.dataset().setSelectedRecordIds([]);
            this.selectedCount = 0;
            this.host.changed();
            return Promise.resolve();
        }
        return this.enqueue(async generation => {
            if (unique.length > MAX_SELECTED_RECORDS) throw new Error('Select no more than 1,000 options at a time.');
            if (children.length && (children.length !== unique.length || children.some(id => !unique.includes(id)))) {
                throw new Error('Select options from one hierarchy level at a time.');
            }
            const sameQuery = children.length === this.childIds.length && children.every(id => this.childIds.includes(id));
            if (!sameQuery) await this.query(children);
            else this.acceptRows = true;
            if (!this.current(generation) || !unique.length) return;

            const found = new Map<string, string>();
            const collect = () => {
                for (const [key, row] of Object.entries(this.host.dataset().records)) {
                    const recordId = row.getRecordId();
                    if (unique.includes(normalize(recordId))) found.set(normalize(recordId), key);
                }
            };
            collect();
            let pages = 0;
            let lastPage = this.host.dataset().sortedRecordIds.join('|');
            while (found.size < unique.length && this.host.dataset().paging.hasNextPage && pages++ < 500) {
                await this.host.load(() => this.host.dataset().paging.loadNextPage(), false);
                if (!this.current(generation)) return;
                collect();
                const page = this.host.dataset().sortedRecordIds.join('|');
                if (page === lastPage) throw new Error('Power Apps returned the same page while preparing commands. Use Refresh and try again.');
                lastPage = page;
            }
            if (found.size !== unique.length) throw new Error('Some selected options are no longer available in this grid view. Refresh the subgrid and select them again.');
            if (!this.current(generation)) return;
            await this.publishSelection(unique.map(id => found.get(id)!), generation);
        });
    }

    private async publishSelection(ids: string[], generation: number): Promise<void> {
        if (!this.current(generation)) return;
        const dataset = this.host.dataset();
        const unique = ids.map(normalize);
        dataset.setSelectedRecordIds(ids);
        // Selection state can be delivered on the next model-driven update.
        for (let attempt = 0; attempt < 40; attempt++) {
            if (!this.current(generation)) return;
            const selected = dataset.getSelectedRecordIds().map(normalize);
            if (selected.length === unique.length && unique.every(id => selected.includes(id))) { this.selectedCount = unique.length; return; }
            await new Promise<void>(resolve => setTimeout(resolve, 50));
        }
        throw new Error('Power Apps did not accept this selection. No parent option has been selected instead. Refresh the subgrid and try again.');
    }

    async loadRootSelectionIds(options?: SelectAllRequest): Promise<SelectAllResult> {
        let result: SelectAllResult | undefined;
        let selectionGeneration = -1;
        await this.enqueue(async generation => {
            selectionGeneration = generation;
            checkSelectionCancelled(options?.signal);
            await this.query([]);
            if (!this.current(generation)) return;
            const selected = new Map<string, string>();
            let pages = 0;
            let previous = '';
            while (this.current(generation)) {
                checkSelectionCancelled(options?.signal);
                const dataset = this.host.dataset();
                const signature = dataset.sortedRecordIds.join('|');
                if (pages && signature === previous) throw new Error('Power Apps returned the same page while selecting root options. Use Refresh and try again.');
                previous = signature;
                let additionalRoots = false;
                for (const key of dataset.sortedRecordIds) {
                    const record = dataset.records[key];
                    if (!record || !isHierarchyRoot(record, this.parentField) || selected.has(normalize(record.getRecordId()))) continue;
                    if (selected.size === MAX_SELECTED_RECORDS) {additionalRoots = true; break;}
                    selected.set(normalize(record.getRecordId()), key);
                }
                if (selected.size === MAX_SELECTED_RECORDS || !dataset.paging.hasNextPage) {
                    result = selectionResult([...selected.values()], additionalRoots || dataset.paging.hasNextPage || dataset.paging.totalResultCount > MAX_SELECTED_RECORDS);
                    break;
                }
                if (++pages > 500) throw new Error('Unable to finish selecting root options. Narrow the root filters first.');
                await this.host.load(() => dataset.paging.loadNextPage(), false);
            }
        });
        checkSelectionCancelled(options?.signal);
        if (!result || !this.current(selectionGeneration)) throw new Error('The selection could not finish. Refresh or narrow the filters and try again.');
        return result;
    }

    async selectAllRoots(): Promise<void> {
        const result = await this.loadRootSelectionIds();
        await this.select(result.ids, []);
    }

    externalRefresh(): void {
        if (this.busy || this.disposed) return;
        this.host.resetSelection();
        this.host.resetRows();
        this.selectedCount = 0;
        this.host.dataset().setSelectedRecordIds([]);
        if (this.childIds.length) void this.refreshRoots(this.rootFilter);
    }

    dispose(): void {
        this.disposed = true;
        this.generation++;
    }
}
