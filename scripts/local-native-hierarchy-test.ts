import assert from 'node:assert/strict';
import { NativeHierarchySelection, hierarchyFilter, isHierarchyRoot } from '../CMModelDrivenGrid/nativeHierarchy';
import { FilterExpression, updateColumnFilter } from '../CMModelDrivenGrid/filtering';
import { CMModelDrivenGrid } from '../CMModelDrivenGrid';

type Dataset = ComponentFramework.PropertyTypes.DataSet;
type Row = ComponentFramework.PropertyHelper.DataSetApi.EntityRecord;
const parent = 'cm_parenttemplateoption';
const primary = 'cm_itemtemplateoptionid';
const rootId = '10000000-0000-4000-8000-000000000001';
const childA = '20000000-0000-4000-8000-000000000001';
const childB = '20000000-0000-4000-8000-000000000002';
const row = (id: string, parentId?: string): Row => ({getRecordId: () => id, getValue: () => parentId ? [{id: parentId}] : null} as unknown as Row);

function testHost(rejectSelection = false) {
    let filter: FilterExpression | undefined;
    let selected: string[] = [];
    let loads = 0;
    let hold: (() => Promise<void>) | undefined;
    let size = 2;
    let resets = 0;
    const errors: string[] = [];
    const all: Record<string, Row> = {[rootId]: row(rootId), [childA]: row(childA, rootId), [childB]: row(childB, rootId)};
    const admitted = () => (filter?.filters?.[1]?.conditions[0]?.value as string[] | undefined) ?? [];
    const dataset = {
        loading: false, error: false, columns: [], sorting: [{name:'cm_order',sortDirection:0}],
        records: {[rootId]: all[rootId]}, sortedRecordIds: [rootId],
        filtering: {getFilter: () => filter, setFilter: (next: FilterExpression) => {filter = next;}},
        paging: {hasNextPage:false, hasPreviousPage:false, totalResultCount:1, setPageSize:(next:number) => {size = next;}, loadNextPage:() => undefined},
        setSelectedRecordIds: (ids: string[]) => {selected = rejectSelection ? [] : ids.filter(id => !!dataset.records[id]);},
        getSelectedRecordIds: () => [...selected],
        refresh: () => {
            const ids = [...admitted(), rootId];
            dataset.records = Object.fromEntries(ids.map(id => [id, all[id]]));
            dataset.sortedRecordIds = ids;
            dataset.paging.totalResultCount = ids.length;
        },
    } as unknown as Dataset;
    const control = new NativeHierarchySelection(parent, primary, {
        dataset: () => dataset, basePageSize: () => 2,
        load: async change => {loads++; if(hold) await hold(); change();},
        changed: () => undefined, resetRows: () => undefined,
        resetSelection: () => {resets++;}, error: message => errors.push(message),
    });
    return {control, dataset, errors, selected: () => selected, loads: () => loads, size: () => size, resets: () => resets, hold: (value: typeof hold) => {hold = value;}};
}

async function run() {
    await testBoundedRootSelection();
    const roots = updateColumnFilter(undefined, 'cm_unitcost', 'greaterThan', '90');
    const composed = hierarchyFilter(parent, primary, roots, [childA]);
    assert.equal(composed.filterOperator, 1);
    assert.deepEqual(composed.filters?.[0].filters, [roots], 'Root conditions must remain inside the root-only branch.');
    assert.equal(composed.filters?.[0].conditions[0].conditionOperator, 12);
    assert.deepEqual(composed.filters?.[1].conditions[0].value, [childA]);
    assert.equal(isHierarchyRoot(row(rootId), parent), true);
    assert.equal(isHierarchyRoot(row(childA, rootId), parent), false);

    const host = testHost();
    await host.control.refreshRoots(roots);
    host.dataset.setSelectedRecordIds([childA]);
    assert.deepEqual(host.selected(), [], 'Strict host rejects a separately fetched child before it is loaded.');
    await host.control.select([childA], [childA]);
    assert.deepEqual(host.selected(), [childA]);
    assert.equal(host.control.selectedCount, 1);
    assert.deepEqual(host.control.rootFilter, roots);
    assert.deepEqual(host.dataset.sorting, [{name:parent,sortDirection:1},{name:'cm_order',sortDirection:0}]);
    assert.equal(host.errors.length, 0);
    await host.control.select([], []);
    assert.equal(host.control.selectedCount, 0, 'Deselecting must clear the confirmed footer count.');

    await host.control.select([childA, childB], [childA, childB]);
    assert.deepEqual(host.selected(), [childA, childB]);
    assert.equal(host.control.selectedCount, 2);
    assert.equal(host.size(), 2);
    await host.control.select([rootId], []);
    assert.deepEqual(host.selected(), [rootId], 'Switching back must target the root, not the old child.');
    assert.deepEqual(host.control.childIds, []);
    await host.control.refreshRoots(undefined, [{name:'cm_order',sortDirection:1}]);
    assert.equal(host.control.rootFilter, undefined, 'Clear filters must really clear the user filter.');
    assert.equal(host.control.selectedCount, 0);

    await host.control.select([childB], [childB]);
    host.control.externalRefresh();
    await host.control.select([], []);
    assert.deepEqual(host.selected(), []);
    assert.deepEqual(host.control.childIds, []);
    await host.control.selectAllRoots();
    assert.deepEqual(host.selected(), [rootId], 'Root Select All must never include children.');
    assert.equal(host.control.selectedCount, 1);

    const raced = testHost();
    let release!: () => void;
    raced.hold(() => new Promise<void>(resolve => {release = resolve;}));
    const first = raced.control.select([childA], [childA]);
    await Promise.resolve(); await Promise.resolve();
    const second = raced.control.select([childB], [childB]);
    raced.hold(undefined); release();
    await Promise.all([first, second]);
    assert.deepEqual(raced.selected(), [childB], 'Late first selection must never overwrite the newer child.');

    const rejected = testHost(true);
    await rejected.control.select([childA], [childA]);
    assert.deepEqual(rejected.selected(), []);
    assert.deepEqual(rejected.control.childIds, []);
    assert.equal(rejected.control.selectedCount, 0);
    assert.match(rejected.errors[0], /did not accept/);
    assert.ok(rejected.resets() > 0);
    const mixed = testHost();
    await mixed.control.select([rootId, childA], [childA]);
    assert.match(mixed.errors[0], /one hierarchy level/);
    assert.deepEqual(mixed.selected(), []);
    const many = testHost();
    await many.control.select(Array.from({length:1001}, (_,index) => `id-${index}`), []);
    assert.match(many.errors[0], /1,000/);
    assert.equal(many.loads(), 0, 'Oversized selections must fail before loading or executing any command.');

    const missing = testHost();
    await missing.control.select(['30000000-0000-4000-8000-000000000001'], []);
    assert.match(missing.errors[0], /no longer available/);
    assert.deepEqual(missing.selected(), []);

    const projection = testHost();
    const component = new CMModelDrivenGrid();
    const rootTwo = '10000000-0000-4000-8000-000000000002';
    const testable = component as unknown as {nativeHierarchy: NativeHierarchySelection; renderGrid: () => void};
    testable.renderGrid = () => undefined;
    testable.nativeHierarchy = new NativeHierarchySelection(parent, primary, {
        dataset: () => projection.dataset, basePageSize: () => 50,
        load: async change => {change();}, changed: () => undefined,
        resetRows: () => {component.accumulatedRecords={};component.accumulatedSortedIds=[];},
        resetSelection: () => undefined, error: message => {throw new Error(message);},
    });
    component.context = {parameters:{records:projection.dataset,InfiniteScroll:{raw:true},MaxRows:{raw:null}},updatedProperties:[],mode:{isVisible:true,allocatedWidth:800,allocatedHeight:500}} as unknown as ComponentFramework.Context<import('../CMModelDrivenGrid/generated/ManifestTypes').IInputs>;
    component.accumulatedRecords = {[rootId]:row(rootId),[rootTwo]:row(rootTwo)};
    component.accumulatedSortedIds = [rootId,rootTwo];
    projection.dataset.records = {[rootId]:row(rootId),[childA]:row(childA,rootId)};
    projection.dataset.sortedRecordIds = [childA,rootId];
    component.updateView(component.context);
    assert.deepEqual(component.sortedRecordsIds,[rootId,rootTwo], 'A native selection update must retain previously loaded root rows.');
    assert.equal(component.records[childA],undefined,'Host-loaded children must never appear twice as root rows.');
    const previousRecords = component.records;
    const previousIds = component.sortedRecordsIds;
    const rootThree = '10000000-0000-4000-8000-000000000003';
    projection.dataset.records = {[rootThree]:row(rootThree)};
    projection.dataset.sortedRecordIds = [rootThree];
    projection.dataset.paging.hasPreviousPage = true;
    component.updateView(component.context);
    assert.notEqual(component.records,previousRecords,'Appending a native hierarchy page must publish a new record-map identity to React.');
    assert.notEqual(component.sortedRecordsIds,previousIds,'Appending a page must publish new ordered IDs to React.');
    assert.deepEqual(previousIds,[rootId,rootTwo],'Previously rendered row IDs must never be mutated in place.');
    assert.deepEqual(component.sortedRecordsIds,[rootId,rootTwo,rootThree]);
    projection.dataset.paging.hasPreviousPage = false;
    projection.dataset.loading = true; component.updateView(component.context);
    projection.dataset.records = {[rootId]:row(rootId)};
    projection.dataset.sortedRecordIds = [rootId];
    projection.dataset.loading = false; component.updateView(component.context);
    assert.deepEqual(component.sortedRecordsIds,[rootId], 'External Refresh must discard stale cached roots.');
    // Model-driven hosts can append pages while hasPreviousPage remains false
    // and can omit dataset/records from updatedProperties altogether.
    const largeIds = Array.from({length:10000},(_,index) => `40000000-0000-4000-8000-${String(index + 1).padStart(12,'0')}`);
    (globalThis as unknown as {window: unknown}).window = {setTimeout:globalThis.setTimeout,clearTimeout:globalThis.clearTimeout};
    let offset = 0, requests = 0;
    const publishBatch = () => {
        const ids = largeIds.slice(offset, offset + 50);
        projection.dataset.records = Object.fromEntries(ids.map(id => [id,row(id)]));
        projection.dataset.sortedRecordIds = ids;
        offset += ids.length;
        projection.dataset.paging.hasPreviousPage = false;
        projection.dataset.paging.hasNextPage = offset < largeIds.length;
        projection.dataset.loading = false;
        component.updateView(component.context);
    };
    component.accumulatedRecords = {}; component.accumulatedSortedIds = []; component.currentPage = 1;
    publishBatch();
    projection.dataset.paging.loadNextPage = () => {
        requests++;
        projection.dataset.loading = true; component.updateView(component.context);
        queueMicrotask(publishBatch);
    };
    while (projection.dataset.paging.hasNextPage) {
        const priorIds = component.sortedRecordsIds;
        const beforeLength = priorIds.length;
        const pending = component.loadNextPage();
        assert.equal(pending, component.loadNextPage(), 'Overlapping scroll requests must share one in-flight request.');
        await pending;
        assert.equal(priorIds.length,beforeLength,'Paging must not mutate the previously rendered page.');
        assert.equal(component.sortedRecordsIds.length,beforeLength+50);
    }
    assert.equal(requests,199);
    assert.equal(component.currentPage,200);
    assert.deepEqual(component.sortedRecordsIds,largeIds,'All 10,000 rows must be reachable once, in order, in 50-row batches.');
    console.log('10,000-row native paging: 200 batches, no duplicates, no missing rows, no overlapping host calls.');
    console.log('Native hierarchy query, strict-host selection, filtering, switching, cancellation and failure checks passed.');
}

async function testBoundedRootSelection(): Promise<void> {
    for (const total of [0,17,1000,2000]) {
        const ids = Array.from({length:total},(_,i)=>`50000000-0000-4000-8000-${String(total-i).padStart(12,'0')}`);
        let selected:string[] = [], page=0, calls=0;
        const errors:string[] = [];
        const dataset = {
            records:{},sortedRecordIds:[],sorting:[],error:false,
            filtering:{getFilter:()=>undefined,setFilter:()=>undefined},
            paging:{hasNextPage:false,totalResultCount:-1,setPageSize:()=>undefined,loadNextPage:()=>{page++;load(false);}},
            refresh:()=>{page=1;load(true);},
            setSelectedRecordIds:(next:string[])=>{selected=next.filter(id=>!!dataset.records[id]);},
            getSelectedRecordIds:()=>selected,
        } as unknown as Dataset;
        const load=(replace:boolean) => {
            calls++;
            const batch=ids.slice((page-1)*75,page*75);
            dataset.sortedRecordIds=batch;
            dataset.records={...(replace?{}:dataset.records),...Object.fromEntries(batch.map(id=>[id,row(id)]))};
            dataset.paging.hasNextPage=page*75<total;
        };
        const control=new NativeHierarchySelection(parent,primary,{
            dataset:()=>dataset,basePageSize:()=>75,load:async change=>{change();},
            changed:()=>undefined,resetRows:()=>undefined,resetSelection:()=>undefined,error:message=>errors.push(message),
        });
        const result=await control.loadRootSelectionIds();
        assert.deepEqual(result.ids,ids.slice(0,1000),'Select All preserves current query order and stops at 1,000 even with an unknown total.');
        assert.equal(result.limitReached,total>1000);
        assert.equal(calls,Math.max(1,Math.ceil(Math.min(total,1000)/75)),'No later pages should be requested after reaching the selection cap.');
        assert.deepEqual(selected,[],'Collecting IDs must not publish a partial selection while paging.');
        await control.select(result.ids,[]);
        assert.deepEqual(selected,result.ids,'The strict native host must accept all collected root IDs.');
        assert.equal(control.selectedCount,Math.min(total,1000));
        assert.deepEqual(errors,[]);
    }
    console.log('Native Select All cap/order/unknown-total checks passed (0, 17, 1,000 and 2,000 roots).');
}
void run().catch(error => {console.error(error); process.exitCode = 1;});
