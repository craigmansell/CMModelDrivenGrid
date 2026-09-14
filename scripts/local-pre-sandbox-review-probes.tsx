import { act, Simulate } from 'react-dom/test-utils';
import type { GridProps } from '../CMModelDrivenGrid/Grid';
import { CMModelDrivenGrid } from '../CMModelDrivenGrid';
import { updateColumnFilter } from '../CMModelDrivenGrid/filtering';

type Row = ComponentFramework.PropertyHelper.DataSetApi.EntityRecord;
type Column = ComponentFramework.PropertyHelper.DataSetApi.Column;
type Context = ComponentFramework.Context<import('../CMModelDrivenGrid/generated/ManifestTypes').IInputs>;

/** Permanent release regressions. All host calls are local mocks, never Dataverse. */
export async function runPreSandboxReviewProbes(harness: {
    container: HTMLElement;
    resources: ComponentFramework.Resources;
    renderGrid: (props: Partial<GridProps>) => void;
    updateGrid: (props: Partial<GridProps>) => void;
    unmount: () => void;
}): Promise<void> {
    const failures: string[] = [];
    const check = (condition: boolean, message: string) => { if (!condition) throw new Error(message); };
    const probe = async (name: string, run: () => Promise<void>) => {
        try { await run(); console.log(`PASS: ${name}`); }
        catch (error) { const detail = error instanceof Error ? error.message : String(error); failures.push(name); console.error(`FAIL: ${name}: ${detail}`); }
        finally { await act(async () => { harness.unmount(); }); }
    };
    const row = {getRecordId: () => 'a', getValue: () => 'Alpha', getFormattedValue: () => 'Alpha'} as unknown as Row;
    const columns = ['name', 'second', 'third'].map((name, order) => ({name, displayName: name, dataType: 'SingleLine.Text', order})) as Column[];
    const key = async (rowIndex: number, col: number, key: string, modifiers = {}) => {
        const cell = harness.container.querySelector<HTMLElement>(`[data-cm-cell-row="${rowIndex}"][data-cm-cell-col="${col}"]`)!;
        check(!!cell, `Keyboard source ${rowIndex},${col} must exist.`);
        await act(async () => {cell.focus(); Simulate.keyDown(cell, {key, ...modifiers});});
        await act(async () => {await new Promise(resolve => setTimeout(resolve, 40));});
    };
    const focusIs = (rowIndex: number, col: number) => document.activeElement?.getAttribute('data-cm-cell-row') === String(rowIndex) && document.activeElement?.getAttribute('data-cm-cell-col') === String(col);

    await probe('Compact End stays on the last rendered field', async () => {
        await act(async () => { harness.renderGrid({columns, records: {a: row}, sortedRecordIds: ['a'], features: {reflow: 'list', listColumns: 2}}); });
        const cell = harness.container.querySelector<HTMLElement>('[data-cm-cell-row="0"][data-cm-cell-col="0"]')!;
        check(!!cell, 'Compact starting cell must exist.');
        await act(async () => { cell.focus(); Simulate.keyDown(cell, {key: 'End'}); });
        await act(async () => { await new Promise(resolve => setTimeout(resolve, 150)); });
        const actual = document.activeElement?.getAttribute('data-cm-cell-col');
        check(actual === '1', `Expected visible column 1; focus remained at column ${actual}.`);
    });

    await probe('Compact keyboard ranges copy only displayed fields and reset on reflow', async () => {
        const makeRow = (id: string) => ({getRecordId: () => id, getValue: (name: string) => `${id}:${name}`, getFormattedValue: (name: string) => `${id}:${name}`}) as unknown as Row;
        const records = {a: makeRow('a'), b: makeRow('b')};
        const copies: string[] = [];
        Object.defineProperty(navigator, 'clipboard', {value: {writeText: async (text: string) => {copies.push(text);}}, configurable: true});
        const copy = async () => {await act(async () => {document.activeElement?.dispatchEvent(new window.KeyboardEvent('keydown', {key:'c', ctrlKey:true, bubbles:true}));});};
        await act(async () => {harness.renderGrid({columns, records, sortedRecordIds: ['a','b'], features: {reflow:'list', listColumns:2}});});
        await key(0,0,'End'); check(focusIs(0,1), 'End must reach the final card field.');
        await key(0,1,'ArrowRight'); check(focusIs(0,1), 'Right at the card boundary must not target a hidden field.');
        await key(0,1,'Home'); check(focusIs(0,0), 'Home must return to the first card field.');
        await key(0,0,'End',{ctrlKey:true,shiftKey:true}); check(focusIs(1,1), 'Ctrl+Shift+End must reach the final visible loaded cell.');
        await copy(); check(copies[0] === 'a:name\ta:second\r\nb:name\tb:second', 'Compact copying must not include the undisplayed third column.');
        await act(async () => {harness.updateGrid({features:{reflow:'grid'}});});
        await copy(); check(copies.length === 1, 'Reflow must clear the previous compact range.');
        await key(0,0,'End'); check(focusIs(0,2), 'The same mounted component must use all columns in table mode.');
        await key(0,2,'ArrowLeft'); check(focusIs(0,1), 'Table arrow navigation must remain intact.');
        await act(async () => {harness.updateGrid({features:{reflow:'list',listColumns:1}});});
        await key(0,0,'End'); check(focusIs(0,0), 'A changed list field limit must immediately update keyboard bounds.');
        await key(0,0,'ArrowDown'); check(focusIs(1,0), 'Down must stay within the rendered card fields.');
        await key(1,0,'ArrowUp'); check(focusIs(0,0), 'Up must return to the previous card.');
    });

    await probe('Reflow cancels a keyboard target waiting on the next page', async () => {
        let requests = 0;
        await act(async () => {harness.renderGrid({columns, records:{a:row},sortedRecordIds:['a'],features:{reflow:'grid'},infiniteScroll:true,hasNextPage:true,loadNextPage:()=>{requests++;harness.updateGrid({itemsLoading:true});}});});
        await key(0,2,'ArrowDown'); check(requests === 1, 'Crossing the loaded boundary should request one page.');
        const second = {...row,getRecordId:()=> 'b'};
        await act(async () => {harness.updateGrid({features:{reflow:'list',listColumns:1},records:{a:row,b:second},sortedRecordIds:['a','b'],hasNextPage:false,itemsLoading:false});});
        await key(0,0,'Home');
        await act(async () => {await new Promise(resolve => setTimeout(resolve, 150));});
        check(focusIs(0,0), 'A stale table target must not steal focus after reflow and page completion.');
        check(!harness.container.querySelector('[data-cm-cell-col="2"]'), 'The old target column must really be absent.');
    });

    await probe('Rejected ordinary filter restores the previous state and releases loading', async () => {
        const controller = new CMModelDrivenGrid();
        let physical: ComponentFramework.PropertyHelper.DataSetApi.FilterExpression | undefined;
        let refreshes = 0;
        const dataset = {
            columns: [columns[0]], records: {a: row}, sortedRecordIds: ['a'], sorting: [], loading: false,
            getTargetEntityType: () => 'account', getViewId: () => 'view-1', setSelectedRecordIds: () => undefined,
            filtering: {getFilter: () => physical, setFilter: (_filter?: ComponentFramework.PropertyHelper.DataSetApi.FilterExpression): void => {throw new Error('Injected host filter rejection');}, clearFilter: () => {physical = undefined;}},
            paging: {hasNextPage: false, hasPreviousPage: false, totalResultCount: 1}, refresh: () => {refreshes++;},
        };
        controller.context = {mode: {allocatedWidth: 900, allocatedHeight: 500, isControlDisabled: false}, parameters: {
            records: dataset, SubGridHeight: {raw: 500}, HighlightValue: {raw: null}, HighlightColor: {raw: null},
            EnableLookupLinks: {raw: false}, EnableInlineEdit: {raw: true}, InfiniteScroll: {raw: true}, EnableJumpBar: {raw: true},
        }} as unknown as Context;
        controller.container = harness.container as HTMLDivElement;
        controller.resources = harness.resources;
        controller.records = {a: row}; controller.sortedRecordsIds = ['a'];
        const access = controller as unknown as {renderGrid: () => void; getRootFilter: () => unknown};
        await act(async () => { access.renderGrid(); });
        const letter = Array.from(harness.container.querySelectorAll<HTMLButtonElement>('.cm-grid-jump-bar button')).find(button => button.textContent === 'A')!;
        check(!!letter && !letter.disabled, 'The real controller must render an enabled alphabet filter.');
        for (let attempt=0;attempt<2;attempt++) {
            await act(async () => { letter.click(); });
            await act(async () => { await new Promise(resolve => setTimeout(resolve, 40)); });
            const overlay = !!harness.container.querySelector('.cm-grid-loading-overlay');
            const logical = access.getRootFilter();
            check(!overlay && logical === physical, JSON.stringify({attempt,loadingOverlay: overlay, logicalFilter: logical, physicalFilter: physical ?? null, refreshes}));
        }
        dataset.filtering.setFilter = (filter?: ComponentFramework.PropertyHelper.DataSetApi.FilterExpression) => {physical=filter;};
        dataset.refresh = () => {refreshes++;dataset.loading=true;access.renderGrid();};
        await act(async () => {letter.click();});
        check(!!harness.container.querySelector('.cm-grid-loading-overlay'), 'The successful retry must still show genuine loading.');
        await act(async () => {dataset.loading=false;access.renderGrid();});
        check(!harness.container.querySelector('.cm-grid-loading-overlay'), 'An unchanged-result retry must release loading.');
        check(refreshes === 1 && JSON.stringify(physical)?.includes('A%'), 'Retry must apply the requested filter once.');
    });

    await probe('Filter/sort rejection restores nested state; failed rollback blocks selection', async () => {
        const controller = new CMModelDrivenGrid();
        const previous = updateColumnFilter(updateColumnFilter(undefined,'name','equals','Alpha'),'second','in',['One','Two'])!;
        let physical = previous;
        let selected: string[] = [];
        let rejectAll = false;
        const originalSorting = [{name:'second',sortDirection:0},{name:'name',sortDirection:0}];
        const dataset = {columns, records:{a:row},sortedRecordIds:['a'],sorting:originalSorting.map(sort=>({...sort})),loading:false,
            getTargetEntityType:()=> 'account',getViewId:()=> 'view-1',setSelectedRecordIds:(ids:string[])=>{selected=ids;},
            filtering:{getFilter:()=>physical,setFilter:(filter:typeof previous)=>{if(rejectAll)throw new Error('Rollback denied');physical=filter;},clearFilter:()=>undefined},
            paging:{hasNextPage:false,hasPreviousPage:false,totalResultCount:1},refresh:():void=>{throw new Error('Refresh rejected');}};
        controller.context={mode:{allocatedWidth:900,allocatedHeight:500,isControlDisabled:false},parameters:{records:dataset,
            SubGridHeight:{raw:500},HighlightValue:{raw:null},HighlightColor:{raw:null},EnableLookupLinks:{raw:false},EnableInlineEdit:{raw:true},InfiniteScroll:{raw:true}}} as unknown as Context;
        controller.container=harness.container as HTMLDivElement;controller.resources=harness.resources;controller.records={a:row};controller.sortedRecordsIds=['a'];
        const access=controller as unknown as {renderGrid:()=>void;getRootFilter:()=>unknown;onSelectAllRecords:()=>Promise<unknown>;rootFilterRecoveryRequired:boolean};
        await act(async()=>{access.renderGrid();controller.onSort('name','desc');});
        check(JSON.stringify(dataset.sorting) === JSON.stringify(originalSorting), 'Rejected refresh must restore both sort priorities.');
        check(JSON.stringify(physical) === JSON.stringify(previous) && JSON.stringify(access.getRootFilter()) === JSON.stringify(previous), 'Nested physical and logical filters must both be restored.');
        rejectAll=true;
        await act(async()=>{controller.onFilter('name','equals','Beta');});
        check(access.rootFilterRecoveryRequired, 'A rejected rollback must demand recovery, not silently enable commands.');
        check(!harness.container.querySelector('[role="checkbox"]'), 'Failed recovery must disable grid row selection.');
        controller.setSelectedRecords(['a']);check(selected.length===0,'Unsafe query state must not publish selection IDs.');
        let rejected=false;try {await access.onSelectAllRecords();} catch {rejected=true;}
        check(rejected, 'Select All must reject an unrecovered query.');
        rejectAll=false;dataset.refresh=()=>undefined;
        await act(async()=>{controller.onFilter('name','equals','Alpha');access.renderGrid();});
        check(!access.rootFilterRecoveryRequired && !!harness.container.querySelector('[role="checkbox"]'), 'An accepted replacement query must permit selection again.');
    });

    await probe('Recovery waits for a successful refresh and follows the accepted view filter', async () => {
        const controller = new CMModelDrivenGrid();
        let physical = updateColumnFilter(undefined,'name','equals','Alpha')!;
        const dataset = {columns,records:{a:row},sortedRecordIds:['a'],sorting:[],loading:false,error:false,
            getTargetEntityType:()=> 'account',getViewId:()=> 'view-1',setSelectedRecordIds:()=>undefined,
            filtering:{getFilter:()=>physical,setFilter:():void=>{throw new Error('Rollback denied');},clearFilter:():void=>{throw new Error('Rollback denied');}},
            paging:{hasNextPage:false,hasPreviousPage:false,totalResultCount:1},refresh:()=>undefined};
        controller.context={updatedProperties:[],mode:{allocatedWidth:900,allocatedHeight:500,isControlDisabled:false},parameters:{records:dataset,
            SubGridHeight:{raw:500},HighlightValue:{raw:null},HighlightColor:{raw:null},EnableLookupLinks:{raw:false},EnableInlineEdit:{raw:true},InfiniteScroll:{raw:true}}} as unknown as Context;
        controller.container=harness.container as HTMLDivElement;controller.resources=harness.resources;
        const access=controller as unknown as {getRootFilter:()=>unknown;rootFilterRecoveryRequired:boolean;columnOptionsFetchedFor:string;lookupNavPropertyFetchedFor:string};
        access.columnOptionsFetchedFor='account';access.lookupNavPropertyFetchedFor='account';
        const update = async () => {await act(async()=>{controller.updateView(controller.context);});};
        await update();
        await act(async()=>{controller.onFilter('name','equals','Beta');});
        check(access.rootFilterRecoveryRequired,'The rejected rollback must require recovery.');
        physical=updateColumnFilter(undefined,'name','equals','Host accepted')!;
        dataset.loading=true;await update();dataset.loading=false;dataset.error=true;await update();
        check(access.rootFilterRecoveryRequired,'An unsuccessful host refresh must not release the command safety gate.');
        dataset.loading=true;dataset.error=false;await update();dataset.loading=false;await update();
        check(!access.rootFilterRecoveryRequired && access.getRootFilter()===physical,'Successful refresh must display the actual accepted host filter.');
        await act(async()=>{controller.onFilter('name','equals','Gamma');});
        check(access.rootFilterRecoveryRequired,'A subsequent rejected rollback must still block selection.');
        physical=updateColumnFilter(undefined,'name','equals','New view')!;dataset.getViewId=()=> 'view-2';await update();
        check(!access.rootFilterRecoveryRequired && access.getRootFilter()===physical,'An independent new view must not inherit the old query recovery state.');
    });

    await probe('Lookup save rechecks column read-only state after metadata', async () => {
        const controller = new CMModelDrivenGrid();
        const column = {name: 'lookup', dataType: 'Lookup.Simple', isReadOnly: false};
        let writes = 0;
        controller.context = {mode: {isControlDisabled: false}, parameters: {EnableInlineEdit: {raw: true}, records: {
            getTargetEntityType: () => 'account', columns: [column], refresh: () => undefined,
        }}, webAPI: {updateRecord: async () => {writes++;}}, navigation: {openAlertDialog: async () => undefined}} as unknown as Context;
        Object.assign(controller, {columnEditabilityEntity:'account', columnEditability:{lookup:true}});
        const access = controller as unknown as {getEntityMeta: () => Promise<{entitySetName: string}>; resolveLookupBindNames: () => Promise<string[]>};
        access.getEntityMeta = async () => ({entitySetName: 'contacts'});
        access.resolveLookupBindNames = async () => {column.isReadOnly = true; return ['lookup'];};
        try {await controller.onUpdateLookupCell('00000000-0000-4000-8000-000000000001', 'lookup', '00000000-0000-4000-8000-000000000002', 'contact');} catch { /* Rejection is the expected safe result. */ }
        check(writes === 0, `Expected zero writes after the column became read only; observed ${writes}.`);
    });

    for (const change of ['isReadOnly','readOnly','disableEditing','removed','alias','type','view','table','retry'] as const) {
        await probe(`Lookup blocks a late ${change} change, including clear/retry`, async () => {
            for (const clear of [false,true]) {
                const controller=new CMModelDrivenGrid();
                const column={name:'lookup',dataType:'Lookup.Simple',alias:'lookup',isReadOnly:false,readOnly:false,disableEditing:false};
                const dataset={getTargetEntityType:()=> 'account',getViewId:()=> 'view-1',columns:[column],refresh:()=>undefined};
                let writes=0;
                controller.context={mode:{isControlDisabled:false},parameters:{EnableInlineEdit:{raw:true},records:dataset},
                    webAPI:{updateRecord:async()=>{writes++;if(change==='retry'){column.isReadOnly=true;throw new Error("An undeclared property 'lookup' was found in the payload.");}}},navigation:{openAlertDialog:async()=>undefined}} as unknown as Context;
                Object.assign(controller, {columnEditabilityEntity:'account', columnEditability:{lookup:true}});
                const access=controller as unknown as {getEntityMeta:()=>Promise<{entitySetName:string}>;resolveLookupBindNames:()=>Promise<string[]>};
                access.getEntityMeta=async()=>({entitySetName:'contacts'});
                access.resolveLookupBindNames=async()=>{
                    if(change==='removed')dataset.columns=[];
                    else if(change==='alias')column.alias='joined.lookup';
                    else if(change==='type')column.dataType='SingleLine.Text';
                    else if(change==='view')dataset.getViewId=()=> 'view-2';
                    else if(change==='table')dataset.getTargetEntityType=()=> 'contact';
                    else if(change!=='retry')column[change]=true;
                    return ['lookup','validLookup'];
                };
                let rejected=false;
                try {await controller.onUpdateLookupCell('00000000-0000-4000-8000-000000000001','lookup',clear?'':'00000000-0000-4000-8000-000000000002','contact');} catch {rejected=true;}
                check(rejected && writes===(change==='retry'?1:0), `${change}, clear=${clear}: no update may follow the state change (observed ${writes}).`);
            }
        });
    }

    if (failures.length) throw new Error(`${failures.length} pre-sandbox regressions failed: ${failures.join('; ')}`);
}
