import assert from 'node:assert/strict';
import { FilterExpression, updateColumnFilter } from '../CMModelDrivenGrid/filtering';
import { buildFilterResolutionFetch, collectFilterCandidates, filterForIds, needsServerFilter, supportedFilterSuperset, resolveFilterIds } from '../CMModelDrivenGrid/datasetFiltering';
import { evaluateDemoFilter } from './demo-filter-evaluator';
import { CMModelDrivenGrid } from '../CMModelDrivenGrid';

type Column = ComponentFramework.PropertyHelper.DataSetApi.Column;
const columns=[{name:'name',dataType:'SingleLine.Text'},{name:'amount',dataType:'Decimal'},{name:'lookup',dataType:'Lookup.Simple'}] as Column[];
const id=(n:number) => `00000000-0000-4000-8000-${String(n).padStart(12,'0')}`;
const view='<fetch><entity name="account"><attribute name="name"/><filter type="or"><condition attribute="statecode" operator="eq" value="0"/><condition attribute="name" operator="eq" value="Special"/></filter></entity></fetch>';
const options={getValue:(row:Record<string,unknown>,field:string)=>row[field],columnKinds:{name:'text' as const,amount:'number' as const,lookup:'lookup' as const}};
const matches=(row:Record<string,unknown>,filter:FilterExpression|undefined) => evaluateDemoFilter(row,filter,options);

export async function testDatasetFiltering(): Promise<void> {
    for (const field of ['name','amount','lookup']) {
        const logical=updateColumnFilter(undefined,field,'containsData');
        const compiled=supportedFilterSuperset(logical,columns);
        assert.equal(needsServerFilter(logical),true,'Presence must use Dataverse NOT NULL, not the host EQ/NE disjunction that includes blanks.');
        assert.equal(compiled,undefined,'Presence candidate scope must include all bound rows before server matching.');
        const query=new DOMParser().parseFromString(buildFilterResolutionFetch(view,'account','accountid',logical!,[id(1)]),'text/xml');
        assert.ok(query.querySelector(`condition[attribute="${field}"][operator="not-null"]`));
        assert.equal(query.querySelector('condition[operator="not-null"]')?.hasAttribute('value'),false);
    }
    const negative=updateColumnFilter(undefined,'name','doesNotContain','alpha')!;
    const combined=updateColumnFilter(negative,'amount','greaterThan','10')!;
    assert.equal(needsServerFilter(combined),true);
    const superset=supportedFilterSuperset(combined,columns);
    assert.equal(matches({name:'alpha',amount:12},superset),true);
    assert.equal(matches({name:'beta',amount:9},superset),false);
    const or:FilterExpression={filterOperator:1,conditions:[],filters:[negative,updateColumnFilter(undefined,'amount','equals','2')!]};
    assert.equal(supportedFilterSuperset(or,columns),undefined,'An OR with an unknown predicate must remain a superset, including potential matches.');
    const and:FilterExpression={filterOperator:0,conditions:[],filters:[or,updateColumnFilter(undefined,'amount','greaterThan','10')!]};
    assert.equal(matches({name:'beta',amount:12},supportedFilterSuperset(and,columns)),true);
    assert.equal(matches({name:'beta',amount:2},supportedFilterSuperset(and,columns)),false);

    const query=new DOMParser().parseFromString(buildFilterResolutionFetch(view,'account','accountid',negative,[id(1),id(2)]),'text/xml');
    assert.ok(query.querySelector('condition[operator="not-like"][value="%alpha%"]'));
    assert.deepEqual(Array.from(query.querySelectorAll('condition[attribute="accountid"] value')).map(node=>node.textContent),[id(1),id(2)]);
    assert.ok(query.querySelector('filter[type="or"]'),'The selected view must be retained.');
    assert.equal(query.querySelectorAll('attribute').length,1,'Only record identities should be retrieved.');
    assert.equal(query.querySelector('attribute')?.getAttribute('name'),'accountid');

    let pageIndex=0;
    assert.deepEqual(await collectFilterCandidates({page:()=>({ids:pageIndex ? [id(2),id(3)] : [id(1),id(2)],hasNextPage:pageIndex===0}),next:async()=>{pageIndex++;},assertCurrent:()=>undefined}),[id(1),id(2),id(3)]);
    await assert.rejects(collectFilterCandidates({page:()=>({ids:[id(1)],hasNextPage:true}),next:async()=>undefined,assertCurrent:()=>undefined}),/repeated/);
    await assert.rejects(collectFilterCandidates({page:()=>({ids:Array.from({length:5001},(_,n)=>id(n+1)),hasNextPage:false}),next:async()=>undefined,assertCurrent:()=>undefined}),/5,000/);
    await assert.rejects(collectFilterCandidates({page:()=>({ids:[],hasNextPage:false,error:true}),next:async()=>undefined,assertCurrent:()=>undefined}),/could not load/);

    let apiCalls=0;
    const api={retrieveMultipleRecords:async(_entity:string,options:string)=>{
        apiCalls++;
        const xml=new DOMParser().parseFromString(decodeURIComponent(options.slice('?fetchXml='.length)),'text/xml');
        const candidate=xml.querySelector('condition[attribute="accountid"] value')!.textContent!;
        return {entities:[{accountid:candidate}]};
    }} as unknown as ComponentFramework.WebApi;
    const candidates=Array.from({length:201},(_,n)=>id(n+1));
    assert.deepEqual(await resolveFilterIds(api,view,'account','accountid',negative,candidates,()=>undefined),[id(1),id(201)]);
    assert.equal(apiCalls,2);
    const incomplete={retrieveMultipleRecords:async()=>({entities:[],nextLink:'?more'})} as unknown as ComponentFramework.WebApi;
    await assert.rejects(resolveFilterIds(incomplete,view,'account','accountid',negative,[id(1)],()=>undefined),/incomplete/);
    const outside={retrieveMultipleRecords:async()=>({entities:[{accountid:id(99)}]})} as unknown as ComponentFramework.WebApi;
    await assert.rejects(resolveFilterIds(outside,view,'account','accountid',negative,[id(1)],()=>undefined),/outside/);
    assert.equal(matches({accountid:id(1)},filterForIds('accountid',[])),false);
    assert.equal(matches({accountid:id(201)},filterForIds('accountid',candidates)),true);
    await testControlFilterTransaction(negative);
    await testControlFilterTransaction(updateColumnFilter(undefined,'lookup','containsData')!);
    console.log('Server predicate compilation, full bound-dataset scope, paging, matching, limits and rollback tests passed.');
}

async function testControlFilterTransaction(negative:FilterExpression): Promise<void> {
    const root=new CMModelDrivenGrid();
    const control=root as unknown as {
        context:ComponentFramework.Context<never>;
        renderGrid:()=>void;
        loadDatasetPageAndWait:(change:()=>void)=>Promise<void>;
        getEntityMeta:()=>Promise<{primaryIdAttr:string}>;
        applyServerRootFilter:(filter:FilterExpression,sorting:[])=>Promise<void>;
        getRootFilter:()=>FilterExpression|undefined;
        rootFilterBusy:boolean;
        rootFilterError:string;
    };
    const rows=[{accountid:id(1),name:'alpha'},{accountid:id(2),name:'beta'},{accountid:id(3),name:'gamma'}];
    let activeFilter:FilterExpression|undefined, index=0, fail=false;
    const submitted:FilterExpression[]=[];
    const dataset={columns,sorting:[],records:{},sortedRecordIds:[] as string[],error:false,getTargetEntityType:()=> 'account',getViewId:()=>id(99),setSelectedRecordIds:()=>undefined,
        filtering:{getFilter:()=>activeFilter,setFilter:(filter:FilterExpression)=>{assert.equal(needsServerFilter(filter),false,'The native dataset must never receive unsupported NOT LIKE.');activeFilter=filter;submitted.push(filter);},clearFilter:()=>{activeFilter=undefined;}},
        paging:{hasNextPage:false,hasPreviousPage:false,loadNextPage:()=>{index++;load();},loadExactPage:(page:number)=>{index=page-1;load();}},
        refresh:()=>{index=0;load();},
    };
    function load() {
        const available=rows.filter(row=>matches(row,activeFilter));
        const current=available.slice(index,index+1);
        dataset.sortedRecordIds=current.map(row=>row.accountid);
        dataset.records=Object.fromEntries(current.map(row=>[row.accountid,{getRecordId:()=>row.accountid,getValue:(field:string)=>row[field as keyof typeof row]}]));
        dataset.paging.hasNextPage=index+1<available.length;dataset.paging.hasPreviousPage=index>0;
    }
    control.context={parameters:{records:dataset},webAPI:{retrieveRecord:async()=>({returnedtypecode:'account',fetchxml:view}),retrieveMultipleRecords:async(_entity:string,options:string)=>{
        if (fail) throw new Error('Server check denied');
        const xml=new DOMParser().parseFromString(decodeURIComponent(options.slice('?fetchXml='.length)),'text/xml');
        const candidates=Array.from(xml.querySelectorAll('condition[attribute="accountid"] value')).map(node=>node.textContent);
        assert.deepEqual(candidates,rows.map(row=>row.accountid),'Must check every bound page, not only the first row.');
        return {entities:rows.filter(row=>matches(row,negative)).map(row=>({accountid:row.accountid}))};
    }}} as unknown as ComponentFramework.Context<never>;
    control.renderGrid=()=>undefined;
    control.getEntityMeta=async()=>({primaryIdAttr:'accountid'});
    control.loadDatasetPageAndWait=async change=>{change();};
    await control.applyServerRootFilter(negative,[]);
    assert.equal(control.rootFilterBusy,false);assert.equal(control.rootFilterError,'');
    assert.deepEqual(rows.filter(row=>matches(row,activeFilter)).map(row=>row.accountid),rows.filter(row=>matches(row,negative)).map(row=>row.accountid));
    assert.deepEqual(control.getRootFilter(),negative,'The UI must retain the chosen text rule, not show record-ID conditions.');
    const successful=activeFilter;
    fail=true;
    await control.applyServerRootFilter(updateColumnFilter(undefined,'name','doesNotContain','beta')!,[]);
    assert.deepEqual(activeFilter,successful,'A failed server check must restore the exact prior dataset filter.');
    assert.match(control.rootFilterError,/Server check denied/);
    assert.equal(control.rootFilterBusy,false);
    assert.ok(submitted.length>=2);
    const nativeFailure=root as unknown as {
        nativeHierarchy:{refreshRoots:()=>Promise<void>};
        nativeSelectionError:string;
        refreshPhysicalFilter:(filter:FilterExpression,sorting:[])=>Promise<void>;
    };
    nativeFailure.nativeHierarchy={refreshRoots:async()=>{nativeFailure.nativeSelectionError='Native paging timed out';}};
    await assert.rejects(nativeFailure.refreshPhysicalFilter(negative,[]),/Native paging timed out/,'A native controller failure must abort resolution even if the dataset has not set its error flag.');
}
