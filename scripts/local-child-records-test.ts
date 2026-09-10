import assert from 'node:assert/strict';
import { JSDOM } from 'jsdom';
import { buildChildFetch, buildChildAvailabilityFetch, childRecord, createChildLoader, createChildAvailabilityLoader } from '../CMModelDrivenGrid/childRecords';
import { testDatasetFiltering } from './local-dataset-filter-test';
import { updateColumnFilter } from '../CMModelDrivenGrid/filtering';
import { CMModelDrivenGrid } from '../CMModelDrivenGrid';
import { createChildActions } from '../CMModelDrivenGrid/childActions';
import { createDataverseChildPaging } from '../CMModelDrivenGrid/dataverseChildPaging';
const dom = new JSDOM('');
globalThis.DOMParser = dom.window.DOMParser;
globalThis.XMLSerializer = dom.window.XMLSerializer;
const parent = '00000000-0000-4000-8000-000000000001';
const child = '00000000-0000-4000-8000-000000000002';
const view = '<fetch distinct="true"><entity name="cm_itemtemplateoption"><attribute name="cm_name"/><filter type="or"><condition attribute="statecode" operator="eq" value="0"/><condition attribute="cm_name" operator="eq" value="Special"/></filter><order attribute="cm_order"/></entity></fetch>';
const filtering = updateColumnFilter(undefined,'cm_name','contains',`A&B <test> "quoted"`);
const result = buildChildFetch(view,'cm_itemtemplateoption','cm_itemtemplateoptionid','cm_parenttemplateoption',{parentId:parent,page:2,sorting:[{name:'cm_name',sortDirection:1}],filtering});
const xml = new DOMParser().parseFromString(result,'text/xml');
assert.equal(xml.querySelector('parsererror'),null);
assert.equal(xml.documentElement.getAttribute('page'),'2');
assert.equal(xml.querySelector('filter[type="or"]')?.children.length,2);
assert.equal(xml.querySelector('condition[attribute="cm_parenttemplateoption"]')?.getAttribute('value'),parent);
assert.ok(xml.querySelector('condition[operator="like"]')?.getAttribute('value')?.includes('A&B <test> "quoted"'));
assert.deepEqual(Array.from(xml.querySelectorAll('order')).map(node=>node.getAttribute('attribute')),['cm_name','cm_itemtemplateoptionid']);
assert.throws(()=>buildChildFetch(view,'account','accountid','parentaccountid',{parentId:parent,page:1,sorting:[]}));
assert.throws(()=>buildChildFetch(view,'cm_itemtemplateoption','cm_itemtemplateoptionid','bad"><filter',{parentId:parent,page:1,sorting:[]}));
assert.throws(()=>buildChildFetch(view,'cm_itemtemplateoption','cm_itemtemplateoptionid','cm_parenttemplateoption',{parentId:'invalid',page:1,sorting:[]}));
const row = childRecord({cm_itemtemplateoptionid:child,cm_name:'Child',_cm_component_value:parent,'_cm_component_value@Microsoft.Dynamics.CRM.lookuplogicalname':'cm_component','_cm_component_value@OData.Community.Display.V1.FormattedValue':'Fabric',cm_totalcost:12,'cm_totalcost@OData.Community.Display.V1.FormattedValue':'£12.00'},'cm_itemtemplateoption','cm_itemtemplateoptionid','cm_name');
assert.equal(row.getRecordId(),child); assert.equal(row.getFormattedValue('cm_component'),'Fabric'); assert.equal(row.getFormattedValue('cm_totalcost'),'£12.00'); assert.equal(row.getValue('missing'),null);
assert.equal((row.getValue('cm_component') as unknown as {etn:string}).etn,'cm_component');
async function run() {
	await testClientPagingContract();
	await testHttpPagingCookie();
	await testAvailability();
	await testDatasetFiltering();
    await testActions();
    let reads=0;
    const api = {retrieveRecord:async()=>{reads++;return {fetchxml:view,returnedtypecode:'cm_itemtemplateoption'};},retrieveMultipleRecords:async(_entity:string,options:string)=>{assert.ok(decodeURIComponent(options).includes(parent));return {entities:[{cm_itemtemplateoptionid:child,cm_name:'Child'}]};}} as unknown as ComponentFramework.WebApi;
    const load = createChildLoader(api,'cm_itemtemplateoption','cm_itemtemplateoptionid','cm_name','cm_parenttemplateoption',parent);
    const first = await load({parentId:parent,page:1,sorting:[]}); await load({parentId:parent,page:2,sorting:[]});
    assert.equal(reads,1);assert.equal(first.hasNextPage,false);assert.equal(first.records.length,1);
    const control = new CMModelDrivenGrid() as unknown as {
        childRecords: Record<string, ComponentFramework.PropertyHelper.DataSetApi.EntityRecord>;
        resolveLookupFilterTokens: (clause:string, recordId:string)=>Promise<string>;
    };
    control.childRecords[child] = row;
    assert.equal(await control.resolveLookupFilterTokens('_cm_component_value eq {{cm_component.id}}',child), `_cm_component_value eq ${parent}`);
    console.log('Child query, escaping, scope, paging and record-adapter tests passed.');
}
async function testClientPagingContract() {
    const rows=Array.from({length:64},(_,index)=>({cm_itemtemplateoptionid:`00000000-0000-4000-8000-${String(index+100).padStart(12,'0')}`,cm_name:`Child ${index+1}`}));
    const nextLink='https://example.invalid/api/data/v9.2/cm_itemtemplateoptions?$skiptoken=%253Ccookie%2520page%253D%25222%2522%253E';
    const requests:string[]=[];
    const api={retrieveRecord:async()=>({fetchxml:view,returnedtypecode:'cm_itemtemplateoption'}),retrieveMultipleRecords:async(_entity:string,options:string,size:number)=>{
        requests.push(options);assert.equal(size,50);
        if(options.startsWith('?$skiptoken=')){assert.equal(options,nextLink.slice(nextLink.indexOf('?')));return {entities:rows.slice(50)};}
        assert.ok(options.startsWith('?fetchXml=<fetch'),'Client API FetchXML must be unencoded so the host can preserve its page number.');
        const doc=new DOMParser().parseFromString(options.slice('?fetchXml='.length),'text/xml');
        assert.equal(doc.querySelector('parsererror'),null);
        return doc.documentElement.getAttribute('page') === '2' ? {entities:rows.slice(50)} : {entities:rows.slice(0,50),nextLink};
    }} as unknown as ComponentFramework.WebApi;
    const load=createChildLoader(api,'cm_itemtemplateoption','cm_itemtemplateoptionid','cm_name','cm_parenttemplateoption',parent);
    const first=await load({parentId:parent,page:1,sorting:[],filtering:updateColumnFilter(undefined,'cm_name','contains',"100% & O'Brien")});
    const second=await load({parentId:parent,page:2,sorting:[],filtering:updateColumnFilter(undefined,'cm_name','contains',"100% & O'Brien"),continuation:first.continuation});
    assert.equal(first.records.length+second.records.length,64);assert.equal(second.hasNextPage,false);
    await assert.rejects(load({parentId:child,page:2,sorting:[],continuation:first.continuation}),/different filter, parent or page/);
    const simple=await load({parentId:parent,page:2,sorting:[]});assert.equal(simple.records.length,14,'Simple paging must retain an explicit page number if no continuation is supplied.');
    assert.equal(requests.length,3,'A mismatched continuation must not issue a query.');
    console.log('Raw client FetchXML and query-scoped continuation paging passed (50 + 14).');
}
async function testHttpPagingCookie() {
    const rows=Array.from({length:64},(_,index)=>({cm_itemtemplateoptionid:`00000000-0000-4000-8000-${String(index+200).padStart(12,'0')}`,cm_name:`Child ${index+1}`}));
    const inner='<cookie page="1"><cm_name last="O&apos;Brien &amp; Sons" first="Alpha"/></cookie>';
    const cookie=`<cookie pagenumber="2" pagingcookie="${encodeURIComponent(encodeURIComponent(inner))}" istracking="False"/>`;
    const requests:string[]=[];
    const fetcher=(async (url:RequestInfo | URL,init?:RequestInit) => {
        assert.equal(init?.credentials,'same-origin');assert.equal(init?.redirect,'error');
        assert.ok(String(url).startsWith('/api/data/v9.2/cm_itemtemplateoptions?fetchXml='));
        assert.equal((init?.headers as Record<string,string>).Prefer,'odata.include-annotations="*"');
        const xml=new DOMParser().parseFromString(decodeURIComponent(String(url).split('?fetchXml=')[1]),'text/xml');
        requests.push(xml.documentElement.getAttribute('page')!);
        const second=xml.documentElement.getAttribute('page') === '2';
        if(second) assert.equal(xml.documentElement.getAttribute('paging-cookie'),inner,'Use the server cookie unchanged after decoding and XML attribute escaping.');
        return new Response(JSON.stringify(second ? {value:rows.slice(50)} : {value:rows.slice(0,50),'@Microsoft.Dynamics.CRM.morerecords':true,'@Microsoft.Dynamics.CRM.fetchxmlpagingcookie':cookie}),{status:200});
    }) as typeof fetch;
    const api={retrieveRecord:async()=>({fetchxml:view,returnedtypecode:'cm_itemtemplateoption'}),retrieveMultipleRecords:async()=>{throw new Error('Do not use the repeating host wrapper for child pages.');}} as unknown as ComponentFramework.WebApi;
    const load=createChildLoader(api,'cm_itemtemplateoption','cm_itemtemplateoptionid','cm_name','cm_parenttemplateoption',parent,createDataverseChildPaging('cm_itemtemplateoptions',fetcher));
    const first=await load({parentId:parent,page:1,sorting:[]});
    await assert.rejects(load({parentId:parent,page:3,sorting:[],continuation:first.continuation}),/different filter, parent or page/);
    const second=await load({parentId:parent,page:2,sorting:[],continuation:first.continuation});
    assert.deepEqual(requests,['1','2']);assert.equal(first.records.length+second.records.length,64);assert.equal(second.hasNextPage,false);
    await assert.rejects(createDataverseChildPaging('cm_itemtemplateoptions',(async()=>new Response('',{status:403})) as typeof fetch)('<fetch/>'),/HTTP 403/);
    assert.throws(()=>createDataverseChildPaging('../accounts',fetcher),/Invalid/);
    console.log('Same-origin Dataverse paging cookies, 50+14 rows, end-of-data, scope and HTTP failure tests passed.');
}
async function testAvailability() {
    const richView=view.replace('<order attribute="cm_order"/>','<link-entity name="cm_component" from="cm_componentid" to="cm_component" alias="c"><attribute name="cm_name"/><filter><condition attribute="statecode" operator="eq" value="0"/></filter></link-entity><order attribute="cm_order"/>');
    const parsed=new DOMParser().parseFromString(buildChildAvailabilityFetch(richView,'cm_itemtemplateoption','cm_parenttemplateoption',[parent,child]),'text/xml');
    assert.equal(parsed.querySelectorAll('attribute').length,1);
    assert.equal(parsed.querySelector('attribute')?.getAttribute('alias'),'drivengrid_parent');
    assert.equal(parsed.documentElement.getAttribute('distinct'),'true');
    assert.ok(parsed.querySelector('link-entity[alias="c"] condition[attribute="statecode"]'));
    assert.ok(parsed.querySelector('filter[type="or"]'));
    assert.equal(parsed.querySelectorAll('order').length,0);
    assert.throws(()=>buildChildAvailabilityFetch(view,'cm_itemtemplateoption','cm_parenttemplateoption',['bad']));
    let viewReads=0, calls=0, fail=false, incomplete=false, invalid=false;
    const api={retrieveRecord:async()=>{viewReads++;return {fetchxml:richView,returnedtypecode:'cm_itemtemplateoption'};},retrieveMultipleRecords:async(_entity:string,query:string)=>{
        calls++;
        if (fail) throw new Error('Access denied');
        const xml=new DOMParser().parseFromString(decodeURIComponent(query.slice('?fetchXml='.length)),'text/xml');
        const values=Array.from(xml.querySelectorAll('condition[attribute="cm_parenttemplateoption"] value')).map(value=>value.textContent!);
        assert.ok(values.length <= 50);
        return {entities:[{drivengrid_parent:invalid ? 'ffffffff-ffff-4fff-8fff-ffffffffffff' : values[0]}],nextLink:incomplete ? '?next' : undefined};
    }} as unknown as ComponentFramework.WebApi;
    const check=createChildAvailabilityLoader(api,'cm_itemtemplateoption','cm_parenttemplateoption',parent);
    assert.deepEqual(await check([parent,child]),{[parent]:true,[child]:false});
    const ids=Array.from({length:51},(_,index)=>`00000000-0000-4000-8000-${String(index+1).padStart(12,'0')}`);
    await check(ids); assert.equal(viewReads,1);assert.equal(calls,3,'51 parents should use two batched checks, not 51 child queries.');
    incomplete=true;await assert.rejects(check([parent,child]),/incomplete/);incomplete=false;
    invalid=true;await assert.rejects(check([parent]),/outside/);invalid=false;
    fail=true;await assert.rejects(check([parent]),/Access denied/);fail=false;
    assert.equal((await check([parent]))[parent],true,'A failed check must be retryable.');
    console.log('Batched leaf-arrow availability, view scope, security failures and retry checks passed.');
}
async function testActions() {
    const template = '00000000-0000-4000-8000-000000000003';
    let confirmations=0, deleted=0, descendant=false, confirmed=true, reparent=false;
    let actualParent=parent;
    let opened: Record<string,string> | undefined;
    const actions = createChildActions({entity:'cm_itemtemplateoption',idField:'cm_itemtemplateoptionid',nameField:'cm_name',parentField:'cm_parenttemplateoption',contextField:'cm_itemtemplate',
        webAPI:{
            retrieveRecord:async(_entity:string,id:string)=>id===parent ? {cm_name:'Parent',_cm_itemtemplate_value:template} : {cm_name:'Child',_cm_parenttemplateoption_value:actualParent},
            retrieveMultipleRecords:async()=>({entities:descendant ? [{}] : []}),
            deleteRecord:async(_entity:string,id:string)=>{assert.equal(id,child);deleted++;},
        } as unknown as ComponentFramework.WebApi,
        navigation:{
            openForm:async(_options:unknown,parameters:Record<string,string>)=>{opened=parameters;return {savedEntityReference:[{id:child}]};},
            openConfirmDialog:async()=>{confirmations++;if(reparent) actualParent=template;return {confirmed};},
        } as unknown as ComponentFramework.Navigation,
    });
    assert.equal(await actions.add(parent),true);
    assert.equal(opened?.cm_parenttemplateoption,parent);
    assert.equal(opened?.cm_itemtemplate,template);
    descendant=true;
    await assert.rejects(actions.delete(parent,child),/has children/);
    assert.equal(confirmations,0);assert.equal(deleted,0);
    descendant=false;confirmed=false;
    assert.equal(await actions.delete(parent,child),false);assert.equal(deleted,0);
    confirmed=true;reparent=true;
    await assert.rejects(actions.delete(parent,child),/no longer belongs/);assert.equal(deleted,0);
    actualParent=parent;reparent=false;
    assert.equal(await actions.delete(parent,child),true);assert.equal(deleted,1);
    await assert.rejects(actions.delete(parent,parent),/parent cannot/);
    console.log('Child command scope, inherited template, cancellation, branch and reparent guards passed.');
}
void run().catch(error=>{console.error(error);process.exitCode=1;});
