import assert from 'node:assert/strict';
import {ComponentAvailabilityClient,availabilityClause} from '../CMModelDrivenGrid/componentAvailability';

const parent='11111111-1111-4111-8111-111111111111',row='22222222-2222-4222-8222-222222222222';
const part='33333333-3333-4333-8333-333333333333',other='44444444-4444-4444-8444-444444444444',foreign='55555555-5555-4555-8555-555555555555';
let width:number|null=null,savedWidth:number|null=null,browseChoice=other,malformed=false,empty=false,saves=0,revalidations=0;
class Channel {
    onmessage?: (event:MessageEvent)=>void;
    postMessage(message:Record<string,unknown>):void {queueMicrotask(()=>{
        if(message.type==='save'){savedWidth=width;saves++;}
        if(message.type==='revalidate'){assert.equal(message.optionTable,'cm_quotationlineoption');revalidations++;}
        this.onmessage?.({data:{type:'response',nonce:message.nonce,ready:true,context:{width,drop:null,itemId:parent,departmentId:null},...(message.type==='browse'?{selection:[{id:browseChoice,name:'Selected component',entityType:'cm_component'}]}:{})}} as MessageEvent);
    });}
    close():void { /* no external resources */ }
}
Object.assign(globalThis,{window:globalThis,sessionStorage:{getItem:()=> 'quotation-tab-token'},BroadcastChannel:Channel});
const request=async (_url:unknown,options:RequestInit):Promise<Response>=>{
    const body=JSON.parse(String(options.body)) as {ContextJson:string};
    const context=JSON.parse(body.ContextJson) as {width?:number|null};
    const measure=Object.prototype.hasOwnProperty.call(context,'width')?context.width:savedWidth;
    const excluded=measure!==null&&measure!==undefined&&(measure<=0||measure>=1000)?[part]:[];
    const result={entryTable:'cm_quotationline',entryId:parent,excludedComponentIds:excluded,governedComponentIds:[part],itemOptionScopeApplied:true,allowedComponentIds:empty?[]:[part,other].filter(id=>!excluded.includes(id))};
    if(malformed)delete (result as Partial<typeof result>).allowedComponentIds;
    return {ok:true,json:async()=>({ResultJson:JSON.stringify(result)})} as Response;
};
async function run():Promise<void>{
    const client=new ComponentAvailabilityClient(()=>({retrieveRecord:async()=>({_cm_quotationline_value:parent})}),()=>{},request as typeof fetch);
    const blank=await client.current('cm_quotationlineoption',row);
    assert.deepEqual(blank.result.allowedComponentIds,[part,other]);
    const clause=availabilityClause(blank.result.excludedComponentIds,blank.result.allowedComponentIds);
    assert.ok(clause.includes('cm_componentid eq '+part)&&clause.includes('cm_componentid eq '+other));
    assert.equal(clause.includes(foreign),false);
    assert.match(availabilityClause([],[]),/00000000-0000-0000-0000-000000000000/,'an empty assignment cannot widen to the catalogue');
    await client.assertAllowed('cm_quotationlineoption',row,part);assert.equal(saves,0,'blank measurements permit the configured provisional choice without an automatic save');
    width=900;await client.assertAllowed('cm_quotationlineoption',row,part);assert.equal(saves,0);
    width=1000;assert.deepEqual((await client.current('cm_quotationlineoption',row)).result.allowedComponentIds,[other]);
    await assert.rejects(()=>client.assertAllowed('cm_quotationlineoption',row,part),/not available/);
    await assert.rejects(()=>client.assertAllowed('cm_quotationlineoption',row,foreign),/not available/);
    await client.assertAllowed('cm_quotationlineoption',row,other);assert.equal(saves,0,'replacement may save before an incompatible parent width');
    assert.equal((await client.browse('cm_quotationlineoption',row,'Other'))?.id,other);
    browseChoice=foreign;await assert.rejects(()=>client.browse('cm_quotationlineoption',row,'Foreign'),/no longer available/);
    empty=true;assert.deepEqual((await client.current('cm_quotationlineoption',row)).result.allowedComponentIds,[]);
    await assert.rejects(()=>client.assertAllowed('cm_quotationlineoption',row,other),/not available/);
    malformed=true;await assert.rejects(()=>client.current('cm_quotationlineoption',row),/Invalid component rule response/);
    await client.selectionSaved('cm_quotationlineoption',row);assert.equal(revalidations,1,'a committed component edit or clear revalidates the dirty parent');
    await client.selectionSaved('cm_saleorderlineoption',row);assert.equal(revalidations,1,'order/template behavior is unchanged');
    const detached=new ComponentAvailabilityClient(()=>({retrieveRecord:async()=>{throw new Error('Closed host')}}),()=>{},request as typeof fetch);
    await detached.selectionSaved('cm_quotationlineoption',row);detached.dispose();
    client.dispose();await client.selectionSaved('cm_quotationlineoption',row);assert.equal(revalidations,1,'closed hosts cannot turn an already committed edit into a failed save');
    console.log('Quotation component scope passed: assigned non-default choices, blank/900/1000, no foreign catalogue choices, empty/malformed fail closed, full browse and no reseeding/save side effects.');
}
void run().catch(error=>{console.error(error);process.exitCode=1;});
