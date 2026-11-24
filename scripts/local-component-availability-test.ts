import assert from 'node:assert/strict';
import {ComponentAvailabilityClient,availabilityClause} from '../CMModelDrivenGrid/componentAvailability';

const parent='11111111-1111-4111-8111-111111111111',row='22222222-2222-4222-8222-222222222222',r8024='33333333-3333-4333-8333-333333333333',r8027='44444444-4444-4444-8444-444444444444';
const store=new Map<string,string>();let width=900,savedWidth=1000,changes=0,closed=0,delayed=false,release: (()=>void)|undefined;
const peers=new Set<Channel>();
class Channel {
    onmessage?: (event:MessageEvent)=>void;
    constructor(public name:string){peers.add(this);}
    postMessage(message:Record<string,unknown>):void {
        if(message.type==='state'||message.type==='save'||message.type==='browse')queueMicrotask(()=>{
            if(message.type==='save')savedWidth=width;
            this.onmessage?.({data:{type:'response',nonce:message.nonce,ready:true,context:{width,drop:3000,itemId:parent,departmentId:null},...(message.type==='browse'?{selection:[{id:r8027,name:'R8027',entityType:'cm_component'}]}:{})}} as MessageEvent);
        });
    }
    close():void{closed++;peers.delete(this);}
}
Object.assign(globalThis,{window:globalThis,sessionStorage:{getItem:(k:string)=>store.get(k)??null},BroadcastChannel:Channel});
store.set('cm-component-availability-host-cm_itemtemplate-'+parent,'tab-specific-token');
const calls:Record<string,unknown>[]=[];
const request=async (_url:unknown,options:RequestInit):Promise<Response>=>{
    const body=JSON.parse(String(options.body)) as {ContextJson:string};calls.push(body);
    const context=JSON.parse(body.ContextJson) as {width?:number};const current=context.width??savedWidth;
    if(delayed)await new Promise<void>(resolve=>{release=resolve;});
    return {ok:true,json:async()=>({ResultJson:JSON.stringify({entryTable:'cm_itemtemplate',entryId:parent,excludedComponentIds:current>0&&current<1000?[]:[r8024],governedComponentIds:[r8024]})})} as Response;
};
async function run():Promise<void>{
    const client=new ComponentAvailabilityClient(()=>({retrieveRecord:async()=>({_cm_itemtemplate_value:parent})}),()=>changes++,request as typeof fetch);
    assert.equal(client.supports('cm_itemtemplateoption','cm_component','cm_component'),true);assert.equal(client.supports('cm_itemoption','cm_component','cm_component'),false);
    assert.equal((await client.current('cm_itemtemplateoption',row)).result.excludedComponentIds.length,0);
    await client.assertAllowed('cm_itemtemplateoption',row,r8024);assert.equal(savedWidth,900,'save unsaved width before committing a component that needs it');
    width=1000;assert.deepEqual((await client.current('cm_itemtemplateoption',row)).result.excludedComponentIds,[r8024]);
    await assert.rejects(()=>client.assertAllowed('cm_itemtemplateoption',row,r8024));
    await client.assertAllowed('cm_itemtemplateoption',row,r8027);assert.equal(savedWidth,900,'a valid replacement can save before widening parent');
    assert.deepEqual(await client.browse('cm_itemtemplateoption',row,'R80'),{id:r8027,name:'R8027'});
    assert.equal(availabilityClause([r8024]),'(cm_componentid ne '+r8024+')');
    delayed=true;const pending=client.current('cm_itemtemplateoption',row);while(!release)await new Promise(resolve=>setTimeout(resolve,0));
    for(const peer of peers)peer.onmessage?.({data:{type:'changed'}} as MessageEvent);release();delayed=false;
    await assert.rejects(()=>pending,/dimensions changed/);assert.ok(changes>0);
    client.dispose();assert.ok(closed>0);
    const absent=new ComponentAvailabilityClient(()=>({retrieveRecord:async()=>({_cm_itemtemplate_value:row})}),()=>{},request as typeof fetch);
    await assert.rejects(()=>absent.current('cm_itemtemplateoption',row),/Refresh/);absent.dispose();
    console.log('Passed component lookup connection: unsaved width, strict boundary, replacement sequencing, filtered browse, stale response, per-tab bridge, disposal, opt-in scope.');
}
void run().catch(error=>{console.error(error);process.exitCode=1;});
