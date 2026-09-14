import {strict as assert} from 'assert';
import {editableAttribute, readColumnEditability} from '../CMModelDrivenGrid/columnEditability';
import {CMModelDrivenGrid} from '../CMModelDrivenGrid';

async function run(): Promise<void> {
    assert.equal(editableAttribute({IsValidForUpdate:false,SourceType:0}),false,'Created On must never become editable merely because it is a date.');
    for (const source of [1,2,3]) assert.equal(editableAttribute({IsValidForUpdate:true,SourceType:source}),false,'Calculated, rollup and formula fields stay read only.');
    assert.equal(editableAttribute({IsValidForUpdate:true,AttributeOf:'money'}),false,'Derived base/companion attributes stay read only.');
    assert.equal(editableAttribute({}),false,'Unknown metadata fails closed.');
    assert.equal(editableAttribute({IsValidForUpdate:true,SourceType:0}),true);
    const urls:string[]=[];
    const request=(async(url:string)=>{
        urls.push(url);
        return {ok:true,json:async()=>({value:[
            {LogicalName:'createdon',IsValidForUpdate:false,SourceType:0},
            {LogicalName:'cm_cost',IsValidForUpdate:true,SourceType:0},
            {LogicalName:'not_requested',IsValidForUpdate:true},
        ]})};
    }) as unknown as typeof fetch;
    const allowed=await readColumnEditability('cm_itemtemplateoption',['createdon','CM_COST','joined.field','missing'], 'https://sandbox.example',request);
    assert.deepEqual(allowed,{createdon:false,cm_cost:true});
    assert(!urls[0].includes('joined.field'));
    await assert.rejects(readColumnEditability("invalid'table",[], 'https://sandbox.example',request));
    await assert.rejects(readColumnEditability('account',['name'],'https://sandbox.example',(async()=>({ok:false})) as unknown as typeof fetch));
    urls.length=0;
    await readColumnEditability('account',Array.from({length:85},(_,i)=>`field${i}`),'https://sandbox.example',request);
    assert.equal(urls.length,3,'Column metadata requests must be URL-bounded.');

    const grid=new CMModelDrivenGrid();let writes=0;
    grid.context={mode:{isControlDisabled:false},parameters:{EnableInlineEdit:{raw:true},records:{
        getTargetEntityType:()=> 'account',columns:[{name:'createdon',dataType:'DateAndTime.DateOnly'},{name:'cm_cost',dataType:'Money'}],refresh:()=>undefined,
    }},webAPI:{updateRecord:async()=>{writes++;}},navigation:{openAlertDialog:async()=>undefined}} as unknown as typeof grid.context;
    await assert.rejects(grid.onUpdateCell('row','cm_cost','5','Money'));
    assert.equal(writes,0,'A pending/failed metadata request must prevent saves.');
    Object.assign(grid,{columnEditabilityEntity:'account',columnEditability:allowed});
    await assert.rejects(grid.onUpdateCell('row','createdon','2026-09-01','DateAndTime.DateOnly'));
    assert.equal(writes,0,'The controller must reject system-managed columns, including child saves.');
    await grid.onUpdateCell('row','cm_cost','5','Money');
    assert.equal(writes,1,'Known editable fields remain usable.');
    grid.context.parameters.records.getTargetEntityType=()=> 'contact';
    await assert.rejects(grid.onUpdateCell('row','cm_cost','6','Money'));
    assert.equal(writes,1,'Metadata from the previous table must not authorize a save.');

    const originalFetch=globalThis.fetch;
    const originalWindow=globalThis.window;
    const pending: ((response: Response)=>void)[]=[];
    const access=grid as unknown as {columnEditabilityRevision:number;columnEditability:Record<string,boolean>;fetchColumnEditability:(entity:string,columns:string[],revision:number)=>Promise<void>;renderGrid:()=>void};
    try {
        Object.assign(globalThis,{window:{location:{origin:'https://sandbox.example'}}});
        globalThis.fetch=(() => new Promise<Response>(resolve=>pending.push(resolve))) as typeof fetch;
        access.renderGrid=()=>undefined;
        access.columnEditabilityRevision=1;
        const older=access.fetchColumnEditability('account',['cm_cost'],1);
        access.columnEditabilityRevision=2;
        const newer=access.fetchColumnEditability('account',['cm_cost'],2);
        pending[1]({ok:true,json:async()=>({value:[{LogicalName:'cm_cost',IsValidForUpdate:false}]})} as Response);
        await newer;
        pending[0]({ok:true,json:async()=>({value:[{LogicalName:'cm_cost',IsValidForUpdate:true}]})} as Response);
        await older;
        assert.equal(access.columnEditability.cm_cost,false,'A stale metadata response cannot reopen an editor after switching away and back.');
    } finally {globalThis.fetch=originalFetch;Object.assign(globalThis,{window:originalWindow});}
    console.log('Column editability metadata and guarded-save tests passed.');
}
void run().catch(error=>{console.error(error);process.exitCode=1;});
