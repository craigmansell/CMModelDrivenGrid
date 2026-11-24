// Opt-in bridge to the saved Dataverse Calculator/Condition/Result service.
// No catalogue IDs or business thresholds live in this reusable grid.
const parents: Record<string,string> = {cm_itemtemplateoption:'cm_itemtemplate',cm_quotationlineoption:'cm_quotationline',cm_saleorderlineoption:'cm_saleorderline'};
const guid = (value: string): string => {
    const id=value.replace(/[{}]/g,'').toLowerCase();
    if (!/^[0-9a-f]{8}(-[0-9a-f]{4}){3}-[0-9a-f]{12}$/.test(id)) throw new Error('Invalid component rule context.');
    return id;
};
export interface AvailabilityContext {width:number|null;drop:number|null;itemId:string|null;departmentId:string|null}
interface HostResponse {type:string;nonce?:string;ready?:boolean;error?:string;context?:AvailabilityContext;selection?:{id:string;name:string;entityType:string}[]}
export interface AvailabilityResult {entryTable:string;entryId:string;excludedComponentIds:string[];governedComponentIds:string[]}
interface WebApi {retrieveRecord(table:string,id:string,options?:string):Promise<ComponentFramework.WebApi.Entity>}
export function availabilityClause(ids: string[]): string {
    return ids.map(id=>`cm_componentid ne ${guid(id)}`).map(clause=>`(${clause})`).join(' and ');
}
export class ComponentAvailabilityClient {
    revision=0;
    private channels=new Map<string,BroadcastChannel>();
    private pending=new Map<string,{resolve:(value:HostResponse)=>void;reject:(error:Error)=>void;timer:number}>();
    private disposed=false;
    constructor(private webApi:()=>WebApi,private changed:()=>void,private request:typeof fetch=fetch) {}
    supports(table:string,column?:string,target?:string):boolean {return !!parents[table]&&column?.toLowerCase()==='cm_component'&&target==='cm_component';}
    private async owner(table:string,id:string):Promise<{table:string;id:string;key:string}> {
        const parent=parents[table];if(!parent)throw new Error('Unsupported component rule entry.');
        const row=await this.webApi().retrieveRecord(table,guid(id),`?$select=_${parent}_value`);
        const parentId=guid(String(row[`_${parent}_value`]??''));return {table:parent,id:parentId,key:parent+'-'+parentId};
    }
    private host(parent:{key:string},type:string,extra:Record<string,unknown>={}):Promise<HostResponse> {
        if(this.disposed)throw new Error('The component grid has closed.');
        const token=sessionStorage.getItem('cm-component-availability-host-'+parent.key);
        if(!token)throw new Error('Refresh this entry to load its component rules.');
        const key=parent.key+'-'+token;
        let channel=this.channels.get(key);
        if(!channel){
            channel=new BroadcastChannel('cm-component-availability-v1-'+key);
            channel.onmessage=(event:MessageEvent<HostResponse>)=>{
                const message=event.data;if(!message)return;
                if(message.type==='changed'){this.revision++;this.changed();return;}
                if(message.type!=='response'||!message.nonce)return;
                const waiting=this.pending.get(message.nonce);if(!waiting)return;
                this.pending.delete(message.nonce);window.clearTimeout(waiting.timer);
                if(message.error||!message.ready)waiting.reject(new Error(message.error??'Entry dimensions are updating. Search again.'));else waiting.resolve(message);
            };
            this.channels.set(key,channel);
        }
        const nonce=crypto.randomUUID();
        return new Promise((resolve,reject)=>{
            const timer=window.setTimeout(()=>{this.pending.delete(nonce);reject(new Error('Component rules did not respond. Refresh the entry.'));},type==='browse'?300000:15000);
            this.pending.set(nonce,{resolve,reject,timer});channel!.postMessage({type,nonce,...extra});
        });
    }
    private async evaluate(table:string,id:string,context:AvailabilityContext|null):Promise<AvailabilityResult> {
        const response=await this.request('/api/data/v9.2/cm_GetEntryComponentAvailability',{method:'POST',credentials:'same-origin',headers:{'Content-Type':'application/json','OData-Version':'4.0','OData-MaxVersion':'4.0'},body:JSON.stringify({OptionTable:table,OptionRowId:guid(id),ContextJson:JSON.stringify(context??{})})});
        const body=await response.json() as {error?:{message?:string};ResultJson?:string};
        if(!response.ok||typeof body.ResultJson!=='string')throw new Error(body.error?.message??'Component availability could not be checked.');
        const result=JSON.parse(body.ResultJson) as AvailabilityResult;
        if(!Array.isArray(result.excludedComponentIds)||!Array.isArray(result.governedComponentIds))throw new Error('Invalid component rule response.');
        result.excludedComponentIds.forEach(guid);return result;
    }
    async current(table:string,id:string):Promise<{result:AvailabilityResult;assertCurrent:()=>void}> {
        const parent=await this.owner(table,id),state=await this.host(parent,'state');
        const revision=this.revision,result=await this.evaluate(table,id,state.context!);
        const assertCurrent=():void=>{if(this.disposed||revision!==this.revision||result.entryTable!==parent.table||guid(result.entryId)!==parent.id)throw new Error('Entry dimensions changed. Search for the component again.');};
        assertCurrent();return {result,assertCurrent};
    }
    async assertAllowed(table:string,id:string,component:string):Promise<void> {
        const chosen=guid(component),current=await this.current(table,id);
        if(current.result.excludedComponentIds.includes(chosen))throw new Error('This component is not available for the current dimensions. Choose another component before saving.');
        current.assertCurrent();
        // Save the parent only when the chosen component needs its unsaved size.
        // A replacement valid at both sizes can be saved first, allowing an
        // existing incompatible choice to be replaced before widening the entry.
        const saved=await this.evaluate(table,id,null);current.assertCurrent();
        if(saved.excludedComponentIds.includes(chosen)){
            const parent=await this.owner(table,id),state=await this.host(parent,'state');
            await this.host(parent,'save',{context:state.context});
            const committed=await this.evaluate(table,id,null);
            if(committed.excludedComponentIds.includes(chosen))throw new Error('Save the entry dimensions before choosing this component.');
        }
    }
    async browse(table:string,id:string,searchTerm:string):Promise<{id:string;name:string}|undefined> {
        const parent=await this.owner(table,id),response=await this.host(parent,'browse',{optionTable:table,optionRowId:guid(id),searchTerm});
        const selected=response.selection;if(!selected?.length)return undefined;
        if(selected.length!==1||selected[0].entityType!=='cm_component')throw new Error('Choose one component.');
        const choice={id:guid(selected[0].id),name:selected[0].name};
        const current=await this.current(table,id);if(current.result.excludedComponentIds.includes(choice.id))throw new Error('The component is no longer available for these dimensions.');return choice;
    }
    dispose():void {
        this.disposed=true;for(const channel of this.channels.values())channel.close();this.channels.clear();
        for(const pending of this.pending.values()){window.clearTimeout(pending.timer);pending.reject(new Error('The component grid has closed.'));}this.pending.clear();
    }
}
