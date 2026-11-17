import * as React from 'react';
import * as ReactDOM from 'react-dom';
import {initializeIcons} from '@fluentui/react/lib/Icons';
import {BlindTypeSetup} from '../blindtype-setup';
import {dataverseStore, SetupState, SetupStore} from '../blindtype-setup-store';
import {NativeSetupHost, SetupCommandState} from '../blindtype-native-commands';
import {NativeHierarchySelection} from '../../CMModelDrivenGrid/nativeHierarchy';
import {FilterExpression} from '../../CMModelDrivenGrid/filtering';
import '../../CMModelDrivenGrid/css/CMModelDrivenGrid.css';
import '../blindtype-setup.css';

interface Inputs {records: ComponentFramework.PropertyTypes.DataSet}
interface PageWaiter {baseline:string;resolve:()=>void;reject:(failure:Error)=>void;timer:ReturnType<typeof setTimeout>}
type NativeContext = ComponentFramework.Context<Inputs> & {events: {SetupCommandsChanged: (state: SetupCommandState) => void}};
const guid = (value: string): string => value.replace(/[{}]/g,'').toLowerCase();
const lookupId = (value: unknown): string | undefined => {
    if(typeof value==='string')return /^[0-9a-f-]{36}$/i.test(guid(value))?guid(value):undefined;
    const row = (Array.isArray(value) ? value[0] : value) as {id?: string | {guid?:string}} | undefined;
    const id = typeof row?.id === 'string' ? row.id : row?.id?.guid;
    return id && /^[0-9a-f-]{36}$/i.test(guid(id)) ? guid(id) : undefined;
};
export class CMBlindTypeSetupGrid implements ComponentFramework.StandardControl<Inputs, Record<string, never>> {
    private context!: NativeContext;
    private container!: HTMLDivElement;
    private native!: NativeHierarchySelection;
    private store?: SetupStore;
    private itemId?: string;
    private sourceIds = new Map<string,string>();
    private destroyed = false;
    private changed?: SetupCommandState;
    private binding?: string;
    private waiters = new Set<PageWaiter>();
    private host: NativeSetupHost = {
        publish: state => {this.changed = state; this.raise();},
        select: async (ids, children) => {
            const translate = (source: string): string => {const id=this.sourceIds.get(source);if(!id)throw new Error('This setup row is not available in the native subgrid. Refresh the subgrid.');return id;};
            await this.native.select(ids.map(translate),children.map(translate));
        },
        refreshed: () => {this.context.parameters.records.refresh();},
    };
    public init(context: ComponentFramework.Context<Inputs>, _notifyOutputChanged:()=>void, _state: ComponentFramework.Dictionary, container: HTMLDivElement): void {
        this.context = context as NativeContext; this.container = container;
        initializeIcons(undefined,{disableWarnings:true}); context.mode.trackContainerResize(true);
        this.native = new NativeHierarchySelection('cm_parentrow','cm_blindtypesetuprowid',{
            dataset:()=>this.context.parameters.records,basePageSize:()=>50,
            load:change=>this.loadPage(change),changed:()=>this.raise(),resetRows:()=>undefined,resetSelection:()=>undefined,
            error:failure=>{this.changed=undefined;this.renderError(failure);this.raise();},
        });
    }
    private raise = ():void => {
        const event = this.context.events?.SetupCommandsChanged;
        if (!event) return;
        const state = this.changed;
        const ready = !this.destroyed && !!state?.ready && !this.native?.busy && !this.context.mode.isControlDisabled && !this.context.parameters.records.error;
        event({...state,recordId:this.itemId,bindRecord:this.itemId?undefined:this.bindRecord,ready,selected:state?.selected??[],canNew:ready&&!!state?.canNew,canEdit:ready&&!!state?.canEdit,canDelete:ready&&!!state?.canDelete,canActivate:ready&&!!state?.canActivate,canDeactivate:ready&&!!state?.canDeactivate,canRefresh:ready&&!!state?.canRefresh,execute:command=>{if(ready)state?.execute(command);}});
    };
    private bindRecord = (value:string):void => {
        const itemId=guid(value);
        if(this.destroyed||this.itemId||this.binding||!/^[0-9a-f-]{36}$/.test(itemId))return;
        this.binding=itemId;
        this.itemId=itemId;
        const inner=dataverseStore(window.location.origin,itemId);
        this.store={kind:'dataverse',load:async()=>this.mapRows(await inner.load()),save:async(before,document)=>this.mapRows(await inner.save(before,document))};
        const scope:FilterExpression={filterOperator:0,conditions:[{attributeName:'cm_blindtype',conditionOperator:0,value:itemId}],filters:this.native.rootFilter?[this.native.rootFilter]:[]};
        this.updateView(this.context);
        void this.native.refreshRoots(scope);
    };
    private loadPage(change:()=>void):Promise<void> {
        return new Promise((resolve,reject)=>{
            const waiter:PageWaiter={baseline:this.signature(),resolve:()=>{clearTimeout(waiter.timer);this.waiters.delete(waiter);resolve();},reject:(failure:Error)=>{clearTimeout(waiter.timer);this.waiters.delete(waiter);reject(failure);},timer:setTimeout(()=>waiter.reject(new Error('Power Apps did not refresh the setup subgrid.')),30000)};
            this.waiters.add(waiter);try{change();}catch(failure){waiter.reject(failure as Error);}
        });
    }
    private signature():string {const data=this.context.parameters.records;return JSON.stringify([data.sortedRecordIds,Object.keys(data.records),data.paging.hasNextPage,data.paging.hasPreviousPage]);}
    private async mapRows(state: SetupState):Promise<SetupState> {
        const rows = await this.context.webAPI.retrieveMultipleRecords('cm_blindtypesetuprow',`?$select=cm_blindtypesetuprowid,cm_sourcekey&$filter=_cm_blindtype_value eq ${state.id}&$top=5000`);
        if(rows.nextLink)throw new Error('The native setup contains too many rows for this editor.');
        this.sourceIds = new Map(rows.entities.map(row=>[String(row.cm_sourcekey),guid(String(row.cm_blindtypesetuprowid))]));
        return state;
    }
    private renderError(failure:string):void {ReactDOM.render(React.createElement('div',{role:'alert'},failure),this.container);}
    public updateView(context: ComponentFramework.Context<Inputs>):void {
        if(this.destroyed)return;this.context=context as NativeContext;
        const dataset=context.parameters.records;
        if(!dataset.loading){for(const waiter of [...this.waiters]){if(dataset.error)waiter.reject(new Error('Power Apps could not load the selected setup rows.'));else if(context.updatedProperties.some(name=>name==='records'||name==='dataset')||waiter.baseline!==this.signature())waiter.resolve();}}
        if(dataset.getTargetEntityType()!=='cm_blindtypesetuprow'){this.renderError('Bind Blind Type Setup Grid only to Blind type setup rows.');return;}
        const ids = new Set(Object.values(dataset.records).map(row=>lookupId(row.getValue('cm_blindtype'))).filter((id):id is string=>!!id));
        if(ids.size>1){this.renderError('This subgrid must be related to one blind type.');return;}
        if(this.itemId&&ids.size&&![...ids].includes(this.itemId)){this.changed=undefined;this.renderError('The setup rows do not belong to this blind type.');this.raise();return;}
        if(!this.itemId){this.renderError('Loading blind-type setup…');this.raise();return;}
        const width=context.mode.allocatedWidth>0?context.mode.allocatedWidth:undefined;
        const height=context.mode.allocatedHeight>0?context.mode.allocatedHeight:undefined;
        this.container.style.height=height?`${height}px`:'100%';this.container.style.width=width?`${width}px`:'100%';
        ReactDOM.render(React.createElement(BlindTypeSetup,{store:this.store!,nativeHost:this.host}),this.container);
        this.raise();
    }
    public getOutputs():Record<string,never>{return {};}
    public destroy():void {this.destroyed=true;this.changed=undefined;this.raise();this.native.dispose();for(const waiter of [...this.waiters])waiter.reject(new Error('The setup control was closed.'));ReactDOM.unmountComponentAtNode(this.container);}
}
