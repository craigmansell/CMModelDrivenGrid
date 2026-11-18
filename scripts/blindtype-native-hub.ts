import type {SetupCommand, SetupCommandState} from './blindtype-native-commands';
import type {SetupCommandEnvelope} from './blindtype-native-channel';

interface NativeGrid {getName(): string; getEntityName?(): string; refreshRibbon?(): void}
interface NativeForm {data: {entity: {getEntityName(): string; getId(): string}}; getControl(name: string): NativeGrid | null; ui?: {refreshRibbon(all: boolean): void}}
interface NativeApi {
    onLoad(context: {getFormContext(): NativeForm}): void;
    can(form: NativeForm, grid: NativeGrid, command: string, table?: string): boolean;
    run(form: NativeForm, grid: NativeGrid, command: string, table?: string): void;
}
interface Entry {element: HTMLElement; envelope: SetupCommandEnvelope; signature: string; binding: boolean; primed: boolean; scopeId?: string}
export interface OwnedSetupHub {protocol: 1; revision: 2; api: NativeApi; publish(element: HTMLElement, envelope: SetupCommandEnvelope, refresh?: () => void): void}
interface OwnedNamespace {[name: string]: unknown; BlindTypeSetupOwnedHub?: OwnedSetupHub}
type OwnedWindow = Window & {CM?: OwnedNamespace};
const gridName = 'Subgrid_BlindTypeSetup';
const normalize = (value: string | undefined): string => String(value || '').replace(/[{}]/g, '').toLowerCase();
const validId = (id: string): boolean => /^[0-9a-f]{8}(-[0-9a-f]{4}){3}-[0-9a-f]{12}$/.test(id);
const flags: Record<SetupCommand, keyof SetupCommandState> = {new:'canNew',edit:'canEdit',delete:'canDelete',activate:'canActivate',deactivate:'canDeactivate',refresh:'canRefresh'};

// Ribbon functions receive the documented PrimaryControl/SelectedControl args.
// Component state stays in our own registry; no host DOM or private SDK access.
export function ownedSetupHub(view: Window): OwnedSetupHub {
    const host = view as OwnedWindow;
    const namespace = host.CM || (host.CM = {});
    if (namespace.BlindTypeSetupOwnedHub?.protocol === 1 && namespace.BlindTypeSetupOwnedHub.revision === 2) return namespace.BlindTypeSetupOwnedHub;
    const entries = new Map<string, Entry>();
    const forms = new Map<string, {form: NativeForm; grid: NativeGrid}>();
    const prune = (): void => {for (const [key, entry] of entries) if (!entry.element.isConnected) entries.delete(key);};
    const identity = (form: NativeForm, grid: NativeGrid, table?: string): string | undefined => {
        if (form.data.entity.getEntityName() !== 'cm_item' || grid.getName() !== gridName || normalize(table || grid.getEntityName?.()) !== 'cm_blindtypesetuprow') return;
        const current = form.getControl(gridName), id = normalize(form.data.entity.getId());
        if (current?.getName() === gridName && validId(id)) return id;
    };
    const notify = (id: string | undefined): void => {
        const context = id && forms.get(id);
        if (context) {try {if (context.grid.refreshRibbon) context.grid.refreshRibbon(); else context.form.ui?.refreshRibbon(true);} catch { /* A navigated form has no active command bar. */ }}
    };
    const resolve = (id: string, bind: boolean): Entry | undefined => {
        prune();
        let matches = [...entries.values()].filter(entry => normalize(entry.envelope.payload.recordId || entry.envelope.recordHint || entry.scopeId) === id);
        if (!matches.length && bind) matches = [...entries.values()].filter(entry => !entry.envelope.payload.recordId && !entry.envelope.recordHint && !entry.scopeId && typeof entry.envelope.payload.bindRecord==='function');
        if (matches.length !== 1) return;
        const entry = matches[0], payload = entry.envelope.payload;
        if (payload.recordId && normalize(payload.recordId) !== id) return;
        if (entry.envelope.recordHint && normalize(entry.envelope.recordHint) !== id) return;
        if (entry.scopeId && entry.scopeId !== id) return;
        entry.scopeId=id;
        if (!payload.recordId && bind && !entry.envelope.loading && payload.bindRecord && !entry.binding) {entry.binding = true;try {payload.bindRecord(id);} catch {entry.binding = false;}}
        return entry;
    };
    const pin = (): void => {namespace.BlindTypeSetup = api;for (const version of [4,5,6,7,8,9,10]) namespace[`BlindTypeSetupV${version}`] = api;};
    const api: NativeApi = {
        onLoad(context) {
            try {const form=context.getFormContext(),grid=form.getControl(gridName);if(!grid)return;const id=identity(form,grid,'cm_blindtypesetuprow');if(id){forms.set(id,{form,grid});resolve(id,true);}} catch { /* Commands stay disabled until a valid native context arrives. */ }
        },
        can(form,grid,command,table) {
            try {
                const id=identity(form,grid,table);if(!id||!Object.prototype.hasOwnProperty.call(flags,command))return false;
                forms.set(id,{form,grid});const entry=resolve(id,true),payload=entry?.envelope.payload;
                return !!(entry&&!entry.envelope.destroyed&&!entry.envelope.loading&&payload?.ready&&typeof payload.execute==='function'&&normalize(payload.recordId)===id&&payload[flags[command as SetupCommand]]===true);
            } catch {return false;}
        },
        run(form,grid,command,table) {if(api.can(form,grid,command,table)){const id=identity(form,grid,table)!;resolve(id,false)?.envelope.payload.execute(command as SetupCommand);}},
    };
    const hub: OwnedSetupHub = {
        protocol:1,revision:2,api,
        publish(element,envelope,refresh) {
            if(!element.isConnected||envelope.protocol!==1||typeof envelope.loading!=='boolean'||!envelope.payload||envelope.component!=='CMBlindTypeSetupGridV2'||envelope.controlName!==gridName)return;
            const previous=entries.get(envelope.instanceId);if(previous&&previous.element!==element)return;
            const id=normalize(envelope.payload.recordId||envelope.recordHint||previous?.scopeId);
            if(envelope.destroyed){entries.delete(envelope.instanceId);notify(id);return;}
            const signature=JSON.stringify([envelope.loading,envelope.payload]);
            const entry:Entry=previous||{element,envelope,signature:'',binding:false,primed:false};
            const changed=entry.signature!==signature;entry.envelope=envelope;entry.signature=signature;entry.binding=false;entries.set(envelope.instanceId,entry);pin();
            if(!envelope.loading&&envelope.recordHint&&!envelope.payload.recordId&&validId(id))resolve(id,true);
            if(changed)notify(id);
            if(!entry.primed&&!envelope.loading&&!envelope.payload.recordId&&!envelope.recordHint&&!forms.has(id)&&refresh){entry.primed=true;refresh();}
        },
    };
    namespace.BlindTypeSetupOwnedHub=hub;pin();return hub;
}
