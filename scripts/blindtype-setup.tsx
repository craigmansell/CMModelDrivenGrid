import * as React from 'react';
import {ResponsiveGrid} from '../CMModelDrivenGrid/ResponsiveGrid';
import {ChildRequest} from '../CMModelDrivenGrid/childRecords';
import {FilterExpression, updateColumnFilter} from '../CMModelDrivenGrid/filtering';
import {selectionResult, checkSelectionCancelled} from '../CMModelDrivenGrid/recordSelection';
import {evaluateDemoFilter, DemoFilterExpression, sortDemoRows} from './demo-filter-evaluator';
import {pageDemoRows} from './demo-data';
import {availableComponents, Catalogue, clone, componentChoices, isCustomOption, optionGroup, LinkedSetup, newCondition, OptionSetup, SetupComponent, SetupCondition, setupColumns, SetupDocument, setupRecord, SetupResult, SetupRow, setupRows, validateSetup, withComponent, withOption, withoutSetupRow} from './blindtype-setup-model';
import {SetupState, SetupStore} from './blindtype-setup-store';
import {ComponentEntryRow, componentLabel, resolveComponent} from './blindtype-component-row';
import {NativeSetupHost, SetupCommand} from './blindtype-native-commands';
import '../CMModelDrivenGrid/css/CMModelDrivenGrid.css';
import './blindtype-setup.css';

type Editor = {kind: 'option'; value: OptionSetup} | {kind: 'link'; value: LinkedSetup} | {kind: 'condition'; linkId: string; value: SetupCondition};
interface ComponentDraft {kind: 'component' | 'result'; rowId: string; parentId: string; groupId: string; replacingId?: string; linkId?: string; conditionId?: string; resultId?: string; order?: number}
const kinds = {name: 'text', kind: 'text', code: 'text', order: 'number', details: 'text', count: 'number'} as const;
const resources = {getString: (key: string) => ({Label_Grid_Footer_RecordCount: '{0} records ({1} Selected)', Label_Grid_Footer: 'Page {0}', Label_NoRecords: 'No matching records', Label_SortAZ: 'A to Z', Label_SortZA: 'Z to A', Label_DoesNotContainData: 'Does not contain data'}[key] ?? key), getResource: () => undefined} as unknown as ComponentFramework.Resources;
const id = (): string => `draft-${crypto.randomUUID()}`;
const query = <T extends object,>(rows: T[], filtering: FilterExpression | undefined, sorting: ChildRequest['sorting']): T[] => {
    const getValue = (row: T, column: string): unknown => row[column as keyof T];
    return sortDemoRows(rows.filter(row => evaluateDemoFilter(row, filtering as unknown as DemoFilterExpression, {getValue, columnKinds: kinds})),
        sorting.map(sort => ({columnName: sort.name, direction: sort.sortDirection === 1 ? 'desc' as const : 'asc' as const})), {getValue, getDisplayValue: (row, column) => String(getValue(row, column) ?? ''), columnKinds: kinds});
};
const message = (error: unknown): string => error instanceof Error ? error.message : 'The setup could not be changed.';

export function BlindTypeSetup({store, nativeHost}: {store: SetupStore; nativeHost?: NativeSetupHost}): React.ReactElement {
    const [saved, setSaved] = React.useState<SetupState>();
    const document = saved?.document;
    const [mode, setMode] = React.useState<'options' | 'links'>('options');
    const [revision, setRevision] = React.useState(0);
    const [selected, setSelected] = React.useState<string[]>([]);
    const [filtering, setFiltering] = React.useState<FilterExpression>();
    const [sorting, setSorting] = React.useState<ChildRequest['sorting']>([]);
    const [editor, setEditor] = React.useState<Editor>();
    const [componentDraft, setComponentDraft] = React.useState<ComponentDraft>();
    const [componentDirty, setComponentDirty] = React.useState(false);
    const componentInput = React.useRef({text: '', makeDefault: false});
    const componentInitial = React.useRef(componentInput.current);
    const editorInitial = React.useRef('');
    const [expandRequest, setExpandRequest] = React.useState<{id: string; sequence: number}>();
    const expandSequence = React.useRef(0);
    const [confirm, setConfirm] = React.useState<SetupRow>();
    const [busy, setBusy] = React.useState(false);
    const saveLock = React.useRef(false);
    const [error, setError] = React.useState('');
    const [status, setStatus] = React.useState('');
    const load = React.useCallback(async () => {setBusy(true); setError(''); try {setSaved(await store.load()); setSelected([]); setRevision(value => value + 1); setStatus('');} catch (failure) {setError(message(failure));} finally {setBusy(false);}}, [store]);
    React.useEffect(() => {void load();}, [load]);
    const pending = componentDirty || (!!editor && JSON.stringify(editor) !== editorInitial.current) || busy && !!saved;
    React.useEffect(() => {const prevent = (event: BeforeUnloadEvent): void => {if (pending) {event.preventDefault(); event.returnValue = '';}}; window.addEventListener('beforeunload', prevent); return () => window.removeEventListener('beforeunload', prevent);}, [pending]);
    const rows = React.useMemo(() => {
        if (!saved || !document) return [];
        const values = setupRows(saved.catalogue, document, mode);
        if (!componentDraft) return values;
        const draft: SetupRow = {id: componentDraft.rowId, parentId: componentDraft.parentId, groupId: componentDraft.groupId, kind: componentDraft.kind === 'result' ? 'Result' : 'Component', name: '', code: '', order: componentDraft.order ?? 0, details: 'Not saved', count: null};
        return [...values.filter(row => row.id !== draft.id), draft];
    }, [saved, document, mode, componentDraft]);
    const byId = React.useMemo(() => new Map(rows.map(row => [row.id, row])), [rows]);
    const branches = React.useMemo(() => {const index = new Map<string, typeof rows>(); for (const row of rows) if (row.parentId) index.set(row.parentId, [...index.get(row.parentId) ?? [], row]); return index;}, [rows]);
    const visible = React.useMemo(() => query(rows.filter(row => !row.parentId), filtering, sorting), [rows, filtering, sorting]);
    const records = React.useMemo(() => Object.fromEntries(visible.map(row => [row.id, setupRecord(row)])), [visible]);
    const ids = React.useMemo(() => visible.map(row => row.id), [visible]);
    const loadChildren = React.useCallback(async (request: ChildRequest) => {
        const children = branches.get(request.parentId) ?? [];
        const entry = children.filter(row => row.id === componentDraft?.rowId);
        const page = pageDemoRows([...entry, ...query(children.filter(row => row.id !== componentDraft?.rowId), request.filtering, request.sorting)], request.page);
        return {records: page.rows.map(setupRecord), hasNextPage: page.hasNextPage};
    }, [branches, componentDraft]);
    const hasChildren = React.useCallback(async (parents: string[]) => Object.fromEntries(parents.map(parent => [parent, !!branches.get(parent)?.length])), [branches]);
    const row = selected.length === 1 ? byId.get(selected[0]) : undefined;
    const link = document?.links.find(item => item.id === row?.linkId);
    const condition = link?.conditions.find(item => item.id === row?.conditionId);
    const locked = busy || !!editor || !!componentDraft || !!confirm;
    const commit = React.useCallback(async (update: (draft: SetupDocument) => void | SetupDocument): Promise<boolean> => {
        if (!saved || saveLock.current) return false;
        saveLock.current = true; setBusy(true); setError(''); setStatus('Saving…');
        try {
            const draft = clone(saved.document); const next = update(draft) || draft;
            validateSetup(saved.catalogue, next);
            const state = await store.save(saved, next);
            setSaved(state); setSelected([]); setRevision(value => value + 1); setStatus(store.kind === 'dataverse' ? 'Saved in Dynamics' : 'Demo saved');
            nativeHost?.refreshed();
            return true;
        } catch (failure) {setError(message(failure)); setStatus('Not saved'); return false;}
        finally {saveLock.current = false; setBusy(false);}
    }, [saved, store, nativeHost]);
    const openEditor = (value: Editor): void => {setError(''); editorInitial.current = JSON.stringify(value); setEditor(value);};
    const beginComponent = (draft: ComponentDraft, text = '', makeDefault = false): void => {
        componentInput.current = {text, makeDefault}; componentInitial.current = componentInput.current; setComponentDirty(false); setError(''); setComponentDraft(draft);
        setSelected([]); setRevision(value => value + 1); setExpandRequest({id: draft.parentId, sequence: ++expandSequence.current});
    };
    const closeComponent = React.useCallback(() => {setComponentDraft(undefined); setComponentDirty(false); setError(''); setRevision(value => value + 1);}, []);
    const edit = (): void => {
        if (!row || !document || !saved) return;
        if (row.kind === 'Option') openEditor({kind: 'option', value: clone(document.options.find(item => item.groupId === row.groupId)!)});
        else if (row.kind === 'Component') {
            const option = document.options.find(item => item.groupId === row.groupId)!;
            const component = optionGroup(saved.catalogue, document, row.groupId).components.find(item => item.id === row.componentId)!;
            beginComponent({kind: 'component', rowId: row.id, parentId: row.parentId!, groupId: row.groupId, replacingId: row.componentId}, componentLabel(component), option.defaultId === row.componentId);
        } else if (row.kind === 'Linked option' && link) openEditor({kind: 'link', value: clone(link)});
        else if (row.kind === 'Condition' && link && condition) openEditor({kind: 'condition', linkId: link.id, value: clone(condition)});
        else if (row.kind === 'Result' && link && condition) beginComponent({kind: 'result', rowId: row.id, parentId: row.parentId!, groupId: row.groupId, linkId: link.id, conditionId: condition.id, resultId: row.resultId, order: row.order}, row.code);
    };
    const saveEditor = async (): Promise<void> => {
        if (!editor || !saved) return;
        if (await commit(draft => {
            if (editor.kind === 'option') return withOption(saved.catalogue, draft, editor.value);
            if (editor.kind === 'link') {const index = draft.links.findIndex(item => item.id === editor.value.id); if (index < 0) draft.links.push(clone(editor.value)); else draft.links[index] = clone(editor.value);}
            else {const target = draft.links.find(item => item.id === editor.linkId); if (!target) throw new Error('This linked option is no longer available.'); const index = target.conditions.findIndex(item => item.id === editor.value.id); if (index < 0) target.conditions.push(clone(editor.value)); else target.conditions[index] = clone(editor.value);}
        })) setEditor(undefined);
    };
    const remove = async (): Promise<void> => {
        if (!confirm || !saved) return;
        if (await commit(draft => withoutSetupRow(saved.catalogue, draft, confirm))) setConfirm(undefined);
    };
    const choices = React.useMemo(() => {
        if (!saved || !document || !componentDraft) return [];
        if (componentDraft.kind === 'component') return componentChoices(saved.catalogue, document, componentDraft.groupId, componentDraft.replacingId);
        const target = document.links.find(item => item.id === componentDraft.linkId)?.conditions.find(item => item.id === componentDraft.conditionId);
        return availableComponents(saved.catalogue, document, componentDraft.groupId).filter(component => !target?.results.some(result => result.id !== componentDraft.resultId && result.componentId === component.id));
    }, [saved, document, componentDraft]);
    const saveComponent = React.useCallback(async (text: string, makeDefault: boolean): Promise<void> => {
        if (!componentDraft || !saved) return;
        let selectedComponent: SetupComponent;
        try {selectedComponent = resolveComponent(choices, text);} catch (failure) {setError(message(failure)); return;}
        if (await commit(draft => {
            if (componentDraft.kind === 'component') return withComponent(saved.catalogue, draft, componentDraft.groupId, selectedComponent.id, makeDefault, componentDraft.replacingId);
            const target = draft.links.find(item => item.id === componentDraft.linkId)?.conditions.find(item => item.id === componentDraft.conditionId);
            if (!target) throw new Error('This condition is no longer available.');
            const value: SetupResult = {id: componentDraft.resultId!, componentId: selectedComponent.id, order: componentDraft.order!};
            const index = target.results.findIndex(item => item.id === value.id); if (index < 0) target.results.push(value); else target.results[index] = value;
        })) closeComponent();
    }, [componentDraft, saved, choices, commit, closeComponent]);
    const renderCell = React.useCallback((record: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord, column: string): React.ReactNode | undefined => {
        if (!componentDraft || record.getRecordId() !== componentDraft.rowId || column !== 'name') return undefined;
        return <ComponentEntryRow key={componentDraft.rowId} id={componentDraft.rowId} choices={choices} initial={componentInput.current} allowDefault={componentDraft.kind === 'component'} busy={busy}
            onChange={value => {componentInput.current = value; setComponentDirty(JSON.stringify(value) !== JSON.stringify(componentInitial.current));}} onSave={(text, makeDefault) => void saveComponent(text, makeDefault)} onCancel={closeComponent}/>;
    }, [componentDraft, choices, busy, saveComponent, closeComponent]);
    const commandAction = React.useRef<(command: SetupCommand) => void>(() => undefined);
    const executeCommand = React.useCallback((command: SetupCommand) => commandAction.current(command), []);
    React.useEffect(() => {
        if (!nativeHost || !saved || !document) return;
        let active = true;
        nativeHost.publish({ready:false, selected:[], canNew:false, canEdit:false, canDelete:false, canActivate:false, canDeactivate:false, canRefresh:false, execute:executeCommand});
        void nativeHost.select(selected, selected.filter(id => !!byId.get(id)?.parentId)).then(() => {
            if (!active) return;
            const option = document.options.find(item => item.groupId === row?.groupId);
            const eligibleOption = !!row && ['Option','Component'].includes(row.kind) && componentChoices(saved.catalogue, document, row.groupId).length > 0;
            let removable = !!row && (row.kind !== 'Option' || !!option && isCustomOption(option));
            if (row && removable) {try {withoutSetupRow(saved.catalogue, document, row);} catch {removable = false;}}
            let canDeactivate = row?.kind === 'Option' && !!option?.enabled;
            if (canDeactivate) {try {const next = clone(document); next.options.find(item => item.groupId === row!.groupId)!.enabled = false; validateSetup(saved.catalogue, next);} catch {canDeactivate = false;}}
            const unlinked = document.options.some(option => option.enabled && !document.links.some(item => item.groupId === option.groupId));
            const resultAvailable = !!link && !!condition && availableComponents(saved.catalogue, document, link.groupId).some(component => !condition.results.some(result => result.componentId === component.id));
            nativeHost.publish({ready:!locked, selected, canNew:!locked && selected.length <= 1 && (mode === 'options' ? selected.length === 0 && document.options.length < 500 || eligibleOption : selected.length === 0 ? unlinked : condition ? resultAvailable : !!link),
                canEdit:!locked && !!row && (row.kind !== 'Component' || !!option?.enabled), canDelete:!locked && removable, canActivate:!locked && row?.kind === 'Option' && option?.enabled === false,
                canDeactivate:!locked && canDeactivate, canRefresh:!locked, execute:executeCommand});
        }).catch(failure => {if (active) setError(message(failure));});
        return () => {active = false;};
    }, [nativeHost, saved, document, selected, byId, row, link, mode, locked, executeCommand]);
    commandAction.current = command => {
        if (!saved || !document || locked) return;
        if (command === 'refresh') {void load(); nativeHost?.refreshed(); return;}
        if (command === 'edit') {edit(); return;}
        if (command === 'delete') {if (row) {setError(''); setConfirm(row);} return;}
        if (command === 'activate' || command === 'deactivate') {if (row?.kind === 'Option') void commit(draft => {draft.options.find(item => item.groupId === row.groupId)!.enabled = command === 'activate';}); return;}
        if (mode === 'options') {
            if (row && ['Option','Component'].includes(row.kind)) beginComponent({kind:'component', rowId:id(), parentId:`option-${row.groupId}`, groupId:row.groupId});
            else if (!row && document.options.length < 500) openEditor({kind:'option', value:{groupId:`custom-${crypto.randomUUID()}`, name:'', sourceGroupId:'', order:Math.min(100000, Math.max(0,...document.options.map(option=>option.order))+1), enabled:true, hidden:false, freeText:false, componentIds:[], defaultId:null}});
        } else if (condition && link) {const resultId=id(); beginComponent({kind:'result',rowId:`result-${resultId}`,parentId:`condition-${condition.id}`,groupId:link.groupId,linkId:link.id,conditionId:condition.id,resultId,order:Math.max(0,...condition.results.map(item=>item.order))+1});}
        else if (link) openEditor({kind:'condition',linkId:link.id,value:newCondition(id(),Math.max(0,...link.conditions.map(item=>item.order))+1)});
        else {const target=document.options.find(option=>option.enabled&&!document.links.some(item=>item.groupId===option.groupId));if(target)openEditor({kind:'link',value:{id:id(),groupId:target.groupId,order:Math.max(0,...document.links.map(item=>item.order))+1,conditions:[]}});}
    };
    if (!saved || !document) return <main className="blindtype-setup"><h1>Blind type setup</h1><p role={error ? 'alert' : 'status'}>{error || 'Loading setup…'}</p>{error && <button onClick={() => void load()}>Try again</button>}</main>;
    const catalogue = saved.catalogue;
    const canAddComponent = mode === 'options' && !!row && ['Option', 'Component'].includes(row.kind) && componentChoices(catalogue, document, row.groupId).length > 0;
    return <div className={`blindtype-setup${nativeHost ? ' setup-native-host' : ''}`}>
        <header className="setup-heading"><div><p>Blind type setup</p><h1>{catalogue.blindType.name}</h1></div><span role="status">{busy ? 'Saving…' : status || (store.kind === 'demo' ? 'Local preview' : 'Saved in Dynamics')}</span></header>
        <div className="setup-tabs" role="group" aria-label="Setup view"><button disabled={locked} aria-pressed={mode === 'options'} onClick={() => {setMode('options'); setSelected([]); setFiltering(undefined); setSorting([]);}}>Options</button><button disabled={locked} aria-pressed={mode === 'links'} onClick={() => {setMode('links'); setSelected([]); setFiltering(undefined); setSorting([]);}}>Linked options</button><span>{mode === 'options' ? 'Choose available components and defaults. Save each edit directly.' : 'Set conditions and results. Save each edit directly.'}</span></div>
        {!nativeHost && <div className="setup-commandbar" role="toolbar" aria-label="Blind type setup commands">
            {mode === 'options' && <>
                <button disabled={locked || document.options.length >= 500} onClick={() => openEditor({kind: 'option', value: {groupId: `custom-${crypto.randomUUID()}`, name: '', sourceGroupId: '', order: Math.min(100000, Math.max(0, ...document.options.map(option => option.order)) + 1), enabled: true, hidden: false, freeText: false, componentIds: [], defaultId: null}})}>Add option</button>
                <button disabled={locked || !canAddComponent} onClick={() => {if (row) beginComponent({kind: 'component', rowId: id(), parentId: `option-${row.groupId}`, groupId: row.groupId});}}>Add component</button>
            </>}
            <button disabled={!row || locked} onClick={edit}>{row?.kind === 'Result' ? 'Choose component' : 'Edit selected'}</button>
            {mode === 'options' && <button disabled={row?.kind !== 'Component' || locked} onClick={() => {if (row) void commit(draft => {draft.options.find(item => item.groupId === row.groupId)!.defaultId = row.componentId!;});}}>Set as default</button>}
            {mode === 'links' && <>
                <button disabled={locked || !document.options.some(option => option.enabled && !document.links.some(item => item.groupId === option.groupId))} onClick={() => {const target = document.options.find(option => option.enabled && !document.links.some(item => item.groupId === option.groupId)); if (target) openEditor({kind: 'link', value: {id: id(), groupId: target.groupId, order: Math.max(0, ...document.links.map(item => item.order)) + 1, conditions: []}});}}>Add linked option</button>
                <button disabled={!link || locked} onClick={() => {if (link) openEditor({kind: 'condition', linkId: link.id, value: newCondition(id(), Math.max(0, ...link.conditions.map(item => item.order)) + 1)});}}>Add condition</button>
                <button disabled={!link || !condition || locked} onClick={() => {if (link && condition) {const resultId = id(); beginComponent({kind: 'result', rowId: `result-${resultId}`, parentId: `condition-${condition.id}`, groupId: link.groupId, linkId: link.id, conditionId: condition.id, resultId, order: Math.max(0, ...condition.results.map(item => item.order)) + 1});}}}>Add result</button>
            </>}
            <button disabled={!row || locked} onClick={() => {setError(''); setConfirm(row);}}>{row?.kind === 'Option' ? isCustomOption(document.options.find(option => option.groupId === row.groupId)!) ? 'Remove option' : 'Disable option' : 'Remove selected'}</button>
        </div>}
        {error && <div className="setup-error" role="alert">{error}</div>}
        {confirm && <section className="setup-confirm" aria-label="Confirm removal"><span>Remove {confirm.name}? Child conditions and results are included. This saves immediately.</span><button disabled={busy} onClick={() => void remove()}>Remove</button><button disabled={busy} onClick={() => {setConfirm(undefined); setError('');}}>Cancel</button></section>}
        {editor && <SetupEditor editor={editor} setEditor={setEditor} catalogue={catalogue} document={document} error={error} busy={busy} onSave={() => void saveEditor()} onCancel={() => {if (!busy) {setEditor(undefined); setError('');}}}/>}
        <div className="setup-grid"><ResponsiveGrid key={mode} columns={setupColumns} records={records} sortedRecordIds={ids}
            layoutKey={`blindtype-setup:${saved.id}:${mode}:v1`} features={{showTypeIcons: true, locale: 'en-GB', accessibleLabel: 'Blind type setup'}}
            childRows={{loadPage: loadChildren, hasChildren, revision, readOnly: true, accessibleLabel: 'Components and linked options'}} selectionResetVersion={revision} renderCell={renderCell} expandChildRequest={expandRequest}
            hasNextPage={false} hasPreviousPage={false} totalResultCount={visible.length} currentPage={1} sorting={sorting} filtering={filtering!} resources={resources} itemsLoading={busy}
            highlightValue={null} highlightColor={null} enableLookupLinks={false} enableInlineEdit={false} setSelectedRecords={setSelected}
            onNavigate={() => undefined} onOpenLookup={() => undefined} onSelectAllRecords={request => {checkSelectionCancelled(request?.signal); return selectionResult(ids);}}
            onFilter={(name, operation, value) => setFiltering(current => updateColumnFilter(current, name, operation, value))} onClearAllFilters={() => setFiltering(undefined)}
            onSort={(name, direction, additive) => setSorting(current => {const next = additive || direction === 'none' ? current.filter(item => item.name !== name) : []; if (direction !== 'none') next.push({name, sortDirection: direction === 'desc' ? 1 : 0}); return next;})}
            loadFirstPage={() => undefined} loadNextPage={() => undefined} loadPreviousPage={() => undefined} onUpdateCell={() => {throw new Error('Use the setup editors to change this row.');}}
        /></div>
        <div className="setup-footer">{selected.length} selected · {document.options.length} options · {catalogue.groups.reduce((sum, group) => sum + group.components.filter(component => component.available).length, 0)} components available for {catalogue.blindType.code}</div>
    </div>;
}
function SetupEditor({editor, setEditor, catalogue, document, busy, onSave, onCancel}: {editor: Editor; setEditor: (value: Editor) => void; catalogue: Catalogue; document: SetupDocument; error: string; busy: boolean; onSave: () => void; onCancel: () => void}): React.ReactElement {
    const [search, setSearch] = React.useState('');
    const linked = editor.kind === 'condition' ? document.links.find(item => item.id === editor.linkId)! : undefined;
    const groupId = editor.kind === 'option' ? editor.value.groupId : editor.kind === 'link' ? editor.value.groupId : linked!.groupId;
    const isNewOption = editor.kind === 'option' && !document.options.some(option => option.groupId === groupId);
    const custom = editor.kind === 'option' && isCustomOption(editor.value);
    const group = editor.kind === 'option'
        ? {...(catalogue.groups.find(source => source.id === (editor.value.sourceGroupId ?? groupId)) ?? {id: groupId, name: 'New option', order: editor.value.order, components: []}), id: groupId, ...(editor.value.name ? {name: editor.value.name} : {})}
        : optionGroup(catalogue, document, groupId);
    const patch = (values: object): void => setEditor({...editor, value: {...editor.value, ...values}} as Editor);
    const numberField = (label: string, name: string, value: number | null, nullable = false) => <label>{label}<input aria-label={label} type="number" min="0" step="any" value={value ?? ''} onChange={event => patch({[name]: event.target.value === '' && nullable ? null : Number(event.target.value)})}/></label>;
    const order = numberField('Order', 'order', editor.value.order);
    const title = editor.kind === 'option' ? isNewOption ? 'Add option' : `Set up ${group.name}` : editor.kind === 'link' ? 'Set up linked option' : `${group.name} condition`;
    return <section className="setup-edit-panel" aria-label={title}>
        <h2>{title}</h2><p>{catalogue.blindType.code} · Save this edit when you are ready.</p>
        <div className="setup-editor"><fieldset disabled={busy}>
            {editor.kind === 'option' && <>
                {custom && <div className="setup-fields"><label>Option name<input aria-label="Option name" maxLength={100} value={editor.value.name ?? ''} onChange={event => patch({name: event.target.value})}/></label><label>Component type<select aria-label="Component type" value={editor.value.sourceGroupId ?? ''} onChange={event => {setSearch(''); patch({sourceGroupId: event.target.value, componentIds: [], defaultId: null});}}><option value="">Choose component type</option>{catalogue.groups.map(source => <option key={source.id} value={source.id} disabled={!source.components.some(component => component.available)}>{source.name}</option>)}</select></label></div>}
                <div className="setup-fields">{order}<label>Default component<select aria-label="Default component" value={editor.value.defaultId ?? ''} onChange={event => patch({defaultId: event.target.value || null})}><option value="">No default</option>{group.components.filter(component => editor.value.componentIds.includes(component.id)).map(component => <option key={component.id} value={component.id}>{component.code} — {component.name}</option>)}</select></label></div>
                <div className="setup-checks"><label><input type="checkbox" checked={editor.value.enabled} onChange={event => patch({enabled: event.target.checked})}/>Enabled</label><label><input type="checkbox" checked={!editor.value.hidden} onChange={event => patch({hidden: !event.target.checked})}/>Show option</label><label><input type="checkbox" checked={editor.value.freeText} onChange={event => patch({freeText: event.target.checked})}/>Allow free text</label></div>
                <h3>Components available for {catalogue.blindType.code}</h3><input aria-label="Search available components" placeholder="Search code or description" value={search} onChange={event => setSearch(event.target.value)}/>
                <div className="setup-component-list">{group.components.filter(component => component.available && `${component.code} ${component.name}`.toLocaleLowerCase().includes(search.toLocaleLowerCase())).map(component => <label key={component.id}><input type="checkbox" checked={editor.value.componentIds.includes(component.id)} onChange={event => {const ids = event.target.checked ? [...editor.value.componentIds, component.id] : editor.value.componentIds.filter(item => item !== component.id); patch({componentIds: ids, defaultId: editor.value.defaultId && ids.includes(editor.value.defaultId) ? editor.value.defaultId : null});}}/><span><strong>{component.code}</strong> {component.name}</span></label>)}{!group.components.length && <p>{custom && !editor.value.sourceGroupId ? 'Choose a component type to see the available components.' : 'No components are available for this option.'}</p>}</div>
            </>}
            {editor.kind === 'link' && <div className="setup-fields"><label>Target option<select aria-label="Target option" disabled={!!editor.value.conditions.length} value={editor.value.groupId} onChange={event => patch({groupId: event.target.value})}>{document.options.filter(option => option.enabled && !document.links.some(link => link.groupId === option.groupId && link.id !== editor.value.id)).map(option => <option key={option.groupId} value={option.groupId}>{optionGroup(catalogue, document, option.groupId).name}</option>)}</select></label>{order}</div>}
            {editor.kind === 'condition' && <>
                <div className="setup-fields">{order}<label>Match<select aria-label="Match" value={editor.value.mode} onChange={event => patch({mode: event.target.value})}><option value="all">All of these choices (AND)</option><option value="any">Any of these choices (OR)</option></select></label></div>
                <h3>When these options are selected</h3>
                {editor.value.tests.map((test, index) => <div className="setup-condition-test" key={index}><label>Option<select aria-label={`Condition option ${index + 1}`} value={test.groupId} onChange={event => patch({tests: editor.value.tests.map((item, current) => current === index ? {groupId: event.target.value, componentId: ''} : item)})}><option value="">Choose option</option>{document.options.filter(option => option.enabled && option.groupId !== groupId && !editor.value.tests.some((item, current) => current !== index && item.groupId === option.groupId)).map(option => <option key={option.groupId} value={option.groupId}>{optionGroup(catalogue, document, option.groupId).name}</option>)}</select></label><label>Component<select aria-label={`Condition component ${index + 1}`} value={test.componentId} onChange={event => patch({tests: editor.value.tests.map((item, current) => current === index ? {...item, componentId: event.target.value} : item)})}><option value="">Choose component</option>{test.groupId && availableComponents(catalogue, document, test.groupId).map(component => <option key={component.id} value={component.id}>{component.code} — {component.name}</option>)}</select></label><button aria-label={`Remove condition choice ${index + 1}`} onClick={() => patch({tests: editor.value.tests.filter((_, current) => current !== index)})}>Remove</button></div>)}
                <button disabled={editor.value.tests.length >= 20} onClick={() => patch({tests: [...editor.value.tests, {groupId: '', componentId: ''}]})}>Add option condition</button>
                <h3>Size limits</h3><div className="setup-fields">{numberField('Minimum width (mm)', 'minWidth', editor.value.minWidth)}{numberField('Maximum width (mm)', 'maxWidth', editor.value.maxWidth, true)}{numberField('Minimum drop (mm)', 'minDrop', editor.value.minDrop)}{numberField('Maximum drop (mm)', 'maxDrop', editor.value.maxDrop, true)}</div>
                <details><summary>Additional limits and formula</summary><div className="setup-fields">{numberField('Maximum area', 'maxArea', editor.value.maxArea)}{numberField('Minimum chain length', 'minChainLength', editor.value.minChainLength)}{numberField('Maximum chain length', 'maxChainLength', editor.value.maxChainLength, true)}{numberField('Minimum panels', 'minPanels', editor.value.minPanels)}{numberField('Maximum panels', 'maxPanels', editor.value.maxPanels, true)}</div><label>Formula<textarea aria-label="Formula" value={editor.value.formula ?? ''} onChange={event => patch({formula: event.target.value || null})}/></label></details>
            </>}
        </fieldset></div>
        <div className="setup-editor-actions"><button className="setup-save" disabled={busy} onClick={onSave}>{busy ? 'Saving…' : 'Save'}</button><button disabled={busy} onClick={onCancel}>Cancel</button></div>
    </section>;
}
