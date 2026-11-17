import * as React from 'react';
import {Dialog, DialogFooter} from '@fluentui/react/lib/Dialog';
import {DefaultButton, PrimaryButton} from '@fluentui/react/lib/Button';
import {ResponsiveGrid} from '../CMModelDrivenGrid/ResponsiveGrid';
import {ChildRequest} from '../CMModelDrivenGrid/childRecords';
import {FilterExpression, updateColumnFilter} from '../CMModelDrivenGrid/filtering';
import {selectionResult, checkSelectionCancelled} from '../CMModelDrivenGrid/recordSelection';
import {evaluateDemoFilter, DemoFilterExpression, sortDemoRows} from './demo-filter-evaluator';
import {pageDemoRows} from './demo-data';
import {availableComponents, Catalogue, clone, isCustomOption, optionGroup, LinkedSetup, newCondition, OptionSetup, SetupCondition, setupColumns, SetupDocument, setupRecord, SetupResult, setupRows, validateSetup, withOption, withoutOption} from './blindtype-setup-model';
import {SetupState, SetupStore} from './blindtype-setup-store';
import '../CMModelDrivenGrid/css/CMModelDrivenGrid.css';
import './blindtype-setup.css';

type Editor = {kind: 'option'; value: OptionSetup} | {kind: 'link'; value: LinkedSetup} |
    {kind: 'condition'; linkId: string; value: SetupCondition} | {kind: 'result'; linkId: string; conditionId: string; value: SetupResult};
const kinds = {name: 'text', kind: 'text', code: 'text', order: 'number', details: 'text', count: 'number'} as const;
const resources = {getString: (key: string) => ({Label_Grid_Footer_RecordCount: '{0} records ({1} Selected)', Label_Grid_Footer: 'Page {0}', Label_NoRecords: 'No matching records', Label_SortAZ: 'A to Z', Label_SortZA: 'Z to A', Label_DoesNotContainData: 'Does not contain data'}[key] ?? key), getResource: () => undefined} as unknown as ComponentFramework.Resources;
const id = (): string => `draft-${crypto.randomUUID()}`;
const query = <T extends object,>(rows: T[], filtering: FilterExpression | undefined, sorting: ChildRequest['sorting']): T[] => {
    const getValue = (row: T, column: string): unknown => row[column as keyof T];
    return sortDemoRows(rows.filter(row => evaluateDemoFilter(row, filtering as unknown as DemoFilterExpression, {getValue, columnKinds: kinds})),
        sorting.map(sort => ({columnName: sort.name, direction: sort.sortDirection === 1 ? 'desc' as const : 'asc' as const})), {getValue, getDisplayValue: (row, column) => String(getValue(row, column) ?? ''), columnKinds: kinds});
};
const message = (error: unknown): string => error instanceof Error ? error.message : 'The setup could not be changed.';

export function BlindTypeSetup({store}: {store: SetupStore}): React.ReactElement {
    const [saved, setSaved] = React.useState<SetupState>();
    const [document, setDocument] = React.useState<SetupDocument>();
    const [mode, setMode] = React.useState<'options' | 'links'>('options');
    const [history, setHistory] = React.useState<SetupDocument[]>([]);
    const [revision, setRevision] = React.useState(0);
    const [generation, setGeneration] = React.useState(0);
    const [selected, setSelected] = React.useState<string[]>([]);
    const [filtering, setFiltering] = React.useState<FilterExpression>();
    const [sorting, setSorting] = React.useState<ChildRequest['sorting']>([]);
    const [editor, setEditor] = React.useState<Editor>();
    const [confirm, setConfirm] = React.useState<'remove' | 'reload'>();
    const [busy, setBusy] = React.useState(false);
    const [error, setError] = React.useState('');
    const [status, setStatus] = React.useState('');
    const reload = React.useCallback(async () => {setBusy(true); setError(''); try {const state = await store.load(); setSaved(state); setDocument(clone(state.document)); setHistory([]); setSelected([]); setRevision(value => value + 1); setGeneration(value => value + 1); setStatus('');} catch (failure) {setError(message(failure));} finally {setBusy(false);}}, [store]);
    React.useEffect(() => {void reload();}, [reload]);
    const dirty = !!saved && JSON.stringify(document) !== JSON.stringify(saved.document);
    React.useEffect(() => {const prevent = (event: BeforeUnloadEvent): void => {if (dirty) {event.preventDefault(); event.returnValue = '';}}; window.addEventListener('beforeunload', prevent); return () => window.removeEventListener('beforeunload', prevent);}, [dirty]);
    const rows = React.useMemo(() => saved && document ? setupRows(saved.catalogue, document, mode) : [], [saved, document, mode]);
    const byId = React.useMemo(() => new Map(rows.map(row => [row.id, row])), [rows]);
    const branches = React.useMemo(() => {const index = new Map<string, typeof rows>(); for (const row of rows) if (row.parentId) index.set(row.parentId, [...index.get(row.parentId) ?? [], row]); return index;}, [rows]);
    const visible = React.useMemo(() => query(rows.filter(row => !row.parentId), filtering, sorting), [rows, filtering, sorting]);
    const records = React.useMemo(() => Object.fromEntries(visible.map(row => [row.id, setupRecord(row)])), [visible]);
    const ids = React.useMemo(() => visible.map(row => row.id), [visible]);
    const loadChildren = React.useCallback(async (request: ChildRequest) => {const page = pageDemoRows(query(branches.get(request.parentId) ?? [], request.filtering, request.sorting), request.page); return {records: page.rows.map(setupRecord), hasNextPage: page.hasNextPage};}, [branches]);
    const hasChildren = React.useCallback(async (parents: string[]) => Object.fromEntries(parents.map(parent => [parent, !!branches.get(parent)?.length])), [branches]);
    const row = selected.length === 1 ? byId.get(selected[0]) : undefined;
    const link = document?.links.find(item => item.id === row?.linkId);
    const condition = link?.conditions.find(item => item.id === row?.conditionId);
    const change = (update: (draft: SetupDocument) => void | SetupDocument): boolean => {
        if (!saved || !document || busy) return false;
        try {const draft = clone(document); const next = update(draft) || draft; validateSetup(saved.catalogue, next); setHistory(items => [...items.slice(-19), document]); setDocument(next); setRevision(value => value + 1); setSelected([]); setError(''); setStatus('Unsaved changes'); return true;}
        catch (failure) {setError(message(failure)); return false;}
    };
    const edit = (): void => {
        if (!row || !document) return;
        if (row.kind === 'Option' || row.kind === 'Component') setEditor({kind: 'option', value: clone(document.options.find(item => item.groupId === row.groupId)!)});
        else if (row.kind === 'Linked option' && link) setEditor({kind: 'link', value: clone(link)});
        else if (row.kind === 'Condition' && link && condition) setEditor({kind: 'condition', linkId: link.id, value: clone(condition)});
        else if (row.kind === 'Result' && link && condition) setEditor({kind: 'result', linkId: link.id, conditionId: condition.id, value: clone(condition.results.find(item => item.id === row.resultId)!)});
    };
    const saveEditor = (): void => {
        if (!editor || !saved) return;
        if (change(draft => {
            if (editor.kind === 'option') return withOption(saved.catalogue, draft, editor.value);
            else if (editor.kind === 'link') {const index = draft.links.findIndex(item => item.id === editor.value.id); if (index < 0) draft.links.push(clone(editor.value)); else draft.links[index] = clone(editor.value);}
            else if (editor.kind === 'condition') {const target = draft.links.find(item => item.id === editor.linkId)!; const index = target.conditions.findIndex(item => item.id === editor.value.id); if (index < 0) target.conditions.push(clone(editor.value)); else target.conditions[index] = clone(editor.value);}
            else {const target = draft.links.find(item => item.id === editor.linkId)!.conditions.find(item => item.id === editor.conditionId)!; const index = target.results.findIndex(item => item.id === editor.value.id); if (index < 0) target.results.push(clone(editor.value)); else target.results[index] = clone(editor.value);}
        })) setEditor(undefined);
    };
    const remove = (): void => {
        if (!row || !saved) return;
        change(draft => {
            if (row.kind === 'Linked option') draft.links = draft.links.filter(item => item.id !== row.linkId);
            else if (row.kind === 'Condition') draft.links.find(item => item.id === row.linkId)!.conditions = draft.links.find(item => item.id === row.linkId)!.conditions.filter(item => item.id !== row.conditionId);
            else if (row.kind === 'Result') {const target = draft.links.find(item => item.id === row.linkId)!.conditions.find(item => item.id === row.conditionId)!; target.results = target.results.filter(item => item.id !== row.resultId);}
            else if (row.kind === 'Option') {
                if (isCustomOption(draft.options.find(item => item.groupId === row.groupId)!)) return withoutOption(saved.catalogue, draft, row.groupId);
                draft.options.find(item => item.groupId === row.groupId)!.enabled = false;
            }
            else {const option = draft.options.find(item => item.groupId === row.groupId)!; option.componentIds = option.componentIds.filter(item => item !== row.componentId); if (option.defaultId === row.componentId) option.defaultId = null;}
        }); setConfirm(undefined);
    };
    const save = async (): Promise<void> => {
        if (!saved || !document || busy) return;
        setBusy(true); setError('');
        try {const state = await store.save(saved, document); setSaved(state); setDocument(clone(state.document)); setHistory([]); setStatus(store.kind === 'dataverse' ? 'Saved in Dynamics' : 'Demo saved');}
        catch (failure) {setError(message(failure));} finally {setBusy(false);}
    };
    if (!saved || !document) return <main className="blindtype-setup"><h1>Blind type setup</h1><p role={error ? 'alert' : 'status'}>{error || 'Loading setup…'}</p>{error && <button onClick={() => void reload()}>Try again</button>}</main>;
    const catalogue = saved.catalogue;
    return <div className="blindtype-setup">
        <header className="setup-heading"><div><p>Blind type setup</p><h1>{catalogue.blindType.name}</h1></div><span role="status">{busy ? 'Working…' : status || (store.kind === 'demo' ? 'Local preview' : 'Saved in Dynamics')}</span></header>
        <div className="setup-tabs" role="group" aria-label="Setup view"><button aria-pressed={mode === 'options'} onClick={() => {setMode('options'); setSelected([]); setFiltering(undefined); setSorting([]);}}>Options</button><button aria-pressed={mode === 'links'} onClick={() => {setMode('links'); setSelected([]); setFiltering(undefined); setSorting([]);}}>Linked options</button><span>{mode === 'options' ? 'Add options, choose allowed components and set defaults.' : 'Set conditions and the component choices they return.'}</span></div>
        <div className="setup-commandbar" role="toolbar" aria-label="Blind type setup commands">
            {mode === 'options' && <button disabled={busy || document.options.length >= 500} onClick={() => {setError(''); setEditor({kind: 'option', value: {groupId: `custom-${crypto.randomUUID()}`, name: '', sourceGroupId: '', order: Math.max(0, ...document.options.map(option => option.order)) + 1, enabled: true, hidden: false, freeText: false, componentIds: [], defaultId: null}});}}>Add option</button>}
            <button disabled={!row || busy} onClick={edit}>{row?.kind === 'Result' ? 'Choose component' : 'Edit selected'}</button>
            {mode === 'options' && <button disabled={row?.kind !== 'Component' || busy} onClick={() => {if (row) change(draft => {draft.options.find(item => item.groupId === row.groupId)!.defaultId = row.componentId!;});}}>Set as default</button>}
            {mode === 'links' && <>
                <button disabled={busy || !document.options.some(option => option.enabled && !document.links.some(item => item.groupId === option.groupId))} onClick={() => {const target = document.options.find(option => option.enabled && !document.links.some(item => item.groupId === option.groupId)); if (target) setEditor({kind: 'link', value: {id: id(), groupId: target.groupId, order: Math.max(0, ...document.links.map(item => item.order)) + 1, conditions: []}});}}>Add linked option</button>
                <button disabled={!link || busy} onClick={() => {if (link) setEditor({kind: 'condition', linkId: link.id, value: newCondition(id(), Math.max(0, ...link.conditions.map(item => item.order)) + 1)});}}>Add condition</button>
                <button disabled={!link || !condition || busy} onClick={() => {if (link && condition) setEditor({kind: 'result', linkId: link.id, conditionId: condition.id, value: {id: id(), componentId: '', order: Math.max(0, ...condition.results.map(item => item.order)) + 1}});}}>Add result</button>
            </>}
            <button disabled={!row || busy} onClick={() => setConfirm('remove')}>{row?.kind === 'Option' ? isCustomOption(document.options.find(option => option.groupId === row.groupId)!) ? 'Remove option' : 'Disable option' : 'Remove selected'}</button>
            <span className="setup-command-spacer"/>
            <button disabled={!history.length || busy} onClick={() => {setDocument(clone(history[history.length - 1])); setHistory(items => items.slice(0, -1)); setRevision(value => value + 1); setSelected([]); setError(''); setStatus('Unsaved changes');}}>Undo</button>
            <button disabled={busy} onClick={() => dirty ? setConfirm('reload') : void reload()}>Reload</button>
            <button className="setup-save" disabled={!dirty || busy} onClick={() => void save()}>{store.kind === 'dataverse' ? 'Save configuration' : 'Save demo'}</button>
        </div>
        {error && <div className="setup-error" role="alert">{error}</div>}
        <div className="setup-grid"><ResponsiveGrid key={`${generation}:${mode}`} columns={setupColumns} records={records} sortedRecordIds={ids}
            layoutKey={`blindtype-setup:${saved.id}:${mode}:v1`} features={{showTypeIcons: true, locale: 'en-GB', accessibleLabel: 'Blind type setup'}}
            childRows={{loadPage: loadChildren, hasChildren, revision, readOnly: true}} selectionResetVersion={revision}
            hasNextPage={false} hasPreviousPage={false} totalResultCount={visible.length} currentPage={1} sorting={sorting} filtering={filtering!} resources={resources} itemsLoading={busy}
            highlightValue={null} highlightColor={null} enableLookupLinks={false} enableInlineEdit={false} setSelectedRecords={setSelected}
            onNavigate={() => undefined} onOpenLookup={() => undefined} onSelectAllRecords={request => {checkSelectionCancelled(request?.signal); return selectionResult(ids);}}
            onFilter={(name, operation, value) => setFiltering(current => updateColumnFilter(current, name, operation, value))} onClearAllFilters={() => setFiltering(undefined)}
            onSort={(name, direction, additive) => setSorting(current => {const next = additive || direction === 'none' ? current.filter(item => item.name !== name) : []; if (direction !== 'none') next.push({name, sortDirection: direction === 'desc' ? 1 : 0}); return next;})}
            loadFirstPage={() => undefined} loadNextPage={() => undefined} loadPreviousPage={() => undefined} onUpdateCell={() => {throw new Error('Use the setup editors to change this row.');}}
        /></div>
        <div className="setup-footer">{selected.length} selected · {document.options.length} options · {catalogue.groups.reduce((sum, group) => sum + group.components.filter(component => component.available).length, 0)} components available for {catalogue.blindType.code}</div>
        {editor && <SetupEditor editor={editor} setEditor={setEditor} catalogue={catalogue} document={document} error={error} onSave={saveEditor} onCancel={() => {setEditor(undefined); setError('');}}/>}
        <Dialog hidden={!confirm} onDismiss={() => setConfirm(undefined)} dialogContentProps={{title: confirm === 'remove' ? `Remove ${row?.name ?? 'selection'}?` : 'Discard unsaved changes?', subText: confirm === 'remove' ? 'Child conditions and results are included. You can undo this before saving.' : 'Reload the last saved configuration from Dynamics.'}}>
            <DialogFooter><PrimaryButton text={confirm === 'remove' ? 'Remove' : 'Discard and reload'} onClick={() => {if (confirm === 'remove') remove(); else {setConfirm(undefined); void reload();}}}/><DefaultButton text="Cancel" onClick={() => setConfirm(undefined)}/></DialogFooter>
        </Dialog>
    </div>;
}

function SetupEditor({editor, setEditor, catalogue, document, error, onSave, onCancel}: {editor: Editor; setEditor: (value: Editor) => void; catalogue: Catalogue; document: SetupDocument; error: string; onSave: () => void; onCancel: () => void}): React.ReactElement {
    const [search, setSearch] = React.useState('');
    const linked = editor.kind === 'condition' || editor.kind === 'result' ? document.links.find(item => item.id === editor.linkId)! : undefined;
    const groupId = editor.kind === 'option' ? editor.value.groupId : editor.kind === 'link' ? editor.value.groupId : linked!.groupId;
    const isNewOption = editor.kind === 'option' && !document.options.some(option => option.groupId === groupId);
    const custom = editor.kind === 'option' && isCustomOption(editor.value);
    const group = editor.kind === 'option'
        ? {...(catalogue.groups.find(source => source.id === (editor.value.sourceGroupId ?? groupId)) ?? {id: groupId, name: 'New option', order: editor.value.order, components: []}), id: groupId, ...(editor.value.name ? {name: editor.value.name} : {})}
        : optionGroup(catalogue, document, groupId);
    const patch = (values: object): void => setEditor({...editor, value: {...editor.value, ...values}} as Editor);
    const numberField = (label: string, name: string, value: number | null, nullable = false) => <label>{label}<input aria-label={label} type="number" min="0" step="any" value={value ?? ''} onChange={event => patch({[name]: event.target.value === '' && nullable ? null : Number(event.target.value)})}/></label>;
    const order = numberField('Order', 'order', editor.value.order);
    const title = editor.kind === 'option' ? isNewOption ? 'Add option' : `Set up ${group.name}` : editor.kind === 'link' ? 'Set up linked option' : editor.kind === 'condition' ? `${group.name} condition` : `Choose ${group.name} component`;
    return <Dialog hidden={false} onDismiss={onCancel} modalProps={{isBlocking: true, styles: {main: {width: 740, maxWidth: '96vw'}}}} maxWidth={800} dialogContentProps={{title, subText: `${catalogue.blindType.code} · changes are applied to your draft until you save the configuration.`}}>
        <div className="setup-editor">
            {error && <div className="setup-error" role="alert">{error}</div>}
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
            {editor.kind === 'result' && <>
                <p>Only components enabled for {catalogue.blindType.code} → {group.name} are listed.</p><input aria-label="Search available components" placeholder="Search code or description" value={search} onChange={event => setSearch(event.target.value)}/>
                <label>Component<select size={7} aria-label="Component" value={editor.value.componentId} onChange={event => patch({componentId: event.target.value})}>{availableComponents(catalogue, document, groupId).filter(component => `${component.code} ${component.name}`.toLocaleLowerCase().includes(search.toLocaleLowerCase()) && !linked!.conditions.find(item => item.id === editor.conditionId)!.results.some(result => result.id !== editor.value.id && result.componentId === component.id)).map(component => <option key={component.id} value={component.id}>{component.code} — {component.name}</option>)}</select></label>{order}
            </>}
        </div>
        <DialogFooter><PrimaryButton text="Apply to draft" onClick={onSave}/><DefaultButton text="Cancel" onClick={onCancel}/></DialogFooter>
    </Dialog>;
}
