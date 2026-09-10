import { act, Simulate } from 'react-dom/test-utils';
import type { GridProps } from '../CMModelDrivenGrid/Grid';
import { CMModelDrivenGrid } from '../CMModelDrivenGrid';

type Row = ComponentFramework.PropertyHelper.DataSetApi.EntityRecord;
type Column = ComponentFramework.PropertyHelper.DataSetApi.Column;
type Context = ComponentFramework.Context<import('../CMModelDrivenGrid/generated/ManifestTypes').IInputs>;

/** Opt-in, read-only review probes. All host calls are local mocks, never Dataverse. */
export async function runPreSandboxReviewProbes(harness: {
    container: HTMLElement;
    resources: ComponentFramework.Resources;
    renderGrid: (props: Partial<GridProps>) => void;
    unmount: () => void;
}): Promise<void> {
    const failures: string[] = [];
    const check = (condition: boolean, message: string) => { if (!condition) throw new Error(message); };
    const probe = async (name: string, run: () => Promise<void>) => {
        try { await run(); console.log(`PASS: ${name}`); }
        catch (error) { const detail = error instanceof Error ? error.message : String(error); failures.push(name); console.error(`FAIL: ${name}: ${detail}`); }
        finally { await act(async () => { harness.unmount(); }); }
    };
    const row = {getRecordId: () => 'a', getValue: () => 'Alpha', getFormattedValue: () => 'Alpha'} as unknown as Row;
    const columns = ['name', 'second', 'third'].map((name, order) => ({name, displayName: name, dataType: 'SingleLine.Text', order})) as Column[];

    await probe('Compact End stays on the last rendered field', async () => {
        await act(async () => { harness.renderGrid({columns, records: {a: row}, sortedRecordIds: ['a'], features: {reflow: 'list', listColumns: 2}}); });
        const cell = harness.container.querySelector<HTMLElement>('[data-cm-cell-row="0"][data-cm-cell-col="0"]')!;
        check(!!cell, 'Compact starting cell must exist.');
        await act(async () => { cell.focus(); Simulate.keyDown(cell, {key: 'End'}); });
        await act(async () => { await new Promise(resolve => setTimeout(resolve, 150)); });
        const actual = document.activeElement?.getAttribute('data-cm-cell-col');
        check(actual === '1', `Expected visible column 1; focus remained at column ${actual}.`);
    });

    await probe('Rejected ordinary filter restores the previous state and releases loading', async () => {
        const controller = new CMModelDrivenGrid();
        let physical: ComponentFramework.PropertyHelper.DataSetApi.FilterExpression | undefined;
        let refreshes = 0;
        const dataset = {
            columns: [columns[0]], records: {a: row}, sortedRecordIds: ['a'], sorting: [], loading: false,
            getTargetEntityType: () => 'account', getViewId: () => 'view-1', setSelectedRecordIds: () => undefined,
            filtering: {getFilter: () => physical, setFilter: () => {throw new Error('Injected host filter rejection');}, clearFilter: () => {physical = undefined;}},
            paging: {hasNextPage: false, hasPreviousPage: false, totalResultCount: 1}, refresh: () => {refreshes++;},
        };
        controller.context = {mode: {allocatedWidth: 900, allocatedHeight: 500, isControlDisabled: false}, parameters: {
            records: dataset, SubGridHeight: {raw: 500}, HighlightValue: {raw: null}, HighlightColor: {raw: null},
            EnableLookupLinks: {raw: false}, EnableInlineEdit: {raw: true}, InfiniteScroll: {raw: true}, EnableJumpBar: {raw: true},
        }} as unknown as Context;
        controller.container = harness.container as HTMLDivElement;
        controller.resources = harness.resources;
        controller.records = {a: row}; controller.sortedRecordsIds = ['a'];
        const access = controller as unknown as {renderGrid: () => void; getRootFilter: () => unknown};
        await act(async () => { access.renderGrid(); });
        const letter = Array.from(harness.container.querySelectorAll<HTMLButtonElement>('.cm-grid-jump-bar button')).find(button => button.textContent === 'A')!;
        check(!!letter && !letter.disabled, 'The real controller must render an enabled alphabet filter.');
        await act(async () => { letter.click(); });
        await act(async () => { await new Promise(resolve => setTimeout(resolve, 150)); });
        const overlay = !!harness.container.querySelector('.cm-grid-loading-overlay');
        const logical = access.getRootFilter();
        check(!overlay && logical === physical, JSON.stringify({loadingOverlay: overlay, logicalFilter: logical, physicalFilter: physical ?? null, refreshes}));
    });

    await probe('Lookup save rechecks column read-only state after metadata', async () => {
        const controller = new CMModelDrivenGrid();
        const column = {name: 'lookup', dataType: 'Lookup.Simple', isReadOnly: false};
        let writes = 0;
        controller.context = {mode: {isControlDisabled: false}, parameters: {EnableInlineEdit: {raw: true}, records: {
            getTargetEntityType: () => 'account', columns: [column], refresh: () => undefined,
        }}, webAPI: {updateRecord: async () => {writes++;}}, navigation: {openAlertDialog: async () => undefined}} as unknown as Context;
        const access = controller as unknown as {getEntityMeta: () => Promise<{entitySetName: string}>; resolveLookupBindNames: () => Promise<string[]>};
        access.getEntityMeta = async () => ({entitySetName: 'contacts'});
        access.resolveLookupBindNames = async () => {column.isReadOnly = true; return ['lookup'];};
        try {await controller.onUpdateLookupCell('00000000-0000-4000-8000-000000000001', 'lookup', '00000000-0000-4000-8000-000000000002', 'contact');} catch { /* Rejection is the expected safe result. */ }
        check(writes === 0, `Expected zero writes after the column became read only; observed ${writes}.`);
    });

    if (failures.length) throw new Error(`${failures.length} supplementary pre-sandbox probes failed: ${failures.join('; ')}`);
}
