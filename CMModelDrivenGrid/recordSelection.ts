import { Selection, SelectionMode, IObjectWithKey } from '@fluentui/react/lib/Selection';

// Power Apps modern-grid Select All with infinite scrolling is capped at 1,000.
export const MAX_SELECTED_RECORDS = 1000;
export interface SelectAllRequest { signal?: AbortSignal }
export interface SelectAllResult { ids: string[]; limitReached: boolean }

export function selectionResult(ids: string[], hasMore = false): SelectAllResult {
    const unique = [...new Set(ids)];
    return {ids: unique.slice(0, MAX_SELECTED_RECORDS), limitReached: unique.length > MAX_SELECTED_RECORDS || (unique.length >= MAX_SELECTED_RECORDS && hasMore)};
}

export function checkSelectionCancelled(signal?: AbortSignal): void {
    if (signal?.aborted) throw new Error('Selection cancelled.');
}

type Row = ComponentFramework.PropertyHelper.DataSetApi.EntityRecord & IObjectWithKey;
export interface RecordSelection extends Selection { syncIds: (ids: string[]) => void }

/** Explicit IDs, never Fluent's open-ended "all current and future items" mode. */
export function createRecordSelection(options: {
    ids: () => string[];
    changed: (ids: string[]) => void;
    toggleAll: () => void;
}): RecordSelection {
    let syncing = false;
    const selection = new Selection({
        selectionMode: SelectionMode.multiple,
        getKey: item => (item as Row).getRecordId(),
        onSelectionChanged: () => {
            if (syncing) return;
            const visible = new Set((selection.getItems() as Row[]).map(item => item.getRecordId()));
            const selected = (selection.getSelection() as Row[]).map(item => item.getRecordId());
            options.changed([...options.ids().filter(id => !visible.has(id)), ...selected]);
        },
    }) as RecordSelection;
    const setAll = selection.setAllSelected.bind(selection);
    const setItems = selection.setItems.bind(selection);
    const isAllSelected = selection.isAllSelected.bind(selection);
    // The header remains ticked at capacity, without selecting later page items.
    selection.isAllSelected = () => options.ids().length >= MAX_SELECTED_RECORDS || isAllSelected();
    selection.syncIds = ids => {
        const wasSyncing = syncing;
        syncing = true;
        selection.setChangeEvents(false);
        try {
            setAll(false);
            ids.forEach(id => selection.setKeySelected(id, true, false));
        } finally {
            selection.setChangeEvents(true);
            syncing = wasSyncing;
        }
    };
    selection.setItems = items => {
        // A replaced/virtualized page must not publish a partial selection.
        syncing = true;
        selection.setChangeEvents(false);
        try { setItems(items, false); selection.syncIds(options.ids()); }
        finally { selection.setChangeEvents(true); syncing = false; }
    };
    selection.toggleAllSelected = () => options.toggleAll();
    selection.setAllSelected = selected => {
        if (selected) options.toggleAll();
        else { selection.syncIds([]); options.changed([]); }
    };
    return selection;
}
