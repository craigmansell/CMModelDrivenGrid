import { childGuid, childLogicalName } from './childRecords';

interface ChildActionOptions {
    entity: string;
    idField: string;
    nameField: string;
    parentField: string;
    contextField?: string;
    webAPI: ComponentFramework.WebApi;
    navigation: ComponentFramework.Navigation;
}

/** Commands never use the hosting dataset selection. All targets are explicit. */
export function createChildActions(options: ChildActionOptions) {
    const {webAPI,navigation} = options;
    const entity = childLogicalName(options.entity);
    const idField = childLogicalName(options.idField);
    const nameField = childLogicalName(options.nameField);
    const parentField = childLogicalName(options.parentField);
    const contextField = options.contextField ? childLogicalName(options.contextField) : '';
    if (contextField === parentField) throw new Error('The template lookup must differ from the parent-option lookup.');
    return {
        add: async (parentId: string): Promise<boolean> => {
            const id = childGuid(parentId);
            const fields = [nameField, ...(contextField ? [`_${contextField}_value`] : [])];
            const parent = await webAPI.retrieveRecord(entity,id,`?$select=${fields.join(',')}`);
            const parameters: Record<string,string> = {[parentField]:id,[`${parentField}name`]:String(parent[nameField] ?? id)};
            if (contextField) {
                const key = `_${contextField}_value`;
                if (!parent[key]) throw new Error('The parent option has no template link. Set it before adding children.');
                parameters[contextField] = childGuid(String(parent[key]));
                parameters[`${contextField}name`] = String(parent[`${key}@OData.Community.Display.V1.FormattedValue`] ?? parameters[contextField]);
            }
            const result = await navigation.openForm({entityName:entity,useQuickCreateForm:true},parameters);
            return !!result?.savedEntityReference?.length;
        },
        delete: async (parentId: string, recordId: string): Promise<boolean> => {
            const parent = childGuid(parentId), id = childGuid(recordId);
            if (id === parent) throw new Error('A parent cannot be deleted through its child grid.');
            const readTarget = async () => {
                const row = await webAPI.retrieveRecord(entity,id,`?$select=${nameField},_${parentField}_value`);
                if (childGuid(String(row[`_${parentField}_value`] ?? '')) !== parent) throw new Error('This option no longer belongs to this parent. Refresh the children.');
                // Include inactive descendants too; never silently cascade a branch.
                const descendants = await webAPI.retrieveMultipleRecords(entity,`?$select=${idField}&$top=1&$filter=${encodeURIComponent(`_${parentField}_value eq ${id}`)}`);
                if (descendants.entities.length) throw new Error('This option has children. Remove its children first; branch deletion is not supported.');
                return row;
            };
            const row = await readTarget();
            const result = await navigation.openConfirmDialog({title:'Delete child option?',text:`Delete ${String(row[nameField] ?? id)}? This permanently deletes the record. Dataverse relationship rules may affect other related records.`,confirmButtonLabel:'Delete',cancelButtonLabel:'Cancel'});
            if (!result.confirmed) return false;
            await readTarget(); // Recheck relationship after the user spends time in the dialog.
            await webAPI.deleteRecord(entity,id);
            return true;
        },
    };
}
