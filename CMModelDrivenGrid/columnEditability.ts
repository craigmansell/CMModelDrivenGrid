/** Dataverse, not the editor's data type, decides whether an existing value can be updated. */
export interface UpdateMetadata {
    LogicalName?: string;
    IsValidForUpdate?: boolean | null;
    AttributeOf?: string | null;
    SourceType?: number | null;
}

export function editableAttribute(attribute: UpdateMetadata): boolean {
    return attribute.IsValidForUpdate === true && !attribute.AttributeOf &&
        (attribute.SourceType === undefined || attribute.SourceType === null || attribute.SourceType === 0);
}

export async function readColumnEditability(
    entity: string, columns: string[], origin: string,
    request: typeof fetch = fetch
): Promise<Record<string, boolean>> {
    if (!/^[a-z][a-z0-9_]*$/i.test(entity)) throw new Error('Invalid table name.');
    const names = [...new Set(columns.filter(name => /^[a-z][a-z0-9_]*$/i.test(name)).map(name => name.toLowerCase()))];
    const result: Record<string, boolean> = {};
    // Bounded URLs; missing attributes and incomplete/failed metadata stay read only.
    for (let offset = 0; offset < names.length; offset += 40) {
        const batch = names.slice(offset, offset + 40);
        const filter = batch.map(name => `LogicalName eq '${name}'`).join(' or ');
        const response = await request(`${origin}/api/data/v9.2/EntityDefinitions(LogicalName='${entity}')/Attributes` +
            `?$select=LogicalName,IsValidForUpdate,AttributeOf,SourceType&$filter=${encodeURIComponent(filter)}`, {
            credentials: 'include', headers: {Accept: 'application/json', 'OData-Version': '4.0', 'OData-MaxVersion': '4.0'},
        });
        if (!response.ok) throw new Error('Column editability metadata is unavailable.');
        const data = await response.json() as {value?: UpdateMetadata[]};
        if (!Array.isArray(data.value)) throw new Error('Invalid column editability metadata.');
        for (const attribute of data.value) {
            const name = attribute.LogicalName?.toLowerCase();
            if (name && batch.includes(name)) result[name] = editableAttribute(attribute);
        }
    }
    return result;
}
