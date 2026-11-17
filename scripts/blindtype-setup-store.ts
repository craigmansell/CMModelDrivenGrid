import {Catalogue, clone, initialR200sSetup, SetupDocument, validateSetup} from './blindtype-setup-model';

export interface SetupState {id: string; name: string; catalogue: Catalogue; document: SetupDocument; etag: string}
export interface SetupStore {kind: 'demo' | 'dataverse'; load(): Promise<SetupState>; save(state: SetupState, document: SetupDocument): Promise<SetupState>}
const selectedColumns = 'cm_itemid,cm_name,cm_blindtypecataloguejson,cm_blindtypesetupjson';
export function dataverseStore(origin: string, id: string, request: typeof fetch = fetch): SetupStore {
    if (!/^https:\/\/[a-z0-9-]+\.crm[0-9]*\.dynamics\.com$/i.test(origin) || !/^[0-9a-f-]{36}$/i.test(id)) throw new Error('Open this setup from a valid Dynamics blind-type record.');
    const url = `${origin}/api/data/v9.2/cm_items(${id})?$select=${selectedColumns}`;
    const headers = {Accept: 'application/json', 'OData-Version': '4.0', 'OData-MaxVersion': '4.0'};
    const parse = (row: Record<string, unknown>): SetupState => {
        if (row.cm_itemid !== id || typeof row['@odata.etag'] !== 'string' || typeof row.cm_blindtypecataloguejson !== 'string' || typeof row.cm_blindtypesetupjson !== 'string') throw new Error('This blind type has no complete setup catalogue.');
        const catalogue = JSON.parse(row.cm_blindtypecataloguejson) as Catalogue;
        const document = JSON.parse(row.cm_blindtypesetupjson) as SetupDocument;
        validateSetup(catalogue, document);
        return {id, name: String(row.cm_name ?? catalogue.blindType.name), catalogue, document, etag: row['@odata.etag']};
    };
    const load = async (): Promise<SetupState> => {
        const response = await request(url, {credentials: 'include', headers});
        if (!response.ok) throw new Error(response.status === 403 ? 'You do not have permission to read this blind type.' : 'The blind-type setup could not be loaded from Dynamics.');
        return parse(await response.json() as Record<string, unknown>);
    };
    return {kind: 'dataverse', load, save: async (state, document) => {
        // Re-read the current catalogue before saving; never trust a stale or edited picker list.
        const current = await load();
        if (current.etag !== state.etag) throw new Error('Another user changed this blind type. Your entry has not been saved. Refresh this page to load the latest version before editing again.');
        validateSetup(current.catalogue, document);
        const response = await request(url, {method: 'PATCH', credentials: 'include', headers: {...headers, 'Content-Type': 'application/json', 'If-Match': state.etag, Prefer: 'return=representation'},
            body: JSON.stringify({cm_blindtypesetupjson: JSON.stringify(document)})});
        if (response.status === 412) throw new Error('Another user changed this blind type. Your entry has not overwritten theirs. Refresh this page to load the latest version before editing again.');
        if (!response.ok) throw new Error(response.status === 403 ? 'You do not have permission to save this blind type. Your entry is still available.' : 'Dynamics did not save this entry. It is still available to retry.');
        const saved = response.status === 204 ? await load() : parse(await response.json() as Record<string, unknown>);
        if (JSON.stringify(saved.document) !== JSON.stringify(document)) throw new Error('The saved version could not be confirmed. Refresh this page before editing again.');
        return saved;
    }};
}
export function demoSetupStore(): SetupStore {
    const seed = initialR200sSetup();
    let revision = 1;
    let state: SetupState = {id: 'r200s-local-demo', name: seed.catalogue.blindType.name, ...seed, etag: '1'};
    return {kind: 'demo', load: async () => clone(state), save: async (current, document) => {
        if (current.etag !== state.etag) throw new Error('This demo was changed. Reload before saving.');
        validateSetup(state.catalogue, document);
        state = {...state, document: clone(document), etag: String(++revision)};
        return clone(state);
    }};
}
