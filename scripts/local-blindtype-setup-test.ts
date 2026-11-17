import assert from 'node:assert/strict';
import {availableComponents, clone, initialR200sSetup, setupRows, validateSetup, newCondition} from './blindtype-setup-model';
import {dataverseStore, demoSetupStore} from './blindtype-setup-store';

async function main(): Promise<void> {
    const {catalogue, document} = initialR200sSetup();
    assert.equal(catalogue.groups.length, 15); assert.equal(catalogue.groups.flatMap(group => group.components).length, 82);
    assert.equal(setupRows(catalogue, document, 'options').length, 97);
    assert.equal(setupRows(catalogue, document, 'links').length, 104);
    const mechanism = availableComponents(catalogue, document, '147');
    assert.equal(mechanism.length, 12); assert.ok(mechanism.some(item => item.code === 'R3050-SL1-WH'));
    assert.ok(!mechanism.some(item => item.code === 'H-WHITE'), 'Hardware colour cannot be selected as a mechanism.');
    assert.equal(catalogue.legacyResults!.length, 2);
    const invalid = (change: (draft: typeof document) => void, pattern: RegExp): void => {const draft = clone(document); change(draft); assert.throws(() => validateSetup(catalogue, draft), pattern);};
    invalid(draft => {draft.blindTypeCode = 'R200R';}, /different blind type/);
    invalid(draft => {draft.links[0].conditions[0].results[0].componentId = '52752';}, /not available/);
    invalid(draft => {draft.options.find(item => item.groupId === '147')!.defaultId = '52752';}, /default component/);
    invalid(draft => {draft.options.find(item => item.groupId === '305')!.componentIds = [];}, /different enabled option/);
    invalid(draft => {draft.links[0].conditions[0].maxWidth = -1;}, /Size limits/);
    invalid(draft => {draft.links[0].conditions[0].results.push({...draft.links[0].conditions[0].results[0], id: 'new'});}, /only once/);
    invalid(draft => {draft.links[0].conditions[0].results.push({id: 'new-legacy', componentId: '51176', order: 2});}, /not available/);
    invalid(draft => {const source = newCondition('cycle', 1); source.tests = [{groupId: '147', componentId: '52344'}]; draft.links.find(item => item.groupId === '305')!.conditions.push(source);}, /loop/);
    const store = demoSetupStore(); const state = await store.load(); const draft = clone(state.document);
    draft.options.find(item => item.groupId === '147')!.defaultId = '52344';
    const saved = await store.save(state, draft); assert.equal(saved.etag, '2'); assert.equal((await store.load()).document.options.find(item => item.groupId === '147')!.defaultId, '52344');
    await assert.rejects(() => store.save(state, draft), /changed/);
    const recordId = '11111111-1111-4111-8111-111111111111';
    const payload = {'@odata.etag': 'W/"1"', cm_itemid: recordId, cm_name: 'R200S setup demo', cm_blindtypecataloguejson: JSON.stringify(catalogue), cm_blindtypesetupjson: JSON.stringify(document)};
    const calls: {url: string; init: RequestInit | undefined}[] = [];
    const api = dataverseStore('https://capricornsandbox.crm4.dynamics.com', recordId, (async (url, init) => {
        calls.push({url: String(url), init});
        if (init?.method === 'PATCH') {
            const body = JSON.parse(String(init.body)); assert.deepEqual(Object.keys(body), ['cm_blindtypesetupjson']);
            assert.equal((init.headers as Record<string, string>)['If-Match'], 'W/"1"');
            return new Response(JSON.stringify({...payload, '@odata.etag': 'W/"2"', cm_blindtypesetupjson: body.cm_blindtypesetupjson}), {status: 200});
        }
        return new Response(JSON.stringify(payload), {status: 200});
    }) as typeof fetch);
    const live = await api.load(); await api.save(live, draft);
    assert.equal(calls.filter(call => call.init?.method === 'PATCH').length, 1);
    const badDraft = clone(draft); badDraft.options[0].componentIds.push('foreign-component');
    await assert.rejects(() => api.save(live, badDraft), /available/);
    assert.equal(calls.filter(call => call.init?.method === 'PATCH').length, 1, 'Invalid components never reach a write request.');
    console.log('Blind-type setup passed: scoped component pools, defaults, invalid references, sizes, duplicate results, cycles, preserved legacy results and conditional Dataverse saves.');
}
void main().catch(error => {console.error(error); process.exitCode = 1;});
