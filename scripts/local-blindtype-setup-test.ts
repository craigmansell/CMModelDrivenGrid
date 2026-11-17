import assert from 'node:assert/strict';
import {availableComponents, clone, initialR200sSetup, setupRows, validateSetup, newCondition, OptionSetup, optionGroup, withOption, withoutOption} from './blindtype-setup-model';
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
    const custom: OptionSetup = {groupId: 'custom-test-option', name: ' Extra mechanism ', sourceGroupId: '147', order: 100, enabled: true, hidden: false, freeText: false, componentIds: ['52344', '52343'], defaultId: '52344'};
    const added = withOption(catalogue, document, custom);
    assert.equal(document.options.length, 15, 'Adding an option does not mutate the saved document.');
    assert.equal(added.version, 2); assert.equal(added.options.length, 16);
    assert.equal(optionGroup(catalogue, added, custom.groupId).name, 'Extra mechanism');
    assert.deepEqual(availableComponents(catalogue, added, custom.groupId).map(item => item.id).sort(), ['52343', '52344']);
    assert.equal(setupRows(catalogue, added, 'options').find(row => row.id === `option-${custom.groupId}`)?.name, 'Extra mechanism');
    assert.ok(setupRows(catalogue, added, 'options').some(row => row.groupId === custom.groupId && row.componentId === '52344' && row.details.includes('Default')));
    for (const [values, pattern] of [
        [{name: ''}, /option name/], [{name: 'x'.repeat(101)}, /option name/], [{name: '  MECHANISM  '}, /already exists/],
        [{name: 'Ｍｅｃｈａｎｉｓｍ'}, /already exists/], [{sourceGroupId: 'foreign'}, /not available/],
        [{componentIds: []}, /at least one/], [{componentIds: ['52752']}, /select only components/],
        [{defaultId: '52345'}, /default component/],
    ] as [Partial<OptionSetup>, RegExp][]) assert.throws(() => withOption(catalogue, document, {...custom, ...values}), pattern);
    const unavailable = clone(catalogue); unavailable.groups.find(group => group.id === '147')!.components.find(item => item.id === '52344')!.available = false;
    assert.throws(() => withOption(unavailable, document, custom), /select only components/);
    const independent = clone(added); independent.options.find(item => item.groupId === '147')!.componentIds = [];
    assert.equal(availableComponents(catalogue, independent, custom.groupId).length, 2, 'New option has its own allowed subset of the blind-type catalogue.');
    const edited = withOption(catalogue, added, {...added.options[added.options.length - 1], name: 'Renamed mechanism', componentIds: ['52343'], defaultId: '52343'});
    assert.equal(edited.options.length, 16); assert.equal(edited.options[15].defaultId, '52343');
    const linkedCustom = clone(added); const customCondition = newCondition('custom-condition', 1);
    customCondition.tests = [{groupId: '259', componentId: availableComponents(catalogue, added, '259')[0].id}];
    customCondition.results = [{id: 'custom-result', componentId: '52344', order: 1}];
    linkedCustom.links.push({id: 'custom-link', groupId: custom.groupId, order: 99, conditions: [customCondition]});
    validateSetup(catalogue, linkedCustom);
    assert.ok(setupRows(catalogue, linkedCustom, 'links').some(row => row.groupId === custom.groupId && row.kind === 'Result' && row.code === 'R3050-SL1-WH'));
    assert.throws(() => withoutOption(catalogue, linkedCustom, custom.groupId), /used by linked rules/);
    const customSource = clone(added); customSource.links[0].conditions.push({...newCondition('custom-source', 99), tests: [{groupId: custom.groupId, componentId: '52343'}]});
    validateSetup(catalogue, customSource);
    assert.throws(() => withoutOption(catalogue, customSource, custom.groupId), /used by linked rules/);
    const customCycle = clone(linkedCustom); customCycle.links.push({id: 'custom-cycle', groupId: '259', order: 100, conditions: [{...newCondition('custom-cycle-condition', 1), tests: [{groupId: custom.groupId, componentId: '52344'}]}]});
    assert.throws(() => validateSetup(catalogue, customCycle), /loop/);
    assert.deepEqual(withoutOption(catalogue, added, custom.groupId), document, 'Removing the last custom option restores the exact original v1 document.');
    assert.throws(() => withoutOption(catalogue, added, '147'), /Catalogue options/);
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
    const live = await api.load(); const savedCustom = await api.save(live, withOption(catalogue, draft, custom));
    assert.equal(savedCustom.document.version, 2); assert.equal(savedCustom.document.options[15].defaultId, '52344');
    assert.equal(calls.filter(call => call.init?.method === 'PATCH').length, 1);
    const badDraft = clone(draft); badDraft.options[0].componentIds.push('foreign-component');
    await assert.rejects(() => api.save(live, badDraft), /available/);
    assert.equal(calls.filter(call => call.init?.method === 'PATCH').length, 1, 'Invalid components never reach a write request.');
    console.log('Blind-type setup passed: new named options, component scope, defaults, custom linked targets/sources, safe removal, v1/v2 persistence, validation, cycles, legacy results and conditional Dataverse saves.');
}
void main().catch(error => {console.error(error); process.exitCode = 1;});
