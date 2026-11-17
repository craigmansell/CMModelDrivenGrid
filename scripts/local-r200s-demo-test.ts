import assert from 'node:assert/strict';
import {createR200sDemoRows, r200sRecord, r200sSnapshot} from './r200s-demo-data';
import {pageDemoRows} from './demo-data';

const rows = createR200sDemoRows();
const byId = new Map(rows.map(row => [row.id, row]));
const children = (id: string) => rows.filter(row => row.parentId === id);
assert.equal(r200sSnapshot.blindType.id, 461);
assert.equal(r200sSnapshot.blindType.code, 'R200S');
assert.equal(rows.length, 104, 'Four linked options, 32 conditions and 68 stored result rows.');
assert.equal(byId.size, rows.length, 'Repeated choice codes must retain distinct result-link IDs.');
assert.deepEqual(rows.filter(row => !row.parentId).map(row => row.name), ['Barrel Type', 'Mechanism', 'Bracket Cover', 'Chain/Cord Tidies']);
assert.equal(rows.filter(row => row.kind === 'Condition').length, 32);
assert.equal(rows.filter(row => row.kind === 'Result').length, 68);
assert.deepEqual(children('r200s-link-103').map(row => row.sequence), [1,2,3,4,5,6,7,8,9]);
assert.deepEqual(children('r200s-condition-314').map(row => row.code), ['R3050-SL1-WH']);
assert.equal(byId.get('r200s-condition-314')!.details, 'Barrel Type AND Hardware Colour');
assert.deepEqual(children('r200s-condition-2154').map(row => row.code), ['NOBKTCOV', 'R3080-40-WH', 'R3042-WH']);
assert.deepEqual(children('r200s-condition-2738').map(row => row.code), ['R8024 (MUST BE BATCH AR)', 'R8021', 'R8027']);
assert.equal(byId.get('r200s-condition-2738')!.name, 'Width 0–2,500 mm');
assert.equal(byId.get('r200s-condition-2739')!.name, 'Width 2,501–3,000 mm');
for (const id of [2667, 2673]) {
    const condition = byId.get(`r200s-condition-${id}`)!;
    assert.ok(condition.details.includes('No results configured'));
    assert.equal(r200sRecord(condition).getFormattedValue('resultCount'), '0');
    assert.equal(children(condition.id).length, 0, 'Do not invent result choices for an empty stored rule.');
}
for (const row of rows) {
    assert.equal(r200sRecord(row).getRecordId(), row.id);
    if (row.parentId) {
        const parent = byId.get(row.parentId)!;
        assert.ok(parent, `Missing parent for ${row.id}`);
        assert.equal(parent.kind, row.kind === 'Condition' ? 'Linked option' : 'Condition');
    }
    if (row.kind === 'Result') assert.equal(children(row.id).length, 0);
}
const covers = children('r200s-link-105');
assert.equal(covers.length, 19);
assert.equal(pageDemoRows(covers, 1).hasNextPage, false);
assert.equal(pageDemoRows(covers, 2).rows.length, 0);
console.log('R200S snapshot passed: 4 linked options, 32 conditions, 68 result links, exact example chains and both empty rules preserved.');
