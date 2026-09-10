import assert from 'node:assert/strict';
import {demoScenario, seedDemoRows, indexDemoRows, pageDemoRows} from './demo-data';
import {evaluateDemoFilter, sortDemoRows} from './demo-filter-evaluator';

const started = performance.now();
const rows = seedDemoRows(new Date(2026, 8, 10), 'parents-2000-children-20');
const index = indexDemoRows(rows);
assert.equal(rows.length, 42000);
assert.equal(index.parents.length, 2000);
assert.equal(index.byId.size, 42000, 'Every parent and child must have a globally unique ID.');
assert.equal(index.children.size, 2000, 'Every parent, and no child, has a branch.');

const visited = new Set<string>();
let batches = 0;
for (let page = 1; ; page++) {
    const batch = pageDemoRows(index.parents, page);
    batches++;
    assert.equal(batch.rows.length, 25);
    for (const parent of batch.rows) {
        assert.ok(!visited.has(parent.id)); visited.add(parent.id);
        const children = index.children.get(parent.id)!;
        assert.equal(children.length, 20, `${parent.id} must have exactly 20 children.`);
        const childPage = pageDemoRows(children, 1);
        assert.equal(childPage.rows.length, 20);
        assert.equal(childPage.hasNextPage, false, 'Twenty children fit in one 25-row child page.');
        assert.equal(pageDemoRows(children, 2).rows.length, 0);
        for (const [offset, child] of childPage.rows.entries()) {
            assert.equal(child.parentId, parent.id);
            assert.equal(child.id, `${parent.id}-child-${offset + 1}`);
            assert.ok(!index.children.has(child.id), 'No unexpected grandchildren.');
            assert.ok(!visited.has(child.id)); visited.add(child.id);
        }
    }
    if (!batch.hasNextPage) break;
}
assert.equal(batches, 80);
assert.equal(visited.size, 42000, 'The complete paged hierarchy has no missing or duplicated records.');
assert.equal(index.parents[1999].name, 'Beacon project 2000');
assert.equal(index.children.get('sample-2000')![19].id, 'sample-2000-child-20');
assert.equal(pageDemoRows(index.parents, 81).hasNextPage, false);
assert.equal(pageDemoRows(index.parents, 81).rows.length, 0);
assert.throws(() => pageDemoRows(index.parents, 0));
assert.throws(() => pageDemoRows(index.parents, 1, 0));
assert.throws(() => indexDemoRows([rows[0], rows[0]]), /Duplicate sample ID/);

const filterOptions = {getValue: (row: typeof rows[number], column: string) => row[column as keyof typeof row], columnKinds: {name: 'text' as const, amount: 'number' as const}};
const lastParent = index.parents.filter(row => evaluateDemoFilter(row, {filterOperator: 0, conditions: [{attributeName: 'name', conditionOperator: 0, value: 'Beacon project 2000'}]}, filterOptions));
assert.deepEqual(lastParent.map(row => row.id), ['sample-2000'], 'Filtering must inspect parents beyond the first loaded page.');
const branch = index.children.get(lastParent[0].id)!;
const filteredChildren = branch.filter(row => evaluateDemoFilter(row, {filterOperator: 0, conditions: [{attributeName: 'amount', conditionOperator: 4, value: 250}]}, filterOptions));
assert.equal(filteredChildren.length, 11);
const descending = sortDemoRows(filteredChildren, [{columnName: 'amount', direction: 'desc'}], filterOptions);
assert.equal(descending[0].id, 'sample-2000-child-20');
assert.equal(descending[10].id, 'sample-2000-child-10');
assert.equal(branch[0].id, 'sample-2000-child-1', 'Sorting must not mutate the stored branch.');
assert.equal(index.children.get('sample-1')!.length, 20, 'Sibling branches remain separate.');

for (const [scenario, parentCount, childCount] of [['standard', 180, 101], ['large', 10000, 2019]] as const) {
    const old = seedDemoRows(new Date(2026, 8, 10), scenario);
    const oldIndex = indexDemoRows(old);
    assert.equal(oldIndex.parents.length, parentCount);
    assert.equal(old.length - parentCount, childCount);
    assert.equal(oldIndex.children.get('sample-1-child-1')!.length, 1);
}
assert.equal(demoScenario('parents-2000-children-20'), 'parents-2000-children-20');
assert.equal(demoScenario('large'), 'large');
assert.equal(demoScenario('untrusted-scenario'), 'standard');
assert.equal(demoScenario(null), 'standard');
console.log(`42,000-record hierarchy passed: 2,000 parents, exactly 20 children each, 80 parent batches, all 2,000 child branches, scope/filter/sort/end-of-data and existing scenarios (${Math.round(performance.now() - started)} ms).`);
