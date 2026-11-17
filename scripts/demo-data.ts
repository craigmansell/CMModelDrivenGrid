/** Local fixtures only. These records never connect to Dataverse. */
export type DemoRow = { id: string; parentId?: string; name: string; amount: number | null; status: number | null; owner: number | null; start: string | null; active: boolean | null };
export type DemoScenario = 'standard' | 'large' | 'parents-2000-children-20';

export function demoScenario(value: string | null): DemoScenario {
    return value === 'large' || value === 'parents-2000-children-20' ? value : 'standard';
}

function localDateWithOffset(base: Date, offsetDays: number): string {
    const date = new Date(base.getFullYear(), base.getMonth(), base.getDate() + offsetDays);
    return `${String(date.getFullYear()).padStart(4, '0')}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

export function seedDemoRows(now = new Date(), scenario: DemoScenario = 'standard'): DemoRow[] {
    const everyParent = scenario === 'parents-2000-children-20';
    const large = scenario === 'large';
    const parents: DemoRow[] = Array.from({length: everyParent ? 2000 : large ? 10000 : 180}, (_, index) => ({
        id: `sample-${index + 1}`,
        name: `${['Atlas', 'Beacon', 'Cedar', 'Delta', 'Evergreen', 'Falcon'][index % 6]} project ${String(index + 1).padStart(3, '0')}`,
        amount: index % 11 === 0 ? null : [1250, 4500, 9750.5, 18000, 32500][index % 5],
        status: index % 13 === 0 ? null : index % 3,
        owner: index % 9 === 0 ? null : index % 4,
        start: index % 17 === 0 ? null : localDateWithOffset(now, ((index % 180) - 90) * 5),
        active: index % 7 === 0 ? null : index % 2 === 0,
    }));
    const withChildren = everyParent ? parents : [...parents.slice(0, 6), ...(large ? [parents[50]] : [])];
    const children = withChildren.flatMap((parent, index) => Array.from({length: everyParent ? 20 : index === 0 ? (large ? 2000 : 85) : 3}, (_, child) => ({
        ...parent, id: `${parent.id}-child-${child + 1}`, parentId: parent.id,
        name: `Option ${String(child + 1).padStart(3, '0')}`, amount: (child + 1) * 25, status: child % 3,
    })));
    // The requested wide hierarchy is exactly two levels: no extra grandchild.
    return everyParent ? [...parents, ...children] : [...parents, ...children, {...children[0], id: 'sample-grandchild', parentId: children[0].id, name: 'Nested option', amount: 10}];
}

export function indexDemoRows(rows: DemoRow[]): {parents: DemoRow[]; byId: Map<string, DemoRow>; children: Map<string, DemoRow[]>} {
    const parents: DemoRow[] = [];
    const byId = new Map<string, DemoRow>();
    const children = new Map<string, DemoRow[]>();
    for (const row of rows) {
        if (byId.has(row.id)) throw new Error(`Duplicate sample ID: ${row.id}`);
        byId.set(row.id, row);
        if (row.parentId) {
            const branch = children.get(row.parentId) ?? [];
            branch.push(row);
            children.set(row.parentId, branch);
        } else parents.push(row);
    }
    return {parents, byId, children};
}

export function pageDemoRows<T>(rows: T[], page: number, pageSize = 25): {rows: T[]; hasNextPage: boolean} {
    if (!Number.isInteger(page) || page < 1 || !Number.isInteger(pageSize) || pageSize < 1) throw new Error('Invalid sample page.');
    return {rows: rows.slice((page - 1) * pageSize, page * pageSize), hasNextPage: page * pageSize < rows.length};
}
