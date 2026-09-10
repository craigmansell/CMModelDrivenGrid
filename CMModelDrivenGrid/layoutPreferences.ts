export interface ColumnLayout {order: string[]; widths: Record<string, number>}
export interface LayoutStorage {getItem(key: string): string | null; setItem(key: string, value: string): void; removeItem(key: string): void}
const prefix = 'cm.drivengrid.layout.v1:';
const safeName = (value: unknown): value is string => typeof value === 'string' && value.length > 0 && value.length <= 200 && !['__proto__', 'prototype', 'constructor'].includes(value);

function storage(): LayoutStorage | undefined {
    try { return typeof window === 'undefined' ? undefined : window.localStorage; } catch { return undefined; }
}
export function readColumnLayout(key: string | undefined, source: LayoutStorage | undefined = storage()): ColumnLayout | undefined {
    if (!key || !source) return undefined;
    try {
        const text = source.getItem(prefix + key);
        if (!text || text.length > 50000) return undefined;
        const candidate = JSON.parse(text) as {version?: unknown; order?: unknown; widths?: unknown};
        if (candidate.version !== 1 || !Array.isArray(candidate.order) || !candidate.widths || typeof candidate.widths !== 'object') return undefined;
        const order = [...new Set(candidate.order.filter(safeName))].slice(0, 200);
        const widths: Record<string, number> = {};
        for (const [name, value] of Object.entries(candidate.widths).slice(0, 200)) {
            if (safeName(name) && typeof value === 'number' && Number.isFinite(value) && value >= 80 && value <= 2000) widths[name] = value;
        }
        return {order, widths};
    } catch { return undefined; }
}
/** Store only presentation: never rows, lookup values, filters, drafts or selections. */
export function writeColumnLayout(key: string | undefined, layout: ColumnLayout, target: LayoutStorage | undefined = storage()): boolean {
    if (!key) return true;
    if (!target) return false;
    try { target.setItem(prefix + key, JSON.stringify({version: 1, order: layout.order, widths: layout.widths})); return true; }
    catch { return false; }
}

export function allowLookupNavigation(enabled: boolean, mode: 'all' | 'primary' | 'none' | undefined, isPrimary: boolean | undefined): boolean {
    return enabled && mode !== 'none' && (mode !== 'primary' || isPrimary === true);
}
