export interface GridFeatures {
    accessibleLabel?: string;
    enableGrouping?: boolean;
    enableAggregation?: boolean;
    enableColumnMoving?: boolean;
    enableMultipleSelection?: boolean;
    enableRangeSelection?: boolean;
    enableFiltering?: boolean;
    enableSorting?: boolean;
    showTypeIcons?: boolean;
    showStatusIcons?: boolean;
    showChoiceColors?: boolean;
    reflow?: 'grid' | 'list' | 'auto';
    listColumns?: number;
    showAvatar?: boolean;
    showListColumnNames?: boolean;
    bandedListRows?: boolean;
    enableJumpBar?: boolean;
    navigationTypes?: 'all' | 'primary' | 'none';
    firstDayOfWeek?: number;
    locale?: string;
    isRTL?: boolean;
    dateTimeOffsetMinutes?: (date: Date) => number;
}

export type AggregateMode = 'sum' | 'average' | 'minimum' | 'maximum';
export function aggregateNumbers(values: readonly unknown[], mode: AggregateMode): number | null {
    const numbers = values.filter((value): value is number => typeof value === 'number' && Number.isFinite(value));
    if (!numbers.length) return null;
    if (mode === 'minimum') return numbers.reduce((a,b) => Math.min(a,b));
    if (mode === 'maximum') return numbers.reduce((a,b) => Math.max(a,b));
    // Compensated sum avoids accumulated rounding drift in large decimal columns.
    let sum = 0, correction = 0;
    for (const value of numbers) { const adjusted = value - correction; const next = sum + adjusted; correction = (next - sum) - adjusted; sum = next; }
    return mode === 'average' ? sum / numbers.length : sum;
}

export function moveColumn(order: readonly string[], from: number, to: number): string[] {
    if (!Number.isInteger(from) || !Number.isInteger(to) || from < 0 || to < 0 || from >= order.length || to >= order.length) return [...order];
    const next = [...order]; next.splice(to, 0, next.splice(from, 1)[0]); return next;
}

export function reconcileColumnOrder(available: readonly string[], saved: readonly string[]): string[] {
    return [...new Set([...saved.filter(key => available.includes(key)), ...available])];
}

export function numericColumnWidth(value: string): number {
    const width = Number(value);
    if (!Number.isInteger(width) || width < 80 || width > 2000) throw new Error('Enter a width from 80 to 2,000 pixels.');
    return width;
}

/** Preserve blank/value identity without a sentinel that can collide with real text. */
export function groupingKey(value: string | null): string {
    return JSON.stringify(value);
}

/** Group by raw values/IDs: two lookups with the same display name stay distinct. */
export function groupRecords<T>(records: readonly T[], getKey: (row: T) => string, getLabel: (row: T) => string, getSortKey?: (row: T) => string) {
    const buckets = new Map<string, { key: string; name: string; rows: T[]; sortKey?: string }>();
    for (const row of records) {
        const key = getKey(row);
        const bucket = buckets.get(key) ?? { key, name: getLabel(row) || '(Blanks)', rows: [], sortKey: getSortKey?.(row) };
        bucket.rows.push(row); buckets.set(key, bucket);
    }
    const items: T[] = [];
    const groups = [...buckets.values()].sort((a,b) => getSortKey ? (a.sortKey ?? '').localeCompare(b.sortKey ?? '') : a.name.localeCompare(b.name, undefined, {numeric: true, sensitivity: 'base'})).map(bucket => {
        const group = {key: bucket.key, name: bucket.name, startIndex: items.length, count: bucket.rows.length};
        items.push(...bucket.rows); return group;
    });
    return {items, groups};
}

export function columnTypeIcon(dataType: string): string {
    const type = dataType.toLowerCase();
    if (type.includes('lookup') || type.includes('customer') || type.includes('owner')) return 'Link';
    if (type.includes('date')) return 'Calendar';
    if (type.includes('money') || type.includes('currency')) return 'Money';
    if (/whole|decimal|float|double|^fp$/.test(type)) return 'NumberField';
    if (/optionset|twooptions|boolean|status|state/.test(type)) return 'BulletedList';
    return 'TextField';
}

export function choiceColorStyle(color: string | undefined): {backgroundColor: string; color: string} | undefined {
    if (!color || !/^#[0-9a-f]{6}$/i.test(color)) return undefined;
    const rgb = [1,3,5].map(index => parseInt(color.slice(index,index+2),16) / 255).map(value => value <= 0.04045 ? value / 12.92 : ((value + 0.055) / 1.055) ** 2.4);
    const luminance = rgb[0] * 0.2126 + rgb[1] * 0.7152 + rgb[2] * 0.0722;
    return {backgroundColor: color, color: luminance > 0.179 ? '#000000' : '#ffffff'};
}
