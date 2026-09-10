export type DateGroupingInterval = 'day' | 'week' | 'month' | 'quarter' | 'year';
export const dateGroupingIntervals: readonly DateGroupingInterval[] = ['day', 'week', 'month', 'quarter', 'year'];
export interface DateGroupOptions {
    dateOnly?: boolean;
    firstDayOfWeek?: number;
    locale?: string;
    /** Same sign as Date.getTimezoneOffset(): UTC minus local time, in minutes. */
    offsetMinutes?: (date: Date) => number;
}

/** Calendar grouping never groups by a formatted label or converts DateOnly through local time. */
export function dateGroup(value: unknown, interval: DateGroupingInterval, options: DateGroupOptions = {}): {key: string; label: string; sortKey: string} {
    if (value === null || value === undefined || value === '') return {key: 'date:blank', label: '(Blanks)', sortKey: 'z:blank'};
    const date = value instanceof Date ? new Date(value.getTime()) : typeof value === 'string' ? new Date(value) : new Date(NaN);
    if (!Number.isFinite(date.getTime())) return {key: 'date:invalid', label: 'Invalid date', sortKey: 'z:invalid'};
    const isCalendarValue = options.dateOnly === true || typeof value === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(value);
    const offset = isCalendarValue ? 0 : options.offsetMinutes?.(date) ?? date.getTimezoneOffset();
    if (!Number.isFinite(offset)) return {key: 'date:invalid', label: 'Invalid date', sortKey: 'z:invalid'};
    const local = new Date(date.getTime() - offset * 60000);
    const start = new Date(0);
    start.setUTCFullYear(local.getUTCFullYear(), local.getUTCMonth(), local.getUTCDate());
    start.setUTCHours(0, 0, 0, 0);
    if (interval === 'week') {
        const first = Number.isInteger(options.firstDayOfWeek) && options.firstDayOfWeek! >= 0 && options.firstDayOfWeek! <= 6 ? options.firstDayOfWeek! : 1;
        start.setUTCDate(start.getUTCDate() - (start.getUTCDay() - first + 7) % 7);
    } else if (interval === 'month') start.setUTCDate(1);
    else if (interval === 'quarter') start.setUTCMonth(Math.floor(start.getUTCMonth() / 3) * 3, 1);
    else if (interval === 'year') start.setUTCMonth(0, 1);
    const stamp = start.toISOString().slice(0, 10);
    const format = (date: Date, settings: Intl.DateTimeFormatOptions): string => {
        try { return date.toLocaleDateString(options.locale, {...settings, timeZone: 'UTC'}); }
        catch { return date.toLocaleDateString('en-GB', {...settings, timeZone: 'UTC'}); }
    };
    let label = format(start, {year: 'numeric', month: 'short', day: 'numeric'});
    if (interval === 'week') {
        const end = new Date(start); end.setUTCDate(end.getUTCDate() + 6);
        label += ` – ${format(end, {year: 'numeric', month: 'short', day: 'numeric'})}`;
    } else if (interval === 'month') label = format(start, {year: 'numeric', month: 'long'});
    else if (interval === 'quarter') label = `Q${Math.floor(start.getUTCMonth() / 3) + 1} ${start.getUTCFullYear()}`;
    else if (interval === 'year') label = String(start.getUTCFullYear());
    return {key: `date:${interval}:${stamp}`, label, sortKey: stamp};
}
