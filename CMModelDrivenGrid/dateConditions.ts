/** Dataverse SDK operator IDs + FetchXML names. Fiscal boundaries are evaluated by Dataverse, not guessed locally. */
export const EXTRA_DATE_CONDITIONS = {
    nextWeek: ['Next week', 21, 'next-week', 'none'],
    nextMonth: ['Next month', 24, 'next-month', 'none'],
    nextYear: ['Next year', 30, 'next-year', 'none'],
    nextXHours: ['Next X hours', 32, 'next-x-hours', 'count'],
    lastXHours: ['Last X hours', 31, 'last-x-hours', 'count'],
    nextXWeeks: ['Next X weeks', 36, 'next-x-weeks', 'count'],
    lastXWeeks: ['Last X weeks', 35, 'last-x-weeks', 'count'],
    nextXYears: ['Next X years', 40, 'next-x-years', 'count'],
    lastXYears: ['Last X years', 39, 'last-x-years', 'count'],
    thisFiscalYear: ['This fiscal year', 58, 'this-fiscal-year', 'none'],
    thisFiscalPeriod: ['This fiscal period', 59, 'this-fiscal-period', 'none'],
    nextFiscalYear: ['Next fiscal year', 60, 'next-fiscal-year', 'none'],
    nextFiscalPeriod: ['Next fiscal period', 61, 'next-fiscal-period', 'none'],
    lastFiscalYear: ['Last fiscal year', 62, 'last-fiscal-year', 'none'],
    lastFiscalPeriod: ['Last fiscal period', 63, 'last-fiscal-period', 'none'],
    lastXFiscalYears: ['Last X fiscal years', 64, 'last-x-fiscal-years', 'count'],
    lastXFiscalPeriods: ['Last X fiscal periods', 65, 'last-x-fiscal-periods', 'count'],
    nextXFiscalYears: ['Next X fiscal years', 66, 'next-x-fiscal-years', 'count'],
    nextXFiscalPeriods: ['Next X fiscal periods', 67, 'next-x-fiscal-periods', 'count'],
    inFiscalYear: ['In fiscal year', 68, 'in-fiscal-year', 'year'],
    inFiscalPeriod: ['In fiscal period', 69, 'in-fiscal-period', 'period'],
    inFiscalPeriodAndYear: ['In fiscal period and year', 70, 'in-fiscal-period-and-year', 'periodYear'],
    inOrBeforeFiscalPeriodAndYear: ['In or before fiscal period and year', 71, 'in-or-before-fiscal-period-and-year', 'periodYear'],
    inOrAfterFiscalPeriodAndYear: ['In or after fiscal period and year', 72, 'in-or-after-fiscal-period-and-year', 'periodYear'],
    olderThanXMonths: ['Older than X months', 53, 'olderthan-x-months', 'count'],
    olderThanXYears: ['Older than X years', 82, 'olderthan-x-years', 'count'],
    olderThanXWeeks: ['Older than X weeks', 83, 'olderthan-x-weeks', 'count'],
    olderThanXDays: ['Older than X days', 84, 'olderthan-x-days', 'count'],
    olderThanXHours: ['Older than X hours', 85, 'olderthan-x-hours', 'count'],
    olderThanXMinutes: ['Older than X minutes', 86, 'olderthan-x-minutes', 'count'],
} as const;

export type ExtraDateMode = keyof typeof EXTRA_DATE_CONDITIONS;
export const extraDateEntries = Object.entries(EXTRA_DATE_CONDITIONS) as [ExtraDateMode, typeof EXTRA_DATE_CONDITIONS[ExtraDateMode]][];
export const extraDateOperatorIds = new Set<number>(extraDateEntries.map(([, definition]) => definition[1]));

export function extraDateCondition(mode: string): typeof EXTRA_DATE_CONDITIONS[ExtraDateMode] | undefined {
    return Object.prototype.hasOwnProperty.call(EXTRA_DATE_CONDITIONS, mode) ? EXTRA_DATE_CONDITIONS[mode as ExtraDateMode] : undefined;
}

export function fiscalInteger(value: string, maximum: number): boolean {
    return /^[1-9]\d*$/.test(value) && Number.isSafeInteger(Number(value)) && Number(value) <= maximum;
}
