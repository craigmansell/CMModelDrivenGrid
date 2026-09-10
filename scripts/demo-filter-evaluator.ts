/**
 * Small, dependency-free filter and sort evaluator used by the interactive
 * harness.  It intentionally mirrors the shape of a Dataverse dataset filter
 * without importing the PCF typings, so the local demo can also be tested in
 * isolation.
 */

export interface DemoCondition {
	attributeName: string;
	conditionOperator: number;
	value?: unknown;
}

export interface DemoFilterExpression {
	filterOperator?: number;
	conditions?: DemoCondition[];
	filters?: DemoFilterExpression[];
}

export type DemoColumnKind = "text" | "number" | "dateOnly" | "boolean" | "option" | "lookup";
export type DemoClock = Date | (() => Date);
export type DemoSortDirection = "asc" | "desc";
export interface DemoSortLevel {
	columnName: string;
	direction: DemoSortDirection;
}

export interface DemoEvaluatorOptions<Row> {
	getValue: (row: Row, columnName: string) => unknown;
	getDisplayValue?: (row: Row, columnName: string) => string | null | undefined;
	columnKinds?: Record<string, DemoColumnKind | undefined>;
	/** Injectable local clock. Defaults to the current machine time. */
	now?: DemoClock;
}

const DAY_MS = 24 * 60 * 60 * 1000;

interface DateParts {
	year: number;
	month: number;
	day: number;
}

function isFiniteNumber(value: unknown): value is number {
	return typeof value === "number" && Number.isFinite(value);
}

function isNullish(value: unknown): value is null | undefined {
	return value === null || value === undefined;
}

function isBlank(value: unknown): boolean {
	return isNullish(value) || value === "";
}

function getClockDate(now?: DemoClock): Date {
	const value = typeof now === "function" ? now() : now;
	if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
	return new Date();
}

function datePartsToOrdinal(parts: DateParts): number {
	return Date.UTC(parts.year, parts.month - 1, parts.day) / DAY_MS;
}

function ordinalToDateParts(ordinal: number): DateParts {
	const date = new Date(ordinal * DAY_MS);
	return { year: date.getUTCFullYear(), month: date.getUTCMonth() + 1, day: date.getUTCDate() };
}

function addDays(ordinal: number, days: number): number {
	return ordinal + days;
}

function addMonths(ordinal: number, months: number): number {
	const source = ordinalToDateParts(ordinal);
	// Date.UTC handles month overflow, while using day one avoids end-of-month
	// clamping changing the meaning of a relative month filter.
	const firstOfTargetMonth = new Date(Date.UTC(source.year, source.month - 1 + months, 1));
	return Date.UTC(firstOfTargetMonth.getUTCFullYear(), firstOfTargetMonth.getUTCMonth(), 1) / DAY_MS;
}

function addCalendarMonths(ordinal: number, months: number): number {
	const source = ordinalToDateParts(ordinal);
	const targetMonth = new Date(Date.UTC(source.year, source.month - 1 + months, 1));
	const targetYear = targetMonth.getUTCFullYear();
	const targetMonthNumber = targetMonth.getUTCMonth() + 1;
	const targetDay = Math.min(source.day, daysInMonth(targetYear, targetMonthNumber));
	return datePartsToOrdinal({ year: targetYear, month: targetMonthNumber, day: targetDay });
}

function daysInMonth(year: number, month: number): number {
	return new Date(Date.UTC(year, month, 0)).getUTCDate();
}

function parseDateParts(yearText: string, monthText: string, dayText: string): DateParts | null {
	const year = Number(yearText);
	const month = Number(monthText);
	const day = Number(dayText);
	if (!Number.isInteger(year) || !Number.isInteger(month) || !Number.isInteger(day)) return null;
	if (month < 1 || month > 12 || day < 1 || day > daysInMonth(year, month)) return null;
	return { year, month, day };
}

/**
 * Parse a value as a calendar date. DateOnly strings deliberately use their
 * first ten characters, rather than Date.parse(), so an offset in a timestamp
 * cannot move a DateOnly value across a local-day boundary.
 */
export function getLocalDateOrdinal(value: unknown): number | null {
	if (value instanceof Date) {
		if (Number.isNaN(value.getTime())) return null;
		return datePartsToOrdinal({ year: value.getFullYear(), month: value.getMonth() + 1, day: value.getDate() });
	}

	if (typeof value !== "string") return null;
	const dateText = value.trim();
	const match = /^(\d{4})-(\d{2})-(\d{2})(?:$|[T\s])/.exec(dateText);
	if (match) {
		const parts = parseDateParts(match[1], match[2], match[3]);
		return parts ? datePartsToOrdinal(parts) : null;
	}

	const parsed = new Date(dateText);
	if (Number.isNaN(parsed.getTime())) return null;
	return datePartsToOrdinal({ year: parsed.getFullYear(), month: parsed.getMonth() + 1, day: parsed.getDate() });
}

export function getLocalTodayOrdinal(now?: DemoClock): number {
	const date = getClockDate(now);
	return datePartsToOrdinal({ year: date.getFullYear(), month: date.getMonth() + 1, day: date.getDate() });
}

function getColumnKind<Row>(columnName: string, value: unknown, options: DemoEvaluatorOptions<Row>): DemoColumnKind {
	const configured = options.columnKinds?.[columnName];
	if (configured) return configured;
	if (typeof value === "number") return "number";
	if (value instanceof Date || getLocalDateOrdinal(value) !== null && /^\d{4}-\d{2}-\d{2}/.test(String(value))) {
		return "dateOnly";
	}
	return "text";
}

function parseNumber(value: unknown): number | null {
	if (isFiniteNumber(value)) return value;
	if (typeof value !== "string" || value.trim() === "") return null;
	const parsed = Number(value.trim());
	return Number.isFinite(parsed) ? parsed : null;
}

function parseBoolean(value: unknown): boolean | null {
	if (typeof value === "boolean") return value;
	if (typeof value !== "string") return null;
	switch (value.trim().toLocaleLowerCase()) {
		case "true":
		case "yes":
		case "1":
			return true;
		case "false":
		case "no":
		case "0":
			return false;
		default:
			return null;
	}
}

function compareNonNull<Row>(
	actual: unknown,
	expected: unknown,
	row: Row,
	columnName: string,
	options: DemoEvaluatorOptions<Row>
): number | null {
	if (isNullish(actual)) return null;
	const kind = getColumnKind(columnName, actual, options);

	if (kind === "number") {
		const left = parseNumber(actual);
		const right = parseNumber(expected);
		return left === null || right === null ? null : left - right;
	}
	if (kind === "dateOnly") {
		const left = getLocalDateOrdinal(actual);
		const right = getLocalDateOrdinal(expected);
		return left === null || right === null ? null : left - right;
	}
	if (kind === "boolean") {
		const left = parseBoolean(actual);
		const right = parseBoolean(expected);
		return left === null || right === null ? null : Number(left) - Number(right);
	}

	return String(actual).toLocaleLowerCase().localeCompare(String(expected ?? "").toLocaleLowerCase(), undefined, {
		numeric: true,
		sensitivity: "base",
	});
}

function isEqual<Row>(actual: unknown, expected: unknown, row: Row, columnName: string, options: DemoEvaluatorOptions<Row>): boolean {
	return compareNonNull(actual, expected, row, columnName, options) === 0;
}

function likePattern(value: string): RegExp {
	let pattern = "";
	for (let index = 0; index < value.length; index++) {
		const character = value[index];
		if (character === '[' && value[index + 2] === ']' && ['%', '_', '['].includes(value[index + 1])) {
			pattern += value[index + 1].replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); index += 2; continue;
		}
		if (character === "%") pattern += ".*";
		else if (character === "_") pattern += ".";
		else pattern += character.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
	}
	return new RegExp(`^${pattern}$`, "i");
}

function matchesLike(actual: unknown, expected: unknown): boolean {
	if (isNullish(actual) || expected === undefined || expected === null) return false;
	return likePattern(String(expected)).test(String(actual));
}

function parseRelativeCount(value: unknown): number | null {
	const parsed = parseNumber(Array.isArray(value) ? value[0] : value);
	return parsed !== null && Number.isInteger(parsed) && parsed >= 0 ? parsed : null;
}

function startOfWeek(ordinal: number): number {
	const date = new Date(ordinal * DAY_MS);
	return ordinal - date.getUTCDay();
}

function matchesDateOperator(operator: number, actual: unknown, value: unknown, now?: DemoClock): boolean {
	const actualOrdinal = getLocalDateOrdinal(actual);
	if (actualOrdinal === null) return false;
	const today = getLocalTodayOrdinal(now);
	if (operator === 25 || operator === 26 || operator === 27) {
		const expectedOrdinal = getLocalDateOrdinal(value);
		if (expectedOrdinal === null) return false;
		if (operator === 25) return actualOrdinal === expectedOrdinal;
		if (operator === 26) return actualOrdinal <= expectedOrdinal;
		return actualOrdinal >= expectedOrdinal;
	}

	if (operator === 14) return actualOrdinal === addDays(today, -1);
	if (operator === 15) return actualOrdinal === today;
	if (operator === 16) return actualOrdinal === addDays(today, 1);
	if (operator === 17) return actualOrdinal >= addDays(today, -6) && actualOrdinal <= today;
	if (operator === 18) return actualOrdinal >= today && actualOrdinal <= addDays(today, 7);

	const thisWeek = startOfWeek(today);
	if (operator === 19) return actualOrdinal >= addDays(thisWeek, -7) && actualOrdinal < thisWeek;
	if (operator === 20) return actualOrdinal >= thisWeek && actualOrdinal < addDays(thisWeek, 7);
	if (operator === 21) return actualOrdinal >= addDays(thisWeek, 7) && actualOrdinal < addDays(thisWeek, 14);

	const todayParts = ordinalToDateParts(today);
	const thisMonth = datePartsToOrdinal({ year: todayParts.year, month: todayParts.month, day: 1 });
	if (operator === 24) return actualOrdinal >= addMonths(thisMonth, 1) && actualOrdinal < addMonths(thisMonth, 2);
	if (operator === 22) {
		const previousMonth = addMonths(thisMonth, -1);
		return actualOrdinal >= previousMonth && actualOrdinal < thisMonth;
	}
	if (operator === 23) {
		const nextMonth = addMonths(thisMonth, 1);
		return actualOrdinal >= thisMonth && actualOrdinal < nextMonth;
	}

	const thisYear = datePartsToOrdinal({ year: todayParts.year, month: 1, day: 1 });
	if (operator === 30) return actualOrdinal >= addMonths(thisYear,12) && actualOrdinal < addMonths(thisYear,24);
	if (operator === 28) {
		const previousYear = datePartsToOrdinal({ year: todayParts.year - 1, month: 1, day: 1 });
		return actualOrdinal >= previousYear && actualOrdinal < thisYear;
	}
	if (operator === 29) {
		const nextYear = datePartsToOrdinal({ year: todayParts.year + 1, month: 1, day: 1 });
		return actualOrdinal >= thisYear && actualOrdinal < nextYear;
	}

	if (operator === 33 || operator === 34) {
		const count = parseRelativeCount(value);
		if (count === null) return false;
		return operator === 33
			? actualOrdinal >= addDays(today, -count) && actualOrdinal <= today
			: actualOrdinal >= today && actualOrdinal <= addDays(today, count);
	}
	if (operator === 37 || operator === 38) {
		const count = parseRelativeCount(value);
		if (count === null) return false;
		const boundary = addCalendarMonths(today, count * (operator === 37 ? -1 : 1));
		return operator === 37
			? actualOrdinal >= boundary && actualOrdinal <= today
			: actualOrdinal >= today && actualOrdinal <= boundary;
	}
	// The harness uses a labelled January–December, monthly fiscal calendar.
	// Live fiscal conditions are sent unchanged to Dataverse's organisation calendar.
	const count = parseRelativeCount(value);
	const within = (from: number, to: number) => actualOrdinal >= from && actualOrdinal < to;
	if (operator === 58) return within(thisYear, addMonths(thisYear,12));
	if (operator === 59) return within(thisMonth, addMonths(thisMonth,1));
	if (operator === 60) return within(addMonths(thisYear,12), addMonths(thisYear,24));
	if (operator === 61) return within(addMonths(thisMonth,1), addMonths(thisMonth,2));
	if (operator === 62) return within(addMonths(thisYear,-12), thisYear);
	if (operator === 63) return within(addMonths(thisMonth,-1), thisMonth);
	if (operator === 64 && count !== null) return within(addMonths(thisYear,-12*count),thisYear);
	if (operator === 65 && count !== null) return within(addMonths(thisMonth,-count),thisMonth);
	if (operator === 66 && count !== null) return within(addMonths(thisYear,12),addMonths(thisYear,12*(count+1)));
	if (operator === 67 && count !== null) return within(addMonths(thisMonth,1),addMonths(thisMonth,count+1));
	const actualParts = ordinalToDateParts(actualOrdinal);
	if (operator === 68) return actualParts.year === count;
	if (operator === 69) return actualParts.month === count;
	if (operator >= 70 && operator <= 72 && Array.isArray(value)) {
		const period = Number(value[0]), year = Number(value[1]);
		if (!Number.isInteger(period) || period < 1 || period > 12 || !Number.isInteger(year)) return false;
		const boundary = datePartsToOrdinal({year,month:period,day:1});
		return operator === 70 ? within(boundary,addMonths(boundary,1)) : operator === 71 ? actualOrdinal < addMonths(boundary,1) : actualOrdinal >= boundary;
	}
	if (count === null) return false;
	if (operator === 35) return actualOrdinal >= today - count*7 && actualOrdinal <= today;
	if (operator === 36) return actualOrdinal >= today && actualOrdinal <= today + count*7;
	if (operator === 39) return actualOrdinal >= addCalendarMonths(today,-12*count) && actualOrdinal <= today;
	if (operator === 40) return actualOrdinal >= today && actualOrdinal <= addCalendarMonths(today,12*count);
	if (operator === 53) return actualOrdinal < addCalendarMonths(today,-count);
	if (operator === 82) return actualOrdinal < addCalendarMonths(today,-count*12);
	if (operator === 83) return actualOrdinal < today - count*7;
	if (operator === 84) return actualOrdinal < today - count;
	const timestamp = actual instanceof Date ? actual.getTime() : typeof actual === 'string' ? new Date(actual).getTime() : NaN;
	const nowMs = getClockDate(now).getTime();
	if (operator === 31) return timestamp >= nowMs - count*3600000 && timestamp <= nowMs;
	if (operator === 32) return timestamp >= nowMs && timestamp <= nowMs + count*3600000;
	if (operator === 85) return timestamp < nowMs - count*3600000;
	if (operator === 86) return timestamp < nowMs - count*60000;
	return false;
}

function matchesCondition<Row>(row: Row, condition: DemoCondition, options: DemoEvaluatorOptions<Row>): boolean {
	const actual = options.getValue(row, condition.attributeName);
	const operator = condition.conditionOperator;

	if (operator === 12) return isBlank(actual);
	if (operator === 13) return !isBlank(actual);
	// Dataverse comparisons do not treat null as an ordinary value. In
	// particular, null must not satisfy not-equal, in, or like.
	if (isNullish(actual)) return false;

	if (operator >= 14) {
		return matchesDateOperator(operator, actual, condition.value, options.now);
	}
	if (operator === 6) return matchesLike(actual, condition.value);
	if (operator === 7) return !isBlank(actual) && !matchesLike(actual, condition.value);
	if (operator === 8) {
		const values = Array.isArray(condition.value) ? condition.value : [condition.value];
		return values.some((value) => value !== undefined && isEqual(actual, value, row, condition.attributeName, options));
	}

	const comparison = compareNonNull(actual, condition.value, row, condition.attributeName, options);
	if (comparison === null) return false;
	switch (operator) {
		case 0: return comparison === 0;
		case 1: return comparison !== 0;
		case 2: return comparison > 0;
		case 3: return comparison < 0;
		case 4: return comparison >= 0;
		case 5: return comparison <= 0;
		default: return false;
	}
}

/** Evaluate all conditions before the harness applies paging. */
export function evaluateDemoFilter<Row>(
	row: Row,
	expression: DemoFilterExpression | undefined,
	options: DemoEvaluatorOptions<Row>
): boolean {
	if (!expression) return true;
	const results = [
		...(expression.conditions ?? []).map((condition) => matchesCondition(row, condition, options)),
		...(expression.filters ?? []).map((child) => evaluateDemoFilter(row, child, options)),
	];
	return expression.filterOperator === 1 ? results.some(Boolean) : results.every(Boolean);
}

function sortComparable<Row>(
	row: Row,
	columnName: string,
	options: DemoEvaluatorOptions<Row>
): { blank: boolean; value: number | string } {
	const rawValue = options.getValue(row, columnName);
	if (isBlank(rawValue)) return { blank: true, value: "" };
	const kind = getColumnKind(columnName, rawValue, options);
	if (kind === "number") {
		const numericValue = parseNumber(rawValue);
		return numericValue === null ? { blank: true, value: "" } : { blank: false, value: numericValue };
	}
	if (kind === "dateOnly") {
		const dateValue = getLocalDateOrdinal(rawValue);
		return dateValue === null ? { blank: true, value: "" } : { blank: false, value: dateValue };
	}
	const displayValue = options.getDisplayValue?.(row, columnName) ?? rawValue;
	return isBlank(displayValue)
		? { blank: true, value: "" }
		: { blank: false, value: String(displayValue) };
}

/** Sort the complete filtered result, preserving original order for ties. */
export function sortDemoRows<Row>(
	rows: Row[],
	sortLevels: DemoSortLevel[],
	options: DemoEvaluatorOptions<Row>
): Row[] {
	return rows
		.map((row, index) => ({
			row,
			index,
			comparables: sortLevels.map((level) => sortComparable(row, level.columnName, options)),
		}))
		.sort((left, right) => {
			for (let levelIndex = 0; levelIndex < sortLevels.length; levelIndex++) {
				const leftComparable = left.comparables[levelIndex];
				const rightComparable = right.comparables[levelIndex];
				if (leftComparable.blank || rightComparable.blank) {
					if (leftComparable.blank && rightComparable.blank) continue;
					// Excel keeps blanks at the end for either ascending or descending.
					return leftComparable.blank ? 1 : -1;
				}
				let comparison: number;
				if (typeof leftComparable.value === "number" && typeof rightComparable.value === "number") {
					comparison = leftComparable.value - rightComparable.value;
				} else {
					comparison = String(leftComparable.value).localeCompare(String(rightComparable.value), undefined, {
						numeric: true,
						sensitivity: "base",
					});
				}
				if (comparison !== 0) return sortLevels[levelIndex].direction === "desc" ? -comparison : comparison;
			}
			return left.index - right.index;
		})
		.map(({ row }) => row);
}
