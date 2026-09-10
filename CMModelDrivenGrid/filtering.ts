import { ExtraDateMode, extraDateEntries, extraDateCondition, fiscalInteger } from './dateConditions';
type DataSetFilterExpression = ComponentFramework.PropertyHelper.DataSetApi.FilterExpression;
type DataSetConditionExpression = ComponentFramework.PropertyHelper.DataSetApi.ConditionExpression;
type DataSetConditionOperator = ComponentFramework.PropertyHelper.DataSetApi.Types.ConditionOperator;

/**
 * The PCF typings historically omitted FilterExpression.filters even though the
 * model-driven dataset API supports child filter groups. Keep the public
 * helper typed against the platform shape while retaining nested AND/OR groups.
 */
export type FilterExpression = DataSetFilterExpression & {
	filters?: FilterExpression[];
};

/** Modes which compile to a Dataverse condition (or a small condition group). */
export type FilterRuleMode =
	| ExtraDateMode
	| "contains"
	| "doesNotContain"
	| "containsData"
	| "equals"
	| "notEquals"
	| "beginsWith"
	| "doesNotBeginWith"
	| "endsWith"
	| "doesNotEndWith"
	| "doesNotContainData"
	| "greaterThan"
	| "greaterOrEqual"
	| "lessThan"
	| "lessOrEqual"
	| "between"
	| "on"
	| "onOrBefore"
	| "onOrAfter"
	| "yesterday"
	| "today"
	| "tomorrow"
	| "last7Days"
	| "next7Days"
	| "lastWeek"
	| "thisWeek"
	| "lastMonth"
	| "thisMonth"
	| "lastYear"
	| "thisYear"
	| "lastXDays"
	| "nextXDays"
	| "lastXMonths"
	| "nextXMonths";

export type FilterMode = FilterRuleMode | "in" | "clear" | "custom";

export interface ColumnFilterRule {
	mode: FilterRuleMode;
	value?: string | string[];
}

export interface ColumnFilterGroup {
	operator: "and" | "or";
	rules: ColumnFilterRule[];
}

export type FilterValue = string | string[] | ColumnFilterGroup;

export type FilterKind = "text" | "number" | "date" | "choice" | "boolean" | "lookup";
export type FilterInput = "none" | "value" | "range" | "count" | "period" | "year" | "periodYear";

export interface FilterDefinition {
	key: FilterRuleMode;
	text: string;
	input: FilterInput;
	types: FilterKind[];
}

/**
 * Logical Dataverse operators. NOT LIKE and NOT NULL must be compiled by
 * datasetFiltering before passing to a model-driven dataset. Child rows use
 * FetchXML, which supports these operators directly.
 */
const NULL_OPERATOR = 12 as DataSetConditionOperator;
const EQUAL_OPERATOR = 0 as DataSetConditionOperator;
const NOT_EQUAL_OPERATOR = 1 as DataSetConditionOperator;
const GREATER_THAN_OPERATOR = 2 as DataSetConditionOperator;
const LESS_THAN_OPERATOR = 3 as DataSetConditionOperator;
const GREATER_OR_EQUAL_OPERATOR = 4 as DataSetConditionOperator;
const LESS_OR_EQUAL_OPERATOR = 5 as DataSetConditionOperator;
const LIKE_OPERATOR = 6 as DataSetConditionOperator;
const NOT_LIKE_OPERATOR = 7 as DataSetConditionOperator;
const NOT_NULL_OPERATOR = 13 as DataSetConditionOperator;
const IN_OPERATOR = 8 as DataSetConditionOperator;
const YESTERDAY_OPERATOR = 14 as DataSetConditionOperator;
const TODAY_OPERATOR = 15 as DataSetConditionOperator;
const TOMORROW_OPERATOR = 16 as DataSetConditionOperator;
const LAST_7_DAYS_OPERATOR = 17 as DataSetConditionOperator;
const NEXT_7_DAYS_OPERATOR = 18 as DataSetConditionOperator;
const LAST_WEEK_OPERATOR = 19 as DataSetConditionOperator;
const THIS_WEEK_OPERATOR = 20 as DataSetConditionOperator;
const LAST_MONTH_OPERATOR = 22 as DataSetConditionOperator;
const THIS_MONTH_OPERATOR = 23 as DataSetConditionOperator;
const ON_OPERATOR = 25 as DataSetConditionOperator;
const ON_OR_BEFORE_OPERATOR = 26 as DataSetConditionOperator;
const ON_OR_AFTER_OPERATOR = 27 as DataSetConditionOperator;
const LAST_YEAR_OPERATOR = 28 as DataSetConditionOperator;
const THIS_YEAR_OPERATOR = 29 as DataSetConditionOperator;
const LAST_X_DAYS_OPERATOR = 33 as DataSetConditionOperator;
const NEXT_X_DAYS_OPERATOR = 34 as DataSetConditionOperator;
const LAST_X_MONTHS_OPERATOR = 37 as DataSetConditionOperator;
const NEXT_X_MONTHS_OPERATOR = 38 as DataSetConditionOperator;

const ALL_FILTER_KINDS: FilterKind[] = ["text", "number", "date", "choice", "boolean", "lookup"];
const NON_DATE_FILTER_KINDS: FilterKind[] = ["text", "number", "choice", "boolean", "lookup"];
const NUMERIC_AND_DATE_KINDS: FilterKind[] = ["number", "date"];
const NUMERIC_ONLY_KINDS: FilterKind[] = ["number"];

/** Definitions shared by the UI, FetchXML and supported-dataset compiler. */
export const FILTER_DEFINITIONS: FilterDefinition[] = [
	{ key: "equals", text: "Equals", input: "value", types: [...NON_DATE_FILTER_KINDS] },
	{ key: "notEquals", text: "Does not equal", input: "value", types: [...NON_DATE_FILTER_KINDS] },
	{ key: "containsData", text: "Contains data", input: "none", types: [...ALL_FILTER_KINDS] },
	{ key: "doesNotContainData", text: "Does not contain data", input: "none", types: [...ALL_FILTER_KINDS] },
	{ key: "contains", text: "Contains", input: "value", types: ["text"] },
	{ key: "doesNotContain", text: "Does not contain", input: "value", types: ["text"] },
	{ key: "beginsWith", text: "Begins with", input: "value", types: ["text"] },
	{ key: "doesNotBeginWith", text: "Does not begin with", input: "value", types: ["text"] },
	{ key: "endsWith", text: "Ends with", input: "value", types: ["text"] },
	{ key: "doesNotEndWith", text: "Does not end with", input: "value", types: ["text"] },
	{ key: "greaterThan", text: "Greater than", input: "value", types: [...NUMERIC_AND_DATE_KINDS] },
	{ key: "greaterOrEqual", text: "Greater than or equal to", input: "value", types: [...NUMERIC_ONLY_KINDS] },
	{ key: "lessThan", text: "Less than", input: "value", types: [...NUMERIC_AND_DATE_KINDS] },
	{ key: "lessOrEqual", text: "Less than or equal to", input: "value", types: [...NUMERIC_ONLY_KINDS] },
	{ key: "between", text: "Between", input: "range", types: [...NUMERIC_AND_DATE_KINDS] },
	{ key: "on", text: "On", input: "value", types: ["date"] },
	{ key: "onOrBefore", text: "On or before", input: "value", types: ["date"] },
	{ key: "onOrAfter", text: "On or after", input: "value", types: ["date"] },
	{ key: "yesterday", text: "Yesterday", input: "none", types: ["date"] },
	{ key: "today", text: "Today", input: "none", types: ["date"] },
	{ key: "tomorrow", text: "Tomorrow", input: "none", types: ["date"] },
	{ key: "last7Days", text: "Last 7 days", input: "none", types: ["date"] },
	{ key: "next7Days", text: "Next 7 days", input: "none", types: ["date"] },
	{ key: "lastWeek", text: "Last week", input: "none", types: ["date"] },
	{ key: "thisWeek", text: "This week", input: "none", types: ["date"] },
	{ key: "lastMonth", text: "Last month", input: "none", types: ["date"] },
	{ key: "thisMonth", text: "This month", input: "none", types: ["date"] },
	{ key: "lastYear", text: "Last year", input: "none", types: ["date"] },
	{ key: "thisYear", text: "This year", input: "none", types: ["date"] },
	{ key: "lastXDays", text: "Last X days", input: "count", types: ["date"] },
	{ key: "nextXDays", text: "Next X days", input: "count", types: ["date"] },
	{ key: "lastXMonths", text: "Last X months", input: "count", types: ["date"] },
	{ key: "nextXMonths", text: "Next X months", input: "count", types: ["date"] },
	...extraDateEntries.map(([key, [text, , , input]]): FilterDefinition => ({key, text, input, types: ['date']})),
];

const OPERATOR_BY_MODE: Partial<Record<FilterRuleMode, DataSetConditionOperator>> = {
	greaterThan: GREATER_THAN_OPERATOR,
	lessThan: LESS_THAN_OPERATOR,
	greaterOrEqual: GREATER_OR_EQUAL_OPERATOR,
	lessOrEqual: LESS_OR_EQUAL_OPERATOR,
	on: ON_OPERATOR,
	onOrBefore: ON_OR_BEFORE_OPERATOR,
	onOrAfter: ON_OR_AFTER_OPERATOR,
	yesterday: YESTERDAY_OPERATOR,
	today: TODAY_OPERATOR,
	tomorrow: TOMORROW_OPERATOR,
	last7Days: LAST_7_DAYS_OPERATOR,
	next7Days: NEXT_7_DAYS_OPERATOR,
	lastWeek: LAST_WEEK_OPERATOR,
	thisWeek: THIS_WEEK_OPERATOR,
	lastMonth: LAST_MONTH_OPERATOR,
	thisMonth: THIS_MONTH_OPERATOR,
	lastYear: LAST_YEAR_OPERATOR,
	thisYear: THIS_YEAR_OPERATOR,
	lastXDays: LAST_X_DAYS_OPERATOR,
	nextXDays: NEXT_X_DAYS_OPERATOR,
	lastXMonths: LAST_X_MONTHS_OPERATOR,
	nextXMonths: NEXT_X_MONTHS_OPERATOR,
};

const VALUE_MODES = new Set<FilterRuleMode>([
	"contains",
	"doesNotContain",
	"doesNotBeginWith",
	"doesNotEndWith",
	"equals",
	"notEquals",
	"beginsWith",
	"endsWith",
	"greaterThan",
	"greaterOrEqual",
	"lessThan",
	"lessOrEqual",
	"on",
	"onOrBefore",
	"onOrAfter",
]);

const COUNT_MODES = new Set<FilterRuleMode>(["lastXDays", "nextXDays", "lastXMonths", "nextXMonths"]);
const NO_VALUE_MODES = new Set<FilterRuleMode>([
	"containsData",
	"doesNotContainData",
	"yesterday",
	"today",
	"tomorrow",
	"last7Days",
	"next7Days",
	"lastWeek",
	"thisWeek",
	"lastMonth",
	"thisMonth",
	"lastYear",
	"thisYear",
]);

function isFilterRuleMode(value: unknown): value is FilterRuleMode {
	return typeof value === "string" && FILTER_DEFINITIONS.some((definition) => definition.key === value);
}

/** Classify PCF/Dataverse data type strings without throwing on missing metadata. */
export function getFilterKind(dataType: string): FilterKind {
	const normalized = typeof dataType === "string" ? dataType.toLowerCase().replace(/[\s_.-]+/g, "") : "";
	if (!normalized) return "text";
	if (/(lookup|owner|customer|partylist|entityreference)/.test(normalized)) return "lookup";
	if (/(twooptions|boolean|bool|yesno)/.test(normalized)) return "boolean";
	if (/(optionset|choice|picklist|status|enum|multiselect)/.test(normalized)) return "choice";
	if (/(date|datetime|dateandtime|time)/.test(normalized)) return "date";
	if (/(decimal|currency|money|whole|integer|int|double|float|number|numeric)/.test(normalized)) return "number";
	return "text";
}

/** Return only the supported filter definitions for a column's data type. */
export function getFilterDefinitions(dataType: string): FilterDefinition[] {
	const kind = getFilterKind(dataType);
	return FILTER_DEFINITIONS
		.filter((definition) => definition.types.includes(kind))
		.filter(definition => !dataType.toLowerCase().includes('dateonly') || !/hours|minutes/i.test(definition.key))
		.map((definition) => {
			if (kind !== "date") return definition;
			if (definition.key === "greaterThan") return { ...definition, text: "After" };
			if (definition.key === "lessThan") return { ...definition, text: "Before" };
			return definition;
		});
}

function asFilter(filter: DataSetFilterExpression | undefined): FilterExpression | undefined {
	return filter as FilterExpression | undefined;
}

function cloneFilterWithoutColumn(
	filter: FilterExpression,
	columnName: string
): FilterExpression | undefined {
	const conditions = (filter.conditions ?? []).filter((condition) => condition.attributeName !== columnName);
	const filters = (filter.filters ?? [])
		.map((child) => cloneFilterWithoutColumn(child, columnName))
		.filter((child): child is FilterExpression => child !== undefined);

	if (conditions.length === 0 && filters.length === 0) {
		return undefined;
	}

	const next: FilterExpression = {
		...filter,
		conditions,
	};
	if (filter.filters !== undefined || filters.length > 0) {
		next.filters = filters;
	}
	return next;
}

function appendFilter(filter: FilterExpression | undefined, addition: FilterExpression): FilterExpression {
	if (!filter) {
		return addition;
	}

	// An OR root must stay an OR group. Wrap it before adding a column filter so
	// the new column condition is ANDed with the existing view expression.
	if (filter.filterOperator === 1) {
		return {
			filterOperator: 0,
			conditions: [],
			filters: [filter, addition],
		};
	}

	const isSimpleAndAddition = addition.filterOperator === 0 && (addition.filters ?? []).length === 0;
	const next: FilterExpression = {
		...filter,
		conditions: isSimpleAndAddition
			? [...(filter.conditions ?? []), ...(addition.conditions ?? [])]
			: [...(filter.conditions ?? [])],
	};
	if (isSimpleAndAddition) {
		if (filter.filters !== undefined) next.filters = [...filter.filters];
	} else {
		next.filters = [...(filter.filters ?? []), addition];
	}
	return next;
}

function getSingleStringValue(
	mode: FilterRuleMode,
	value: string | string[] | undefined,
	options: { strict: boolean; required: boolean }
): string | undefined {
	let candidate: string | undefined;
	if (Array.isArray(value)) {
		if (value.length !== 1 || typeof value[0] !== "string") {
			if (options.strict || options.required) {
				throw new Error(`Filter mode "${mode}" requires one string value.`);
			}
			return value[0];
		}
		candidate = value[0];
	} else if (value !== undefined) {
		if (typeof value !== "string") {
			if (options.strict || options.required) {
				throw new Error(`Filter mode "${mode}" requires one string value.`);
			}
			return undefined;
		}
		candidate = value;
	}

	if (candidate === undefined || candidate.trim().length === 0) {
		if (options.required) {
			throw new Error(`Filter mode "${mode}" requires a non-empty value.`);
		}
		return undefined;
	}
	return candidate;
}

function getCountValue(mode: FilterRuleMode, value: string | string[] | undefined): string {
	let candidate = "";
	if (Array.isArray(value)) {
		if (value.length !== 1 || typeof value[0] !== "string") {
			throw new Error(`Filter mode "${mode}" requires a positive integer count value.`);
		}
		candidate = value[0].trim();
	} else if (typeof value === "string") {
		candidate = value.trim();
	}
	if (!/^[1-9]\d*$/.test(candidate) || !Number.isSafeInteger(Number(candidate))) {
		throw new Error(`Filter mode "${mode}" requires a positive integer count value.`);
	}
	return candidate;
}

function getRangeValues(value: string | string[] | undefined): [string, string] {
	if (!Array.isArray(value) || value.length !== 2 || value.some((item) => typeof item !== "string" || item.trim().length === 0)) {
		throw new Error('Filter mode "between" requires exactly two non-empty string values: [lower, upper].');
	}
	return [value[0], value[1]];
}

function isDateRangeValue(value: string): boolean {
	// Date-only values are the normal model-driven date input. Also accept an
	// ISO date-time suffix so a DateAndTime column keeps inclusive date bounds.
	return /^\d{4}-\d{2}-\d{2}(?:$|T\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:Z|[+-]\d{2}:?\d{2})?)$/.test(value.trim());
}

function makeCondition(
	columnName: string,
	mode: FilterRuleMode,
	value?: string | string[],
	strict = false
): DataSetConditionExpression | undefined {
	const extraDate = extraDateCondition(mode);
	if (extraDate) {
		const [, operator, , input] = extraDate;
		let compiledValue: string | string[] = '';
		if (input === 'periodYear') {
			if (!Array.isArray(value) || value.length !== 2 || !fiscalInteger(value[0], 13) || !fiscalInteger(value[1], 9999)) {
				throw new Error('Enter a fiscal period (1–13) and a fiscal year (1–9999).');
			}
			compiledValue = [...value];
		} else if (input !== 'none') {
			compiledValue = getCountValue(mode, value);
			if ((input === 'period' && !fiscalInteger(compiledValue, 13)) || (input === 'year' && !fiscalInteger(compiledValue, 9999))) {
				throw new Error(input === 'period' ? 'Enter a fiscal period from 1 to 13.' : 'Enter a fiscal year from 1 to 9999.');
			}
		}
		return {attributeName: columnName, conditionOperator: operator as DataSetConditionOperator, value: compiledValue};
	}
	if (mode === "doesNotContainData" || mode === "containsData") {
		return {
			attributeName: columnName,
			conditionOperator: mode === "containsData" ? NOT_NULL_OPERATOR : NULL_OPERATOR,
			value: "",
		};
	}

	if (mode === "between") {
		throw new Error('Filter mode "between" is a range and must be compiled as a condition group.');
	}

	if (NO_VALUE_MODES.has(mode)) {
		const conditionOperator = OPERATOR_BY_MODE[mode];
		if (conditionOperator === undefined) {
			throw new Error(`Filter mode "${mode}" is not supported by the model-driven dataset API.`);
		}
		return { attributeName: columnName, conditionOperator, value: "" };
	}

	if (COUNT_MODES.has(mode)) {
		const conditionOperator = OPERATOR_BY_MODE[mode];
		if (conditionOperator === undefined) {
			throw new Error(`Filter mode "${mode}" is not supported by the model-driven dataset API.`);
		}
		return { attributeName: columnName, conditionOperator, value: getCountValue(mode, value) };
	}

	if (!VALUE_MODES.has(mode)) {
		throw new Error(`Filter mode "${mode}" is not supported by the model-driven dataset API.`);
	}

	const rawValue = getSingleStringValue(mode, value, { strict, required: strict });
	if (rawValue === undefined) return undefined;
	const literalValue = rawValue.replace(/[[%_]/g, character => `[${character}]`);

	switch (mode) {
		case "doesNotContain":
			return { attributeName: columnName, conditionOperator: NOT_LIKE_OPERATOR, value: `%${literalValue}%` };
		case "doesNotBeginWith":
			return { attributeName: columnName, conditionOperator: NOT_LIKE_OPERATOR, value: `${literalValue}%` };
		case "doesNotEndWith":
			return { attributeName: columnName, conditionOperator: NOT_LIKE_OPERATOR, value: `%${literalValue}` };
		case "contains":
			return { attributeName: columnName, conditionOperator: LIKE_OPERATOR, value: `%${literalValue}%` };
		case "beginsWith":
			return { attributeName: columnName, conditionOperator: LIKE_OPERATOR, value: `${literalValue}%` };
		case "endsWith":
			return { attributeName: columnName, conditionOperator: LIKE_OPERATOR, value: `%${literalValue}` };
		case "notEquals":
			return { attributeName: columnName, conditionOperator: NOT_EQUAL_OPERATOR, value: rawValue };
		case "greaterThan":
			return { attributeName: columnName, conditionOperator: GREATER_THAN_OPERATOR, value: rawValue };
		case "greaterOrEqual":
			return { attributeName: columnName, conditionOperator: GREATER_OR_EQUAL_OPERATOR, value: rawValue };
		case "lessThan":
			return { attributeName: columnName, conditionOperator: LESS_THAN_OPERATOR, value: rawValue };
		case "lessOrEqual":
			return { attributeName: columnName, conditionOperator: LESS_OR_EQUAL_OPERATOR, value: rawValue };
		case "on":
			return { attributeName: columnName, conditionOperator: ON_OPERATOR, value: rawValue };
		case "onOrBefore":
			return { attributeName: columnName, conditionOperator: ON_OR_BEFORE_OPERATOR, value: rawValue };
		case "onOrAfter":
			return { attributeName: columnName, conditionOperator: ON_OR_AFTER_OPERATOR, value: rawValue };
		case "equals":
			return { attributeName: columnName, conditionOperator: EQUAL_OPERATOR, value: rawValue };
		default:
			throw new Error(`Filter mode "${mode}" is not supported by the model-driven dataset API.`);
	}
}

function makeBetweenFilter(columnName: string, value?: string | string[]): FilterExpression {
	const [lower, upper] = getRangeValues(value);
	const isDateRange = isDateRangeValue(lower) && isDateRangeValue(upper);
	return {
		filterOperator: 0,
		conditions: [
			{
				attributeName: columnName,
				conditionOperator: isDateRange ? ON_OR_AFTER_OPERATOR : GREATER_OR_EQUAL_OPERATOR,
				value: lower,
			},
			{
				attributeName: columnName,
				conditionOperator: isDateRange ? ON_OR_BEFORE_OPERATOR : LESS_OR_EQUAL_OPERATOR,
				value: upper,
			},
		],
	};
}

function makeRuleFilter(columnName: string, rule: ColumnFilterRule, strict: boolean): FilterExpression {
	if (!rule || typeof rule !== "object" || !isFilterRuleMode(rule.mode)) {
		const mode = rule && typeof rule === "object" ? String((rule as { mode?: unknown }).mode) : "";
		throw new Error(`Custom filter rule has unsupported mode "${mode}".`);
	}
	if (rule.mode === "between") return makeBetweenFilter(columnName, rule.value);
	const condition = makeCondition(columnName, rule.mode, rule.value, strict);
	if (!condition) {
		throw new Error(`Custom filter rule "${rule.mode}" requires a non-empty value.`);
	}
	return { filterOperator: 0, conditions: [condition] };
}

function validateCustomGroup(value: FilterValue | undefined): ColumnFilterGroup {
	if (!value || typeof value !== "object" || Array.isArray(value)) {
		throw new Error('Filter mode "custom" requires a ColumnFilterGroup value.');
	}
	const candidate = value as Partial<ColumnFilterGroup>;
	if (candidate.operator !== "and" && candidate.operator !== "or") {
		throw new Error('Custom filter group operator must be "and" or "or".');
	}
	if (!Array.isArray(candidate.rules)) {
		throw new Error("Custom filter group rules must be an array.");
	}
	if (candidate.rules.length < 1 || candidate.rules.length > 2) {
		throw new Error(`Custom filter group must contain between 1 and 2 rules; received ${candidate.rules.length}.`);
	}
	for (let index = 0; index < candidate.rules.length; index++) {
		const rule = candidate.rules[index];
		if (!rule || typeof rule !== "object" || !isFilterRuleMode(rule.mode)) {
			const mode = rule && typeof rule === "object" ? String((rule as { mode?: unknown }).mode) : "";
			throw new Error(`Custom filter rule at index ${index} has unsupported mode "${mode}".`);
		}
	}
	return {
		operator: candidate.operator,
		rules: candidate.rules,
	};
}

function makeCustomFilter(columnName: string, group: ColumnFilterGroup): FilterExpression {
	// Keep every custom rule as a child group. This makes an AND/OR custom group
	// explicit, and prevents a between range from being flattened into its parent.
	return {
		filterOperator: group.operator === "or" ? 1 : 0,
		conditions: [],
		filters: group.rules.map((rule) => makeRuleFilter(columnName, rule, true)),
	};
}

/**
 * Replace the conditions for one column while retaining every unrelated
 * condition and nested child filter. Lookup values are passed using the real
 * dataset attribute name; callers should provide raw lookup GUIDs.
 */
export function updateColumnFilter(
	current: DataSetFilterExpression | undefined,
	columnName: string,
	mode: FilterMode,
	value?: FilterValue
): FilterExpression | undefined {
	let validatedCustom: ColumnFilterGroup | undefined;
	if (mode === "custom") validatedCustom = validateCustomGroup(value);

	const cleaned = cloneFilterWithoutColumn(asFilter(current) ?? { filterOperator: 0, conditions: [] }, columnName);

	if (mode === "clear") return cleaned;

	let addition: FilterExpression | undefined;
	if (mode === "custom") {
		addition = makeCustomFilter(columnName, validatedCustom!);
	} else if (mode === "in") {
		if (!Array.isArray(value) || value.length === 0) return cleaned;
		if (value.some((item) => typeof item !== "string")) {
			throw new Error('Filter mode "in" requires an array of string values.');
		}
		addition = makeInFilter(columnName, value);
	} else if (mode === "between") {
		addition = makeBetweenFilter(columnName, value as string | string[] | undefined);
	} else {
		const condition = makeCondition(columnName, mode, value as string | string[] | undefined);
		addition = condition ? { filterOperator: 0, conditions: [condition] } : undefined;
	}
	if (!addition) return cleaned;

	return appendFilter(cleaned, addition);
}

function makeInFilter(columnName: string, values: string[]): FilterExpression | undefined {
	const hasBlank = values.some((value) => value === "");
	const nonBlank = values.filter((value) => value !== "");
	if (nonBlank.length === 0 && !hasBlank) return undefined;

	const nonBlankCondition: DataSetConditionExpression | undefined =
		nonBlank.length === 0
			? undefined
			: nonBlank.length === 1
				? { attributeName: columnName, conditionOperator: EQUAL_OPERATOR, value: nonBlank[0] }
				: { attributeName: columnName, conditionOperator: IN_OPERATOR, value: nonBlank };

	if (!hasBlank) {
		return nonBlankCondition ? { filterOperator: 0, conditions: [nonBlankCondition] } : undefined;
	}

	// A blank value is a Null condition. It cannot be represented as an empty
	// string inside In because Dataverse treats null and "" differently.
	const conditions: DataSetConditionExpression[] = [
		{ attributeName: columnName, conditionOperator: NULL_OPERATOR, value: "" },
	];
	if (nonBlankCondition) conditions.push(nonBlankCondition);
	return { filterOperator: 1, conditions };
}

function conditionValue(condition: DataSetConditionExpression): string | string[] | undefined {
	if (condition.value === undefined) return undefined;
	if (Array.isArray(condition.value)) {
		return condition.value.every((value) => typeof value === "string")
			? [...condition.value]
			: condition.value.map((value) => String(value));
	}
	return typeof condition.value === "string" ? condition.value : String(condition.value);
}

function asSingleDecodedValue(condition: DataSetConditionExpression): string | string[] | undefined {
	const value = conditionValue(condition);
	if (Array.isArray(value) && value.length === 1) return value[0];
	return value;
}

function decodeLikeRule(condition: DataSetConditionExpression): ColumnFilterRule {
	const value = asSingleDecodedValue(condition);
	const text = typeof value === "string" ? value : "";
	const startsWithWildcard = text.startsWith("%");
	const endsWithWildcard = text.endsWith("%");
	let mode: FilterRuleMode = "contains";
	if (startsWithWildcard && endsWithWildcard && text.length >= 2) mode = "contains";
	else if (endsWithWildcard) mode = "beginsWith";
	else if (startsWithWildcard) mode = "endsWith";
	if (condition.conditionOperator === NOT_LIKE_OPERATOR) {
		mode = mode === "beginsWith" ? "doesNotBeginWith" : mode === "endsWith" ? "doesNotEndWith" : "doesNotContain";
	}
	return {
		mode,
		value: text.replace(/^%/, "").replace(/%$/, "").replace(/\[([%_[])\]/g, '$1'),
	};
}

/** Decode a single supported dataset condition into the public rule model. */
export function deriveFilterRule(condition: DataSetConditionExpression | undefined): ColumnFilterRule | undefined {
	if (!condition) return undefined;
	const operator = Number(condition.conditionOperator);
	const value = asSingleDecodedValue(condition);
	const extraDate = extraDateEntries.find(([, definition]) => definition[1] === operator);
	if (extraDate) return extraDate[1][3] === 'none' ? {mode: extraDate[0]} : {mode: extraDate[0], value};
	const withValue = (mode: FilterRuleMode): ColumnFilterRule =>
		value === undefined ? { mode } : { mode, value };

	switch (operator) {
		case 0: return withValue("equals");
		case 1: return withValue("notEquals");
		case 2: return withValue("greaterThan");
		case 3: return withValue("lessThan");
		case 4: return withValue("greaterOrEqual");
		case 5: return withValue("lessOrEqual");
		case 6: return decodeLikeRule(condition);
		case 7: return decodeLikeRule(condition);
		case 12: return { mode: "doesNotContainData" };
		case 13: return { mode: "containsData" };
		case 14: return { mode: "yesterday" };
		case 15: return { mode: "today" };
		case 16: return { mode: "tomorrow" };
		case 17: return { mode: "last7Days" };
		case 18: return { mode: "next7Days" };
		case 19: return { mode: "lastWeek" };
		case 20: return { mode: "thisWeek" };
		case 22: return { mode: "lastMonth" };
		case 23: return { mode: "thisMonth" };
		case 25: return withValue("on");
		case 26: return withValue("onOrBefore");
		case 27: return withValue("onOrAfter");
		case 28: return { mode: "lastYear" };
		case 29: return { mode: "thisYear" };
		case 33: return withValue("lastXDays");
		case 34: return withValue("nextXDays");
		case 37: return withValue("lastXMonths");
		case 38: return withValue("nextXMonths");
		default:
			// In (8) is a checklist rather than a single condition-editor rule.
			return undefined;
	}
}

function decodeBetweenRule(expression: FilterExpression, columnName: string): ColumnFilterRule | undefined {
	if (expression.filterOperator !== 0 || (expression.filters ?? []).length > 0) return undefined;
	const conditions = expression.conditions ?? [];
	if (conditions.length !== 2 || conditions.some((condition) => condition.attributeName !== columnName)) return undefined;
	const lower = conditions.find((condition) => Number(condition.conditionOperator) === 4 || Number(condition.conditionOperator) === 27);
	const upper = conditions.find((condition) => Number(condition.conditionOperator) === 5 || Number(condition.conditionOperator) === 26);
	if (!lower || !upper || typeof lower.value !== "string" || typeof upper.value !== "string") return undefined;
	const numericPair = Number(lower.conditionOperator) === 4 && Number(upper.conditionOperator) === 5;
	const datePair = Number(lower.conditionOperator) === 27 && Number(upper.conditionOperator) === 26;
	if (!numericPair && !datePair) return undefined;
	return { mode: "between", value: [lower.value, upper.value] };
}

function decodeCustomGroup(expression: FilterExpression, columnName: string): ColumnFilterGroup | undefined {
	if (expression.filterOperator !== 0 && expression.filterOperator !== 1) return undefined;
	const childExpressions = expression.filters ?? [];
	if (childExpressions.length === 0) {
		// A direct between expression is useful to callers reopening a range, but
		// a plain one-condition filter is not a custom group.
		const between = decodeBetweenRule(expression, columnName);
		return between ? { operator: "and", rules: [between] } : undefined;
	}

	const directRules = (expression.conditions ?? []).map((condition) => {
		if (condition.attributeName !== columnName) return undefined;
		return deriveFilterRule(condition);
	});
	if (directRules.some((rule): rule is undefined => rule === undefined)) return undefined;

	const childRules: ColumnFilterRule[] = [];
	for (const child of childExpressions) {
		if ((child.conditions ?? []).some((condition) => condition.attributeName !== columnName)) return undefined;
		const between = decodeBetweenRule(child, columnName);
		if (between) {
			childRules.push(between);
			continue;
		}
		if (child.filterOperator !== 0 || (child.filters ?? []).length > 0 || (child.conditions ?? []).length !== 1) return undefined;
		const rule = deriveFilterRule(child.conditions?.[0]);
		if (!rule) return undefined;
		childRules.push(rule);
	}

	const rules = [...(directRules as ColumnFilterRule[]), ...childRules];
	if (rules.length < 1 || rules.length > 2) return undefined;
	return { operator: expression.filterOperator === 1 ? "or" : "and", rules };
}

/**
 * Find a custom group previously applied to one column. The traversal only
 * accepts groups whose rules all target that column, so unrelated nested view
 * filters are left alone. Between is decoded from its nested >=/<= pair.
 */
export function extractColumnFilterGroup(
	filter: DataSetFilterExpression | undefined,
	columnName: string
): ColumnFilterGroup | undefined {
	if (!filter) return undefined;
	const current = filter as FilterExpression;
	const ownGroup = decodeCustomGroup(current, columnName);
	if (ownGroup) return ownGroup;
	for (const child of current.filters ?? []) {
		const nested = extractColumnFilterGroup(child, columnName);
		if (nested) return nested;
	}
	return undefined;
}

/** Return every condition for the requested column, including child groups. */
export function getColumnConditions(filter: DataSetFilterExpression | undefined, columnName: string): DataSetConditionExpression[] {
	if (!filter) return [];
	const current = filter as FilterExpression;
	return [
		...(current.conditions ?? []).filter((condition) => condition.attributeName === columnName),
		...(current.filters ?? []).flatMap((child) => getColumnConditions(child, columnName)),
	];
}

/** Return unique filtered column names in traversal order. */
export function getFilterColumnNames(filter: DataSetFilterExpression | undefined): string[] {
	if (!filter) return [];
	const current = filter as FilterExpression;
	const names = [...(current.conditions ?? []).map((condition) => condition.attributeName)];
	for (const child of current.filters ?? []) names.push(...getFilterColumnNames(child));
	return Array.from(new Set(names));
}
