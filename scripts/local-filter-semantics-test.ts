import {
	ColumnFilterGroup,
	FILTER_DEFINITIONS,
	deriveFilterRule,
	extractColumnFilterGroup,
	getColumnConditions,
	getFilterDefinitions,
	getFilterKind,
	updateColumnFilter,
} from "../CMModelDrivenGrid/filtering";

function assert(condition: boolean, message: string): void {
	if (!condition) throw new Error(message);
}

function assertDeepEqual(actual: unknown, expected: unknown, message: string): void {
	if (JSON.stringify(actual) !== JSON.stringify(expected)) {
		throw new Error(`${message}\nExpected: ${JSON.stringify(expected)}\nActual: ${JSON.stringify(actual)}`);
	}
}

function assertThrows(action: () => unknown, expectedText: string): void {
	try {
		action();
	} catch (error) {
		const message = error instanceof Error ? error.message : String(error);
		assert(message.includes(expectedText), `Expected error containing "${expectedText}", received "${message}".`);
		return;
	}
	throw new Error(`Expected an error containing "${expectedText}".`);
}

function conditionOperator(filter: unknown): number | undefined {
	const expression = filter as { conditions?: Array<{ conditionOperator?: number }> } | undefined;
	return expression?.conditions?.[0]?.conditionOperator;
}

function runOperatorMappingChecks(): void {
	const scalarMappings: Array<[string, number, string]> = [
		["equals", 0, "x"],
		["notEquals", 1, "x"],
		["contains", 6, "%x%"],
		["beginsWith", 6, "x%"],
		["endsWith", 6, "%x"],
		["greaterThan", 2, "x"],
		["lessThan", 3, "x"],
		["greaterOrEqual", 4, "x"],
		["lessOrEqual", 5, "x"],
		["on", 25, "x"],
		["onOrBefore", 26, "x"],
		["onOrAfter", 27, "x"],
		["lastXDays", 33, "4"],
		["nextXDays", 34, "4"],
		["lastXMonths", 37, "4"],
		["nextXMonths", 38, "4"],
	];
	for (const [mode, operator, value] of scalarMappings) {
		const result = updateColumnFilter(undefined, "createdon", mode as never, value);
		assert(conditionOperator(result) === operator, `${mode} should map to condition operator ${operator}.`);
		assert(
			(result as { conditions: Array<{ value?: string }> }).conditions[0].value === value ||
			mode === "contains" || mode === "beginsWith" || mode === "endsWith",
			`${mode} should preserve its condition value.`
		);
	}

	const relativeMappings: Array<[string, number]> = [
		["yesterday", 14], ["today", 15], ["tomorrow", 16], ["last7Days", 17], ["next7Days", 18],
		["lastWeek", 19], ["thisWeek", 20], ["lastMonth", 22], ["thisMonth", 23], ["lastYear", 28], ["thisYear", 29],
	];
	for (const [mode, operator] of relativeMappings) {
		assert(conditionOperator(updateColumnFilter(undefined, "createdon", mode as never)) === operator, `${mode} should map to condition operator ${operator}.`);
	}

	const between = updateColumnFilter(undefined, "amount", "between", ["10", "20"]);
	assertDeepEqual(
		(between as { conditions: unknown[] }).conditions,
		[
			{ attributeName: "amount", conditionOperator: 4, value: "10" },
			{ attributeName: "amount", conditionOperator: 5, value: "20" },
		],
		"Between should compile to >= lower and <= upper conditions."
	);
	const dateBetween = updateColumnFilter(undefined, "createdon", "between", ["2026-09-01", "2026-09-07"]);
	assertDeepEqual(
		(dateBetween as { conditions: unknown[] }).conditions,
		[
			{ attributeName: "createdon", conditionOperator: 27, value: "2026-09-01" },
			{ attributeName: "createdon", conditionOperator: 26, value: "2026-09-07" },
		],
		"Date between should compile to inclusive on-or-after/on-or-before conditions."
	);
	assertDeepEqual(
		extractColumnFilterGroup(dateBetween, "createdon"),
		{ operator: "and", rules: [{ mode: "between", value: ["2026-09-01", "2026-09-07"] }] },
		"Date range pairs should decode as a between rule when reopened."
	);

	for (const [operator, mode] of [[8, undefined]] as Array<[number, undefined]>) {
		assert(deriveFilterRule({ attributeName: "x", conditionOperator: operator as never, value: "" }) === mode, `Unsupported operator ${operator} should not be decoded.`);
	}
	assertDeepEqual(
		deriveFilterRule({ attributeName: "name", conditionOperator: 6 as never, value: "%Ada%" }),
		{ mode: "contains", value: "Ada" },
		"LIKE contains wildcard semantics should be retained."
	);
	assertDeepEqual(
		deriveFilterRule({ attributeName: "name", conditionOperator: 6 as never, value: "Ada%" }),
		{ mode: "beginsWith", value: "Ada" },
		"LIKE begins-with wildcard semantics should be decoded."
	);
}

function runCustomCompositionChecks(): void {
	const existing = {
		filterOperator: 0,
		conditions: [{ attributeName: "statecode", conditionOperator: 0, value: "0" }],
		filters: [{
			filterOperator: 1,
			conditions: [
				{ attributeName: "category", conditionOperator: 0, value: "A" },
				{ attributeName: "category", conditionOperator: 0, value: "B" },
			],
		}],
	} as Parameters<typeof updateColumnFilter>[0];
	const requested: ColumnFilterGroup = {
		operator: "or",
		rules: [
			{ mode: "equals", value: "Open" },
			{ mode: "between", value: ["10", "20"] },
		],
	};
	const result = updateColumnFilter(existing, "amount", "custom", requested);
	assert(getColumnConditions(result, "statecode").length === 1, "Custom replacement must retain unrelated root conditions.");
	assert(getColumnConditions(result, "category").length === 2, "Custom replacement must retain unrelated nested conditions.");
	const custom = extractColumnFilterGroup(result, "amount");
	assertDeepEqual(custom, requested, "Custom group should reopen with its OR operator and nested between rule.");
	const customExpression = (result as { filters?: Array<{ filterOperator?: number; filters?: unknown[] }> }).filters?.find(
		(group) => group.filterOperator === 1 && Array.isArray(group.filters)
	);
	assert(customExpression !== undefined, "Custom group should remain a nested OR expression beside existing filters.");

	const andGroup: ColumnFilterGroup = {
		operator: "and",
		rules: [
			{ mode: "greaterOrEqual", value: "10" },
			{ mode: "lessOrEqual", value: "20" },
		],
	};
	const directCustom = updateColumnFilter(undefined, "amount", "custom", andGroup);
	assertDeepEqual(extractColumnFilterGroup(directCustom, "amount"), andGroup, "A custom AND group should preserve two independent ordered rules.");

	const customBetweenOnly = updateColumnFilter(undefined, "amount", "custom", {
		operator: "and",
		rules: [{ mode: "between", value: ["10", "20"] }],
	});
	assertDeepEqual(
		extractColumnFilterGroup(customBetweenOnly, "amount"),
		{ operator: "and", rules: [{ mode: "between", value: ["10", "20"] }] },
		"A custom between subgroup should decode as one range rule."
	);
}

function runValidationChecks(): void {
	assertThrows(
		() => updateColumnFilter(undefined, "amount", "between", "10"),
		"between"
	);
	assertThrows(
		() => updateColumnFilter(undefined, "amount", "between", ["10"]),
		"exactly two"
	);
	assertThrows(
		() => updateColumnFilter(undefined, "amount", "between", ["10", ""]),
		"non-empty"
	);
	assertThrows(
		() => updateColumnFilter(undefined, "amount", "custom", { operator: "xor", rules: [{ mode: "equals", value: "x" }] } as never),
		"operator"
	);
	assertThrows(
		() => updateColumnFilter(undefined, "amount", "custom", { operator: "and", rules: [] }),
		"between 1 and 2"
	);
	assertThrows(
		() => updateColumnFilter(undefined, "amount", "custom", {
			operator: "and",
			rules: [
				{ mode: "equals", value: "a" },
				{ mode: "notEquals", value: "b" },
				{ mode: "contains", value: "c" },
			],
		}),
		"between 1 and 2"
	);
	assertThrows(() => updateColumnFilter(undefined, "createdon", "lastXDays"), "positive integer");
	assertThrows(() => updateColumnFilter(undefined, "createdon", "lastXDays", "0"), "positive integer");
	assertThrows(() => updateColumnFilter(undefined, "createdon", "lastXDays", "1.5"), "positive integer");
	assertThrows(() => updateColumnFilter(undefined, "createdon", "lastXDays", ["1", "2"]), "positive integer");

	const ordinaryMissing = updateColumnFilter(undefined, "name", "equals");
	assert(ordinaryMissing === undefined, "Existing ordinary value modes should still clear when no value is supplied.");
}

function runDataTypeChecks(): void {
	assert(getFilterKind("SingleLine.Text") === "text", "Single-line text should classify as text.");
	assert(getFilterKind("Decimal") === "number", "Decimal should classify as number.");
	assert(getFilterKind("Whole.None") === "number", "Whole number should classify as number.");
	assert(getFilterKind("DateAndTime.DateOnly") === "date", "Date-only should classify as date.");
	assert(getFilterKind("OptionSet") === "choice", "OptionSet should classify as choice.");
	assert(getFilterKind("TwoOptions") === "boolean", "TwoOptions should classify as boolean.");
	assert(getFilterKind("Lookup.Simple") === "lookup", "Lookup should classify as lookup.");
	assert(getFilterKind("") === "text", "Missing data type should safely default to text.");

	const textKeys = getFilterDefinitions("SingleLine.Text").map((definition) => definition.key);
	assertDeepEqual(textKeys,['equals','notEquals','containsData','doesNotContainData','contains','doesNotContain','beginsWith','doesNotBeginWith','endsWith','doesNotEndWith'],'Text filters must match the requested Power Apps order.');
	for (const mode of ['containsData','doesNotContain','doesNotBeginWith','doesNotEndWith'] as const) {
		const input=mode === 'containsData' ? undefined : 'Atlas';
		const encoded=updateColumnFilter(undefined,'name','custom',{operator:'and',rules:[input ? {mode,value:input} : {mode}]});
		assertDeepEqual(extractColumnFilterGroup(encoded,'name'),{operator:'and',rules:[input ? {mode,value:input} : {mode}]},`${mode} must round-trip through the editor.`);
	}
	assert(textKeys.includes("equals") && textKeys.includes("notEquals") && textKeys.includes("doesNotContainData"), "Text should expose blank, equals and not-equals filters.");
	assert(textKeys.includes("contains") && textKeys.includes("beginsWith") && textKeys.includes("endsWith"), "Text should expose text pattern filters.");
	assert(!textKeys.includes("between") && !textKeys.includes("today"), "Text should not expose numeric/date-only filters.");

	const numberKeys = getFilterDefinitions("Decimal").map((definition) => definition.key);
	assert(numberKeys.includes("greaterThan") && numberKeys.includes("between"), "Numbers should expose ordered and range filters.");
	assert(!numberKeys.includes("contains") && !numberKeys.includes("today"), "Numbers should not expose text-pattern/date-relative filters.");

	const dateKeys = getFilterDefinitions("DateAndTime.DateOnly").map((definition) => definition.key);
	assert(dateKeys.includes("on") && dateKeys.includes("onOrBefore") && dateKeys.includes("onOrAfter"), "Dates should expose date boundary filters.");
	assert(dateKeys.includes("today") && dateKeys.includes("lastXDays") && dateKeys.includes("between"), "Dates should expose relative/count/range filters.");
	assert(!dateKeys.includes("equals") && !dateKeys.includes("notEquals") && !dateKeys.includes("greaterOrEqual") && !dateKeys.includes("lessOrEqual"), "Dates should not expose generic equality or midnight-sensitive inclusive operators.");
	const dateDefinitions = getFilterDefinitions("DateAndTime.DateOnly");
	assert(dateDefinitions.find((definition) => definition.key === "greaterThan")?.text === "After", "Date greater-than should be labelled After.");
	assert(dateDefinitions.find((definition) => definition.key === "lessThan")?.text === "Before", "Date less-than should be labelled Before.");
	assert(!FILTER_DEFINITIONS.some((definition) => definition.key === ("notLike" as never)), "Unsupported notLike must not be masked as a definition.");
}

function run(): void {
	runOperatorMappingChecks();
	runCustomCompositionChecks();
	runValidationChecks();
	runDataTypeChecks();
	console.log("Local filter-semantics tests passed.");
}

try {
	run();
} catch (error) {
	console.error("Local filter-semantics tests failed:", error);
	process.exitCode = 1;
}
