import {
	DemoFilterExpression,
	evaluateDemoFilter,
	getLocalDateOrdinal,
	sortDemoRows,
} from "./demo-filter-evaluator";

interface DemoTestRow {
	id: string;
	name: string;
	amount: number | null;
	status: number | null;
	ownerId: string | null;
	ownerName: string | null;
	date: string | null;
	active: boolean | null;
}

function assert(condition: boolean, message: string): void {
	if (!condition) throw new Error(message);
}

function dateAt(offsetDays: number): string {
	const base = new Date(2026, 8, 7, 12, 0, 0, 0);
	const date = new Date(base.getFullYear(), base.getMonth(), base.getDate() + offsetDays, 12, 0, 0, 0);
	return `${String(date.getFullYear()).padStart(4, "0")}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

const rows: DemoTestRow[] = [
	{ id: "a", name: "Alpha", amount: 10, status: 0, ownerId: "owner-z", ownerName: "Alex Morgan", date: dateAt(0), active: true },
	{ id: "b", name: "Beta", amount: null, status: 1, ownerId: "owner-a", ownerName: "Taylor Reed", date: null, active: false },
	{ id: "c", name: "alpha child", amount: 2, status: 2, ownerId: "owner-y", ownerName: "Jordan Patel", date: dateAt(-1), active: true },
	{ id: "d", name: "Alpha", amount: 10, status: 0, ownerId: "owner-x", ownerName: "Alex Morgan", date: dateAt(1), active: null },
	{ id: "e", name: "Gamma", amount: 30, status: null, ownerId: null, ownerName: null, date: dateAt(-400), active: false },
];

const options = {
	getValue: (row: DemoTestRow, column: string): unknown => row[column as keyof DemoTestRow],
	getDisplayValue: (row: DemoTestRow, column: string): string | null => column === "ownerId" ? row.ownerName : column === "status"
		? row.status === null ? null : ["Draft", "In progress", "Complete"][row.status]
		: column === "active" ? row.active === null ? null : row.active ? "Yes" : "No"
		: row[column as keyof DemoTestRow] as string | null,
	columnKinds: {
		name: "text" as const,
		amount: "number" as const,
		status: "option" as const,
		ownerId: "lookup" as const,
		date: "dateOnly" as const,
		active: "boolean" as const,
	},
	now: () => new Date(2026, 8, 7, 12, 0, 0, 0),
};

function filter(condition: { attributeName: string; conditionOperator: number; value?: unknown }): DemoFilterExpression {
	return { filterOperator: 0, conditions: [condition] };
}

function matchingIds(expression: DemoFilterExpression): string[] {
	return rows.filter((row) => evaluateDemoFilter(row, expression, options)).map((row) => row.id);
}

function testConditions(): void {
	assert(matchingIds(filter({attributeName:'name',conditionOperator:7,value:'%alpha%'})).join(',') === 'b,e','Does not contain must exclude matching names, case-insensitively.');
	assert(matchingIds(filter({attributeName:'name',conditionOperator:7,value:'alpha%'})).join(',') === 'b,e','Does not begin with must exclude the prefix.');
	assert(matchingIds(filter({attributeName:'name',conditionOperator:7,value:'%child'})).join(',') === 'a,b,d,e','Does not end with must exclude the suffix.');
	assert(matchingIds(filter({attributeName:'amount',conditionOperator:13})).join(',') === 'a,c,d,e','Contains data must exclude null without dropping numeric values.');
	assert(matchingIds(filter({attributeName:'ownerId',conditionOperator:7,value:'%never%'})).join(',') === 'a,b,c,d','Negative patterns must not include blank fields.');
	assert(matchingIds(filter({ attributeName: "amount", conditionOperator: 0, value: "10" })).join(",") === "a,d", "Numeric equality should coerce the dataset literal.");
	assert(matchingIds(filter({ attributeName: "amount", conditionOperator: 1, value: "10" })).join(",") === "c,e", "Null must not match not-equal.");
	assert(matchingIds(filter({ attributeName: "amount", conditionOperator: 2, value: 10 })).join(",") === "e", "Greater-than should compare numbers, not formatted labels.");
	assert(matchingIds(filter({ attributeName: "amount", conditionOperator: 3, value: 10 })).join(",") === "c", "Less-than should compare numbers.");
	assert(matchingIds(filter({ attributeName: "amount", conditionOperator: 4, value: 10 })).join(",") === "a,d,e", "Greater-than-or-equal should include equality.");
	assert(matchingIds(filter({ attributeName: "amount", conditionOperator: 5, value: 10 })).join(",") === "a,c,d", "Less-than-or-equal should include equality.");
	assert(matchingIds(filter({ attributeName: "name", conditionOperator: 6, value: "%alpha%" })).join(",") === "a,c,d", "Like should support percent wildcards case-insensitively.");
	assert(matchingIds(filter({ attributeName: "status", conditionOperator: 8, value: ["0", "2"] })).join(",") === "a,c,d", "In should compare option values rather than labels.");
	assert(matchingIds(filter({ attributeName: "amount", conditionOperator: 12 })).join(",") === "b", "Null should match only the null operator.");
	assert(matchingIds(filter({ attributeName: "amount", conditionOperator: 8, value: ["", "10"] })).join(",") === "a,d", "Null must not match In even when the list contains an empty literal.");
	assert(matchingIds(filter({ attributeName: "amount", conditionOperator: 6, value: "%" })).join(",") === "a,c,d,e", "Null must not match Like.");

	const nested: DemoFilterExpression = {
		filterOperator: 0,
		conditions: [{ attributeName: "amount", conditionOperator: 4, value: 2 }],
		filters: [{
			filterOperator: 1,
			conditions: [
				{ attributeName: "name", conditionOperator: 0, value: "Alpha" },
				{ attributeName: "status", conditionOperator: 0, value: "2" },
			],
		}],
	};
	assert(matchingIds(nested).join(",") === "a,c,d", "Nested AND/OR groups should preserve their own boolean operators.");
}

function testDates(): void {
	assert(getLocalDateOrdinal("2026-09-07T23:30:00-08:00") === getLocalDateOrdinal("2026-09-07"), "DateOnly parsing must ignore timestamp offsets.");
	const dateOptions = {
		getValue: (row: { date: string }) => row.date,
		columnKinds: { date: "dateOnly" as const },
		now: () => new Date(2026, 8, 7, 12, 0, 0, 0),
	};
	const dateCheck = (operator: number, offset: number, expected = ""): boolean => evaluateDemoFilter(
		{ date: dateAt(offset) },
		filter({ attributeName: "date", conditionOperator: operator, value: expected }),
		dateOptions,
	);
	assert(dateCheck(14, -1), "Yesterday should use the injected local date.");
	assert(dateCheck(15, 0), "Today should use the injected local date.");
	assert(dateCheck(16, 1), "Tomorrow should use the injected local date.");
	assert(dateCheck(17, -6) && dateCheck(17, 0) && !dateCheck(17, -7), "Last 7 Days includes today and the preceding six days.");
	assert(dateCheck(18, 0) && dateCheck(18, 7) && !dateCheck(18, 8), "Next 7 Days should use an inclusive local calendar range.");
	assert(dateCheck(19, -7) && dateCheck(19,-2) && !dateCheck(19, -1), "Last Week uses the documented Sunday-Saturday week.");
	assert(dateCheck(20, -1) && dateCheck(20, 5) && !dateCheck(20, 6), "This Week uses the documented Sunday-Saturday week.");
	assert(dateCheck(22, -31) && !dateCheck(22, -38), "Last Month should use the previous calendar month.");
	assert(dateCheck(23, 0) && !dateCheck(23, 24), "This Month should use the current calendar month.");
	assert(dateCheck(25, 0, dateAt(0)) && !dateCheck(25, 1, dateAt(0)), "On should compare local DateOnly values.");
	assert(dateCheck(26, -1, dateAt(0)) && dateCheck(26, 0, dateAt(0)) && !dateCheck(26, 1, dateAt(0)), "On or before should include equality.");
	assert(dateCheck(27, 1, dateAt(0)) && dateCheck(27, 0, dateAt(0)) && !dateCheck(27, -1, dateAt(0)), "On or after should include equality.");
	assert(dateCheck(28, -400) && !dateCheck(28, -800), "Last Year should use the previous calendar year.");
	assert(dateCheck(29, 0) && !dateCheck(29, 400), "This Year should use the current calendar year.");
	assert(dateCheck(33, -3, "3") && !dateCheck(33, -4, "3"), "Last X Days should honor its count.");
	assert(dateCheck(34, 3, "3") && !dateCheck(34, 4, "3"), "Next X Days should honor its count.");
	assert(dateCheck(37, -31, "1") && !dateCheck(37, -62, "1"), "Last X Months should use a rolling local calendar boundary.");
	assert(dateCheck(38, 30, "1") && !dateCheck(38, 31, "1"), "Next X Months should use a rolling local calendar boundary.");
}

function testSorting(): void {
	const sort = (columnName: string, direction: "asc" | "desc", input = rows): string[] => sortDemoRows(input, [{ columnName, direction }], options).map((row) => row.id);
	assert(sort("amount", "asc").join(",") === "c,a,d,e,b", "Ascending numbers should be numeric with blanks last.");
	assert(sort("amount", "desc").join(",") === "e,a,d,c,b", "Descending numbers should remain numeric with blanks last.");
	assert(sort("status", "asc").join(",") === "c,a,d,b,e", "OptionSet sorting should use displayed labels and keep null last.");
	assert(sort("ownerId", "asc").join(",") === "a,d,c,b,e", "Lookup sorting should use displayed names, not ids.");
	assert(sort("date", "asc").join(",") === "e,c,a,d,b", "DateOnly sorting should be chronological with blanks last.");
	assert(sort("date", "desc").join(",") === "d,a,c,e,b", "Descending DateOnly sorting should keep blanks last.");

	const multiRows = [{ ...rows[3], amount: 2 }, rows[0], rows[1], rows[2], rows[4]];
	const multi = sortDemoRows(multiRows, [
		{ columnName: "status", direction: "asc" },
		{ columnName: "amount", direction: "desc" },
	], options).map((row) => row.id);
	assert(multi.join(",") === "c,a,d,b,e", "Multi-sort should apply secondary levels within equal primary labels.");

	const ties = rows.filter((row) => row.id === "a" || row.id === "d");
	assert(sort("amount", "asc", ties).join(",") === "a,d", "Equal values should retain original order.");
}

testConditions();
testDates();
testSorting();
console.log("Demo filter evaluator checks passed.");
