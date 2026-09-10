import { JSDOM } from "jsdom";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { act, Simulate } from "react-dom/test-utils";
import type { ColumnFilterGroup, FilterDefinition, FilterRuleMode } from "../CMModelDrivenGrid/filtering";
import type { ColumnFilterEditorProps } from "../CMModelDrivenGrid/ColumnFilterEditor";

const dom = new JSDOM("<!doctype html><html><body></body></html>", { url: "http://localhost/" });
globalThis.window = dom.window as unknown as Window & typeof globalThis;
globalThis.document = dom.window.document;
(globalThis.window as unknown as { requestAnimationFrame: (callback: FrameRequestCallback) => number }).requestAnimationFrame =
	(callback) => setTimeout(() => callback(0), 0) as unknown as number;
(globalThis.window as unknown as { cancelAnimationFrame: (id: number) => void }).cancelAnimationFrame =
	(id) => clearTimeout(id as unknown as ReturnType<typeof setTimeout>);
Object.defineProperty(globalThis, "navigator", { value: dom.window.navigator, configurable: true });
(globalThis as unknown as { Document: typeof Document }).Document = dom.window.Document;
(globalThis as unknown as { Node: typeof Node }).Node = dom.window.Node;
(globalThis as unknown as { HTMLElement: typeof HTMLElement }).HTMLElement = dom.window.HTMLElement;
(globalThis as unknown as { Event: typeof Event }).Event = dom.window.Event;
(globalThis.window as unknown as { getComputedStyle: typeof window.getComputedStyle }).getComputedStyle =
	() => ({ getPropertyValue: () => "" }) as unknown as CSSStyleDeclaration;

const { ColumnFilterEditor } = require("../CMModelDrivenGrid/ColumnFilterEditor") as {
	ColumnFilterEditor: React.ComponentType<ColumnFilterEditorProps>;
};
const { getFilterDefinitions } = require("../CMModelDrivenGrid/filtering") as {
	getFilterDefinitions: (dataType: string) => FilterDefinition[];
};

function assert(condition: boolean, message: string): void {
	if (!condition) throw new Error(message);
}

function definition(dataType: string, mode: FilterRuleMode): FilterDefinition {
	const candidate = getFilterDefinitions(dataType).find((item) => item.key === mode);
	if (!candidate) throw new Error(`Missing ${mode} definition for ${dataType}.`);
	return candidate;
}

const container = document.createElement("div");
document.body.appendChild(container);

function renderEditor(props: ColumnFilterEditorProps): void {
	act(() => {
		ReactDOM.unmountComponentAtNode(container);
		ReactDOM.render(React.createElement(ColumnFilterEditor, props), container);
	});
}

function findButton(text: string): HTMLButtonElement {
	const button = Array.from(container.querySelectorAll("button")).find((candidate) => candidate.textContent?.trim() === text);
	if (!button) throw new Error(`Missing ${text} button.`);
	return button as HTMLButtonElement;
}

function click(text: string): void {
	act(() => Simulate.click(findButton(text)));
}

function testNumberRangeValidation(): void {
	const between = definition("Decimal", "between");
	let applied: ColumnFilterGroup | undefined;
	renderEditor({
		columnName: "Amount",
		dataType: "Decimal",
		initialGroup: { operator: "and", rules: [{ mode: between.key, value: ["10", "2"] }] },
		onApply: (group) => { applied = group; },
		onCancel: () => undefined,
	});

	const editorBody = container.querySelector(".cm-filter-editor-body");
	assert(!!editorBody?.querySelector("select"), "Filter fields must be inside the single scrolling editor body.");
	assert(!editorBody?.contains(findButton("Apply")) && !editorBody?.contains(findButton("Back")), "Editor header and actions must remain outside the scrolling body.");
	click("Apply");
	assert(!applied, "A reversed number range must not apply.");
	assert(container.textContent?.includes("first value must be less than or equal") === true, "Reversed range should show an inline error.");

	renderEditor({
		columnName: "Amount",
		dataType: "Decimal",
		initialGroup: { operator: "and", rules: [{ mode: between.key, value: ["2", "10"] }] },
		onApply: (group) => { applied = group; },
		onCancel: () => undefined,
	});
	click("Apply");
	assert(JSON.stringify(applied?.rules[0].value) === JSON.stringify(["2", "10"]), "A valid number range should submit both values.");
	const minimum = container.querySelector<HTMLInputElement>('input[id$="-from"]')!;
	const maximum = container.querySelector<HTMLInputElement>('input[id$="-to"]')!;
	act(() => { Simulate.change(minimum, { target: { value: "4" } }); });
	act(() => { Simulate.change(maximum, { target: { value: "18" } }); });
	click("Apply");
	assert(JSON.stringify(applied?.rules[0].value) === JSON.stringify(["4", "18"]), "Editing both bounds must preserve the other bound, not overwrite the whole range.");
}

function testTwoRuleAndOrSubmissionAndReopen(): void {
	const equals = definition("SingleLine.Text", "equals");
	let applied: ColumnFilterGroup | undefined;
	const onApply = (group: ColumnFilterGroup): void => { applied = group; };
	const initialGroup: ColumnFilterGroup = {
		operator: "and",
		rules: [
			{ mode: equals.key, value: "Alpha" },
			{ mode: equals.key, value: "Beta" },
		],
	};
	renderEditor({ columnName: "Name", dataType: "SingleLine.Text", initialGroup, onApply, onCancel: () => undefined });
	assert((container.querySelector('input[type="checkbox"]') as HTMLInputElement).checked, "Two initial rules should enable the second-condition checkbox.");
	click("Apply");
	assert(applied?.operator === "and" && applied.rules.length === 2, "Two rules should submit with AND.");

	const orGroup: ColumnFilterGroup = { ...initialGroup, operator: "or" };
	renderEditor({ columnName: "Name", dataType: "SingleLine.Text", initialGroup: orGroup, onApply, onCancel: () => undefined });
	click("Apply");
	assert(applied?.operator === "or" && applied.rules.length === 2, "Two rules should submit with OR.");

	renderEditor({ columnName: "Name", dataType: "SingleLine.Text", initialGroup: applied, onApply, onCancel: () => undefined });
	const reopenedSelects = Array.from(container.querySelectorAll("select")) as HTMLSelectElement[];
	assert(reopenedSelects[1]?.value === "or", "Reopened editor should restore the OR operator.");
	const reopenedInputs = Array.from(container.querySelectorAll('input[type="text"]')) as HTMLInputElement[];
	assert(reopenedInputs.length === 2 && reopenedInputs[0].value === "Alpha" && reopenedInputs[1].value === "Beta", "Reopened editor should restore both rule values.");
}

function testDateRelativeAndCountInputs(): void {
	const today = definition("DateAndTime.DateOnly", "today");
	const lastXDays = definition("DateAndTime.DateOnly", "lastXDays");
	let applied: ColumnFilterGroup | undefined;
	const onApply = (group: ColumnFilterGroup): void => { applied = group; };
	renderEditor({
		columnName: "Start date",
		dataType: "DateAndTime.DateOnly",
		initialGroup: { operator: "and", rules: [{ mode: today.key }] },
		onApply,
		onCancel: () => undefined,
	});
	assert(container.querySelectorAll('input:not([type="checkbox"])').length === 0, "Date-relative no-input modes should not render a value input.");
	click("Apply");
	assert(applied?.rules[0].mode === today.key && !Object.prototype.hasOwnProperty.call(applied.rules[0], "value"), "Date-relative no-input mode should submit without a value.");

	renderEditor({
		columnName: "Start date",
		dataType: "DateAndTime.DateOnly",
		initialGroup: { operator: "and", rules: [{ mode: lastXDays.key, value: "0" }] },
		onApply,
		onCancel: () => undefined,
	});
	assert(container.querySelector('input[type="number"]') !== null, "Date count mode should render a number input.");
	click("Apply");
	assert(container.textContent?.includes("positive whole number") === true, "Date count should reject zero inline.");
	renderEditor({
		columnName: "Start date",
		dataType: "DateAndTime.DateOnly",
		initialGroup: { operator: "and", rules: [{ mode: lastXDays.key, value: "3" }] },
		onApply,
		onCancel: () => undefined,
	});
	click("Apply");
	assert(applied?.rules[0].mode === lastXDays.key && applied.rules[0].value === "3", "Date count should submit a positive integer string.");
}

function testLookupUsesRawOptionValues(): void {
	const equals = definition("Lookup.Simple", "equals");
	let applied: ColumnFilterGroup | undefined;
	renderEditor({
		columnName: "Owner",
		dataType: "Lookup.Simple",
		initialGroup: { operator: "and", rules: [{ mode: equals.key, value: "guid-123" }] },
		valueOptions: [{ label: "Contoso", value: "guid-123" }],
		onApply: (group) => { applied = group; },
		onCancel: () => undefined,
	});
	const select = (Array.from(container.querySelectorAll("select"))[1] ?? null) as HTMLSelectElement | null;
	assert(select?.value === "guid-123", "Lookup select should preserve the raw option value.");
	assert(select?.options[1]?.text === "Contoso", "Lookup select should display the friendly option label.");
	assert(container.querySelector('input:not([type="checkbox"])') === null, "Lookup filters must not offer an arbitrary display-name/GUID text input.");
	click("Apply");
	assert(applied?.rules[0].value === "guid-123", "Lookup filter should submit the raw option value.");
}

testNumberRangeValidation();
testTwoRuleAndOrSubmissionAndReopen();
testDateRelativeAndCountInputs();
testLookupUsesRawOptionValues();

ReactDOM.unmountComponentAtNode(container);
container.remove();
console.log("local-filter-editor-test: passed");
