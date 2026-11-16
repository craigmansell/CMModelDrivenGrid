import { JSDOM } from "jsdom";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { act, Simulate } from "react-dom/test-utils";
import type { GridProps } from "../CMModelDrivenGrid/Grid";
import type { FilterValue } from "../CMModelDrivenGrid/filtering";
import { fillColumnWidths } from "../CMModelDrivenGrid/columnSizing";
import type { SelectAllResult } from '../CMModelDrivenGrid/recordSelection';

const dom = new JSDOM("<!doctype html><html><body></body></html>", { url: "http://localhost/" });
globalThis.window = dom.window as unknown as Window & typeof globalThis;
globalThis.document = dom.window.document;
// These must exist before Fluent UI modules are required: merge-styles and
// Async capture the jsdom window during module initialization.
(globalThis.window as unknown as { requestAnimationFrame: (callback: FrameRequestCallback) => number }).requestAnimationFrame =
	(callback) => setTimeout(() => callback(0), 0) as unknown as number;
(globalThis.window as unknown as { cancelAnimationFrame: (id: number) => void }).cancelAnimationFrame =
	(id) => clearTimeout(id as unknown as ReturnType<typeof setTimeout>);
// Load Fluent UI only after the jsdom window exists. Its Async helper caches
// the global window at module initialization and Callout uses requestAnimationFrame.
const Utilities = require("@fluentui/react/lib/Utilities") as typeof import("@fluentui/react/lib/Utilities");
const UtilitiesCommonJs = require("@fluentui/react/lib-commonjs/Utilities") as typeof import("@fluentui/react/lib-commonjs/Utilities");
const { List } = require("@fluentui/react/lib/List") as typeof import("@fluentui/react/lib/List");
const { Grid } = require("../CMModelDrivenGrid/Grid") as { Grid: React.ComponentType<GridProps> };
const { selectionResult } = require('../CMModelDrivenGrid/recordSelection') as typeof import('../CMModelDrivenGrid/recordSelection');
const { initializeIcons } = require("@fluentui/react/lib/Icons") as typeof import("@fluentui/react/lib/Icons");
initializeIcons(undefined, { disableWarnings: true });
Object.defineProperty(dom.window.HTMLCanvasElement.prototype, "getContext", {
	value: () => ({ font: "", measureText: (text: string) => ({ width: text.length * 7 }) }),
	configurable: true,
});
Object.defineProperty(globalThis, "navigator", { value: dom.window.navigator, configurable: true });
(globalThis as unknown as { Document: typeof Document }).Document = dom.window.Document;
(globalThis as unknown as { Node: typeof Node }).Node = dom.window.Node;
(globalThis as unknown as { HTMLElement: typeof HTMLElement }).HTMLElement = dom.window.HTMLElement;
(globalThis as unknown as { Event: typeof Event }).Event = dom.window.Event;
(globalThis as unknown as { KeyboardEvent: typeof KeyboardEvent }).KeyboardEvent = dom.window.KeyboardEvent;
(globalThis as unknown as { requestAnimationFrame: (callback: FrameRequestCallback) => number }).requestAnimationFrame =
	(callback) => setTimeout(() => callback(0), 0) as unknown as number;
Object.defineProperty(globalThis.window, "innerHeight", { value: 900, configurable: true });
Object.defineProperty(globalThis.window, "innerWidth", { value: 1400, configurable: true });
Object.defineProperty(globalThis.window, "getComputedStyle", {
	value: () =>
		({
			direction: "ltr",
			overflowY: "auto",
			overflowX: "auto",
			paddingTop: "0px",
			paddingBottom: "0px",
			paddingLeft: "0px",
			paddingRight: "0px",
			getPropertyValue: () => "",
		}) as unknown as CSSStyleDeclaration,
	configurable: true,
});
Object.defineProperty(globalThis.window.HTMLElement.prototype, "offsetHeight", { get: () => 32 });
Object.defineProperty(globalThis.window.HTMLElement.prototype, "offsetWidth", { get: () => 180 });
Object.defineProperty(globalThis.window.HTMLElement.prototype, "clientHeight", { get: () => 600 });
Object.defineProperty(globalThis.window.HTMLElement.prototype, "clientWidth", { get: () => 1000 });
Object.defineProperty(globalThis.window.HTMLElement.prototype, "scrollHeight", { get: () => 1200 });
Object.defineProperty(globalThis.window.HTMLElement.prototype, "scrollWidth", { get: () => 1000 });
globalThis.window.HTMLElement.prototype.scrollTo = () => undefined;
(globalThis.window.HTMLElement.prototype as unknown as {attachEvent: () => void}).attachEvent = () => undefined;
(globalThis.window.HTMLElement.prototype as unknown as {detachEvent: () => void}).detachEvent = () => undefined;
(globalThis.window as unknown as { requestAnimationFrame: (callback: FrameRequestCallback) => number }).requestAnimationFrame =
	(callback) => setTimeout(() => callback(0), 0) as unknown as number;

class StubMutationObserver {
	public constructor(_callback: MutationCallback) { undefined; }
	public disconnect(): void { undefined; }
	public observe(_target: Node, _options?: MutationObserverInit): void { undefined; }
	public takeRecords(): MutationRecord[] { return []; }
}
(globalThis as unknown as { MutationObserver: typeof StubMutationObserver }).MutationObserver = StubMutationObserver;
(globalThis.window as unknown as { MutationObserver: typeof StubMutationObserver }).MutationObserver = StubMutationObserver;

const originalGetWindow = Utilities.getWindow;
(Utilities as unknown as { getWindow: typeof Utilities.getWindow }).getWindow =
	(element?: Element | null) => originalGetWindow(element) ?? (globalThis.window as unknown as Window);
const originalGetWindowCommonJs = UtilitiesCommonJs.getWindow;
(UtilitiesCommonJs as unknown as { getWindow: typeof Utilities.getWindow }).getWindow =
	(element?: Element | null) => originalGetWindowCommonJs(element) ?? (globalThis.window as unknown as Window);
(List as unknown as { prototype: { _shouldVirtualize: () => boolean } }).prototype._shouldVirtualize = () => false;

const eventGroupPrototype = (UtilitiesCommonJs.EventGroup as unknown as {
	prototype: {
		on: (target: EventTarget | undefined, eventName: string, callback: (event: Event) => void, options?: boolean) => void;
	};
}).prototype;
const originalEventGroupOn = eventGroupPrototype.on;
eventGroupPrototype.on = function (target, eventName, callback, options): void {
	if (!target) return;
	originalEventGroupOn.call(this, target, eventName, callback, options);
};

const resources: ComponentFramework.Resources = {
	getResource: () => { throw new Error("Not implemented for local test"); },
	getString: (id: string) => {
		switch (id) {
			case "Label_Grid_Footer_RecordCount": return "{0} records ({1} Selected)";
			case "Label_Grid_Footer": return "Page {0}";
			case "Label_NoRecords": return "No records found";
			case "Label_SortAZ": return "A to Z";
			case "Label_SortZA": return "Z to A";
			default: return id;
		}
	},
};

function assert(condition: boolean, message: string): void {
	if (!condition) throw new Error(message);
}

type TestRecord = ComponentFramework.PropertyHelper.DataSetApi.EntityRecord;
function makeRecord(id: string, rawValues: Record<string, unknown>, formattedValues: Record<string, string>): TestRecord {
	return {
		getRecordId: () => id,
		getValue: (columnName: string) => rawValues[columnName] as ReturnType<TestRecord['getValue']>,
		getFormattedValue: (columnName: string) => formattedValues[columnName] ?? "",
		getNamedReference: () => ({ id: { guid: id }, etn: "account", name: id }),
	};
}

function findEditableCell(container: HTMLElement, columnName: string): HTMLElement | undefined {
	return container.querySelector(`[aria-label^="${columnName}:"]`) as HTMLElement | undefined;
}

let intersectionCallback: ((entries: IntersectionObserverEntry[]) => void) | undefined;
class StubIntersectionObserver {
	public constructor(callback: IntersectionObserverCallback) { intersectionCallback = entries => callback(entries, this as unknown as IntersectionObserver); }
	public disconnect(): void { undefined; }
	public observe(_target: Element): void { undefined; }
	public unobserve(_target: Element): void { undefined; }
}
(globalThis as unknown as { IntersectionObserver: typeof StubIntersectionObserver }).IntersectionObserver = StubIntersectionObserver;
(globalThis.window as unknown as { IntersectionObserver: typeof StubIntersectionObserver }).IntersectionObserver = StubIntersectionObserver;

const container = document.createElement("div");
document.body.appendChild(container);

let currentGridProps: GridProps | undefined;
function createGridProps(overrides: Partial<GridProps>): GridProps {
	const defaults: GridProps = {
		width: 900,
		height: 500,
		columns: [],
		records: {},
		sortedRecordIds: [],
		hasNextPage: false,
		hasPreviousPage: false,
		totalResultCount: 0,
		currentPage: 1,
		sorting: [],
		filtering: undefined as unknown as ComponentFramework.PropertyHelper.DataSetApi.FilterExpression,
		resources,
		itemsLoading: false,
		highlightValue: null,
		highlightColor: null,
		enableLookupLinks: false,
		enableInlineEdit: true,
		setSelectedRecords: () => undefined,
		onNavigate: () => undefined,
		onOpenLookup: () => undefined,
		onSort: () => undefined,
		onFilter: () => undefined,
		onClearAllFilters: () => undefined,
		loadFirstPage: () => undefined,
		loadNextPage: () => undefined,
		loadPreviousPage: () => undefined,
		onUpdateCell: async () => undefined,
		onSelectAllRecords: async () => undefined,
	};
	return { ...defaults, ...overrides };
}

function renderGrid(overrides: Partial<GridProps>): void {
	ReactDOM.unmountComponentAtNode(container);
	currentGridProps = createGridProps(overrides);
	ReactDOM.render(React.createElement(Grid, currentGridProps), container);
}

function updateGrid(overrides: Partial<GridProps>): void {
	if (!currentGridProps) throw new Error("Grid has not been rendered.");
	currentGridProps = { ...currentGridProps, ...overrides };
	ReactDOM.render(React.createElement(Grid, currentGridProps), container);
}

async function testResponsiveSizing(): Promise<void> {
	const { ResponsiveGrid, resolveGridHeight } = require('../CMModelDrivenGrid/ResponsiveGrid') as typeof import('../CMModelDrivenGrid/ResponsiveGrid');
	assert(resolveGridHeight(850, 520) === 850, 'A definite Power Apps allocation must win over the embedded pixel height.');
	assert(resolveGridHeight(-1, 520, 850) === 520, 'An ordinary fixed-height subgrid must retain its configured fallback.');
	assert(resolveGridHeight(-1, undefined, 240) === 240, 'Automatic sizing must fit even a short available tab area.');
	assert(resolveGridHeight(Number.NaN, 0, Infinity) === -1, 'Unavailable or invalid sizes must retain CSS fill sizing.');
	const oldObserver = globalThis.ResizeObserver;
	const observers: SizingObserver[] = [];
	class SizingObserver {
		target?: Element;
		disconnected = false;
		constructor(private callback: ResizeObserverCallback) { observers.push(this); }
		observe(target: Element): void { this.target = target; }
		unobserve(): void { undefined; }
		disconnect(): void { this.disconnected = true; }
		resize(): void { this.callback([], this as unknown as ResizeObserver); }
	}
	Object.defineProperty(globalThis, 'ResizeObserver', { value: SizingObserver, configurable: true, writable: true });
	const column = { name:'name', displayName:'Name', dataType:'SingleLine.Text', order:0, visualSizeFactor:180 } as ComponentFramework.PropertyHelper.DataSetApi.Column;
	const record = makeRecord('parent', {name:'Parent'}, {name:'Parent'});
	let selected: string[] = [], queryCalls = 0, writes = 0;
	const props = createGridProps({width:undefined, height:undefined, columns:[column], records:{parent:record}, sortedRecordIds:['parent'],
		setSelectedRecords: ids => { selected = ids; }, onSort: () => { queryCalls++; }, onFilter: () => { queryCalls++; },
		onUpdateCell: async () => { writes++; }});
	const render = (height?: number, fixedHeight?: number) => ReactDOM.render(<ResponsiveGrid {...props} height={height} fixedHeight={fixedHeight}/>, container);
	try {
		await act(async () => { ReactDOM.unmountComponentAtNode(container); render(); });
		const host = container.querySelector<HTMLElement>('.cm-grid-host')!;
		const observer = observers.find(item => item.target === host)!;
		assert(!!observer, 'Automatic sizing must observe only the component-owned host.');
		let availableHeight = 720;
		Object.defineProperty(host, 'clientHeight', {get: () => availableHeight, configurable: true});
		const grid = () => host.querySelector<HTMLElement>('.cm-model-driven-grid')!;
		for (const height of [720, 900, 576, 240]) {
			availableHeight = height;
			await act(async () => { observer.resize(); await new Promise(resolve => setTimeout(resolve, 10)); });
			assert(grid().style.height === `${height}px` && grid().style.minHeight === '', `The grid must fill ${height}px without retaining a 320px minimum or stale height.`);
			assert(host.style.width === '100%' && grid().style.width === '100%', 'Full-tab sizing must also retain the available width.');
		}
		await act(async () => { (grid().querySelector('.ms-DetailsRow [role="checkbox"]') as HTMLElement).click(); });
		assert(selected.join() === 'parent', 'The resize fixture must have a real selected record.');
		availableHeight = 810;
		await act(async () => { window.dispatchEvent(new Event('resize')); await new Promise(resolve => setTimeout(resolve, 10)); });
		assert(grid().style.height === '810px' && selected.join() === 'parent', 'Window resizing must update the height while retaining selection.');
		await act(async () => { render(850, 520); });
		assert(host.style.height === '850px' && grid().style.height === '850px', 'Native allocated height must not be capped by the old fixed height.');
		assert(observer.disconnected, 'A definite host allocation must release the automatic observer.');
		await act(async () => { render(-1, 520); });
		assert(grid().style.height === '520px' && selected.join() === 'parent', 'Embedded fixed-height mode must remain available without remounting the grid.');
		assert(queryCalls === 0 && writes === 0, 'Sizing changes must never sort, filter or save records.');
		await act(async () => { render(); });
		const finalObserver = observers.filter(item => item.target === host).slice(-1)[0];
		await act(async () => { finalObserver.resize(); ReactDOM.unmountComponentAtNode(container); await new Promise(resolve => setTimeout(resolve, 10)); });
		assert(finalObserver.disconnected, 'Unmount must disconnect sizing observation and cancel queued work.');
	} finally {
		ReactDOM.unmountComponentAtNode(container);
		Object.defineProperty(globalThis, 'ResizeObserver', {value:oldObserver, configurable:true, writable:true});
	}
}

async function testColumnHeaderToggle(): Promise<void> {
	const columns = [
		{ name: 'name', displayName: 'Name', dataType: 'SingleLine.Text', order: 0, visualSizeFactor: 180 },
		{ name: 'amount', displayName: 'Amount', dataType: 'Currency', order: 1, visualSizeFactor: 120 },
	] as ComponentFramework.PropertyHelper.DataSetApi.Column[];
	let filterCalls = 0;
	let sortCalls = 0;
	const parent = makeRecord('parent', { name: 'Parent', amount: 10 }, { name: 'Parent', amount: '10.00' });
	const props = createGridProps({ columns, records: { parent }, sortedRecordIds: ['parent'],
		onFilter: () => { filterCalls++; }, onSort: () => { sortCalls++; } });
	const header = (scope: ParentNode, name: string) => {
		const element = scope.querySelector<HTMLElement>(`.ms-DetailsHeader-cellTitle[aria-label="${name}, column actions"]`);
		assert(!!element, `Missing ${name} column header.`);
		return element!;
	};
	const clickHeader = async (scope: ParentNode, name: string) => {
		await act(async () => {
			const element = header(scope, name);
			element.dispatchEvent(new window.MouseEvent('mousedown', { bubbles: true }));
			element.focus();
			element.dispatchEvent(new window.MouseEvent('mouseup', { bubbles: true }));
			// Target a nested label, as real header text and chevron clicks bubble.
			(element.querySelector('.ms-DetailsHeader-cellName') ?? element)
				.dispatchEvent(new window.MouseEvent('click', { bubbles: true, cancelable: true }));
			await new Promise(resolve => setTimeout(resolve, 5));
		});
	};
	const menu = () => document.querySelector('.cm-grid-column-callout');
	await act(async () => { renderGrid(props); });
	await clickHeader(container, 'Name');
	assert(!!menu() && header(container, 'Name').getAttribute('aria-expanded') === 'true', 'First header click must open the existing dropdown.');
	const search = menu()!.querySelector<HTMLInputElement>('[aria-label="Search filter values"]')!;
	await act(async () => { search.value = 'unsaved'; Simulate.change(search); });
	await clickHeader(container, 'Name');
	assert(!menu() && header(container, 'Name').getAttribute('aria-expanded') === 'false', 'Second click on the same header must close its dropdown.');
	assert(filterCalls === 0 && sortCalls === 0, 'Closing a header must discard unapplied choices without filtering or sorting.');
	await clickHeader(container, 'Name');
	assert(menu()?.querySelector<HTMLInputElement>('[aria-label="Search filter values"]')?.value === '', 'Reopening must start with clean, unapplied search state.');
	await clickHeader(container, 'Amount');
	assert(document.querySelectorAll('.cm-grid-column-callout').length === 1 && menu()?.querySelector('.cm-column-menu-title')?.textContent === 'Amount', 'Clicking another header must switch directly to its menu.');
	assert(header(container, 'Name').getAttribute('aria-expanded') === 'false' && header(container, 'Amount').getAttribute('aria-expanded') === 'true', 'Only the active header may be announced as expanded.');
	await act(async () => {
		const numberFilters = Array.from(menu()!.querySelectorAll('button')).find(button => button.querySelector('.ms-Button-label')?.textContent === 'Number filters');
		assert(!!numberFilters, 'The existing Number filters action must remain available.');
		Simulate.click(numberFilters!);
	});
	assert(!!menu()?.querySelector('.cm-filter-editor'), 'The existing condition editor must remain available.');
	await clickHeader(container, 'Amount');
	assert(!menu(), 'The same header must also close the condition-editor view.');
	await clickHeader(container, 'Amount');
	await act(async () => { document.body.dispatchEvent(new window.MouseEvent('click', { bubbles: true })); });
	assert(!menu(), 'Clicking outside must still dismiss the dropdown.');

	const { HierarchyGrid } = require('../CMModelDrivenGrid/HierarchyGrid') as typeof import('../CMModelDrivenGrid/HierarchyGrid');
	await act(async () => {
		ReactDOM.unmountComponentAtNode(container);
		ReactDOM.render(<HierarchyGrid {...props} childRows={{ loadPage: async () => ({
			records: [makeRecord('child', { name: 'Child', amount: 5 }, { name: 'Child', amount: '5.00' })], hasNextPage: false,
		}) }} />, container);
	});
	await act(async () => { Simulate.click(container.querySelector('button[aria-label="Expand child options for Parent"]')!); await Promise.resolve(); });
	const panel = container.querySelector('.cm-child-panel')!;
	await clickHeader(panel, 'Name');
	assert(!!menu(), 'Child column headers must open the same dropdown.');
	await clickHeader(panel, 'Name');
	assert(!menu(), 'A second child-header click must also close its dropdown.');
	assert(filterCalls === 0 && sortCalls === 0, 'Toggling child menus must not change the parent query.');
}

async function testRawFilterValues(): Promise<void> {
	const column: ComponentFramework.PropertyHelper.DataSetApi.Column = {
		name: "statuscode",
		displayName: "Status",
		dataType: "OptionSet",
		alias: "statuscode",
		order: 0,
		visualSizeFactor: 100,
		isHidden: false,
	};
	const records = {
		a: makeRecord("a", { statuscode: 100 }, { statuscode: "Active" }),
		b: makeRecord("b", { statuscode: 200 }, { statuscode: "Inactive" }),
		c: makeRecord("c", { statuscode: null }, { statuscode: "" }),
		d: makeRecord("d", { statuscode: { unsupported: true } }, { statuscode: "Unsupported" }),
	};
	const filters: Array<{ name: string; mode: string; value?: FilterValue }> = [];
	await act(async () => {
		renderGrid({ columns: [column], records, sortedRecordIds: ["a", "b", "c", "d"], totalResultCount: 4, onFilter: (name, mode, value) => filters.push({ name, mode, value }) });
	});
	const header = container.querySelector(".ms-DetailsHeader-cell:not(.ms-DetailsHeader-cellIsCheck) .ms-DetailsHeader-cellTitle");
	assert(!!header, "Filter regression: DetailsList header was not rendered.");
	await act(async () => {
		(header as HTMLElement).dispatchEvent(new window.MouseEvent("click", { bubbles: true, cancelable: true }));
		await Promise.resolve();
	});
	const callout = document.querySelector(".cm-grid-column-callout");
	assert(!!callout, "Filter regression: column callout did not open.");
	const valueList = callout?.querySelector('.cm-column-value-list[role="group"]');
	assert(!!valueList, "The values must share one bordered, labelled checklist.");
	assert(valueList?.querySelector(".ms-Checkbox .ms-Checkbox-text")?.textContent === "(Select All)", "Select All must be the first row inside the same checklist as the values.");
	assert(!!callout?.querySelector('.cm-column-menu-footer button') && !callout?.querySelector('.cm-column-menu-body .cm-filter-actions'), "Menu actions must stay outside the scrolling body.");
	assert(!callout?.textContent?.includes("Unsupported"), "Unsupported object values must not be offered as blanks.");
	const checkboxes = Array.from(callout?.querySelectorAll<HTMLInputElement>('input[type="checkbox"]') ?? []);
	for (const label of ["Inactive", "(Blanks)"]) {
		const input = checkboxes.find((candidate) => candidate.closest(".ms-Checkbox")?.textContent?.includes(label));
		assert(!!input, `Filter regression: '${label}' option was not rendered.`);
		await act(async () => { input!.checked=false; Simulate.change(input as HTMLInputElement, { target: input }); });
	}
	const ok = Array.from(callout?.querySelectorAll("button") ?? []).find((button) => button.textContent?.trim() === "OK");
	assert(!!ok, "Filter regression: OK button was not rendered.");
	await act(async () => { Simulate.click(ok as Element); await Promise.resolve(); });
	assert(filters.length === 1 && filters[0].mode === "equals" && filters[0].value === "100", `Expected raw option value 100, received ${JSON.stringify(filters)}.`);
}

async function testExcelSearchApplyAndCancel(): Promise<void> {
	const column = { name: "name", displayName: "Name", dataType: "SingleLine.Text", alias: "name", order: 0, visualSizeFactor: 100, isHidden: false };
	const filters: Array<{mode: string; value?: FilterValue}> = [];
	await act(async () => { renderGrid({ columns: [column], records: { a: makeRecord("a", {name: "Alpha"}, {name: "Alpha"}), b: makeRecord("b", {name: "Beta"}, {name: "Beta"}) }, sortedRecordIds: ["a", "b"], onFilter: (_name, mode, value) => filters.push({mode,value}) }); });
	const open = async () => { await act(async () => {
		const header = container.querySelector('.ms-DetailsHeader-cell:not(.ms-DetailsHeader-cellIsCheck) .ms-DetailsHeader-cellTitle');
		(header as HTMLElement).dispatchEvent(new window.MouseEvent("click", {bubbles:true})); await Promise.resolve();
	}); };
	const click = async (label: string) => { await act(async () => {
		const button = Array.from(document.querySelectorAll('.cm-grid-column-callout button')).find((element) => element.textContent?.trim() === label);
		assert(!!button, `Missing filter menu button ${label}`); Simulate.click(button as Element); await Promise.resolve();
	}); };
	await open();
	let search = document.querySelector<HTMLInputElement>('[aria-label="Search filter values"]');
	await act(async () => { if (search) search.value = "Alpha"; Simulate.change(search as HTMLInputElement); });
	await click("Cancel");
	assert(filters.length === 0, "Cancel must not apply a searched value.");
	await open();
	search = document.querySelector<HTMLInputElement>('[aria-label="Search filter values"]');
	await act(async () => { if (search) search.value = "Alpha"; Simulate.change(search as HTMLInputElement); });
	await click("OK");
	assert(filters.length === 1 && filters[0].mode === "equals" && filters[0].value === "Alpha", "Search and OK must apply only the matching values, not clear the filter.");
}

async function testSingleColumnSortMenu(): Promise<void> {
	const column = { name: "name", displayName: "Name", dataType: "SingleLine.Text", alias: "name", order: 0, visualSizeFactor: 100, isHidden: false };
	const sortCalls: Array<{ name: string; direction: string; additive?: boolean }> = [];
	await act(async () => {
		renderGrid({
			columns: [column], records: { a: makeRecord("a", { name: "Alpha" }, { name: "Alpha" }) }, sortedRecordIds: ["a"],
			sorting: [{ name: "createdon", sortDirection: 1 }],
			onSort: (name, direction, additive) => { sortCalls.push({ name, direction, additive }); },
		});
	});
	await act(async () => {
		const header = container.querySelector('.ms-DetailsHeader-cell:not(.ms-DetailsHeader-cellIsCheck) .ms-DetailsHeader-cellTitle');
		(header as HTMLElement).dispatchEvent(new window.MouseEvent("click", { bubbles: true }));
		await Promise.resolve();
	});
	const callout = document.querySelector(".cm-grid-column-callout");
	assert(!!callout && !callout.textContent?.includes("Add to current sort"), "The rejected Add to current sort item must stay absent.");
	const ascending = Array.from(callout?.querySelectorAll("button") ?? []).find((button) => button.querySelector(".ms-Button-label")?.textContent === "A to Z");
	assert(!!ascending, "Ascending sort action must remain available.");
	await act(async () => { Simulate.click(ascending as Element); await Promise.resolve(); });
	assert(sortCalls.length === 1 && sortCalls[0].name === "name" && sortCalls[0].direction === "asc" && !sortCalls[0].additive, "A new column sort must replace, not append to, the current sort.");
	await act(async()=>{updateGrid({sortedRecordIds:['a']});});
	await act(async()=>{(container.querySelector('.ms-DetailsHeader-cell:not(.ms-DetailsHeader-cellIsCheck) .ms-DetailsHeader-cellTitle') as HTMLElement).click();});
	const shifted=Array.from(document.querySelectorAll('.cm-grid-column-callout button')).find(button=>button.textContent?.includes('Z to A'))!;
	await act(async()=>{Simulate.click(shifted,{shiftKey:true});});
	assert(sortCalls[1]?.additive === true && sortCalls[1]?.direction === 'desc','Holding Shift on a direction must use additive sorting without an extra menu item.');
}

async function testColumnFillOnResize(): Promise<void> {
	const columns = ["a", "b"].map((name, order) => ({ name, displayName: name, dataType: "SingleLine.Text", alias: name, order, visualSizeFactor: 100, isHidden: false }));
	await act(async () => { renderGrid({ columns, records: { row: makeRecord("row", { a: "Alpha", b: "Beta" }, { a: "Alpha", b: "Beta" }) }, sortedRecordIds: ["row"] }); });
	const headerWidth = () => Array.from(container.querySelectorAll<HTMLElement>('.ms-DetailsHeader-cell:not(.ms-DetailsHeader-cellIsCheck)')).reduce((sum, element) => sum + parseFloat(element.style.width), 40);
	assert(headerWidth() === 1000, `Initial columns must fill the measured viewport, got ${headerWidth()}.`);
	const scrollPort = container.querySelector('.ms-ScrollablePane--contentContainer')!;
	Object.defineProperty(scrollPort, "clientWidth", { value: 1400, configurable: true });
	await act(async () => { window.dispatchEvent(new window.Event("resize")); });
	assert(headerWidth() === 1400, `Resizing must update Fluent's cached widths, got ${headerWidth()}.`);
}

async function testInfiniteScrollDistanceFallback(): Promise<void> {
	const previousObserver = globalThis.IntersectionObserver;
	Object.defineProperty(globalThis, "IntersectionObserver", { value: undefined, configurable: true, writable: true });
	try {
		for (const infiniteScroll of [undefined, true]) {
			let loadCalls = 0;
			await act(async () => { renderGrid({ infiniteScroll, hasNextPage: true, loadNextPage: () => { loadCalls++; } }); });
			const button = Array.from(container.querySelectorAll("button")).find((element) => element.textContent?.trim() === "Load more");
			assert(!button && !container.querySelector('button[aria-label="Next page"]'), "Default automatic scrolling must not show manual paging buttons.");
			assert(!container.querySelector('.cm-grid-jump-bar'), 'The alphabet bar must remain hidden unless explicitly enabled.');
			const port = container.querySelector('.ms-ScrollablePane--contentContainer') as HTMLElement;
			await act(async () => {port.scrollTop=500; port.dispatchEvent(new dom.window.Event('scroll'));});
			assert(loadCalls === 1, "Default and explicit infinite scrolling must request one page even without IntersectionObserver.");
		}
	} finally {
		Object.defineProperty(globalThis, "IntersectionObserver", { value: previousObserver, configurable: true, writable: true });
	}
}

async function testInfiniteScrollEmptyPageRetry(): Promise<void> {
	await act(async () => {
		renderGrid({ infiniteScroll: true, hasNextPage: true, records: { a: makeRecord("a", {}, {}) }, sortedRecordIds: ["a"],
			loadNextPage: () => { updateGrid({ records: { a: makeRecord("a", {}, {}) }, sortedRecordIds: ["a"] }); },
		});
	});
	assert(!Array.from(container.querySelectorAll("button")).some((button) => button.textContent?.trim() === "Load more"), "Automatic scrolling must not show the manual Load more button.");
	await act(async () => { intersectionCallback?.([{ isIntersecting: true } as IntersectionObserverEntry]); await Promise.resolve(); });
	assert(container.textContent?.includes("No additional records were returned.") === true, "A completed page with no progress must expose Retry instead of silently stalling.");
	assert(Array.from(container.querySelectorAll("button")).some((button) => button.textContent?.trim() === "Retry"), "An empty page must leave a recovery action.");
}

async function testInfiniteScrollWaitsForPagePromise(): Promise<void> {
	let loadCalls = 0;
	let completePage!: () => void;
	const page = new Promise<void>(resolve => { completePage = resolve; });
	const records = { a: makeRecord('a', {}, {}), b: makeRecord('b', {}, {}) };
	await act(async () => {
		renderGrid({ infiniteScroll: true, hasNextPage: true, records: { a: records.a }, sortedRecordIds: ['a'],
			loadNextPage: () => { loadCalls++; return page; },
		});
	});
	await act(async () => { intersectionCallback?.([{ isIntersecting: true } as IntersectionObserverEntry]); });
	assert(loadCalls === 1, 'The first page should be requested once.');
	// Dataverse renders the new payload before the paging promise and its
	// controller-side finally handlers have settled.
	await act(async () => { updateGrid({ records, sortedRecordIds: ['a', 'b'] }); });
	await act(async () => { intersectionCallback?.([{ isIntersecting: true } as IntersectionObserverEntry]); });
	assert(loadCalls === 1, 'New rows must not release a still-pending page promise and request the same host page again.');
	assert(!container.textContent?.includes('No additional records'), 'A successful pending page must not report an empty-page error.');
	await act(async () => { updateGrid({ hasNextPage: false }); completePage(); await page; });
	assert(!container.textContent?.includes('No additional records'), 'Late completion of the final page must not leave a spurious Retry warning.');
	assert(!container.querySelector('.cm-grid-loading-overlay'), 'The completed final page must release loading.');
}

async function testPagingYieldToSelection(): Promise<void> {
	let finish!: () => void;
	const page = new Promise<void>(resolve => {finish=resolve;});
	await act(async()=>{renderGrid({infiniteScroll:true,hasNextPage:true,records:{a:makeRecord('a',{}, {})},sortedRecordIds:['a'],loadNextPage:()=>page});});
	await act(async()=>{intersectionCallback?.([{isIntersecting:true} as IntersectionObserverEntry]);});
	await act(async()=>{updateGrid({itemsLoading:true,commandStatus:'Preparing record commands…'});finish();await page;});
	assert(!container.textContent?.includes('No additional records'),'Yielding an overlapping page to native selection must not report an empty-page error.');
	await act(async()=>{updateGrid({itemsLoading:false,commandStatus:undefined,hasNextPage:false});});
	assert(!container.textContent?.includes('No additional records'),'Selection completion must leave paging ready without a false Retry.');
}

async function testUnsupportedEditAndUntouchedDropdown(): Promise<void> {
	const unsupportedColumn: ComponentFramework.PropertyHelper.DataSetApi.Column = {
		name: "photo", displayName: "Photo", dataType: "Image", alias: "photo", order: 0, visualSizeFactor: 100, isHidden: false,
	};
	await act(async () => {
		renderGrid({ columns: [unsupportedColumn], records: { a: makeRecord("a", { photo: "https://example.invalid/photo" }, { photo: "Photo" }) }, sortedRecordIds: ["a"], totalResultCount: 1 });
	});
	const unsupportedCell = container.querySelector('[aria-label^="Photo:"]');
	assert(!!unsupportedCell, "Edit regression: unsupported cell was not rendered.");
	await act(async () => { Simulate.doubleClick(unsupportedCell as Element); });
	assert(!container.querySelector("input,select"), "Unsupported object/image columns must not enter inline edit mode.");

	const optionColumn: ComponentFramework.PropertyHelper.DataSetApi.Column = {
		name: "statuscode", displayName: "Status", dataType: "OptionSet", alias: "statuscode", order: 0, visualSizeFactor: 100, isHidden: false,
	};
	const updateCalls: string[] = [];
	await act(async () => {
		renderGrid({
			columns: [optionColumn],
			records: { a: makeRecord("a", { statuscode: null }, { statuscode: "" }) },
			sortedRecordIds: ["a"], totalResultCount: 1,
			columnOptions: { statuscode: [{ label: "Active", value: "100" }, { label: "Inactive", value: "200" }] },
			onUpdateCell: async (_id, _name, value) => { updateCalls.push(value); },
		});
	});
	// Let the initial Fluent FocusZone animation-frame focus settle before
	// starting an edit; otherwise its mount focus can immediately blur it.
	await act(async () => { await new Promise((resolve) => setTimeout(resolve, 25)); });
	const editableCell = findEditableCell(container, "Status");
	assert(!!editableCell, "Edit regression: OptionSet cell was not editable.");
	await act(async () => {
		editableCell?.focus();
		Simulate.doubleClick(editableCell as Element);
		await Promise.resolve();
	});
	const select = container.querySelector("select");
	assert(!!select, `Edit regression: OptionSet editor was not rendered. Cell: ${editableCell?.outerHTML}. Current editors: ${Array.from(container.querySelectorAll('input,select,[aria-label^="Status:"]')).map((element) => element.outerHTML).join("\n")}`);
	await act(async () => { Simulate.blur(select as HTMLSelectElement, { target: select }); await Promise.resolve(); });
	assert(updateCalls.length === 0, "Untouched dropdown blur must not commit an update.");
}

async function testClipboardFocusGuard(): Promise<void> {
	let writes = 0;
	Object.defineProperty(navigator, "clipboard", { value: { writeText: async () => { writes++; } }, configurable: true });
	const column: ComponentFramework.PropertyHelper.DataSetApi.Column = {
		name: "name", displayName: "Name", dataType: "SingleLine.Text", alias: "name", order: 0, visualSizeFactor: 100, isHidden: false,
	};
	await act(async () => {
		renderGrid({
			columns: [column],
			records: {
				a: makeRecord("a", { name: "Alpha" }, { name: "Alpha" }),
				b: makeRecord("b", { name: "Beta" }, { name: "Beta" }),
			},
			sortedRecordIds: ["a", "b"],
			totalResultCount: 2,
		});
	});
	const selectedCells = Array.from(container.querySelectorAll<HTMLElement>('[aria-label^="Name:"]'));
	assert(selectedCells.length >= 2, "Clipboard regression: data cells were not rendered.");
	await act(async () => {
		Simulate.mouseDown(selectedCells[0], { button: 0, shiftKey: false });
		Simulate.mouseEnter(selectedCells[1]);
		document.dispatchEvent(new window.MouseEvent("mouseup", { bubbles: true }));
	});
	const input = document.createElement("input");
	container.appendChild(input);
	input.dispatchEvent(new window.KeyboardEvent("keydown", { key: "c", ctrlKey: true, bubbles: true }));
	await new Promise((resolve) => setTimeout(resolve, 0));
	assert(writes === 0, "Grid clipboard handling must not steal Ctrl+C from focused inputs.");
	writes = 0;
	document.dispatchEvent(new window.KeyboardEvent("keydown", { key: "c", ctrlKey: true, bubbles: true }));
	await new Promise((resolve) => setTimeout(resolve, 0));
	assert(writes === 1, "Clipboard regression setup failed: a selected two-row range should copy once.");

	// Reordering/removing rows invalidates a range. A stale global cellSelection
	// must not copy the old rows after the host returns a new dataset order.
	writes = 0;
	await act(async () => {
		updateGrid({
			records: {
				b: makeRecord("b", { name: "Beta" }, { name: "Beta" }),
				a: makeRecord("a", { name: "Alpha" }, { name: "Alpha" }),
			},
			sortedRecordIds: ["b", "a"],
		});
		await Promise.resolve();
	});
	document.dispatchEvent(new window.KeyboardEvent("keydown", { key: "c", ctrlKey: true, bubbles: true }));
	await new Promise((resolve) => setTimeout(resolve, 0));
	assert(writes === 0, "Reordered rows must clear the prior cell range before Ctrl+C.");
}

async function testPagedNextButton(): Promise<void> {
	let loadCalls = 0;
	await act(async () => {
		renderGrid({
			columns: [], records: { a: makeRecord("a", {}, {}) }, sortedRecordIds: ["a"], totalResultCount: 2,
			hasNextPage: true, infiniteScroll: false,
			loadNextPage: () => { loadCalls++; },
		});
	});
	const nextButton = container.querySelector('button[aria-label="Next page"]');
	assert(!!nextButton, "Paging regression: Next page button was not rendered.");
	await act(async () => { Simulate.click(nextButton as Element); await Promise.resolve(); });
	assert(loadCalls === 1, `Paging regression: expected one Next page callback, got ${loadCalls}.`);
}

async function testInfiniteScrollErrorGuard(): Promise<void> {
	let loadCalls = 0;
	intersectionCallback = undefined;
	await act(async () => {
		renderGrid({
			columns: [], records: { a: makeRecord("a", {}, {}) }, sortedRecordIds: ["a"], totalResultCount: 2,
			hasNextPage: true, infiniteScroll: true,
			loadNextPage: () => { loadCalls++; throw new Error("synthetic page failure"); },
		});
	});
	assert(!!intersectionCallback, "Infinite-scroll regression: sentinel observer was not created.");
	await act(async () => { intersectionCallback?.([{ isIntersecting: true } as IntersectionObserverEntry]); await Promise.resolve(); });
	assert(loadCalls === 1, `Expected one failed automatic load, got ${loadCalls}.`);
	await act(async () => { intersectionCallback?.([{ isIntersecting: true } as IntersectionObserverEntry]); await Promise.resolve(); });
	assert(loadCalls === 1, "Observer callbacks must remain blocked while the load error is shown.");
	const retry = Array.from(container.querySelectorAll("button")).find((button) => button.textContent?.trim() === "Retry");
	assert(!!retry, "Infinite-scroll regression: Retry action was not rendered.");
	await act(async () => { Simulate.click(retry as Element); await Promise.resolve(); });
	assert(loadCalls === 2, "Manual Retry should clear the error guard and issue one new request.");
}

async function testInfiniteScrollIdentityGuard(): Promise<void> {
	let loadCalls = 0;
	intersectionCallback = undefined;
	const loadNextPage = (): void => {
		loadCalls++;
		updateGrid({
			records: {
				a: makeRecord("a", {}, {}),
				b: makeRecord("b", {}, {}),
			},
			sortedRecordIds: ["a", "b"],
			totalResultCount: 3,
		});
	};
	await act(async () => {
		renderGrid({
			columns: [], records: { a: makeRecord("a", {}, {}) }, sortedRecordIds: ["a"], totalResultCount: 3,
			hasNextPage: true, infiniteScroll: true, loadNextPage,
		});
	});
	assert(!!intersectionCallback, "Infinite-scroll identity regression: sentinel observer was not created.");
	await act(async () => { intersectionCallback?.([{ isIntersecting: true } as IntersectionObserverEntry]); await Promise.resolve(); });
	assert(loadCalls === 1, "A visible sentinel should request the first page exactly once.");
	await act(async () => { intersectionCallback?.([{ isIntersecting: true } as IntersectionObserverEntry]); await Promise.resolve(); });
	assert(loadCalls === 2, "Appending records should release the identity guard for the next page.");
	await act(async () => { intersectionCallback?.([{ isIntersecting: true } as IntersectionObserverEntry]); await Promise.resolve(); });
	assert(loadCalls === 2, "Repeated callbacks for the same loaded identity must not duplicate requests.");
}

async function run(): Promise<void> {
	console.log('Checking native/full-tab responsive sizing...');
	await testResponsiveSizing();
	console.log('Checking column-header open/close toggles...');
	await testColumnHeaderToggle();
	console.log('Checking grouping, complete totals and compact-list parity...');
	await testPresentationParity();
	await testExcelAdditions();
	await testLookupInteractionSafety();
	console.log('Checking full-view and bounded Select All...');
	await testBoundedSelectAll();
	console.log('Checking retained child panels...');
	await testChildExpansion();
	console.log('Checking hierarchy command selection...');
	await testHierarchySelection();
	await testLeafArrows();
	console.log('Checking filters, editing and scrolling...');
	const filled = fillColumnWidths([100, 100, 100], 1000);
	assert(filled.reduce((sum, value) => sum + value, 40 + 3 * 24) === 1000, "Columns must fill the viewport including selection and cell padding.");
	assert(JSON.stringify(fillColumnWidths([200, 100], 200)) === "[200,100]", "Narrow windows must retain readable widths and horizontal scrolling.");
	assert(fillColumnWidths([200, 100], 1000, [true, false])[0] === 200, "Manually sized columns must retain their width while automatic columns fill the gap.");
	assert(JSON.stringify(fillColumnWidths([200, 100], 1000, [true, true])) === "[200,712]", "The last column must absorb spare space when every column has been resized.");
	assert(fillColumnWidths([], 1000).length === 0, "An empty grid must not create columns.");
	assert(fillColumnWidths([100], Number.NaN)[0] === 100, "Unavailable viewport measurements must preserve preferred widths.");
	await testRawFilterValues();
	await testExcelSearchApplyAndCancel();
	await testSingleColumnSortMenu();
	await testColumnFillOnResize();
	await testInfiniteScrollDistanceFallback();
	await testInfiniteScrollEmptyPageRetry();
	await testInfiniteScrollWaitsForPagePromise();
	await testPagingYieldToSelection();
	await testUnsupportedEditAndUntouchedDropdown();
	await testClipboardFocusGuard();
	await testPagedNextButton();
	await testInfiniteScrollErrorGuard();
	await testInfiniteScrollIdentityGuard();
	await testClippedChildPaging();
	await testAncestorScrollRepaintsVirtualRows();
	await runReviewGates();
	console.log("Local grid regression tests passed.");
}

async function testBoundedSelectAll(): Promise<void> {
	const { HierarchyGrid } = require('../CMModelDrivenGrid/HierarchyGrid') as typeof import('../CMModelDrivenGrid/HierarchyGrid');
	const column = {name:'name',displayName:'Name',dataType:'SingleLine.Text',order:0,visualSizeFactor:180} as ComponentFramework.PropertyHelper.DataSetApi.Column;
	const row = (id:string) => makeRecord(id,{name:id},{name:id});
	const all = Array.from({length:2000},(_,i) => row(`parent-${i + 1}`));
	let selected:string[] = [], calls=0;
	let props = createGridProps({columns:[column],records:Object.fromEntries(all.slice(0,75).map(r=>[r.getRecordId(),r])),sortedRecordIds:all.slice(0,75).map(r=>r.getRecordId()),totalResultCount:2000,hasNextPage:true,setSelectedRecords:ids=>{selected=ids;}});
	const render = () => ReactDOM.render(<HierarchyGrid {...props}/>,container);
	const header = () => container.querySelector('.ms-DetailsHeader [role="checkbox"]') as HTMLElement;
	const click = async(element:HTMLElement) => {await act(async()=>{element.click();await Promise.resolve();});};
	const check = (id:string) => container.querySelector(`[aria-label="Name: ${id}"]`)!.closest('.ms-DetailsRow')!.querySelector('[role="checkbox"]') as HTMLElement;
	props.onSelectAllRecords = async () => {
		calls++;
		await Promise.resolve();
		props = {...props,records:Object.fromEntries(all.slice(0,1000).map(r=>[r.getRecordId(),r])),sortedRecordIds:all.slice(0,1000).map(r=>r.getRecordId())};
		render();
		return selectionResult(all.map(r=>r.getRecordId()));
	};
	await act(async()=>{ReactDOM.unmountComponentAtNode(container);render();});
	await click(header());
	assert(selected.length === 1000 && selected[999] === 'parent-1000','Select All must pass all 1,000 eligible IDs, not the original 75-row batch.');
	assert(container.querySelector('.cm-grid-record-count')?.textContent?.includes('1000 Selected') === true,'Shared selection footer must report the complete count.');
	assert(container.textContent?.includes('Dynamics limit') === true,'A capped selection must explain the 1,000-record limit.');
	await act(async()=>{props={...props,records:Object.fromEntries(all.slice(0,1200).map(r=>[r.getRecordId(),r])),sortedRecordIds:all.slice(0,1200).map(r=>r.getRecordId())};render();});
	assert(selected.length === 1000 && calls === 1,'Scrolling must neither shrink nor automatically extend the completed selection.');
	assert(check('parent-1000').getAttribute('aria-checked') === 'true' && check('parent-1001').getAttribute('aria-checked') === 'false','Later rows beyond the cap must not be ticked.');
	assert(header().getAttribute('aria-checked') === 'true','The capped header stays ticked after later rows arrive.');
	await click(header());
	assert(selected.length === 0,'A second header click clears the entire selection, even when not all loaded rows were selected.');
	await click(check('parent-2'));
	await click(header());
	assert(selected.length === 0 && calls === 1,'Header click with any existing selection clears it without loading more rows.');

	// Filtered sets below the cap and current ordering are preserved.
	const filtered=all.slice(1500,1517).reverse();
	props={...props,records:Object.fromEntries(filtered.map(r=>[r.getRecordId(),r])),sortedRecordIds:filtered.map(r=>r.getRecordId()),totalResultCount:17,hasNextPage:false,filtering:{filterOperator:0,conditions:[{attributeName:'name',conditionOperator:6,value:'filtered'}]},onSelectAllRecords:()=>selectionResult(filtered.map(r=>r.getRecordId()))};
	await act(async()=>{render();});
	await click(header());
	assert(selected.length === 17 && selected[0] === 'parent-1517' && selected[16] === 'parent-1501','Select All must respect the complete filtered/sorted result.');
	await click(header());

	// A late loader must not restore selections after cancel, filter change, or failure.
	let resolve!: (value:SelectAllResult)=>void;
	let signal:AbortSignal|undefined;
	props={...props,onSelectAllRecords:options=>{signal=options?.signal;return new Promise<SelectAllResult>(done=>{resolve=done;});}};
	await act(async()=>{render();});
	await click(header()); await click(header());
	assert(signal?.aborted === true,'The header must cancel an in-progress selection.');
	await act(async()=>{resolve(selectionResult(filtered.map(r=>r.getRecordId())));await Promise.resolve();});
	assert(selected.length === 0,'Cancelled selections must not be reapplied when the loader finishes.');
	await click(header());
	await act(async()=>{props={...props,sorting:[{name:'name',sortDirection:1}]};render();});
	assert(signal?.aborted === true,'Changing query order must cancel a pending selection.');
	await act(async()=>{resolve(selectionResult(filtered.map(r=>r.getRecordId())));await Promise.resolve();});
	assert(selected.length === 0,'An old query must never apply its IDs after a sort change.');
	props={...props,onSelectAllRecords:async()=>{throw new Error('Synthetic selection page failure');}};
	await act(async()=>{render();}); await click(header());
	assert(selected.length === 0 && container.textContent?.includes('Synthetic selection page failure') === true,'Failed paging must report an error without partial command targets.');

	// Children have their own full-view paging and the same cap, never siblings/parents.
	for (const childCount of [20,85,1200]) {
		const children=Array.from({length:childCount},(_,i)=>row(`child-${i+1}`));
		const pageSize = childCount > 1000 ? 250 : 25;
		let childLoads=0;
		const base=createGridProps({columns:[column],records:{p:row('p'),q:row('q')},sortedRecordIds:['p','q'],totalResultCount:2,setSelectedRecords:ids=>{selected=ids;}});
		await act(async()=>{ReactDOM.unmountComponentAtNode(container);ReactDOM.render(<HierarchyGrid {...base} childRows={{loadPage:async request=>{
			childLoads++;assert(request.parentId === 'p','Child Select All must stay inside the expanded parent.');
			return {records:children.slice((request.page-1)*pageSize,request.page*pageSize),hasNextPage:request.page*pageSize<children.length};
		}}}/>,container);});
		await click(container.querySelector('button[aria-label="Expand child options for p"]') as HTMLElement);
		await click(container.querySelector('.cm-child-panel .ms-DetailsHeader [role="checkbox"]') as HTMLElement);
		assert(selected.length === Math.min(1000,childCount) && selected[0] === 'child-1','Child Select All must select the full branch up to the cap.');
		assert(childLoads === Math.ceil(Math.min(1000,childCount)/pageSize),'Child Select All must stop loading at end-of-data or the 1,000-ID cap.');
		assert(container.querySelector('.cm-grid-footer:not(.cm-child-panel .cm-grid-footer) .cm-grid-record-count')?.textContent?.includes(`${selected.length} Selected`) === true,'Parent footer must agree with child command selection.');
		await click(container.querySelector('.cm-child-panel .ms-DetailsHeader [role="checkbox"]') as HTMLElement);
		assert(selected.length === 0,'Child header must clear its complete selection.');
	}
}

async function testAncestorScrollRepaintsVirtualRows(): Promise<void> {
	await act(async () => {ReactDOM.unmountComponentAtNode(container);renderGrid({});});
	const grid = container.querySelector<HTMLElement>('.cm-model-driven-grid')!;
	grid.getBoundingClientRect = () => ({top:100,bottom:600,left:0,right:1000,width:1000,height:500,x:0,y:100,toJSON:()=>({})});
	const prototype = List.prototype;
	const original = prototype.forceUpdate;
	let repaints = 0;
	prototype.forceUpdate = function () {repaints++;original.call(this);};
	try {
		await act(async () => {container.dispatchEvent(new Event('scroll'));await new Promise(resolve => setTimeout(resolve,10));});
		assert(repaints > 0,'Moving an ancestor must remeasure the nested virtual list when it returns to the viewport.');
		const before = repaints;
		grid.getBoundingClientRect = () => ({top:-1000,bottom:-500,left:0,right:1000,width:1000,height:500,x:0,y:-1000,toJSON:()=>({})});
		await act(async () => {container.dispatchEvent(new Event('scroll'));await new Promise(resolve => setTimeout(resolve,10));});
		assert(repaints === before,'Offscreen nested lists must not repaint on every ancestor scroll.');
	} finally {prototype.forceUpdate = original;}
}

async function testClippedChildPaging(): Promise<void> {
	const { HierarchyGrid } = require('../CMModelDrivenGrid/HierarchyGrid') as typeof import('../CMModelDrivenGrid/HierarchyGrid');
	const column = {name:'name',displayName:'Name',dataType:'SingleLine.Text',order:0,visualSizeFactor:180} as ComponentFramework.PropertyHelper.DataSetApi.Column;
	const childRows = Array.from({length:64},(_,index) => makeRecord(`child-${index+1}`,{name:`Child ${index+1}`},{name:`Child ${index+1}`}));
	const requested: number[] = [];
	let complete: ((page: import('../CMModelDrivenGrid/childRecords').ChildPage) => void) | undefined;
	const props = createGridProps({columns:[column],records:{p:makeRecord('p',{name:'Parent'},{name:'Parent'})},sortedRecordIds:['p'],totalResultCount:1,infiniteScroll:true});
	await act(async () => {
		ReactDOM.unmountComponentAtNode(container);
		ReactDOM.render(<HierarchyGrid {...props} childRows={{loadPage:async request => {
			requested.push(request.page);
			if (request.page === 1) return {records:childRows.slice(0,50),hasNextPage:true};
			return new Promise(resolve => {complete=resolve;});
		}}}/>,container);
	});
	await act(async () => {Simulate.click(container.querySelector('button[aria-label="Expand child options for Parent"]')!);await Promise.resolve();});
	assert(requested.join() === '1','Expansion must fetch only the first child batch.');
	const panel = container.querySelector('.cm-child-panel')!;
	const port = panel.querySelector('.ms-ScrollablePane--contentContainer')!;
	await act(async () => {port.scrollTop=550;port.dispatchEvent(new Event('scroll'));await Promise.resolve();});
	assert(requested.join() === '1,2','A clipped child sentinel must not prevent the scroll-distance fallback from loading page 2.');
	assert(!panel.textContent?.includes('Loading records'),'Loading another child batch must not block already loaded rows.');
	await act(async () => {port.scrollTop=0;complete?.({records:childRows.slice(50),hasNextPage:false});await Promise.resolve();});
	assert(!!panel.querySelector('[aria-label="Name: Child 64"]'),'The final child must be reachable after the next batch.');
	assert(!panel.querySelector('.cm-grid-load-error'),'Successful child paging must not show a false no-progress error.');
	assert(panel.querySelectorAll('.ms-DetailsRow').length === 64,'Child pages must merge without duplicates.');
}

async function testChildExpansion(): Promise<void> {
	const { HierarchyGrid } = require('../CMModelDrivenGrid/HierarchyGrid') as typeof import('../CMModelDrivenGrid/HierarchyGrid');
	const column = {name:'name',displayName:'Name',dataType:'SingleLine.Text',order:0,visualSizeFactor:180} as ComponentFramework.PropertyHelper.DataSetApi.Column;
	let resolvePage: ((page: import('../CMModelDrivenGrid/childRecords').ChildPage) => void) | undefined;
	let calls = 0;
	let added = 0;
	const config = {revision:0,onAdd:async()=>{added++;return false;},onDelete:async()=>false,loadPage: async () => {calls++;return new Promise<import('../CMModelDrivenGrid/childRecords').ChildPage>(resolve => {resolvePage=resolve;});}};
	const parentProps = createGridProps({columns:[column],records:{p:makeRecord('p',{name:'Parent'},{name:'Parent'})},sortedRecordIds:['p'],enableInlineEdit:true});
	await act(async () => {
		ReactDOM.unmountComponentAtNode(container);
		ReactDOM.render(<HierarchyGrid {...parentProps} childRows={config}/>,container);
	});
	assert(calls === 0,'Collapsed parents must not query children.');
	const toggle = () => container.querySelector('button[aria-label="Expand child options for Parent"],button[aria-label="Collapse child options for Parent"]')!;
	assert(!!toggle().closest('.ms-DetailsRow-cellCheck'), 'The chevron must sit beside the checkbox, outside the first data column.');
	assert(!!toggle().querySelector('[data-icon-name="ChevronRight"]'), 'Collapsed rows must use a Fluent chevron, not a solid triangle.');
	assert(!!container.querySelector('.cm-grid-selection-check'), 'Hierarchy rows must have an always-visible square selection indicator.');
	assert(!container.querySelector('[aria-label="Name: Parent"]')?.querySelector('.cm-child-toggle'), 'The first data value must not contain the hierarchy control.');
	const hierarchyHeaderWidth = Array.from(container.querySelectorAll<HTMLElement>('.ms-DetailsHeader-cell:not(.ms-DetailsHeader-cellIsCheck)')).reduce((sum, element) => sum + parseFloat(element.style.width), 56);
	assert(hierarchyHeaderWidth === 1000, `The compact hierarchy gutter and data columns must fill the viewport, got ${hierarchyHeaderWidth}.`);
	await act(async()=>{Simulate.click(toggle());await Promise.resolve();});
	assert(calls === 1,'Expanding must issue one child request.');
	assert(toggle().getAttribute('aria-expanded') === 'true','Expansion must expose accessible state.');
	assert(!!toggle().querySelector('[data-icon-name="ChevronDown"]'), 'Expanded rows must use a downward chevron.');
	await act(async()=>{Simulate.click(toggle());await Promise.resolve();});
	await act(async()=>{resolvePage?.({records:[makeRecord('late',{name:'Stale child'},{name:'Stale child'})],hasNextPage:false});await Promise.resolve();});
	assert(toggle().getAttribute('aria-expanded') === 'false' && !!container.querySelector('.cm-child-retained[hidden]'),'A late response must stay hidden after collapse.');
	await act(async()=>{Simulate.click(toggle());await Promise.resolve();});
	assert(calls === 1,'Reopening a retained panel must not repeat a completed query.');
	await act(async()=>{config.revision++;ReactDOM.render(<HierarchyGrid {...parentProps} childRows={config}/>,container);await Promise.resolve();});
	assert(calls === 2,'A host command-bar refresh must reload expanded children without a panel button.');
	await act(async()=>{resolvePage?.({records:[makeRecord('c',{name:'Loaded child'},{name:'Loaded child'})],hasNextPage:false});await Promise.resolve();});
	assert(!!container.textContent?.includes('Loaded child'),'Expanded children must render their values.');
	assert(!!container.querySelector('button[aria-label="Expand child options for Loaded child"]'),'Nested expansion must remain available.');
	const childPanel = container.querySelector('.cm-child-panel')!;
	assert(!childPanel.querySelector('.cm-child-toolbar'), 'Expanded children must not have a visible loaded-count heading or toolbar.');
	assert(!childPanel.textContent?.includes('Child options ·'), 'The child loaded-count heading must be removed, not only hidden.');
	assert(childPanel.querySelectorAll('[data-selection-toggle="true"]').length > 0,'Child rows must expose normal selection checkboxes.');
	assert(!childPanel.querySelector('button[title="Delete child option"]'),'Child commands must use the shared selection, not a per-row trash icon.');
	assert(!childPanel.querySelector('.cm-child-toolbar button'),'Child headers must contain no action buttons, even with legacy action callbacks configured.');
	const scroll = childPanel.querySelector('.ms-ScrollablePane--contentContainer')!;
	scroll.scrollTop = 123;
	await act(async()=>{Simulate.click(toggle());});
	await act(async()=>{Simulate.click(toggle());});
	assert(childPanel === container.querySelector('.cm-child-panel'),'Collapse must retain the same child component.');
	assert(scroll.scrollTop === 123,'Collapse must preserve child scroll position.');
	assert(added === 0,'Child panels must not execute legacy Add callbacks.');
	await act(async()=>{ReactDOM.render(<HierarchyGrid {...parentProps} childRows={{...config,readOnly:true}}/>,container);});
	assert(childPanel.getAttribute('aria-label') === 'Child options (read only)','Read-only child status must remain accessible without a visible heading.');
	assert(!Array.from(childPanel.querySelectorAll('button')).some(button=>button.textContent === 'Add child' || button.getAttribute('title') === 'Delete child option'),'Read-only children must have no mutation commands.');
	await act(async()=>{ReactDOM.render(<HierarchyGrid {...parentProps} disableSelection={true} childRows={config}/>,container);});
	assert(!!toggle() && !toggle().closest('.ms-DetailsRow-cellCheck'), 'Expansion must remain available when selection is disabled.');
}

async function testHierarchySelection(): Promise<void> {
	const { HierarchyGrid } = require('../CMModelDrivenGrid/HierarchyGrid') as typeof import('../CMModelDrivenGrid/HierarchyGrid');
	const column = {name:'name',displayName:'Name',dataType:'SingleLine.Text',order:0,visualSizeFactor:180} as ComponentFramework.PropertyHelper.DataSetApi.Column;
	const row = (id: string) => makeRecord(id,{name:id},{name:id});
	let selected: string[] = [];
	let loads = 0;
	const parentProps = createGridProps({columns:[column],records:{p:row('p'),q:row('q')},sortedRecordIds:['p','q'],totalResultCount:2,
		enableInlineEdit:true,setSelectedRecords:ids => {selected=ids;}});
	const config = {revision:0,loadPage:async ({parentId}: {parentId:string}) => {
		loads++;
		return {records:parentId === 'p' ? [row('a'),row('b')] : parentId === 'q' ? [row('c')] : [row('nested')],hasNextPage:false};
	}};
	const render = (props=parentProps) => ReactDOM.render(<HierarchyGrid {...props} childRows={config}/>,container);
	const toggle = (id:string) => container.querySelector(`button[aria-label="Expand child options for ${id}"],button[aria-label="Collapse child options for ${id}"]`) as HTMLElement;
	const check = (id:string) => container.querySelector(`[aria-label="Name: ${id}"]`)!.closest('.ms-DetailsRow')!.querySelector('[role="checkbox"]') as HTMLElement;
	const click = async (element:HTMLElement) => {await act(async()=>{element.click();await Promise.resolve();});};
	await act(async()=>{ReactDOM.unmountComponentAtNode(container);render();});
	await click(check('p'));
	assert(selected.join() === 'p','Parent selection must still reach the command bar.');
	await click(toggle('p'));
	assert(selected.join() === 'p','Clicking the adjacent chevron must not toggle the selection checkbox.');
	await click(check('a'));
	assert(selected.join() === 'a','Selecting a child must publish the child ID, not its parent.');
	assert(check('p').getAttribute('aria-checked') === 'false','Selecting a child must clear parent selection.');
	await click(check('b'));
	assert(selected.join() === 'a,b','Child multi-selection must publish both child IDs.');
	const rootFooter = container.querySelector('.cm-grid-footer:not(.cm-child-panel .cm-grid-footer) .cm-grid-record-count');
	assert(rootFooter?.textContent?.includes('2 Selected') === true,'The root footer must include selected child rows.');
	assert(loads === 1,'Selecting children must not reload them.');
	await click(toggle('q'));
	await click(check('c'));
	assert(selected.join() === 'c' && check('a').getAttribute('aria-checked') === 'false','Switching child panels must clear the old selection.');
	await click(check('p'));
	assert(selected.join() === 'p' && check('c').getAttribute('aria-checked') === 'false','Selecting a parent must clear the child selection.');
	await click(check('a'));
	await click(toggle('a'));
	await click(check('nested'));
	assert(selected.join() === 'nested','Nested descendants must use the same command selection.');
	await click(toggle('p'));
	assert(selected.length === 0,'Collapsing an ancestor must clear hidden descendant selection.');
	await click(toggle('p'));
	assert(check('nested').getAttribute('aria-checked') === 'false','Reopening must not restore hidden selected descendants.');
	await click(check('b'));
	await act(async()=>{config.revision++;render();await Promise.resolve();});
	assert(selected.length === 0,'Refreshing from the host command bar must clear stale child command targets.');
	await click(check('b'));
	await act(async()=>{ReactDOM.render(<HierarchyGrid {...parentProps} childRows={config} selectionResetVersion={1}/>,container);await Promise.resolve();});
	assert(check('b').getAttribute('aria-checked') === 'false','A rejected native selection must clear the child checkbox without publishing a parent.');
	await click(check('a'));
	await act(async()=>{render({...parentProps,records:{q:row('q')},sortedRecordIds:['q']});await Promise.resolve();});
	assert(selected.length === 0,'Removing a parent from the view must clear its selected children.');
}

async function testLeafArrows(): Promise<void> {
	const { HierarchyGrid } = require('../CMModelDrivenGrid/HierarchyGrid') as typeof import('../CMModelDrivenGrid/HierarchyGrid');
	const column={name:'name',displayName:'Name',dataType:'SingleLine.Text',order:0,visualSizeFactor:180} as ComponentFramework.PropertyHelper.DataSetApi.Column;
	const row=(id:string) => makeRecord(id,{name:id},{name:id});
	const parentProps=createGridProps({columns:[column],records:{p:row('p'),q:row('q')},sortedRecordIds:['p','q'],totalResultCount:2});
	let checks=0, loads=0, parentHasChildren=true;
	let resolveFirst!: (result:Record<string,boolean>) => void;
	const config={revision:0,hasChildren:async(ids:string[]) => {
		checks++;
		if (checks === 1) return new Promise<Record<string,boolean>>(resolve => {resolveFirst=resolve;});
		return Object.fromEntries(ids.map(id => [id,id === 'p' && parentHasChildren]));
	},loadPage:async()=>{loads++;return {records:[row('a'),row('b')],hasNextPage:false};}};
	const render=() => ReactDOM.render(<HierarchyGrid {...parentProps} childRows={config}/>,container);
	const arrow=(id:string) => container.querySelector(`button[aria-label="Expand child options for ${id}"],button[aria-label="Collapse child options for ${id}"]`) as HTMLButtonElement | null;
	await act(async()=>{ReactDOM.unmountComponentAtNode(container);render();});
	assert(!arrow('p') && !arrow('q'),'Unknown child existence must not create an empty expansion affordance.');
	await act(async()=>{resolveFirst({p:true,q:false});await Promise.resolve();});
	assert(!!arrow('p') && !arrow('q'),'Only a parent with actual children may show an arrow.');
	const leafCheck=container.querySelector('[aria-label="Name: q"]')!.closest('.ms-DetailsRow')!.querySelector('[role="checkbox"]');
	assert(!!leafCheck,'A leaf must retain its selection checkbox.');
	await act(async()=>{arrow('p')!.click();await Promise.resolve();});
	assert(!arrow('a') && !arrow('b'),'Leaf children must also hide their arrows.');
	assert(loads === 1,'Availability checks must not download child grids just to draw arrows.');
	await act(async()=>{render();await Promise.resolve();});
	assert(checks === 2,'Ordinary renders must reuse one batched existence check per level.');
	parentHasChildren=false;
	await act(async()=>{config.revision++;render();await Promise.resolve();});
	assert(!arrow('p'),'Refreshing after the final child disappears must hide the arrow.');
	assert(!Array.from(container.querySelectorAll('.cm-child-panel')).some(panel => !panel.closest('[hidden]')),'A confirmed empty parent must collapse its old child panel.');
	parentHasChildren=true;
	await act(async()=>{config.revision++;render();await Promise.resolve();});
	assert(!!arrow('p'),'Refreshing after a first child is added must restore the arrow.');
}

async function testPresentationParity(): Promise<void> {
	const columns = [{name:'name',displayName:'Name',dataType:'SingleLine.Text',order:0},{name:'amount',displayName:'Amount',dataType:'Decimal',order:1},{name:'status',displayName:'Status',dataType:'OptionSet',order:2}] as ComponentFramework.PropertyHelper.DataSetApi.Column[];
	const records = {a:makeRecord('a',{name:'Alpha',amount:10,status:1},{name:'Alpha',amount:'10',status:'Open'}),b:makeRecord('b',{name:'Beta',amount:20,status:1},{name:'Beta',amount:'20',status:'Open'}),c:makeRecord('c',{name:'Cedar',amount:null,status:2},{name:'Cedar',amount:'',status:'Closed'})};
	const cache = new Map<string,unknown>([['groupColumn','status'],['aggregates',{amount:'sum'}]]);
	let requests=0;
	await act(async()=>{renderGrid({columns,records,sortedRecordIds:['a','b','c'],stateCache:cache,infiniteScroll:true,hasNextPage:true,loadNextPage:()=>{requests++;}});});
	assert(!container.querySelector('[aria-label="Totals — all matching records"]'),'Never display complete totals before the last page.');
	assert(requests === 1,'Starting an analysis must request the next page without scrolling.');
	await act(async()=>{updateGrid({hasNextPage:false,itemsLoading:false});});
	assert(container.querySelector('[aria-label="Totals — all matching records"]')?.textContent?.includes('30'),'Complete sum must include all non-blank values.');
	assert(container.querySelectorAll('.ms-GroupHeader').length === 2,'Group by raw choices must render both groups.');
	assert(container.querySelectorAll('.ms-GroupHeader-expand[aria-expanded="false"]').length === 2,'New groups must start collapsed.');
	const collapse=container.querySelector('.ms-GroupHeader-expand') as HTMLElement;
	await act(async()=>{collapse.click();});
	assert(Object.values(cache.get('groupCollapse') as Record<string,boolean>).filter(value=>!value).length===1,'Group expansion must be retained in level state.');
	const collapseAll = container.querySelector('.ms-DetailsHeader-collapseButton') as HTMLElement;
	assert(!!collapseAll, 'Grouped columns must expose the global collapse control.');
	await act(async()=>{collapseAll.click();});
	assert(Object.values(cache.get('groupCollapse') as Record<string,boolean>).filter(Boolean).length===2, 'Collapse all must retain every group key.');
	await act(async()=>{updateGrid({selectionCountOverride:1});});
	assert(container.querySelectorAll('.ms-GroupHeader-expand[aria-expanded="false"]').length===2, 'An unrelated update must not reopen globally collapsed groups.');
	await act(async()=>{collapseAll.click();});
	assert(Object.values(cache.get('groupCollapse') as Record<string,boolean>).every(value=>!value), 'Expand all must clear retained collapse state.');
	await act(async()=>{updateGrid({dataError:'Page failed'});});
	assert(!container.querySelector('[aria-label="Totals — all matching records"]'),'A failed child page must not produce a complete-looking total.');
	await act(async()=>{updateGrid({dataError:undefined,features:{enableGrouping:false,enableAggregation:false}});});
	assert(!container.querySelector('.ms-GroupHeader') && !container.querySelector('.cm-grid-totals'),'Disabled grouping/aggregation must clear saved presentation settings.');
	let selected:string[]=[];
	await act(async()=>{renderGrid({columns,records,sortedRecordIds:['a','b','c'],features:{reflow:'list',listColumns:2,enableMultipleSelection:false,showAvatar:true},setSelectedRecords:ids=>{selected=ids;}});});
	assert(container.querySelectorAll('.cm-grid-list-row').length===3,'List mode must render compact record cards.');
	assert(container.querySelectorAll('.cm-grid-list-field').length===6,'List column count must limit visible card fields.');
	const checks=container.querySelectorAll<HTMLInputElement>('.cm-grid-list-row input[type="checkbox"]');
	await act(async()=>{checks[0].checked=true; Simulate.change(checks[0],{target:checks[0]});});
	assert(selected.join()==='a','List checkbox must publish the exact record ID.');
	await act(async()=>{checks[1].checked=true; Simulate.change(checks[1],{target:checks[1]});});
	assert(selected.join()==='b','Single-selection setting must replace, not append, the selected ID.');
	await act(async()=>{renderGrid({columns,records,sortedRecordIds:['a','b','c'],infiniteScroll:true,totalResultCount:-1,hasNextPage:true});});
	assert(container.querySelector('.cm-grid-record-count')?.textContent?.includes('3+'), 'An unknown count must describe loaded records, not invent 5,000 matches.');
}

async function testExcelAdditions(): Promise<void> {
	const {updateColumnFilter,getColumnConditions} = require('../CMModelDrivenGrid/filtering') as typeof import('../CMModelDrivenGrid/filtering');
	const columns=[{name:'name',displayName:'Name',dataType:'SingleLine.Text',order:0,isPrimary:true},{name:'amount',displayName:'Amount',dataType:'Decimal',order:1}] as ComponentFramework.PropertyHelper.DataSetApi.Column[];
	const records={a:makeRecord('a',{name:'Alpha blue',amount:10},{name:'Alpha blue',amount:'10'})};
	const base=updateColumnFilter(undefined,'name','contains','blue');
	let filter=base;
	await act(async()=>{renderGrid({columns,records,sortedRecordIds:['a'],features:{enableJumpBar:true},filtering:filter!,onFilter:(name,mode,value)=>{filter=updateColumnFilter(filter,name,mode,value);updateGrid({filtering:filter!,records:{...records},sortedRecordIds:['a']});}});});
	const jump=(letter:string)=>Array.from(container.querySelectorAll<HTMLButtonElement>('.cm-grid-jump-bar button')).find(button=>button.textContent===letter)!;
	await act(async()=>{jump('A').click();});
	assert(getColumnConditions(filter,'name').length===2,'Jump bar must keep the Excel filter on the same column.');
	await act(async()=>{jump('All').click();});
	assert(JSON.stringify(filter)===JSON.stringify(base),'Jump All must restore the exact non-alphabet filter.');
	await act(async()=>{updateGrid({sorting:[{name:'amount',sortDirection:0}]});});
	assert(jump('A').disabled,'Alphabet buttons must be disabled on a numeric sort.');
	const layoutKey='regression:org:user:view';
	await act(async()=>{renderGrid({columns,records,sortedRecordIds:['a'],layoutKey,stateCache:new Map([['columnOrder',['amount','name']],['widths',{amount:200}]])});});
	await act(async()=>{renderGrid({columns,records,sortedRecordIds:['a'],layoutKey});});
	const order=Array.from(container.querySelectorAll('.ms-DetailsHeader-cellTitle')).map(header=>header.textContent);
	assert(order[0]?.includes('Amount')===true,'Column order must survive a fresh grid mount for the same scope.');
	await act(async()=>{renderGrid({columns,records,sortedRecordIds:['a'],layoutKey:'different:user:view'});});
	assert(container.querySelector('.ms-DetailsHeader-cellTitle')?.textContent?.includes('Name')===true,'A different user/view must not inherit the layout.');
	let opened=0;
	await act(async()=>{renderGrid({columns,records,sortedRecordIds:['a'],enableLookupLinks:true,features:{navigationTypes:'primary'},onNavigate:()=>{opened++;}});});
	const primary=container.querySelector('[aria-label="Name: Alpha blue"] a') as HTMLElement;
	assert(!!primary,'Primary-only navigation must link an ordinary primary text column.');
	await act(async()=>{primary.click();}); assert(opened===1,'Primary name link must open the actual row.');
	await act(async()=>{updateGrid({features:{navigationTypes:'none'}});});
	assert(!container.querySelector('[aria-label="Name: Alpha blue"] a'),'Navigation none must remove primary links.');
	await act(async()=>{updateGrid({features:{navigationTypes:'none',reflow:'list'}});});
	assert(!container.querySelector('.cm-grid-record-link'),'Navigation none must also remove compact-list record links.');
}

async function testLookupInteractionSafety(): Promise<void> {
	const column={name:'lookup',displayName:'Lookup',dataType:'Lookup.Simple',order:0,alias:'lookup',visualSizeFactor:180,targets:['contact']} as ComponentFramework.PropertyHelper.DataSetApi.Column;
	const records={a:makeRecord('a',{lookup:{id:'old',entityType:'contact',name:'Old'}},{lookup:'Old'})};
	let writes:string[]=[];
	let resolveBrowse:(value:{id:string;name:string}|undefined)=>void=()=>undefined;
	const open=async()=>{
		await act(async()=>{renderGrid({columns:[column],records,sortedRecordIds:['a'],onSearchLookup:async()=>[{id:'new',name:'New name'}],onBrowseLookup:()=>new Promise(resolve=>{resolveBrowse=resolve;}),onUpdateLookupCell:async(_row,_col,id)=>{writes.push(id);}});});
		await act(async()=>{await new Promise(resolve=>setTimeout(resolve,25));});
		await act(async()=>{Simulate.doubleClick(container.querySelector('[aria-label="Lookup: Old"]')!); await Promise.resolve();});
	};
	await open();
	await act(async()=>{Simulate.keyDown(container.querySelector('input')!,{key:'Tab'});});
	assert(writes.length===0,'Tab must not silently accept the first lookup search result.');
	await open();
	await act(async()=>{Simulate.keyDown(container.querySelector('input')!,{key:'ArrowDown'});});
	assert(container.querySelector('input')?.getAttribute('aria-activedescendant')?.includes('results-')===true,'Lookup keyboard result must be announced accessibly.');
	await act(async()=>{Simulate.keyDown(container.querySelector('input')!,{key:'Enter'});await Promise.resolve();});
	assert(writes.join()==='new','ArrowDown + Enter must accept the explicitly highlighted lookup exactly once.');
	writes=[]; await open();
	const more=()=>Array.from(document.querySelectorAll<HTMLButtonElement>('button')).find(button=>button.textContent?.includes('More records'))!;
	await act(async()=>{more().click();});
	await act(async()=>{Simulate.blur(container.querySelector('input')!);await new Promise(resolve=>setTimeout(resolve,250));resolveBrowse(undefined);await Promise.resolve();});
	assert(writes.length===0 && !!container.querySelector('input') && document.activeElement===container.querySelector('input'),'Cancel full lookup must preserve the editor and refocus it, not clear the value.');
	await act(async()=>{more().click();resolveBrowse({id:'chosen',name:'Chosen'});await Promise.resolve();});
	assert(writes.join()==='chosen','Full dialog choice must use the existing guarded save path once.');
	writes=[]; await open();
	await act(async()=>{more().click();ReactDOM.unmountComponentAtNode(container);resolveBrowse({id:'late',name:'Late'});await Promise.resolve();});
	assert(writes.length===0,'A dialog closing after unmount must not save.');
}

// The release gates now run in the normal suite; retain the focused entry point.
const runReviewGates = () => (require('./local-pre-sandbox-review-probes') as typeof import('./local-pre-sandbox-review-probes')).runPreSandboxReviewProbes({container, resources, renderGrid, updateGrid, unmount: () => { ReactDOM.unmountComponentAtNode(container); }});
const runSelectedSuite = process.argv.includes('--layout') ? testResponsiveSizing : process.argv.includes('--pre-sandbox-review')
	? runReviewGates
	: run;
runSelectedSuite()
	.then(() => {
		ReactDOM.unmountComponentAtNode(container);
		process.exit(0);
	})
	.catch((error: unknown) => {
		console.error("Local grid regression tests failed:", error);
		ReactDOM.unmountComponentAtNode(container);
		process.exit(1);
	});
