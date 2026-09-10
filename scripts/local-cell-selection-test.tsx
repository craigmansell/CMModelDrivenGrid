import { JSDOM } from "jsdom";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { act, Simulate } from "react-dom/test-utils";
import * as Utilities from "@fluentui/react/lib/Utilities";
import * as UtilitiesCommonJs from "@fluentui/react/lib-commonjs/Utilities";
import { List } from "@fluentui/react/lib/List";
import { Grid } from "../CMModelDrivenGrid/Grid";
import { initializeIcons } from "@fluentui/react/lib/Icons";

const dom = new JSDOM("<!doctype html><html><body></body></html>", {
	url: "http://localhost/",
});

globalThis.window = dom.window as unknown as Window & typeof globalThis;
globalThis.document = dom.window.document;
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
(globalThis as unknown as { requestAnimationFrame: (cb: FrameRequestCallback) => number }).requestAnimationFrame = (cb) =>
	setTimeout(() => cb(0), 0) as unknown as number;

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
		}) as CSSStyleDeclaration,
	configurable: true,
});

if (!globalThis.window.HTMLElement.prototype.getBoundingClientRect) {
	globalThis.window.HTMLElement.prototype.getBoundingClientRect = () =>
		({ x: 0, y: 0, top: 0, left: 0, right: 1000, bottom: 600, width: 1000, height: 600, toJSON: () => "" }) as DOMRect;
}
Object.defineProperty(globalThis.window.HTMLElement.prototype, "offsetHeight", { get: () => 32 });
Object.defineProperty(globalThis.window.HTMLElement.prototype, "offsetWidth", { get: () => 180 });
Object.defineProperty(globalThis.window.HTMLElement.prototype, "clientHeight", { get: () => 600 });
Object.defineProperty(globalThis.window.HTMLElement.prototype, "clientWidth", { get: () => 1000 });
Object.defineProperty(globalThis.window.HTMLElement.prototype, "scrollHeight", { get: () => 1200 });
Object.defineProperty(globalThis.window.HTMLElement.prototype, "scrollWidth", { get: () => 1000 });
(globalThis.window.HTMLElement.prototype as unknown as { scrollTo: () => void }).scrollTo = () => undefined;
(globalThis.window.HTMLElement.prototype as unknown as { attachEvent?: () => void }).attachEvent = () => undefined;
(globalThis.window.HTMLElement.prototype as unknown as { detachEvent?: () => void }).detachEvent = () => undefined;

if (!(globalThis as unknown as { MutationObserver?: unknown }).MutationObserver) {
	class StubMutationObserver {
		public constructor(_cb: MutationCallback) { undefined; }
		public disconnect(): void { undefined; }
		public observe(_t: Node, _o?: MutationObserverInit): void { undefined; }
		public takeRecords(): MutationRecord[] { return []; }
	}
	(globalThis as unknown as { MutationObserver: unknown }).MutationObserver = StubMutationObserver;
	(globalThis.window as unknown as { MutationObserver: unknown }).MutationObserver = StubMutationObserver;
}

const origGetWindow = Utilities.getWindow;
(Utilities as unknown as { getWindow: typeof Utilities.getWindow }).getWindow = (el?: Element | null) =>
	origGetWindow(el) ?? (globalThis.window as unknown as Window);
const origGetWindowCjs = UtilitiesCommonJs.getWindow;
(UtilitiesCommonJs as unknown as { getWindow: typeof UtilitiesCommonJs.getWindow }).getWindow = (el?: Element | null) =>
	origGetWindowCjs(el) ?? (globalThis.window as unknown as Window);

(List as unknown as { prototype: { _shouldVirtualize: () => boolean } }).prototype._shouldVirtualize = () => false;

// Fluent SelectionZone attaches listeners via EventGroup; ignore undefined targets in jsdom.
const eventGroupPrototype = (UtilitiesCommonJs.EventGroup as unknown as {
	prototype: {
		on: (target: EventTarget | undefined, eventName: string, callback: (event: Event) => void, options?: boolean) => void;
	};
}).prototype;
const originalEventGroupOn = eventGroupPrototype.on;
eventGroupPrototype.on = function (
	target: EventTarget | undefined,
	eventName: string,
	callback: (event: Event) => void,
	options?: boolean
): void {
	if (!target) return;
	originalEventGroupOn.call(this, target, eventName, callback, options);
};

// ── Mock clipboard ────────────────────────────────────────────────────────────
let clipboardText = "";
Object.defineProperty(globalThis.navigator, "clipboard", {
	value: { writeText: async (text: string) => { clipboardText = text; } },
	configurable: true,
	writable: true,
});

// ── Test data ─────────────────────────────────────────────────────────────────
function makeRecord(id: string, vals: Record<string, string>): ComponentFramework.PropertyHelper.DataSetApi.EntityRecord {
	return {
		getRecordId: () => id,
		getFormattedValue: (col: string) => vals[col] ?? "",
		getValue: (col: string) => vals[col] ?? "",
		getNamedReference: () => ({ id: { guid: id }, etn: "account", name: id }),
	};
}

const records = {
	r1: makeRecord("r1", { col1: "A1", col2: "B1" }),
	r2: makeRecord("r2", { col1: "A2", col2: "B2" }),
};

const columns: ComponentFramework.PropertyHelper.DataSetApi.Column[] = [
	{ name: "col1", displayName: "Col1", dataType: "SingleLine.Text", alias: "col1", order: 0, visualSizeFactor: 120, isHidden: false },
	{ name: "col2", displayName: "Col2", dataType: "SingleLine.Text", alias: "col2", order: 1, visualSizeFactor: 120, isHidden: false },
];

const resources: ComponentFramework.Resources = {
	getResource: () => { throw new Error("not impl"); },
	getString: (id: string) => {
		const m: Record<string, string> = {
			Label_Grid_Footer_RecordCount: "{0} records ({1} Selected)",
			Label_Grid_Footer: "Page {0}",
			Label_ShowFullScreen: "Show Full Screen",
			Label_NoRecords: "No records found",
			Label_SortAZ: "A to Z",
			Label_SortZA: "Z to A",
		};
		return m[id] ?? id;
	},
};

// ── Find cell selection divs (outer wrapper div with both onMouseDown and onMouseEnter) ──
function findCellDivs(root: HTMLElement): HTMLDivElement[] {
	const result: HTMLDivElement[] = [];
	function walk(node: Element): void {
		const propKey = Object.keys(node).find(
			(k) => k.startsWith("__reactProps$") || k.startsWith("__reactEventHandlers$")
		);
		if (propKey && node.tagName === "DIV") {
			const props = (node as unknown as Record<string, unknown>)[propKey] as Record<string, unknown>;
			if (typeof props?.onMouseDown === "function" && typeof props?.onMouseEnter === "function") {
				result.push(node as HTMLDivElement);
			}
		}
		for (const child of Array.from(node.children)) {
			walk(child);
		}
	}
	walk(root);
	return result;
}

function assert(condition: boolean, msg: string): void {
	if (!condition) throw new Error(`FAIL: ${msg}`);
}

// With the overlay approach the border/background styles are on a sibling overlay div.
// Walk backwards through siblings to find the one with a blue background-color.
function overlayOf(cell: HTMLDivElement | undefined): HTMLElement | null {
	if (!cell) return null;
	let sib = cell.previousElementSibling;
	while (sib instanceof HTMLElement) {
		if (sib.style.backgroundColor !== "") return sib;
		sib = sib.previousElementSibling;
	}
	return null;
}

// ── Run ───────────────────────────────────────────────────────────────────────
const container = document.createElement("div");
document.body.appendChild(container);

async function run(): Promise<void> {
	// Render the grid.
	await act(async () => {
		ReactDOM.render(
			React.createElement(Grid, {
				width: 900,
				height: 500,
				columns,
				records,
				sortedRecordIds: ["r1", "r2"],
				hasNextPage: false,
				hasPreviousPage: false,
				totalResultCount: 2,
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
				onSelectAllRecords: async () => undefined,
				onSort: () => undefined,
				onFilter: () => undefined,
				loadFirstPage: () => undefined,
				loadNextPage: () => undefined,
				loadPreviousPage: () => undefined,
				onUpdateCell: async () => undefined,
			}),
			container
		);
	});

	// ── Test 1: click single cell → blue outline ──────────────────────────────
	console.log("\n── Test 1: click cell (0,0) ──");
	let cells = findCellDivs(container);
	console.log(`  Found ${cells.length} cell divs`);
	assert(cells.length >= 4, `Expected ≥4 cell divs (2 rows × 2 cols), got ${cells.length}`);

	// cells[0] = row0/col0, cells[1] = row0/col1, cells[2] = row1/col0, cells[3] = row1/col1
	await act(async () => {
		Simulate.mouseDown(cells[0], { shiftKey: false, ctrlKey: false, button: 0 });
	});

	cells = findCellDivs(container);
	console.log(`  After click, cell[0] style:`, overlayOf(cells[0])?.style?.cssText);
	assert(
		overlayOf(cells[0])?.style?.borderTop?.includes("2px"),
		`Expected 2px top border on selected cell, got: "${overlayOf(cells[0])?.style?.borderTop}"`
	);
	assert(
		overlayOf(cells[0])?.style?.backgroundColor !== "",
		`Expected background colour on selected cell`
	);
	// Non-selected cells should have no border
	assert(
		!overlayOf(cells[1])?.style?.borderTop?.includes("2px"),
		`cell[1] should not have 2px border, got: "${overlayOf(cells[1])?.style?.borderTop}"`
	);
	console.log("  PASS: cell[0] has blue 2px outline, others clean");

	// ── Test 2: mouseenter while dragging → extends selection ─────────────────
	console.log("\n── Test 2: drag to cell (1,1) ──");
	// isDraggingRef was set to true during the mousedown above.
	await act(async () => {
		Simulate.mouseEnter(cells[3]); // row1/col1
	});

	cells = findCellDivs(container);
	console.log(`  After drag, cell[0] style:`, overlayOf(cells[0])?.style?.cssText);
	console.log(`  After drag, cell[3] style:`, overlayOf(cells[3])?.style?.cssText);

	// All 4 cells should now be in the selection range.
	for (let i = 0; i < 4; i++) {
		assert(
			overlayOf(cells[i])?.style?.backgroundColor !== "",
			`cell[${i}] should have background in range selection`
		);
	}
	// Outer-edge borders: top of row0, bottom of row1, left of col0, right of col1.
	assert(overlayOf(cells[0])?.style?.borderTop?.includes("2px"),   "cell[0] top edge should be 2px");
	assert(overlayOf(cells[0])?.style?.borderLeft?.includes("2px"),  "cell[0] left edge should be 2px");
	assert(overlayOf(cells[1])?.style?.borderTop?.includes("2px"),   "cell[1] top edge should be 2px");
	assert(overlayOf(cells[1])?.style?.borderRight?.includes("2px"), "cell[1] right edge should be 2px");
	assert(overlayOf(cells[2])?.style?.borderBottom?.includes("2px"),"cell[2] bottom edge should be 2px");
	assert(overlayOf(cells[2])?.style?.borderLeft?.includes("2px"),  "cell[2] left edge should be 2px");
	assert(overlayOf(cells[3])?.style?.borderBottom?.includes("2px"),"cell[3] bottom edge should be 2px");
	assert(overlayOf(cells[3])?.style?.borderRight?.includes("2px"), "cell[3] right edge should be 2px");
	console.log("  PASS: all 4 cells in selection range with correct edge borders");

	// End drag.
	await act(async () => {
		document.dispatchEvent(new (dom.window.MouseEvent)("mouseup"));
	});

	// ── Test 3: shift+click from (1,1) back to (0,0) ─────────────────────────
	console.log("\n── Test 3: shift+click cell (0,0) to extend ──");
	// First click cell (1,1) to set anchor.
	await act(async () => {
		Simulate.mouseDown(cells[3], { shiftKey: false, button: 0 });
	});
	cells = findCellDivs(container);
	// Now shift+click cell (0,0) to extend.
	await act(async () => {
		Simulate.mouseDown(cells[0], { shiftKey: true, button: 0 });
	});
	cells = findCellDivs(container);
	for (let i = 0; i < 4; i++) {
		assert(
			overlayOf(cells[i])?.style?.backgroundColor !== "",
			`cell[${i}] should be selected after shift+click range, got empty bg`
		);
	}
	console.log("  PASS: shift+click extends selection across 2×2 range");

	// ── Test 4: Ctrl+C copies as TSV ─────────────────────────────────────────
	console.log("\n── Test 4: Ctrl+C copies TSV ──");
	clipboardText = "";
	await act(async () => {
		const ev = new (dom.window.KeyboardEvent)("keydown", { key: "c", ctrlKey: true, bubbles: true });
		document.dispatchEvent(ev);
		// clipboard.writeText is async; wait a tick
		await Promise.resolve();
	});
	console.log(`  Clipboard text: "${clipboardText}"`);
	// Selection is 2 rows × 2 cols → 4 values.
	// Row order: r1=[A1, B1], r2=[A2, B2]  (or r2/r1 depending on anchor/head ordering)
	const lines = clipboardText.split("\r\n");
	assert(lines.length === 2, `Expected 2 TSV rows, got ${lines.length}: "${clipboardText}"`);
	const cells0 = lines[0].split("\t");
	assert(cells0.length === 2, `Expected 2 tab-separated values in row 0, got ${cells0.length}`);
	console.log("  PASS: Ctrl+C produced 2×2 TSV");
	Object.defineProperty(navigator, 'clipboard', {value: {writeText: async () => {throw new Error('Clipboard denied');}}, configurable:true});
	await act(async()=>{document.dispatchEvent(new dom.window.KeyboardEvent('keydown',{key:'c',ctrlKey:true,bubbles:true})); await Promise.resolve();});
	assert(container.querySelector('[role="alert"]')?.textContent?.includes('Unable to copy') === true, 'Denied clipboard access must show a useful error.');
	Object.defineProperty(navigator, 'clipboard', {value:undefined, configurable:true});
	await act(async()=>{document.dispatchEvent(new dom.window.KeyboardEvent('keydown',{key:'c',ctrlKey:true,bubbles:true})); await Promise.resolve();});
	assert(container.querySelector('[role="alert"]')?.textContent?.includes('unavailable') === true, 'Unavailable clipboard access must not fail silently.');

	console.log("\n✓ All cell-selection tests passed.");
}

run()
	.then(() => {
		ReactDOM.unmountComponentAtNode(container);
		process.exit(0);
	})
	.catch((err: unknown) => {
		console.error("\n✗ Cell-selection test failed:", err);
		ReactDOM.unmountComponentAtNode(container);
		process.exit(1);
	});
