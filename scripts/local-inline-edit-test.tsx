import { JSDOM } from "jsdom";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { act, Simulate } from "react-dom/test-utils";
import * as Utilities from "@fluentui/react/lib/Utilities";
import * as UtilitiesCommonJs from "@fluentui/react/lib-commonjs/Utilities";
import { List } from "@fluentui/react/lib/List";
import { Grid } from "../CMModelDrivenGrid/Grid";

const dom = new JSDOM("<!doctype html><html><body></body></html>", {
	url: "http://localhost/",
});

// Minimal browser-like globals for React + Fluent UI rendering in Node.
globalThis.window = dom.window as unknown as Window & typeof globalThis;
globalThis.document = dom.window.document;
Object.defineProperty(globalThis, "navigator", {
	value: dom.window.navigator,
	configurable: true,
});
(globalThis as unknown as { Document: typeof Document }).Document = dom.window.Document;
(globalThis as unknown as { Node: typeof Node }).Node = dom.window.Node;
(globalThis as unknown as { HTMLElement: typeof HTMLElement }).HTMLElement = dom.window.HTMLElement;
(globalThis as unknown as { Event: typeof Event }).Event = dom.window.Event;
(globalThis as unknown as { KeyboardEvent: typeof KeyboardEvent }).KeyboardEvent = dom.window.KeyboardEvent;
(globalThis as unknown as { requestAnimationFrame: (cb: FrameRequestCallback) => number }).requestAnimationFrame = (
	cb
) => setTimeout(() => cb(0), 0) as unknown as number;

// Fluent DetailsList relies on viewport/layout APIs that are not implemented in jsdom.
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

const elementPrototype = globalThis.window.HTMLElement.prototype as HTMLElement & {
	scrollTop?: number;
	scrollLeft?: number;
	scrollHeight?: number;
	clientHeight?: number;
	clientWidth?: number;
};

if (!elementPrototype.getBoundingClientRect) {
	elementPrototype.getBoundingClientRect = () =>
		({
			x: 0,
			y: 0,
			top: 0,
			left: 0,
			right: 1000,
			bottom: 600,
			width: 1000,
			height: 600,
			toJSON: () => "",
		}) as DOMRect;
}

Object.defineProperty(globalThis.window.HTMLElement.prototype, "offsetHeight", { get: () => 32 });
Object.defineProperty(globalThis.window.HTMLElement.prototype, "offsetWidth", { get: () => 180 });
Object.defineProperty(globalThis.window.HTMLElement.prototype, "clientHeight", { get: () => 600 });
Object.defineProperty(globalThis.window.HTMLElement.prototype, "clientWidth", { get: () => 1000 });
Object.defineProperty(globalThis.window.HTMLElement.prototype, "scrollHeight", { get: () => 1200 });
Object.defineProperty(globalThis.window.HTMLElement.prototype, "scrollWidth", { get: () => 1000 });

(globalThis.window.HTMLElement.prototype as unknown as { scrollTo: (...args: unknown[]) => void }).scrollTo = () => {
	undefined;
};
(globalThis.window.HTMLElement.prototype as unknown as { attachEvent?: (...args: unknown[]) => void }).attachEvent =
	() => undefined;
(globalThis.window.HTMLElement.prototype as unknown as { detachEvent?: (...args: unknown[]) => void }).detachEvent =
	() => undefined;

if (!(globalThis as unknown as { MutationObserver?: unknown }).MutationObserver) {
	class StubMutationObserver {
		public constructor(_callback: MutationCallback) {
			undefined;
		}
		public disconnect(): void {
			undefined;
		}
		public observe(_target: Node, _options?: MutationObserverInit): void {
			undefined;
		}
		public takeRecords(): MutationRecord[] {
			return [];
		}
	}
	(globalThis as unknown as { MutationObserver: typeof StubMutationObserver }).MutationObserver =
		StubMutationObserver;
	(globalThis.window as unknown as { MutationObserver: typeof StubMutationObserver }).MutationObserver =
		StubMutationObserver;
}

// getWindow can return undefined in jsdom for some internal refs; force a safe fallback.
const originalGetWindow = Utilities.getWindow;
(Utilities as unknown as { getWindow: typeof Utilities.getWindow }).getWindow = (
element?: Element | null
) => originalGetWindow(element) ?? (globalThis.window as unknown as Window);

const originalGetWindowCommonJs = UtilitiesCommonJs.getWindow;
(UtilitiesCommonJs as unknown as { getWindow: typeof UtilitiesCommonJs.getWindow }).getWindow = (
	element?: Element | null
) => originalGetWindowCommonJs(element) ?? (globalThis.window as unknown as Window);

// Force DetailsList/List to render without virtualization for jsdom test stability.
(List as unknown as { prototype: { _shouldVirtualize: () => boolean } }).prototype._shouldVirtualize = () => false;

// Fluent SelectionZone attaches listeners via EventGroup; ignore undefined targets in jsdom.
const eventGroupPrototype = (UtilitiesCommonJs.EventGroup as unknown as {
	prototype: {
		on: (
			target: EventTarget | undefined,
			eventName: string,
			callback: (event: Event) => void,
			options?: boolean
		) => void;
	};
}).prototype;
const originalEventGroupOn = eventGroupPrototype.on;
eventGroupPrototype.on = function (
	target: EventTarget | undefined,
	eventName: string,
	callback: (event: Event) => void,
	options?: boolean
): void {
	if (!target) {
		return;
	}
	originalEventGroupOn.call(this, target, eventName, callback, options);
};

const resources: ComponentFramework.Resources = {
	getResource: () => {
		throw new Error("Not implemented for local test");
	},
	getString: (id: string) => {
		switch (id) {
			case "Label_Grid_Footer_RecordCount":
				return "{0} records ({1} Selected)";
			case "Label_Grid_Footer":
				return "Page {0}";
			case "Label_ShowFullScreen":
				return "Show Full Screen";
			case "Label_NoRecords":
				return "No records found";
			case "Label_SortAZ":
				return "A to Z";
			case "Label_SortZA":
				return "Z to A";
			default:
				return id;
		}
	},
};

const recordId = "record-1";
const record: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord = {
	getRecordId: () => recordId,
	getFormattedValue: (_columnName: string) => "12.34",
	getValue: (_columnName: string) => 12.34,
	getNamedReference: () => ({ id: { guid: recordId }, etn: "account", name: "Record 1" }),
};

const columns: ComponentFramework.PropertyHelper.DataSetApi.Column[] = [
	{
		name: "cm_decimal",
		displayName: "Decimal",
		dataType: "Decimal",
		alias: "cm_decimal",
		order: 0,
		visualSizeFactor: 120,
		isHidden: false,
	},
];

const updateCalls: Array<{ recordId: string; columnName: string; value: string; dataType: string }> = [];

const container = document.createElement("div");
document.body.appendChild(container);

async function run(): Promise<void> {
	await act(async () => {
		ReactDOM.render(
			React.createElement(Grid, {
				width: 900,
				height: 500,
				columns,
				records: { [recordId]: record },
				sortedRecordIds: [recordId],
				hasNextPage: false,
				hasPreviousPage: false,
				totalResultCount: 1,
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
				loadFirstPage: () => undefined,
				loadNextPage: () => undefined,
				loadPreviousPage: () => undefined,
				onUpdateCell: async (rid: string, col: string, value: string, dataType: string) => {
					updateCalls.push({ recordId: rid, columnName: col, value, dataType });
				},
				onFullScreen: () => undefined,
				isFullScreen: false,
			}),
			container
		);
	});

	const editableCell = container.querySelector('span[title="Click to edit"]');
	if (!editableCell) {
		throw new Error("Editable cell was not rendered.");
	}

	await act(async () => {
		const reactPropKey = Object.keys(editableCell).find(
			(key) => key.startsWith("__reactProps$") || key.startsWith("__reactEventHandlers$")
		);
		if (!reactPropKey) {
			throw new Error("Could not locate React props on editable cell.");
		}
		const reactProps = (editableCell as unknown as Record<string, unknown>)[reactPropKey] as
			| { onClick?: (event: { preventDefault: () => void; stopPropagation: () => void }) => void }
			| undefined;
		if (!reactProps?.onClick) {
			throw new Error("Editable cell is missing React onClick handler.");
		}
		try {
			reactProps.onClick({
				preventDefault: () => undefined,
				stopPropagation: () => undefined,
			});
		} catch (error) {
			console.error("Click handler error:", error);
			throw error;
		}
		await Promise.resolve();
	});

	const input = container.querySelector("input");
	if (!input) {
		throw new Error("Inline editor input was not shown after double-click.");
	}

	await act(async () => {
		(input as HTMLInputElement).value = "45.67";
		Simulate.change(input, { target: { value: "45.67" } });
	});

	await act(async () => {
		Simulate.blur(input);
		await Promise.resolve();
	});

	if (updateCalls.length !== 1) {
		throw new Error(`Expected one update call, got ${updateCalls.length}.`);
	}

	const call = updateCalls[0];
	if (call.value !== "45.67") {
		throw new Error(`Expected saved value '45.67', got '${call.value}'.`);
	}

	console.log("Local inline-edit test passed:", call);
}

run()
	.then(() => {
		ReactDOM.unmountComponentAtNode(container);
		process.exit(0);
	})
	.catch((error: unknown) => {
		console.error("Local inline-edit test failed:", error);
		ReactDOM.unmountComponentAtNode(container);
		process.exit(1);
	});
