import { useConst, useForceUpdate, useId } from "@fluentui/react-hooks";
import * as React from "react";
import { GridStateCache, useGridState } from './gridState';
import * as ReactDOM from "react-dom";
import { IObjectWithKey, IRenderFunction, SelectionMode } from "@fluentui/react/lib/Utilities";
import {
	ConstrainMode,
	CheckboxVisibility,
	ColumnActionsMode,
	DetailsList,
	DetailsListLayoutMode,
	DetailsRow,
	DetailsRowCheck,
	IColumn,
	IDetailsList,
	IDetailsHeaderProps,
	IDetailsListProps,
	IDetailsRowProps,
	IDetailsRowStyles,
} from "@fluentui/react/lib/DetailsList";
import { Sticky, StickyPositionType } from "@fluentui/react/lib/Sticky";
import { Callout, DirectionalHint } from "@fluentui/react/lib/Callout";
import { ScrollablePane, ScrollbarVisibility } from "@fluentui/react/lib/ScrollablePane";
import { Stack } from "@fluentui/react/lib/Stack";
import { Overlay } from "@fluentui/react/lib/Overlay";
import { ActionButton, DefaultButton, IconButton, PrimaryButton } from "@fluentui/react/lib/Button";
import { Link } from "@fluentui/react/lib/Link";
import { Icon } from "@fluentui/react/lib/Icon";
import { Text } from "@fluentui/react/lib/Text";
import { TextField } from "@fluentui/react/lib/TextField";
import { Checkbox } from "@fluentui/react/lib/Checkbox";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { useTheme } from "@fluentui/react";
import { getColumnConditions, getFilterColumnNames, deriveFilterRule, extractColumnFilterGroup, getFilterDefinitions, getFilterKind, FilterMode, FilterValue, FilterRuleMode } from "./filtering";
import { ColumnFilterEditor } from "./ColumnFilterEditor";
import { fillColumnWidths } from "./columnSizing";
import { serializeGridClipboard } from './clipboard';
import { GridFeatures, moveColumn, reconcileColumnOrder, numericColumnWidth, AggregateMode, aggregateNumbers, groupRecords, groupingKey, columnTypeIcon, choiceColorStyle } from './gridPresentation';
import { dateGroup, DateGroupingInterval, dateGroupingIntervals } from './dateGrouping';
import { readColumnLayout, writeColumnLayout, allowLookupNavigation } from './layoutPreferences';
import { getJumpFilter } from './filtering';
import { navigateCell } from './cellNavigation';
import { Dialog, DialogFooter } from '@fluentui/react/lib/Dialog';
import { GroupHeader, IGroup } from '@fluentui/react/lib/GroupedList';
import { createRecordSelection, MAX_SELECTED_RECORDS, SelectAllRequest, SelectAllResult, selectionResult } from './recordSelection';

type DataSet = ComponentFramework.PropertyHelper.DataSetApi.EntityRecord & IObjectWithKey;
interface EditingCell {
	recordId: string;
	columnName: string;
	dataType: string;
	value: string;
	originalValue: string;
	/** Entity type of the lookup target — set when editing a lookup field. */
	lookupEntityType?: string;
}

type CellSaveStatus = "saving" | "saved" | "failed";

type FilterByMode = FilterRuleMode;

interface FilterValueOption {
	/** Stable key used by the checkbox list. It is separate from the display label. */
	key: string;
	/** Formatted value shown to the user. */
	label: string;
	/** Raw value submitted to the dataset filter. Empty string represents a blank value. */
	value: string;
	isBlank: boolean;
}

const BLANK_FILTER_OPTION_KEY = "__cm_grid_blank__";
const GRID_CELL_STYLE_PROPS = { cellLeftPadding: 12, cellRightPadding: 12, cellExtraRightPadding: 24 };
// Keep data-column sizing in sync with the 56px checkbox/chevron gutter.
const HIERARCHY_GUTTER_EXTRA_WIDTH = 16;

const renderHierarchyCheckbox: IDetailsListProps['onRenderCheckbox'] = checkboxProps => (
	<span aria-hidden="true" className={`cm-grid-selection-check${checkboxProps?.checked ? ' is-checked' : ''}`}>
		{checkboxProps?.checked && <Icon iconName="CheckMark" />}
	</span>
);

function stringFormat(template: string, ...args: string[]): string {
	args?.forEach((arg, index) => {
		template = template.replace("{" + index + "}", arg);
	});
	return template;
}

// Measure rendered pixel width of a string using an offscreen canvas.
// Fluent DetailsList cells: ~14px Segoe UI; headers: same but with bold weight.
let _measureCanvas: HTMLCanvasElement | null = null;
function measureTextWidth(text: string, font = "14px 'Segoe UI', sans-serif"): number {
	_measureCanvas ??= document.createElement("canvas");
	const ctx = _measureCanvas.getContext("2d");
	if (!ctx) return text.length * 8;
	ctx.font = font;
	return ctx.measureText(text).width;
}

function getUniqueSortedFilterOptions(options: FilterValueOption[]): FilterValueOption[] {
	const unique = new Map<string, FilterValueOption>();
	for (const option of options) {
		if (!unique.has(option.key)) {
			unique.set(option.key, option);
		}
	}
	return Array.from(unique.values()).sort((left, right) => {
		if (left.isBlank !== right.isBlank) return left.isBlank ? 1 : -1;
		return left.label.localeCompare(right.label, undefined, { numeric: true, sensitivity: "base" });
	});
}

function toRawFilterString(rawValue: unknown, depth = 0): string | null {
	if (depth > 3 || rawValue === null || rawValue === undefined) return null;
	if (rawValue instanceof Date) {
		return Number.isNaN(rawValue.getTime()) ? null : rawValue.toISOString();
	}
	if (typeof rawValue === "string" || typeof rawValue === "number" || typeof rawValue === "boolean") {
		return String(rawValue);
	}
	if (Array.isArray(rawValue)) {
		// A single lookup reference can be array-wrapped. Multi-valued fields
		// require different Dataverse operators; do not silently use item zero.
		return rawValue.length === 1 ? toRawFilterString(rawValue[0], depth + 1) : null;
	}
	if (typeof rawValue !== "object") return null;

	const record = rawValue as Record<string, unknown>;
	for (const key of ["value", "Value", "guid", "Guid", "id", "Id", "key", "Key"]) {
		const found = toRawFilterString(record[key], depth + 1);
		if (found !== null) return found;
	}
	return null;
}

function isBlankRawFilterValue(rawValue: unknown): boolean {
	if (rawValue === null || rawValue === undefined || rawValue === "") return true;
	if (Array.isArray(rawValue)) {
		return rawValue.length === 0 || rawValue.every((value) => value === null || value === undefined);
	}
	return false;
}

function createFilterValueOption(label: string, rawValue: unknown): FilterValueOption | null {
	const rawText = toRawFilterString(rawValue);
	const isBlank = isBlankRawFilterValue(rawValue);
	if (!isBlank && rawText === null) {
		// Do not turn an unsupported object/array value into a false "blank" filter.
		return null;
	}
	const value = isBlank ? "" : rawText ?? label;
	const formattedLabel = label.trim() || value;
	const displayLabel = isBlank ? "(Blanks)" : formattedLabel === "(Blanks)" ? "(Blanks) — text value" : formattedLabel;
	return {
		key: isBlank ? BLANK_FILTER_OPTION_KEY : `value:${value}`,
		label: displayLabel,
		value,
		isBlank,
	};
}

export interface GridProps {
	features?: GridFeatures;
	/** User/org/table/view-scoped presentation only. Omit to disable browser persistence. */
	layoutKey?: string;
	dataError?: string;
	stateCache?: GridStateCache;
	selectedRecordIds?: string[];
	onCollapseChild?: (id: string) => void;
	renderChildRows?: (record: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord, expanded: boolean) => React.ReactNode;
	hasChildRows?: (record: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord) => boolean | undefined;
	renderRowActions?: (record: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord) => React.ReactNode;
	disableSelection?: boolean;
	/** Only one level of a hierarchy owns the host command-bar selection. */
	selectionActive?: boolean;
	width?: number;
	height?: number;
	columns: ComponentFramework.PropertyHelper.DataSetApi.Column[];
	records: Record<string, ComponentFramework.PropertyHelper.DataSetApi.EntityRecord>;
	sortedRecordIds: string[];
	hasNextPage: boolean;
	hasPreviousPage: boolean;
	totalResultCount: number;
	currentPage: number;
	sorting: ComponentFramework.PropertyHelper.DataSetApi.SortStatus[];
	filtering: ComponentFramework.PropertyHelper.DataSetApi.FilterExpression;
	resources: ComponentFramework.Resources;
	itemsLoading: boolean;
	highlightValue: string | null;
	highlightColor: string | null;
	enableLookupLinks: boolean;
	enableInlineEdit: boolean;
	columnOptions?: Record<string, { label: string; value: string; color?: string }[]>;
	setSelectedRecords: (ids: string[]) => void;
	onNavigate: (item?: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord) => void;
	onOpenLookup: (entityType: string, id: string) => void;
	onSort: (name: string, direction: "asc" | "desc" | "none", additive?: boolean) => void;
	onFilter: (name: string, mode: FilterMode, value?: FilterValue) => void;
	onClearAllFilters: () => void;
	loadFirstPage: () => void;
	loadNextPage: () => void | Promise<void>;
	loadPreviousPage: () => void;
	onUpdateCell: (recordId: string, columnName: string, value: string, dataType: string) => Promise<void> | void;
	onUpdateLookupCell?: (recordId: string, columnName: string, targetId: string, targetEntityType: string) => Promise<void>;
	lookupTargetEntityMap?: Record<string, string>;
	/** Load eligible IDs in current filter/sort order; Grid publishes only the completed result. */
	onSelectAllRecords: (request?: SelectAllRequest) => Promise<SelectAllResult | undefined> | SelectAllResult | undefined;
	selectionCountOverride?: number;
	selectionResetVersion?: number;
	commandStatus?: string;
	commandError?: string;
	/** An error-only host update can finish a query without changing its records. */
	queryErrorVersion?: number;
	onSearchLookup?: (
		entityType: string,
		searchTerm: string,
		columnName?: string,
		recordId?: string
	) => Promise<{ id: string; name: string }[]>;
	onBrowseLookup?: (entityType: string, searchTerm: string, columnName: string, recordId: string) => Promise<{id: string; name: string} | undefined>;
	infiniteScroll?: boolean;
	item?: DataSet;
}

// Self-contained editing input — manages its own value state so typing always
// works regardless of whether DetailsList re-renders the parent row.
interface EditingTextFieldProps {
	inputType?: 'text' | 'date';
	initialValue?: string | null;
	ariaLabel?: string;
	onCommit: (value: string) => void;
	onCancel: () => void;
}

const EditingTextField: React.FC<EditingTextFieldProps> = ({ initialValue, ariaLabel, inputType, onCommit, onCancel }) => {
	const normalizedInitialValue = initialValue ?? "";
	const [value, setValue] = React.useState(normalizedInitialValue);
	const valueRef = React.useRef(normalizedInitialValue);
	const committedRef = React.useRef(false);

	const handleChange = (_ev: React.FormEvent, newValue?: string) => {
		const v = newValue ?? "";
		setValue(v);
		valueRef.current = v;
	};

	const handleCommit = () => {
		if (committedRef.current) return;
		committedRef.current = true;
		onCommit(valueRef.current);
	};

	const handleKeyDown = (event: React.KeyboardEvent<HTMLInputElement | HTMLTextAreaElement>) => {
		if (event.key === "Enter") {
			event.preventDefault();
			handleCommit();
		} else if (event.key === "Escape") {
			event.preventDefault();
			committedRef.current = true;
			onCancel();
		}
	};

	return (
		<TextField
			type={inputType ?? 'text'}
			value={value}
			autoFocus
			ariaLabel={ariaLabel}
			borderless
			styles={{ root: { width: "100%" }, fieldGroup: { minHeight: 24 } }}
			onChange={handleChange}
			onBlur={handleCommit}
			onKeyDown={handleKeyDown}
		/>
	);
};

// Self-contained dropdown for choice (OptionSet / TwoOptions) fields.
interface EditingDropdownProps {
	initialValue?: string | null;
	options: { label: string; value: string }[];
	ariaLabel?: string;
	onCommit: (value: string) => void;
	onCancel: () => void;
}

const EditingDropdown: React.FC<EditingDropdownProps> = ({ initialValue, options, ariaLabel, onCommit, onCancel }) => {
	const committedRef = React.useRef(false);
	const changedRef = React.useRef(false);
	const normalizedInitialValue = initialValue ?? "";
	const dropdownOptions = React.useMemo(
		() => (options.some((option) => option.value === "") ? options : [{ label: "(No value)", value: "" }, ...options]),
		[options]
	);

	const handleChange = (event: React.ChangeEvent<HTMLSelectElement>) => {
		if (committedRef.current) return;
		changedRef.current = event.target.value !== normalizedInitialValue;
		if (!changedRef.current) return;
		committedRef.current = true;
		onCommit(event.target.value);
	};

	const handleKeyDown = (event: React.KeyboardEvent<HTMLSelectElement>) => {
		if (event.key === "Escape") {
			event.preventDefault();
			committedRef.current = true;
			onCancel();
		}
	};

	const handleBlur = (event: React.FocusEvent<HTMLSelectElement>) => {
		if (committedRef.current) return;
		// A dropdown receives a blur when the row/list rerenders or when focus moves
		// away. Do not turn an untouched empty/default selection into a write.
		if (!changedRef.current || event.target.value === normalizedInitialValue) {
			committedRef.current = true;
			onCancel();
			return;
		}
		committedRef.current = true;
		onCommit(event.target.value);
	};

	return (
		<select
			autoFocus
			defaultValue={normalizedInitialValue}
			aria-label={ariaLabel}
			onChange={handleChange}
			onKeyDown={handleKeyDown}
			onBlur={handleBlur}
			style={{ width: "100%", minHeight: 24, fontSize: "inherit", border: "none", outline: "none" }}
		>
			{dropdownOptions.map((opt) => (
				<option key={opt.value} value={opt.value}>
					{opt.label}
				</option>
			))}
		</select>
	);
};

// Searchable lookup picker — shows a text input and a Callout dropdown of results.
interface EditingLookupProps {
	initialDisplayValue?: string | null;
	ariaLabel?: string;
	onSearch: (term: string) => Promise<{ id: string; name: string }[]>;
	onBrowse?: (term: string) => Promise<{id: string; name: string} | undefined>;
	onCommit: (id: string, name: string) => void;
	onCancel: () => void;
}

const EditingLookup: React.FC<EditingLookupProps> = ({ initialDisplayValue, ariaLabel, onSearch, onBrowse, onCommit, onCancel }) => {
	const normalizedInitialDisplayValue = initialDisplayValue ?? "";
	const [searchTerm, setSearchTerm] = React.useState(normalizedInitialDisplayValue);
	const [results, setResults] = React.useState<{ id: string; name: string }[]>([]);
	const [isSearching, setIsSearching] = React.useState(false);
	const [searchError, setSearchError] = React.useState<string | null>(null);
	const [dropdownRect, setDropdownRect] = React.useState<DOMRect | null>(null);
	const committedRef = React.useRef(false);
	const debounceRef = React.useRef<number | undefined>(undefined);
	const blurRef = React.useRef<number | undefined>(undefined);
	const searchRequestRef = React.useRef(0);
	const isMountedRef = React.useRef(true);
	const containerRef = React.useRef<HTMLDivElement>(null);
	const browsingRef = React.useRef(false);
	const [browsing, setBrowsing] = React.useState(false);
	const [activeResult, setActiveResult] = React.useState(-1);
	const resultListId = useId('cm-lookup-results-');
	React.useEffect(() => {if (!browsing && !committedRef.current) containerRef.current?.querySelector('input')?.focus();}, [browsing]);

	const performSearch = React.useCallback(
		async (term: string) => {
			const requestId = ++searchRequestRef.current;
			setIsSearching(true);
			setSearchError(null);
			try {
				const found = await onSearch(term);
				if (!isMountedRef.current || requestId !== searchRequestRef.current) return;
				setResults(found);
				setActiveResult(-1);
			} catch (err) {
				if (!isMountedRef.current || requestId !== searchRequestRef.current) return;
				setResults([]);
				setSearchError(err instanceof Error ? err.message : "Search failed");
			} finally {
				if (isMountedRef.current && requestId === searchRequestRef.current) {
					setIsSearching(false);
				}
			}
		},
		[onSearch]
	);

	// Capture position and kick off initial search on mount.
	React.useEffect(() => {
		const mountedStateRef = isMountedRef;
		const requestSequenceRef = searchRequestRef;
		const debounceTimerRef = debounceRef;
		const blurTimerRef = blurRef;
		if (containerRef.current) {
			setDropdownRect(containerRef.current.getBoundingClientRect());
		}
		void performSearch(normalizedInitialDisplayValue);
		return () => {
			mountedStateRef.current = false;
			requestSequenceRef.current++;
			if (debounceTimerRef.current !== undefined) {
				window.clearTimeout(debounceTimerRef.current);
			}
			if (blurTimerRef.current !== undefined) {
				window.clearTimeout(blurTimerRef.current);
			}
		};
	// eslint-disable-next-line react-hooks/exhaustive-deps
	}, []);

	const handleChange = (_ev: React.FormEvent, newValue?: string) => {
		const term = newValue ?? "";
		// Do not let an old result remain selectable while a new query is debouncing.
		searchRequestRef.current++; setResults([]); setActiveResult(-1); setIsSearching(true);
		setSearchTerm(term);
		if (debounceRef.current !== undefined) {
			window.clearTimeout(debounceRef.current);
		}
		debounceRef.current = window.setTimeout(() => {
			void performSearch(term);
		}, 300);
	};

	const handleKeyDown = (event: React.KeyboardEvent) => {
		if (browsingRef.current) return;
		if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
			event.preventDefault(); event.stopPropagation();
			setActiveResult(current => Math.max(-1, Math.min(results.length - 1, current + (event.key === 'ArrowDown' ? 1 : -1))));
			return;
		}
		if (event.key === 'Enter' && activeResult >= 0 && results[activeResult]) {
			event.preventDefault(); event.stopPropagation(); handleSelect(results[activeResult].id, results[activeResult].name); return;
		}
		if (event.key === "Escape") {
			event.preventDefault();
			event.stopPropagation();
			committedRef.current = true;
			onCancel();
		} else if (event.key === "Tab") {
			event.preventDefault();
			event.stopPropagation();
			if (activeResult >= 0 && results[activeResult]) {
				handleSelect(results[activeResult].id, results[activeResult].name);
			} else if (searchTerm.trim() === "" && normalizedInitialDisplayValue !== '') {
				handleSelect("", ""); // commit null
			} else {
				committedRef.current = true;
				onCancel();
			}
		}
	};

	// Delay cancel so a mousedown on a result item fires before we cancel.
	const handleBlur = () => {
		if (browsingRef.current) return;
		const termAtBlur = searchTerm;
		blurRef.current = window.setTimeout(() => {
			blurRef.current = undefined;
			if (!committedRef.current && !browsingRef.current && isMountedRef.current) {
				committedRef.current = true;
				if (termAtBlur.trim() === "" && normalizedInitialDisplayValue !== '') {
					onCommit("", ""); // empty field → save as null
				} else {
					onCancel(); // non-empty but nothing selected → revert
				}
			}
		}, 200);
	};

	const handleSelect = (id: string, name: string) => {
		if (committedRef.current) return;
		committedRef.current = true;
		onCommit(id, name);
	};
	const handleBrowse = async () => {
		if (!onBrowse || browsingRef.current || committedRef.current) return;
		browsingRef.current = true; setBrowsing(true); setSearchError(null);
		if (blurRef.current !== undefined) window.clearTimeout(blurRef.current);
		try {
			const selected = await onBrowse(searchTerm);
			if (isMountedRef.current && selected) handleSelect(selected.id, selected.name);
		} catch (error) {
			if (isMountedRef.current) setSearchError(error instanceof Error ? error.message : 'Unable to open record lookup.');
		} finally {
			browsingRef.current = false;
			if (isMountedRef.current) {
				setBrowsing(false);
				if (!committedRef.current) containerRef.current?.querySelector('input')?.focus();
			}
		}
	};

	// Always render the dropdown once we have a position — shows spinner, error, results, or "no results".
	const dropdownPortal =
		dropdownRect !== null && !browsing &&
		ReactDOM.createPortal(
			<div
				data-selection-disabled={true}
				onClick={event => event.stopPropagation()}
				onMouseDown={event => event.stopPropagation()}
				onDoubleClick={event => event.stopPropagation()}
				style={{
					position: "fixed",
					top: dropdownRect.bottom,
					left: dropdownRect.left,
					minWidth: Math.max(dropdownRect.width, 200),
					maxHeight: 200,
					overflowY: "auto",
					background: "white",
					border: "1px solid #ccc",
					boxShadow: "0 4px 8px rgba(0,0,0,0.16)",
					zIndex: 1000000,
					borderRadius: 2,
				}}
			>
				{/* Clear button moved inline next to the text field */}
				{isSearching && (
					<Stack horizontal verticalAlign="center" tokens={{ childrenGap: 6 }} style={{ padding: "6px 8px" }}>
						<Spinner size={SpinnerSize.xSmall} />
						<Text variant="small">Searching...</Text>
					</Stack>
				)}
				{!isSearching && searchError && (
					<Text variant="small" style={{ padding: "6px 8px", display: "block", color: "red" }}>
						{searchError}
					</Text>
				)}
				{!isSearching && !searchError && results.length === 0 && (
					<Text variant="small" style={{ padding: "6px 8px", display: "block", color: "#666" }}>
						No results found
					</Text>
				)}
				<div id={resultListId} role="listbox" aria-label="Lookup results">{!isSearching &&
					!searchError &&
					results.map((r, index) => (
						<ActionButton
							key={r.id}
							id={`${resultListId}-${index}`}
							role="option" aria-selected={activeResult === index} tabIndex={-1}
							styles={{ root: { width: "100%", textAlign: "left", height: 28, paddingLeft: 8, background: activeResult === index ? '#deecf9' : undefined } }}
							aria-pressed={activeResult === index}
							onMouseDown={(e) => e.preventDefault()}
							onClick={() => handleSelect(r.id, r.name)}
						>
							{r.name}
						</ActionButton>
					))}</div>
				{onBrowse && <ActionButton text="More records…" iconProps={{iconName: 'Search'}} onMouseDown={event => event.preventDefault()} onClick={() => void handleBrowse()} />}
			</div>,
			document.body
		);

	return (
		<>
			<div ref={containerRef} style={{ display: "flex", alignItems: "center" }}>
				<TextField
					autoFocus
					ariaLabel={ariaLabel}
					role="combobox" aria-autocomplete="list" aria-expanded={!browsing} aria-controls={resultListId}
					aria-activedescendant={activeResult >= 0 ? `${resultListId}-${activeResult}` : undefined}
					borderless
					value={searchTerm}
					disabled={browsing}
					styles={{ root: { flex: 1, minWidth: 0 }, fieldGroup: { minHeight: 24 } }}
					onChange={handleChange}
					onKeyDown={handleKeyDown}
					onBlur={handleBlur}
				/>			</div>
			{browsing && <Spinner size={SpinnerSize.xSmall} label="Record lookup is open" />}
			{dropdownPortal}
		</>
	);
};

const onRenderDetailsHeader: IRenderFunction<IDetailsHeaderProps> = (props, defaultRender) => {
	if (props && defaultRender) {
		return (
			<Sticky stickyPosition={StickyPositionType.Header} isScrollSynced>
				{defaultRender({
					...props,
				})}
			</Sticky>
		);
	}
	return null;
};

export const Grid = React.memo((props: GridProps) => {
	const features = props.features ?? {};
	const initialLayout = useConst(() => readColumnLayout(props.layoutKey));
	const [columnOrder, setColumnOrder] = useGridState<string[]>(props.stateCache, 'columnOrder', () => initialLayout?.order ?? []);
	const [widthEditor, setWidthEditor] = React.useState<{key: string; label: string; value: string} | undefined>();
	const [widthError, setWidthError] = React.useState('');
	const hierarchySelectionGutter = !!props.renderChildRows && !props.disableSelection;
	const [expandedRows, setExpandedRows] = useGridState(props.stateCache, 'expanded', () => new Set<string>());
	const [visitedChildRows, setVisitedChildRows] = useGridState(props.stateCache, 'visited', () => new Set<string>());
	const previouslyExpanded = React.useRef(expandedRows);
	const onCollapseChild = props.onCollapseChild;
	React.useEffect(() => {
		for (const id of previouslyExpanded.current) if (!expandedRows.has(id)) onCollapseChild?.(id);
		previouslyExpanded.current = expandedRows;
	}, [expandedRows, onCollapseChild]);
	const hasChildRows = props.hasChildRows;
	const hierarchyRecords = props.records;
	React.useEffect(() => {
		if (!hasChildRows) return;
		setExpandedRows(current => {
			const removed = [...current].filter(id => hierarchyRecords[id] && hasChildRows(hierarchyRecords[id]) === false);
			if (!removed.length) return current;
			const next = new Set(current); removed.forEach(id => next.delete(id)); return next;
		});
	}, [hasChildRows,hierarchyRecords,setExpandedRows]);
	const {
		records,
		sortedRecordIds,
		columns,
		width,
		height,
		hasNextPage,
		hasPreviousPage,
		sorting,
		filtering,
		currentPage,
		itemsLoading,
		setSelectedRecords,
		onNavigate,
		onSort,
		onFilter,
		onClearAllFilters,
		resources,
		loadFirstPage,
		loadNextPage,
		loadPreviousPage,
		highlightValue,
		highlightColor,
		enableLookupLinks,
		enableInlineEdit,
		totalResultCount,
		onOpenLookup,
		onUpdateCell,
		onUpdateLookupCell,
		lookupTargetEntityMap,
		onSelectAllRecords,
		onSearchLookup,
		onBrowseLookup,
		columnOptions,
		infiniteScroll,
	} = props;
	const theme = useTheme();
	const forceUpdate = useForceUpdate();
	const onSelectAllRecordsRef = React.useRef(onSelectAllRecords);
	const setSelectedRecordsRef = React.useRef(setSelectedRecords);
	onSelectAllRecordsRef.current = onSelectAllRecords;
	setSelectedRecordsRef.current = setSelectedRecords;
	const selectionIds = React.useRef<string[]>(props.selectedRecordIds ?? []);
	const bulkRequest = React.useRef<AbortController>();
	const [selectingAll, setSelectingAll] = React.useState(false);
	const [selectionMessage, setSelectionMessage] = React.useState('');
	const [selectionError, setSelectionError] = React.useState('');
	const selectionCallbacks = React.useRef<{changed: (ids: string[]) => void; toggleAll: () => void}>({changed: () => undefined, toggleAll: () => undefined});
	const selection = useConst(() => createRecordSelection({
		ids: () => selectionIds.current,
		changed: ids => selectionCallbacks.current.changed(ids),
		toggleAll: () => selectionCallbacks.current.toggleAll(),
	}));
	const publishSelection = (ids: string[]): void => {
		selectionIds.current = ids;
		selection.syncIds(ids);
		setSelectedRecordsRef.current(ids);
		forceUpdate();
	};
	selectionCallbacks.current.changed = ids => {
		bulkRequest.current?.abort(); bulkRequest.current = undefined; setSelectingAll(false);
		const result = selectionResult(features.enableMultipleSelection === false ? ids.slice(-1) : ids);
		setSelectionError('');
		setSelectionMessage(result.limitReached ? 'You can select up to 1,000 records at a time.' : '');
		publishSelection(result.ids);
	};
	selectionCallbacks.current.toggleAll = () => {
		if (bulkRequest.current || selectionIds.current.length) { selectionCallbacks.current.changed([]); return; }
		const controller = new AbortController();
		bulkRequest.current = controller; setSelectingAll(true); setSelectionError(''); setSelectionMessage('');
		void (async () => {
			try {
				const loaded = await onSelectAllRecordsRef.current({signal: controller.signal});
				if (controller.signal.aborted) return;
				if (!loaded && hasNextPage) throw new Error('Unable to select records beyond the loaded page. Please try again.');
				const result = loaded ?? selectionResult(sortedRecordIds);
				const bounded = selectionResult(result.ids, result.limitReached);
				publishSelection(bounded.ids);
				setSelectionMessage(bounded.limitReached ? 'Selected the first 1,000 matching records (Dynamics limit). Narrow the filters to select a different set.' : '');
			} catch (error) {
				if (!controller.signal.aborted) { publishSelection([]); setSelectionError(error instanceof Error ? error.message : 'Unable to select records.'); }
			} finally {
				if (bulkRequest.current === controller) { bulkRequest.current = undefined; setSelectingAll(false); }
			}
		})();
	};
	React.useEffect(() => () => {bulkRequest.current?.abort(); bulkRequest.current = undefined;}, []);
	React.useLayoutEffect(() => {
		if (props.selectionActive !== false) return;
		bulkRequest.current?.abort(); bulkRequest.current = undefined; setSelectingAll(false);
		selectionIds.current = []; selection.syncIds([]); setSelectionMessage(''); setSelectionError('');
		forceUpdate();
	}, [props.selectionActive, selection, forceUpdate]);
	const previousSelectionReset = React.useRef(props.selectionResetVersion);
	React.useLayoutEffect(() => {
		if (previousSelectionReset.current === props.selectionResetVersion) return;
		previousSelectionReset.current = props.selectionResetVersion;
		bulkRequest.current?.abort(); bulkRequest.current = undefined; setSelectingAll(false);
		selectionIds.current = []; selection.syncIds([]); setSelectionMessage(''); setSelectionError('');
		forceUpdate();
	}, [props.selectionResetVersion, selection, forceUpdate]);
	const selectionQuery = JSON.stringify([filtering, sorting]);
	const previousSelectionQuery = React.useRef(selectionQuery);
	React.useLayoutEffect(() => {
		if (previousSelectionQuery.current === selectionQuery) return;
		previousSelectionQuery.current = selectionQuery;
		if (bulkRequest.current || selectionIds.current.length) selectionCallbacks.current.changed([]);
	}, [selectionQuery]);

	const [isComponentLoading, setIsLoading] = React.useState<boolean>(false);
	const [filterColumn, setFilterColumn] = React.useState<IColumn | undefined>();
	const [filterTarget, setFilterTarget] = React.useState<HTMLElement | undefined>();
	const [allFilterOptions, setAllFilterOptions] = React.useState<FilterValueOption[]>([]);
	const [checkedFilterOptionKeys, setCheckedFilterOptionKeys] = React.useState<string[]>([]);
	const [actionError, setActionError] = React.useState<string>();
	const [clipboardError, setClipboardError] = React.useState('');
	const [groupColumn, setGroupColumn] = useGridState<string | undefined>(props.stateCache, 'groupColumn', () => undefined);
	const [aggregateColumns, setAggregateColumns] = useGridState<Record<string, AggregateMode>>(props.stateCache, 'aggregates', () => ({}));
	const [groupCollapse, setGroupCollapse] = useGridState<Record<string, boolean>>(props.stateCache, 'groupCollapse', () => ({}));
	const [dateGrouping, setDateGrouping] = useGridState<DateGroupingInterval>(props.stateCache, 'dateGrouping', () => 'day');
	const analysisRequested = !!groupColumn || Object.keys(aggregateColumns).length > 0;
	const analysisReady = !hasNextPage && !itemsLoading && !isComponentLoading && !props.dataError && !props.commandError;
	const [analysisPaused, setAnalysisPaused] = React.useState(false);
	const [searchText, setSearchText] = React.useState<string>("");
	const [filterView, setFilterView] = React.useState<"main" | "filterBy">("main");
	const appliedValueFilterOptionsRef = React.useRef<Map<string, FilterValueOption[]>>(new Map());
	// Preserve the last value-picker options used for each applied column filter so
	// editing an active filter chip can reopen with the same choices.

	// Keep column widths stable across paging/filtering; resize explicitly.
	const containerRef = React.useRef<HTMLDivElement>(null);
	const listColumnMenuRef = React.useRef<HTMLDivElement>(null);
	const [columnWidths, setColumnWidths] = useGridState(props.stateCache, 'widths', () => initialLayout?.widths ?? ({} as Record<string, number>));
	const [layoutMessage, setLayoutMessage] = React.useState('');
	React.useEffect(() => {
		setLayoutMessage(writeColumnLayout(props.layoutKey, {order: columnOrder, widths: columnWidths}) ? '' : 'Column layout will last for this session. Browser storage is unavailable.');
	}, [props.layoutKey, columnOrder, columnWidths]);
	const [availableColumnWidth, setAvailableColumnWidth] = React.useState(0);
	const listMode = features.reflow === 'list' || features.reflow === 'auto' && availableColumnWidth > 0 && availableColumnWidth < 480;
	const visibleColumns = React.useMemo(() => {
		const available = columns.filter(column => !column.isHidden && (typeof column.order !== 'number' || column.order >= 0))
			.sort((a,b) => (a.order ?? 0) - (b.order ?? 0));
		const order = reconcileColumnOrder(available.map(column => column.name), columnOrder);
		return available.sort((a,b) => order.indexOf(a.name) - order.indexOf(b.name));
	}, [columns, columnOrder]);
	const renderedColumnCount = listMode ? Math.min(visibleColumns.length, Math.max(1, Math.min(10, features.listColumns ?? 3))) : visibleColumns.length;
	const rangeLayoutKey = JSON.stringify([listMode, visibleColumns.slice(0, renderedColumnCount).map(column => column.name)]);
	const detailsListRef = React.useRef<IDetailsList>(null);
	React.useEffect(() => {
		let frame: number | undefined;
		// Fluent List watches its own scroll port. An expanded list also moves
		// when an ancestor scrolls, so refresh its cached viewport on re-entry.
		const ancestorScrolled = (event: Event): void => {
			const root = containerRef.current;
			const target = event.target;
			if (!root || (target !== document && target !== window && !(target instanceof window.Element && target.contains(root)))) return;
			if (frame !== undefined) return;
			frame = window.requestAnimationFrame(() => {
				frame = undefined;
				const rect = containerRef.current?.getBoundingClientRect();
				if (rect && rect.height > 0 && rect.bottom > 0 && rect.top < window.innerHeight) detailsListRef.current?.forceUpdate();
			});
		};
		window.addEventListener('scroll', ancestorScrolled, true);
		return () => {
			window.removeEventListener('scroll', ancestorScrolled, true);
			if (frame !== undefined) window.cancelAnimationFrame(frame);
		};
	}, []);
	const syncingColumnWidthsRef = React.useRef(false);
	const appliedColumnWidthsRef = React.useRef(new Map<string, number>());
	React.useLayoutEffect(() => {
		const scrollPort = containerRef.current?.querySelector<HTMLElement>(".ms-ScrollablePane--contentContainer");
		if (!scrollPort) return;
		const measure = (): void => {
			const nextWidth = scrollPort.clientWidth;
			setAvailableColumnWidth((current) => current === nextWidth ? current : nextWidth);
		};
		measure();
		const observer = typeof ResizeObserver !== "undefined" ? new ResizeObserver(measure) : undefined;
		observer?.observe(scrollPort);
		window.addEventListener("resize", measure);
		return () => { observer?.disconnect(); window.removeEventListener("resize", measure); };
	}, [width, height]);
	const queryIdentity = JSON.stringify({ sorting, filtering });
	React.useEffect(() => {
		const port = containerRef.current?.querySelector<HTMLElement>(".ms-ScrollablePane--contentContainer");
		if (!port) return;
		const cached = props.stateCache?.get('scroll') as {query: string; top: number} | undefined;
		port.scrollTo({top: cached?.query === queryIdentity ? cached.top : 0});
		const remember = () => props.stateCache?.set('scroll', {query: queryIdentity, top: port.scrollTop});
		port.addEventListener('scroll', remember);
		return () => { port.removeEventListener('scroll', remember); };
	}, [queryIdentity, props.stateCache]);

	// Infinite scroll refs.
	const sentinelRef = React.useRef<HTMLDivElement>(null);
	const isLoadingNextPageRef = React.useRef(false);
	const lastInfiniteLoadIdentityRef = React.useRef<string | null>(null);
	const observedInfiniteLoadIdentityRef = React.useRef<string | null>(null);
	const [infiniteScrollError, setInfiniteScrollError] = React.useState<string | null>(null);
	const [pagingInProgress, setPagingInProgress] = React.useState(false);
	const pageRequestSerial = React.useRef(0);
	// PCF can render the new page before loadNextPage's promise/finally settles.
	// Keep this separate from the rendered-row identity to prevent reusing the
	// controller's still-active promise as if it were a request for another page.
	const pagePromisePendingRef = React.useRef(false);
	const [pagePromisePending, setPagePromisePending] = React.useState(false);
	React.useEffect(() => () => { pageRequestSerial.current++; }, []);

	// Multi-cell editing state — each key is `${recordId}-${columnName}`.
	const [editingCells, setEditingCells] = useGridState(props.stateCache, 'editing', () => new Map<string, EditingCell>());
	const editingCellsRef = React.useRef<Map<string, EditingCell>>(new Map());
	editingCellsRef.current = editingCells;
	const editingAllowedRef = React.useRef(enableInlineEdit);
	editingAllowedRef.current = enableInlineEdit;
	React.useLayoutEffect(() => {
		if (!enableInlineEdit && editingCellsRef.current.size) {
			editingCellsRef.current = new Map();
			setEditingCells(new Map());
		}
	}, [enableInlineEdit, setEditingCells]);

	// Per-cell save status — same key format as editingCells.
	const [cellSaveStatuses, setCellSaveStatuses] = React.useState<Map<string, CellSaveStatus>>(() => new Map());
	const cellSaveTimeoutsRef = React.useRef<Map<string, number>>(new Map());

	// Cleanup all pending timeouts on unmount.
	React.useEffect(() => {
		// eslint-disable-next-line react-hooks/exhaustive-deps
		return () => { cellSaveTimeoutsRef.current.forEach((id) => window.clearTimeout(id)); };
	}, []);

	// ── Excel-like cell range selection ──────────────────────────────────
	interface CellCoord { row: number; col: number; }
	interface CellSelection { anchor: CellCoord; head: CellCoord; }
	const [cellSelection, setCellSelection] = React.useState<CellSelection | null>(null);
	const cellSelectionRef = React.useRef<CellSelection | null>(null);
	const isDraggingRef = React.useRef(false);
	const cellMouseDownRef = React.useRef(false);
	const rangeColumnsRef = React.useRef<IColumn[]>([]);
	const displayItemsRef = React.useRef<DataSet[]>([]);
	const visibleRangeRowsRef = React.useRef<Set<number>>();
	const [keyboardTarget, setKeyboardTarget] = React.useState<{head: CellCoord; anchor: CellCoord; query: string; layout: string} | undefined>();
	const previousRangeOrderRef = React.useRef<{ rows: string[]; layout: string }>();
	React.useEffect(() => {
		if (features.enableRangeSelection !== false) return;
		cellSelectionRef.current = null;
		isDraggingRef.current = false;
		setCellSelection(null);
	}, [features.enableRangeSelection]);
	React.useEffect(() => {
		if (features.enableGrouping === false) setGroupColumn(undefined);
		if (features.enableAggregation === false) setAggregateColumns(current => Object.keys(current).length ? {} : current);
	}, [features.enableGrouping, features.enableAggregation, setGroupColumn, setAggregateColumns]);

	const setCellTransientStatus = React.useCallback((cellKey: string, status: "saved" | "failed", timeoutMs: number) => {
		const existing = cellSaveTimeoutsRef.current.get(cellKey);
		if (existing !== undefined) {
			window.clearTimeout(existing);
		}
		setCellSaveStatuses((prev) => {
			const next = new Map(prev);
			next.set(cellKey, status);
			return next;
		});
		const id = window.setTimeout(() => {
			cellSaveTimeoutsRef.current.delete(cellKey);
			setCellSaveStatuses((prev) => {
				const next = new Map(prev);
				next.delete(cellKey);
				return next;
			});
		}, timeoutMs);
		cellSaveTimeoutsRef.current.set(cellKey, id);
	}, []);

	const isInlineEditableDataType = React.useCallback((dataType: string): boolean => {
		const normalizedType = dataType.toLowerCase().replace(/[\s_-]+/g, "");
		if (!normalizedType) return false;
		const unsupportedMarkers = [
			"customer",
			"partylist",
			"regarding",
			"image",
			"file",
			"multiselectoptionset",
			"multiselect",
			"primarykey",
			"uniqueidentifier",
			"state",
			"status",
			"managedproperty",
			"virtual",
		];
		if (unsupportedMarkers.some((marker) => normalizedType.includes(marker))) return false;

		return [
			"fp",
			"singleline",
			"multiple",
			"memo",
			"text",
			"decimal",
			"currency",
			"money",
			"whole",
			"number",
			"double",
			"float",
			"integer",
			"int",
			"date",
			"datetime",
			"dateandtime",
			"twooptions",
			"boolean",
			"optionset",
			"lookup",
			"owner",
		].some((marker) => normalizedType.includes(marker));
	}, []);

	const isLookupDataType = React.useCallback((dataType: string): boolean => {
		const n = dataType.toLowerCase();
		return n.includes("lookup") || n.includes("owner");
	}, []);

	const getLookupReference = React.useCallback(
		(
			item: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord,
			fieldName: string
		): { entityType: string; id: string } | null => {
			try {
				const rawValue = item.getValue(fieldName);
				const candidate = Array.isArray(rawValue) ? rawValue[0] : rawValue;
				if (!candidate || typeof candidate !== "object") {
					return null;
				}

				if ("entityType" in candidate && "id" in candidate) {
					const lookupCandidate = candidate as { entityType?: unknown; id?: unknown };
					const entityType = typeof lookupCandidate.entityType === "string" ? lookupCandidate.entityType : "";
					const id = typeof lookupCandidate.id === "string" ? lookupCandidate.id : "";
					return entityType && id ? { entityType, id } : null;
				}

				if ("etn" in candidate && "id" in candidate) {
					const entityReferenceCandidate = candidate as { etn?: unknown; id?: { guid?: unknown } };
					const entityType = typeof entityReferenceCandidate.etn === "string" ? entityReferenceCandidate.etn : "";
					const id = typeof entityReferenceCandidate.id?.guid === "string" ? entityReferenceCandidate.id.guid : "";
					return entityType && id ? { entityType, id } : null;
				}

				return null;
			} catch {
				return null;
			}
		},
		[]
	);

	const beginEditCell = React.useCallback(
		(
			item: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord,
			column: IColumn,
			dataType: string,
			formattedValue: string
		) => {
			const cellKey = `${item.getRecordId()}-${column.key}`;
			if (editingCellsRef.current.has(cellKey)) return;

			const initialRawValue = item.getValue(column.key);
			let initialValue = formattedValue;
			if (typeof initialRawValue === "string") {
				initialValue = initialRawValue;
			} else if (typeof initialRawValue === "number" || typeof initialRawValue === "boolean") {
				initialValue = initialRawValue.toString();
			} else if (initialRawValue instanceof Date) {
				initialValue = dataType.toLowerCase().includes("dateonly")
					? initialRawValue.toISOString().slice(0, 10)
					: initialRawValue.toISOString();
			}

			// For lookup fields, capture the target entity type from the existing reference.
			let lookupEntityType: string | undefined;
			if (isLookupDataType(dataType)) {
				const ref = getLookupReference(item, column.key);
				lookupEntityType = ref?.entityType;
				// Fallback: extract etn directly from opaque lookup objects where id is not a string.
				if (!lookupEntityType) {
					const rawVal = item.getValue(column.key);
					const cand = Array.isArray(rawVal) ? rawVal[0] : rawVal;
					if (cand && typeof cand === "object" && "etn" in cand) {
						const etnVal = (cand as { etn?: unknown }).etn;
						if (typeof etnVal === "string" && etnVal) lookupEntityType = etnVal;
					}
				}
				// If the cell is empty (no existing ref), fall back to the column's declared targets.
				if (!lookupEntityType) {
					const pcfCol = column.data as { targets?: string[] } | undefined;
					lookupEntityType = pcfCol?.targets?.[0];
				}
				// Last resort: scan loaded rows for any non-null ref on the same column.
				if (!lookupEntityType) {
					for (const otherItem of displayItemsRef.current) {
						const otherRef = getLookupReference(otherItem, column.key);
						if (otherRef?.entityType) {
							lookupEntityType = otherRef.entityType;
							break;
						}
						// Also try opaque etn extraction for other rows.
						const otherRaw = otherItem.getValue(column.key);
						const otherCand = Array.isArray(otherRaw) ? otherRaw[0] : otherRaw;
						if (otherCand && typeof otherCand === "object" && "etn" in otherCand) {
							const otherEtn = (otherCand as { etn?: unknown }).etn;
							if (typeof otherEtn === "string" && otherEtn) {
								lookupEntityType = otherEtn;
								break;
							}
						}
					}
				}
				// Final fallback: use the metadata-driven lookup target entity map.
				if (!lookupEntityType && lookupTargetEntityMap) {
					lookupEntityType = lookupTargetEntityMap[column.key.toLowerCase()];
				}
				// Always use the formatted display value for lookups (the raw value is an object).
				initialValue = formattedValue;
			}

			setEditingCells((prev) => {
				const next = new Map(prev);
				next.set(cellKey, {
					recordId: item.getRecordId(),
					columnName: column.key,
					dataType,
					value: initialValue,
					originalValue: initialValue,
					lookupEntityType,
				});
				return next;
			});
		},
		[getLookupReference, isLookupDataType, lookupTargetEntityMap, setEditingCells]
	);

	const cancelEditCell = React.useCallback((cellKey: string) => {
		setEditingCells((prev) => {
			if (!prev.has(cellKey)) return prev;
			const next = new Map(prev);
			next.delete(cellKey);
			return next;
		});
	}, [setEditingCells]);

	const commitEditCell = React.useCallback(
		async (cellKey: string, finalValue: string) => {
			if (!editingAllowedRef.current) return;
			const current = editingCellsRef.current.get(cellKey);
			if (!current) return;
			// Lookup fields must be saved via commitLookupCell — never via plain text update.
			if (isLookupDataType(current.dataType)) return;

			setEditingCells((prev) => {
				const next = new Map(prev);
				next.delete(cellKey);
				return next;
			});

			if (finalValue === current.originalValue) return;

			const existing = cellSaveTimeoutsRef.current.get(cellKey);
			if (existing !== undefined) {
				window.clearTimeout(existing);
				cellSaveTimeoutsRef.current.delete(cellKey);
			}

			setCellSaveStatuses((prev) => {
				const next = new Map(prev);
				next.set(cellKey, "saving");
				return next;
			});

			try {
				await onUpdateCell(current.recordId, current.columnName, finalValue, current.dataType);
				setCellTransientStatus(cellKey, "saved", 2000);
			} catch {
				setCellTransientStatus(cellKey, "failed", 4000);
			}
		},
		[isLookupDataType, onUpdateCell, setCellTransientStatus, setEditingCells]
	);

	const commitLookupCell = React.useCallback(
		async (cellKey: string, targetId: string, targetEntityType: string) => {
			if (!editingAllowedRef.current) return;
			const current = editingCellsRef.current.get(cellKey);
			if (!current || !onUpdateLookupCell) return;

			setEditingCells((prev) => {
				const next = new Map(prev);
				next.delete(cellKey);
				return next;
			});

			const existing = cellSaveTimeoutsRef.current.get(cellKey);
			if (existing !== undefined) {
				window.clearTimeout(existing);
				cellSaveTimeoutsRef.current.delete(cellKey);
			}

			setCellSaveStatuses((prev) => {
				const next = new Map(prev);
				next.set(cellKey, "saving");
				return next;
			});

			try {
				await onUpdateLookupCell(current.recordId, current.columnName, targetId, targetEntityType);
				setCellTransientStatus(cellKey, "saved", 2000);
			} catch {
				setCellTransientStatus(cellKey, "failed", 4000);
			}
		},
		[onUpdateLookupCell, setCellTransientStatus, setEditingCells]
	);

	const items: DataSet[] = React.useMemo(() => {
		const sortedRecords: DataSet[] = sortedRecordIds
			.map((id) => {
				const record = records[id];
				return record as DataSet | undefined;
			})
			.filter((record): record is DataSet => record !== undefined);

		if (sortedRecords.length > 0) {
			return sortedRecords;
		}

		return Object.values(records).filter((record): record is DataSet => record !== undefined);
	}, [records, sortedRecordIds]);

	const grouped = React.useMemo(() => {
		const column = columns.find(column => column.name === groupColumn);
		if (!groupColumn || !analysisReady || !column) return undefined;
		if (getFilterKind(column.dataType) === 'date') {
			const dates = new Map(items.map(item => [item, dateGroup(item.getValue(groupColumn), dateGrouping, {dateOnly: column.dataType.toLowerCase().includes('dateonly'), firstDayOfWeek: features.firstDayOfWeek, locale: features.locale, offsetMinutes: features.dateTimeOffsetMinutes})]));
			const groupFor = (item: DataSet) => dates.get(item)!;
			return groupRecords(items, item => groupFor(item).key, item => groupFor(item).label, item => groupFor(item).sortKey);
		}
		return groupRecords(items, item => groupingKey(toRawFilterString(item.getValue(groupColumn))), item => item.getFormattedValue(groupColumn));
	}, [items, groupColumn, analysisReady, columns, dateGrouping, features.firstDayOfWeek, features.locale, features.dateTimeOffsetMinutes]);
	const orderedItems = grouped?.items ?? items;
	const groups: IGroup[] | undefined = React.useMemo(() => grouped?.groups.map(group => ({...group, isCollapsed: groupCollapse[group.key] ?? true})), [grouped, groupCollapse]);
	const visibleRangeRows = React.useMemo(() => groups ? groups.filter(group => !group.isCollapsed).flatMap(group => Array.from({length: group.count}, (_, index) => group.startIndex + index)) : undefined, [groups]);
	visibleRangeRowsRef.current = visibleRangeRows ? new Set(visibleRangeRows) : undefined;
	React.useLayoutEffect(() => {
		const next = { rows: orderedItems.map((item) => item.getRecordId()), layout: rangeLayoutKey };
		const previous = previousRangeOrderRef.current;
		previousRangeOrderRef.current = next;
		if (!previous) return;
		const sameColumns = previous.layout === next.layout;
		const existingRowsUnmoved = previous.rows.length <= next.rows.length && previous.rows.every((id, index) => id === next.rows[index]);
		if (!sameColumns || !existingRowsUnmoved) {
			setKeyboardTarget(undefined);
			cellSelectionRef.current = null;
			isDraggingRef.current = false;
			setCellSelection(null);
		}
	}, [orderedItems, rangeLayoutKey]);

	// cellSelection is included so the List component (which only re-renders pages when `items`
	// changes reference) picks up the updated onRenderItemColumn selection styles on every click/drag.
	// eslint-disable-next-line react-hooks/exhaustive-deps
	const displayItems = React.useMemo(() => [...orderedItems], [orderedItems, editingCells, cellSaveStatuses, cellSelection, expandedRows, visitedChildRows, props.renderRowActions, props.renderChildRows, props.hasChildRows, enableInlineEdit, enableLookupLinks, features, columnOptions, lookupTargetEntityMap, onBrowseLookup]);
	displayItemsRef.current = displayItems;
	const infiniteLoadIdentity = React.useMemo(() => {
		const ids = displayItems.map((item) => item.getRecordId()).join("|");
		return `${hasNextPage ? "more" : "end"}:${ids}`;
	}, [displayItems, hasNextPage]);
	const latestInfiniteIdentity = React.useRef(infiniteLoadIdentity);
	latestInfiniteIdentity.current = infiniteLoadIdentity;
	const latestPagingBusy = React.useRef(itemsLoading || isComponentLoading);
	latestPagingBusy.current = itemsLoading || isComponentLoading;
	React.useLayoutEffect(() => {
		if (!props.selectedRecordIds || props.selectionActive === false) return;
		selectionIds.current = props.selectedRecordIds;
		selection.syncIds(props.selectedRecordIds);
	}, [props.selectedRecordIds, props.selectionActive, selection, displayItems]);

	// Clear local loading state whenever the PCF dataset finishes loading OR records update.
	// This prevents the overlay getting permanently stuck if the filter result is identical
	// to the current data (records reference unchanged, items useMemo never re-runs).
	const prevItemsLoadingRef = React.useRef(false);
	React.useEffect(() => {
		if (!props.commandError && !props.dataError) return;
		setIsLoading(false);
		setPagingInProgress(false);
	}, [props.commandError, props.dataError, props.queryErrorVersion]);
	React.useEffect(() => {
		const wasLoading = prevItemsLoadingRef.current;
		prevItemsLoadingRef.current = itemsLoading;
		if (wasLoading && !itemsLoading) {
			setIsLoading(false);
			if (!pagePromisePendingRef.current) setPagingInProgress(false);
		}
	}, [itemsLoading]);

	// Also clear local loading when records change (belt-and-braces for the fast-path case
	// where PCF never sets dataset.loading=true because it responds before the next render).
	React.useEffect(() => {
		if (itemsLoading || pagePromisePendingRef.current) return;
		setIsLoading(false);
		setPagingInProgress(false);
		if (!itemsLoading && isLoadingNextPageRef.current && lastInfiniteLoadIdentityRef.current === infiniteLoadIdentity) {
			isLoadingNextPageRef.current = false;
			setInfiniteScrollError("No additional records were returned. Try again.");
		}
	}, [records, sortedRecordIds, itemsLoading, infiniteLoadIdentity]);

	// The host can finish a paging request without appending records (for
	// example when the page is empty or the dataset reports the same page
	// twice). Release only the in-flight flag once both loading indicators have
	// settled; the identity guard below still prevents an automatic retry for
	// the same dataset state, while the footer Retry action can recover it.
	React.useEffect(() => {
		if (!itemsLoading && !isComponentLoading && !pagePromisePendingRef.current) {
			if (isLoadingNextPageRef.current && hasNextPage && lastInfiniteLoadIdentityRef.current === infiniteLoadIdentity) {
				setInfiniteScrollError("No additional records were returned. Try again.");
			}
			isLoadingNextPageRef.current = false;
		}
	}, [isComponentLoading, itemsLoading, hasNextPage, infiniteLoadIdentity]);

	// A list rerender can move the sentinel without loading a new page. Clear the
	// request guard only when the dataset identity actually changes, otherwise a
	// visible sentinel can issue duplicate loadNextPage calls.
	React.useEffect(() => {
		if (observedInfiniteLoadIdentityRef.current === infiniteLoadIdentity) return;
		observedInfiniteLoadIdentityRef.current = infiniteLoadIdentity;
		lastInfiniteLoadIdentityRef.current = null;
		if (!pagePromisePendingRef.current) isLoadingNextPageRef.current = false;
		setInfiniteScrollError(null);
	}, [infiniteLoadIdentity]);

	const requestNextPage = React.useCallback((force = false) => {
		if (!infiniteScroll || !hasNextPage || itemsLoading || isComponentLoading || bulkRequest.current) return;
		if (pagePromisePending || pagePromisePendingRef.current) return;
		if (infiniteScrollError && !force) return;
		// A manual Retry is allowed to recover even if the host left the local
		// guard set after a failed/silent page request. Automatic observer
		// callbacks remain blocked while the error is visible.
		if (force) {
			isLoadingNextPageRef.current = false;
			lastInfiniteLoadIdentityRef.current = null;
		}
		if (isLoadingNextPageRef.current) return;
		if (!force && lastInfiniteLoadIdentityRef.current === infiniteLoadIdentity) return;

		lastInfiniteLoadIdentityRef.current = infiniteLoadIdentity;
		isLoadingNextPageRef.current = true;
		setInfiniteScrollError(null);
		setPagingInProgress(true);
		const serial = ++pageRequestSerial.current;
		pagePromisePendingRef.current = true;
		setPagePromisePending(true);
		const handleLoadError = () => {
			if (serial !== pageRequestSerial.current) return;
			pagePromisePendingRef.current = false;
			setPagePromisePending(false);
			lastInfiniteLoadIdentityRef.current = null;
			isLoadingNextPageRef.current = false;
			setIsLoading(false);
			setPagingInProgress(false);
			setInfiniteScrollError("Unable to load more records. Try again.");
		};

		try {
			const result = loadNextPage() as unknown;
			if (result && typeof (result as { then?: unknown }).then === "function") {
				void Promise.resolve(result as PromiseLike<unknown>).then(() => {
					if (serial !== pageRequestSerial.current) return;
					pagePromisePendingRef.current = false;
					isLoadingNextPageRef.current = false;
					setPagingInProgress(false);
					if (latestPagingBusy.current) lastInfiniteLoadIdentityRef.current = null;
					else if (latestInfiniteIdentity.current === infiniteLoadIdentity) setInfiniteScrollError("No additional records were returned. Try again.");
					// Re-arm prefetch only after the host promise has fully settled.
					setPagePromisePending(false);
					return undefined;
				}, handleLoadError);
			} else {
				pagePromisePendingRef.current = false;
				setPagePromisePending(false);
			}
		} catch {
			handleLoadError();
		}
	}, [hasNextPage, infiniteLoadIdentity, infiniteScroll, infiniteScrollError, isComponentLoading, itemsLoading, loadNextPage, pagePromisePending]);
	React.useEffect(() => {
		if (!keyboardTarget) return;
		if (keyboardTarget.query !== selectionQuery || keyboardTarget.layout !== rangeLayoutKey || features.enableRangeSelection === false || infiniteScrollError) {setKeyboardTarget(undefined); return;}
		if (keyboardTarget.head.row >= orderedItems.length) {
			if (hasNextPage && !groups) requestNextPage(); else setKeyboardTarget(undefined);
			return;
		}
		const next = {anchor: keyboardTarget.anchor, head: keyboardTarget.head};
		cellSelectionRef.current = next; setCellSelection(next);
		let frame: number | undefined;
		let attempts = 0;
		const focus = () => {
			const root = containerRef.current;
			const candidates = root?.querySelectorAll<HTMLElement>(`[data-cm-cell-row="${next.head.row}"][data-cm-cell-col="${next.head.col}"]`);
			const cell = candidates && Array.from(candidates).find(cell => cell.closest('.cm-model-driven-grid') === root);
			if (cell) {cell.focus({preventScroll: true}); cell.scrollIntoView?.({block: 'nearest', inline: 'nearest'}); setKeyboardTarget(undefined);}
			else if (++attempts <= 30) {
				if (attempts === 1) detailsListRef.current?.focusIndex(next.head.row);
				frame = window.requestAnimationFrame(focus);
			}
			else {setKeyboardTarget(undefined); cellSelectionRef.current = null; setCellSelection(null);}
		};
		focus();
		return () => {if (frame !== undefined) window.cancelAnimationFrame(frame);};
	}, [keyboardTarget, selectionQuery, rangeLayoutKey, features.enableRangeSelection, orderedItems, hasNextPage, groups, requestNextPage, infiniteScrollError]);
	// Exact summaries need the whole filtered level. Page through the bound dataset,
	// retaining its form relationship and security scope; never aggregate a base view alone.
	React.useEffect(() => {
		if (analysisRequested && !analysisPaused && hasNextPage && !itemsLoading && !isComponentLoading && !props.dataError) requestNextPage();
	}, [analysisRequested, analysisPaused, hasNextPage, itemsLoading, isComponentLoading, requestNextPage, infiniteLoadIdentity, props.dataError]);

	// Prefetch within the grid's scroll viewport before the last row is reached.
	React.useEffect(() => {
		if (!infiniteScroll || !sentinelRef.current) return;
		const sentinel = sentinelRef.current;
		const port = containerRef.current?.querySelector<HTMLElement>(".ms-ScrollablePane--contentContainer");
		// A nested grid's sentinel may be clipped by the parent's viewport even
		// at the end of its own scrollbar. Use the actual scroll distance too.
		const checkDistance = () => {
			if (port && port.clientHeight > 0 && port.scrollHeight - port.scrollTop - port.clientHeight <= 160) requestNextPage();
		};
		port?.addEventListener('scroll', checkDistance, {passive:true});
		const observer = typeof IntersectionObserver === 'undefined' ? undefined : new IntersectionObserver(
			([entry]) => {
				if (entry.isIntersecting) requestNextPage();
			},
			{ root: port ?? null, rootMargin: "0px 0px 160px 0px", threshold: 0 }
		);
		observer?.observe(sentinel);
		checkDistance();
		return () => {observer?.disconnect();port?.removeEventListener('scroll',checkDistance);};
	}, [infiniteScroll, requestNextPage]);

	// Global mouse-up ends drag; Ctrl+C copies selected range as TSV.
	React.useEffect(() => {
		const onMouseUp = () => { isDraggingRef.current = false; };
		const onMouseDownOutside = () => {
			if (cellMouseDownRef.current) {
				cellMouseDownRef.current = false;
				return;
			}
			// Don't disturb displayItems while a cell is being edited — re-rendering
			// the List would remount the active editor (killing in-progress searches).
			if (editingCellsRef.current.size > 0) return;
			// Click was outside any cell – clear the selection.
			cellSelectionRef.current = null;
			setCellSelection(null);
		};
		const onKeyDown = (e: KeyboardEvent) => {
			const owningGrid = (e.target as HTMLElement | null)?.closest?.('.cm-model-driven-grid');
			if (owningGrid && owningGrid !== containerRef.current) return;
			if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'c') {
				const target = e.target as HTMLElement | null;
				const targetTag = target?.tagName?.toLowerCase();
				if (
					editingCellsRef.current.size > 0 ||
					targetTag === "input" ||
					targetTag === "textarea" ||
					targetTag === "select" ||
					target?.isContentEditable
				) {
					return;
				}
				const sel = cellSelectionRef.current;
				if (!sel) return;
				const { anchor, head } = sel;
				const minRow = Math.min(anchor.row, head.row);
				const maxRow = Math.max(anchor.row, head.row);
				const minCol = Math.min(anchor.col, head.col);
				const maxCol = Math.max(anchor.col, head.col);
				const cols = rangeColumnsRef.current.slice(minCol, maxCol + 1);
				const rows = displayItemsRef.current.slice(minRow, maxRow + 1).filter((_, index) => !visibleRangeRowsRef.current || visibleRangeRowsRef.current.has(minRow + index));
				const tsv = serializeGridClipboard(rows
					.map((row) =>
						cols
							.map((col) => {
								try { return row.getFormattedValue(col.fieldName ?? col.key) ?? ''; }
								catch { return ''; }
							})
					)
				);
				setClipboardError('');
				if (navigator.clipboard?.writeText) {
					e.preventDefault();
					void navigator.clipboard.writeText(tsv).catch(() => setClipboardError('Unable to copy. Check clipboard access and try again.'));
				} else {
					setClipboardError('Clipboard access is unavailable in this browser.');
				}
			}
		};
		document.addEventListener('mouseup', onMouseUp);
		document.addEventListener('keydown', onKeyDown);
		document.addEventListener('mousedown', onMouseDownOutside);
		return () => {
			document.removeEventListener('mouseup', onMouseUp);
			document.removeEventListener('keydown', onKeyDown);
			document.removeEventListener('mousedown', onMouseDownOutside);
		};
	}, []);

	const deriveFilterByState = React.useCallback((condition?: ComponentFramework.PropertyHelper.DataSetApi.ConditionExpression) => {
		const rule = condition ? deriveFilterRule(condition) : undefined;
		return { mode: rule?.mode ?? "equals" as FilterByMode, value: rule && Array.isArray(rule.value) ? rule.value.join(" – ") : rule?.value ?? "" };
	}, []);
	const editorFilterGroup = React.useMemo(() => {
		if (!filterColumn) return undefined;
		const group = extractColumnFilterGroup(filtering, filterColumn.key);
		if (group) return group;
		const conditions = getColumnConditions(filtering, filterColumn.key);
		const rule = conditions.length === 1 ? deriveFilterRule(conditions[0]) : undefined;
		return rule ? { operator: "and" as const, rules: [rule] } : undefined;
	}, [filtering, filterColumn]);

	const activeFilters = React.useMemo(() => {
		const filteredColumnNames = getFilterColumnNames(filtering);
		if (filteredColumnNames.length === 0) return [];

		const seenColumns = new Set<string>();
		return filteredColumnNames.flatMap((attributeName) => {
			const col = columns.find((c) => c.name === attributeName);
			const columnKey = col?.name ?? attributeName;
			if (seenColumns.has(columnKey)) return [];
			seenColumns.add(columnKey);

			const conditions = getColumnConditions(filtering, attributeName);
			if (conditions.length === 0) return [];
			const persistedOptions = appliedValueFilterOptionsRef.current.get(columnKey) ?? [];
			const labelForRawValue = (rawValue: unknown): string => {
				const rawText = toRawFilterString(rawValue) ?? "";
				return persistedOptions.find((option) => option.value === rawText)?.label ?? (rawText || "(Blanks)");
			};
			const modeShort: Partial<Record<FilterByMode, string>> = {
				equals: "=", notEquals: "≠",
				contains: "contains",
				beginsWith: "starts",
				endsWith: "ends",
			};
			const hasBlankCondition = conditions.some((condition) => condition.conditionOperator === 12);
			const valueLabels = conditions.flatMap((condition) => {
				if (condition.conditionOperator === 12 || condition.value === undefined) return [];
				return (Array.isArray(condition.value) ? condition.value : [condition.value]).map(labelForRawValue);
			});
			const uniqueValueLabels = Array.from(new Set(valueLabels));
			let valueLabel: string;
			if (hasBlankCondition && uniqueValueLabels.length === 0) {
				valueLabel = "is blank";
			} else if (hasBlankCondition) {
				const nonBlankLabel = uniqueValueLabels.length <= 3 ? uniqueValueLabels.join(", ") : `${uniqueValueLabels.length} values`;
				valueLabel = `is blank or ${nonBlankLabel}`;
			} else if (uniqueValueLabels.length > 1 || conditions.some((condition) => Array.isArray(condition.value))) {
				valueLabel = uniqueValueLabels.length <= 3 ? uniqueValueLabels.join(", ") : `${uniqueValueLabels.length} values`;
			} else {
				const condition = conditions[0];
				const { mode, value } = deriveFilterByState(condition);
				valueLabel = `${modeShort[mode] ?? mode} "${labelForRawValue(value)}"`;
			}
			const group = extractColumnFilterGroup(filtering, attributeName);
			if (group) {
				const definitions = getFilterDefinitions(col?.dataType ?? "");
				valueLabel = group.rules.map((rule) => {
					const definition = definitions.find((entry) => entry.key === rule.mode);
					const label = definition?.text ?? rule.mode;
					if (definition?.input === "none") return label;
					const values = Array.isArray(rule.value) ? rule.value : [rule.value ?? ""];
					return `${label} ${values.map(labelForRawValue).join(" – ")}`;
				}).join(group.operator === "or" ? " OR " : " AND ");
			}

			return [{ columnKey, displayName: col?.displayName ?? attributeName, valueLabel }];
		});
	}, [columns, deriveFilterByState, filtering]);

	const onFilterDismiss = React.useCallback(() => {
		setFilterColumn(undefined);
		setFilterTarget(undefined);
		setFilterView("main");
		setAllFilterOptions([]);
		setCheckedFilterOptionKeys([]);
		setSearchText("");
	}, []);

	const persistValueFilterOptions = React.useCallback((columnKey: string, options: FilterValueOption[]) => {
		const normalizedOptions = getUniqueSortedFilterOptions(options).map((option) => ({ ...option }));
		if (normalizedOptions.length === 0) {
			appliedValueFilterOptionsRef.current.delete(columnKey);
			return;
		}

		appliedValueFilterOptionsRef.current.set(columnKey, normalizedOptions);
	}, []);

	const clearPersistedValueFilterOptions = React.useCallback((columnKey?: string) => {
		if (columnKey) {
			appliedValueFilterOptionsRef.current.delete(columnKey);
			return;
		}

		appliedValueFilterOptionsRef.current.clear();
	}, []);

	const openFilterCallout = React.useCallback(
		(column: IColumn, target: HTMLElement, preservedOptions?: FilterValueOption[]) => {
			const postFilterOptions = getUniqueSortedFilterOptions(
				(items ?? [])
					.filter((item): item is DataSet => item !== undefined)
					.map((item) => {
						let rawValue: unknown;
						let rawValueRead = true;
						try {
							rawValue = item.getValue(column.key);
						} catch {
							rawValueRead = false;
							rawValue = undefined;
						}
						const formattedValue = item.getFormattedValue(column.key) ?? "";
						if ((!rawValueRead || rawValue === undefined) && formattedValue.trim().length > 0) {
							return null;
						}
						return createFilterValueOption(formattedValue, rawValue);
					})
					.filter((option): option is FilterValueOption => option !== null)
			);

			const existingConditions = getColumnConditions(filtering, column.key);
			const existingCondition = existingConditions.find((condition) => condition.conditionOperator !== 12) ?? existingConditions[0];
			const inCondition = existingConditions.find(
				(condition) => condition.conditionOperator === 8 && Array.isArray(condition.value)
			);
			let initialCheckedOptions = [...postFilterOptions];

			const hasBlankCondition = existingConditions.some((condition) => condition.conditionOperator === 12);
			const nonBlankValues = existingConditions.flatMap((condition) => {
				if (condition.conditionOperator === 12 || condition.value === undefined) return [];
				return Array.isArray(condition.value) ? condition.value : [condition.value];
			});
			if (inCondition || (hasBlankCondition && nonBlankValues.length > 0)) {
				// "in" or mixed blank/value conditions — restore every stored value.
				const restoredValues = [
					...(hasBlankCondition ? [""] : []),
					...nonBlankValues.map((value) => (typeof value === "string" ? value : String(value ?? ""))),
				];
				initialCheckedOptions = Array.from(new Set(restoredValues)).map((value) => {
					const rawText = typeof value === "string" ? value : String(value ?? "");
					const matchingOption = postFilterOptions.find((option) => option.value === rawText);
					return matchingOption ?? ({
						key: rawText.length === 0 ? BLANK_FILTER_OPTION_KEY : `value:${rawText}`,
						label: rawText || "(Blanks)",
						value: rawText,
						isBlank: rawText.length === 0,
					} satisfies FilterValueOption);
				});
			} else if (hasBlankCondition) {
				// doesNotContainData — only blanks are selected.
				initialCheckedOptions = [{
					key: BLANK_FILTER_OPTION_KEY,
					label: "(Blanks)",
					value: "",
					isBlank: true,
				}];
			} else if (
				(existingCondition?.conditionOperator === 0) &&
				typeof existingCondition.value === "string"
			) {
				const rawText = existingCondition.value;
				const matchingOption = postFilterOptions.find((option) => option.value === rawText);
				initialCheckedOptions = [matchingOption ?? ({
					key: rawText.length === 0 ? BLANK_FILTER_OPTION_KEY : `value:${rawText}`,
					label: rawText || "(Blanks)",
					value: rawText,
					isBlank: rawText.length === 0,
				} satisfies FilterValueOption)];
			}

			const availableOptions = getUniqueSortedFilterOptions([
				...(preservedOptions ?? appliedValueFilterOptionsRef.current.get(column.key) ?? []),
				...postFilterOptions,
				...initialCheckedOptions,
			]);
			const availableKeys = new Set(availableOptions.map((option) => option.key));

			setActionError(undefined);
			setFilterColumn(column);
			setFilterTarget(target);
			setFilterView("main");
			setAllFilterOptions(availableOptions);
			setCheckedFilterOptionKeys(initialCheckedOptions.filter((option) => availableKeys.has(option.key)).map((option) => option.key));
			setSearchText("");
		},
			[filtering, items]
		);


	const filteredValueOptions = React.useMemo(() => {
		if (!searchText.trim()) {
			return allFilterOptions;
		}

		const loweredSearch = searchText.toLocaleLowerCase();
		return allFilterOptions.filter((option) => option.label.toLocaleLowerCase().includes(loweredSearch));
	}, [allFilterOptions, searchText]);

	const allVisibleSelected =
		filteredValueOptions.length > 0 &&
		filteredValueOptions.every((option) => checkedFilterOptionKeys.includes(option.key));

	const menuRowStyles = React.useMemo(
		() => ({
			root: {
				width: "100%",
				justifyContent: "flex-start",
				textAlign: "left",
				minHeight: 36,
				height: 36,
				paddingLeft: 12,
				paddingRight: 12,
				paddingTop: 0,
				paddingBottom: 0,
				borderRadius: 0,
				fontWeight: "400",
			},
			label: {
				fontWeight: "400",
				fontSize: 14,
				lineHeight: "36px",
			},
			rootHovered: {
				backgroundColor: theme.palette.neutralLighter,
			},
		}),
		[theme.palette.neutralLighter]
	);

	const compactCheckboxStyles = React.useMemo(
		() => ({
			root: {
				marginTop: 0,
				marginBottom: 0,
				marginLeft: 0,
				marginRight: 0,
				paddingTop: 0,
				paddingBottom: 0,
				minHeight: 26,
			},
			label: {
				fontSize: 14,
				alignItems: "center",
				minWidth: 0,
				width: "100%",
			},
			checkbox: {
				width: 16,
				height: 16,
				flexShrink: 0,
				marginRight: 6,
			},
			text: {
				marginLeft: 0,
				minWidth: 0,
				overflowWrap: "anywhere",
				lineHeight: "20px",
			},
		}),
		[]
	);

	const onApplyValueFilter = React.useCallback(() => {
		if (!filterColumn) {
			return;
		}

		const selectedOptions = filteredValueOptions.filter((option) => checkedFilterOptionKeys.includes(option.key));
		if (selectedOptions.length === 0) return;
		const allOptionsSelected =
			!searchText.trim() && allFilterOptions.length > 0 && allFilterOptions.every((option) => checkedFilterOptionKeys.includes(option.key));

		setIsLoading(true);
		if (allOptionsSelected) {
			clearPersistedValueFilterOptions(filterColumn.key);
			onFilter(filterColumn.key, "clear");
		} else {
			persistValueFilterOptions(filterColumn.key, allFilterOptions);
			const rawValues = selectedOptions.map((option) => option.value);
			const hasBlank = rawValues.includes("");
			const nonBlankValues = rawValues.filter((v) => v !== "");

			if (hasBlank && nonBlankValues.length === 0) {
				// Only "(Blanks)" selected — use doesNotContainData so Dataverse returns null/empty rows
				onFilter(filterColumn.key, "doesNotContainData");
			} else if (rawValues.length === 1) {
				onFilter(filterColumn.key, "equals", rawValues[0]);
			} else {
				onFilter(filterColumn.key, "in", rawValues);
			}
		}

		onFilterDismiss();
	}, [
		allFilterOptions,
		filteredValueOptions,
		searchText,
		checkedFilterOptionKeys,
		clearPersistedValueFilterOptions,
		filterColumn,
		onFilter,
		onFilterDismiss,
		persistValueFilterOptions,
	]);

	const onColumnContextMenu = React.useCallback(
		(column?: IColumn, ev?: React.MouseEvent<HTMLElement>) => {
			if (column && ev) {
				ev.preventDefault();
				openFilterCallout(column, ev.currentTarget as HTMLElement);
			}
		},
		[openFilterCallout]
	);

	// Dynamics 365 grid headers open a column-actions menu. Sorting remains
	// available in that menu, alongside filtering, instead of being a hidden
	// right-click-only interaction.
	const onColumnClick = React.useCallback(
		(ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
			ev.preventDefault();
			openFilterCallout(column, ev.currentTarget as HTMLElement);
		},
		[openFilterCallout]
	);

	const onNextPage = React.useCallback(() => {
		if (infiniteScroll) {
			requestNextPage(true);
		} else {
			setIsLoading(true);
			void Promise.resolve(loadNextPage()).catch(() => {setIsLoading(false);setInfiniteScrollError('Unable to load more records. Try again.');});
		}
	}, [infiniteScroll, loadNextPage, requestNextPage]);

	const onPreviousPage = React.useCallback(() => {
		setIsLoading(true);
		loadPreviousPage();
	}, [loadPreviousPage, setIsLoading]);

	const onFirstPage = React.useCallback(() => {
		setIsLoading(true);
		loadFirstPage();
	}, [loadFirstPage, setIsLoading]);

	const gridColumns = React.useMemo(() => {
		const visibleCols = visibleColumns;
		if (visibleCols.length === 0) return [] as IColumn[];

		// Fluent adds cell padding separately. Do not shrink columns below readable
		// sizes to force a fit: horizontal scrolling keeps headers and cells aligned.
		const preferredWidths = visibleCols.map((col) => {
			const kind = getFilterKind(col.dataType);
			const preferred = kind === "text" || kind === "lookup" ? 176 : kind === "boolean" ? 80 : 128;
			const factor = Number.isFinite(col.visualSizeFactor) && col.visualSizeFactor >= 60 ? col.visualSizeFactor : preferred;
			return columnWidths[col.name] ?? Math.max(80, Math.min(320, Math.max(factor, measureTextWidth(col.displayName) + 40)));
		});
		const resolvedWidths = fillColumnWidths(preferredWidths, availableColumnWidth - (hierarchySelectionGutter ? HIERARCHY_GUTTER_EXTRA_WIDTH : 0), visibleCols.map((col) => columnWidths[col.name] !== undefined));

		return visibleCols.map((col, i) => {
			const sortOn = sorting?.find((s) => s.name === col.name);
			const filteredColumnNames = getFilterColumnNames(filtering);
			const filtered = filteredColumnNames.some((name) => name === col.name);
			const colW = resolvedWidths[i];
			return {
				key: col.name,
				name: col.displayName,
				ariaLabel: `${col.displayName}, column actions`,
				fieldName: col.name,
				isSorted: sortOn != null,
				isSortedDescending: sortOn?.sortDirection === 1,
				isResizable: true,
				isFiltered: filtered,
				isMenuOpen: filterColumn?.key === col.name && filterTarget != null,
				columnActionsMode: ColumnActionsMode.hasDropdown,
				data: col,
				iconName: features.showTypeIcons ? columnTypeIcon(col.dataType) : undefined,
				minWidth: 80,
				maxWidth: Math.max(600, colW),
				calculatedWidth: colW,
				currentWidth: colW,
				className: getFilterKind(col.dataType) === "number" ? "cm-grid-number-cell" : undefined,
				styles: {
					cellTitle: { justifyContent: "flex-start", textAlign: "left" },
					cellName: { justifyContent: "flex-start", textAlign: "left" },
				},
				onColumnContextMenu: onColumnContextMenu,
				onColumnClick,
			} as IColumn;
		});
	}, [visibleColumns, sorting, filtering, filterColumn?.key, filterTarget, onColumnContextMenu, onColumnClick, columnWidths, availableColumnWidth, hierarchySelectionGutter, features.showTypeIcons]);
	rangeColumnsRef.current = gridColumns.slice(0, renderedColumnCount);
	React.useLayoutEffect(() => {
		const list = detailsListRef.current;
		if (!list) return;
		// Fluent caches fixed-column widths; changing calculatedWidth props alone
		// does not refresh that cache. Use its public API without treating automatic
		// expansion as a manual user resize or remounting the scrolling list.
		syncingColumnWidthsRef.current = true;
		try {
			for (const column of gridColumns) {
				const nextWidth = column.calculatedWidth;
				if (nextWidth !== undefined && appliedColumnWidthsRef.current.get(column.key) !== nextWidth) {
					list.updateColumn(column, { width: nextWidth });
					appliedColumnWidthsRef.current.set(column.key, nextWidth);
				}
			}
		} finally { syncingColumnWidthsRef.current = false; }
	}, [gridColumns]);

	const onActiveFilterClick = React.useCallback(
		(columnKey: string, target: HTMLElement) => {
			const column = gridColumns.find((c) => c.key === columnKey);
			if (!column) return;
			openFilterCallout(column, target, appliedValueFilterOptionsRef.current.get(columnKey));
		},
		[gridColumns, openFilterCallout]
	);

	const clearAllFilters = React.useCallback(() => {
		clearPersistedValueFilterOptions();
		setIsLoading(true);
		onClearAllFilters();
	}, [clearPersistedValueFilterOptions, onClearAllFilters]);

	const rootContainerStyle: React.CSSProperties = React.useMemo(() => {
		const hasExplicitHeight = typeof height === "number" && Number.isFinite(height) && height > 0;
		return {
			height: hasExplicitHeight ? height : "100%",
			minHeight: hasExplicitHeight ? undefined : 320,
			width: typeof width === "number" && Number.isFinite(width) && width > 0 ? width : "100%",
		};
	}, [width, height]);

	const onRenderItemColumn = React.useCallback(
		(item?: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord, _index?: number, column?: IColumn) => {
			if (!column?.fieldName || !item) {
				return <></>;
			}

			const formattedValue = item.getFormattedValue(column.fieldName);
			const columnMetadata = column.data as ComponentFramework.PropertyHelper.DataSetApi.Column | undefined;
			let rawCellValue: unknown;
			try {
				rawCellValue = item.getValue(column.fieldName);
			} catch {
				rawCellValue = undefined;
			}
			const dataType = columnMetadata?.dataType ?? "";
			const effectiveDataType =
				dataType ||
				(typeof rawCellValue === "number"
					? "decimal"
					: typeof rawCellValue === "string"
						? "singleline.text"
						: typeof rawCellValue === "boolean"
							? "twooptions"
							: "");
			const alias = (columnMetadata as { alias?: string } | undefined)?.alias;
			const isAliasedColumn = typeof alias === "string" && alias.length > 0 && alias !== column.key;
			const recordId = item.getRecordId();
			const cellKey = `${recordId}-${column.fieldName}`;
			const isEditingThisCell = editingCells.has(cellKey);
			const cellSaveStatus = cellSaveStatuses.get(cellKey);
			const normalizedType = effectiveDataType.toLowerCase();
			const isLookup = isLookupDataType(effectiveDataType);
			const lookupTarget = isLookup
				? getLookupReference(item, column.fieldName)?.entityType ??
					(columnMetadata as { targets?: string[] } | undefined)?.targets?.[0] ??
					lookupTargetEntityMap?.[column.key.toLowerCase()]
				: undefined;

			const metadataFlags = columnMetadata as
				| { isReadOnly?: boolean; readOnly?: boolean; disableEditing?: boolean }
				| undefined;
			const isReadOnly = metadataFlags?.isReadOnly === true || metadataFlags?.readOnly === true || metadataFlags?.disableEditing === true;

			const canEditThisCell =
				enableInlineEdit &&
				!isAliasedColumn &&
				!isReadOnly &&
				!itemsLoading &&
				!isComponentLoading &&
				(isLookup
					? !!lookupTarget && !!onSearchLookup && !!onUpdateLookupCell
					: isInlineEditableDataType(effectiveDataType));

			// ── Excel-like cell selection ─────────────────────────────────────
			const rowIndex = _index ?? 0;
			const colIndex = gridColumns.findIndex((c) => c.key === column.key);
			// Use cellSelection state (not ref) here so this callback re-evaluates on selection changes.
			const sel = cellSelection;
			const isInSelection = !!(sel &&
				rowIndex >= Math.min(sel.anchor.row, sel.head.row) &&
				rowIndex <= Math.max(sel.anchor.row, sel.head.row) &&
				colIndex >= Math.min(sel.anchor.col, sel.head.col) &&
				colIndex <= Math.max(sel.anchor.col, sel.head.col));
			const isTopEdge    = isInSelection && !!sel && rowIndex === Math.min(sel.anchor.row, sel.head.row);
			const isBottomEdge = isInSelection && !!sel && rowIndex === Math.max(sel.anchor.row, sel.head.row);
			const isLeftEdge   = isInSelection && !!sel && colIndex === Math.min(sel.anchor.col, sel.head.col);
			const isRightEdge  = isInSelection && !!sel && colIndex === Math.max(sel.anchor.col, sel.head.col);

			const handleCellMouseDown = (e: React.MouseEvent) => {
				if (e.button !== 0 || features.enableRangeSelection === false) return;
				setKeyboardTarget(undefined);
				cellMouseDownRef.current = true;
				if (isEditingThisCell) return;
				if (e.shiftKey && cellSelectionRef.current) {
					const updated = { anchor: cellSelectionRef.current.anchor, head: { row: rowIndex, col: colIndex } };
					cellSelectionRef.current = updated;
					setCellSelection(updated);
				} else {
					const newSel = { anchor: { row: rowIndex, col: colIndex }, head: { row: rowIndex, col: colIndex } };
					cellSelectionRef.current = newSel;
					setCellSelection(newSel);
					isDraggingRef.current = true;
				}
			};

			const handleCellMouseEnter = () => {
				if (!isDraggingRef.current || isEditingThisCell || features.enableRangeSelection === false) return;
				const current = cellSelectionRef.current;
				if (!current) return;
				const updated = { ...current, head: { row: rowIndex, col: colIndex } };
				cellSelectionRef.current = updated;
				setCellSelection(updated);
			};

			const handleCellKeyDown = (e: React.KeyboardEvent<HTMLDivElement>) => {
				if (isEditingThisCell) return;
				const target = e.target as HTMLElement;
				const targetTag = target.tagName.toLowerCase();
				if (targetTag === "a" || targetTag === "button" || targetTag === "input" || targetTag === "select" || targetTag === "textarea" || target.isContentEditable) {
					return;
				}
				if (features.enableRangeSelection !== false && !e.altKey) {
					const current = {row: rowIndex, col: colIndex};
					const edge = e.ctrlKey || e.metaKey;
					const head = navigateCell(current, e.key, {rows: displayItemsRef.current.length + (!edge && hasNextPage && !groups ? 1 : 0), columns: renderedColumnCount, visibleRows: visibleRangeRows, rtl: features.isRTL, edge,
						pageSize: Math.max(1, Math.floor((containerRef.current?.clientHeight ?? 430) / (listMode ? 44 + renderedColumnCount * 36 : 43)) - 1)});
					if (head) {
						e.preventDefault(); e.stopPropagation();
						setKeyboardTarget({head, anchor: e.shiftKey ? cellSelectionRef.current?.anchor ?? current : head, query: selectionQuery, layout: rangeLayoutKey});
						return;
					}
					if (e.key === 'Escape') {e.preventDefault(); e.stopPropagation(); setKeyboardTarget(undefined); cellSelectionRef.current = null; setCellSelection(null); return;}
				}
				if (canEditThisCell && (e.key === "Enter" || e.key === "F2" || e.key === ' ')) {
					e.preventDefault();
					e.stopPropagation(); setKeyboardTarget(undefined);
					beginEditCell(item, column, effectiveDataType, formattedValue);
				}
			};

			const cellOverlay = isInSelection ? (
				<div style={{
					position: 'absolute',
					top: 0, right: 0, bottom: 0, left: 0,
					backgroundColor: 'rgba(0, 120, 212, 0.10)',
					borderTop:    isTopEdge    ? '2px solid #0078d4' : '1px solid rgba(0,120,212,0.25)',
					borderBottom: isBottomEdge ? '2px solid #0078d4' : '1px solid rgba(0,120,212,0.25)',
					borderLeft:   isLeftEdge   ? '2px solid #0078d4' : '1px solid rgba(0,120,212,0.25)',
					borderRight:  isRightEdge  ? '2px solid #0078d4' : '1px solid rgba(0,120,212,0.25)',
					pointerEvents: 'none',
					boxSizing: 'border-box',
				}} />
			) : null;

			// ── Active editor ──────────────────────────────────────────────────
			if (isEditingThisCell) {
				const cellData = editingCells.get(cellKey)!;

				// Lookup editor
				if (isLookup && onSearchLookup && cellData.lookupEntityType) {
					const entityType = cellData.lookupEntityType;
					return (
						<EditingLookup
							key={cellKey}
							initialDisplayValue={cellData.value}
							ariaLabel={`Edit ${column.name}`}
							onSearch={(term) => onSearchLookup(entityType, term, cellData.columnName, cellData.recordId)}
							onBrowse={onBrowseLookup ? term => onBrowseLookup(entityType, term, cellData.columnName, cellData.recordId) : undefined}
							onCommit={(id) => void commitLookupCell(cellKey, id, entityType)}
							onCancel={() => cancelEditCell(cellKey)}
						/>
					);
				}

				// Choice (OptionSet) editor
				const isOptionSet = normalizedType.includes("optionset") && !normalizedType.includes("multiselect");
				const isTwoOptions = normalizedType.includes("twooptions") || normalizedType.includes("boolean");
				const options = columnOptions?.[column.fieldName];

				if (isOptionSet && options && options.length > 0) {
					return (
						<EditingDropdown
							key={cellKey}
							initialValue={cellData.value}
							ariaLabel={`Edit ${column.name}`}
							options={options}
							onCommit={(value) => void commitEditCell(cellKey, value)}
							onCancel={() => cancelEditCell(cellKey)}
						/>
					);
				}

				if (isTwoOptions) {
					return (
						<EditingDropdown
							key={cellKey}
							initialValue={cellData.value}
							ariaLabel={`Edit ${column.name}`}
							options={[
								{ label: "Yes", value: "true" },
								{ label: "No", value: "false" },
							]}
							onCommit={(value) => void commitEditCell(cellKey, value)}
							onCancel={() => cancelEditCell(cellKey)}
						/>
					);
				}

				// Plain text editor
				return (
					<EditingTextField
						key={cellKey}
						inputType={normalizedType.includes('dateonly') ? 'date' : 'text'}
						initialValue={cellData.value}
						ariaLabel={`Edit ${column.name}`}
						onCommit={(finalValue) => void commitEditCell(cellKey, finalValue)}
						onCancel={() => cancelEditCell(cellKey)}
					/>
				);
			}

			// ── Per-cell save status icon (shown at far left of cell value) ───
			const statusIcon =
				cellSaveStatus === "saving" ? (
					<Spinner size={SpinnerSize.xSmall} />
				) : cellSaveStatus === "saved" ? (
					<Icon iconName="CheckMark" style={{ color: theme.palette.green, fontSize: 12 }} />
				) : cellSaveStatus === "failed" ? (
					<Icon iconName="Error" style={{ color: theme.palette.redDark, fontSize: 12 }} />
				) : null;

			const withStatus = (content: React.ReactNode) =>
				statusIcon ? (
					<Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
						{statusIcon}
						<span>{content}</span>
					</Stack>
				) : (
					<>{content}</>
				);

			// ── Lookup link (read-only or editable) ───────────────────────────
			const navigationAllowed = allowLookupNavigation(enableLookupLinks, features.navigationTypes, columnMetadata?.isPrimary);
			const lookupReference = navigationAllowed ? getLookupReference(item, column.fieldName) : null;

			let innerContent: React.ReactNode;

			if (!isEditingThisCell && lookupReference && formattedValue) {
				const linkContent = (
					<Link
						href="#"
								onClick={(event) => {
							event.preventDefault();
							event.stopPropagation();
							onOpenLookup(lookupReference.entityType, lookupReference.id);
						}}>
						{formattedValue}
					</Link>
				);

				if (canEditThisCell && onSearchLookup) {
					innerContent = withStatus(linkContent);
				} else {
					innerContent = withStatus(linkContent);
				}
			} else if (navigationAllowed && columnMetadata?.isPrimary && formattedValue) {
				innerContent = withStatus(<Link href="#" onClick={event => {event.preventDefault(); event.stopPropagation(); onNavigate(item);}}>{formattedValue}</Link>);
			} else if (canEditThisCell) {
				// ── Editable plain cell ───────────────────────────────────────
				innerContent = withStatus(formattedValue);
			} else {
				innerContent = withStatus(formattedValue);
			}

			const choiceStyle = features.showChoiceColors ? choiceColorStyle(columnOptions?.[column.fieldName]?.find(option => option.value === toRawFilterString(rawCellValue))?.color) : undefined;
			if (choiceStyle) innerContent = <span className="cm-grid-choice-badge" style={choiceStyle}>{innerContent}</span>;
			if (features.showStatusIcons && (normalizedType === 'state' || normalizedType === 'status')) innerContent = <><Icon iconName={String(rawCellValue) === '0' ? 'StatusCircleCheckmark' : 'StatusCircleRing'} aria-hidden="true" />&nbsp;{innerContent}</>;
			return (
				<>
					{cellOverlay}
					{canEditThisCell && (
						<div
							style={{ position: 'absolute', top: 0, right: 0, bottom: 0, left: 0, cursor: 'text', zIndex: 1 }}
							onDoubleClick={() => { beginEditCell(item, column, effectiveDataType, formattedValue); }}
						/>
					)}
					<div
						style={{
							display: 'flex',
							alignItems: 'center',
							height: '100%',
							width: '100%',
							boxSizing: 'border-box',
							userSelect: 'none',
							position: 'relative',
							zIndex: 2,
						}}
						tabIndex={canEditThisCell || features.enableRangeSelection !== false ? 0 : undefined}
						data-cm-cell-row={rowIndex}
						data-cm-cell-col={colIndex}
						aria-label={`${column.name}: ${formattedValue || "(blank)"}`}
						onMouseDown={handleCellMouseDown}
						onMouseEnter={handleCellMouseEnter}
						onKeyDown={handleCellKeyDown}
						onDoubleClick={canEditThisCell ? () => { beginEditCell(item, column, effectiveDataType, formattedValue); } : undefined}
					>
						{innerContent}
					</div>
				</>
			);
		},
		[
			beginEditCell,
			cancelEditCell,
			commitEditCell,
			commitLookupCell,
			editingCells,
			cellSaveStatuses,
			cellSelection,
			columnOptions,
			features.enableRangeSelection,
			features.navigationTypes,
			features.isRTL,
			groups,
			hasNextPage,
			selectionQuery,
			rangeLayoutKey,
			renderedColumnCount,
			listMode,
			visibleRangeRows,
			features.showChoiceColors,
			features.showStatusIcons,
			enableInlineEdit,
			enableLookupLinks,
			getLookupReference,
			gridColumns,
			isComponentLoading,
			isInlineEditableDataType,
			isLookupDataType,
			itemsLoading,
			lookupTargetEntityMap,
			onOpenLookup,
			onNavigate,
			onSearchLookup,
			onBrowseLookup,
			onUpdateLookupCell,
			theme.palette.green,
			theme.palette.redDark,
		]
	);

	function renderChildToggle(record: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord) {
		if (props.hasChildRows && props.hasChildRows(record) !== true) return null;
		const id = record.getRecordId();
		const expanded = expandedRows.has(id);
		const firstColumn = gridColumns[0];
		const label = record.getFormattedValue(firstColumn?.fieldName ?? firstColumn?.key ?? '') || id;
		return <button type="button" className="cm-child-toggle"
			aria-label={`${expanded ? 'Collapse' : 'Expand'} child options for ${label}`}
			aria-expanded={expanded} data-selection-disabled={true}
			onMouseDown={event => event.stopPropagation()} onDoubleClick={event => event.stopPropagation()}
			onKeyDown={event => { if (event.key === 'Enter' || event.key === ' ') event.stopPropagation(); }}
			onClick={event => {
				event.stopPropagation();
				setVisitedChildRows(current => new Set(current).add(id));
				setExpandedRows(current => { const next = new Set(current); if (next.has(id)) next.delete(id); else next.add(id); return next; });
			}}><Icon iconName={expanded ? 'ChevronDown' : 'ChevronRight'} aria-hidden="true" /></button>;
	}

	const onRenderRow: IDetailsListProps["onRenderRow"] = (rowProps?: IDetailsRowProps) => {
		const customStyles: Partial<IDetailsRowStyles> = {};

		if (rowProps?.item) {
			if (highlightColor && highlightValue) {
				let indicatorValue: unknown;
				try {
					const itemRecord = rowProps.item as ComponentFramework.PropertyHelper.DataSetApi.EntityRecord;
					indicatorValue = itemRecord.getValue("HighlightIndicator");
				} catch {
					indicatorValue = undefined;
				}
				if (indicatorValue == highlightValue) {
					customStyles.root = { backgroundColor: highlightColor };
				}
			}
			const record = rowProps.item as ComponentFramework.PropertyHelper.DataSetApi.EntityRecord;
			if (listMode) {
				const cardColumns = gridColumns.slice(0, renderedColumnCount);
				const id = record.getRecordId();
				const firstValue = record.getFormattedValue(cardColumns[0]?.fieldName ?? '') || id;
				return <React.Fragment><div role="row" aria-selected={selectionIds.current.includes(id)} className={`cm-grid-list-row${features.bandedListRows !== false && rowProps.itemIndex % 2 ? ' is-banded' : ''}`} style={{height: 44 + cardColumns.length * 36}}>
					<div className="cm-grid-list-row-head">
						{!props.disableSelection && <input type="checkbox" aria-label={`Select ${firstValue}`} checked={selectionIds.current.includes(id)} onChange={event => selection.setKeySelected(id, event.target.checked, false)} />}
						{props.renderChildRows && renderChildToggle(record)}
						{features.showAvatar && <span className="cm-grid-avatar" aria-hidden="true">{firstValue.split(/\s+/).slice(0,2).map(word => word[0]).join('').toLocaleUpperCase()}</span>}
						{enableLookupLinks && features.navigationTypes !== 'none' ? <button type="button" className="cm-grid-record-link" onClick={() => onNavigate(record)}>{firstValue}</button> : <span>{firstValue}</span>}
					</div>
					{cardColumns.map(column => <div role="gridcell" className="cm-grid-list-field" key={column.key}>
						{features.showListColumnNames !== false && <span className="cm-grid-list-field-label">{column.name}</span>}
						<div className="cm-grid-list-field-value">{onRenderItemColumn(record,rowProps.itemIndex,column)}</div>
					</div>)}
				</div>{visitedChildRows.has(id) && <div hidden={!expandedRows.has(id)} className="cm-child-retained">{renderChildren(record)}</div>}</React.Fragment>;
			}
			return <React.Fragment><DetailsRow {...rowProps} styles={customStyles}
				onRenderCheck={hierarchySelectionGutter ? checkProps => <>
					<DetailsRowCheck {...checkProps} />{renderChildToggle(record)}
				</> : rowProps.onRenderCheck} />
				{visitedChildRows.has(record.getRecordId()) && <div hidden={!expandedRows.has(record.getRecordId())} className="cm-child-retained">{renderChildren(record)}</div>}
			</React.Fragment>;
		}

		return null;
	};

	function renderChildren(record: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord) {
		return props.renderChildRows?.(record, expandedRows.has(record.getRecordId()));
	}
	function renderTotals(rows: readonly DataSet[], label: string) {
		return <div className="cm-grid-totals" role="status" aria-label={label}><strong>{label}</strong>{Object.entries(aggregateColumns).filter(([key]) => columns.some(column => column.name === key)).map(([key, mode]) => {
			const value = aggregateNumbers(rows.map(row => row.getValue(key)), mode);
			return <span key={key}>{columns.find(column => column.name === key)?.displayName} ({mode}): <strong>{value === null ? '—' : value.toLocaleString(undefined, {maximumFractionDigits: 10})}</strong></span>;
		})}</div>;
	}

	const jumpColumn = columns.find(column => column.name === sorting[0]?.name) ?? columns.find(column => column.isPrimary) ?? columns[0];
	const jumpFilter = getJumpFilter(filtering);
	const canJump = !!jumpColumn && getFilterKind(jumpColumn.dataType) === 'text' && !jumpColumn.isHidden && features.enableFiltering !== false;
	return (
		<div
			ref={containerRef}
			dir={features.isRTL ? 'rtl' : undefined}
			className={`cm-model-driven-grid${hierarchySelectionGutter ? ' cm-grid-hierarchy' : ''}${listMode ? ' cm-grid-list-mode' : ''}`}
			style={{ ...rootContainerStyle, display: "flex", flexDirection: "column", minWidth: 0 }}
		>
		<Stack verticalFill grow className="cm-grid-layout" style={{ height: "100%", width: "100%" }}>
			{listMode && <div className="cm-grid-list-actions">
				{!props.disableSelection && features.enableMultipleSelection !== false && <DefaultButton text={selectionIds.current.length ? 'Clear selection' : 'Select all'} onClick={() => selectionCallbacks.current.toggleAll()} />}
				<div ref={listColumnMenuRef}><ActionButton text="Columns" menuProps={{items: gridColumns.map(column => ({key:column.key, text:column.name, onClick: () => {if (listColumnMenuRef.current) openFilterCallout(column, listColumnMenuRef.current);}}))}} /></div>
			</div>}
			{clipboardError && <div role="alert" className="cm-grid-load-error">{clipboardError}</div>}
			{actionError && !filterTarget && <div role="alert" className="cm-grid-load-error">{actionError}</div>}
			{layoutMessage && <div role="status" className="cm-grid-analysis-status">{layoutMessage}</div>}
			{analysisRequested && <div className="cm-grid-analysis-status">
				{groupColumn && <span>Grouped by {columns.find(column => column.name === groupColumn)?.displayName}</span>}
				{!analysisReady && <span role="status">{analysisPaused ? 'Summary loading paused.' : `Preparing complete results (${items.length.toLocaleString()} records loaded)…`}</span>}
				{hasNextPage && <DefaultButton text={analysisPaused ? 'Resume' : 'Pause'} onClick={() => setAnalysisPaused(!analysisPaused)} />}
				<DefaultButton text="Clear grouping and totals" onClick={() => {setGroupColumn(undefined); setAggregateColumns({}); setAnalysisPaused(false);}} />
			</div>}
			{analysisReady && Object.keys(aggregateColumns).length > 0 && renderTotals(items, 'Totals — all matching records')}
			{widthEditor && <Dialog hidden={false} onDismiss={() => setWidthEditor(undefined)} dialogContentProps={{title: `${widthEditor.label} column width`}}>
				<TextField label="Width (pixels)" type="number" min={80} max={2000} value={widthEditor.value} errorMessage={widthError}
					onChange={(_, value) => {setWidthEditor({...widthEditor, value: value ?? ''}); setWidthError('');}} />
				<DialogFooter><PrimaryButton text="Apply" onClick={() => {
					try {const next = numericColumnWidth(widthEditor.value); setColumnWidths(current => ({...current, [widthEditor.key]: next})); setWidthEditor(undefined);}
					catch (error) {setWidthError(error instanceof Error ? error.message : 'Invalid width.');}
				}}/><DefaultButton text="Cancel" onClick={() => setWidthEditor(undefined)}/></DialogFooter>
			</Dialog>}
			<Stack.Item grow className="cm-grid-surface" style={{ position: "relative", backgroundColor: "white", zIndex: 0, minHeight: 0 }}>
				{!itemsLoading && !isComponentLoading && items?.length === 0 && (
					<Stack grow horizontalAlign="center" verticalAlign="center" className="noRecords">
						<Icon iconName="PageList"></Icon>
						<Text variant="large">{resources.getString("Label_NoRecords")}</Text>
					</Stack>
				)}
				<ScrollablePane className="cm-grid-scroll-pane" style={{ height: "100%" }} scrollbarVisibility={ScrollbarVisibility.auto}>
					<DetailsList
						componentRef={detailsListRef}
						cellStyleProps={GRID_CELL_STYLE_PROPS}
						className="cm-grid-details"
						ariaLabelForGrid={features.accessibleLabel?.trim() ? features.accessibleLabel : 'Records'}
						key={infiniteScroll ? "details-infinite" : `details-${currentPage}`}
						columns={gridColumns}
						isHeaderVisible={!listMode}
						groups={groups}
						groupProps={{
							isAllGroupsCollapsed: !!groups?.length && groups.every(group => group.isCollapsed),
							onToggleCollapseAll: collapsed => setGroupCollapse(Object.fromEntries((groups ?? []).map(group => [group.key, !!collapsed]))),
							onRenderHeader: header => header ? <><GroupHeader {...header} onToggleCollapse={group => {if (group) setGroupCollapse(current => ({...current, [group.key]: !group.isCollapsed}));}} />{header.group && Object.keys(aggregateColumns).length > 0 && renderTotals(displayItems.slice(header.group.startIndex, header.group.startIndex + header.group.count), 'Group totals')}</> : null
						}}
						columnReorderOptions={features.enableColumnMoving === false ? undefined : {onColumnDrop: ({draggedIndex, targetIndex}) => setColumnOrder(moveColumn(gridColumns.map(column => column.key), draggedIndex, targetIndex))}}
						onColumnResize={(column, newWidth) => {
							if (!column || newWidth === undefined || syncingColumnWidthsRef.current) return;
							const key = column.key;
							appliedColumnWidthsRef.current.set(key, newWidth);
							setColumnWidths((current) => current[key] === newWidth ? current : ({ ...current, [key]: newWidth }));
						}}
						onRenderItemColumn={(item, index, column) => {
							const content = onRenderItemColumn(item as ComponentFramework.PropertyHelper.DataSetApi.EntityRecord | undefined, index, column);
							if ((!props.renderRowActions && (!props.renderChildRows || hierarchySelectionGutter)) || !item || column?.key !== gridColumns[0]?.key) return content;
							const record = item as ComponentFramework.PropertyHelper.DataSetApi.EntityRecord;
							return <div className="cm-child-cell">{props.renderChildRows && !hierarchySelectionGutter && renderChildToggle(record)}<div className="cm-child-value">{content}</div>{props.renderRowActions?.(record)}</div>;
						}}
						onShouldVirtualize={() => true}
						listProps={{renderedWindowsAhead: 1, renderedWindowsBehind: 1,
							getPageHeight: (start = 0, _rect, count = 10) => displayItems.slice(start, start + count).reduce((height, row) => height + (listMode ? 44 + renderedColumnCount * 36 : 43) + (expandedRows.has(row.getRecordId()) ? 380 : 0), 0)}}
						selectionMode={props.disableSelection ? SelectionMode.none : features.enableMultipleSelection === false ? SelectionMode.single : SelectionMode.multiple}
						checkboxVisibility={hierarchySelectionGutter ? CheckboxVisibility.always : CheckboxVisibility.onHover}
						onRenderCheckbox={hierarchySelectionGutter ? renderHierarchyCheckbox : undefined}
						onRenderDetailsHeader={onRenderDetailsHeader}
						items={displayItems}
						setKey={infiniteScroll ? "set-infinite" : `set${currentPage}`}
						initialFocusedIndex={0}
						checkButtonAriaLabel="select row"
						ariaLabelForSelectAllCheckbox={selectingAll ? 'Cancel selecting records' : selectionIds.current.length ? 'Clear selection' : 'Select all records (up to 1,000)'}
						layoutMode={DetailsListLayoutMode.fixedColumns}
						constrainMode={ConstrainMode.unconstrained}
						useReducedRowRenderer={false}
						selection={selection}
						onItemInvoked={enableInlineEdit || !enableLookupLinks || features.navigationTypes === 'none' ? undefined : onNavigate}
						onRenderRow={onRenderRow}></DetailsList>
					{infiniteScroll && <div ref={sentinelRef} style={{ height: 1 }} />}
					{filterTarget && filterColumn && (
						<Callout
							className={`cm-grid-column-callout${filterView === "filterBy" ? " cm-grid-column-callout-dynamic" : ""}`}
							ariaLabel={`Filter ${filterColumn.name}`}
							target={filterTarget}
							onDismiss={onFilterDismiss}
							directionalHint={DirectionalHint.bottomLeftEdge}
							gapSpace={0}
							isBeakVisible={false}
							calloutMaxHeight={560}
							calloutWidth={320}
							styles={{ root: { height: filterView === "main" ? "min(560px, calc(100vh - 24px))" : "auto", maxWidth: "calc(100vw - 24px)" } }}
							style={{ maxHeight: "inherit" }}
							hideOverflow
							setInitialFocus>
							<div className="cm-column-menu">
							{filterView === "main" ? <>
								<div className="cm-column-menu-body">
								<div className="cm-column-menu-title">{filterColumn.name}</div>
								<ActionButton styles={menuRowStyles} text="Column options" iconProps={{iconName: 'ColumnOptions'}} menuProps={{items: [
									{key: 'width', text: 'Column width', onClick: () => {setWidthError(''); setWidthEditor({key: filterColumn.key, label: filterColumn.name, value: String(Math.round(filterColumn.calculatedWidth ?? 176))}); onFilterDismiss();}},
									{key: 'left', text: 'Move left', disabled: features.enableColumnMoving === false || gridColumns[0]?.key === filterColumn.key, onClick: () => {const index = gridColumns.findIndex(column => column.key === filterColumn.key); setColumnOrder(moveColumn(gridColumns.map(column => column.key), index, index - 1)); onFilterDismiss();}},
									{key: 'right', text: 'Move right', disabled: features.enableColumnMoving === false || gridColumns[gridColumns.length - 1]?.key === filterColumn.key, onClick: () => {const index = gridColumns.findIndex(column => column.key === filterColumn.key); setColumnOrder(moveColumn(gridColumns.map(column => column.key), index, index + 1)); onFilterDismiss();}},
									{key: 'group', text: groupColumn === filterColumn.key ? 'Ungroup' : 'Group by this column', disabled: features.enableGrouping === false || !infiniteScroll, onClick: () => {setGroupColumn(groupColumn === filterColumn.key ? undefined : filterColumn.key); setAnalysisPaused(false); onFilterDismiss();}},
									...(getFilterKind((filterColumn.data as ComponentFramework.PropertyHelper.DataSetApi.Column).dataType) === 'date' ? [{key: 'date-group', text: 'Group dates by', disabled: features.enableGrouping === false || !infiniteScroll, subMenuProps: {items: dateGroupingIntervals.map(interval => ({key: interval, text: interval[0].toUpperCase() + interval.slice(1), canCheck: true, checked: groupColumn === filterColumn.key && dateGrouping === interval, onClick: () => {setDateGrouping(interval); setGroupColumn(filterColumn.key); setAnalysisPaused(false); onFilterDismiss();}}))}}] : []),
									{key: 'reset-layout', text: 'Reset column layout', onClick: () => {setColumnOrder([]); setColumnWidths({}); onFilterDismiss();}},
									{key: 'total', text: 'Totals', disabled: features.enableAggregation === false || !infiniteScroll || getFilterKind((filterColumn.data as ComponentFramework.PropertyHelper.DataSetApi.Column).dataType) !== 'number', subMenuProps: {items: [...(['sum','average','minimum','maximum'] as const).map(mode => ({key: mode, text: mode[0].toUpperCase() + mode.slice(1), canCheck: true, checked: aggregateColumns[filterColumn.key] === mode, onClick: () => {setAggregateColumns(current => ({...current, [filterColumn.key]: mode})); setAnalysisPaused(false); onFilterDismiss();}})), {key: 'none', text: 'None', onClick: () => {setAggregateColumns(current => {const next = {...current}; delete next[filterColumn.key]; return next;}); onFilterDismiss();}}]}},
								]}} />
								{actionError && <div role="alert" className="cm-filter-error">{actionError}</div>}
								{sorting.length > 0 && <div className="cm-filter-sort-summary">Sort order: {sorting.map((sort) => `${columns.find((column) => column.name === sort.name)?.displayName ?? sort.name} ${sort.sortDirection === 1 ? "↓" : "↑"}`).join(" → ")}</div>}
								{(["asc", "desc"] as const).map((direction) => {
									const kind = getFilterKind((filterColumn.data as ComponentFramework.PropertyHelper.DataSetApi.Column).dataType);
									const label = kind === "number" ? (direction === "asc" ? "Smallest to largest" : "Largest to smallest") : kind === "date" ? (direction === "asc" ? "Oldest to newest" : "Newest to oldest") : direction === "asc" ? "A to Z" : "Z to A";
									return <ActionButton key={direction} styles={menuRowStyles} text={label} iconProps={{ iconName: direction === "asc" ? "SortUp" : "SortDown" }}
										aria-pressed={filterColumn.isSorted && filterColumn.isSortedDescending === (direction === "desc")}
										disabled={features.enableSorting === false || (filterColumn.data as ComponentFramework.PropertyHelper.DataSetApi.Column).disableSorting}
										title="Hold Shift while choosing a direction to sort by another column too."
										onClick={(event) => { setIsLoading(true); try { onSort(filterColumn.key, direction, event.shiftKey); onFilterDismiss(); } catch (error) { setIsLoading(false); setActionError(error instanceof Error ? error.message : "Unable to sort records."); } }} />;
								})}
								<ActionButton styles={menuRowStyles} text="Clear sort" iconProps={{ iconName: "Clear" }} disabled={features.enableSorting === false || (filterColumn.data as ComponentFramework.PropertyHelper.DataSetApi.Column).disableSorting === true || !filterColumn.isSorted}
									onClick={() => { setIsLoading(true); try { onSort(filterColumn.key, "none"); onFilterDismiss(); } catch (error) { setIsLoading(false); setActionError(error instanceof Error ? error.message : "Unable to sort records."); } }} />
								<div className="cm-column-menu-separator" />
								<ActionButton styles={menuRowStyles} text="Clear Filter" iconProps={{ iconName: "ClearFilter" }} disabled={!filterColumn.isFiltered}
									onClick={() => { clearPersistedValueFilterOptions(filterColumn.key); setIsLoading(true); onFilter(filterColumn.key, "clear"); onFilterDismiss(); }} />
								<ActionButton styles={menuRowStyles} text={getFilterKind((filterColumn.data as ComponentFramework.PropertyHelper.DataSetApi.Column).dataType) === "number" ? "Number filters" : getFilterKind((filterColumn.data as ComponentFramework.PropertyHelper.DataSetApi.Column).dataType) === "date" ? "Date filters" : "Filter by"}
									iconProps={{ iconName: "Filter" }} disabled={features.enableFiltering === false} menuIconProps={{ iconName: "ChevronRight" }} onClick={() => setFilterView("filterBy")} />
								<div className="cm-column-menu-separator" />
								<div className="cm-column-menu-values">
									<TextField placeholder="Search values" ariaLabel="Search filter values" value={searchText}
										onChange={(_event, value) => {
											const query = value ?? ""; setSearchText(query);
											setCheckedFilterOptionKeys(allFilterOptions.filter((option) => option.label.toLocaleLowerCase().includes(query.toLocaleLowerCase())).map((option) => option.key));
										}} />
									<div className="cm-column-value-list" role="group" aria-label={`${filterColumn.name} filter values`}>
										<Checkbox label="(Select All)" styles={compactCheckboxStyles} checked={allVisibleSelected}
											indeterminate={!allVisibleSelected && filteredValueOptions.some((option) => checkedFilterOptionKeys.includes(option.key))}
											onChange={(_event, checked) => setCheckedFilterOptionKeys((current) => checked ? Array.from(new Set([...current, ...filteredValueOptions.map((option) => option.key)])) : current.filter((key) => !filteredValueOptions.some((option) => option.key === key)))} />
										{filteredValueOptions.map((option) => <Checkbox key={option.key} label={option.label} styles={compactCheckboxStyles}
											checked={checkedFilterOptionKeys.includes(option.key)}
											onChange={(_event, checked) => setCheckedFilterOptionKeys((current) => checked ? Array.from(new Set([...current, option.key])) : current.filter((key) => key !== option.key))} />)}
										{filteredValueOptions.length === 0 && <div role="status">No matching values</div>}
									</div>
								</div>
								</div>
								<div className="cm-column-menu-footer">
									<div className="cm-filter-note">Values from loaded records. Conditions filter the full dataset.</div>
									<div className="cm-filter-actions"><PrimaryButton text="OK" onClick={onApplyValueFilter} disabled={features.enableFiltering === false || !filteredValueOptions.some((option) => checkedFilterOptionKeys.includes(option.key))} /><DefaultButton text="Cancel" onClick={onFilterDismiss} /></div>
								</div>
							</> : <ColumnFilterEditor columnName={filterColumn.name}
								dataType={(filterColumn.data as ComponentFramework.PropertyHelper.DataSetApi.Column).dataType}
								initialGroup={editorFilterGroup}
								valueOptions={allFilterOptions.filter((option) => !option.isBlank).map(({label, value}) => ({label, value}))}
								onCancel={() => setFilterView("main")}
								onApply={(group) => {
									setIsLoading(true);
									try { onFilter(filterColumn.key, "custom", group); persistValueFilterOptions(filterColumn.key, allFilterOptions); onFilterDismiss(); }
									catch (error) { setIsLoading(false); throw error; }
								}} />}
							</div>
						</Callout>
					)}
				</ScrollablePane>
				{(itemsLoading || isComponentLoading) && !(pagingInProgress && items.length > 0) && (
					<Overlay className="cm-grid-loading-overlay">
						<Spinner size={SpinnerSize.medium} label="Loading records" />
					</Overlay>
				)}
			</Stack.Item>
			{activeFilters.length > 0 && (
				<Stack.Item className="cm-grid-active-filters">
					<Stack
						horizontal
						wrap
						verticalAlign="center"
						tokens={{ childrenGap: 4 }}
						style={{
							padding: "4px 8px",
							borderTop: `1px solid ${theme.palette.neutralLight}`,
							background: theme.palette.neutralLighterAlt,
						}}
					>
						<Text variant="small" style={{ color: theme.palette.neutralSecondary, whiteSpace: "nowrap" }}>
							Filtered by:
						</Text>
						{activeFilters.map((f) => (
							<div
								key={f.columnKey}
								role="group"
								aria-label={`Filter ${f.displayName}: ${f.valueLabel}`}
								style={{
									display: "inline-flex",
									alignItems: "center",
									background: theme.palette.white,
									border: `1px solid ${theme.palette.neutralTertiaryAlt}`,
									borderRadius: 2,
									padding: "1px 4px 1px 6px",
									fontSize: 12,
									gap: 2,
								}}
							>
								<button
									type="button"
									aria-label={`Edit filter on ${f.displayName}`}
									onClick={(event) => onActiveFilterClick(f.columnKey, event.currentTarget)}
									style={{
										border: 0,
										background: "transparent",
										color: "inherit",
										font: "inherit",
										padding: 0,
										cursor: "pointer",
									}}
								>
									<strong>{f.displayName}</strong>
									<span style={{ color: theme.palette.neutralSecondary, marginLeft: 2 }}>{f.valueLabel}</span>
								</button>
								<IconButton
									iconProps={{ iconName: "Cancel" }}
									ariaLabel={`Clear filter on ${f.displayName}`}
									styles={{ root: { height: 20, width: 20, minWidth: 20 }, icon: { fontSize: 10 } }}
									onClick={() => {
										clearPersistedValueFilterOptions(f.columnKey);
										setIsLoading(true);
										onFilter(f.columnKey, "clear");
									}}
								/>
							</div>
						))}
						{activeFilters.length > 1 && (
							<DefaultButton
								text="Clear All"
								styles={{ root: { height: 22, minWidth: 0, padding: "0 8px" }, label: { fontSize: 12 } }}
								onClick={clearAllFilters}
							/>
						)}
					</Stack>
				</Stack.Item>
			)}
			{features.enableJumpBar && <nav className="cm-grid-jump-bar" aria-label={`Jump by ${jumpColumn?.displayName ?? 'name'}`}>
				{['All', ...'ABCDEFGHIJKLMNOPQRSTUVWXYZ'].map(letter => <button key={letter} type="button"
					aria-pressed={letter === 'All' ? !jumpFilter : jumpFilter?.column === jumpColumn?.name && jumpFilter.letter === letter}
					disabled={itemsLoading || isComponentLoading || (letter === 'All' ? !jumpFilter : !canJump)}
					title={letter === 'All' ? 'Remove only the alphabet filter' : canJump ? `${jumpColumn?.displayName} begins with ${letter}` : 'Sort by a text column to use the alphabet bar'}
					onClick={() => {setIsLoading(true); try {onFilter(jumpColumn?.name ?? '', 'jump', letter === 'All' ? '' : letter);} catch (error) {setIsLoading(false); setActionError(error instanceof Error ? error.message : 'Unable to filter records.');}}}>{letter}</button>)}
			</nav>}
			<Stack.Item className="cm-grid-footer">
				<Stack horizontal verticalAlign="center" style={{ width: "100%" }}>
					<Stack.Item align="center">
						<Text variant="small" className="cm-grid-record-count">
							{stringFormat(
								resources.getString("Label_Grid_Footer_RecordCount"),
								totalResultCount < 0 ? `${items.length}${hasNextPage || !infiniteScroll ? '+' : ''}` : totalResultCount.toString(),
								(props.selectionCountOverride ?? selectionIds.current.length).toString()
							)}
							{activeFilters.length > 0 && (
								<span style={{ marginLeft: 6, color: theme.palette.themePrimary, fontStyle: "italic" }}>
									(filtered)
								</span>
							)}
						</Text>
						{selectingAll && <Text variant="small" role="status">Selecting matching records (up to {MAX_SELECTED_RECORDS.toLocaleString('en-GB')})…</Text>}
						{(selectionError || selectionMessage) && <Text variant="small" role={selectionError ? 'alert' : 'status'} styles={{root: {display: 'block'}}}>{selectionError || selectionMessage}</Text>}
						{(props.commandError ?? props.commandStatus) && <Text variant="small" role={props.commandError ? 'alert' : 'status'} styles={{root: {display: 'block', color: props.commandError ? theme.palette.redDark : theme.palette.neutralSecondary}}}>{props.commandError ?? props.commandStatus}</Text>}
					</Stack.Item>
					<Stack.Item grow />
					{infiniteScroll ? (
						<Stack.Item align="center" className="cm-grid-load-more">
							{infiniteScrollError ? (
								<Stack horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
									<Text variant="small" className="cm-grid-load-error">{infiniteScrollError}</Text>
									<DefaultButton text="Retry" onClick={onNextPage} />
								</Stack>
							) : (itemsLoading || isComponentLoading || pagingInProgress) && hasNextPage ? (
								<Spinner size={SpinnerSize.small} label="Loading..." labelPosition="right" />
							) : null}
						</Stack.Item>
					) : (
						<>
						<IconButton
							ariaLabel="First page"
							iconProps={{ iconName: "Rewind" }}
							disabled={!hasPreviousPage || isComponentLoading || itemsLoading}
							onClick={onFirstPage}
						/>
						<IconButton
							ariaLabel="Previous page"
							iconProps={{ iconName: "Previous" }}
							disabled={!hasPreviousPage || isComponentLoading || itemsLoading}
							onClick={onPreviousPage}
						/>
						<Stack.Item align="center">
							{stringFormat(
								resources.getString("Label_Grid_Footer"),
								currentPage.toString(),
								selection.getSelectedCount().toString()
							)}
						</Stack.Item>
						<IconButton
							ariaLabel="Next page"
							iconProps={{ iconName: "Next" }}
							disabled={!hasNextPage || isComponentLoading || itemsLoading}
							onClick={onNextPage}
						/>
						</>
					)}
				</Stack>
			</Stack.Item>
		</Stack>
		</div>
	);
});

Grid.displayName = "Grid";
