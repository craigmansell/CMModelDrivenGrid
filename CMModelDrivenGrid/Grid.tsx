import { useConst, useForceUpdate } from "@fluentui/react-hooks";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { IObjectWithKey, IRenderFunction, SelectionMode } from "@fluentui/react/lib/Utilities";
import {
	ConstrainMode,
	DetailsList,
	DetailsListLayoutMode,
	DetailsRow,
	IColumn,
	IDetailsHeaderProps,
	IDetailsListProps,
	IDetailsRowStyles,
} from "@fluentui/react/lib/DetailsList";
import { Sticky, StickyPositionType } from "@fluentui/react/lib/Sticky";
import { Callout, DirectionalHint } from "@fluentui/react/lib/Callout";
import { ScrollablePane, ScrollbarVisibility } from "@fluentui/react/lib/ScrollablePane";
import { Stack } from "@fluentui/react/lib/Stack";
import { Overlay } from "@fluentui/react/lib/Overlay";
import { ActionButton, DefaultButton, IconButton, PrimaryButton } from "@fluentui/react/lib/Button";
import { Selection } from "@fluentui/react/lib/Selection";
import { Link } from "@fluentui/react/lib/Link";
import { Icon } from "@fluentui/react/lib/Icon";
import { Text } from "@fluentui/react/lib/Text";
import { TextField } from "@fluentui/react/lib/TextField";
import { Checkbox } from "@fluentui/react/lib/Checkbox";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { useTheme } from "@fluentui/react";

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

type FilterByMode =
	| "equals"
	| "notEquals"
	| "containsData"
	| "doesNotContainData"
	| "contains"
	| "notContains"
	| "beginsWith"
	| "notBeginsWith"
	| "endsWith"
	| "notEndsWith";

function stringFormat(template: string, ...args: string[]): string {
	args?.forEach((arg, index) => {
		template = template.replace("{" + index + "}", arg);
	});
	return template;
}

export interface GridProps {
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
	columnOptions?: Record<string, { label: string; value: string }[]>;
	setSelectedRecords: (ids: string[]) => void;
	onNavigate: (item?: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord) => void;
	onOpenLookup: (entityType: string, id: string) => void;
	onSort: (name: string, direction: "asc" | "desc" | "none") => void;
	onFilter: (
		name: string,
		mode:
			| "contains"
			| "notContains"
			| "equals"
			| "notEquals"
			| "beginsWith"
			| "notBeginsWith"
			| "endsWith"
			| "notEndsWith"
			| "containsData"
			| "doesNotContainData"
			| "in"
			| "clear",
		value?: string | string[]
	) => void;
	loadFirstPage: () => void;
	loadNextPage: () => void;
	loadPreviousPage: () => void;
	onUpdateCell: (recordId: string, columnName: string, value: string, dataType: string) => Promise<void> | void;
	onUpdateLookupCell?: (recordId: string, columnName: string, targetId: string, targetEntityType: string) => Promise<void>;
	onSearchLookup?: (entityType: string, searchTerm: string) => Promise<{ id: string; name: string }[]>;
	onFullScreen: () => void;
	isFullScreen: boolean;
	item?: DataSet;
}

// Self-contained editing input — manages its own value state so typing always
// works regardless of whether DetailsList re-renders the parent row.
interface EditingTextFieldProps {
	initialValue: string;
	onCommit: (value: string) => void;
	onCancel: () => void;
}

const EditingTextField: React.FC<EditingTextFieldProps> = ({ initialValue, onCommit, onCancel }) => {
	const [value, setValue] = React.useState(initialValue);
	const valueRef = React.useRef(initialValue);
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
			value={value}
			autoFocus
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
	initialValue: string;
	options: { label: string; value: string }[];
	onCommit: (value: string) => void;
	onCancel: () => void;
}

const EditingDropdown: React.FC<EditingDropdownProps> = ({ initialValue, options, onCommit, onCancel }) => {
	const committedRef = React.useRef(false);

	const handleChange = (event: React.ChangeEvent<HTMLSelectElement>) => {
		if (committedRef.current) return;
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
		committedRef.current = true;
		onCommit(event.target.value);
	};

	return (
		<select
			autoFocus
			defaultValue={initialValue}
			onChange={handleChange}
			onKeyDown={handleKeyDown}
			onBlur={handleBlur}
			style={{ width: "100%", minHeight: 24, fontSize: "inherit", border: "none", outline: "none" }}
		>
			{options.map((opt) => (
				<option key={opt.value} value={opt.value}>
					{opt.label}
				</option>
			))}
		</select>
	);
};

// Searchable lookup picker — shows a text input and a Callout dropdown of results.
interface EditingLookupProps {
	initialDisplayValue: string;
	onSearch: (term: string) => Promise<{ id: string; name: string }[]>;
	onCommit: (id: string, name: string) => void;
	onCancel: () => void;
}

const EditingLookup: React.FC<EditingLookupProps> = ({ initialDisplayValue, onSearch, onCommit, onCancel }) => {
	const [searchTerm, setSearchTerm] = React.useState(initialDisplayValue);
	const [results, setResults] = React.useState<{ id: string; name: string }[]>([]);
	const [isSearching, setIsSearching] = React.useState(false);
	const [dropdownRect, setDropdownRect] = React.useState<DOMRect | null>(null);
	const committedRef = React.useRef(false);
	const debounceRef = React.useRef<number | undefined>(undefined);
	const containerRef = React.useRef<HTMLDivElement>(null);

	const performSearch = React.useCallback(
		async (term: string) => {
			setIsSearching(true);
			try {
				const found = await onSearch(term);
				setResults(found);
			} catch {
				setResults([]);
			} finally {
				setIsSearching(false);
			}
		},
		[onSearch]
	);

	// Capture position and kick off initial search on mount.
	React.useEffect(() => {
		if (containerRef.current) {
			setDropdownRect(containerRef.current.getBoundingClientRect());
		}
		void performSearch(initialDisplayValue);
		return () => {
			if (debounceRef.current !== undefined) {
				window.clearTimeout(debounceRef.current);
			}
		};
	// eslint-disable-next-line react-hooks/exhaustive-deps
	}, []);

	const handleChange = (_ev: React.FormEvent, newValue?: string) => {
		const term = newValue ?? "";
		setSearchTerm(term);
		if (debounceRef.current !== undefined) {
			window.clearTimeout(debounceRef.current);
		}
		debounceRef.current = window.setTimeout(() => {
			void performSearch(term);
		}, 300);
	};

	const handleKeyDown = (event: React.KeyboardEvent) => {
		if (event.key === "Escape") {
			event.preventDefault();
			committedRef.current = true;
			onCancel();
		}
	};

	// Delay cancel so a mousedown on a result item fires before we cancel.
	const handleBlur = () => {
		window.setTimeout(() => {
			if (!committedRef.current) {
				committedRef.current = true;
				onCancel();
			}
		}, 200);
	};

	const handleSelect = (id: string, name: string) => {
		if (committedRef.current) return;
		committedRef.current = true;
		onCommit(id, name);
	};

	const dropdownPortal =
		dropdownRect !== null &&
		(results.length > 0 || isSearching) &&
		ReactDOM.createPortal(
			<div
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
				{isSearching && (
					<Stack horizontal verticalAlign="center" tokens={{ childrenGap: 6 }} style={{ padding: "6px 8px" }}>
						<Spinner size={SpinnerSize.xSmall} />
						<Text variant="small">Searching...</Text>
					</Stack>
				)}
				{!isSearching && results.length === 0 && (
					<Text variant="small" style={{ padding: "6px 8px", display: "block", color: "#666" }}>
						No results found
					</Text>
				)}
				{!isSearching &&
					results.map((r) => (
						<ActionButton
							key={r.id}
							styles={{ root: { width: "100%", textAlign: "left", height: 28, paddingLeft: 8 } }}
							onMouseDown={(e) => e.preventDefault()}
							onClick={() => handleSelect(r.id, r.name)}
						>
							{r.name}
						</ActionButton>
					))}
			</div>,
			document.body
		);

	return (
		<>
			<div ref={containerRef}>
				<TextField
					autoFocus
					borderless
					value={searchTerm}
					styles={{ root: { width: "100%" }, fieldGroup: { minHeight: 24 } }}
					onChange={handleChange}
					onKeyDown={handleKeyDown}
					onBlur={handleBlur}
				/>
			</div>
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
		resources,
		loadFirstPage,
		loadNextPage,
		loadPreviousPage,
		onFullScreen,
		isFullScreen,
		highlightValue,
		highlightColor,
		enableLookupLinks,
		enableInlineEdit,
		totalResultCount,
		onOpenLookup,
		onUpdateCell,
		onUpdateLookupCell,
		onSearchLookup,
		columnOptions,
	} = props;
	const theme = useTheme();
	const blankValueLabel = "(Blanks)";

	const forceUpdate = useForceUpdate();
	const onSelectionChanged = (): void => {
		const items = selection.getItems() as DataSet[];
		const selected = selection.getSelectedIndices().map((index: number) => {
			const item: DataSet | undefined = items[index];
			return item && items[index].getRecordId();
		});

		setSelectedRecords(selected);
		forceUpdate();
	};

	const selection: Selection = useConst(() => {
		return new Selection({
			selectionMode: SelectionMode.multiple,
			onSelectionChanged: onSelectionChanged,
		});
	});

	const [isComponentLoading, setIsLoading] = React.useState<boolean>(false);
	const [filterColumn, setFilterColumn] = React.useState<IColumn | undefined>();
	const [filterTarget, setFilterTarget] = React.useState<HTMLElement | undefined>();
	const [allFilterValues, setAllFilterValues] = React.useState<string[]>([]);
	const [checkedFilterValues, setCheckedFilterValues] = React.useState<string[]>([]);
	const [searchText, setSearchText] = React.useState<string>("");
	const [filterByTarget, setFilterByTarget] = React.useState<HTMLElement | undefined>();
	const [filterByMode, setFilterByMode] = React.useState<FilterByMode>("equals");
	const [filterByValue, setFilterByValue] = React.useState<string>("");

	// Multi-cell editing state — each key is `${recordId}-${columnName}`.
	const [editingCells, setEditingCells] = React.useState<Map<string, EditingCell>>(() => new Map());
	const editingCellsRef = React.useRef<Map<string, EditingCell>>(new Map());
	editingCellsRef.current = editingCells;

	// Per-cell save status — same key format as editingCells.
	const [cellSaveStatuses, setCellSaveStatuses] = React.useState<Map<string, CellSaveStatus>>(() => new Map());
	const cellSaveTimeoutsRef = React.useRef<Map<string, number>>(new Map());

	// Cleanup all pending timeouts on unmount.
	React.useEffect(() => {
		// eslint-disable-next-line react-hooks/exhaustive-deps
		return () => { cellSaveTimeoutsRef.current.forEach((id) => window.clearTimeout(id)); };
	}, []);

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
		const normalizedType = dataType.toLowerCase();
		const unsupportedMarkers = [
			"customer",
			"partylist",
			"regarding",
			"image",
			"file",
			"multiselectoptionset",
		];
		return !unsupportedMarkers.some((marker) => normalizedType.includes(marker));
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
				initialValue = initialRawValue.toISOString();
			}

			// For lookup fields, capture the target entity type from the existing reference.
			let lookupEntityType: string | undefined;
			if (isLookupDataType(dataType)) {
				const ref = getLookupReference(item, column.key);
				lookupEntityType = ref?.entityType;
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
		[getLookupReference, isLookupDataType]
	);

	const cancelEditCell = React.useCallback((cellKey: string) => {
		setEditingCells((prev) => {
			if (!prev.has(cellKey)) return prev;
			const next = new Map(prev);
			next.delete(cellKey);
			return next;
		});
	}, []);

	const commitEditCell = React.useCallback(
		async (cellKey: string, finalValue: string) => {
			const current = editingCellsRef.current.get(cellKey);
			if (!current) return;

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
		[onUpdateCell, setCellTransientStatus]
	);

	const commitLookupCell = React.useCallback(
		async (cellKey: string, targetId: string, targetEntityType: string) => {
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
		[onUpdateLookupCell, setCellTransientStatus]
	);

	const items: DataSet[] = React.useMemo(() => {
		setIsLoading(false);

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
	}, [records, sortedRecordIds, setIsLoading]);

	// eslint-disable-next-line react-hooks/exhaustive-deps
	const displayItems = React.useMemo(() => [...items], [items, editingCells, cellSaveStatuses]);

	const isFilterValueRequired = React.useCallback((mode: FilterByMode): boolean => {
		return mode !== "containsData" && mode !== "doesNotContainData";
	}, []);

	const deriveFilterByState = React.useCallback(
		(
			condition?: ComponentFramework.PropertyHelper.DataSetApi.ConditionExpression
		): { mode: FilterByMode; value: string } => {
			if (!condition) {
				return { mode: "equals", value: "" };
			}

			const conditionValue = typeof condition.value === "string" ? condition.value : "";
			const parseLikeValue = (value: string): "contains" | "beginsWith" | "endsWith" => {
				if (value.startsWith("%") && value.endsWith("%") && value.length >= 2) {
					return "contains";
				}
				if (value.endsWith("%")) {
					return "beginsWith";
				}
				if (value.startsWith("%")) {
					return "endsWith";
				}
				return "contains";
			};
			const stripLikeWildcards = (value: string): string => value.replace(/^%/, "").replace(/%$/, "");

			const operator = condition.conditionOperator as number;
			switch (operator) {
				case 0:
					return { mode: "equals", value: conditionValue };
				case 1:
					return { mode: "notEquals", value: conditionValue };
				case 6: {
					const parsedMode = parseLikeValue(conditionValue);
					return { mode: parsedMode, value: stripLikeWildcards(conditionValue) };
				}
				case 7: {
					const parsedMode = parseLikeValue(conditionValue);
					const negativeMode: FilterByMode =
						parsedMode === "contains"
							? "notContains"
							: parsedMode === "beginsWith"
								? "notBeginsWith"
								: "notEndsWith";
					return { mode: negativeMode, value: stripLikeWildcards(conditionValue) };
				}
				case 12:
					return { mode: "doesNotContainData", value: "" };
				case 13:
					return { mode: "containsData", value: "" };
				default:
					return { mode: "equals", value: "" };
			}
		},
		[]
	);

	const onFilterDismiss = React.useCallback(() => {
		setFilterColumn(undefined);
		setFilterTarget(undefined);
		setFilterByTarget(undefined);
		setAllFilterValues([]);
		setCheckedFilterValues([]);
		setSearchText("");
		setFilterByMode("equals");
		setFilterByValue("");
	}, []);

	const openFilterCallout = React.useCallback(
		(column: IColumn, target: HTMLElement) => {
			const uniqueValues = Array.from(
				new Set(
					(items ?? [])
						.filter((item): item is DataSet => item !== undefined)
						.map((item) => {
							const formattedValue = item.getFormattedValue(column.key);
							return formattedValue && formattedValue.trim().length > 0 ? formattedValue : blankValueLabel;
						})
				)
			).sort((left, right) => left.localeCompare(right));

			const existingCondition = filtering?.conditions?.find((condition) => condition.attributeName === column.key);
			const filterByState = deriveFilterByState(existingCondition);
			let initialCheckedValues = [...uniqueValues];

			if (existingCondition?.conditionOperator === 8 && Array.isArray(existingCondition.value)) {
				initialCheckedValues = existingCondition.value.map((value) =>
					value && value.trim().length > 0 ? value : blankValueLabel
				);
			} else if (
				(existingCondition?.conditionOperator === 0 || existingCondition?.conditionOperator === 1) &&
				typeof existingCondition.value === "string"
			) {
				initialCheckedValues = [
					existingCondition.value && existingCondition.value.trim().length > 0
						? existingCondition.value
						: blankValueLabel,
				];
			}

			setFilterColumn(column);
			setFilterTarget(target);
			setAllFilterValues(uniqueValues);
			setCheckedFilterValues(initialCheckedValues.filter((value) => uniqueValues.includes(value)));
			setSearchText("");
			setFilterByMode(filterByState.mode);
			setFilterByValue(filterByState.value);
		},
		[blankValueLabel, deriveFilterByState, filtering?.conditions, items]
	);

	const openFilterByPopup = React.useCallback((target: HTMLElement) => {
		setFilterByTarget(target);
	}, []);

	const closeFilterByPopup = React.useCallback(() => {
		setFilterByTarget(undefined);
	}, []);

	const applyFilterBy = React.useCallback(() => {
		if (!filterColumn) {
			return;
		}

		const trimmedValue = filterByValue.trim();
		if (isFilterValueRequired(filterByMode) && trimmedValue.length === 0) {
			return;
		}

		if (isFilterValueRequired(filterByMode)) {
			onFilter(filterColumn.key, filterByMode, trimmedValue);
		} else {
			onFilter(filterColumn.key, filterByMode);
		}

		setIsLoading(true);
		onFilterDismiss();
	}, [filterByMode, filterByValue, filterColumn, isFilterValueRequired, onFilter, onFilterDismiss]);

	const clearFilterBy = React.useCallback(() => {
		if (!filterColumn) {
			return;
		}

		onFilter(filterColumn.key, "clear");
		setIsLoading(true);
		onFilterDismiss();
	}, [filterColumn, onFilter, onFilterDismiss]);

	const filteredValueOptions = React.useMemo(() => {
		if (!searchText.trim()) {
			return allFilterValues;
		}

		const loweredSearch = searchText.toLocaleLowerCase();
		return allFilterValues.filter((value) => value.toLocaleLowerCase().includes(loweredSearch));
	}, [allFilterValues, searchText]);

	const allVisibleSelected = filteredValueOptions.every((value) => checkedFilterValues.includes(value));

	const menuRowStyles = React.useMemo(
		() => ({
			root: {
				width: "100%",
				justifyContent: "flex-start",
				textAlign: "left",
				minHeight: 22,
				height: 28,
				paddingLeft: 4,
				paddingRight: 4,
				paddingTop: 0,
				paddingBottom: 0,
				borderRadius: 0,
				fontWeight: "400",
			},
			label: {
				fontWeight: "400",
				fontSize: 12,
				lineHeight: "28px",
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
				minHeight: 20,
			},
			label: {
				fontSize: 12,
			},
		}),
		[]
	);

	const filterByOptions = React.useMemo(
		() => [
			{ key: "equals", text: "Equals" },
			{ key: "notEquals", text: "Does not equal" },
			{ key: "containsData", text: "Contains data" },
			{ key: "doesNotContainData", text: "Does not contain data" },
			{ key: "contains", text: "Contains" },
			{ key: "notContains", text: "Does not contain" },
			{ key: "beginsWith", text: "Begins with" },
			{ key: "notBeginsWith", text: "Does not begin with" },
			{ key: "endsWith", text: "Ends with" },
			{ key: "notEndsWith", text: "Does not end with" },
		] as { key: FilterByMode; text: string }[],
		[]
	);

	const onApplyValueFilter = React.useCallback(() => {
		if (!filterColumn) {
			return;
		}

		if (checkedFilterValues.length === 0 || checkedFilterValues.length === allFilterValues.length) {
			onFilter(filterColumn.key, "clear");
		} else {
			const rawValues = checkedFilterValues.map((value) => (value === blankValueLabel ? "" : value));
			if (rawValues.length === 1) {
				onFilter(filterColumn.key, "equals", rawValues[0]);
			} else {
				onFilter(filterColumn.key, "in", rawValues);
			}
		}

		setIsLoading(true);
		onFilterDismiss();
	}, [allFilterValues.length, blankValueLabel, checkedFilterValues, filterColumn, onFilter, onFilterDismiss]);

	const onColumnContextMenu = React.useCallback(
		(column?: IColumn, ev?: React.MouseEvent<HTMLElement>) => {
			if (column && ev) {
				ev.preventDefault();
				openFilterCallout(column, ev.currentTarget as HTMLElement);
			}
		},
		[openFilterCallout]
	);

	const onColumnClick = React.useCallback(
		(_ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
			if (column) {
				if (column.isSorted && column.isSortedDescending) {
					onSort(column.key, "none");
				} else if (column.isSorted) {
					onSort(column.key, "desc");
				} else {
					onSort(column.key, "asc");
				}

				if (items && items.length > 0) {
					setIsLoading(true);
				}
			}
		},
		[items, onSort, setIsLoading]
	);

	const onNextPage = React.useCallback(() => {
		setIsLoading(true);
		loadNextPage();
	}, [loadNextPage, setIsLoading]);

	const onPreviousPage = React.useCallback(() => {
		setIsLoading(true);
		loadPreviousPage();
	}, [loadPreviousPage, setIsLoading]);

	const onFirstPage = React.useCallback(() => {
		setIsLoading(true);
		loadFirstPage();
	}, [loadFirstPage, setIsLoading]);

	const gridColumns = React.useMemo(() => {
		return columns
			.filter((col) => !col.isHidden && (typeof col.order !== "number" || col.order >= 0))
			.sort((a, b) => (a.order ?? 0) - (b.order ?? 0))
			.map((col) => {
				const sortOn = sorting?.find((s) => s.name === col.name);
				const filtered = filtering?.conditions?.find((f) => f.attributeName == col.name);
				return {
					key: col.name,
					name: col.displayName,
					fieldName: col.name,
					isSorted: sortOn != null,
					isSortedDescending: sortOn?.sortDirection === 1,
					isResizable: true,
					isFiltered: filtered != null,
					data: col,
					minWidth: col.visualSizeFactor > 100 ? col.visualSizeFactor : 100,
					onColumnContextMenu: onColumnContextMenu,
					onColumnClick: onColumnClick,
				} as IColumn;
			});
	}, [columns, sorting, filtering?.conditions, onColumnContextMenu, onColumnClick]);

	const rootContainerStyle: React.CSSProperties = React.useMemo(() => {
		return {
			height: typeof height === "number" && Number.isFinite(height) && height > 0 ? height : "100%",
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
			const recordId = item.getRecordId();
			const cellKey = `${recordId}-${column.fieldName}`;
			const isEditingThisCell = editingCells.has(cellKey);
			const cellSaveStatus = cellSaveStatuses.get(cellKey);
			const normalizedType = effectiveDataType.toLowerCase();
			const isLookup = isLookupDataType(effectiveDataType);

			const isPrimitiveRawValue =
				rawCellValue === null ||
				rawCellValue === undefined ||
				typeof rawCellValue === "string" ||
				typeof rawCellValue === "number" ||
				typeof rawCellValue === "boolean";

			const canEditThisCell =
				enableInlineEdit &&
				!itemsLoading &&
				!isComponentLoading &&
				((isPrimitiveRawValue || (!!effectiveDataType && isInlineEditableDataType(effectiveDataType))) ||
					(isLookup && !!onSearchLookup));

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
							onSearch={(term) => onSearchLookup(entityType, term)}
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
						initialValue={cellData.value}
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
			const lookupReference = enableLookupLinks ? getLookupReference(item, column.fieldName) : null;
			if (!isEditingThisCell && lookupReference && formattedValue) {
				const linkContent = (
					<Link
						onClick={(event) => {
							event?.preventDefault();
							event?.stopPropagation();
							onOpenLookup(lookupReference.entityType, lookupReference.id);
						}}>
						{formattedValue}
					</Link>
				);

				if (canEditThisCell && onSearchLookup) {
					return (
						<span
							onDoubleClick={(event) => {
								event.preventDefault();
								event.stopPropagation();
								beginEditCell(item, column, effectiveDataType, formattedValue);
							}}
							style={{ display: "block", width: "100%", minHeight: "1em" }}
							title="Double-click to edit">
							{withStatus(linkContent)}
						</span>
					);
				}

				return withStatus(linkContent);
			}

			// ── Editable plain cell ───────────────────────────────────────────
			if (canEditThisCell) {
				return (
					<span
						onDoubleClick={(event) => {
							event.preventDefault();
							event.stopPropagation();
							beginEditCell(item, column, effectiveDataType, formattedValue);
						}}
						style={{ cursor: "text", display: "block", width: "100%", minHeight: "1em" }}
						title="Double-click to edit">
						{withStatus(formattedValue)}
					</span>
				);
			}

			return withStatus(formattedValue);
		},
		[
			beginEditCell,
			cancelEditCell,
			commitEditCell,
			commitLookupCell,
			editingCells,
			cellSaveStatuses,
			columnOptions,
			enableInlineEdit,
			enableLookupLinks,
			getLookupReference,
			isComponentLoading,
			isInlineEditableDataType,
			isLookupDataType,
			itemsLoading,
			onOpenLookup,
			onSearchLookup,
			theme.palette.green,
			theme.palette.redDark,
		]
	);

	const onRenderRow: IDetailsListProps["onRenderRow"] = (props) => {
		const customStyles: Partial<IDetailsRowStyles> = {};

		if (props?.item) {
			if (highlightColor && highlightValue) {
				let indicatorValue: unknown;
				try {
					const itemRecord = props.item as ComponentFramework.PropertyHelper.DataSetApi.EntityRecord;
					indicatorValue = itemRecord.getValue("HighlightIndicator");
				} catch {
					indicatorValue = undefined;
				}
				if (indicatorValue == highlightValue) {
					customStyles.root = { backgroundColor: highlightColor };
				}
			}
			return <DetailsRow {...props} styles={customStyles} />;
		}

		return null;
	};

	return (
		<Stack verticalFill grow style={rootContainerStyle}>
			<Stack.Item grow style={{ position: "relative", backgroundColor: "white", zIndex: 0, minHeight: 0 }}>
				{!itemsLoading && !isComponentLoading && items?.length === 0 && (
					<Stack grow horizontalAlign="center" className={"noRecords"}>
						<Icon iconName="PageList"></Icon>
						<Text variant="large">{resources.getString("Label_NoRecords")}</Text>
					</Stack>
				)}
				<ScrollablePane style={{ height: "100%" }} scrollbarVisibility={ScrollbarVisibility.auto}>
					<DetailsList
						key={`details-${currentPage}`}
						columns={gridColumns}
						onRenderItemColumn={onRenderItemColumn}
						onRenderDetailsHeader={onRenderDetailsHeader}
						items={displayItems}
						setKey={`set${currentPage}`}
						initialFocusedIndex={0}
						checkButtonAriaLabel="select row"
						layoutMode={DetailsListLayoutMode.fixedColumns}
						constrainMode={ConstrainMode.unconstrained}
						useReducedRowRenderer={false}
						selection={selection}
						onItemInvoked={enableInlineEdit ? undefined : onNavigate}
						onRenderRow={onRenderRow}></DetailsList>
					{filterTarget && filterColumn && (
						<Callout
							target={filterTarget}
							onDismiss={onFilterDismiss}
							directionalHint={DirectionalHint.bottomLeftEdge}
							gapSpace={8}
							calloutMaxHeight={560}
							setInitialFocus>
							<Stack tokens={{ childrenGap: 0 }} style={{ width: 280, padding: 4 }}>
								<ActionButton
									styles={menuRowStyles}
									text={resources.getString("Label_SortAZ")}
									iconProps={{ iconName: "SortUp" }}
									onClick={() => {
										onSort(filterColumn.key, "asc");
										setIsLoading(true);
										onFilterDismiss();
									}}
								/>
								<ActionButton
									styles={menuRowStyles}
									text={resources.getString("Label_SortZA")}
									iconProps={{ iconName: "SortDown" }}
									onClick={() => {
										onSort(filterColumn.key, "desc");
										setIsLoading(true);
										onFilterDismiss();
									}}
								/>
							<div style={{ height: 1, background: theme.palette.neutralLight, margin: "2px 0" }} />
								<ActionButton
									styles={menuRowStyles}
									text="Clear Filter"
									iconProps={{ iconName: "ClearFilter" }}
									onClick={() => {
										onFilter(filterColumn.key, "clear");
										setIsLoading(true);
										onFilterDismiss();
									}}
								/>
							<div style={{ height: 1, background: theme.palette.neutralLight, margin: "2px 0" }} />
								<ActionButton
									styles={menuRowStyles}
									text="Filter by"
									iconProps={{ iconName: "Filter" }}
									menuIconProps={{ iconName: "ChevronRight" }}
									onClick={(event) => {
										if (event?.currentTarget) {
											openFilterByPopup(event.currentTarget as HTMLElement);
										}
									}}
								/>
							<div style={{ height: 1, background: theme.palette.neutralLight, margin: "2px 0" }} />
								<TextField
									placeholder="Search"
									value={searchText}
									styles={{ root: { marginTop: 2, marginBottom: 2 }, fieldGroup: { minHeight: 24 } }}
									onChange={(_event, newValue) => setSearchText(newValue ?? "")}
								/>
								<Checkbox
									label="(Select All)"
									styles={compactCheckboxStyles}
									checked={allVisibleSelected}
									onChange={(_event, checked) => {
										if (checked) {
											setCheckedFilterValues((currentValues) =>
												Array.from(new Set([...currentValues, ...filteredValueOptions]))
											);
										} else {
											setCheckedFilterValues((currentValues) =>
												currentValues.filter((value) => !filteredValueOptions.includes(value))
											);
										}
									}}
								/>
								<Stack
									styles={{ root: { maxHeight: 150, overflowY: "auto", padding: 2, borderStyle: "solid", borderWidth: 1, borderColor: theme.palette.neutralLight, marginTop: 2 } }}
									tokens={{ childrenGap: 2 }}>
									{filteredValueOptions.map((value) => (
										<Checkbox
											key={value}
											label={value}
											styles={compactCheckboxStyles}
											checked={checkedFilterValues.includes(value)}
											onChange={(_event, checked) => {
												if (checked) {
													setCheckedFilterValues((currentValues) =>
														currentValues.includes(value) ? currentValues : [...currentValues, value]
													);
												} else {
													setCheckedFilterValues((currentValues) =>
														currentValues.filter((currentValue) => currentValue !== value)
													);
												}
											}}
										/>
									))}
								</Stack>
								<Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 6 }} style={{ marginTop: 2 }}>
									<PrimaryButton text="OK" onClick={onApplyValueFilter} disabled={checkedFilterValues.length === 0} />
									<DefaultButton text="Cancel" onClick={onFilterDismiss} />
								</Stack>
							</Stack>
						</Callout>
					)}
					{filterByTarget && filterColumn && (
						<Callout
							target={filterByTarget}
							onDismiss={closeFilterByPopup}
							directionalHint={DirectionalHint.rightTopEdge}
							gapSpace={8}
							setInitialFocus>
							<Stack tokens={{ childrenGap: 10 }} style={{ width: 280, padding: 12 }}>
								<Stack horizontal horizontalAlign="space-between" verticalAlign="center">
									<Text variant="mediumPlus">Filter by</Text>
									<IconButton iconProps={{ iconName: "Cancel" }} onClick={closeFilterByPopup} />
								</Stack>
								<select
									value={filterByMode}
									onChange={(event) => setFilterByMode(event.currentTarget.value as FilterByMode)}
									style={{ width: "100%", minHeight: 32 }}>
									{filterByOptions.map((option) => (
										<option key={option.key} value={option.key}>
											{option.text}
										</option>
									))}
								</select>
								{isFilterValueRequired(filterByMode) && (
									<TextField
										value={filterByValue}
										onChange={(_event, newValue) => setFilterByValue(newValue ?? "")}
									/>
								)}
								<Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 8 }}>
									<PrimaryButton
										text="Apply"
										onClick={applyFilterBy}
										disabled={isFilterValueRequired(filterByMode) && filterByValue.trim().length === 0}
									/>
									<DefaultButton text="Clear" onClick={clearFilterBy} />
								</Stack>
							</Stack>
						</Callout>
					)}
				</ScrollablePane>
				{(itemsLoading || isComponentLoading) && <Overlay />}
			</Stack.Item>
			<Stack.Item>
				<Stack horizontal style={{ width: "100%", paddingLeft: 8, paddingRight: 8 }}>
					<Stack.Item align="center">
						<Text>
							{stringFormat(
								resources.getString("Label_Grid_Footer_RecordCount"),
								totalResultCount === -1 ? "5000+" : totalResultCount.toString(),
								selection.getSelectedCount().toString()
							)}
						</Text>
					</Stack.Item>
					<Stack.Item grow align="center" style={{ textAlign: "center" }}>
						{!isFullScreen && <Link onClick={onFullScreen}>{resources.getString("Label_ShowFullScreen")}</Link>}
					</Stack.Item>
					<IconButton
						alt="First Page"
						iconProps={{ iconName: "Rewind" }}
						disabled={!hasPreviousPage || isComponentLoading || itemsLoading}
						onClick={onFirstPage}
					/>
					<IconButton
						alt="Previous Page"
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
						alt="Next Page"
						iconProps={{ iconName: "Next" }}
						disabled={!hasNextPage || isComponentLoading || itemsLoading}
						onClick={onNextPage}
					/>
				</Stack>
			</Stack.Item>
		</Stack>
	);
});

Grid.displayName = "Grid";
