import { useConst, useForceUpdate } from "@fluentui/react-hooks";
import * as React from "react";
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
import { IContextualMenuProps } from "@fluentui/react/lib/ContextualMenu";
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
import { useTheme } from "@fluentui/react";

type DataSet = ComponentFramework.PropertyHelper.DataSetApi.EntityRecord & IObjectWithKey;

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
	setSelectedRecords: (ids: string[]) => void;
	onNavigate: (item?: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord) => void;
	onSort: (name: string, direction: "asc" | "desc" | "none") => void;
	onFilter: (
		name: string,
		mode: "contains" | "equals" | "notEquals" | "beginsWith" | "endsWith" | "in" | "clear",
		value?: string | string[]
	) => void;
	loadFirstPage: () => void;
	loadNextPage: () => void;
	loadPreviousPage: () => void;
	onFullScreen: () => void;
	isFullScreen: boolean;
	item?: DataSet;
}

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

const onRenderItemColumn = (
	item?: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord,
	index?: number,
	column?: IColumn
) => {
	if (column?.fieldName && item) {
		return <>{item?.getFormattedValue(column.fieldName)}</>;
	}
	return <></>;
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
		totalResultCount,
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

	const onFilterDismiss = React.useCallback(() => {
		setFilterColumn(undefined);
		setFilterTarget(undefined);
		setAllFilterValues([]);
		setCheckedFilterValues([]);
		setSearchText("");
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
		},
		[blankValueLabel, filtering?.conditions, items]
	);

	const promptTextFilter = React.useCallback(
		(mode: "contains" | "equals" | "notEquals" | "beginsWith" | "endsWith", label: string) => {
			if (!filterColumn) {
				return;
			}

			const input = window.prompt(`${label} filter for '${filterColumn.name}'. Leave empty to clear filter:`, "");
			if (input === null) {
				return;
			}

			if (input.trim().length === 0) {
				onFilter(filterColumn.key, "clear");
			} else {
				onFilter(filterColumn.key, mode, input.trim());
			}

			setIsLoading(true);
			onFilterDismiss();
		},
		[filterColumn, onFilter, onFilterDismiss]
	);

	const filteredValueOptions = React.useMemo(() => {
		if (!searchText.trim()) {
			return allFilterValues;
		}

		const loweredSearch = searchText.toLocaleLowerCase();
		return allFilterValues.filter((value) => value.toLocaleLowerCase().includes(loweredSearch));
	}, [allFilterValues, searchText]);

	const allVisibleSelected = filteredValueOptions.every((value) => checkedFilterValues.includes(value));

	const textFiltersMenuProps = React.useMemo<IContextualMenuProps>(
		() => ({
			items: [
				{ key: "equals", text: "Equals...", onClick: () => promptTextFilter("equals", "Equals") },
				{
					key: "notEquals",
					text: "Does Not Equal...",
					onClick: () => promptTextFilter("notEquals", "Does Not Equal"),
				},
				{
					key: "beginsWith",
					text: "Begins With...",
					onClick: () => promptTextFilter("beginsWith", "Begins With"),
				},
				{ key: "endsWith", text: "Ends With...", onClick: () => promptTextFilter("endsWith", "Ends With") },
				{ key: "contains", text: "Contains...", onClick: () => promptTextFilter("contains", "Contains") },
			],
			directionalHint: DirectionalHint.rightTopEdge,
			isBeakVisible: false,
		}),
		[promptTextFilter]
	);

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
			.filter((col) => !col.isHidden && col.order >= 0)
			.sort((a, b) => a.order - b.order)
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

	const onRenderRow: IDetailsListProps["onRenderRow"] = (props) => {
		const customStyles: Partial<IDetailsRowStyles> = {};

		if (props?.item) {
			if (highlightColor && highlightValue) {
				let indicatorValue: unknown;
				try {
					// HighlightIndicator is optional in this fork and may not be bound on all app grids.
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
						columns={gridColumns}
						onRenderItemColumn={onRenderItemColumn}
						onRenderDetailsHeader={onRenderDetailsHeader}
						items={items}
						setKey={`set${currentPage}`} // Ensures that the selection is reset when paging
						initialFocusedIndex={0}
						checkButtonAriaLabel="select row"
						layoutMode={DetailsListLayoutMode.fixedColumns}
						constrainMode={ConstrainMode.unconstrained}
						selection={selection}
						onItemInvoked={onNavigate}
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
									text="Text Filters"
									menuIconProps={{ iconName: "ChevronRight" }}
									menuProps={textFiltersMenuProps}
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
				</ScrollablePane>
				{(itemsLoading || isComponentLoading) && <Overlay />}
			</Stack.Item>
			<Stack.Item>
				<Stack horizontal style={{ width: "100%", paddingLeft: 8, paddingRight: 8 }}>
					<Stack.Item align="center">
						{stringFormat(
							resources.getString("Label_Grid_Footer_RecordCount"),
							totalResultCount === -1 ? "5000+" : totalResultCount.toString(),
							selection.getSelectedCount().toString()
						)}
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
