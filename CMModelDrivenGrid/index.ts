import { initializeIcons } from "@fluentui/react/lib/Icons";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { Grid } from "./Grid";

// Register icons - but ignore warnings if they have been already registered by Power Apps
initializeIcons(undefined, { disableWarnings: true });

export class CMModelDrivenGrid implements ComponentFramework.StandardControl<IInputs, IOutputs> {
	notifyOutputChanged: () => void;
	container: HTMLDivElement;
	context: ComponentFramework.Context<IInputs>;
	sortedRecordsIds: string[] = [];
	resources: ComponentFramework.Resources;
	isTestHarness: boolean;
	records: Record<string, ComponentFramework.PropertyHelper.DataSetApi.EntityRecord>;
	currentPage = 1;
	isFullScreen = false;
	columnOptions: Record<string, { label: string; value: string }[]> = {};
	private columnOptionsFetchedFor: string | undefined;
	private entityMetaCache: Record<string, { entitySetName: string; primaryNameAttr: string }> = {};

	setSelectedRecords = (ids: string[]): void => {
		this.context.parameters.records.setSelectedRecordIds(ids);
	};

	onNavigate = (item?: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord): void => {
		if (item) {
			this.context.parameters.records.openDatasetItem(item.getNamedReference());
		}
	};

	private normalizeGuid = (id: string): string => id.replace(/[{}]/g, "");

	private parseInlineEditValue = (value: string, dataType: string): string | number | boolean | null => {
		const trimmedValue = value.trim();
		if (trimmedValue.length === 0) {
			return null;
		}

		const normalizedType = dataType.toLowerCase();
		if (normalizedType.includes("twooptions") || normalizedType.includes("boolean")) {
			const loweredValue = trimmedValue.toLowerCase();
			if (loweredValue === "true" || loweredValue === "yes" || loweredValue === "1") {
				return true;
			}
			if (loweredValue === "false" || loweredValue === "no" || loweredValue === "0") {
				return false;
			}
			throw new Error("Please enter true/false, yes/no, or 1/0.");
		}

		if (normalizedType.includes("dateandtime") || normalizedType.includes("datetime") || normalizedType.includes("date")) {
			const parsedDate = new Date(trimmedValue);
			if (Number.isNaN(parsedDate.getTime())) {
				throw new Error("Please enter a valid date value.");
			}
			return parsedDate.toISOString();
		}

		if (normalizedType.includes("whole.") || normalizedType.includes("optionset")) {
			const parsedInteger = Number.parseInt(trimmedValue, 10);
			if (Number.isNaN(parsedInteger)) {
				throw new Error("Please enter a whole number.");
			}
			return parsedInteger;
		}

		if (
			normalizedType.includes("decimal") ||
			normalizedType.includes("currency") ||
			normalizedType.includes("fp")
		) {
			const parsedDecimal = Number.parseFloat(trimmedValue);
			if (Number.isNaN(parsedDecimal)) {
				throw new Error("Please enter a numeric value.");
			}
			return parsedDecimal;
		}

		return trimmedValue;
	};

	onOpenLookup = (entityType: string, id: string): void => {
		if (!entityType || !id) {
			return;
		}

		void this.context.navigation.openForm({
			entityName: entityType,
			entityId: this.normalizeGuid(id),
		});
	};

	onUpdateCell = async (recordId: string, columnName: string, value: string, dataType: string): Promise<void> => {
		try {
			const entityType = this.context.parameters.records.getTargetEntityType();
			if (!entityType) {
				throw new Error("Unable to determine target table for inline editing.");
			}

			const parsedValue = this.parseInlineEditValue(value, dataType);
			const updateData: ComponentFramework.WebApi.Entity = {
				[columnName]: parsedValue,
			};

			await this.context.webAPI.updateRecord(entityType, this.normalizeGuid(recordId), updateData);
			this.context.parameters.records.refresh();
		} catch (error) {
			const message = error instanceof Error ? error.message : "Unable to save the field value.";
			await this.context.navigation.openAlertDialog({ text: message });
			throw error;
		}
	};

	onSort = (name: string, direction: "asc" | "desc" | "none"): void => {
		const sorting = this.context.parameters.records.sorting;
		while (sorting.length > 0) {
			sorting.pop();
		}

		if (direction !== "none") {
			this.context.parameters.records.sorting.push({
				name: name,
				sortDirection: direction === "desc" ? 1 : 0,
			});
		}

		this.context.parameters.records.refresh();
	};

	onFilter = (
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
	): void => {
		const filtering = this.context.parameters.records.filtering;
		const currentFilter = filtering.getFilter();
		const nextConditions = (currentFilter?.conditions ?? []).filter((condition) => condition.attributeName !== name);

		if (mode === "clear") {
			if (nextConditions.length === 0) {
				filtering.clearFilter();
			} else {
				filtering.setFilter({
					filterOperator: 0,
					conditions: nextConditions,
				});
			}
			this.context.parameters.records.refresh();
			return;
		}

		if (mode === "containsData" || mode === "doesNotContainData") {
			nextConditions.push({
				attributeName: name,
				conditionOperator: (mode === "containsData" ? 13 : 12) as ComponentFramework.PropertyHelper.DataSetApi.Types.ConditionOperator,
				value: "",
			});

			filtering.setFilter({
				filterOperator: 0,
				conditions: nextConditions,
			});
			this.context.parameters.records.refresh();
			return;
		}

		if ((typeof value === "string" && value.trim().length === 0) || value === undefined) {
			if (nextConditions.length === 0) {
				filtering.clearFilter();
			} else {
				filtering.setFilter({
					filterOperator: 0,
					conditions: nextConditions,
				});
			}
			this.context.parameters.records.refresh();
			return;
		}

		if (mode === "in") {
			if (!Array.isArray(value) || value.length === 0) {
				if (nextConditions.length === 0) {
					filtering.clearFilter();
				} else {
					filtering.setFilter({
						filterOperator: 0,
						conditions: nextConditions,
					});
				}
				this.context.parameters.records.refresh();
				return;
			}

			nextConditions.push({
				attributeName: name,
				conditionOperator: 8,
				value: value,
			});
		} else {
			const filterValue = Array.isArray(value) ? value[0] : value;
			if (!filterValue || filterValue.trim().length === 0) {
				if (nextConditions.length === 0) {
					filtering.clearFilter();
				} else {
					filtering.setFilter({
						filterOperator: 0,
						conditions: nextConditions,
					});
				}
				this.context.parameters.records.refresh();
				return;
			}

			let conditionOperator = 0;
			let mappedValue: string = filterValue;

			switch (mode) {
				case "contains":
					conditionOperator = 6;
					mappedValue = `%${filterValue}%`;
					break;
				case "notContains":
					conditionOperator = 7;
					mappedValue = `%${filterValue}%`;
					break;
				case "beginsWith":
					conditionOperator = 6;
					mappedValue = `${filterValue}%`;
					break;
				case "notBeginsWith":
					conditionOperator = 7;
					mappedValue = `${filterValue}%`;
					break;
				case "endsWith":
					conditionOperator = 6;
					mappedValue = `%${filterValue}`;
					break;
				case "notEndsWith":
					conditionOperator = 7;
					mappedValue = `%${filterValue}`;
					break;
				case "notEquals":
					conditionOperator = 1;
					mappedValue = filterValue;
					break;
				case "equals":
				default:
					conditionOperator = 0;
					mappedValue = filterValue;
					break;
			}

			nextConditions.push({
				attributeName: name,
				conditionOperator:
					conditionOperator as ComponentFramework.PropertyHelper.DataSetApi.Types.ConditionOperator,
				value: mappedValue,
			});
		}

		if (nextConditions.length > 0) {
			filtering.setFilter({
				filterOperator: 0,
				conditions: nextConditions,
			});
		} else {
			filtering.clearFilter();
		}

		this.context.parameters.records.refresh();
	};

	loadFirstPage = (): void => {
		this.currentPage = 1;
		this.context.parameters.records.paging.loadExactPage(1);
	};

	loadNextPage = (): void => {
		this.currentPage++;
		this.context.parameters.records.paging.loadExactPage(this.currentPage);
	};

	loadPreviousPage = (): void => {
		this.currentPage--;
		this.context.parameters.records.paging.loadExactPage(this.currentPage);
	};

	onFullScreen = (): void => {
		this.context.mode.setFullScreen(true);
	};

	/* eslint-disable @typescript-eslint/no-unsafe-assignment, @typescript-eslint/no-unsafe-call, @typescript-eslint/no-unsafe-member-access, @typescript-eslint/no-explicit-any */
	private getEntityMeta = async (entityType: string): Promise<{ entitySetName: string; primaryNameAttr: string }> => {
		if (this.entityMetaCache[entityType]) return this.entityMetaCache[entityType];
		const utils = this.context.utils as any;
		const meta = await utils.getEntityMetadata(entityType);
		const result = {
			entitySetName: String(meta?.EntitySetName ?? entityType + "s"),
			primaryNameAttr: String(meta?.PrimaryNameAttribute ?? "name"),
		};
		this.entityMetaCache[entityType] = result;
		return result;
	};
	/* eslint-enable @typescript-eslint/no-unsafe-assignment, @typescript-eslint/no-unsafe-call, @typescript-eslint/no-unsafe-member-access, @typescript-eslint/no-explicit-any */

	onSearchLookup = async (entityType: string, searchTerm: string): Promise<{ id: string; name: string }[]> => {
		try {
			const { primaryNameAttr } = await this.getEntityMeta(entityType);
			const safeSearch = searchTerm.trim().replace(/'/g, "''");
			const filter = safeSearch ? `&$filter=contains(${primaryNameAttr},'${safeSearch}')` : "";
			const query = `?$select=${primaryNameAttr}&$top=20&$orderby=${primaryNameAttr} asc${filter}`;
			const results = await this.context.webAPI.retrieveMultipleRecords(entityType, query);
			return results.entities.map((e) => {
				// eslint-disable-next-line @typescript-eslint/no-base-to-string
				const id = String((e as Record<string, unknown>)[`${entityType}id`] ?? (e as Record<string, unknown>).id ?? "");
				// eslint-disable-next-line @typescript-eslint/no-base-to-string
				const name = String((e as Record<string, unknown>)[primaryNameAttr] ?? "(unknown)");
				return { id, name };
			});
		} catch {
			return [];
		}
	};

	onUpdateLookupCell = async (
		recordId: string,
		columnName: string,
		targetId: string,
		targetEntityType: string
	): Promise<void> => {
		const entityType = this.context.parameters.records.getTargetEntityType();
		if (!entityType) throw new Error("Unable to determine target table for inline editing.");
		try {
			const { entitySetName } = await this.getEntityMeta(targetEntityType);
			const updateData: ComponentFramework.WebApi.Entity = {
				[`${columnName}@odata.bind`]: `/${entitySetName}(${this.normalizeGuid(targetId)})`,
			};
			await this.context.webAPI.updateRecord(entityType, this.normalizeGuid(recordId), updateData);
			this.context.parameters.records.refresh();
		} catch (error) {
			const message = error instanceof Error ? error.message : "Unable to save the lookup field.";
			await this.context.navigation.openAlertDialog({ text: message });
			throw error;
		}
	};

	private fetchColumnOptions = async (
		entityType: string,
		columns: ComponentFramework.PropertyHelper.DataSetApi.Column[]
	): Promise<void> => {
		const optionSetCols = columns.filter((col) => {
			const t = col.dataType.toLowerCase();
			return t.includes("optionset") && !t.includes("multiselectoptionset");
		});

		if (optionSetCols.length === 0) return;

		/* eslint-disable @typescript-eslint/no-unsafe-assignment, @typescript-eslint/no-unsafe-call, @typescript-eslint/no-unsafe-member-access */
		try {
			// EntityMetadata is typed as [key: string]: any; we use explicit any-casts throughout.
			// eslint-disable-next-line @typescript-eslint/no-explicit-any
			const utils = this.context.utils as any;
			const metadata = await utils.getEntityMetadata(entityType, optionSetCols.map((c) => c.name));

			let changed = false;
			optionSetCols.forEach((col) => {
				try {
					// eslint-disable-next-line @typescript-eslint/no-explicit-any
					const attrMeta: any =
						typeof metadata?.Attributes?.getByName === "function"
							? metadata.Attributes.getByName(col.name)
							: metadata?.[col.name];
					// eslint-disable-next-line @typescript-eslint/no-explicit-any
					const rawOptions: any[] = attrMeta?.OptionSet?.Options ?? attrMeta?.AttributeType?.Options ?? [];
					if (Array.isArray(rawOptions) && rawOptions.length > 0) {
						// eslint-disable-next-line @typescript-eslint/no-explicit-any
						this.columnOptions[col.name] = rawOptions.map((opt: any) => ({
							label: String(
								opt.Label?.UserLocalizedLabel?.Label ?? opt.Label ?? opt.DisplayName ?? opt.Value
							),
							value: String(opt.Value),
						}));
						changed = true;
					}
				} catch {
					// ignore per-column errors
				}
			});

			if (changed) {
				this.renderGrid();
			}
		} catch {
			// Metadata fetch is best-effort; fall back to text input for OptionSet columns.
		}
		/* eslint-enable @typescript-eslint/no-unsafe-assignment, @typescript-eslint/no-unsafe-call, @typescript-eslint/no-unsafe-member-access */
	};

	private renderGrid = (): void => {
		if (!this.container || !this.context) return;
		const dataset = this.context.parameters.records;
		const paging = dataset.paging;

		const rawAllocatedWidth = this.context.mode.allocatedWidth as unknown as string;
		const rawAllocatedHeight = this.context.mode.allocatedHeight as unknown as string;
		let allocatedWidth = parseInt(rawAllocatedWidth, 10);
		let allocatedHeight = parseInt(rawAllocatedHeight, 10);

		if (!Number.isFinite(allocatedWidth) || allocatedWidth <= 0) {
			allocatedWidth = -1;
		}
		if (!this.isFullScreen && this.context.parameters.SubGridHeight.raw) {
			allocatedHeight = this.context.parameters.SubGridHeight.raw;
		}
		if (!Number.isFinite(allocatedHeight) || allocatedHeight <= 0) {
			allocatedHeight = 420;
		}

		ReactDOM.render(
			React.createElement(Grid, {
				width: allocatedWidth,
				height: allocatedHeight,
				columns: dataset.columns,
				records: this.records ?? {},
				sortedRecordIds: this.sortedRecordsIds ?? [],
				hasNextPage: paging.hasNextPage,
				hasPreviousPage: paging.hasPreviousPage,
				currentPage: this.currentPage,
				totalResultCount: paging.totalResultCount,
				sorting: dataset.sorting,
				filtering: dataset.filtering?.getFilter(),
				resources: this.resources,
				itemsLoading: dataset.loading,
				highlightValue: this.context.parameters.HighlightValue.raw,
				highlightColor: this.context.parameters.HighlightColor.raw,
				enableLookupLinks: this.context.parameters.EnableLookupLinks.raw ?? true,
				enableInlineEdit: true,
				columnOptions: this.columnOptions,
				onSearchLookup: this.onSearchLookup,
				onUpdateLookupCell: this.onUpdateLookupCell,
				setSelectedRecords: this.setSelectedRecords,
				onNavigate: this.onNavigate,
				onOpenLookup: this.onOpenLookup,
				onSort: this.onSort,
				onFilter: this.onFilter,
				loadFirstPage: this.loadFirstPage,
				loadNextPage: this.loadNextPage,
				loadPreviousPage: this.loadPreviousPage,
				onUpdateCell: this.onUpdateCell,
				isFullScreen: this.isFullScreen,
				onFullScreen: this.onFullScreen,
			}),
			this.container
		);
	};

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(
		context: ComponentFramework.Context<IInputs>,
		notifyOutputChanged: () => void,
		state: ComponentFramework.Dictionary,
		container: HTMLDivElement
	): void {
		this.notifyOutputChanged = notifyOutputChanged;
		this.container = container;
		this.context = context;
		this.context.mode.trackContainerResize(true);
		this.resources = this.context.resources;
		this.isTestHarness = document.getElementById("control-dimensions") !== null;
	}

	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void {
		const dataset = context.parameters.records;

		// In MDAs, the initial population of the dataset does not provide updatedProperties.
		// Ensure we always hydrate local record caches at least once.
		const initialLoad = this.records === undefined;
		const datasetChanged =
			context.updatedProperties.includes("dataset") || context.updatedProperties.includes("records") || initialLoad;
		const resetPaging = datasetChanged && !dataset.loading && !dataset.paging.hasPreviousPage && this.currentPage !== 1;

		if (context.updatedProperties.includes("fullscreen_close")) {
			this.isFullScreen = false;
		}
		if (context.updatedProperties.includes("fullscreen_open")) {
			this.isFullScreen = true;
		}

		if (resetPaging) {
			this.currentPage = 1;
		}

		// Always keep local caches in sync with the latest dataset payload.
		// Some MDA update cycles don't flag updatedProperties as expected.
		this.records = dataset.records ?? {};
		this.sortedRecordsIds = dataset.sortedRecordIds ?? [];

		// Kick off a one-time metadata fetch for OptionSet columns whenever the
		// entity type changes (e.g. on first load or navigation to a different table).
		const entityType = dataset.getTargetEntityType?.();
		if (entityType && entityType !== this.columnOptionsFetchedFor) {
			this.columnOptionsFetchedFor = entityType;
			void this.fetchColumnOptions(entityType, dataset.columns);
		}

		this.renderGrid();
	}

	/**
	 * It is called by the framework prior to a control receiving new data.
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs {
		return {} as IOutputs;
	}

	/**
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void {
		ReactDOM.unmountComponentAtNode(this.container);
	}
}
