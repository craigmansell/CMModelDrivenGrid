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
	// Maps source-entity column logical name → Dataverse navigation property name.
	// Populated eagerly from full entity metadata at updateView time.
	private lookupNavPropertyMap: Record<string, string> = {};
	private lookupNavPropertyFetchedFor: string | undefined;
	// Stored promise so onUpdateLookupCell can await map population before saving.
	private lookupNavPropertyFetchPromise: Promise<void> | undefined;

	setSelectedRecords = (ids: string[]): void => {
		this.context.parameters.records.setSelectedRecordIds(ids);
	};

	onNavigate = (item?: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord): void => {
		if (item) {
			this.context.parameters.records.openDatasetItem(item.getNamedReference());
		}
	};

	private normalizeGuid = (id: string): string => id.replace(/[{}]/g, "");

	private isInvalidPropertyPayloadError = (error: unknown): boolean => {
		const message = error instanceof Error ? error.message.toLowerCase() : String(error).toLowerCase();
		return (
			message.includes("invalid property") ||
			message.includes("undeclared property") ||
			message.includes("does not exist on type")
		);
	};

	private isAliasedDatasetColumn = (column: ComponentFramework.PropertyHelper.DataSetApi.Column): boolean => {
		const alias = (column as { alias?: string }).alias;
		return typeof alias === "string" && alias.length > 0 && alias !== column.name;
	};

	private getEditableDatasetColumn = (columnName: string): ComponentFramework.PropertyHelper.DataSetApi.Column => {
		const column = this.context.parameters.records.columns.find((c) => c.name === columnName);
		if (!column) {
			throw new Error(`Column '${columnName}' is not available for inline editing.`);
		}
		if (this.isAliasedDatasetColumn(column)) {
			throw new Error("This column comes from a related table/view and cannot be edited inline.");
		}
		return column;
	};

	private resolveLookupBindNames = async (
		entityType: string,
		columnName: string,
		targetEntityType?: string
	): Promise<string[]> => {
		const bindCandidates = new Set<string>();
		const normalizedTarget = typeof targetEntityType === "string" ? targetEntityType.trim() : "";
		const normalizedTargetLower = normalizedTarget.toLowerCase();

		const addCandidate = (candidate: unknown): void => {
			if (typeof candidate !== "string") {
				return;
			}
			const trimmed = candidate.trim();
			if (trimmed.length > 0) {
				bindCandidates.add(trimmed);
			}
		};

		const addPolymorphicCandidate = (baseName: string): void => {
			if (!normalizedTarget) {
				return;
			}
			const trimmedBase = baseName.trim();
			if (trimmedBase.length === 0) {
				return;
			}
			const lowerBase = trimmedBase.toLowerCase();
			if (lowerBase.endsWith(`_${normalizedTargetLower}`)) {
				return;
			}
			addCandidate(`${trimmedBase}_${normalizedTarget}`);
			if (normalizedTargetLower !== normalizedTarget) {
				addCandidate(`${trimmedBase}_${normalizedTargetLower}`);
			}
		};

		const baseCandidates = [columnName];
		if (!columnName.endsWith("id")) {
			baseCandidates.push(`${columnName}id`);
		} else if (columnName.length > 2) {
			baseCandidates.push(columnName.slice(0, -2));
		}

		/* eslint-disable @typescript-eslint/no-unsafe-assignment, @typescript-eslint/no-unsafe-call, @typescript-eslint/no-unsafe-member-access, @typescript-eslint/no-explicit-any */
		try {
			const utils = this.context.utils as any;
			const metadata = await utils.getEntityMetadata(entityType, [columnName]);
			const attributeMeta =
				typeof metadata?.Attributes?.getByName === "function"
					? metadata.Attributes.getByName(columnName)
					: metadata?.[columnName];

			addCandidate(attributeMeta?.NavigationPropertyName);
			addCandidate(attributeMeta?.SchemaName);
			addCandidate(attributeMeta?.LogicalName);
			addCandidate(attributeMeta?.ODataName);

			const relationships: any[] = Array.isArray(metadata?.ManyToOneRelationships)
				? metadata.ManyToOneRelationships
				: [];

			const matchingRelationships = relationships.filter((relationship) => {
				const referencingAttribute = relationship?.ReferencingAttribute;
				if (typeof referencingAttribute !== "string") {
					return false;
				}
				return referencingAttribute.toLowerCase() === columnName.toLowerCase();
			});

			const getNavName = (relationship: any): unknown =>
				relationship?.ReferencingEntityNavigationPropertyName ?? relationship?.NavigationPropertyName;

			const getReferencedEntity = (relationship: any): string => {
				const candidate = relationship?.ReferencedEntity ?? relationship?.ReferencedEntityName;
				return typeof candidate === "string" ? candidate : "";
			};

			if (normalizedTarget) {
				matchingRelationships
					.filter((relationship) => getReferencedEntity(relationship).toLowerCase() === normalizedTargetLower)
					.forEach((relationship) => addCandidate(getNavName(relationship)));
			}

			matchingRelationships.forEach((relationship) => addCandidate(getNavName(relationship)));
		} catch {
			// Ignore metadata resolution failures and rely on fallback candidates.
		}
		/* eslint-enable @typescript-eslint/no-unsafe-assignment, @typescript-eslint/no-unsafe-call, @typescript-eslint/no-unsafe-member-access, @typescript-eslint/no-explicit-any */

		// Fall back to common naming conventions.
		baseCandidates.forEach(addCandidate);

		// Polymorphic lookups (customer/regarding/custom polymorphic) require a target-type-suffixed bind property.
		baseCandidates.forEach((candidate) => addPolymorphicCandidate(candidate));

		return Array.from(bindCandidates);
	};

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

			const editableColumn = this.getEditableDatasetColumn(columnName);
			const parsedValue = this.parseInlineEditValue(value, editableColumn.dataType || dataType);
			const updateData: ComponentFramework.WebApi.Entity = {
				[editableColumn.name]: parsedValue,
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

		// For lookup columns, conditions are stored under the "{name}name" companion attribute
		// (the display-name virtual field), so we must clear both variants.
		const columnMeta = this.context.parameters.records.columns.find((c) => c.name === name);
		const isLookup = columnMeta ? /lookup|owner/i.test(columnMeta.dataType) : false;
		const filterAttributeName = isLookup ? `${name}name` : name;

		const nextConditions = (currentFilter?.conditions ?? []).filter(
			(condition) => condition.attributeName !== name && condition.attributeName !== `${name}name`
		);

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
				attributeName: filterAttributeName,
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

	onClearAllFilters = (): void => {
		this.context.parameters.records.filtering.clearFilter();
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

		// For custom entities (publisher prefix like "cm_"), Dataverse always names the
		// primary name field "${prefix}_name" — e.g. cm_componentoption → cm_name.
		const publisherPrefix = entityType.includes("_") ? entityType.split("_")[0] : "";
		const conventionName = publisherPrefix ? `${publisherPrefix}_name` : "name";

		try {
			const utils = this.context.utils as any;
			const meta = await utils.getEntityMetadata(entityType);
			const metaEntitySet = typeof meta?.EntitySetName === "string" ? meta.EntitySetName : "";
			const metaPrimaryName = typeof meta?.PrimaryNameAttribute === "string" ? meta.PrimaryNameAttribute : "";
			const result = {
				entitySetName: metaEntitySet !== "" ? metaEntitySet : `${entityType}s`,
				// If metadata returns "name" for a custom entity, the convention is more reliable.
				primaryNameAttr: metaPrimaryName && !(metaPrimaryName === "name" && publisherPrefix) ? metaPrimaryName : conventionName,
			};
			this.entityMetaCache[entityType] = result;
			return result;
		} catch {
			// Fall back to naming conventions if metadata fetch fails.
			const fallback = { entitySetName: `${entityType}s`, primaryNameAttr: conventionName };
			this.entityMetaCache[entityType] = fallback;
			return fallback;
		}
	};
	/* eslint-enable @typescript-eslint/no-unsafe-assignment, @typescript-eslint/no-unsafe-call, @typescript-eslint/no-unsafe-member-access, @typescript-eslint/no-explicit-any */

	onSearchLookup = async (entityType: string, searchTerm: string): Promise<{ id: string; name: string }[]> => {
		const { primaryNameAttr } = await this.getEntityMeta(entityType);
		const safeSearch = (searchTerm ?? "").trim().replace(/'/g, "''");
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
			const editableColumn = this.getEditableDatasetColumn(columnName);

			// Ensure the nav property map is fully populated before building candidates.
			if (this.lookupNavPropertyFetchPromise) {
				await this.lookupNavPropertyFetchPromise;
			}

			// Use the eagerly-fetched nav property map first (most reliable).
			// Fall back to the runtime metadata resolver for any remaining candidates.
			const prebuiltNavProp = this.lookupNavPropertyMap[editableColumn.name.toLowerCase()];
			const isClear = targetId === "";
			const entitySetName = isClear ? "" : (await this.getEntityMeta(targetEntityType)).entitySetName;
			const runtimeCandidates = await this.resolveLookupBindNames(entityType, editableColumn.name, targetEntityType);
			const bindCandidates = prebuiltNavProp
				? [prebuiltNavProp, ...runtimeCandidates.filter((c) => c !== prebuiltNavProp)]
				: runtimeCandidates;

			console.info(
				`[CMModelDrivenGrid] onUpdateLookupCell: column=${columnName}, prebuiltNavProp=${prebuiltNavProp ?? "(none)"},`,
				`candidates=[${bindCandidates.join(", ")}], entitySetName=${entitySetName}, targetId=${targetId}`
			);
			console.info(`[CMModelDrivenGrid] lookupNavPropertyMap:`, JSON.stringify(this.lookupNavPropertyMap));

			let lastError: unknown;
			for (const bindName of bindCandidates) {
				try {
					const updateData: ComponentFramework.WebApi.Entity = isClear
						? { [`${bindName}@odata.bind`]: null }
						: { [`${bindName}@odata.bind`]: `/${entitySetName}(${this.normalizeGuid(targetId)})` };
					console.info(`[CMModelDrivenGrid] Trying bind: ${bindName}@odata.bind (clear=${isClear})`);
					await this.context.webAPI.updateRecord(entityType, this.normalizeGuid(recordId), updateData);
					console.info(`[CMModelDrivenGrid] SUCCESS with bind: ${bindName}@odata.bind`);
					this.context.parameters.records.refresh();
					return;
				} catch (error) {
					console.info(`[CMModelDrivenGrid] FAILED with bind: ${bindName}@odata.bind`, error);
					lastError = error;
					// Always try the next candidate — PCF wraps OData errors in opaque
					// objects so we cannot reliably detect "undeclared property" from
					// the error message alone.
				}
			}

			if (lastError instanceof Error) {
				throw lastError;
			}
			throw new Error("Unable to save the lookup field.");
		} catch (error) {
			const message = error instanceof Error ? error.message : "Unable to save the lookup field.";
			await this.context.navigation.openAlertDialog({ text: message });
			throw error;
		}
	};

	private fetchLookupNavProperties = async (
		entityType: string,
		columns: ComponentFramework.PropertyHelper.DataSetApi.Column[]
	): Promise<void> => {
		const lookupCols = columns.filter((col) => {
			const t = col.dataType.toLowerCase();
			return t.includes("lookup") || t.includes("owner");
		});
		if (lookupCols.length === 0) return;

		// PCF's getEntityMetadata does NOT return ManyToOneRelationships.
		// Query the Dataverse Web API metadata endpoint directly instead.
		try {
			const baseUrl = window.location.origin;
			const url =
				`${baseUrl}/api/data/v9.0/EntityDefinitions(LogicalName='${entityType}')/ManyToOneRelationships` +
				`?$select=ReferencingAttribute,ReferencingEntityNavigationPropertyName,ReferencedEntity`;

			const response = await fetch(url, {
				headers: {
					Accept: "application/json",
					"OData-MaxVersion": "4.0",
					"OData-Version": "4.0",
				},
				credentials: "include",
			});

			if (!response.ok) {
				console.info(`[CMModelDrivenGrid] fetchLookupNavProperties: HTTP ${response.status} for ${entityType}`);
				return;
			}

			const data = await response.json() as {
				value: {
					ReferencingAttribute: string;
					ReferencingEntityNavigationPropertyName: string;
					ReferencedEntity: string;
				}[];
			};

			const lookupColNames = new Set(lookupCols.map((c) => c.name.toLowerCase()));

			console.info(
				`[CMModelDrivenGrid] fetchLookupNavProperties: entityType=${entityType},`,
				`relationship count=${data.value?.length ?? 0},`,
				`lookup columns=${lookupCols.map((c) => c.name).join(", ")}`
			);

			(data.value ?? []).forEach((rel) => {
				const attr = (rel.ReferencingAttribute ?? "").toLowerCase();
				const navProp = rel.ReferencingEntityNavigationPropertyName ?? "";
				if (attr && navProp && lookupColNames.has(attr)) {
					if (!this.lookupNavPropertyMap[attr]) {
						this.lookupNavPropertyMap[attr] = navProp;
						console.info(`[CMModelDrivenGrid] Mapped lookup: ${attr} → nav property: ${navProp}`);
					}
				}
			});

			console.info(`[CMModelDrivenGrid] lookupNavPropertyMap after fetch:`, JSON.stringify(this.lookupNavPropertyMap));
		} catch (err) {
			console.info(`[CMModelDrivenGrid] fetchLookupNavProperties error:`, err);
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
				onClearAllFilters: this.onClearAllFilters,
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
		if (entityType && entityType !== this.lookupNavPropertyFetchedFor) {
			this.lookupNavPropertyFetchedFor = entityType;
			this.lookupNavPropertyMap = {};
			this.lookupNavPropertyFetchPromise = this.fetchLookupNavProperties(entityType, dataset.columns);
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
