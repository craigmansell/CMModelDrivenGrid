import { initializeIcons } from "@fluentui/react/lib/Icons";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { HierarchyGrid as Grid } from "./HierarchyGrid";
import { createChildLoader, createChildAvailabilityLoader, ChildRowsConfig, ChildRequest, childGuid, childLogicalName } from "./childRecords";
import { createDataverseChildPaging } from "./dataverseChildPaging";
import { FilterExpression, FilterMode, FilterValue, updateColumnFilter } from "./filtering";
import { needsServerFilter, supportedFilterSuperset, collectFilterCandidates, resolveFilterIds, filterForIds } from './datasetFiltering';
import { NativeHierarchySelection, isHierarchyRoot } from "./nativeHierarchy";
import { MAX_SELECTED_RECORDS, SelectAllRequest, SelectAllResult, selectionResult, checkSelectionCancelled } from './recordSelection';

interface LookupFilterConfig {
	default?: string;
	byEntity?: Record<string, string>;
	byColumn?: Record<string, string>;
	byColumnAndEntity?: Record<string, string>;
}

interface EntityMetadata {
	entitySetName: string;
	primaryIdAttr: string;
	primaryNameAttr: string;
	lookupFields: Map<string, string>;
}

interface PageLoadWaiter {
	baseline: string;
	timerId: number;
	settled: boolean;
	resolve: () => void;
	reject: (error: Error) => void;
}

interface PageLoadWaitHandle {
	promise: Promise<void>;
	reject: (error: Error) => void;
}

const CONTROL_DESTROYED_MESSAGE = "The control was destroyed while loading a dataset page.";

function copyFilter(filter: FilterExpression | undefined): FilterExpression | undefined {
	return filter ? {...filter, conditions: filter.conditions.map(condition => ({...condition,
		value: Array.isArray(condition.value) ? [...condition.value] : condition.value})),
		...(filter.filters ? {filters: filter.filters.map(child => copyFilter(child)!)} : {})} : undefined;
}

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
	accumulatedRecords: Record<string, ComponentFramework.PropertyHelper.DataSetApi.EntityRecord> = {};
	accumulatedSortedIds: string[] = [];
	currentPage = 1;
	columnOptions: Record<string, { label: string; value: string; color?: string }[]> = {};
	private columnOptionsFetchedFor: string | undefined;
	private entityMetaCache: Record<string, EntityMetadata> = {};
	// Maps source-entity column logical name → Dataverse navigation property name.
	// Populated eagerly from full entity metadata at updateView time.
	private lookupNavPropertyMap: Record<string, string> = {};
	// Maps source-entity column logical name → target entity logical name (from ManyToOneRelationships).
	private lookupTargetEntityMap: Record<string, string> = {};
	private lookupNavPropertyFetchedFor: string | undefined;
	// Stored promise so onUpdateLookupCell can await map population before saving.
	private lookupNavPropertyFetchPromise: Promise<void> | undefined;
	private appliedPageSize: number | undefined;
	private lookupTokenGuidCache: Record<string, string | null> = {};
	private lookupTokenGuidByNameCache: Record<string, string | null> = {};
	private visibilityChangeHandler: (() => void) | undefined;
	private windowFocusHandler: (() => void) | undefined;
	private tabRefreshTimeoutId: number | undefined;
	private lastTabRefreshAt = 0;
	private lastControlVisible: boolean | undefined;
	private lastVisibilityRefreshAt = 0;
	private pageLoadWaiters = new Set<PageLoadWaiter>();
	private nextPageRequest: Promise<void> | undefined;
	private appendingPage = false;
	private selectAllAcrossPagesInProgress = false;
	private selectionRevision = 0;
	private nativeHierarchy: NativeHierarchySelection | undefined;
	private nativeHierarchyKey = '';
	private nativeHierarchyInitializing = false;
	private nativeSelectionReset = 0;
	private nativeSelectionError = '';
	private isDestroyed = false;
	private logicalRootFilter: FilterExpression | undefined;
	private ownsRootFilter = false;
	private filterViewKey = '';
	private rootFilterBusy = false;
	private rootFilterError = '';
	private queryErrorVersion = 0;
	private rootFilterRecoveryRequired = false;
	private pendingRootFilter: {filter: FilterExpression | undefined; sorting: ComponentFramework.PropertyHelper.DataSetApi.SortStatus[]} | undefined;
	private getRootFilter = (): FilterExpression | undefined => this.ownsRootFilter ? this.logicalRootFilter
		: this.nativeHierarchy ? this.nativeHierarchy.rootFilter : this.context.parameters.records.filtering.getFilter();

	private refreshPhysicalFilter = async (filter: FilterExpression | undefined, sorting: ComponentFramework.PropertyHelper.DataSetApi.SortStatus[]): Promise<void> => {
		const dataset=this.context.parameters.records;
		if (this.nativeHierarchy) {
			this.nativeSelectionError='';
			await this.nativeHierarchy.refreshRoots(filter,sorting);
			if (this.nativeSelectionError) throw new Error(this.nativeSelectionError);
		}
		else {
			dataset.sorting.splice(0,dataset.sorting.length,...sorting);
			if (filter) dataset.filtering.setFilter(filter); else dataset.filtering.clearFilter();
			this.currentPage=1;
			await this.loadDatasetPageAndWait(() => dataset.refresh());
		}
		if (dataset.error) throw new Error('Power Apps could not apply this filter.');
	};

	private applyRootFilter = (filter: FilterExpression | undefined, sorting = this.nativeHierarchy?.rootSorting ?? this.context.parameters.records.sorting): void => {
		const nextSorting=sorting.map(sort => ({...sort}));
		if (this.rootFilterBusy) { this.pendingRootFilter={filter,sorting:nextSorting}; return; }
		this.rootFilterError='';
		if (needsServerFilter(filter)) { void this.applyServerRootFilter(filter!,nextSorting); return; }
		const dataset=this.context.parameters.records;
		const previousPhysical=copyFilter(dataset.filtering.getFilter());
		const previousLogical=copyFilter(this.getRootFilter());
		const previousSorting=dataset.sorting.map(sort => ({...sort}));
		const native=this.nativeHierarchy;
		const previousNativeFilter=copyFilter(native?.rootFilter);
		const previousNativeSorting=native?.rootSorting.map(sort => ({...sort}));
		try {
			const physical=supportedFilterSuperset(filter,this.context.parameters.records.columns);
			this.logicalRootFilter=filter; this.ownsRootFilter=true;
			if (this.nativeHierarchy) void this.nativeHierarchy.refreshRoots(physical,nextSorting);
			else {
				dataset.sorting.splice(0,dataset.sorting.length,...nextSorting);
				if (physical) dataset.filtering.setFilter(physical); else dataset.filtering.clearFilter();
				dataset.refresh();
			}
			this.rootFilterRecoveryRequired=false;
		} catch (error) {
			this.logicalRootFilter=previousLogical; this.ownsRootFilter=true;
			this.rootFilterError=error instanceof Error ? error.message : 'Unable to apply the filter.';
			// Sorting can clear the host filter before setFilter/refresh rejects.
			// Restore the whole query, not just the text shown in the filter chip.
			try {
				if (native) {native.rootFilter=previousNativeFilter; native.rootSorting=previousNativeSorting!;}
				dataset.sorting.splice(0,dataset.sorting.length,...previousSorting);
				if (previousPhysical) dataset.filtering.setFilter(previousPhysical); else dataset.filtering.clearFilter();
			} catch {
				this.rootFilterRecoveryRequired=true;
				this.rootFilterError+=' Refresh the subgrid before selecting records; the previous query could not be restored.';
				dataset.setSelectedRecordIds([]); this.nativeSelectionReset++;
			}
			// An identical error on a second attempt still needs a completion event.
			this.queryErrorVersion++;
			this.renderGrid();
		}
	};

	private applyServerRootFilter = async (filter: FilterExpression, sorting: ComponentFramework.PropertyHelper.DataSetApi.SortStatus[]): Promise<void> => {
		const dataset=this.context.parameters.records;
		const entity=dataset.getTargetEntityType();
		const viewId=dataset.getViewId();
		const native=this.nativeHierarchy;
		const previousPhysical=native ? native.rootFilter : dataset.filtering.getFilter();
		const previousSorting=(native?.rootSorting ?? dataset.sorting).map(sort => ({...sort}));
		const previousLogical=this.getRootFilter();
		const assertCurrent=() => {
			if (this.isDestroyed || this.context.parameters.records.getTargetEntityType() !== entity || this.context.parameters.records.getViewId() !== viewId || this.nativeHierarchy !== native) {
				throw new Error('The grid changed while checking this filter. Apply it again in the current view.');
			}
		};
		this.logicalRootFilter=filter; this.ownsRootFilter=true;
		this.rootFilterBusy=true; this.nativeSelectionReset++;
		dataset.setSelectedRecordIds([]); this.renderGrid();
		try {
			const meta=await this.getEntityMeta(entity);
			const view=await this.context.webAPI.retrieveRecord('savedquery',viewId.replace(/[{}]/g,''),'?$select=fetchxml,returnedtypecode');
			assertCurrent();
			if (view.returnedtypecode !== entity) throw new Error('The selected view belongs to another table.');
			await this.refreshPhysicalFilter(supportedFilterSuperset(filter,dataset.columns),sorting);
			assertCurrent();
			const candidates=await collectFilterCandidates({assertCurrent,
				page:() => {
					const current=this.context.parameters.records;
					return {ids:current.sortedRecordIds.map(id => current.records[id]).filter(row => !!row && (!native || isHierarchyRoot(row,native.parentField))).map(row => row.getRecordId()),hasNextPage:current.paging.hasNextPage,error:current.error};
				},
				next:() => this.loadNextPage(),
			});
			const matches=await resolveFilterIds(this.context.webAPI,String(view.fetchxml),entity,meta.primaryIdAttr,filter,candidates,assertCurrent);
			assertCurrent();
			await this.refreshPhysicalFilter(filterForIds(meta.primaryIdAttr,matches),sorting);
			assertCurrent();
			this.logicalRootFilter=filter; this.ownsRootFilter=true;
			this.rootFilterRecoveryRequired=false;
		} catch (error) {
			if (!this.isDestroyed && this.nativeHierarchy === native && this.context.parameters.records.getViewId() === viewId) {
				this.logicalRootFilter=previousLogical; this.ownsRootFilter=true;
				this.rootFilterError=error instanceof Error ? error.message : 'Unable to check the text filter.';
				try { await this.refreshPhysicalFilter(previousPhysical,previousSorting); }
				catch { this.rootFilterError+=' Refresh the subgrid before selecting records.'; }
			}
		} finally {
			this.rootFilterBusy=false;
			if (!this.isDestroyed) {
				this.childRevision++; this.renderGrid();
				const pending=this.pendingRootFilter; this.pendingRootFilter=undefined;
				if (pending) this.applyRootFilter(pending.filter,pending.sorting);
			}
		}
	};
	private childLoaderKey = '';
	private childRecords: Record<string, ComponentFramework.PropertyHelper.DataSetApi.EntityRecord> = {};
	private childLoader: ChildRowsConfig['loadPage'] | undefined;
	private childAvailabilityKey = '';
	private childAvailabilityLoader: ChildRowsConfig['hasChildren'];
	private checkChildRows = (parentIds: string[]): Promise<Record<string,boolean>> => {
		const entity = this.context.parameters.records.getTargetEntityType();
		const parent = this.context.parameters.ChildParentLookup?.raw?.trim() ?? '';
		const view = this.context.parameters.ChildViewId?.raw?.trim() ?? '';
		const key = `${entity}|${parent}|${view}`;
		if (!this.childAvailabilityLoader || key !== this.childAvailabilityKey) {
			this.childAvailabilityLoader = createChildAvailabilityLoader(this.context.webAPI,entity,parent,view);
			this.childAvailabilityKey = key;
		}
		return this.childAvailabilityLoader(parentIds);
	};
	private childRevision = 0;
	private childDatasetWasLoading = false;
	private loadChildPage = async (request: ChildRequest) => {
		const entity = this.context.parameters.records.getTargetEntityType();
		const parentField = this.context.parameters.ChildParentLookup?.raw?.trim() ?? '';
		const view = this.context.parameters.ChildViewId?.raw?.trim() ?? '';
		const key = `${entity}|${parentField}|${view}`;
		if (!this.childLoader || key !== this.childLoaderKey) {
			const meta = await this.getEntityMeta(entity);
			if (this.isDestroyed) throw new Error(CONTROL_DESTROYED_MESSAGE);
			this.childRecords = {};
			this.childLoader = createChildLoader(this.context.webAPI, entity, meta.primaryIdAttr, meta.primaryNameAttr, parentField, view, createDataverseChildPaging(meta.entitySetName));
			this.childLoaderKey = key;
		}
		const page = await this.childLoader(request);
		if (!this.isDestroyed && key === this.childLoaderKey) {
			for (const row of page.records) this.childRecords[row.getRecordId()] = row;
		}
		return page;
	};

	private toODataLiteral = (value: unknown): string => {
		if (value === null || value === undefined) return "null";
		if (typeof value === "number") return Number.isFinite(value) ? value.toString() : "null";
		if (typeof value === "boolean") return value ? "true" : "false";
		if (value instanceof Date) return `'${value.toISOString()}'`;
		if (typeof value === "string") {
			const trimmed = value.trim();
			if (/^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[1-5][0-9a-fA-F]{3}-[89abAB][0-9a-fA-F]{3}-[0-9a-fA-F]{12}$/.test(trimmed)) {
				return trimmed;
			}
			return `'${trimmed.replace(/'/g, "''")}'`;
		}
		return "null";
	};

	private extractLookupTokenValue = (raw: unknown, mode?: string): unknown => {
		const suffix = (mode ?? "").toLowerCase();
		const guidRegex = /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[1-5][0-9a-fA-F]{3}-[89abAB][0-9a-fA-F]{3}-[0-9a-fA-F]{12}$/;
		const guidInTextRegex = /[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[1-5][0-9a-fA-F]{3}-[89abAB][0-9a-fA-F]{3}-[0-9a-fA-F]{12}/;
		const toGuidIfPossible = (value: unknown, depth = 0): string | undefined => {
			if (depth > 2 || value === null || value === undefined) return undefined;
			if (typeof value === "string") {
				const trimmed = value.trim();
				const cleaned = trimmed.replace(/[{}]/g, "");
				if (guidRegex.test(cleaned)) return this.normalizeGuid(cleaned);
				const embedded = guidInTextRegex.exec(cleaned)?.[0];
				if (embedded) return this.normalizeGuid(embedded);
				return undefined;
			}
			if (Array.isArray(value)) {
				for (const item of value) {
					const found = toGuidIfPossible(item, depth + 1);
					if (found) return found;
				}
				return undefined;
			}
			if (typeof value === "object") {
				const rec = value as Record<string, unknown>;

				try {
					const toStringFn = (value as { toString?: () => string }).toString;
					if (typeof toStringFn === "function" && toStringFn !== Object.prototype.toString) {
						const asString = toStringFn.call(value);
						const foundInString = toGuidIfPossible(asString, depth + 1);
						if (foundInString) return foundInString;
					}
				} catch {
					// Ignore string conversion failures.
				}

				try {
					const valueOfResult = (value as { valueOf?: () => unknown }).valueOf?.();
					const foundInValueOf = toGuidIfPossible(valueOfResult, depth + 1);
					if (foundInValueOf) return foundInValueOf;
				} catch {
					// Ignore valueOf failures.
				}

				for (const key of ["guid", "Guid", "id", "Id", "value", "Value", "key", "Key"]) {
					const found = toGuidIfPossible(rec[key], depth + 1);
					if (found) return found;
				}

				try {
					const ownNames = Object.getOwnPropertyNames(value);
					for (const name of ownNames) {
						const found = toGuidIfPossible(rec[name], depth + 1);
						if (found) return found;
					}
				} catch {
					// Ignore non-enumerable property access failures.
				}

				for (const nested of Object.values(rec)) {
					const found = toGuidIfPossible(nested, depth + 1);
					if (found) return found;
				}
			}
			return undefined;
		};
		const pickString = (rec: Record<string, unknown>, keys: string[]): string | undefined => {
			for (const key of keys) {
				const value = rec[key];
				if (typeof value === "string" && value.length > 0) return value;
			}
			return undefined;
		};

		if (typeof raw === "string") {
			if (suffix === "name") return raw;
			if (suffix === "entitytype") return null;
			if (suffix === "id" || !suffix) {
				const normalized = toGuidIfPossible(raw);
				if (normalized) return normalized;
				return suffix === "id" ? null : raw;
			}
		}

		if (Array.isArray(raw) && raw.length > 0) {
			const first = raw[0] as unknown;
			if (first && typeof first === "object") {
				const rec = first as Record<string, unknown>;
				const id = toGuidIfPossible(rec);
				const name = pickString(rec, ["name", "Name"]);
				const entityType = pickString(rec, ["entityType", "entitytype", "EntityType", "logicalName", "LogicalName", "etn", "Etn"]);
				if (suffix === "name") return name ?? null;
				if (suffix === "entitytype") return entityType ?? null;
				if (suffix === "id" || !suffix) return id ?? null;
			}
			return null;
		}

		if (raw && typeof raw === "object") {
			const rec = raw as Record<string, unknown>;
			const id = toGuidIfPossible(rec);
			const name = pickString(rec, ["name", "Name"]);
			const entityType = pickString(rec, ["entityType", "entitytype", "EntityType", "logicalName", "LogicalName", "etn", "Etn"]);
			if (suffix === "name") return name ?? null;
			if (suffix === "entitytype") return entityType ?? null;
			if (suffix === "id" || !suffix) return id ?? null;
			return null;
		}

		if (suffix) {
			return null;
		}

		return raw;
	};

	private getLookupTokenGuidFromSourceRow = async (recordId: string, attributeName: string): Promise<string | null> => {
		const sourceEntityType = this.context.parameters.records.getTargetEntityType();
		if (!sourceEntityType) return null;

		const normalizedRecordId = this.normalizeGuid(recordId);
		const attribute = attributeName.toLowerCase();
		const cacheKey = `${sourceEntityType}|${normalizedRecordId}|${attribute}`;
		if (cacheKey in this.lookupTokenGuidCache) return this.lookupTokenGuidCache[cacheKey];

		const lookupIdAttribute = `_${attribute}_value`;
		try {
			const rowResult = await this.context.webAPI.retrieveRecord(
				sourceEntityType,
				normalizedRecordId,
				`?$select=${lookupIdAttribute}`
			);
			const idRaw = (rowResult as Record<string, unknown>)[lookupIdAttribute];
			const guid = typeof idRaw === "string" && idRaw.trim().length > 0 ? this.normalizeGuid(idRaw) : null;
			this.lookupTokenGuidCache[cacheKey] = guid;
			return guid;
		} catch {
			this.lookupTokenGuidCache[cacheKey] = null;
			return null;
		}
	};

	private getLookupTokenGuidFromLookupObject = async (raw: unknown): Promise<string | null> => {
		if (!raw || typeof raw !== "object") return null;
		const rec = raw as Record<string, unknown>;
		const lookupEntityType = typeof rec.etn === "string"
			? rec.etn
			: typeof rec.entityType === "string"
				? rec.entityType
				: typeof rec.logicalName === "string"
					? rec.logicalName
					: "";
		const lookupName = typeof rec.name === "string"
			? rec.name
			: typeof rec.Name === "string"
				? rec.Name
				: "";

		if (!lookupEntityType || !lookupName) return null;
		const cacheKey = `${lookupEntityType}|${lookupName.toLowerCase()}`;
		if (cacheKey in this.lookupTokenGuidByNameCache) return this.lookupTokenGuidByNameCache[cacheKey];

		try {
			const { primaryIdAttr, primaryNameAttr } = await this.getEntityMeta(lookupEntityType);
			const escapedName = lookupName.replace(/'/g, "''");
			// Names are not unique identifiers. Fetch enough to detect ambiguity,
			// and keep URL delimiters in a name from becoming extra query options.
			const query = `?$select=${primaryIdAttr},${primaryNameAttr}&$top=2&$filter=${encodeURIComponent(`${primaryNameAttr} eq '${escapedName}'`)}`;
			const result = await this.context.webAPI.retrieveMultipleRecords(lookupEntityType, query);
			const first = result.entities.length === 1 && !result.nextLink
				? result.entities[0] as Record<string, unknown>
				: undefined;
			const idRaw = first?.[primaryIdAttr];
			const guid = typeof idRaw === "string" && idRaw.trim().length > 0 ? this.normalizeGuid(idRaw) : null;
			this.lookupTokenGuidByNameCache[cacheKey] = guid;
			return guid;
		} catch {
			this.lookupTokenGuidByNameCache[cacheKey] = null;
			return null;
		}
	};

	private resolveLookupFilterTokens = async (clause: string, recordId?: string): Promise<string> => {
		if (!recordId || !clause.includes("{{")) return clause;
		const row = this.records?.[recordId] ?? this.accumulatedRecords?.[recordId] ?? this.childRecords[recordId];
		if (!row) {
			return clause;
		}


		const tokenRegex = /\{\{\s*([a-zA-Z0-9_]+)(?:\.([a-zA-Z0-9_]+))?\s*\}\}/g;
		let resolvedClause = clause;

		for (const tokenMatch of clause.matchAll(tokenRegex)) {
			const fullToken = tokenMatch[0];
			const attributeName = tokenMatch[1] ?? "";
			const tokenMode = tokenMatch[2];
			try {
				const mode = (tokenMode ?? "").toLowerCase();
				const raw = row.getValue(attributeName);
				let tokenValue = this.extractLookupTokenValue(raw, tokenMode);

				if ((tokenValue === null || tokenValue === undefined) && (mode === "" || mode === "id")) {
					const fallbackKeys = [`_${attributeName}_value`, `${attributeName}id`, `${attributeName}_id`];
					for (const key of fallbackKeys) {
						try {
							const fallbackRaw = row.getValue(key);
							if (typeof fallbackRaw === "string") {
								tokenValue = this.normalizeGuid(fallbackRaw);
								break;
							}
							const extracted = this.extractLookupTokenValue(fallbackRaw, "id");
							if (extracted !== null && extracted !== undefined) {
								tokenValue = extracted;
								break;
							}
						} catch {
							// continue trying other fallback keys
						}
					}
				}

				if ((tokenValue === null || tokenValue === undefined) && mode === "name") {
					try {
						tokenValue = row.getFormattedValue(attributeName);
					} catch {
						// ignore
					}
				}

				if ((tokenValue === null || tokenValue === undefined) && (mode === "" || mode === "id")) {
					tokenValue = await this.getLookupTokenGuidFromSourceRow(recordId, attributeName);
				}

				if ((tokenValue === null || tokenValue === undefined) && (mode === "" || mode === "id")) {
					tokenValue = await this.getLookupTokenGuidFromLookupObject(raw);
				}

				const literal = this.toODataLiteral(tokenValue);
				resolvedClause = resolvedClause.replace(fullToken, literal);
			} catch {
				resolvedClause = resolvedClause.replace(fullToken, "null");
			}
		}

		return resolvedClause;
	};

	private normalizeLookupComparisonFields = (clause: string, lookupFields: Map<string, string>, nullFallbackFields?: Set<string>): string => {
		const isGuidLiteral = (value: string): boolean => {
			const cleaned = value.trim().replace(/^'|'$/g, "").replace(/[{}]/g, "");
			return /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[1-5][0-9a-fA-F]{3}-[89abAB][0-9a-fA-F]{3}-[0-9a-fA-F]{12}$/.test(cleaned);
		};
		const isNullLiteral = (value: string): boolean => {
			const cleaned = value.trim().replace(/^'|'$/g, "").toLowerCase();
			return cleaned === "null";
		};

		return clause.replace(
			/(^|\(|\s)([a-zA-Z][a-zA-Z0-9_]*)\s+(eq|ne)\s+([^\s)]+)/g,
			(match, prefix: string, fieldName: string, op: string, rhs: string) => {
				if (fieldName.startsWith("_") || fieldName.endsWith("_value")) {
					return match;
				}
				const fieldKey = fieldName.toLowerCase();
				const canonicalLogicalName = lookupFields.get(fieldKey);
				if (canonicalLogicalName) {
					return `${prefix}_${canonicalLogicalName}_value ${op} ${rhs}`;
				}

				// Metadata can be partial in some clients. As a safe fallback, if the comparison
				// is against a GUID literal, treat it as a lookup-style comparison.
				if (isGuidLiteral(rhs) || (nullFallbackFields?.has(fieldKey) && isNullLiteral(rhs))) {
					return `${prefix}_${fieldKey}_value ${op} ${rhs}`;
				}

				return match;
			}
		);
	};

	private getMissingODataPropertyName = (error: unknown): string | null => {
		const message = (() => {
			if (typeof error === "string") return error;
			if (error && typeof error === "object" && "message" in error) {
				const m = (error as { message?: unknown }).message;
				return typeof m === "string" ? m : "";
			}
			return "";
		})();

		const match = /(?:Could not find a property named|undeclared property) '([A-Za-z0-9_]+)'/i.exec(message);
		return match?.[1] ?? null;
	};

	private getLookupFilterClause = (entityType: string, columnName?: string): string => {
		const rawConfig = [
			this.context.parameters.LookupFilterConfig?.raw,
			this.context.parameters.LookupFilterConfig2?.raw,
			this.context.parameters.LookupFilterConfig3?.raw,
			this.context.parameters.LookupFilterConfig4?.raw,
			this.context.parameters.LookupFilterConfig5?.raw,
		]
			.filter((part): part is string => typeof part === "string" && part.length > 0)
			.join("")
			.trim();
		if (!rawConfig) return "";

		const entityKey = entityType.toLowerCase();
		const columnKey = columnName?.toLowerCase();
		const columnEntityKey = columnKey ? `${columnKey}@${entityKey}` : undefined;

		try {
			const parsed = JSON.parse(rawConfig) as LookupFilterConfig;
			const fromColumnEntity =
				columnEntityKey && parsed.byColumnAndEntity ? parsed.byColumnAndEntity[columnEntityKey] : undefined;
			const fromColumn = columnKey && parsed.byColumn ? parsed.byColumn[columnKey] : undefined;
			const fromEntity = parsed.byEntity?.[entityKey];
			const fallback = parsed.default;
			return (fromColumnEntity ?? fromColumn ?? fromEntity ?? fallback ?? "").trim();
		} catch {
			// Backward-compatible mode: if not JSON, treat raw text as a global filter clause.
			return rawConfig;
		}
	};

	private getConfiguredPageSize = (): number | undefined => {
		const raw = this.context.parameters.MaxRows?.raw;
		if (raw === null || raw === undefined) return undefined;
		const parsed = Number(raw);
		if (!Number.isFinite(parsed) || parsed <= 0) return undefined;
		return Math.floor(parsed);
	};

	setSelectedRecords = (ids: string[]): void => {
		if (this.rootFilterBusy || this.rootFilterRecoveryRequired) { this.context.parameters.records.setSelectedRecordIds([]); return; }
		if (this.isDestroyed) return;
		this.selectionRevision++;
		if (this.nativeHierarchy) {
			this.nativeSelectionError = '';
			const children = ids.filter(id => !!this.childRecords[id] && !this.records?.[id]);
			void this.nativeHierarchy.select(ids, children);
			return;
		}
		this.context.parameters.records.setSelectedRecordIds(ids);
	};

	private initializeNativeHierarchy = (): void => {
		const dataset = this.context.parameters.records;
		if (!this.context.parameters.EnableNativeChildSelection?.raw || !this.context.parameters.EnableChildRows?.raw) return;
		const parent = this.context.parameters.ChildParentLookup?.raw?.trim() ?? '';
		const entity = dataset.getTargetEntityType();
		const key = `${entity}|${dataset.getViewId()}|${parent}`;
		if (key === this.nativeHierarchyKey) return;
		this.nativeHierarchyKey = key;
		this.nativeHierarchy?.dispose();
		this.nativeHierarchy = undefined;
		this.nativeHierarchyInitializing = true;
		void (async () => {
			try {
				if (!/^[a-zA-Z][a-zA-Z0-9_]*$/.test(parent)) throw new Error('Configure the child parent lookup before enabling native child selection.');
				const metadata = await this.getEntityMeta(entity);
				if (this.isDestroyed || key !== this.nativeHierarchyKey) return;
				if (!dataset.addColumn) throw new Error('This host does not support native child dataset columns.');
				if (!dataset.columns.some(column => column.name === parent)) await Promise.resolve(dataset.addColumn(parent));
				if (this.isDestroyed || key !== this.nativeHierarchyKey) return;
				this.nativeHierarchy = new NativeHierarchySelection(parent, metadata.primaryIdAttr, {
					dataset: () => this.context.parameters.records,
					basePageSize: () => this.getConfiguredPageSize() ?? 50,
					load: async (change, firstPage) => {
						// Public PCF paging calls may not overlap. Finish an in-flight
						// scroll batch before changing the command-selection query.
						await this.nextPageRequest;
						this.currentPage = firstPage ? 1 : this.currentPage + 1;
						return this.loadDatasetPageAndWait(change);
					},
					changed: () => this.renderGrid(),
					resetRows: () => { this.accumulatedRecords = {}; this.accumulatedSortedIds = []; },
					resetSelection: () => { this.nativeSelectionReset++; },
					error: message => { this.nativeSelectionError = message; },
				});
				this.nativeHierarchyInitializing = false;
				await this.nativeHierarchy.refreshRoots(this.nativeHierarchy.rootFilter);
			} catch (error) {
				if (this.isDestroyed || key !== this.nativeHierarchyKey) return;
				this.nativeSelectionError = error instanceof Error ? error.message : 'Unable to initialize native child selection.';
				this.nativeHierarchyInitializing = false;
				this.renderGrid();
			}
		})();
	};

	private waitForNextDatasetPage = (timeoutMs = 20000): PageLoadWaitHandle => {
		const dataset = this.context.parameters.records;
		const baseline = this.getDatasetPageSignature(dataset);
		let resolvePromise!: () => void;
		let rejectPromise!: (error: Error) => void;
		const promise = new Promise<void>((resolve, reject) => {
			resolvePromise = resolve;
			rejectPromise = reject;
		});

		const settle = (error?: Error): void => {
			if (waiter.settled) return;
			waiter.settled = true;
			window.clearTimeout(waiter.timerId);
			this.pageLoadWaiters.delete(waiter);
			if (error) {
				rejectPromise(error);
			} else {
				resolvePromise();
			}
		};

		const waiter: PageLoadWaiter = {
			baseline,
			timerId: 0,
			settled: false,
			resolve: () => settle(),
			reject: (error: Error) => settle(error),
		};
		waiter.timerId = window.setTimeout(
			() => waiter.reject(new Error("Timed out while loading a dataset page.")),
			timeoutMs
		);

		if (this.isDestroyed) {
			waiter.reject(new Error(CONTROL_DESTROYED_MESSAGE));
		} else {
			this.pageLoadWaiters.add(waiter);
		}

		return { promise, reject: waiter.reject };
	};

	private loadDatasetPageAndWait = (loadPage: () => void): Promise<void> => {
		const waitHandle = this.waitForNextDatasetPage();
		try {
			if (this.isDestroyed) {
				waitHandle.reject(new Error(CONTROL_DESTROYED_MESSAGE));
			} else {
				loadPage();
			}
		} catch (error) {
			const normalizedError = error instanceof Error ? error : new Error(String(error));
			// A synchronous host failure must remove the waiter immediately instead
			// of leaving a timer and a stale selection operation behind.
			waitHandle.reject(normalizedError);
		}
		return waitHandle.promise;
	};

	private getDatasetPageSignature = (dataset: ComponentFramework.PropertyTypes.DataSet): string => {
		const recordIds = dataset.sortedRecordIds ?? [];
		const recordKeys = Object.keys(dataset.records ?? {});
		return [
			dataset.paging.firstPageNumber,
			dataset.paging.lastPageNumber,
			dataset.paging.hasPreviousPage,
			dataset.paging.hasNextPage,
			recordIds.join(","),
			recordKeys.join(","),
		].join("|");
	};

	private onSelectAllRecords = async (options?: SelectAllRequest): Promise<SelectAllResult | undefined> => {
		if (this.rootFilterRecoveryRequired) throw new Error('Refresh the subgrid before selecting records.');
		if (this.rootFilterBusy) throw new Error('Wait for filtering to finish before selecting records.');
		if (this.nativeHierarchy) {
			this.selectionRevision++;
			return this.nativeHierarchy.loadRootSelectionIds(options);
		}
		if (this.selectAllAcrossPagesInProgress || this.isDestroyed) return;
		const selectionRevision = this.selectionRevision;
		this.selectAllAcrossPagesInProgress = true;
		const dataset = this.context.parameters.records;
		const originalPage = this.currentPage;
		let loadedPage = originalPage;
		let result: SelectAllResult | undefined;
		const checkCurrent = () => {
			checkSelectionCancelled(options?.signal);
			if (selectionRevision !== this.selectionRevision) throw new Error('The selection or view changed. Select the records again.');
		};
		try {
			// Keep automatic scrolling and editing from issuing competing requests
			// during the gaps between pages in a select-all operation.
			this.renderGrid();
			const selectedIds = new Set<string>();
			let additionalRows = false;
			const addCurrentPageIds = (): void => {
				for (const id of dataset.sortedRecordIds ?? []) {
					if (selectedIds.has(id)) continue;
					if (selectedIds.size === MAX_SELECTED_RECORDS) {additionalRows = true; break;}
					selectedIds.add(id);
				}
			};

			checkCurrent();
			if (originalPage !== 1) {
				this.currentPage = 1;
				await this.loadDatasetPageAndWait(() => dataset.paging.loadExactPage(1));
				if (this.isDestroyed) return;
				checkCurrent();
				loadedPage = 1;
			}

			addCurrentPageIds();
			let pagesVisited = 0;
			let previousPage = dataset.sortedRecordIds.join('|');
			while (selectedIds.size < MAX_SELECTED_RECORDS && dataset.paging.hasNextPage && pagesVisited < 500) {
				if (this.isDestroyed) return;
				checkCurrent();
				pagesVisited++;
				const nextPage = loadedPage + 1;
				this.currentPage = nextPage;
				await this.loadDatasetPageAndWait(() => dataset.paging.loadNextPage());
				if (this.isDestroyed) return;
				checkCurrent();
				loadedPage = nextPage;
				const signature = dataset.sortedRecordIds.join('|');
				if (signature === previousPage) throw new Error('The service repeated the previous page while selecting records. Refresh and try again.');
				previousPage = signature;
				addCurrentPageIds();
			}
			if (selectedIds.size < MAX_SELECTED_RECORDS && dataset.paging.hasNextPage) {
				throw new Error("Unable to select all records because the dataset exceeds 500 pages.");
			}

			result = selectionResult([...selectedIds], additionalRows || dataset.paging.hasNextPage || dataset.paging.totalResultCount > MAX_SELECTED_RECORDS);
		} catch (error) {
			if (this.isDestroyed || options?.signal?.aborted) return;
			const message = error instanceof Error ? error.message : "Unable to select every record.";
			await this.context.navigation.openAlertDialog({ text: message });
			throw error;
		} finally {
			try {
				if (!this.isDestroyed && selectionRevision === this.selectionRevision && loadedPage !== originalPage) {
					await this.loadDatasetPageAndWait(() => dataset.paging.loadExactPage(originalPage));
				}
			} finally {
				if (selectionRevision === this.selectionRevision) this.currentPage = originalPage;
				this.selectAllAcrossPagesInProgress = false;
				if (!this.isDestroyed) this.renderGrid();
			}
		}
		if (this.isDestroyed || options?.signal?.aborted) return;
		checkCurrent();
		return result;
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
		const flags = column as typeof column & { isReadOnly?: boolean; readOnly?: boolean; disableEditing?: boolean };
		if (flags.isReadOnly || flags.readOnly || flags.disableEditing) throw new Error('This column is read only.');
		return column;
	};

	/** Recheck at save time: the host can disable an already-open editor. */
	private assertEditingAllowed = (recordId: string): void => {
		if (this.isDestroyed || this.context.mode?.isControlDisabled || this.context.parameters.EnableInlineEdit?.raw === false) {
			throw new Error('This grid is read only.');
		}
		if (this.childRecords[this.normalizeGuid(recordId)] && this.context.parameters.DisableChildEditing?.raw) {
			throw new Error('Child records are read only.');
		}
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
		// Text whitespace is user data. Only typed number/date/choice values are trimmed.
		if (/singleline|multiple|memo|text/i.test(dataType)) return value.length ? value : null;
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

		if (normalizedType.includes("dateonly")) {
			if (!/^\d{4}-\d{2}-\d{2}$/.test(trimmedValue)) {
				throw new Error("Please enter a valid date in YYYY-MM-DD format.");
			}
			const [year, month, day] = trimmedValue.split("-").map(Number);
			const parsedDate = new Date(Date.UTC(year, month - 1, day));
			if (
				parsedDate.getUTCFullYear() !== year ||
				parsedDate.getUTCMonth() !== month - 1 ||
				parsedDate.getUTCDate() !== day
			) {
				throw new Error("Please enter a valid date in YYYY-MM-DD format.");
			}
			return trimmedValue;
		}

		if (normalizedType.includes("dateandtime") || normalizedType.includes("datetime") || normalizedType.includes("date")) {
			const parsedDate = new Date(trimmedValue);
			if (Number.isNaN(parsedDate.getTime())) {
				throw new Error("Please enter a valid date value.");
			}
			return parsedDate.toISOString();
		}

		if (normalizedType.includes("whole.") || normalizedType.includes("optionset")) {
			const parsedInteger = Number(trimmedValue);
			if (!Number.isFinite(parsedInteger) || !Number.isInteger(parsedInteger)) {
				throw new Error("Please enter a whole number.");
			}
			return parsedInteger;
		}

		if (
			normalizedType.includes("decimal") ||
			normalizedType.includes("currency") || normalizedType.includes("money") ||
			normalizedType.includes("fp") || normalizedType.includes("double") || normalizedType.includes("float")
		) {
			const parsedDecimal = Number(trimmedValue);
			if (!Number.isFinite(parsedDecimal)) {
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
			this.assertEditingAllowed(recordId);
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

	onSort = (name: string, direction: "asc" | "desc" | "none", additive = false): void => {
		const records = this.context.parameters.records;
		const column = records.columns.find((candidate) => candidate.name === name);
		if (this.context.parameters.EnableSorting?.raw === false || column?.disableSorting) return;

		// Dataverse resets filtering when the dataset sorting collection is changed.
		// Snapshot the complete expression and reapply it after sorting so a sort
		// does not silently broaden the current result set.
		const existingFilter = this.getRootFilter();
		const visibleSorting = this.nativeHierarchy?.rootSorting ?? records.sorting;
		const nextSorting = additive || direction === "none" ? visibleSorting.map((sort) => ({ ...sort })) : [];
		const existingIndex = nextSorting.findIndex((sort) => sort.name === name);
		if (existingIndex >= 0) nextSorting.splice(existingIndex, 1);
		if (direction !== "none") {
			const nextSort: ComponentFramework.PropertyHelper.DataSetApi.SortStatus = { name, sortDirection: direction === "desc" ? 1 : 0 };
			if (existingIndex >= 0) nextSorting.splice(existingIndex, 0, nextSort);
			else nextSorting.push(nextSort);
		}
		this.applyRootFilter(existingFilter,nextSorting);
	};

	onFilter = (
		name: string,
		mode: FilterMode,
		value?: FilterValue
	): void => {
		const nextFilter = updateColumnFilter(this.pendingRootFilter?.filter ?? this.getRootFilter(), name, mode, value);
		this.applyRootFilter(nextFilter);
	};

	onClearAllFilters = (): void => {
		this.applyRootFilter(undefined);
	};

	loadFirstPage = (): void => {
		const previousPage = this.currentPage;
		this.currentPage = 1;
		try {
			this.context.parameters.records.paging.loadExactPage(1);
		} catch (error) {
			this.currentPage = previousPage;
			throw error;
		}
	};

	loadNextPage = (): Promise<void> => {
		if (this.nextPageRequest) return this.nextPageRequest;
		const previousPage = this.currentPage;
		this.currentPage++;
		this.appendingPage = true;
		const request = this.loadDatasetPageAndWait(() => this.context.parameters.records.paging.loadNextPage()).catch((error: unknown) => {
			this.currentPage = previousPage;
			throw error;
		}).finally(() => { if (this.nextPageRequest === request) this.nextPageRequest = undefined; this.appendingPage = false; });
		this.nextPageRequest = request;
		return request;
	};

	loadPreviousPage = (): void => {
		const previousPage = this.currentPage;
		this.currentPage--;
		try {
			this.context.parameters.records.paging.loadExactPage(this.currentPage);
		} catch (error) {
			this.currentPage = previousPage;
			throw error;
		}
	};

	/* eslint-disable @typescript-eslint/no-unsafe-assignment, @typescript-eslint/no-unsafe-call, @typescript-eslint/no-unsafe-member-access, @typescript-eslint/no-explicit-any */
	private getEntityMeta = async (entityType: string): Promise<EntityMetadata> => {
		if (this.entityMetaCache[entityType]) return this.entityMetaCache[entityType];

		// Convention is a fallback only; a custom table may choose another primary name.
		const publisherPrefix = entityType.includes("_") ? entityType.split("_")[0] : "";
		const conventionName = publisherPrefix ? `${publisherPrefix}_name` : "name";

		try {
			const utils = this.context.utils as any;
			const meta = await utils.getEntityMetadata(entityType);
			const metaEntitySet = typeof meta?.EntitySetName === "string" ? meta.EntitySetName : "";
			const metaPrimaryId = typeof meta?.PrimaryIdAttribute === "string" ? meta.PrimaryIdAttribute : "";
			const metaPrimaryName = typeof meta?.PrimaryNameAttribute === "string" ? meta.PrimaryNameAttribute : "";

			// Extract lookup field logical names from entity attribute metadata.
			const lookupFields = new Map<string, string>();
			try {
				const attrs = meta?.Attributes;
				const attrArray: any[] = typeof attrs?.getAll === "function" ? attrs.getAll()
					: typeof attrs?.get === "function" ? attrs.get()
					: Array.isArray(attrs) ? attrs : [];
				for (const attr of attrArray) {
					const attrType = (attr?.AttributeType ?? attr?.attributeType ?? "") as string;
					if (attrType === "Lookup" || attrType === "Customer" || attrType === "Owner"
						|| attrType === "lookup" || attrType === "customer" || attrType === "owner") {
						const logicalName = (attr?.LogicalName ?? attr?.logicalName ?? "") as string;
						if (logicalName) {
							const canonical = logicalName.toLowerCase();
							lookupFields.set(canonical, canonical);
						}
					}
				}
			} catch {
				// Attribute metadata unavailable — lookupFields stays empty (no normalization).
			}

			const result = {
				entitySetName: metaEntitySet !== "" ? metaEntitySet : `${entityType}s`,
				primaryIdAttr: metaPrimaryId !== "" ? metaPrimaryId : `${entityType}id`,
				primaryNameAttr: metaPrimaryName !== '' ? metaPrimaryName : conventionName,
				lookupFields,
			};
			this.entityMetaCache[entityType] = result;
			return result;
		} catch {
			// Fall back to naming conventions if metadata fetch fails.
			const fallback = {
				entitySetName: `${entityType}s`,
				primaryIdAttr: `${entityType}id`,
				primaryNameAttr: conventionName,
				lookupFields: new Map<string, string>(),
			};
			this.entityMetaCache[entityType] = fallback;
			return fallback;
		}
	};
	/* eslint-enable @typescript-eslint/no-unsafe-assignment, @typescript-eslint/no-unsafe-call, @typescript-eslint/no-unsafe-member-access, @typescript-eslint/no-explicit-any */

	onSearchLookup = async (
		entityType: string,
		searchTerm: string,
		columnName?: string,
		recordId?: string
	): Promise<{ id: string; name: string }[]> => {
		const { primaryIdAttr, primaryNameAttr, lookupFields } = await this.getEntityMeta(entityType);
		const safeSearch = (searchTerm ?? "").trim().replace(/'/g, "''");
		const searchClause = safeSearch ? `contains(${primaryNameAttr},'${safeSearch}')` : "";
		const rawConfiguredClause = await this.resolveLookupFilterTokens(
			this.getLookupFilterClause(entityType, columnName),
			recordId
		);
		const toQuery = (configuredClause: string): string => {
			const combinedFilter = [searchClause, configuredClause]
				.filter((clause) => clause && clause.trim().length > 0)
				.map((clause) => `(${clause.trim()})`)
				.join(" and ");
			const filter = combinedFilter ? `&$filter=${encodeURIComponent(combinedFilter)}` : "";
			return `?$select=${primaryIdAttr},${primaryNameAttr}&$top=20&$orderby=${primaryNameAttr} asc${filter}`;
		};
		const mapResults = (results: { entities: ComponentFramework.WebApi.Entity[] }): { id: string; name: string }[] => results.entities.map((e) => {
			// eslint-disable-next-line @typescript-eslint/no-base-to-string
			const id = String((e as Record<string, unknown>)[primaryIdAttr] ?? (e as Record<string, unknown>).id ?? "");
			// eslint-disable-next-line @typescript-eslint/no-base-to-string
			const name = String((e as Record<string, unknown>)[primaryNameAttr] ?? "(unknown)");
			return { id, name };
		});

		const configuredClause = this.normalizeLookupComparisonFields(rawConfiguredClause, lookupFields);
		try {
			const query = toQuery(configuredClause);
			const results = await this.context.webAPI.retrieveMultipleRecords(entityType, query);
			return mapResults(results);
		} catch (error) {
			const missingProperty = this.getMissingODataPropertyName(error)?.toLowerCase();
			if (!missingProperty) throw error;

			const retriedConfiguredClause = this.normalizeLookupComparisonFields(
				rawConfiguredClause,
				lookupFields,
				new Set([missingProperty])
			);
			if (retriedConfiguredClause === configuredClause) throw error;

			const retryQuery = toQuery(retriedConfiguredClause);
			const retryResults = await this.context.webAPI.retrieveMultipleRecords(entityType, retryQuery);
			return mapResults(retryResults);
		}
	};

	/** Supported PCF lookup dialog. Its result is revalidated against this column's restrictions. */
	onBrowseLookup = async (entityType: string, searchTerm: string, columnName: string, recordId: string): Promise<{id: string; name: string} | undefined> => {
		this.assertEditingAllowed(recordId);
		const dataset = this.context.parameters.records;
		const view = dataset.getViewId?.();
		const sourceEntity = dataset.getTargetEntityType();
		const assertCurrent = () => {
			this.assertEditingAllowed(recordId); this.getEditableDatasetColumn(columnName);
			if (this.context.parameters.records.getViewId?.() !== view || this.context.parameters.records.getTargetEntityType() !== sourceEntity) throw new Error('The grid changed. Open the lookup again.');
		};
		assertCurrent();
		childLogicalName(entityType);
		const options: ComponentFramework.UtilityApi.LookupOptions & {searchText: string} = {allowMultiSelect: false, defaultEntityType: entityType, entityTypes: [entityType], searchText: searchTerm};
		const result = await this.context.utils.lookupObjects(options);
		assertCurrent();
		if (!result?.length) return undefined;
		if (result.length !== 1 || result[0].entityType.toLowerCase() !== entityType.toLowerCase()) throw new Error('Choose one record from the correct table.');
		const id = childGuid(result[0].id);
		const meta = await this.getEntityMeta(entityType);
		const rawClause = await this.resolveLookupFilterTokens(this.getLookupFilterClause(entityType, columnName), recordId);
		assertCurrent();
		const clause = this.normalizeLookupComparisonFields(rawClause, meta.lookupFields);
		const filter = [`${childLogicalName(meta.primaryIdAttr)} eq ${id}`, clause].filter(Boolean).map(part => `(${part})`).join(' and ');
		const allowed = await this.context.webAPI.retrieveMultipleRecords(entityType, `?$select=${childLogicalName(meta.primaryIdAttr)},${childLogicalName(meta.primaryNameAttr)}&$top=1&$filter=${encodeURIComponent(filter)}`, 1);
		assertCurrent();
		const row = allowed.entities.find(row => typeof row[meta.primaryIdAttr] === 'string' && childGuid(row[meta.primaryIdAttr] as string) === id);
		if (!row) throw new Error('This record does not meet the lookup restrictions. Choose another record.');
		return {id, name: typeof row[meta.primaryNameAttr] === 'string' ? row[meta.primaryNameAttr] as string : result[0].name ?? ''};
	};

	onUpdateLookupCell = async (
		recordId: string,
		columnName: string,
		targetId: string,
		targetEntityType: string
	): Promise<void> => {
		const entityType = this.context.parameters.records.getTargetEntityType();
		const viewId = this.context.parameters.records.getViewId?.();
		if (!entityType) throw new Error("Unable to determine target table for inline editing.");
		try {
			this.assertEditingAllowed(recordId);
			const editableColumn = this.getEditableDatasetColumn(columnName);
			const originalDataType = editableColumn.dataType;

			// Ensure the nav property map is fully populated before building candidates.
			if (this.lookupNavPropertyFetchPromise) {
				await this.lookupNavPropertyFetchPromise;
			}

			// Use the eagerly-fetched nav property map first (most reliable).
			// Fall back to the runtime metadata resolver for any remaining candidates.
			const key = editableColumn.name.toLowerCase();
			// A polymorphic column can have a different binding per target table.
			const prebuiltNavProp = this.lookupNavPropertyMap[`${key}:${targetEntityType.toLowerCase()}`]
				?? (this.lookupTargetEntityMap[key] === targetEntityType.toLowerCase() ? this.lookupNavPropertyMap[key] : undefined);
			const isClear = targetId === "";
			const entitySetName = isClear ? "" : (await this.getEntityMeta(targetEntityType)).entitySetName;
			const runtimeCandidates = await this.resolveLookupBindNames(entityType, editableColumn.name, targetEntityType);
			const bindCandidates = prebuiltNavProp
				? [prebuiltNavProp, ...runtimeCandidates.filter((c) => c !== prebuiltNavProp)]
				: runtimeCandidates;

			let lastError: unknown;
			for (const bindName of bindCandidates) {
				this.assertEditingAllowed(recordId);
				const currentColumn = this.getEditableDatasetColumn(columnName);
				if (this.context.parameters.records.getTargetEntityType() !== entityType || this.context.parameters.records.getViewId?.() !== viewId || currentColumn.dataType !== originalDataType) {
					throw new Error('The grid changed. Open the lookup again.');
				}
				try {
					const updateData: ComponentFramework.WebApi.Entity = isClear
						? { [`${bindName}@odata.bind`]: null }
						: { [`${bindName}@odata.bind`]: `/${entitySetName}(${this.normalizeGuid(targetId)})` };
					await this.context.webAPI.updateRecord(entityType, this.normalizeGuid(recordId), updateData);
					this.context.parameters.records.refresh();
					return;
				} catch (error) {
					lastError = error;
					// Retry only a confirmed property-name rejection. A timeout, access
					// failure, conflict or business rule must not repeat a possibly applied save.
					if (this.getMissingODataPropertyName(error)?.toLowerCase() !== bindName.toLowerCase()) throw error;
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

			(data.value ?? []).forEach((rel) => {
				const attr = (rel.ReferencingAttribute ?? "").toLowerCase();
				const navProp = rel.ReferencingEntityNavigationPropertyName ?? "";
				const referencedEntity = (rel.ReferencedEntity ?? "").toLowerCase();
				if (attr && navProp && lookupColNames.has(attr)) {
					if (referencedEntity) this.lookupNavPropertyMap[`${attr}:${referencedEntity}`] = navProp;
					if (!this.lookupNavPropertyMap[attr]) {
						this.lookupNavPropertyMap[attr] = navProp;
					}
					if (referencedEntity && !this.lookupTargetEntityMap[attr]) {
						this.lookupTargetEntityMap[attr] = referencedEntity;
					}
				}
			});

		} catch {
			// Runtime metadata resolution remains available; do not log record/query data.
		}
	};

	private fetchColumnOptions = async (
		entityType: string,
		columns: ComponentFramework.PropertyHelper.DataSetApi.Column[]
	): Promise<void> => {
		const optionSetCols = columns.filter((col) => {
			const t = col.dataType.toLowerCase();
			return (t.includes("optionset") && !t.includes("multiselectoptionset")) || t === 'state' || t === 'status';
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
							color: typeof opt.Color === 'string' ? opt.Color : undefined,
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

	private refreshDatasetFromTabActivation = (): void => {
		const dataset = this.context?.parameters?.records;
		if (!dataset || dataset.loading || this.rootFilterBusy) return;

		const now = Date.now();
		if (now - this.lastTabRefreshAt < 1500) return;
		this.lastTabRefreshAt = now;

		dataset.refresh();
	};

	private scheduleTabRefresh = (): void => {
		if (typeof window === "undefined") return;
		if (this.tabRefreshTimeoutId !== undefined) {
			window.clearTimeout(this.tabRefreshTimeoutId);
		}
		this.tabRefreshTimeoutId = window.setTimeout(() => {
			this.tabRefreshTimeoutId = undefined;
			this.refreshDatasetFromTabActivation();
		}, 120);
	};

	private getControlVisibleState = (context: ComponentFramework.Context<IInputs>): boolean => {
		const mode = context.mode as unknown as { isVisible?: boolean; allocatedWidth?: number | string; allocatedHeight?: number | string };
		if (typeof mode.isVisible === "boolean") {
			return mode.isVisible;
		}

		const width = Number(mode.allocatedWidth ?? 0);
		const height = Number(mode.allocatedHeight ?? 0);
		return Number.isFinite(width) && Number.isFinite(height) && width > 0 && height > 0;
	};

	private renderGrid = (): void => {
		if (this.isDestroyed || !this.container || !this.context) return;
		const dataset = this.context.parameters.records;
		const paging = dataset.paging;
		const native = this.nativeHierarchy;
		const waitingForNative = (!!this.context.parameters.EnableNativeChildSelection?.raw && !native) || this.rootFilterBusy;

		const parseAllocatedSize = (value: unknown): number => {
			if (typeof value === "number") {
				return Number.isFinite(value) ? value : Number.NaN;
			}
			if (typeof value !== "string") {
				return Number.NaN;
			}
			const trimmed = value.trim();
			// Only accept plain numeric strings; reject values like "100%".
			if (!/^[-+]?\d+(\.\d+)?$/.test(trimmed)) {
				return Number.NaN;
			}
			const parsed = Number(trimmed);
			return Number.isFinite(parsed) ? parsed : Number.NaN;
		};

		const rawAllocatedWidth = this.context.mode.allocatedWidth as unknown;
		const rawAllocatedHeight = this.context.mode.allocatedHeight as unknown;
		let allocatedWidth = parseAllocatedSize(rawAllocatedWidth);
		let allocatedHeight = parseAllocatedSize(rawAllocatedHeight);

		if (!Number.isFinite(allocatedWidth) || allocatedWidth <= 0) {
			allocatedWidth = -1;
		}
		if (this.context.parameters.SubGridHeight.raw) {
			allocatedHeight = this.context.parameters.SubGridHeight.raw;
		}
		// Pass -1 when no valid height is available — Grid will use 100% to fill its container.
		if (!Number.isFinite(allocatedHeight) || allocatedHeight <= 0) {
			allocatedHeight = -1;
		}

		const viewKey = `${dataset.getTargetEntityType?.() ?? ''}|${dataset.getViewId?.() ?? ''}|${this.context.parameters.ChildViewId?.raw ?? ''}`;
		const userId = this.context.userSettings?.userId;
		const layoutKey = this.context.parameters.RememberColumnLayout?.raw !== false && userId && dataset.getViewId?.()
			? `${window.location.origin}|${userId.toLowerCase()}|${viewKey}` : undefined;
		ReactDOM.render(
			React.createElement(Grid, {
				key: `${userId ?? ''}|${viewKey}`,
				layoutKey,
				features: {
					accessibleLabel: this.context.parameters.AccessibleLabel?.raw?.trim() ? this.context.parameters.AccessibleLabel.raw : 'Records',
					enableGrouping: this.context.parameters.EnableGrouping?.raw ?? true,
					enableAggregation: this.context.parameters.EnableAggregation?.raw ?? true,
					enableColumnMoving: this.context.parameters.EnableColumnMoving?.raw ?? true,
					enableMultipleSelection: this.context.parameters.EnableMultipleSelection?.raw ?? true,
					enableRangeSelection: this.context.parameters.EnableRangeSelection?.raw ?? true,
					enableFiltering: this.context.parameters.EnableFiltering?.raw ?? true,
					enableSorting: this.context.parameters.EnableSorting?.raw ?? true,
					showTypeIcons: this.context.parameters.ShowTypeIcons?.raw ?? false,
					showStatusIcons: this.context.parameters.ShowStatusIcons?.raw ?? false,
					showChoiceColors: this.context.parameters.ShowChoiceColors?.raw ?? false,
					reflow: this.context.parameters.ReflowBehavior?.raw === 'list' ? 'list' : this.context.parameters.ReflowBehavior?.raw === 'grid' ? 'grid' : 'auto',
					listColumns: this.context.parameters.ListColumns?.raw ?? 3,
					showAvatar: this.context.parameters.ShowAvatar?.raw ?? false,
					showListColumnNames: this.context.parameters.ShowListColumnNames?.raw ?? true,
					bandedListRows: this.context.parameters.BandedListRows?.raw ?? true,
					enableJumpBar: this.context.parameters.EnableJumpBar?.raw ?? false,
					navigationTypes: this.context.parameters.NavigationTypes?.raw === 'primary' ? 'primary' : this.context.parameters.NavigationTypes?.raw === 'none' ? 'none' : 'all',
					firstDayOfWeek: this.context.userSettings?.dateFormattingInfo?.firstDayOfWeek,
					isRTL: this.context.userSettings?.isRTL,
				},
				selectionCountOverride: native?.selectedCount,
				selectionResetVersion: this.nativeSelectionReset,
				commandStatus: this.rootFilterBusy ? 'Checking filters across all grid pages…' : this.nativeHierarchyInitializing || native?.busy ? 'Preparing record commands…' : undefined,
				commandError: this.rootFilterError || this.nativeSelectionError || undefined,
				queryErrorVersion: this.queryErrorVersion,
				disableSelection: this.rootFilterRecoveryRequired,
				childRows: this.context.parameters.EnableChildRows?.raw ? {loadPage:this.loadChildPage,hasChildren:this.checkChildRows,revision:this.childRevision,
					readOnly:!!this.context.parameters.DisableChildEditing?.raw || this.context.mode.isControlDisabled} : undefined,
				width: allocatedWidth,
				height: allocatedHeight,
				columns: native ? dataset.columns.filter(column => column.name !== native.parentField) : dataset.columns,
				records: waitingForNative ? {} : this.records ?? {},
				sortedRecordIds: waitingForNative ? [] : this.sortedRecordsIds ?? [],
				hasNextPage: paging.hasNextPage,
				hasPreviousPage: paging.hasPreviousPage,
				currentPage: this.currentPage,
				totalResultCount: paging.totalResultCount < 0 ? paging.totalResultCount : Math.max(0, paging.totalResultCount - (native?.childIds.length ?? 0)),
				sorting: native?.rootSorting ?? dataset.sorting,
				filtering: this.getRootFilter() ?? {filterOperator: 0, conditions: []},
				resources: this.resources,
				itemsLoading: this.rootFilterBusy || dataset.loading || this.selectAllAcrossPagesInProgress || this.nativeHierarchyInitializing || !!native?.busy,
				highlightValue: this.context.parameters.HighlightValue.raw,
				highlightColor: this.context.parameters.HighlightColor.raw,
				enableLookupLinks: this.context.parameters.EnableLookupLinks.raw ?? true,
				enableInlineEdit: !this.context.mode.isControlDisabled && (this.context.parameters.EnableInlineEdit.raw ?? true),
				infiniteScroll: this.context.parameters.InfiniteScroll.raw ?? true,
				columnOptions: this.columnOptions,
				lookupTargetEntityMap: this.lookupTargetEntityMap,
				onSearchLookup: this.onSearchLookup,
				onBrowseLookup: this.context.parameters.EnableLookupDialog?.raw !== false && typeof this.context.utils?.lookupObjects === 'function' ? this.onBrowseLookup : undefined,
				onUpdateLookupCell: this.onUpdateLookupCell,
				onSelectAllRecords: this.onSelectAllRecords,
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
		this.isDestroyed = false;
		this.notifyOutputChanged = notifyOutputChanged;
		this.container = container;
		this.container.style.height = "100%";
		this.context = context;
		this.context.mode.trackContainerResize(true);
		this.resources = this.context.resources;
		this.isTestHarness = document.getElementById("control-dimensions") !== null;

		if (typeof document !== "undefined") {
			this.visibilityChangeHandler = () => {
				if (document.visibilityState === "visible") {
					this.scheduleTabRefresh();
				}
			};
			document.addEventListener("visibilitychange", this.visibilityChangeHandler);
		}

		if (typeof window !== "undefined") {
			this.windowFocusHandler = () => {
				this.scheduleTabRefresh();
			};
			window.addEventListener("focus", this.windowFocusHandler);
		}
	}

	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void {
		if (this.isDestroyed) return;
		const dataset = context.parameters.records;
		this.context = context;
		const filterViewKey=`${dataset.getTargetEntityType?.() ?? ''}|${dataset.getViewId?.() ?? ''}`;
		if (this.filterViewKey && this.filterViewKey !== filterViewKey) {
			this.ownsRootFilter=false; this.logicalRootFilter=undefined; this.pendingRootFilter=undefined; this.rootFilterError='';
			this.rootFilterRecoveryRequired=false;
		}
		this.filterViewKey=filterViewKey;
		this.initializeNativeHierarchy();
		const native = this.nativeHierarchy;
		const completedExternalRefresh=this.childDatasetWasLoading && !dataset.loading && !this.appendingPage && !this.selectAllAcrossPagesInProgress && !this.rootFilterBusy && !native?.busy && (!native || !dataset.paging.hasPreviousPage);
		if (completedExternalRefresh) {
			if (!dataset.error) {
				// After a failed rollback, the accepted host query is authoritative.
				// Do not re-enable commands with an old logical filter still displayed.
				if (this.rootFilterRecoveryRequired) {this.ownsRootFilter=false; this.logicalRootFilter=undefined;}
				this.rootFilterRecoveryRequired=false; this.rootFilterError='';
			}
			this.childRevision++;
			native?.externalRefresh();
			if (!native) {this.selectionRevision++; this.nativeSelectionReset++; dataset.setSelectedRecordIds([]);}
		}
		this.childDatasetWasLoading = dataset.loading;

		const isVisibleNow = this.getControlVisibleState(context);
		const becameVisible = this.lastControlVisible === false && isVisibleNow;
		this.lastControlVisible = isVisibleNow;
		if (becameVisible && !dataset.loading && !this.rootFilterBusy) {
			const now = Date.now();
			if (now - this.lastVisibilityRefreshAt > 1500) {
				this.lastVisibilityRefreshAt = now;
				dataset.refresh();
				return;
			}
		}

		const configuredPageSize = this.getConfiguredPageSize();
		if (!context.parameters.EnableNativeChildSelection?.raw && configuredPageSize && configuredPageSize !== this.appliedPageSize) {
			try {
				dataset.paging.setPageSize(configuredPageSize);
				this.appliedPageSize = configuredPageSize;
				this.currentPage = 1;
				dataset.refresh();
				return;
			} catch {
				// Ignore and continue rendering with platform defaults.
			}
		}

		// In MDAs, the initial population of the dataset does not provide updatedProperties.
		// Ensure we always hydrate local record caches at least once.
		const initialLoad = this.records === undefined;
		const datasetChanged =
			context.updatedProperties.includes("dataset") || context.updatedProperties.includes("records") || initialLoad;
		if (datasetChanged && !dataset.loading) {
			this.lookupTokenGuidCache = {};
			this.lookupTokenGuidByNameCache = {};
		}
		const resetPaging = datasetChanged && !dataset.loading && !this.appendingPage && !dataset.paging.hasPreviousPage && this.currentPage !== 1;

		if (resetPaging) {
			this.currentPage = 1;
		}

		// Always keep local caches in sync with the latest dataset payload.
		// Some MDA update cycles don't flag updatedProperties as expected.
		const infiniteScroll = this.context.parameters.InfiniteScroll.raw ?? true;
		const pageRecords = native ? Object.fromEntries(Object.entries(dataset.records ?? {}).filter(([, row]) => isHierarchyRoot(row, native.parentField))) : dataset.records ?? {};
		const pageIds = (dataset.sortedRecordIds ?? []).filter(id => !!pageRecords[id]);
		if (infiniteScroll && !dataset.loading && native?.acceptRows !== false) {
			if (!dataset.paging.hasPreviousPage && !native && !this.appendingPage) {
				// Fresh load (filter/sort change or first page) — reset accumulated records.
				// Do this even when a host omits its dataset/records updatedProperties flag;
				// otherwise the row count can update while stale row objects remain visible.
				this.accumulatedRecords = { ...pageRecords };
				this.accumulatedSortedIds = [...pageIds];
			} else {
				// Additional page — merge new records in.
				this.accumulatedRecords = { ...this.accumulatedRecords, ...pageRecords };
				const existingIds = new Set(this.accumulatedSortedIds);
				this.accumulatedSortedIds = [...this.accumulatedSortedIds, ...pageIds.filter(id => !existingIds.has(id))];
			}
			this.records = this.accumulatedRecords;
			this.sortedRecordsIds = this.accumulatedSortedIds;
		} else if (!infiniteScroll && native?.acceptRows !== false) {
			this.records = pageRecords;
			this.sortedRecordsIds = pageIds;
		}

		if (!dataset.loading && this.pageLoadWaiters.size > 0) {
			// Most hosts flag a dataset update through updatedProperties. Some model-driven
			// app cycles omit that flag, so also release a waiter once the page payload or
			// paging state differs from the snapshot taken before the load request.
			const currentSignature = this.getDatasetPageSignature(dataset);
			const waiters = Array.from(this.pageLoadWaiters).filter(
				(waiter) => datasetChanged || waiter.baseline !== currentSignature
			);
			for (const waiter of waiters) {
				if (dataset.error) waiter.reject(new Error('Power Apps could not load the next batch of records.'));
				else waiter.resolve();
			}
		}

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
			this.lookupTargetEntityMap = {};
			this.lookupNavPropertyFetchPromise = this.fetchLookupNavProperties(entityType, dataset.columns);
		}

		this.renderGrid();
		if (completedExternalRefresh && needsServerFilter(this.getRootFilter())) this.applyRootFilter(this.getRootFilter());
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
		this.isDestroyed = true;
		this.nativeHierarchy?.dispose();
		this.childRecords = {};
		this.childLoader = undefined;
		this.childAvailabilityLoader = undefined;
		if (this.visibilityChangeHandler && typeof document !== "undefined") {
			document.removeEventListener("visibilitychange", this.visibilityChangeHandler);
			this.visibilityChangeHandler = undefined;
		}
		if (this.windowFocusHandler && typeof window !== "undefined") {
			window.removeEventListener("focus", this.windowFocusHandler);
			this.windowFocusHandler = undefined;
		}
		if (this.tabRefreshTimeoutId !== undefined && typeof window !== "undefined") {
			window.clearTimeout(this.tabRefreshTimeoutId);
			this.tabRefreshTimeoutId = undefined;
		}
		const waiters = Array.from(this.pageLoadWaiters);
		for (const waiter of waiters) {
			waiter.reject(new Error(CONTROL_DESTROYED_MESSAGE));
		}
		this.pageLoadWaiters.clear();
		if (this.container) {
			ReactDOM.unmountComponentAtNode(this.container);
		}
	}
}
