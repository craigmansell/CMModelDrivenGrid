import { CMModelDrivenGrid } from "../CMModelDrivenGrid";
import { getColumnConditions, getFilterColumnNames, updateColumnFilter } from "../CMModelDrivenGrid/filtering";
import { SelectAllRequest, SelectAllResult } from '../CMModelDrivenGrid/recordSelection';

type TestContext = ComponentFramework.Context<import("../CMModelDrivenGrid/generated/ManifestTypes").IInputs>;

interface TestDatasetState {
	loading: boolean;
	records: Record<string, unknown>;
	sortedRecordIds: string[];
	columns: unknown[];
	sorting: unknown[];
	paging: {
		firstPageNumber: number;
		lastPageNumber: number;
		hasPreviousPage: boolean;
		hasNextPage: boolean;
		totalResultCount: number;
		loadNextPage: () => void;
		loadExactPage: (page: number) => void;
		setPageSize: () => void;
	};
	filtering: {
		getFilter: () => unknown;
		setFilter: (filter: unknown) => void;
	};
	getTargetEntityType: () => string;
	setSelectedRecordIds: (ids: string[]) => void;
	refresh: () => void;
}

function assert(condition: boolean, message: string): asserts condition {
	if (!condition) throw new Error(message);
}

function assertThrows(action: () => unknown, expectedMessage: string): void {
	try {
		action();
	} catch (error) {
		const message = error instanceof Error ? error.message : String(error);
		assert(message.includes(expectedMessage), `Expected '${expectedMessage}', received '${message}'.`);
		return;
	}
	throw new Error(`Expected an error containing '${expectedMessage}'.`);
}

function installNodeWindowShim(): void {
	const globalWithWindow = globalThis as unknown as {
		window?: {
			setTimeout: (handler: () => void, timeout: number) => number;
			clearTimeout: (id: number) => void;
		};
	};
	if (globalWithWindow.window) return;
	globalWithWindow.window = {
		setTimeout: (handler, timeout) => globalThis.setTimeout(handler, timeout) as unknown as number,
		clearTimeout: (id) => globalThis.clearTimeout(id as unknown as ReturnType<typeof globalThis.setTimeout>),
	};
}

function createPagingControl(options: {
	loadNextPage?: () => void;
	loadExactPage?: (page: number) => void;
} = {}): {
	control: CMModelDrivenGrid;
	testableControl: {
		onSelectAllRecords: (request?: SelectAllRequest) => Promise<SelectAllResult | void>;
		pageLoadWaiters: Set<unknown>;
		renderGrid: () => void;
		records: Record<string, unknown>;
		sortedRecordsIds: string[];
	};
	dataset: TestDatasetState;
	context: TestContext;
	selectedIds: string[];
} {
	installNodeWindowShim();
	const control = new CMModelDrivenGrid();
	const selectedIds: string[] = [];
	const dataset = {
		loading: false,
		records: { r1: {} },
		sortedRecordIds: ["r1"],
		columns: [],
		sorting: [],
		paging: {
			firstPageNumber: 1,
			lastPageNumber: 1,
			hasPreviousPage: false,
			hasNextPage: true,
			totalResultCount: 2,
			loadNextPage: (): void => options.loadNextPage?.(),
			loadExactPage: (page: number): void => options.loadExactPage?.(page),
			setPageSize: (): void => undefined,
		},
		filtering: {
			getFilter: (): undefined => undefined,
			setFilter: (_filter: unknown): void => undefined,
		},
		getTargetEntityType: (): string => "",
		setSelectedRecordIds: (ids: string[]): void => {
			selectedIds.splice(0, selectedIds.length, ...ids);
		},
		refresh: (): void => undefined,
	} satisfies TestDatasetState;
	const context = {
		parameters: {
			records: dataset,
			InfiniteScroll: { raw: true },
			MaxRows: { raw: null },
			SubGridHeight: { raw: null },
			HighlightValue: { raw: "" },
			HighlightColor: { raw: "" },
			EnableLookupLinks: { raw: true },
			EnableInlineEdit: { raw: true },
			LookupFilterConfig: { raw: "" },
			LookupFilterConfig2: { raw: "" },
			LookupFilterConfig3: { raw: "" },
			LookupFilterConfig4: { raw: "" },
			LookupFilterConfig5: { raw: "" },
		},
		mode: { isVisible: true, allocatedWidth: 800, allocatedHeight: 400 },
		updatedProperties: [],
		resources: {},
		navigation: { openAlertDialog: async (): Promise<void> => undefined },
	} as unknown as TestContext;
	const testableControl = control as unknown as {
		onSelectAllRecords: (request?: SelectAllRequest) => Promise<SelectAllResult | void>;
		pageLoadWaiters: Set<unknown>;
		renderGrid: () => void;
		records: Record<string, unknown>;
		sortedRecordsIds: string[];
	};
	control.context = context;
	testableControl.records = dataset.records;
	testableControl.sortedRecordsIds = dataset.sortedRecordIds;
	testableControl.renderGrid = (): void => undefined;
	return { control, testableControl, dataset, context, selectedIds };
}

async function run(): Promise<void> {
	const readonlyControl = new CMModelDrivenGrid();
	let readonlyWrites = 0;
	readonlyControl.context = {mode:{isControlDisabled:true}, parameters:{EnableInlineEdit:{raw:true},records:{getTargetEntityType:()=> 'account',columns:[{name:'name',dataType:'SingleLine.Text'}],refresh:()=>undefined}},webAPI:{updateRecord:async()=>{readonlyWrites++;}},navigation:{openAlertDialog:async()=>undefined}} as unknown as TestContext;
	let readonlyRejected = false;
	try {await readonlyControl.onUpdateCell('00000000-0000-4000-8000-000000000001','name','changed','SingleLine.Text');} catch {readonlyRejected=true;}
	assert(readonlyRejected && readonlyWrites===0,'A disabled control must reject edits before calling Dataverse.');
	readonlyControl.context.mode.isControlDisabled=false;
	readonlyControl.context.parameters.EnableInlineEdit.raw=false;
	readonlyRejected=false;
	try {await readonlyControl.onUpdateCell('00000000-0000-4000-8000-000000000001','name','changed','SingleLine.Text');} catch {readonlyRejected=true;}
	assert(readonlyRejected && readonlyWrites===0,'Inline-edit configuration must be enforced in the save path too.');
	const control = new CMModelDrivenGrid();
	const testableControl = control as unknown as {
		parseInlineEditValue: (value: string, dataType: string) => string | number | boolean | null;
	};
	assert(testableControl.parseInlineEditValue('  text\n ', 'SingleLine.Text') === '  text\n ', 'Text whitespace must not be silently trimmed.');
	assert(testableControl.parseInlineEditValue('', 'Multiple') === null, 'An empty text editor must still clear its value.');
	for (const type of ['FP','Double','Float','Money']) assert(testableControl.parseInlineEditValue(' 12.5 ',type) === 12.5, `${type} must save a number, not numeric text.`);
	await testLookupWriteSafety();
	await testLookupDialogSafety();

	assertThrows(
		() => testableControl.parseInlineEditValue("12.5", "Whole.None"),
		"whole number"
	);
	assertThrows(
		() => testableControl.parseInlineEditValue("12abc", "Decimal"),
		"numeric value"
	);
	assert(
		testableControl.parseInlineEditValue("12.5", "Decimal") === 12.5,
		"A valid decimal value should be preserved."
	);
	assert(
		testableControl.parseInlineEditValue("2026-09-04", "DateAndTime.DateOnly") === "2026-09-04",
		"A date-only value should not be converted through a timezone."
	);
	assertThrows(
		() => testableControl.parseInlineEditValue("2026-02-30", "DateAndTime.DateOnly"),
		"valid date"
	);

	let capturedQuery = "";
	control.context = {
		parameters: {
			records: {},
			LookupFilterConfig: { raw: "" },
			LookupFilterConfig2: { raw: "" },
			LookupFilterConfig3: { raw: "" },
			LookupFilterConfig4: { raw: "" },
			LookupFilterConfig5: { raw: "" },
		},
		utils: {
			getEntityMetadata: async () => ({
				EntitySetName: "emails",
				PrimaryIdAttribute: "activityid",
				PrimaryNameAttribute: "subject",
				Attributes: { getAll: () => [] },
			}),
		},
		webAPI: {
			retrieveMultipleRecords: async (_entityType: string, query: string) => {
				capturedQuery = query;
				return { entities: [{ activityid: "id-1", subject: "Welcome" }] };
			},
		},
	} as unknown as typeof control.context;

	const results = await control.onSearchLookup("email", "Wel");
	assert(capturedQuery.includes("$select=activityid,subject"), "Lookup search should select the metadata primary key.");
	assert(results.length === 1 && results[0].id === "id-1", "Lookup search should map the metadata primary key.");
	await control.onSearchLookup("email", "O'Brien & Sons #1 + 20%");
	const searchOptions = new URLSearchParams(capturedQuery);
	let searchOptionCount = 0;
	searchOptions.forEach(() => { searchOptionCount++; });
	assert(searchOptionCount === 4, "Special characters in a lookup name must not create extra URL query options.");
	assert(
		searchOptions.get("$filter") === "(contains(subject,'O''Brien & Sons #1 + 20%'))",
		"Lookup search must preserve URL-special characters and escape OData apostrophes."
	);
	const lookupTokenControl = control as unknown as {
		getLookupTokenGuidFromLookupObject: (raw: unknown) => Promise<string | null>;
	};
	control.context.webAPI.retrieveMultipleRecords = async (_entity, query) => {
		capturedQuery = query ?? "";
		return { entities: [
			{ activityid: "00000000-0000-0000-0000-000000000001" },
			{ activityid: "00000000-0000-0000-0000-000000000002" },
		], nextLink: '' };
	};
	assert(
		await lookupTokenControl.getLookupTokenGuidFromLookupObject({ etn: "email", name: "Shared & Name" }) === null,
		"Lookup token fallback must not choose an arbitrary record when names are duplicated."
	);
	assert(new URLSearchParams(capturedQuery).get("$top") === "2", "Name fallback must check for duplicate matches.");

	const nestedFilter = {
		filterOperator: 0,
		conditions: [{ attributeName: "statecode", conditionOperator: 0, value: "0" }],
		filters: [
			{
				filterOperator: 1,
				conditions: [
					{ attributeName: "category", conditionOperator: 0, value: "A" },
					{ attributeName: "category", conditionOperator: 0, value: "B" },
				],
			},
		],
	} as ComponentFramework.PropertyHelper.DataSetApi.FilterExpression;

	const withColumnFilter = updateColumnFilter(nestedFilter, "name", "equals", "Alpha");
	assert(!!withColumnFilter, "Adding a column filter should produce a filter expression.");
	assert(
		getFilterColumnNames(withColumnFilter).sort().join(",") === "category,name,statecode",
		"Adding a column filter should preserve nested filter columns."
	);
	assert(
		getColumnConditions(withColumnFilter, "category").length === 2,
		"Nested conditions should remain addressable after adding a column filter."
	);

	const clearedNested = updateColumnFilter(withColumnFilter, "category", "clear");
	assert(
		getFilterColumnNames(clearedNested).sort().join(",") === "name,statecode",
		"Clearing a column should remove matching conditions recursively and preserve unrelated conditions."
	);

	const blankAndValue = updateColumnFilter(undefined, "name", "in", ["", "Alpha", "Beta"]);
	assert(blankAndValue?.filterOperator === 1, "Blank plus value filtering should use an OR group.");
	const blankConditions = getColumnConditions(blankAndValue, "name");
	assert(
		blankConditions.some((condition) => condition.conditionOperator === 12) &&
		blankConditions.some((condition) => condition.conditionOperator === 8 && Array.isArray(condition.value)),
		"Blank plus values should be represented as Null OR In, rather than In with an empty string."
	);
	assert(
		blankConditions.every((condition) => !Array.isArray(condition.value) || !condition.value.includes("")),
		"An empty string must not be sent inside an In value list."
	);
	const blankAndValueWithExisting = updateColumnFilter(
		{ filterOperator: 0, conditions: [{ attributeName: "statecode", conditionOperator: 0, value: "0" }] },
		"name",
		"in",
		["", "Alpha", "Beta"]
	);
	assert(
		blankAndValueWithExisting?.filterOperator === 0 &&
		getColumnConditions(blankAndValueWithExisting, "statecode").length === 1 &&
		getColumnConditions(blankAndValueWithExisting, "name").some((condition) => condition.conditionOperator === 12),
		"A blank plus value filter must stay grouped under the existing AND expression."
	);
	const blankValueGroup = blankAndValueWithExisting?.filters?.find(
		(group) => group.filterOperator === 1 && getColumnConditions(group, "name").length > 0
	);
	assert(
		blankValueGroup !== undefined,
		"A blank plus value filter must preserve its nested OR group beside unrelated AND conditions."
	);

	const lookupFilter = updateColumnFilter(undefined, "primarycontactid", "equals", "00000000-0000-0000-0000-000000000001");
	assert(
		getColumnConditions(lookupFilter, "primarycontactid")[0]?.attributeName === "primarycontactid",
		"Lookup filters should use the real lookup attribute for raw GUID values."
	);

	// Dataverse clears the filter collection while sorting is reconfigured. The control
	// must restore the complete expression before refreshing the dataset.
	const sortControl = createPagingControl();
	sortControl.dataset.columns = [{ name: "name", disableSorting: false }];
	const sortFilter = updateColumnFilter(undefined, "name", "equals", "Alpha");
	let activeSortFilter = sortFilter;
	sortControl.dataset.filtering = {
		getFilter: (): unknown => activeSortFilter,
		setFilter: (filter: unknown): void => {
			activeSortFilter = filter as typeof sortFilter;
		},
	};
	const sorting = sortControl.dataset.sorting;
	const originalPop = sorting.pop.bind(sorting);
	const originalPush = sorting.push.bind(sorting);
	sorting.pop = (): unknown => {
		activeSortFilter = undefined;
		return originalPop();
	};
	sorting.push = (...items: unknown[]): number => {
		activeSortFilter = undefined;
		return originalPush(...items);
	};
	const sortTestable = sortControl.control as unknown as {
		onSort: (name: string, direction: "asc" | "desc" | "none", additive?: boolean) => void;
	};
	sortTestable.onSort("name", "asc");
	assert(activeSortFilter === sortFilter, "Sorting should restore the existing dataset filter.");
	sortTestable.onSort("amount", "desc", true);
	assert(JSON.stringify(Array.from(sorting)) === JSON.stringify([{ name: "name", sortDirection: 0 }, { name: "amount", sortDirection: 1 }]), "Additive sorting must retain priority order.");
	sortTestable.onSort("name", "desc", true);
	assert(JSON.stringify(Array.from(sorting)) === JSON.stringify([{ name: "name", sortDirection: 1 }, { name: "amount", sortDirection: 1 }]), "Changing a sort direction must not change its priority.");
	sortTestable.onSort("name", "none");
	assert(JSON.stringify(Array.from(sorting)) === JSON.stringify([{ name: "amount", sortDirection: 1 }]), "Clear sort must remove only its column.");
	assert(activeSortFilter === sortFilter, "Multi-column sorting must retain filters.");

	// A synchronous host exception must not leave the visual page counter advanced,
	// otherwise a retry skips the failed page.
	const failedNavigation = createPagingControl();
	const navigationTestable = failedNavigation.control as unknown as {
		currentPage: number;
		loadFirstPage: () => void;
		loadNextPage: () => Promise<void>;
		loadPreviousPage: () => void;
	};
	failedNavigation.dataset.paging.loadExactPage = (): void => {
		throw new Error("synthetic navigation failure");
	};
	failedNavigation.dataset.paging.loadNextPage = () => failedNavigation.dataset.paging.loadExactPage(3);
	navigationTestable.currentPage = 2;
	let nextFailure = '';
	try { await navigationTestable.loadNextPage(); } catch (error) { nextFailure = String(error); }
	assert(nextFailure.includes('synthetic navigation failure'), 'Next-page failure must reject so scrolling can offer Retry.');
	assert(navigationTestable.currentPage === 2, "Next-page failure should restore the prior page counter.");
	assertThrows(() => navigationTestable.loadPreviousPage(), "synthetic navigation failure");
	assert(navigationTestable.currentPage === 2, "Previous-page failure should restore the prior page counter.");
	assertThrows(() => navigationTestable.loadFirstPage(), "synthetic navigation failure");
	assert(navigationTestable.currentPage === 2, "First-page failure should restore the prior page counter.");

	// The host can synchronously call updateView from loadNextPage and omit updatedProperties.
	// The waiter must be registered before the load call, and the final page restoration
	// must also complete through the same path.
	const synchronousPageLoad = createPagingControl();
	const busyStates: boolean[] = [];
	const busyControl = synchronousPageLoad.control as unknown as { selectAllAcrossPagesInProgress: boolean };
	synchronousPageLoad.testableControl.renderGrid = () => {
		busyStates.push(busyControl.selectAllAcrossPagesInProgress);
	};
	let synchronousNextPageCalls = 0;
	const setSynchronousPage = (page: number): void => {
		synchronousPageLoad.dataset.records = page === 1 ? { r1: {} } : { r2: {} };
		synchronousPageLoad.dataset.sortedRecordIds = page === 1 ? ["r1"] : ["r2"];
		synchronousPageLoad.dataset.paging.firstPageNumber = page;
		synchronousPageLoad.dataset.paging.lastPageNumber = page;
		synchronousPageLoad.dataset.paging.hasPreviousPage = page > 1;
		synchronousPageLoad.dataset.paging.hasNextPage = page < 2;
		// No updatedProperties flag is supplied here by design.
		synchronousPageLoad.control.updateView(synchronousPageLoad.context);
	};
	synchronousPageLoad.dataset.paging.loadNextPage = (): void => {
		synchronousNextPageCalls++;
		setSynchronousPage(2);
	};
	synchronousPageLoad.dataset.paging.loadExactPage = (page): void => setSynchronousPage(page);
	const synchronousSelection = await synchronousPageLoad.testableControl.onSelectAllRecords();
	assert(busyStates[0] === true && busyStates[busyStates.length - 1] === false,
		"Select-all should render a busy state before paging and release it after restoration.");
	assert(synchronousNextPageCalls === 1, "Select-all should request the next page exactly once.");
	assert(
		!!synchronousSelection && synchronousSelection.ids.join(",") === "r1,r2" && synchronousPageLoad.selectedIds.length === 0,
		"Select-all should include records returned by a synchronous updateView callback."
	);
	assert(
		synchronousPageLoad.testableControl.pageLoadWaiters.size === 0,
		"Synchronous page loads should not leave a waiter behind after page restoration."
	);

	// A synchronous paging exception must reject and clean up its waiter immediately;
	// otherwise the selection operation would retain a 20-second timer and stale callback.
	const failedPageLoad = createPagingControl();
	const setFailedPage = (page: number): void => {
		failedPageLoad.dataset.records = page === 1 ? { r1: {} } : { r2: {} };
		failedPageLoad.dataset.sortedRecordIds = page === 1 ? ["r1"] : ["r2"];
		failedPageLoad.dataset.paging.firstPageNumber = page;
		failedPageLoad.dataset.paging.lastPageNumber = page;
		failedPageLoad.dataset.paging.hasPreviousPage = page > 1;
		failedPageLoad.dataset.paging.hasNextPage = page < 2;
		failedPageLoad.control.updateView(failedPageLoad.context);
	};
	failedPageLoad.dataset.paging.loadNextPage = (): void => {
		throw new Error("synthetic paging failure");
	};
	failedPageLoad.dataset.paging.loadExactPage = (page): void => setFailedPage(page);
	let failedSelectionRejected = false;
	try {
		await failedPageLoad.testableControl.onSelectAllRecords();
	} catch (error) {
		failedSelectionRejected = error instanceof Error && error.message.includes("synthetic paging failure");
	}
	assert(failedSelectionRejected, "A synchronous paging error should reject select-all with the host error.");
	assert(
		failedPageLoad.testableControl.pageLoadWaiters.size === 0,
		"A synchronous paging error must not leave a stale page-load waiter."
	);

	// Destroying the control while a page is pending must cancel selection and its waiter;
	// it must not select a partial page or attempt to restore paging after unmount.
	const destroyedDuringPageLoad = createPagingControl();
	destroyedDuringPageLoad.dataset.paging.loadNextPage = (): void => undefined;
	const pendingSelection = destroyedDuringPageLoad.testableControl.onSelectAllRecords();
	globalThis.setTimeout(() => destroyedDuringPageLoad.control.destroy(), 0);
	await pendingSelection;
	assert(
		destroyedDuringPageLoad.selectedIds.length === 0,
		"Destroying during select-all must not apply a partial selection."
	);
	assert(
		destroyedDuringPageLoad.testableControl.pageLoadWaiters.size === 0,
		"Destroying during select-all must reject and remove its pending page waiter."
	);

	// Standard (non-hierarchy) dataset paging has the same cap, including unknown totals.
	const capped = createPagingControl();
	let cappedPage = 1, cappedNextCalls = 0;
	const capIds = Array.from({length:2000},(_,index)=>`cap-${index + 1}`);
	const applyCappedPage = (page:number) => {
		cappedPage = page;
		const batch = capIds.slice((page - 1)*75,page*75);
		capped.dataset.records = Object.fromEntries(batch.map(id=>[id,{}]));
		capped.dataset.sortedRecordIds = batch;
		capped.dataset.paging.firstPageNumber = page; capped.dataset.paging.lastPageNumber = page;
		capped.dataset.paging.hasPreviousPage = page > 1; capped.dataset.paging.hasNextPage = page*75<2000;
		capped.dataset.paging.totalResultCount = -1;
		capped.control.updateView(capped.context);
	};
	capped.dataset.paging.loadExactPage = applyCappedPage;
	capped.dataset.paging.loadNextPage = () => {cappedNextCalls++;applyCappedPage(cappedPage + 1);};
	applyCappedPage(1);
	const capResult = await capped.testableControl.onSelectAllRecords();
	assert(!!capResult && capResult.ids.length === 1000 && capResult.ids[999] === 'cap-1000' && capResult.limitReached,'Standard dataset Select All must return the first 1,000 IDs, not reject the whole 2,000-row view.');
	assert(cappedNextCalls === 13 && cappedPage === 1,'Standard Select All must stop after reaching its cap and restore the original page.');
	assert(capped.selectedIds.length === 0,'The loader must not publish intermediate selections.');
	capped.control.setSelectedRecords(capResult!.ids);
	capped.dataset.loading = true; capped.control.updateView(capped.context);
	capped.dataset.loading = false; capped.control.updateView(capped.context);
	assert(capped.selectedIds.length === 0,'An external dataset refresh must clear stale standard-grid selection.');

	const cancelled = new AbortController();
	capped.dataset.paging.loadNextPage = () => {applyCappedPage(cappedPage + 1);cancelled.abort();};
	const cancelledResult = await capped.testableControl.onSelectAllRecords({signal:cancelled.signal});
	assert(!cancelledResult && capped.selectedIds.length === 0,'Cancelling standard Select All must not return or select a partial result.');
	console.log("Local core-logic tests passed.");
}

async function testLookupWriteSafety(): Promise<void> {
	const rowId='00000000-0000-4000-8000-000000000001';
	const targetId='00000000-0000-4000-8000-000000000002';
	const create = (fail: (attempt: number) => void) => {
		const grid=new CMModelDrivenGrid();
		let writes=0;
		grid.context={mode:{isControlDisabled:false},parameters:{EnableInlineEdit:{raw:true},records:{getTargetEntityType:()=> 'account',columns:[{name:'lookup',dataType:'Lookup.Simple'}],refresh:()=>undefined}},webAPI:{updateRecord:async()=>{writes++;fail(writes);}},navigation:{openAlertDialog:async()=>undefined}} as unknown as TestContext;
		const testable=grid as unknown as {resolveLookupBindNames:()=>Promise<string[]>;getEntityMeta:()=>Promise<{entitySetName:string}>;childRecords:Record<string,unknown>};
		testable.resolveLookupBindNames=async()=>['lookup','validLookup'];
		testable.getEntityMeta=async()=>({entitySetName:'contacts'});
		return {grid,testable,writes:()=>writes};
	};
	for (const message of ['Missing privilege','Business rule rejected','Network timeout','429 throttled','412 conflict',"Could not find a property named 'otherField'"]) {
		const test=create(()=>{throw new Error(message);});
		let rejected=false;
		try {await test.grid.onUpdateLookupCell(rowId,'lookup',targetId,'contact');} catch {rejected=true;}
		assert(rejected && test.writes()===1, `Do not replay uncertain or rejected lookup writes: ${message}`);
	}
	const fallback=create(attempt=>{if(attempt===1) throw new Error("An undeclared property 'lookup' was found in the payload.");});
	await fallback.grid.onUpdateLookupCell(rowId,'lookup',targetId,'contact');
	assert(fallback.writes()===2,'Only a confirmed rejected bind name may try the next metadata candidate.');
	const disabled=create(()=>undefined);
	disabled.testable.resolveLookupBindNames=async()=>{disabled.grid.context.mode.isControlDisabled=true;return ['lookup'];};
	let rejected=false;
	try {await disabled.grid.onUpdateLookupCell(rowId,'lookup',targetId,'contact');} catch {rejected=true;}
	assert(rejected && disabled.writes()===0,'Recheck read-only state after asynchronous metadata reads.');
	const child=create(()=>undefined);
	child.testable.childRecords[rowId]={};
	child.grid.context.parameters.DisableChildEditing={raw:true} as ComponentFramework.PropertyTypes.TwoOptionsProperty;
	rejected=false;
	try {await child.grid.onUpdateLookupCell(`{${rowId.toUpperCase()}}`,'lookup',targetId,'contact');} catch {rejected=true;}
	assert(rejected && child.writes()===0,'Read-only children must be guarded with normalized record IDs.');
}

async function testLookupDialogSafety(): Promise<void> {
	const rowId = '00000000-0000-4000-8000-000000000001';
	const targetId = '00000000-0000-4000-8000-000000000002';
	const create = () => {
		const grid = new CMModelDrivenGrid();
		const seen = {dialogs:0,queries:[] as string[],writes:0};
		grid.context = {mode:{isControlDisabled:false},parameters:{EnableInlineEdit:{raw:true}, LookupFilterConfig:{raw:JSON.stringify({byColumn:{lookup:'statecode eq 0'}})},records:{getViewId:()=> 'view-1',getTargetEntityType:()=> 'account',columns:[{name:'lookup',dataType:'Lookup.Simple'}]}},
			utils:{getEntityMetadata:async()=>({EntitySetName:'cm_tags',PrimaryIdAttribute:'cm_tagid',PrimaryNameAttribute:'name',Attributes:{getAll:()=>[]}}),lookupObjects:async(options:ComponentFramework.UtilityApi.LookupOptions)=>{seen.dialogs++;assert(options.allowMultiSelect === false && options.entityTypes.join() === 'cm_tag','Picker must be single-select for the correct table.');return [{id:`{${targetId}}`,entityType:'cm_tag',name:'From dialog'}];}},
			webAPI:{retrieveMultipleRecords:async(_entity:string,query:string)=>{seen.queries.push(query);return {entities:[{cm_tagid:targetId,name:'Verified name'}]};},updateRecord:async()=>{seen.writes++;}}} as unknown as TestContext;
		return {grid,seen};
	};
	const good=create();
	const chosen=await good.grid.onBrowseLookup('cm_tag',"O'Brien & Sons",'lookup',rowId);
	assert(chosen?.id === targetId && chosen.name === 'Verified name','Picker must normalize ID and use verified metadata name, including custom tables using name.');
	const params = new URLSearchParams(good.seen.queries[0]);
	assert(params.get('$select') === 'cm_tagid,name','Authoritative name metadata must not be replaced by a publisher convention.');
	assert(params.get('$filter') === `(cm_tagid eq ${targetId}) and (statecode eq 0)`,'Full lookup must enforce both exact ID and configured restrictions.');
	assert(good.seen.writes === 0,'Opening and choosing a lookup must not perform its own duplicate update.');
	const cancelled=create(); cancelled.grid.context.utils.lookupObjects=async()=>[];
	assert(await cancelled.grid.onBrowseLookup('cm_tag','','lookup',rowId) === undefined && cancelled.seen.queries.length === 0,'Cancel must never clear or update a lookup.');
	const reject = async(test:ReturnType<typeof create>,label:string) => {
		let failed=false; try {await test.grid.onBrowseLookup('cm_tag','','lookup',rowId);} catch {failed=true;}
		assert(failed && test.seen.writes===0,label);
	};
	const restricted=create(); restricted.grid.context.webAPI.retrieveMultipleRecords=async()=>({entities:[],nextLink:''});
	await reject(restricted,'An out-of-filter record must be rejected.');
	const disabled=create(); disabled.grid.context.mode.isControlDisabled=true;
	await reject(disabled,'Read-only lookup cannot open.'); assert(disabled.seen.dialogs===0,'Read-only check must run before opening dialog.');
	const changed=create(); changed.grid.context.utils.lookupObjects=async()=>{changed.grid.context.mode.isControlDisabled=true;return [{id:targetId,entityType:'cm_tag'}];};
	await reject(changed,'Read-only state must be rechecked after the dialog.');
	const moved=create(); moved.grid.context.utils.lookupObjects=async()=>{moved.grid.context.parameters.records.getViewId=()=> 'view-2';return [{id:targetId,entityType:'cm_tag'}];};
	await reject(moved,'A dialog result from a previous view must not be applied.');
	const wrong=create(); wrong.grid.context.utils.lookupObjects=async()=>[{id:targetId,entityType:'contact'}];
	await reject(wrong,'Wrong-table results must be rejected.');
	const invalid=create(); invalid.grid.context.utils.lookupObjects=async()=>[{id:"bad'guid",entityType:'cm_tag'}];
	await reject(invalid,'Malformed lookup IDs must not reach an OData query.');
	const afterRead=create(); afterRead.grid.context.webAPI.retrieveMultipleRecords=async()=>{afterRead.grid.context.mode.isControlDisabled=true;return {entities:[{cm_tagid:targetId,name:'Verified'}],nextLink:''};};
	await reject(afterRead,'Read-only state must be rechecked after result validation.');
}

void run().catch((error: unknown) => {
	console.error("Local core-logic tests failed:", error);
	process.exitCode = 1;
});
