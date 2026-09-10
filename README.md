# DrivenGrid

DrivenGrid (`CraigMansell.CMModelDrivenGrid`) is a Power Apps component framework (PCF) dataset control for
model-driven app views and subgrids. It started from Microsoft's model-driven grid
sample and now adds editable cells, lookup links and lookup search, richer filters,
multi-cell copy, record highlighting, configurable paging, and optional infinite
scrolling.

**Candidate 1.0.111 / unmanaged solution 1.0.90.0 is not deployed.** The sandbox
still runs 1.0.110 / 1.0.89.0. See the [full updated feature comparison](FULL-FEATURE-COMPARISON-2026-09-10.md)
for local checks, earlier live evidence and remaining native-grid differences.

## Main capabilities

- Sort and filter dataset columns, including a value picker and active-filter chips.
- Edit supported scalar, choice, Boolean, and lookup columns in place.
- Open lookup values as related Dataverse records.
- Select rows, select records across dataset pages, and copy a selected cell range as TSV.
- Highlight rows using the bound `HighlightIndicator` property set.
- Use normal paging or automatically load additional pages.
- Refresh the dataset when the browser tab or control becomes visible again.
- Group a complete grid level and display sum, average, minimum and maximum.
- Move columns, enter precise content widths, and use a compact responsive list.
- Configure selection/filter/sort capabilities, choice colours, type icons and labels.

Inline editing uses the Dataverse Web API and still respects the current user's table
and column privileges. Related-table/aliased columns, unsupported data types, and
lookup columns whose target table cannot be resolved remain read-only.

The filter value picker lists values from records already loaded into the grid;
it is not a server-wide list of distinct values. Labels are displayed to the user,
but filters use the underlying values (including lookup IDs and choice codes).
Text-matching operators are available for text columns, not numeric/lookup fields.
Text columns offer all ten requested Power Apps-style conditions, including negative
contains/prefix/suffix and data-presence checks. Advanced root text/presence checks retain
the bound subgrid scope and native selection; up to 5,000 candidate rows are checked
through Dataverse, with a clear narrowing message above that limit. This advanced
root-resolution path currently requires a public saved view; personal-view failures
keep the previous filter and show an error.

The column menu now includes type-specific sorting and a two-condition AND/OR
filter editor. Numeric columns offer comparisons and ranges; date columns include
calendar and relative-date filters. Selecting a new sort replaces the previous
sort while retaining filters. See [filter and sorting behaviour](FILTERS-AND-SORTING.md) for the
supported options and differences from the native Power Apps grid and Excel.

## Configuration

### Expandable child rows

Child panels retain loaded rows, nested expansion, filter/sort settings and scroll position when collapsed and reopened. This is in-memory state, not saved across form navigation or reload. Refreshing data can reset paging; parent filtering that removes a row also removes its panel. Row virtualization stays enabled after expansion: the data/state cache survives recycled row DOM. Loaded data can still accumulate in memory; this is not a bounded-cache guarantee.

Set `DisableChildEditing=true` to make child cells read-only while leaving parents editable. Child panels open directly onto the column headers, without a loaded-count heading. Right/down chevrons sit beside always-visible square checkboxes in a compact leading gutter. The child-region name and read-only status remain accessible to screen readers. Child panels contain no Add, Refresh, Delete or other record-action buttons: the single subgrid command bar is the required action surface. The legacy `EnableChildCommands` and `ChildContextLookup` inputs remain for existing form compatibility but no longer enable embedded commands. Refreshing the host dataset refreshes expanded children and clears stale selection; a failed child load directs users to that command-bar Refresh action.

Child checkboxes and multi-selection use one hierarchy level at a time. Selecting children clears parents/siblings; collapsing an ancestor clears hidden selection. The legacy setup fetches children separately, so their IDs are not recognised by the native command bar. **Control 1.0.104 adds opt-in `EnableNativeChildSelection`: Power Apps retrieves selected children through its primary dataset before their IDs are published to native commands.** The control checks selection readback and fails safely without substituting a parent. The footer reports confirmed native selection, including selected children. Do not treat the local harness's simulated toolbar as proof of native command integration. Successful inline saves request a refresh; Item Template cost/totals scripts still require a signed-in integration test.

The owner explicitly rejected a separate child toolbar: all record actions must come from the existing subgrid command bar. Do not reintroduce a child toolbar, inject records into platform internals or bypass native permission rules. **Add existing child** scoped automatically to the expanded parent, separate child-view column layouts and different-table children remain unimplemented. The existing host Add Existing command retains its normal subgrid relationship scope; selecting a child does not re-parent it or change that scope.

Enable `EnableChildRows`, set `ChildParentLookup` to the self-referencing lookup logical name, and set `ChildViewId` to a public view GUID of the same table. The child view must include the same columns as the parent grid and must not exclude children. For Item Template Options use `cm_parenttemplateoption` and the existing **Sub Options** view.

For native child commands, additionally enable `EnableNativeChildSelection` and bind the main dataset to a **dedicated view that includes both roots and children**. Do not enable this setting on a view that filters the parent lookup to null. The sandbox deployment helper clones Options as **DrivenGrid hierarchy options**, preserves its active-record condition, ordering and nine visible columns, and adds the hidden parent/name attributes. Only the target Item Template subgrid is rebound; the shared Options view and relationship remain unchanged. The PCF applies root-only filters internally, temporarily admitting explicitly selected child IDs and hiding those child rows from the root display. Root sorting/filtering and child-panel filters remain independent. The current candidate selects the first 1,000 matching records when a level is larger; filter or sort to choose another subset. Native selection preparation is serialised, cancelled by newer requests and checked for missing/rejected targets. Local strict-host, root-cache and refresh tests pass. Signed-in sandbox checks for the earlier 1.0.104 release confirmed native child commands, exact-child Edit navigation, multi-selection, level switching, collapse cleanup and independent root/child filtering. Native asynchronous permission rules can take a moment to settle after selection. Destructive and data-changing commands were not executed during that read-only smoke test. Candidate 1.0.111 still needs its own signed-in acceptance.

Children load on expansion, with independent sorting, filtering, editing, refresh and infinite scrolling. Filters on the main grid affect parents only; use each child grid's menus to filter its children. Nested levels are supported with cycle protection. Candidate Child Select All pages the matching direct children up to 1,000, not just loaded children. Child paging retains server continuation cookies and has a 1,000-page safety guard; narrow filters if it is reached. Expand arrows appear only on rows with accessible children in the configured child view. Distinct-parent checks are batched and cached per refresh, without downloading each child grid. Leaves retain their aligned selection checkbox; native Refresh rechecks added/removed children. A failed existence check shows an error and never pretends that the row is confirmed empty.

### Grouping, totals and compact layout (candidate)

Use **Column options** for Move left/right, Column width, Group by this column,
and numeric Totals. Width is the content width (80–2,000 pixels); cell padding is
additional. Spare space still fills the grid. Grouping and totals request all
matching pages of the current level; the grid never labels a partial result as a
complete total. They do not open every child's branch. Pause/Resume controls the
summary fetch; normal scrolling can still fetch a page. Failed loads withhold totals.
Use Clear grouping and totals to return to ordinary virtualized browsing.

These are client-computed summaries, not server aggregate queries. They require
infinite scrolling and may take time/memory for large results. Numbers use raw
values without currency conversion; do not sum mixed transaction currencies as
though they were one currency. Date groups use exact values, not native date intervals.

`ReflowBehavior` accepts `auto`, `grid` or `list`. Auto uses cards below 480 pixels.
The list supports 1–10 fields, optional labels, banding and initials (not photographs).
Its Columns menu retains filtering/sorting; this is a custom behaviour, not a claim
that every native list restriction is duplicated.

### Properties

| Property | Purpose |
| --- | --- |
| `HighlightIndicator` | Dataset column compared with `HighlightValue`. |
| `HighlightValue` / `HighlightColor` | Value and CSS colour used to highlight matching rows. |
| `EnableLookupLinks` | Opens lookup values as related records. |
| `EnableInlineEdit` | Enables or disables in-grid editing. |
| `LookupFilterConfig` through `LookupFilterConfig5` | One JSON lookup-filter configuration split across fields when needed. |
| `InfiniteScroll` | Loads the next page when the end-of-grid sentinel becomes visible. Enabled by default. |
| `MaxRows` | Optional Dataverse dataset page size. |
| `SubGridHeight` | Optional fixed height, in pixels, for form subgrids. |
| `EnableGrouping` / `EnableAggregation` | Enable complete-level groups and numeric summaries. Defaults true in this candidate. |
| `EnableColumnMoving` | Header drag/reorder and Move left/right. |
| `EnableMultipleSelection` / `EnableRangeSelection` | Multiple-row selection and mouse rectangular copy, independently switchable. |
| `EnableFiltering` / `EnableSorting` | Enable new header conditions/sorts. Existing filters can still be cleared. |
| `ShowTypeIcons` / `ShowChoiceColors` | Header type icons and metadata-based choice colour pills. |
| `ShowStatusIcons` | State/Status value icons, not Microsoft's editing-status column. |
| `ReflowBehavior` / `ListColumns` | Auto/grid/list layout and 1–10 list fields. |
| `ShowAvatar` / `ShowListColumnNames` / `BandedListRows` | List initials, labels and alternate row bands. |
| `AccessibleLabel` | Accessible name for the grid. |

Lookup-filter JSON supports `default`, `byEntity`, `byColumn`, and
`byColumnAndEntity`. A non-JSON value is treated as a global OData filter clause.
Dynamic clauses can use row tokens such as `{{attribute}}`, `{{attribute.id}}`,
`{{attribute.name}}`, and `{{attribute.entityType}}`.

## Development

Prerequisites are Node.js, npm, and the Microsoft Power Platform CLI.

### Try the interactive sample

```powershell
npm run demo
```

Open [the interactive harness](http://127.0.0.1:8183/). It renders the actual grid
with 180 disposable sample parents and text, currency, choice, lookup, date, and
Yes/No columns, including blanks. Click a column heading for its filter options,
combine filters, scroll to load more rows, or double-click a cell to edit. Scroll
horizontally to reach the remaining columns in a narrow browser panel.

**Reset sample data** restores the original records and clears filters and sorting.
The demo runs only on the local loopback interface and never writes to Dataverse.
It watches the grid source; reload the browser after source changes. Stop it with
Ctrl+C. The installed `tsx` tool supplies its bundler; no extra packages are needed.

For the requested hierarchy stress test, open
[2,000 parents with 20 children each](http://127.0.0.1:8183/?scenario=parents-2000-children-20)
or choose **Use 2,000 parents × 20 children**. This scenario has exactly 42,000
unique records, with no grandchildren. Parents arrive in 25-row batches as you
scroll; children are requested only when their parent is expanded. The visible
request counter shows child pages and rows returned. Do not use **Preload local
rows for rendering test** when checking automatic paging.

`npm run test:demo-hierarchy` checks every parent and child for correct scope,
counts, unique IDs, paging, end-of-data, filtering and sorting. It is also part of
`npm run test:local`. The harness keeps its generated fixture in local memory;
these tests do not create 42,000 Dataverse records or establish server performance
at that size. Native commands and Dataverse throttling still need signed-in tests.

The header selection checkbox selects all matching records up to **1,000**,
following the current filter and sort order, including records beyond the loaded
page. This matches Microsoft's documented
[modern-grid Select All limit](https://learn.microsoft.com/en-us/power-platform/released-versions/common-data-service/monthly-2312).
The parent-grid header selects parents only; a child-grid header selects that parent's
matching direct children only. If any records are already selected, clicking the
header clears the selection. Clicking it while selection is loading cancels the
operation. Paging does not add further records to a completed selection.

A limit message explains when only the first 1,000 matches were selected. Narrow
the filters or change the ordering to choose a different set. Native command
availability and action-specific limits still apply; the control does not bypass
Dynamics command rules. This selection update is currently local/unreleased.

Try [compact list mode](http://127.0.0.1:8183/?scenario=parents-2000-children-20&layout=list).
Fiscal demo filters simulate January–December with monthly periods; the real control
sends fiscal conditions to Dataverse, which uses the organisation's configuration.

### Validate the control

```powershell
npm run lint
npm run test:local
npm run build
```

Use `npm start` for the standard PCF test harness. `npm run start:watch` rebuilds
and reloads the harness after edits. `npm run dev:watch` watches source files and
builds locally; `npm run dev:loop` runs local tests and a build. Neither command
increments versions or deploys to Dataverse.

For sample data in the standard harness, choose
`assets/harness-filter-test-data.csv` under **records**, map `HighlightIndicator`,
then click **Apply**. The Microsoft harness exposes the manifest's bound property
set; its mock refresh does not actually apply dataset filters. Use `npm run demo`
for interactive filter and scrolling checks. Neither local harness reproduces
Dataverse security, metadata, or its Web API. Validate real permissions, lookup
metadata, filtering, and server paging in a signed-in
development model-driven app before deploying the control to production.

The solution project is explicitly configured to produce an unmanaged package.
`npm run push:dev` targets the Capricorn development environment and should only be
used while authenticated to the intended tenant.
