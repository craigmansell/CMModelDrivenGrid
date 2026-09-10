# Filtering and sorting

The grid uses a Power Apps-style surface with Excel-style column menus. It is a
custom dataset PCF, not a replacement implementation of every native-grid or Excel
feature.

## How to use it

- Click a column heading to open its menu. Drag a header divider to resize.
- Columns expand to fill spare grid width. Narrow windows retain readable widths
  and horizontal scrolling; after every column is manually resized, the last
  column fills any remaining gap.
- Sort text A to Z/Z to A, numbers smallest/largest, and dates oldest/newest.
- Choosing a sort replaces previous sort levels. **Clear sort** removes the
  selected column's sort; sorting always retains active filters.
- The main checklist uses a consistent 320 by 560px size; the condition editor
  grows and shrinks to fit its fields. Both are capped to the available host space.
  The bordered checklist aligns **Select All** with the values. Only one region
  scrolls at a time and the action buttons remain visible.
- Search the checkbox list, then select **OK** to apply the matching selections.
  **Cancel** leaves the applied filter unchanged. Unchecking every result disables
  OK; it never silently clears the filter.
- Open **Number filters**, **Date filters**, or **Filter by** for conditions. Add
  a second condition and join them with AND or OR. Separate columns combine with
  AND; selected checkbox values within one column combine with OR.
- Clear a column using its menu or filter chip. Sorting retains active filters.
  A new query starts at the top; infinite scrolling preserves position as it appends.
- Scroll near the bottom to load more records automatically. There is no Load more
  button. Paging failures show an error/recovery action instead of endless loading.

## Supported field filters

| Field | Conditions |
| --- | --- |
| Text | Equals, does not equal, contains data, does not contain data, contains, does not contain, begins with, does not begin with, ends with, does not end with |
| Number / currency | Equals, does not equal, greater/less than, greater/less than or equal, between, contains data, does not contain data |
| Date/time | 53 conditions: the 50 observed native date choices plus After, Before and Between. Includes calendar, relative hours/weeks/years, fiscal period/year and Older than X units. Date-only columns omit four hour/minute conditions. |
| Choice / Yes-No / lookup | Equals, does not equal, contains data, does not contain data, checkbox value selection |

The custom editor supports up to two conditions, including ranges. Numbers must be
valid decimals; dates use a date picker; relative counts are positive whole numbers.
Choice/lookup controls submit their raw codes/IDs, not formatted display names.
Fiscal periods accept 1–13 and fiscal years 1–9999. Period/year combinations submit
both values in Dataverse's required order. The server decides valid periods and
their dates for its actual fiscal configuration. All added conditions are covered
by compiler tests and accepted by read-only sandbox queries, not a complete
time-zone/fiscal-boundary acceptance test of every expected record ID.

## Important differences

- The checkbox picker lists values from **loaded records**, plus retained choices
  from the applied filter. It is not a full-table distinct-value query. Conditions
  submitted to Dataverse filter the dataset on the server, not just loaded rows.
- The documented model-driven PCF dataset does not support NotLike or NotNull directly.
  DrivenGrid resolves Contains data, negative/literal-wildcard text and added date
  conditions on the server.
  These filters enumerate the **actual bound grid scope**, retaining its view
  and subgrid relationship, then ask Dataverse to match those record IDs before
  applying supported ID conditions to the primary dataset. This preserves native
  commands and Dataverse text/null semantics; it never filters only the loaded page.
  The check is capped at 5,000 candidate rows and 500 pages. Above that, narrow another
  column first; errors restore the previous filter instead of showing partial results.
  This advanced root-resolution path currently requires a **public saved view**;
  personal views are not supported by it and produce an error without replacing
  the previous filter.
  Child filters use FetchXML directly and remain server-paged. Refresh/sort repeats
  root text matching; filters and expansion are in-memory, not saved views.
- Related-table advanced-find conditions, hierarchy conditions, multi-select
  choice operators, Excel colour filters, Top 10, and date-tree checkbox grouping
  are not implemented by this menu. Use the native Power Apps grid/view designer
  when those capabilities are required.
- Lookup/choice menus still have four conditions. The six label-pattern conditions
  seen in native lookup/Status Reason menus are not implemented; text matching a
  lookup GUID or numeric choice code would not reproduce label matching.
- Literal percent, underscore and left bracket are escaped in text pattern filters.
  Scoped sandbox checks returned the exact expected percent/underscore record IDs.
- In Dynamics, ordering and relative-date interpretation follow Dataverse, including
  its collation, user time zone and calendar rules. The local demo approximates these
  with local dates and label-based, stable sorting; it is not evidence of exact
  Excel/Dataverse parity. Test in the signed-in model-driven app before deployment.
- The demo has 180 disposable records, evaluates filters and sorts before taking a
  25-row page, and never connects to Dataverse. Its dates are relative to today.

Platform references: [PCF dataset operators and sorting](https://learn.microsoft.com/en-us/power-apps/developer/component-framework/reference/dataset),
[Power Apps grid capabilities](https://learn.microsoft.com/en-us/power-apps/maker/model-driven-apps/the-power-apps-grid-control),
[Excel filtering](https://support.microsoft.com/en-gb/office/filter-data-in-a-range-or-table-in-excel-01832226-31b5-4568-8806-38c37dcc180e).
