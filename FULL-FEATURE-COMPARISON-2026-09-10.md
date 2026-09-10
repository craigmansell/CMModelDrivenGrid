# DrivenGrid vs Power Apps grid — updated full review

10 September 2026. Candidate **1.0.111 / unmanaged solution 1.0.90.0**.

## Outcome

This batch adds substantial functionality and fixes the confirmed correctness issues,
but **DrivenGrid is not a complete native-grid replacement**. The code and local
harness include these changes; the sandbox still runs **1.0.110 / 1.0.89.0**.
No deployment or business-data change was made during this implementation batch.

Added: 30 date conditions; complete-level groups and totals; column movement and
width editor; compact list/reflow; choice colours/type icons; configurable selection,
filtering, sorting and accessibility labels. Corrected: read-only writes, literal
wildcards, ambiguous lookup-save retries, clipboard quoting/errors, unknown counts,
text whitespace and floating-point editing. The earlier full-view Select All work
is included in the owner-approved dependent source batch.

The important remaining gaps are different-table/independent child datasets,
lookup/status label-pattern filters, advanced lookup and specialist editors,
native event/customizer contracts, keyboard range extension, and full device,
localization, permission and live-scale acceptance. Matching menu labels is not
proof of matching execution. A third-party dataset control's public API does not
establish access to all Microsoft OneGrid host integration. Do not bridge that gap
by injecting records, overriding command rules or adding the rejected child toolbar.

## Evidence and meaning of the results

The reference inventory was captured earlier on the same day from installed
`Microsoft.PowerApps.PowerAppsOneGrid` **1.0.356**: 45 direct properties, 26 non-hidden
and 19 hidden, plus Items/ChildItems datasets. DrivenGrid now has 35 direct inputs
and one records dataset. Property counts are **not a parity percentage**.

- **Earlier paired live**: same sandbox demo parent/two children and nine columns;
  native and DrivenGrid expansion, selection, command availability, exact-child Edit,
  child numeric filter and explicit sort matched. Diagnostic form-onload suppression
  was used; normal business-handler acceptance is still required.
- **Earlier native-only live**: all 101 isolated QA options in a flat main grid;
  full menu inventory, Select All, groups and sums. Active 96 / Inactive 5; sums
  2,198.08 / 230.00 / 2,428.08 independently matched a scoped Dataverse aggregate.
- **This candidate, local**: component regressions, type/lint checks and actual
  harness UI. Local rows are fixtures, not Dataverse throughput measurements.
- **This candidate, server queries**: read-only conditions on the 101 existing QA
  records, with sandbox URL and organization locks. All 30 added date operators
  accepted; percent and underscore matches compared with independently derived IDs.
  This did not install the candidate or exercise each condition through the live UI.
- **Open** means no adequate acceptance evidence, not automatically a defect.
  **Host** means supplied by Dynamics around the PCF, not implemented by this grid.

Historical evidence remains in the CrmPackage workspace under
`outputs/qa/drivengrid-one-to-one-20260910/`; current source baseline and read-only
query script are under `outputs/qa/drivengrid-feature-parity-20260910/`. Historical
results have not been relabelled as candidate tests.

## Complete feature inventory

### Records, loading and navigation

| ID | Feature | Candidate against native grid |
| --- | --- | --- |
| D01 | View rows/columns | Reads bound columns and formatted values; earlier paired nine-column fixture matched. |
| D02 | View selector | Host display requested; view management and every form configuration open. |
| D03 | Quick Find/search | Host display requested; full hierarchy/related-field scope open. |
| D04 | Initial load/empty/error | Implemented and locally regressed; no matched cold-load benchmark. |
| D05 | Infinite scrolling | Automatic root/child paging and virtualization; earlier live paging, local 42k fixture. No 42k server test. |
| D06 | Pagination | Optional conventional root paging; full native paging/hierarchy combination open. |
| D07 | Refresh | Native refresh integrates with rows/children/selection; external-update combinations open. |
| D08 | Counts | Fixed unknown totals: loaded count with `+` while incomplete, rather than assuming 5,000. |
| D09 | Open own record | Earlier paired native Edit opened exact child; compact sample Open also targets exact child. |
| D10 | Lookup links | Boolean switch, with corrected click event. Native all/primary-only/none setting still absent. |
| D11 | Query/view changes | Local stale-request/cache tests; full host view/search matrix open. |
| D12 | Blanks/formatting | Raw blank/zero identity and supplied formatting covered; full locale/currency/timezone matrix open. |

### Selection and clipboard

| ID | Feature | Candidate against native grid |
| --- | --- | --- |
| S01 | Single parent/child | Exact identities, not proxy-parent selection; earlier paired live plus local card check. |
| S02 | Multiple children | Shared active branch; earlier paired two-child selection. |
| S03 | Select All beyond viewport | Pages matching IDs up to 1,000; local 75-to-1,000 and 2,000-row tests. Native live showed 101. |
| S04 | Selection announcements | Uses confirmed count; does not copy the observed native stale 50-row announcement. Screen-reader audit open. |
| S05 | 1,000 limit | Later-loaded row 1,001 remains unticked; boundary/cancellation tests local, candidate live acceptance open. |
| S06 | Filtered Select All | Current filter/sort scope; previous 334-match browser test and automated cases. |
| S07 | Progress/cancel | Implemented and locally tested, including stale results and view changes. |
| S08 | Clear selection | Header click clears/cancels; earlier live clear matched. |
| S09 | Parent/child scope | One active hierarchy level/branch; collapse clears hidden selection. |
| S10 | Offscreen identity | IDs survive row recycling; large native comparison remains open. |
| S11 | Recursive Select All | Not implied: selecting parents does not select every descendant. No native recursive promise established. |
| S12 | Single-select option | Added EnableMultipleSelection; tested replacement selection in list mode. |
| S13 | Rectangular copy | Mouse drag/Shift-click; quoting fixed. Native keyboard range extension remains absent. |
| S14 | Embedded tabs/newlines/quotes | Quoted TSV serializer tested. Actual paste into Excel not acceptance-tested. |
| S15 | Paste/fill/undo | Not implemented as spreadsheet features; not established as universal native capabilities either. |
| S16 | Clipboard denied/unavailable | Visible error feedback added and component-tested. Browser permission matrix open. |

### Child options

| ID | Feature | Candidate against native grid |
| --- | --- | --- |
| H01 | Expand/collapse | Earlier paired same-parent expansion; current compact layout also tested. |
| H02 | Leaf arrows | No arrow for confirmed leaves; batched existence queries, failure remains visible. |
| H03 | Child title/actions | Intentionally no loaded-count title or Add/Refresh/Delete strip. Accessible region remains. |
| H04 | Shared command bar | Earlier paired command availability/exact Edit target. Candidate requires a fresh signed-in smoke test. |
| H05 | Different child table | **Gap**: native separate child dataset; custom self-related same-table loader. |
| H06 | Independent child columns | **Gap**: custom reuses root columns; child view supplies query scope. |
| H07 | Multiple expanded branches | Implemented; local tests, full live combinations open. |
| H08 | Grandchildren | Recursive custom branches with cycle/depth guards. Native arbitrary-depth parity not established. |
| H09 | Independent filters/sorts | Earlier paired numeric child filtering/sorting; custom root/child query tests. |
| H10 | Default ordering | A child view without explicit order yielded different initial orders. Explicit Order sort matched. |
| H11 | Parent filter searching descendants | Not recursive; each level's filters are independent. |
| H12 | Child paging | Earlier live 64-child branch and local cookie/continuation tests; 1,000-page guard. |
| H13 | Child editing disabled | Configuration and host read-only guards enforced before writes; role/field-security UI tests open. |
| H14 | Virtualization retention | Data/drafts/filter/order/width/scroll retained outside recycled DOM; not persisted across reload. |
| H15 | Native command prerequisites | Requires a bound same-table view admitting selected children; extra configuration compared with native. |
| H16 | New/Add Existing association | Outer command retains normal subgrid relationship context. Child selection does not automatically set a new child's parent. |

### Filtering

| Field family | Native observed menu | Candidate |
| --- | ---: | --- |
| Text | 10 | 10; literal `%`, `_`, `[` fixed. |
| Currency | 8 | 9, with additional Between. |
| Date/time | 50 | 53: all 50 corresponding choices, plus After/Before/Between. |
| Date-only | Not separately inventoried | 49; excludes Last/Next X hours and Older than X hours/minutes. |
| Lookup | 10 | 4; six label-pattern conditions still absent. |
| Status Reason | 10 | 4; same six label-pattern gaps. Do not extrapolate to every choice field. |
| Boolean | Not live-inventoried | 4, plus loaded-values checklist; live parity open. |
| Multichoice/party list | Not in matched fixture | No complete specialist implementation. |

All 50 native date choices are accounted for:

| Family | Conditions |
| --- | --- |
| Existing 20 equivalents | On; On or after/before; Today/Yesterday/Tomorrow; This week/month/year; Next 7 days; Next X days/months; Last week/7 days/month/year; Last X days/months; Contains data; Does not contain data. |
| Added calendar 3 | Next week, month, year. |
| Added fiscal adjacent 6 | This/Next/Last fiscal period and fiscal year. |
| Added relative time 6 | Next/Last X hours, weeks and years. |
| Added relative fiscal 4 | Next/Last X fiscal periods and fiscal years. |
| Added age 6 | Older than X minutes, hours, days, weeks, months, years. |
| Added specific fiscal 5 | In fiscal year; In fiscal period; In fiscal period and year; In or after/before fiscal period and year. |

| ID | Behaviour | Candidate against native grid |
| --- | --- | --- |
| F01 | Values checklist | Search, Select All, blanks, dotted guides; loaded values plus retained choices, not server-wide distinct values. |
| F02 | Multi-value filter | Custom value selection combines with OR; native also documents multi-value text filtering. |
| F03 | Conditions | Two AND/OR conditions per column; columns AND together. Not a general nested query designer. |
| F04 | Validation | Number/date/count/range/fiscal inputs validated; period 1–13, year 1–9999, server fiscal configuration authoritative. |
| F05 | Beyond first page | Dataset/server conditions, not only visible rows; not every live operator/combination accepted. |
| F06 | Root server fallback | Public saved view, 5,000 candidates/500 pages; errors restore prior query. Applies to unsupported/presence/negative/literal/additional-date predicates. |
| F07 | Literal wildcards | Local compiler and independent sandbox percent/underscore IDs passed. Positive `[` server fixture absent. |
| F08 | Lookup candidates | Loaded checklist/basic lookup search differs from native contextual full-view candidate list. |
| F09 | Dates/time zones | Added mappings accepted by server. Full fiscal calendars, DateOnly behaviour vs format, DST and boundary-ID tests still open. |
| F10 | Advanced related-table filters | Host capability, not an equivalent custom column query builder. |
| F11 | Cancel/clear/zero/combined | Local regressions; a subset paired live. Full live cross-product open. |
| F12 | Excel extras | Colour/Top 10/date-tree/formula filters absent; not established native grid features. |

The local fiscal demo uses January–December monthly periods; the real query uses
Dataverse's organization calendar. See [filter details](FILTERS-AND-SORTING.md).
The added server operator mappings follow Microsoft's [FetchXML operators](https://learn.microsoft.com/en-us/power-apps/developer/data-platform/fetchxml/reference/operators)
and [ConditionOperator definitions](https://learn.microsoft.com/en-us/dotnet/api/microsoft.xrm.sdk.query.conditionoperator?view=dataverse-sdk-latest).

### Sorting, columns, grouping, totals

| ID | Feature | Candidate against native grid |
| --- | --- | --- |
| C01 | Single sort | Ascending/descending, numeric/date semantics; earlier paired child sorts. |
| C02 | Additive sort | Intentionally simple replace-sort menu, per owner request. Native Shift-add sorting not implemented in UI. |
| C03 | Initial view order | Supported where supplied; ensure explicit stable order/tie-breaker for a fair comparison. |
| C04 | Resize | Drag plus new 80–2,000px content-width editor; 220px yielded 244px including cell padding in browser. |
| C05 | Fill width | Spare-width distribution, minimums and overflow retained. |
| C06 | Column movement | New header drag/drop callback and Move left/right. Browser Budget moved before Project; drag gesture itself not browser-tested. |
| C07 | Edit view columns | Host designer integration, not a custom view-management implementation. |
| C08 | Groups | New raw-value single-column groups, individual/global collapse retained. Initially expanded, unlike documented native default. |
| C09 | Date intervals | **Gap**: exact-value grouping only; native date interval submenu not reproduced. |
| C10 | Totals | Sum/average/min/max; all pages of current level before displaying complete totals. No currency conversion/server aggregate plan. |
| C11 | Groups with hierarchy/paging | Custom groups require infinite scrolling; no full native combination signoff. Group header selection still obeys 1,000 cap. |
| C12 | Save layout | In-memory per-level state, not saved personal views or durable native personalization. |

Grouping/totals fetch all matching records in one level, but do not expand children.
They may consume more memory/network than server aggregates. Pause/Resume affects
summary fetching; normal viewport scrolling can still load a page. Failures withhold
totals. Raw transaction-currency sums must not be treated as converted base currency.

### Editing

| ID | Feature | Candidate against native grid |
| --- | --- | --- |
| E01 | Grid read-only | Fixed root renderer/write guard; rechecked after asynchronous lookup metadata. |
| E02 | Column read-only | Save and UI flags honoured; role/field-security/calculated-field matrix open. |
| E03 | Text/memo | Preserve whitespace now; empty input clears. No full native multiline editor. |
| E04 | Whole/decimal/currency | Typed parser and errors; metadata precision/min/max/localized input matrix open. |
| E05 | Floating point | FP UI allowlist and Float/Double/Money numeric parsing corrected/tested. Live metadata cases open. |
| E06 | Boolean | Yes/No editor; localized option labels not fully equivalent. |
| E07 | Choice | Metadata options plus new validated colour pills; dynamic/translated states open. |
| E08 | State/status | Inline transition editing blocked; use native Activate/Deactivate. New value icons are not row-save status icons. |
| E09 | Date | Date-only calendar input added; DateTime editor/timezone/locale remains partial. |
| E10 | Lookup | Search/choose/clear; 20-result cap and primary-name field. Native candidate experience broader. |
| E11 | Advanced lookup | **Gap**: MRU/view selector/create target/advanced dialog/polymorphic picker not fully implemented. |
| E12 | Customer/regarding/party list | No generic specialist inline editors. Native availability must be checked for each configuration. |
| E13 | Multi-select choice | No complete inline editor. |
| E14 | File/image/virtual/ID | Guarded out/no specialist editor; not claimed universally editable in native either. |
| E15 | Save/cancel/feedback | Enter/blur commits, Escape cancels, cell saving/error feedback; native row timing can differ. |
| E16 | Validation/concurrency | Server validation remains authoritative. Retry only confirmed invalid bind names; no replay after ambiguous failure. Full conflict/plugin matrix open. |
| E17 | Native events | **Gap/open**: direct Web API writes do not reproduce all native grid OnChange/OnSave/customizer contracts. |
| E18 | Bulk editing | Native command availability is not proof every bulk action succeeds at scale. |

### Commands and app integration

| ID | Feature | Candidate against native grid |
| --- | --- | --- |
| A01 | Child Edit | Earlier paired exact child navigation passed, no save; candidate smoke still needed. |
| A02 | Activate/Deactivate | Earlier shared availability and prior disposable action checks, not complete candidate pair. |
| A03 | Delete | Shared command, no deletion in this batch; old disposable QA is not full new acceptance. |
| A04 | Assign/Share/Email | Host commands; visibility inspected earlier, effects not tested here. |
| A05 | Flow/Report/Templates/Export | Host features; no generated documents, flows or exports executed in this review. |
| A06 | New/Add Existing | Host relationship and security context; no replacement child toolbar. |
| A07 | Associated/Show As | Earlier menu availability, full navigation combinations open. |
| A08 | Enable/display rules | Normal native rules, no bypass; additional roles/states/custom rules open. |
| A09 | Form scripts | Normal handlers must be included in final candidate live acceptance. |
| A10 | Native customizer | **Gap**: no compatible OneGrid customizer-control contract. |
| A11 | Charts/Power BI/AI | Host integrations, not newly implemented child-grid features. |

Public [dataset APIs](https://learn.microsoft.com/en-us/power-apps/developer/component-framework/reference/dataset)
and Microsoft's [grid customizer interface](https://learn.microsoft.com/en-us/power-apps/developer/component-framework/customize-editable-grid-control)
are separate from private OneGrid host integration. Native command execution must
remain on the actual command surface, evaluated for the actual selected records.

### Appearance, accessibility and devices

| ID | Feature | Candidate against native grid |
| --- | --- | --- |
| U01 | Dynamics appearance | Fluent table, square row checkboxes and chevrons; not pixel-perfect certification. |
| U02 | Compact child UI | Requested title/action-strip removal preserved. |
| U03 | Fullscreen | Intentionally absent. |
| U04 | Filter sizing | Bounded checklist, dynamic condition height, dotted guides; compact fiscal editor visually checked. Full zoom matrix open. |
| U05 | Formatting | New type icons/choice colours/state value icons; native row-status column still different. |
| U06 | Keyboard/edit | Fluent navigation and custom edit keys; complete native keyboard parity open. |
| U07 | Keyboard copy ranges | Mouse range/Shift-click only; keyboard extension and cross-page range copying remain gaps. |
| U08 | Accessibility | New configurable grid name plus existing region/expanded/status labels. No screen-reader certification. |
| U09 | Forced colours | Explicit CSS/focus styles; comprehensive visual/contrast audit open. |
| U10 | Small screens/touch/zoom | Auto cards below 480px, forced-list browser check; actual devices/touch/full zoom open. |
| U11 | List presentation | New 1–10 fields, labels, bands, initials. Custom list also exposes Columns/Select All, not all native restrictions. |
| U12 | Languages/RTL | English resources and hard-coded labels remain; full localization/RTL gap. |

### Scale, security and maintenance

| ID | Feature | Candidate against native grid |
| --- | --- | --- |
| R01 | 2,000 × 20 | Local 42,000-record fixture exhaustively covers all 2,000 branches/80 root batches; not a server benchmark. |
| R02 | Virtualization | Kept with expanded branches; candidate grouped browser check had 53 row nodes. Counts are observations, not a fixed limit. |
| R03 | Memory | Data/cache can grow; no comprehensive long-session memory/eviction benchmark. |
| R04 | Paging integrity | Cookie/scope/dedup/end-of-data tests; prior live 64-child paging. Equal-data native server scale open. |
| R05 | Failures | Timeouts/cancellation/stale results/error recovery tested locally; offline/429/concurrent-update matrix incomplete. |
| R06 | Security/scope | Signed-in Dataverse calls; strengthened write guards and removed record/query console diagnostics. Restricted-user acceptance open. |
| R07 | Offline/mobile | No custom offline contract or full mobile acceptance established. |
| R08 | Maintenance | Custom Fluent/React/PCF code must be maintained separately from Microsoft's control. |
| R09 | Documentation | Corrected stale virtualization, loaded-only selection and filter coverage statements. |
| R10 | Deployment/rollback | Candidate not installed. Fresh exact rollback, locks, unmanaged assertion, publish/readback and live smoke required. |

## Every installed native configuration declaration accounted for

| Non-hidden native property | DrivenGrid mapping |
| --- | --- |
| AccessibleLabel | Added AccessibleLabel. |
| EnableEditing | EnableInlineEdit plus host/write guards. |
| DisableChildItemsEditing | DisableChildEditing. |
| EnableFiltering | Added EnableFiltering. |
| EnableSorting | Added EnableSorting. |
| EnableGrouping | Added EnableGrouping; client complete-level implementation. |
| EnableAggregation | Added EnableAggregation; client raw-number implementation. |
| EnableColumnMoving | Added EnableColumnMoving. |
| EnableMultipleSelection | Added EnableMultipleSelection. |
| EnableRangeSelection | Added EnableRangeSelection; mouse-range implementation. |
| EnableJumpBar | Not implemented. |
| EnablePagination | Inverse InfiniteScroll choice; not identical contract. |
| EnableDropdownColor | Added ShowChoiceColors. |
| EnableStatusIcons | Native edit-row status icons not equivalent to custom ShowStatusIcons value icons. |
| EnableTypeIcons | Added ShowTypeIcons. |
| NavigationTypesAllowed | EnableLookupLinks boolean only. |
| ReflowBehavior | Added auto/grid/list. |
| ShowAvatar | Added initials. |
| ShowColumnNamesForListView | Added ShowListColumnNames. |
| EnableBandedRowsForListView | Added BandedListRows. |
| NumberOfListColumns | Added ListColumns, 1–10. |
| ContextualLookupColumnFilters | No full-view equivalent. |
| LookupFilterBeginsWith | No maker toggle; search uses contains. |
| UseFirstColumnForLookupEdits | No; target primary-name search. |
| GridCustomizerControlFullName | No native customizer contract. |
| EnableStatusColumn | No equivalent native row-status column. |

All 19 hidden declarations were inventoried, not treated as supported maker requirements:
HeadlessRendering, ShowFooter, ShowColumnHeaders, GridActions, BasePaletteColor,
Font, FontColor, FontSize, FontWeight, HeaderFont, HeaderFontColor, HeaderFontSize,
HeaderFontWeight, NoDataText, TabIndex, DateTimeFormat, DateOnlyFormat,
DisableStatusColumn, DisableStatusColumnSavingIcons.

Native Items and ChildItems/ParentId datasets are not equivalent to DrivenGrid's
single records dataset with a self-lookup loader. DrivenGrid's original 19 inputs
and 16 additions are documented in its manifest/README, including legacy no-op
EnableChildCommands/ChildContextLookup retained for installed form compatibility.

Microsoft's [configuration guidance](https://learn.microsoft.com/en-us/power-apps/maker/model-driven-apps/the-power-apps-grid-control)
describes grid/list settings and limitations. Actual installed declarations and
earlier live observations underpin this inventory; documentation alone is not a
claim that a particular combination is enabled on the test form.

## Current validation record

- Full local regression suite: passed after all save/logging refinements and
  compatible dependency-security updates. Includes the exhaustive 42,000-record
  fixture, 10,000-row native paging simulation and bounded Select All cases.
- Control and focused harness/core/presentation-test TypeScript checks passed.
  The final authoritative Release build, including ESLint, passed. Webpack retained
  three bundle-size/performance warnings; the production bundle is 725,880 bytes
  (about 709 KiB). This is not a matched cold-load performance result.
- Final unmanaged ZIP readback verified solution 1.0.90.0, control 1.0.111,
  CraigMansell/cm publisher, one dataset, 35 distinct inputs and all referenced
  resource keys. No import or publish was performed.
- Added compiler tests cover all 30 date mappings, fiscal inputs, literal patterns,
  duplicate-name group identity, totals including blank/zero/decimal values,
  reorder/width bounds, colour contrast and quoted TSV.
- Added component tests cover complete-vs-partial/error totals, automatic summary
  paging, grouping/collapse-all persistence, disabled presentation settings, list
  selection and honest unknown counts.
- Safety tests cover host/child read-only enforcement after asynchronous metadata,
  preserving text whitespace, numeric formats and no replay of ambiguous lookup writes.
- Browser: total 24,018,432 independently matched fixture data for 2,000 parents;
  zero child requests. Blank group 154 / 1,848,014 and Complete group 615 /
  7,392,806.5 matched independent sums. Compact fiscal period 3/year 2026 returned
  62 parents; selected child opened sample-53-child-2. Width/Move left passed.
- Read-only sandbox at 12:43 UTC: all 30 added date query operators accepted.
  `%` and `_` matched independently expected two records each; `[` matched zero.
  Fiscal boundary expectations were not inferred from the monthly local fixture.
- After restarting the patched harness, Select All again selected exactly 1,000
  parents with zero child requests. Clearing then expanding one branch returned
  20 children in one request; Open resolved selected child sample-1-child-2.
  The preview was left unfiltered, unselected and in its original column order.

Package: `CMModelDrivenGridSolution/bin/Release/CMModelDrivenGridSolution.zip`.
SHA-256: `ECA9C5F62242C578D890ADD0A66274C26C83B8EE4302B8B89C4D8E3FEE99F7CD`.
Bundled JavaScript SHA-256:
`23E00D24DC0E83CD6F78FF314CA450A35117FEE4CEEABA721F210932E5826347`.
These identify the locally built candidate, not an installed sandbox resource.

### Dependency-security review

Compatible lockfile updates reduced `npm audit` from 35 reported vulnerable
package entries to nine (five moderate, four high). The final
`npm audit --omit=dev` returned zero; this is an advisory check, not proof of
complete security. Updated tsx to 4.23.13 and brace-expansion to 1.1.18 within
existing dependency ranges. The test-data server was restarted with the patched
esbuild dependency, addressing the known
[Windows development-server path issue](https://github.com/evanw/esbuild/security/advisories/GHSA-g7r4-m6w7-qqqr).

The remaining entries belong to PCF development tooling, not shipped runtime
dependencies: pcf-scripts/pcf-start use applicationinsights 2.x and its OpenTelemetry
1.x chain; pcf-start's BrowserSync/UI use Immutable 3.x. The reported fixed lines
require incompatible major replacements in those chains. No forced overrides or
toolchain downgrades were applied merely to clear the audit. Keep development
servers loopback-only and do not expose them to untrusted users; a separately
validated upstream-toolchain upgrade remains recommended. The local sample server
binds to 127.0.0.1. Full tests, types and production packaging passed after updates.

## Remaining acceptance / recommended next step

The source batch is publishable as an **unreleased candidate**, not as an all-feature
parity certificate. Do not promote it to production on the basis of harness tests.

For a sandbox deployment: capture a fresh exact unmanaged rollback, lock installed
versions/form/view/publisher state, install only the versioned PCF batch, verify
published resource hashes, and run the normal-handler QA form against the matching
native configuration. Test candidate Select All at the boundary, both grid levels,
all condition expected IDs/time zones, editable/disabled and restricted-role paths.
Use disposable records for actions; Assign/Share/Flow/Email are not blanket-authorized
by permission to edit/delete isolated sample options.

Full native-equivalent children from another table require an architecture decision:
use the actual Microsoft grid where its independent child dataset and host command
contract are required, or separately design supported custom integration. Native
cell customization does not automatically grant replacement of the Excel-style
header menus. The owner's answer makes native behaviour the target, not permission
to bypass private interfaces or silently replace every existing DrivenGrid form.

Do not create 42,000 sandbox records without a capacity-approved fixture plan; the
existing environment warns of low storage. The current large-data fixture remains
local. Keep the requested no-fullscreen/no-extra-child-toolbar/automatic-scroll
design choices while closing the remaining gaps.
