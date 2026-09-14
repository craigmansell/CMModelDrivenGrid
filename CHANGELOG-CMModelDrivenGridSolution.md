# CMModelDrivenGridSolution changelog

## Sandbox release 1.0.96.0 / control 1.0.117 — 2026-09-14

- Published and verified the exact locally tested unmanaged package in Capricorn sandbox. Full local suite, focused test TypeScript and Release build pass (three existing bundle-size warnings). Published assets match; all Item Template forms, option views and publisher remain unchanged. No production/schema/ribbon/security changes.
- Final fresh-tab acceptance: initial paging; 33-root/64-child Select All; scope switching/collapse; exact-child native Edit; locked system-managed dates; reversible parent save. The failing selection-to-totals path now completes with sum 799.36, Active 30/sum 679.36 and Inactive 3/sum 120. Six final-version filter spot-checks match exact reference IDs, following the 174/174 full live matrix on 1.0.114.
- All 106 original QA fixtures restored with zero changed business fields. One new explicitly disposable child was deleted through the native command bar and its absence verified; historical deletion preserved. Excel keyboard ranges pass live, but clipboard readback remains a manual check because the test browser's virtual clipboard blocked that workflow. No all-native-feature or production sign-off.
- Reuse the exact PCF-only batch rollback under the unchanged freshness/scope locks. Final input ZIP SHA-256 `2425AC6F4AD302469CDB3F13C987678F89F62983BD55F3F4BFF5733745B30058`. Detailed versioned results, package evidence and remaining boundaries: `SANDBOX-VALIDATION-2026-09-14.md`. Nine upstream development-tool audit entries remain; shipped-dependency audit is clean.

## Sandbox dataset-snapshot paging fix candidate 1.0.96.0 / control 1.0.117 — 2026-09-14

- A fresh sandbox page on 1.0.116 exposed the immutable-host variant: Power Apps replaces the dataset/context object during updateView, leaving the object which issued the page request stale. Read the delivered dataset after awaiting each page, not the issuing snapshot. A regression failed with the same repeated-page error before this fix and passes afterwards; the existing mutable-host, 10,000-row, duplicate-page and selection-handover checks remain.
- Reuse the verified PCF-only rollback baseline and unchanged component boundaries. Final signed-in checks passed as recorded in the release entry above; 1.0.116 is superseded and not accepted for use.

## Sandbox retained-row paging fix candidate 1.0.95.0 / control 1.0.116 — 2026-09-14

- The final selection/totals sequence exposed a second paging case: switching native command selection restarts physical host paging but deliberately retains visible roots. A page containing only retained rows was mistaken for an empty page, preventing complete totals/grouping. A regression reproduced the failure (one catch-up request instead of two). Consume overlapping physical pages under one promise until a new root or EOF, retaining existing rows, with repeat-page and 500-page safeguards. Yield to pending selection and do not display a false empty-page error while that query is taking over.
- Version 1.0.115 is published. A genuinely fresh tab confirmed system-managed date fields locked for both levels and normal cost editing available; an already-open tab reused earlier scripts after reload and was excluded from that acceptance check. Root Select All selected all 33 records; child Select All selected all 64 across pages, switching scope deselected parents, and collapse cleared child selection. The native command bar successfully deleted only the newly-created, marked disposable child; deletion verified by exact ID and marker.

## Sandbox read-only fix candidate 1.0.94.0 / control 1.0.115 — 2026-09-14

- Live testing found that system-managed Created On / Record Created On fields could open editors because dataset columns did not supply the optional read-only flags. Read bounded public Dataverse attribute metadata once per table/column set, propagate read-only flags to parent and child grids, and enforce the same check immediately before saving. Unknown, failed, missing, calculated, rollup and derived attributes fail closed; a failed metadata load can be retried with the subgrid Refresh command. This is column eligibility, not a replacement for server record/field security or custom form business rules.
- Add metadata, batch-size, pending/failure, previous-table and guarded-save regression checks; retain the existing late read-only lookup tests with explicitly confirmed editable fixture metadata.
- Version 1.0.114 is published in sandbox: all 174 live filter cases (87 root/87 child) match independent Dataverse references; child paging 50+14, nested G01/G02, shared native Edit targeting and reversible child cost/lookup/state actions passed. The final 1.0.115 validation and record-restoration audit are pending.

## Sandbox paging fix candidate 1.0.93.0 / control 1.0.114 — 2026-09-14

- Live 1.0.113 testing found that a new page could render before the host paging promise settled. Prefetch then reused that promise as a new request, leaving a false "No additional records" warning even after all 33 QA roots arrived.
- Added a regression which fails on 1.0.113, then a separate promise-lifetime guard which prevents overlapping prefetch and re-arms scrolling after completion. Preserve genuine empty-page Retry recovery and all Excel/hierarchy features. The focused grid regression suite passes; full validation and deployment evidence will be recorded below before acceptance.
- Owner resumed sandbox publication. The fresh check already found 1.0.113 published, so no additional Publish All action was needed. Its three resources match the exact tested 1.0.92.0 package and forms/views/publisher remain unchanged. That version's live test is not final acceptance because of the paging defect.

## Sandbox 1.0.92.0 import — publication approval required — 2026-09-14

- Re-ran the full local suite successfully, including all 16 review regressions, the 42,000-record fixture and strict-host paging. Exact candidate ZIP remains `948FADDD33A48623DC8AF3764AA9415EFECE60ACBBF79AFAAC8C39A69DBDE69E`; no product-source or version changes.
- Verified sandbox organization, installed 1.0.110 / 1.0.89.0, publisher and existing hierarchy bindings. Created and verified fresh PCF and full-release unmanaged rollback exports at CrmPackage `outputs/deployment-backups/drivengrid-v113-sandbox-20260914-095030/`. Published/draft Item Template forms and option views matched and the audited unpublished UI inventory was clear before import.
- Imported the exact unmanaged 1.0.92.0 package successfully (100% at 08:54:18 UTC). Targeted resource publication installed the tested bundle hash, but the published control definition is still 1.0.110. The attempted audited Publish All action was rejected by the safety reviewer before execution because broader publication requires explicit owner approval. Do not describe this mixed publication state as a completed deployment or approved for use.
- Added locked prepare/apply/publish/verify/export and read-only QA helpers. Snapshot captured all 106 remaining disposable fixtures, preserving the historical C65 deletion. All 174 independent filter-reference queries were accepted after correcting a reserved-variable issue in the new QA generator; these are server reference results, not PCF UI passes. No fixture values, forms, views, schema, security or production data were changed.
- Full live UI/action tests and final export are paused until publication is explicitly authorized and completed. Current status and resumption instructions: `SANDBOX-VALIDATION-2026-09-14.md`. The old deployment baseline must not be reused if it is stale or component scope changes.

## Unreleased candidate 1.0.92.0 / control 1.0.113 — fix pre-sandbox review gates — 2026-09-10

- Fix synchronous ordinary dataset filter/sort rejection: snapshot and restore nested physical/logical filters and sort priorities; release local loading even when records are unchanged or the same error occurs twice. If restoring the host query also fails, block row selection/Select All until recovery; a successful refresh follows the accepted host filter and a new view does not inherit the old recovery state.
- Bound compact keyboard navigation and Excel range copying to displayed fields. Clear ranges and pending off-page focus on reflow/column changes; table-mode navigation still uses all visible columns. Keep Excel menus, Shift sort/ranges, quoted TSV, lazy child expansion, shared outer commands and infinite scrolling.
- Recheck current column editability, presence, alias, type, view and table immediately before lookup writes and permitted property-name retries, including clearing a lookup. Late read-only/changed-view results must not issue another update; existing no-retry rules for ambiguous or permission failures remain.
- Promote the review probes into the normal regression suite and expand to 16 groups. Final full suite, focused harness/test TypeScript and authoritative Release build with ESLint pass; includes exhaustive 42,000-record hierarchy and 10,000-row/200-batch strict-host paging checks. Browser compact/table keyboard, Excel menu/sort and shared exact-child sample Open checks pass.
- Verify local unmanaged package 1.0.92.0, control 1.0.113, publisher CraigMansell/cm, one control/one dataset/39 inputs, all 91 resource keys. Bundle 743,630 bytes; three existing Webpack performance warnings remain. Package SHA-256 `948FADDD33A48623DC8AF3764AA9415EFECE60ACBBF79AFAAC8C39A69DBDE69E`; bundle SHA-256 `F30DC138710B6242005B200AA521A97BB39BDC8F0E967E4123CCA13EC7A798C4`.
- No Dataverse import/publish, form/view/ribbon/schema/record writes or production action. Last recorded sandbox verification is still 1.0.110 / 1.0.89.0, not re-read in this fix batch. A fresh exact rollback, environment/version/component locks and signed-in acceptance remain mandatory before deployment. Evidence: `VALIDATION-FIXES-2026-09-10.md`; historical failing review retained.

## Pre-sandbox review of candidate 1.0.91.0 / control 1.0.112 — HOLD — 2026-09-10

- Re-ran the complete local suite, type checks, unmanaged Release build, package/resource readback and dependency audit. Existing regressions pass, including all 42,000 fixture records and 10,000-row strict-host paging; shipped-dependency audit is clean, nine development-tool entries remain.
- Added opt-in supplementary release probes. They reproduce three unfixed defects: ordinary filter rejection can leave loading/UI filter state stuck, compact End targets an undisplayed column, and an asynchronous lookup save misses a column becoming read-only during metadata resolution. Hold release sign-off until these pass; no product-source or version change in this review.
- Rechecked local child Open/Activate targeting, collapse selection clearing, 1,000-row Select All, Excel menu, Contains Atlas (334 matches) and descending sort preserving the filter. No Dataverse deployment, record writes, component changes or production action.
- Evidence, reproduction command, exact rebuilt archive identity and acceptance gaps: `REVIEW-PRE-SANDBOX-2026-09-10.md`. Earlier feature-validation/comparison documents now link to the hold decision.

## Unreleased candidate 1.0.91.0 / control 1.0.112 — keep Excel, close further gaps — 2026-09-10

- Keep the independent DrivenGrid control, Excel-style searchable value checklist, dotted guides, dynamic condition panel, AND/OR filters, range copying and single shared command bar. Do not introduce a Microsoft grid dependency, fullscreen, a child toolbar or an Add to current sort item.
- Add Shift-assisted multi-column sorting to the existing direction buttons. Child sort updates retain other priorities and clearing removes only the chosen column.
- Add day/week/month/quarter/year date groups in chronological order; new groups start collapsed. Weeks honour the provided first day, DateOnly does not shift through timezones. DateTime grouping currently uses browser calendar time; full Dataverse user-timezone/TimeZoneIndependent parity remains open. Complete-level grouping/totals still fetch all matching rows in that level, not descendants.
- Add keyboard arrows, Home/End, PageUp/PageDown, Ctrl-edge and Shift-range extension. Focus crosses virtual rows and can request the next page; Ctrl+End means the last currently loaded row, not an unbounded fetch. Copy excludes collapsed groups and preserves quoted Excel TSV. Read-only cells can still be navigated and copied.
- Add opt-in EnableJumpBar: A–Z on the primary sorted text column, preserving existing Excel conditions. All removes only the jump condition. Strip internal jump markers before calling the public dataset API.
- Add RememberColumnLayout (default true), scoped by browser origin, user, entity, view and root/child level. Store only column names/order/widths, not records, filters, drafts or selections. Handle blocked/corrupt storage; Reset column layout restores defaults. Add NavigationTypes all/primary/none, including primary text and compact-list links.
- Add EnableLookupDialog (default true): More records uses supported PCF lookupObjects; validate its single-table/single-record result against metadata and configured restrictions before the existing save. Guard cancellation, unmount, changed view and read-only changes. Keyboard lookup highlights are explicit; Tab no longer picks an unchosen first result. Prevent portalled lookup clicks from reaching row selection, honour actual custom-table primary-name metadata and choose target-specific polymorphic bindings.
- Local suite passes, including new lookup UI/safety, persistent-layout, jump-filter, keyboard, initial-collapse and observer-free automatic-scroll cases, plus existing filter/child/42,000-fixture tests. No Dataverse deployment, form/view/ribbon changes or business-data writes in this batch. Release package validation is recorded in the comparison document.

## Unreleased candidate 1.0.90.0 / control 1.0.111 — feature comparison batch — 2026-09-10

- Add 30 date conditions with shared SDK/FetchXML mappings, fiscal period/year validation and server-side fiscal interpretation. Date/time now exposes 53 conditions (all 50 observed native choices plus After, Before and Between); date-only excludes four hour/minute conditions. Keep common conditions before advanced fiscal choices. Read-only sandbox queries accepted all 30 additions; this is not complete expected-ID/time-zone acceptance of each date condition.
- Correct literal `%`, `_` and `[` text patterns. An isolated 101-row sandbox query returned exactly the independently expected IDs for percent/underscore (two each); bracket returned the expected zero, without claiming a positive bracket fixture. Root predicates retain the public-view/5,000-candidate/500-page safeguards.
- Add complete-level grouping, sum/average/min/max, retained individual/global collapse, column reorder and explicit 80–2,000px content widths. Summaries automatically fetch remaining pages without opening children, support pause/resume, and stay hidden on incomplete/error results. These are client-side raw-number summaries, not native server aggregates or mixed-currency conversion.
- Add compact grid/list/auto reflow, list initials/labels/banding, metadata choice colours, type/state icons and 16 configurable inputs. Preserve the single native record command surface, same-table hierarchy, leaf checkboxes, no child action/count strip, no fullscreen and no Load more button. Different-table/independent child datasets, advanced lookup, label-pattern lookup filters, native customizer/events and full keyboard/localization parity remain open.
- Enforce host/column/child read-only state before writes and again after asynchronous lookup metadata. Only retry a lookup save after a confirmed rejected property name; never repeat ambiguous/network/permission/business-rule failures. Remove record/query diagnostic logging. Preserve text whitespace, align FP/Money/Float/Double parsing and show an honest loaded-count lower bound when total count is unknown.
- Quote tabs/newlines/quotes in copied Excel TSV and show clipboard-denied/unavailable errors. Fix lookup-link event handling and stale range/presentation state when settings change. Remove unused feature placeholders.
- Local browser: independently verified all-parent total 24,018,432 for 2,000 parents, zero child requests; grouping/collapse with 53 rendered row nodes at a checked viewport; compact fiscal filter returned 62 matches; child Open targeted sample-53-child-2, not its parent; column width and Move left passed. After dependency updates, Select All again selected exactly 1,000 parents without child requests; expanding one branch loaded 20 children and Open targeted sample-1-child-2. These are local fixture checks, not a 42,000-record Dataverse benchmark.
- Apply compatible development-dependency security updates, including tsx 4.23.13 and brace-expansion 1.1.18, and restart the local server on patched esbuild. Audit reduced from 35 to nine development-tool entries (five moderate/four high); shipped-dependency audit reports zero. Remaining upstream PCF telemetry/BrowserSync dependency chains need a separately validated incompatible upgrade, not forced overrides.
- Final full regression suite, control/focused harness TypeScript checks and authoritative unmanaged Release build including ESLint passed after all refinements and dependency updates. Package readback verified versions/publisher, 35 distinct inputs, one dataset and all resource keys. Bundle is 725,880 bytes; three Webpack size/performance warnings remain. Candidate ZIP SHA-256 `ECA9C5F62242C578D890ADD0A66274C26C83B8EE4302B8B89C4D8E3FEE99F7CD`; bundle SHA-256 `23E00D24DC0E83CD6F78FF314CA450A35117FEE4CEEABA721F210932E5826347`.
- This is a local candidate only. No Dataverse import/publish, form/view/ribbon/schema/data/production changes in this batch; installed sandbox remains 1.0.110 / 1.0.89.0. A fresh exact rollback, sandbox/version locks and signed-in candidate smoke test are mandatory before deployment. Historical rollback exports are not a new baseline.
- The owner explicitly approved committing/pushing the complete dependent DrivenGrid-only batch on 2026-09-10, resolving the earlier Git-scope blocker. Unrelated CrmPackage edits are excluded. Source baseline and read-only QA evidence: CrmPackage `outputs/qa/drivengrid-feature-parity-20260910/`. Full inventory and remaining work: `FULL-FEATURE-COMPARISON-2026-09-10.md`.

## Unreleased - full-view Select All with Dynamics limit - 2026-09-10

- Fix the header/footer mismatch reported with 75 selected rows. Track explicit record IDs across page replacement and virtualization instead of Fluent's open-ended all-current-and-future-items selection. Publish only the completed selection, keeping the shared footer and command targets in agreement.
- Select the first 1,000 matching records in current filter/sort order, or all matches when fewer exist, following Microsoft's modern-grid behaviour. Keep parent and child scopes separate. Native root and standard dataset paging stop at the cap instead of rejecting an entire large view or selecting it without a limit. Child Select All now loads later child pages, stopping at end-of-data or 1,000.
- Show selection progress/errors and an explanatory limit message. A second header click (or a click with an existing selection) clears all IDs; a pending operation can be cancelled. Query changes, scope switches, refresh and disposal cannot reapply stale selections. Newly scrolled rows beyond the cap stay unticked, while the capacity-reached header stays ticked.
- Full local suite passed, including exact 75-to-1,000 selection, later-row boundaries, 17 filtered/reversed records, cancellation/query races/failures, child branches of 20/85/1,200, standard dataset caps/refresh cleanup and strict native-host caps for 0/17/1,000/2,000 roots with unknown totals. Control/harness TypeScript checks and the normal local PCF build including ESLint passed. Existing Browserslist/development-bundle warnings remain; no dependency changes.
- Browser verified: 50 loaded parents to exactly 1,000 selected; row 1,000 ticked and row 1,001 unticked after scrolling; 334 filtered parents all selected; all 20 children selected without their parent; a 2,000-child branch selected exactly 1,000 in 40 pages. Native command execution and signed-in sandbox acceptance are not claimed for this update.
- No manifest/solution version bump, Dataverse deployment, form/view/ribbon change or data mutation. The sandbox retains the earlier installed release until an explicitly validated deployment. Evidence: CrmPackage `outputs/qa/drivengrid-select-all-20260910/REVIEW.md`. Git publication remains blocked by overlapping pre-existing modified/untracked feature dependencies; nothing unrelated has been staged or pushed.

## Harness-only hierarchy stress QA - 2026-09-10 (control unchanged)

- Add a directly linkable local scenario with exactly 2,000 parents and 20 children per parent: 42,000 unique records, two levels, no grandchildren. Preserve the existing standard and 10,000-parent scenarios. Index the local fixture by ID and parent so expanding a branch does not scan all 42,000 records.
- Keep automatic 25-row parent batches and lazy child requests; display fixture totals and child request counters. No actual PCF runtime, manifest, solution version, form, view, record or environment changes. Sandbox remains at the previously verified control 1.0.110 / unmanaged solution 1.0.89.0; no deployment in this batch.
- Added exhaustive fixture regression: all 80 parent batches and all 2,000 branches account for every record exactly once, with correct child scope, leaf status and end-of-data. Last-parent filtering and child numeric filtering/descending sorting pass. Full local regression suite, control ESLint and focused harness TypeScript checks passed.
- Actual browser scrolling reached parent 2,000 without using preload. Checked row DOM counts peaked at 120; unopened branches did not request child data. Parents 1, 1,000 and 2,000 each returned 20 children. Child selection survived parent scrolling, and the sample Open command resolved the exact first/last child IDs. Last-branch descending sort and Budget >= 250 returned 11 children without altering the first branch; collapse cleared selection and reopening reused its filtered/sorted cache. A fresh root filter found parent 2,000 with only 50 parents initially loaded.
- This is local fixture/rendering evidence, not a 42,000-record Dataverse load test or proof of native command-bar performance. Existing server filter limitations are unchanged. Detailed evidence: CrmPackage `outputs/qa/drivengrid-2000x20-20260910/REVIEW.md`. Scoped Git publication remains blocked by overlapping pre-existing, uncommitted DrivenGrid dependencies; they have not been staged or published indiscriminately.

## 1.0.89.0 / control 1.0.110 - server child paging cookies 2026-09-09

- Signed-in 1.0.109 still repeated the first 50 children when page 2 was requested. Its explicit duplicate-page error correctly stopped the loop; the raw client-query adjustment did not fix the host wrapper and is not accepted as final.
- Use the authenticated app's same-origin public Dataverse endpoint for child page reads, preserving server paging cookies, parent/view/filter/order scope and formatted values. Request continuation annotations, double-decode the supplied cookie according to Microsoft's FetchXML guidance, XML-escape it, and stop on the server's end-of-data signal. Keep permission/HTTP errors visible, reject redirects and time out stalled requests; no private host APIs or data writes.
- Full integrated regression and authoritative Release build passed, including 50+14 HTTP/cookie pages, escaped values, scope/page rejection, server end-of-data and HTTP failure. Published sandbox 1.0.110 / unmanaged 1.0.89.0 using the verified 154902 rollback and exact installed-109 guards. All three published resources match the tested ZIP; published/draft forms/views and publisher remain unchanged. Input ZIP SHA-256 `108BEDB1C5869F3D9C23D94BB9FD1F70A0C709219EDF37B1A18DFD9309DE2A70`; bundle `1D2FF185B58EDB4F67DAB81B5F5F04C19FF3B6B6C89B756D96F46D1C42D1EDF4`.
- Fresh signed-in session passed: all 33 parents and all 64 children visibly reachable by scrolling, no stuck Loading records overlay, C64 selectable with native Edit/Activate/Deactivate and overflow commands, and Edit opened exact child ID `3f8257c1-4b8f-4320-b735-0b26050adb77` without saving. Root/child Name Equals Alpha red returned R01/R10 and C01/C10 respectively; clearing child filtering restored 50+14 paging. C64 native commands also passed with the root filter active; collapse cleared hidden selection. Parent selection also showed native commands without a blocking overlay.
- The previously open form retained an older control resource after reload. Acceptance used a new form session with the updated resource cache key; open the test link in a new tab if an old tab still shows the previous error. The QA form is left unfiltered and unselected. No record action that changes data was executed; fixture readback found no changes to the 106 remaining QA records or the outside-scope sentinel (C65 was deleted in the earlier approved QA).
- Final unmanaged exports verified in the same 154902 backup directory: PCF SHA-256 `62D6336A2E705E4D2B3A7AD8F1D3A0ABD9CD83938646236D89C6883AAB9E1059`; Capricorn release `00C8EABA4C6C51447FCBF6B98435C352529086B9926F935012DB5D22DFADB46F`. Local rendering stress passed 10,000 parents/2,000 children with 80–100 row DOM elements at checked points and retained virtualized branch selection. Existing literal-percent/underscore filter defects and the 5,000-candidate advanced-root-filter limit are not resolved. Scoped Git publication remains blocked by overlapping pre-existing dependencies and the unrelated release-repository branch.

## 1.0.88.0 / control 1.0.109 - client child paging correction 2026-09-09

- Intermediate 1.0.108 fixed the stuck native parent paging: signed-in QA reached all 33 root records, without a blocking Loading records overlay during append. All three published resources match the tested ZIP; published/draft forms/views and publisher are unchanged.
- Signed-in child acceptance still stopped at 50. Read-only calls against the exact saved child view return C01–C50 and then C51–C64 (50 + 14), so the remaining records are present and authorized. Correct the client API's raw FetchXML contract, retain returned next-link options without re-encoding, scope continuations to the exact parent/filter/order, and reject duplicate pages explicitly instead of silently advancing.
- Add focused raw-query/continuation regressions including literal XML characters and mismatched parent rejection. Child/native/grid regressions passed. Authoritative Release build, second PCF-only push and signed-in acceptance pending. Reuse the exact original 1.0.107 rollback for this same batch; additionally lock the installed 1.0.108 bundle hash before import. No form/view/schema/ribbon/record changes.
- Release build passed (input ZIP `B6414855976C76480500145EE575F7D9E71946809D68BF8A7DDF0389180AD3A9`). The intended baseline reuse was invalidated by independent published changes to two Curtain forms; the guard stopped before import. Fresh full unmanaged rollback at CrmPackage `outputs/deployment-backups/drivengrid-loading-scale-sandbox-20260909-154902/` preserves those changes and installed 1.0.108. The actual second push uses that new baseline; no unrelated form change is reverted.

## 1.0.87.0 / control 1.0.108 - loading and large-list correction 2026-09-09

- Fix the reproduced native-hierarchy cache mutation: publish fresh record maps/ordered IDs for each batch and update, so React renders appended and refreshed records.
- Use the documented next-page API with a shared in-flight promise, timeout/error recovery and page-counter rollback. Do not confuse a scroll batch with an external refresh; serialize native command-selection queries after pending paging.
- Keep loaded rows readable during paging. Check scroll distance as well as the intersection marker, because a parent viewport can clip a nested grid's marker indefinitely.
- Keep Fluent row virtualization enabled with expanded children. Account for row/panel heights; retain child data, filters, ordering, widths, expansion, drafts and scroll state outside virtualized row DOM. Clear hidden native selection on explicit collapse/removal, not merely when scrolling a row off screen.
- Add a 10,000-row/200-batch strict paging regression, a clipped-child-scroll regression, and optional 10,000-parent/2,000-child local sample data. Fresh sandbox rollback: CrmPackage `outputs/deployment-backups/drivengrid-loading-scale-sandbox-20260909-150733/`. Deployment and signed-in acceptance are pending; no form/view/schema/ribbon/record-data/production changes are included. Existing literal percent/underscore filtering issues and server-filter candidate limits are not part of this batch.


## 1.0.86.0 / control 1.0.107 - live presence-filter correction 2026-09-09

- The signed-in 1.0.106 smoke test found that a supported EQ/NE presence disjunction still included a blank lookup row, although numeric Equals 999 correctly returned zero. Do not accept the local SQL-style simulation as proof of the native dataset's null semantics.
- Resolve root Contains data with actual Dataverse NOT NULL through the same bounded, relationship-scoped server path as negative text filters. Never send unsupported operators to the primary dataset. Add a complete zero-match presence transaction regression; child NOT NULL remains direct FetchXML.
- Filter/child/native regressions and the authoritative unmanaged Release build, including ESLint, passed. Input ZIP SHA-256 `EDD4E8F27983ABAEB3E43960F5C333F33D5BD3220350A9E28B05892D8AC8F8F0`; bundle SHA-256 `0B5FAB895122E52A76532CBAC663B8B48F050179A49FBDB5ACCA07AFBB5AB025`.
- An independently published change to Quotation Curtain Item Template main form invalidated the first snapshot before the second push. The target DrivenGrid form stayed byte-for-byte unchanged. The original-baseline deployment stopped before import; a fresh full unmanaged PCF/Capricorn release rollback was verified under CrmPackage `outputs/deployment-backups/drivengrid-filters-leafrows-sandbox-20260909-102025/`, preserving the separate change. The v107 helper also locked the installed 1.0.106 / 1.0.85.0 and its exact bundle hash before any write.
- Published sandbox control 1.0.107 / unmanaged solution 1.0.86.0. All three published resource hashes match the tested package; all five published/draft Item Template forms, Item Template Option views and publisher match the fresh baseline. No form/view/schema/ribbon/data/production mutation was performed.
- Signed-in acceptance passed on the existing demo form: root Component Contains data excludes the blank parent, Option Contains data retains the populated parent, and clearing either restores unfiltered rows. Child Component Contains data removes only blank children; clearing restores both. Only the parent has an arrow; both leaf children remain selectable. Native Edit/Activate/Deactivate appear for Child A after clearing a zero-match root filter and for Child B while a matching root filter is active. Collapse and clearing root filters clear selection. The initial pre-existing tab had a stale native command bar for parent and child alike; reloading and repeating both filter sequences passed. No command action that changes data was executed.
- Text-column choices and negative predicates passed local/harness and read-only Dataverse query checks. This form has lookup/numeric columns, so no visible text-column live acceptance is claimed. Advanced root resolution currently requires a public saved view, with the documented 5,000-candidate/500-page limit. Scoped Git publication remains blocked by overlapping earlier modified/untracked DrivenGrid dependencies.
- Final unmanaged exports verified with exact unchanged form/view checks: PCF SHA-256 `F481F10740F52C2DC167F62617C9A270EB8180431FCF2DD5E6E0303FE8BE7489`; Capricorn release SHA-256 `91B428EEE6F65CFEF7003AC85BA1D674F6389F6305278A11380CA10D4686F394`. `runtime-review.md` and export metadata are beside the fresh rollback manifest.

## 1.0.85.0 / control 1.0.106 - text filters and leaf rows 2026-09-09

- Hide expand chevrons on rows with no accessible children in the configured child view. Batch distinct-parent checks (50 per request), cache per refresh, check every hierarchy level, preserve checkbox alignment and clear hidden branch selection when the last child disappears. Failures remain visible and retry on native Refresh; do not assume a failed query means no children.
- Match the requested ten text operators and ordering, adding Contains data, Does not contain, Does not begin with and Does not end with. Keep type-appropriate numeric/date/lookup choices and hide value inputs for presence/blank conditions. Remove the obsolete unsupported-operator footnote.
- Child predicates execute in Dataverse FetchXML. Root Contains data compiles to supported EQ/NE predicates; negative text predicates scan the actual bound dataset scope, check candidates through Dataverse, then use supported ID filters on the primary dataset. Never inject host records or send unsupported NOT LIKE / NOT NULL operators to the model-driven dataset. Preserve native record commands and logical filter labels, re-evaluate on refresh/sort, and restore the previous filter on query failure.
- Bound root text-resolution to 5,000 candidate rows and 500 pages, with an explicit narrowing/error message instead of partial results. Child filters remain server-paged. No form/view/schema/ribbon/data/production changes are planned in this PCF-only batch.
- Published as an intermediate build with matching resource hashes and exact unchanged forms/views. After refreshing stale browser code, the live leaf-arrow and native child selection checks passed. Root Contains data did not exclude a blank lookup; v106 is not accepted as the final filter implementation. Continue with v107. Scoped Git publication remains blocked by overlapping pre-existing DrivenGrid edits; preserve them.

## 1.0.84.0 / control 1.0.105 - compact child expansion 2026-09-09

- Remove the visible Child options / loaded-count heading. Retain an accessible child-region name and read-only status, loading/error feedback and existing record-selection footer.
- Use Fluent right/down chevrons beside always-visible square checkboxes in a fixed leading gutter. Keep the first data column free of expansion controls; preserve expansion when selection is disabled. Stop chevron events from selecting or invoking a record.
- Preserve the native child-selection controller, queries, filters, command bar and existing form/view configuration. Update column sizing for the compact gutter and add expansion/selection/layout regressions.
- Full local regression suite and TypeScript passed. Correct explicit row-renderer typing found by the release lint check; focused grid tests and the authoritative unmanaged Release build then passed, including ESLint. Harness screenshots, Enter/Space expansion and 56px header/row alignment passed; existing bundle-size/Browserslist warnings remain.
- Published sandbox control 1.0.105 / unmanaged solution 1.0.84.0. Signed-in screenshot and native command checks passed in a fresh form session: both children enabled Edit/Activate/Deactivate, the native overflow commands remained available, Edit opened Child A's exact record without saving, and collapsing cleared hidden selection. No child heading or embedded record-action buttons remain.
- Fresh PCF-only rollback, exact unchanged forms/views/publisher checks, deployment and runtime evidence are at CrmPackage `outputs/deployment-backups/drivengrid-child-layout-sandbox-20260909-085328/`. All three published resource hashes match the tested package. Input ZIP SHA-256 `36CA5583CAE4F99F1F65E85669EB3AC21C578609678F3C92F0F3A9A11AE8B7C1`; final unmanaged export SHA-256 `76142C3BA4183C0C05BC0FE7745BD29CDD3CAA3339E78D0102615591D7F32C61`.
- No form, view, schema, relationship, ribbon, record-data or production changes. Existing open forms may need reopening after publication. Scoped Git commit/push remains blocked by overlapping prior modified/untracked DrivenGrid work; source and the task-only baseline are preserved locally.

## 1.0.83.0 / control 1.0.104 - native child dataset selection 2026-09-08

- Add opt-in native child selection backed by a dedicated view that includes parent and child records. Root display filters and sorting remain independent; the internal query admits only explicitly selected children and loads them through the documented primary dataset APIs before notifying native commands.
- Preserve nested panels, parent caches, child filters and the existing subgrid command bar. Do not inject record objects into platform state, bypass native rules or add a replacement toolbar. Scope one hierarchy level at a time, reject mixed/oversized targets, ignore late requests and verify native selection readback.
- Correct the root footer to include child selection and show confirmed native selection counts. Refresh/rejected selections clear checkboxes without accidentally selecting a parent. Add strict-host regression coverage for the original rejection and the revised loading sequence.
- Local checks passed: full regression suite, TypeScript, added strict-host/cache/refresh tests and authoritative Release unmanaged build including ESLint. Bundle is 636 KiB; the existing bundle-size/Browserslist warnings remain. Input package SHA-256 `0B41EE0AED798093B3F6A588735418940055A51C4488AF0498A6D507ABEF115D`; bundle SHA-256 `862837139BD2525B3D8A0130AB5F85DCDD504508FA329603E1CFE6B3CC806AF6`.
- Published to sandbox with the dedicated hierarchy view and native-selection setting on the standard Item Template form. Fresh full rollback is at CrmPackage `outputs/deployment-backups/drivengrid-native-selection-sandbox-20260908-134037/rollback-manifest.json`. Final exact readback confirms only the target subgrid changed, all three form factors retain the same nine columns, and the original Options view, other forms and publisher are untouched.
- Signed-in acceptance passed after reloading the cached form: child selections show the existing native Edit/Activate/Deactivate and overflow commands. Command Checker confirms all Edit enable/display rules True; native Edit opened Child A and Child B by their exact IDs, without saving. Multi-child selection, parent/child switching, independent root/child filters, descending child sort, collapse cleanup and native subgrid Refresh passed. No embedded action buttons, native-grid replacement, rule overrides or business-record changes.
- Final unmanaged exports: PCF SHA-256 `00E246388121C0774FCD9423C1F2A226D096C4830252DF4AE55C644530A051A1`; Capricorn release SHA-256 `3020B2E4EB4B88BCC7423C44F602D83F08811AE519AA8C9285A4DDCB7C2DB40F`. Runtime evidence is `runtime-review.md` beside the rollback manifest. Production remains untouched. Scoped commit/push is blocked by overlapping earlier modified/untracked DrivenGrid files; do not include them indiscriminately or publish an incomplete dependent change.

## 1.0.82.0 / control 1.0.103 - single command surface 2026-09-08

- Owner requires all record actions on the existing subgrid command bar and explicitly rejects an embedded/separate child toolbar. Remove Add child, Refresh children and local Retry buttons; child headers retain only their count/read-only status. Preserve row selection, filters, sorting, editing, expansion and scrolling.
- Remove unused embedded-action wiring and temporary native-command diagnostics from the PCF runtime. Keep legacy manifest inputs for existing form compatibility and correct their misleading descriptions. Do not claim that native child command selection is fixed: the 1.0.102 runtime evidence still applies.
- Add regression coverage that no embedded action buttons appear even when legacy callbacks are supplied; host refresh must reload child data and clear selected descendants. The harness uses one clearly labelled simulated command bar, now including Refresh, and explicitly discloses the remaining native integration gap.
- Full local tests, TypeScript, ESLint and the authoritative Release unmanaged build passed. Harness visual/interaction checks confirmed a single simulated top command bar and no child-header buttons. Signed-in sandbox checks confirmed no Add child, Refresh children or local Retry buttons; the existing subgrid command bar's Refresh reloads expanded children, clears child selection and preserves parent expansion.
- Published control 1.0.103 / unmanaged solution 1.0.82.0 to sandbox only. Input package SHA-256 `B14BFF767538A2A8FF23BED75D3E3881BDDD877B47CD1A7EC7B1A063ACEB5A62`; published and expected bundle SHA-256 `40259F30560B58DCDA986228BBCAC7F2C4E23015D3D0A27544CE4388A19BA45C`. Reused the exact verified PCF-only rollback baseline; all Item Template forms and publisher settings remained unchanged. No record, schema, ribbon or production changes.
- This completes the embedded-button cleanup, not native selected-child command integration. No separate child toolbar or native-grid replacement is authorized. Final integrated export, scoped commit and push remain pending; overlapping pre-existing repository changes must not be included indiscriminately. Runtime evidence: CrmPackage `outputs/deployment-backups/drivengrid-selection-sandbox-20260908-111642/v103-runtime-review.md`.

## 1.0.81.0 / control 1.0.102 - native command-rule investigation 2026-09-08

- Verified the actual native nested Power Apps grid on the existing Quotation Item Template form using Microsoft's session-only onload troubleshooting flag. Once asynchronous command rules completed, selecting Child A enabled Edit; Edit opened that exact child. No data or form changes were made.
- The documented command-retrieval API is present on DrivenGrid's dataset. Extend the temporary read-only diagnostic to request fresh rule evaluation and report the complete bounded command list and selected IDs. Native child command integration remains unaccepted.
- Sandbox outcome: selecting `7ec63dcb-7ef3-4508-9bf7-a42e25dd9293` returned `selected: []`; a refreshed `retrieveRecordCommand` returned only unselected-grid actions, without Edit/Activate/Deactivate/Delete/Assign/Share. Root and child GUID formats match, ruling out the suspected case-format difference. Do not bypass native rules or inject records into private platform state.
- Exact published bundle SHA-256 `F811CEF28B516398FB11CAD5D4875D4CAF579EA7610898BD93C6971777073710` matches the validated unmanaged package (package SHA-256 `AAD5144202B7122A02997BE9C30C769B084E51F40073C0DBCFC60ACB5577749A`). All Item Template forms and publisher settings remained unchanged. Evidence is under CrmPackage `outputs/deployment-backups/drivengrid-selection-sandbox-20260908-111642`.
- Paused native-command rollout for the owner's choice: actual Microsoft grid for exact native behavior, or retain DrivenGrid/Excel-style filters with a separately scoped child toolbar. Temporary diagnostics remain in the investigation build; remove them in the next implementation batch. No accepted final package/export/commit/push yet. Overlapping pre-existing repository changes must not be staged indiscriminately.

## 1.0.80.0 / control 1.0.101 - sandbox integration investigation 2026-09-08

- The live 1.0.100 smoke test exposed a gap not represented by the mock dataset: selecting root rows enables native commands, but separately fetched child rows do not. Switching to a child clears the parent command target safely. Do not treat native child commands as accepted or complete.
- Add temporary read-only diagnostics through the documented dataset `getSelectedRecordIds` and `retrieveRecordCommand` APIs to determine whether the supported command wrapper can supply the native child actions. No commands are executed by these diagnostics.

## 1.0.79.0 / control 1.0.100 - child selection 2026-09-08

- Replace disabled child selection and per-row trash icons with normal checkboxes and multi-selection. Publish selected child IDs through the supported dataset selection API to the existing host command bar. One hierarchy level owns selection at a time; switching levels clears other checkboxes, and collapsing/removing a parent clears hidden selected descendants.
- Prevent an older cross-page parent Select All operation from overwriting a newer child selection. Preserve child filtering, sorting, paging and retained scroll state.
- Compare against the signed-in native Power Apps grid using the same demo options. The native command checker confirms `Mscrm.EditSelectedRecord` and table permission/selection rules. Do not copy private Microsoft implementation code or bypass command enable rules.
- Add a clearly labelled sample command bar (Open, Activate, Deactivate and confirmed Delete) to the local harness. Add hierarchy-selection regression coverage. Sandbox command integration and final deployment verification are pending.

## 1.0.78.0 / control 1.0.99 - sandbox attachment 2026-09-08

- Deployed the exact locally validated unmanaged package to Capricorn sandbox and published control 1.0.99. Targeted web-resource publication alone did not publish the control manifest, so the audited Publish All fallback was used after checking pending form, view, app, chart and sitemap changes. Only this PCF's three resources were pending; publisher settings did not change.
- Attached DrivenGrid to `Subgrid_ItemTemplateOptions` on `Item Template main form` (`52d64db5-87fb-f011-8406-002248a34d35`) for all three form factors. Preserve the parent Options view, relationship and form scripts. Enable child expansion, inline editing and child commands, with the existing Sub Options view, `cm_parenttemplateoption` parent and `cm_itemtemplate` context lookup.
- Verified installed control/solution versions and exact proposed form readback. Other Item Template forms matched their baseline immediately after deployment. Fresh rollback evidence is in the CrmPackage repository at `outputs/deployment-backups/drivengrid-itemtemplate-sandbox-20260908-100837`.
- A later independent Product Template field addition on the Curtain form stopped the combined final-export guard; the DrivenGrid target remains exact. Sample-data setup and signed-in interaction tests are pending because the original `test` template now routes to the Curtain form. No production changes.

## 1.0.78.0 / control 1.0.99 - 2026-09-07

- Retain child panel state across collapse/reopen: loaded pages, filters, sorting, nested expansion and scroll position. State is session-local and is discarded when its parent leaves the rendered dataset.
- Add a separate child-only read-only setting and opt-in Add child / Delete child commands. New child forms inherit the configured parent and template/context lookups. Require confirmation for deletion, check parent scope again afterwards and refuse branch deletion, including inactive descendants.
- Disable mutation commands and cancel active cell editors when child editing is disabled. Keep child commands isolated from the hosting dataset's selection.
- Add on-screen harness creation/deletion dialogs and a child-read-only toggle. Verified sample add, delete cancellation, confirmed deletion, branch protection, parent editing in child-read-only mode, and retained numeric filter/descending sort/scroll position.
- Validation passed: generated manifest types, TypeScript, ESLint, full `npm run test:local`, browser interaction checks, `git diff --check`, and the Release unmanaged solution build (629 KiB bundle warning retained). No deployment of 1.0.99, no test records written to Dataverse, no Publish All. Prior 1.0.98 sandbox import remains pending publication/form binding. Add existing child, independent child columns and parent-total integration verification remain outstanding.

## 1.0.77.0 / control 1.0.98 - 2026-09-07

- Add opt-in expandable child rows with a configured self-referencing parent lookup and child view, preserving the child view filters and scoping every query to its parent.
- Load children only on expansion; support independent filtering/sorting, inline editing, recursive expansion, automatic paging, empty states and retry. Ignore stale requests after collapse or query changes.
- Keep child row selection separate from the hosting dataset command bar; scope clipboard shortcuts to the active grid.
- Add sample child and grandchild options to the local harness, including 85 children under the first project for scrolling checks.
- Full local regression suite and production unmanaged package build passed. Browser checks passed for nested expansion, child filtering/sorting, automatic paging through 85 children and a grandchild numeric edit. Lookup-filter tokens now resolve against loaded child records.
- Sandbox import completed successfully (job `d93a4170-3f4e-4487-b8f8-53b062c5809f`), advancing the unmanaged solution to 1.0.77.0. The published custom control still reports 1.0.70; the version guard stopped before any form write. Publication, form binding, signed-in runtime verification and final exports remain pending. No production changes.

## 1.0.76.0 / control 1.0.97 - 2026-09-07

- Name the control and solution DrivenGrid, preserving the technical control identity and solution unique name for upgrade compatibility.
- Update local harness branding and correct the README's obsolete Add to current sort instructions.
- Local naming change only; no Dataverse deployment. Item Template form preparation identified that the existing native grid expands child options, which this flat dataset grid does not yet reproduce.

## 1.0.75.0 / control 1.0.96 - 2026-09-07

- Add decorative dotted connector lines beside Select All and filter values, matching the supplied Excel reference while preserving aligned checkboxes, wrapping labels and single-region scrolling.
- Expand columns proportionally into spare grid width, measuring the actual scroll viewport and accounting for selection/padding. Preserve readable minimum widths and manual sizing; when all columns are manually sized, the last fills remaining space.
- Recalculate on host resizing and scrollbar changes; add width-allocation regression checks.
- Prefetch pages near the bottom of the grid and hide Load more when automatic observation is supported. Retain the unsupported-host fallback and show Retry for failed or no-progress page requests.
- Keep the guides out of pointer interaction and accessibility content. Local harness only; no Dataverse deployment.

## 1.0.74.0 / control 1.0.95 - 2026-09-07

- Size the condition-filter editor to its content, growing or shrinking as conditions and validation messages change, while preserving the fixed-size main checklist.
- Cap tall editors to the available callout height with a single scrolling field area and visible Apply/Cancel buttons. Local harness only; no Dataverse deployment.

## 1.0.73.0 / control 1.0.94 - 2026-09-07

- Give both column-menu views a consistent 320 by 560px footprint, constrained to the available host viewport, with one scrolling region and always-visible action buttons.
- Align the full-width search and bordered Excel-style checklist; place Select All and values on the same checkbox/text columns with compact rows and wrapping long labels.
- Remove Add to current sort from the menu. Selecting a new sort now always replaces previous sort levels while retaining filters.
- Add regression checks for the checklist structure, fixed editor actions and replacement-sort callback. Local harness only; no Dataverse deployment.

## 1.0.72.0 / control 1.0.93 - 2026-09-07

- Align header and row selection cells; use readable, resizable column widths instead of forced compression, right-align numbers, and keep the demo footer inside the viewport.
- Replace the hover-only filter flyout with a single keyboard-accessible panel and validated two-condition AND/OR editor.
- Add supported numeric comparisons/ranges and calendar/relative-date conditions, with type-appropriate inputs and sort labels.
- Make searched value selections apply only matching choices, preserve Cancel behavior, and show partial Select All state and loaded-values scope.
- Add ordered multi-column sorting and clear only the requested sort level while retaining filters; reset scroll position for new queries, not appended pages.
- Extend the interactive sample evaluator, relative sample dates, regression tests, and filter-support documentation without changing dependencies or deploying to Dataverse.

## 1.0.71.0 / control 1.0.92 - 2026-09-07

- Review and harden filtering, paging, inline editing, cell-range selection, and infinite scrolling; add regression coverage for the corrected cases.
- Use raw field values for picker filters, preserving readable labels, and represent blank values with a proper null condition.
- Preserve unrelated nested filter groups when applying or clearing a column filter.
- Preserve filters when sorting; register page waiters before requesting data, cancel them on destroy, and keep failed page requests retryable without skipping page numbers.
- Encode lookup-search URL values and avoid resolving duplicate display names to an arbitrary record; prefer the source record's actual lookup ID.
- Make the development watcher build locally instead of automatically deploying to Dataverse; keep deployment an explicit action and stop only the watcher's own build processes on exit.
- Add cross-platform development-tool regression checks to the local test suite.
- Make lint failures return a failing command exit code instead of relying on the PCF wrapper's console output.
- Add a local-only interactive harness with 180 sample records, working combined filters, inline editing, and incremental loading across six field types.
- Keep the Dynamics-style layout, removed full-screen action, and default infinite scrolling; correct header alignment and stale cell-range selection after sorting.

## 1.0.70.0 / control 1.0.91 - 2026-09-04

- Refresh the grid surface, headers, rows, menus, loading treatment, filter chips, and footer to align with the modern Dynamics 365 visual language.
- Open the Dynamics-style column actions menu from a normal column-header click while preserving right-click access.
- Remove the full-screen action and its unused control plumbing.
- Enable infinite scrolling by default while retaining the existing manual fallback when automatic viewport observation is unavailable.
- Keep infinite-scroll record caches synchronized when a host refreshes data without reporting a dataset-change flag.

## 1.0.69.0 / control 1.0.90 - 2026-09-04

- Respect the `EnableInlineEdit` control input instead of enabling editing unconditionally.
- Use Dataverse primary-key metadata for lookup searches, including tables whose key is not named `<table>id`.
- Reject malformed whole-number and decimal edits instead of silently truncating them.
- Restore the original dataset page after selecting records across pages and clean up pending page waiters.
- Prevent stale lookup-search responses and timers from updating unmounted editors.
- Prevent lookup cells with no resolvable target table from entering a stuck edit state.
- Remove a render-time React state update, improve selection callback freshness, and harden clipboard handling.
- Scope grid CSS to this control, add missing infinite-scroll resource strings, and improve keyboard labels.
- Package the Dataverse solution as unmanaged only.
- Repair and expose local core-logic and UI validation scripts.
- Key row selection by Dataverse record ID so sorting, filtering, and infinite scrolling cannot retarget selected indexes.
- Remove filter modes whose condition operators are unsupported by model-driven dataset controls.
- Respect dataset columns that disable sorting, preserve date-only values without timezone conversion, and provide a manual infinite-scroll fallback.
