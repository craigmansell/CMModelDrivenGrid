# DrivenGrid sandbox validation — 14 September 2026

## Current decision

**Deployed and verified: control 1.0.117 / unmanaged solution 1.0.96.0.** Full
local suite, final signed-in smoke and fixture restoration audit pass. Ready for
sandbox user acceptance, not production sign-off. Live clipboard readback remains
a manual check. Intermediate 1.0.113–1.0.116 builds are superseded, not accepted releases.

Target: Capricorn sandbox, organization `27252def-2a75-f111-b27b-000d3ab8db73`.
No production, schema, form, view, ribbon, relationship, publisher or security
changes belong to this PCF-only batch.

[Open the QA form with test data](https://capricornsandbox.crm4.dynamics.com/main.aspx?appid=f29c988c-f0cc-f011-8544-6045bde18fc1&pagetype=entityrecord&etn=cm_itemtemplate&id=79357ab7-61b3-4f6f-bc6b-946fa6a95dae&formid=e1c3c7c8-2d9d-4a11-8011-b1b9517f81f3)

Form: **DrivenGrid QA - Full subgrid tests**. Record: **DrivenGrid QA 2026-09-09 -
disposable test records**. Use a fresh tab after the update: one already-open
Dynamics tab reused older scripts even after reload and was excluded from acceptance.

## Issues found and fixed

1. **Paging promise lifetime (1.0.114):** new rows could render before the host
   page promise settled. Prefetch reused it and showed a false empty-page warning.
   Keep the promise guard until actual completion.
2. **System-managed editing (1.0.115):** optional dataset read-only flags did not
   identify Created On / Record Created On. Read bounded public Dataverse attribute
   metadata, fail closed while eligibility is unknown, propagate to both levels,
   and check again immediately before saving. Calculated/rollup/derived attributes
   remain locked. Server security/custom business rules remain authoritative;
   this does not infer every custom form lock.
3. **Retained rows after native selection (1.0.116–1.0.117):** command selection
   can restart physical paging while visible parents remain cached. Catch up through
   cached pages to new rows or EOF, with duplicate/500-page guards and selection
   handover. Power Apps can replace the dataset object on updateView: inspect the
   delivered page, not the issuing snapshot. Both variants failed their new
   regression before correction.

Excel menus, dotted checklist, AND/OR conditions, Shift sorting/ranges, the outer
command bar and lazy hierarchy remain. No fullscreen, Load more button or child
toolbar/count strip was introduced.

## Live coverage completed before the final smoke test

Version labels distinguish observed UI tests from final-binary acceptance.

| Area | Result | Version |
| --- | --- | --- |
| Filter matrix | **174/174 pass**, 87 root + 87 child, exact IDs against independent Dataverse queries | 1.0.114 |
| Text filters | Ten standard predicates, apostrophe, no-match, literal %, underscore and bracket cases | 1.0.114 |
| Other filters | Nine number, four lookup, four status, all 53 DateTime conditions, two optional-date presence cases at each level | 1.0.114 |
| Child paging/nesting | 50+14 children; G01/G02 expansion; no leaf arrows | 1.0.114 |
| Native Edit | Selected C03 opened exact child ID, not its parent | 1.0.114 |
| Reversible cost edit | Escape preserved 20; save 23.45 API-verified; restored 20 | 1.0.114 |
| Reversible lookup | Alpha API-verified, then cleared to original null using Ctrl+A/Backspace/Tab | 1.0.114 |
| Native state commands | C03 deactivated then activated; original state/status 0/1 verified | 1.0.114 |
| Excel checklist | Alpha search/cancel; Alpha red Apply; descending sort kept filter | 1.0.114 |
| Compound queries | Shift secondary sort, same-column AND/OR, cross-column AND: exact keys passed | 1.0.114 |
| Read-only dates | F2 opens zero root/child editors; normal Unit Cost editor remains available | 1.0.115 fresh tab |
| Selection | All 33 roots; child selection clears roots; all 64 children; collapse clears selection | 1.0.115 |
| Native Delete | Only a newly created marked disposable child deleted through outer command bar; ID/marker absence verified | 1.0.115 |

Delete-only child `5e43239e-e5cb-4d81-8171-f3f56767dc9d`, marker
`DGQA20260914:V115-DELETE-ONLY`, has been removed. It was not a business record.
Historical C65 deletion was preserved. No Share, Assign, Email, Flow, distributed
reports or permission changes were executed.

## Local release validation

Full `npm run test:local` passes on 1.0.117: core/parity/metadata, inline editing,
range copy, grid/review regressions, filters, child/native selection, hierarchy,
and development tooling. All 16 previous review gates remain in the normal suite.
New tests cover pending/failed/stale metadata, guarded saves, promise lifetime,
selection handover and mutable/replacement dataset catch-up.

- Strict-host 10,000 rows: 200 batches, no missing/duplicate rows or overlapping calls.
- Local 42,000-record hierarchy: 2,000 parents, exactly 20 children each, every branch
  checked. **Not a 42,000-row Dataverse benchmark.**
- Production dependency audit: zero known vulnerabilities at this check.
- Nine upstream development-tool audit entries remain (five moderate/four high)
  in PCF telemetry/BrowserSync chains; no forced incompatible upgrade.
- Focused test TypeScript checks and authoritative unmanaged Release build including
  ESLint pass. The existing three Webpack size/performance warnings remain (730 KiB).

## Final 1.0.117 signed-in smoke and restoration

Fresh tab, normal handler, published bundle verified against the tested package:

- Initial automatic paging shows 20 roots without Retry; only roots 1/2 have arrows.
- Select All chooses all 33 roots. Switching to C03 clears root selection and enables
  native Edit/Activate/Deactivate. Child Select All loads/selects all 64; collapsing
  the branch clears selection.
- The previously failing selection-to-totals sequence now completes without an
  empty-page warning. Full-level Unit Cost sum **799.36**; Active **30 / 679.36** and
  Inactive **3 / 120**, matching the independent fixture reference.
- Root Created On and child Record Created On open no F2 editor; normal Unit Cost
  editing remains available. Parent R03 saved **21.45**, independently read back,
  then restored to **20**.
- Six representative final-version filter checks passed exact IDs: root text,
  numeric range, lookup, status, date range, and child literal percent.
- Native Edit opens exact C03 ID `35ed8b84-51be-4e40-ac26-d4fe994032ea`, Gamma red,
  Order 1,003. Returned without modifying its form.
- Live Shift-arrow selection renders the expected four cells (R01/R02, Order/Name).
  Copy/paste readback is **not verified**: the automation browser's virtual clipboard
  was empty and blocked paste. This is not counted as a product pass or proven defect.
  Manually press Ctrl+C on a range and paste into Excel to complete this check.
- Final audit: **106 original fixtures, zero changed business fields**. The sole
  newly created delete-test child is absent. Final QA page is unfiltered/unselected,
  with no loading/error overlay. Screenshot inspected for alignment and leaf arrows.

Evidence: `live-final-v117.json`, `records.after.json`, `record-audit.json` and
`v115-delete-fixture-verified.json`. The 174-case full matrix was on 1.0.114, not
silently rerun on 1.0.117; the latter contains focused paging/editability corrections.

Final tested input archive SHA-256:
`2425AC6F4AD302469CDB3F13C987678F89F62983BD55F3F4BFF5733745B30058`.
Published bundle SHA-256:
`A7190A4DE5D9B41D66745372F518209135C3F54F62DFE4FB871F1614DB9B3151`.

Final unmanaged exports are verified in the deployment evidence folder:

- `CMModelDrivenGridSolution.after.unmanaged.zip` — solution 1.0.96.0 / control
  1.0.117; SHA-256 `E3628BF98D587E4AB565B1E7CCBF2BD398741EB9011D9122F1FC17A7812F5B5A`.
  Its 747,627-byte bundle matches the published/tested bundle above.
- `CapricornRelease20260622.after.unmanaged.zip` — companion release snapshot;
  SHA-256 `66A5C5D8F04E2731090EC440EF267AF30B6949879650A199C55BFC555D29023F`.
  This solution contains no PCF control assets and is **not** a substitute for the
  separate DrivenGrid package. No solution-boundary change was made to add it.

## Deployment safeguards and evidence

CrmPackage folders:

- `outputs/deployment-backups/drivengrid-v113-sandbox-20260914-095030/`
- `outputs/qa/drivengrid-v113-20260914/`

One exact rollback baseline at `2026-09-14T08:52:33.0586333Z` is reused under the
unchanged 12-hour guard. Each import checks sandbox identity, versions, previous
bundle hash, publisher, exact candidate hash, unmanaged identity, eight expected
archive entries and absence of unrelated pending UI changes. Published/draft forms
and views must remain unchanged. Targeted resource publication is tried first;
broad publication requires owner approval and a fresh pending-work audit. The
initial 1.0.113 publication pause was resolved after the owner resumed.

Rollback SHA-256:

- PCF: `9BAA0884CE7C3C875DEDAC71DD9AFE336935CFE111DB731623C73A77C4CA9B47`
- Full release: `38D9AAC27646088B234FCE61118E4691DFAFA893699D5966CA42D0C93F7030ED`

Per-version import intents/results, publish receipts when needed, published asset
hashes and deployment results preserve intermediate history. `records.before.json`
contains 106 baseline fixtures, including the outside template/sentinel. Final
readback matches all original business fields, ignoring only annotations, modified-on
and version number.

Filter evidence: `filter-oracle.validated.json`, `live-root-first60.json`,
`live-root-final27.json`, `live-child-first32.json`, `live-child-all87.json`.
The initial `filter-oracle.json` generator failure is preserved and not counted
as a product failure or UI pass. `live-interactions-through-v115.json` includes
the failed totals case requiring the final retest. Delete intent/creation/absence
evidence is separate.

## Remaining boundaries

Sandbox QA is **not production approval or complete native-grid parity**.
See [feature comparison](FULL-FEATURE-COMPARISON-2026-09-10.md).

- Checklist values come from loaded rows, not all distinct server values. Condition
  filters target the full dataset within documented safeguards.
- Groups/totals load the complete matching level client-side, not native server
  aggregates, currency conversion or automatic descendant totals.
- Shared native child commands require the same-table hierarchy setup; arbitrary
  different-table children do not have identical command behavior.
- Multi-user/roles, mobile/touch, assistive technology, offline, all locales/timezones
  and large live-Dataverse performance acceptance are outside this signed-in run.
- Native New/Add Existing were not exercised end-to-end in this final run; Share,
  Assign, Email and Flow actions were deliberately not executed. Do not interpret
  command visibility or the API-created disposable fixture as those UI passes.
- Not every Microsoft platform/customizer feature is reproduced. DrivenGrid remains
  an independent PCF with the requested Excel features.
