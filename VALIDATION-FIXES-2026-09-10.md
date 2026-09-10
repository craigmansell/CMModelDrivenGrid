# DrivenGrid 1.0.113 — pre-sandbox defect fixes

10 September 2026. Unreleased control **1.0.113**, unmanaged solution **1.0.92.0**.

**The three reproduced 1.0.112 review defects are fixed and their local release
checks pass.** This is not signed-in sandbox acceptance or a native-feature parity
certificate. The sandbox was not changed or re-read; its last recorded verified
release remains control 1.0.110 / solution 1.0.89.0.

## Fixes and regression evidence

1. **Ordinary filter rejection/loading:** deep-copy and restore logical/physical
   filters and all sort priorities after synchronous host rejection. An error-only
   completion signal releases local loading even with unchanged records or repeated
   identical errors. Tests cover two identical failures, successful retry, genuine
   loading, unchanged-result completion, nested filters and failed refresh. If
   rollback also fails, row selection and Select All are blocked. Failed host
   refresh remains blocked; successful refresh follows the accepted host filter;
   a new view discards the old recovery state.
2. **Compact keyboard/ranges:** navigation, copying and card rendering share the
   displayed column boundary. End reaches the last displayed field; Right stays
   there; Home, Up/Down and Ctrl+Shift+End work. The exact two-row/two-field TSV
   excludes the third undisplayed field. Reflow clears old ranges, changes bounds
   immediately and cancels a pending next-page keyboard target. Table End still
   reaches the last of all visible columns.
3. **Late lookup editability:** immediately before every permitted write, recheck
   current column flags/presence/alias/type and table/view identity, not only the
   pre-metadata state. Tests cover isReadOnly, readOnly, disableEditing, removal,
   joined alias, changed type, view, table and property-rejection retry, each for
   setting and clearing. No write follows the state change; the retry case records
   only its first confirmed rejected attempt. Existing successful lookup cases,
   cancellation, configured restrictions and unsafe-retry prevention also pass.

The review probes now run as part of `npm run test:local`; no opt-in-only release
gate is left outside the normal suite. Sixteen named review groups pass. Host
failure and save tests use local mocks, not Dataverse mutations. The ordinary
filter fix does not claim to repair every asynchronous host/network failure or
prove the cause of earlier sandbox screenshots.

## Final checks

- **Full local suite: pass.** Core/presentation, inline editing, range/quoted TSV,
  grid/child interactions, filter semantics/editor/evaluator, server-predicate
  compilation/scope/limits/rollback, paging/selection and development tooling.
- **Large fixtures: pass.** Exactly 2,000 parents and 20 children per parent:
  42,000 records, all 80 parent batches and all 2,000 child branches. Separate
  strict-host paging model: 10,000 rows/200 batches, no missing/duplicate rows or
  overlapping host calls. These are local fixtures, not a Dataverse load benchmark.
- **Focused harness/test TypeScript: pass.** Includes the added review tests.
- **Authoritative unmanaged Release build: pass.** Includes manifest checks,
  TypeScript and ESLint; final build completed 16:59 local time. Three Webpack
  bundle/entrypoint/performance warnings remain; no compile or lint errors.
- **Browser: pass.** In the 2,000-by-20 fixture, compact End/Right/Home/Shift+End
  stay within Project/Budget/Status. Table End reaches Active (column 5). Initial
  loading settles without alerts; Excel searchable checklist and A-to-Z sort
  remain available. Expanding Atlas project 001 makes one child request for 20
  rows; selecting Option 002 and using the shared Open command shows exact ID
  `sample-1-child-2`, not its parent. The sample preview was reset afterwards.

Reproduction commands:

```text
npm run test:local
npx tsx scripts/local-grid-regression-test.tsx --pre-sandbox-review
npx tsc --noEmit --target es2020 --module commonjs --jsx react --esModuleInterop --skipLibCheck scripts/harness-demo.tsx scripts/local-core-logic-test.ts scripts/local-parity-logic-test.ts scripts/local-grid-regression-test.tsx scripts/local-cell-selection-test.tsx scripts/local-pre-sandbox-review-probes.tsx
dotnet build CMModelDrivenGridSolution/CMModelDrivenGridSolution.cdsproj --no-restore --configuration Release
```

Preview: [2,000 parents with 20 children each](http://127.0.0.1:8183/?scenario=parents-2000-children-20).
Compact preview: [three displayed fields](http://127.0.0.1:8183/?scenario=parents-2000-children-20&layout=list).
These links require the local demo server; they are not sandbox links.

## Exact local package readback

Archive: `CMModelDrivenGridSolution/bin/Release/CMModelDrivenGridSolution.zip`.

- Solution 1.0.92.0; Managed=0; publisher CraigMansell / cm.
- Only root component: type 66, `cm_CraigMansell.CMModelDrivenGrid`. No entities,
  forms, relationships, roles, workflows or plugins in this package.
- Control 1.0.113; one dataset; 39 direct inputs; all 91 distinct resource keys
  present and every manifest display/description key resolved.
- Bundle: 743,630 bytes.
- Bundle SHA-256: `F30DC138710B6242005B200AA521A97BB39BDC8F0E967E4123CCA13EC7A798C4`.
- ZIP SHA-256: `948FADDD33A48623DC8AF3764AA9415EFECE60ACBBF79AFAAC8C39A69DBDE69E`.

## Release boundary and remaining acceptance

No Dataverse import/publish, schema/form/view/ribbon changes, server record writes
or production actions were performed. The Dynamics safety workflow kept validation
local and informed the immediate pre-write guards; no native command rules were
overridden. Dependencies/lockfile were not changed. The prior review's audit (zero
shipped-dependency findings, nine development-tool entries) was not rerun here.

Before a sandbox write, capture a fresh exact unmanaged rollback and validate
organisation, installed versions, publisher and component scope. After an approved
deployment, use the isolated QA form with normal business handlers to verify actual
parent/child commands, save/readback, lookup dialog, security, filter failures and
server paging. Do not treat earlier-version live results as 1.0.113 acceptance.

Excel features remain intact; independent/cross-table child datasets, label-pattern
lookup/status filtering, specialist editors, native event/customizer contracts,
full timezone/device/accessibility/security and live-scale acceptance remain open.
See the [full comparison](FULL-FEATURE-COMPARISON-2026-09-10.md). This fix batch does
not claim to add those remaining features.
