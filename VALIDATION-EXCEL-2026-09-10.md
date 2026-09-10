# DrivenGrid 1.0.112 — Excel-preserving feature pass

10 September 2026. Independent DrivenGrid PCF retained; no Microsoft grid replacement.
Unmanaged candidate solution **1.0.91.0**, control **1.0.112**. This package is **not
deployed**. Last verified sandbox remains 1.0.89.0 / 1.0.110.

## Verified locally

- Full `npm run test:local` passed: core/filter semantics, condition UI, inline edit,
  range selection/clipboard, grid/hierarchy interactions, native-host paging model,
  complete 42,000-record fixture and development tooling. No suppressed errors in
  the final run. Core suite was repeated after the two final lint refinements.
- Focused TypeScript check passed for harness, core, presentation, grid and cell
  selection regression scripts. Authoritative Release build passed, including
  manifest validation, TypeScript, ESLint and unmanaged solution packaging.
- New tests cover date intervals/week boundaries/timezone-independent DateOnly,
  chronological grouping/collapsed defaults, keyboard movement/extension and exact
  quoted TSV, jump-filter preservation, storage scoping and corrupt/blocked storage,
  primary/none navigation and changes to rendering settings.
- Full lookup: explicit result highlighting and Enter, unchosen Tab cancellation,
  opening/returning/cancelling the dialog, retained editor/refocus, no duplicate save,
  unmount, wrong table/ID, rejected configured restrictions, view change and read-only
  changes across asynchronous work. These are controlled mocks of the supported
  host lookup interface, not a signed-in Dynamics dialog acceptance test.
- Automatic scrolling works without IntersectionObserver through the scroll-distance
  fallback. No Load more or fullscreen button, child toolbar or Add to current sort
  item was added. Existing searchable Excel checklist, dotted guides, conditions,
  filter chips and shared selection/commands remain in place.

## Browser evidence (local fixture, no Dataverse connection)

- A–Z `A` returned 334 matching parents, `All` restored 2,000. Excel column menu
  still shows search, (Select All), values, type-specific conditions, OK/Cancel.
- Moving Project behind Budget survived reload. Reset restored original order.
- Ctrl+End focused loaded row 50; automatic prefetch reached 75. Shift+Down moved
  focus/extended the range to row 51. This is a loaded-page boundary check, not
  full Excel keyboard parity or a paste-to-Excel acceptance test.
- Month grouping completed all 2,000 parents, with 31 initially collapsed groups
  including 118 blanks, chronological June 2025–November 2027 labels, zero row DOM
  nodes while collapsed and zero child requests. Clearing grouping rendered 30
  row nodes instead of 2,000. Grouping still loads the entire current level in memory.
- Fresh Select All reached exactly 1,000 selected parents with zero child requests
  and 50 rendered row nodes. Expanding Atlas project 001 returned its 20 children
  in one request. Selecting Option 002 and using the one outer sample Open command
  opened **sample-1-child-2**, identifying Atlas project 001 as its parent. No child
  toolbar was present.

## Release readback

Package: `CMModelDrivenGridSolution/bin/Release/CMModelDrivenGridSolution.zip`.

- Solution 1.0.91.0; control 1.0.112; Managed = 0.
- Publisher CraigMansell / cm; exactly one control, one records dataset, 39 inputs.
- All referenced display/description resource keys exist.
- Bundle: 741,638 bytes. Webpack reports three existing size/performance categories
  of warning; zero lint errors. This is not a cold-load performance benchmark.
- Package SHA-256: `F371B3E4F3EA28F8D937326BD57E2A2AA2D226362CC7A2FF565515E4D31AA4AC`.
- Bundle SHA-256: `C41841934C71F1D8F72487BB6C361E6369A50B743CFAE29B6E39B19AD308FED3`.

## Limits and release boundary

No Dataverse writes, deployment, form/view/ribbon changes or production actions.
The full current feature matrix is in [the comparison](FULL-FEATURE-COMPARISON-2026-09-10.md).
Remaining gaps include independent/cross-table child datasets, lookup/choice label
pattern filters, specialist editors, native event/customizer contracts, broad
localization/accessibility/security/device coverage and live-scale acceptance.
DateTime grouping currently uses browser calendar time, not verified Dataverse
user-timezone/TimeZoneIndependent semantics. Layout persistence is browser-local,
not native/cloud personal views. The full lookup dialog is constrained to the cell's
current target table; its broader native UX remains unverified.

Before sandbox release: create a fresh exact unmanaged rollback, validate environment
and component locks, deploy the candidate, perform a signed-in QA smoke on the
isolated form without bypassing normal business handlers, and capture final unmanaged
export/readback. No claim of complete 1:1 native parity is made.
