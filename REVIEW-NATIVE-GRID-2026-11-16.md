# DrivenGrid native presentation candidate — 16 November 2026

Prepared control **1.0.118** / unmanaged solution **1.0.97.0**, based on committed
release `4ec3a07be17092283b76b27c0e23120f0e13884c`. This candidate has passed local
automated checks and packaging. It has **not been deployed or visually accepted**.

## Requested behaviour

- Move the surrounding grid appearance closer to the modern Power Apps grid.
- Preserve the existing column dropdown's design, content and sort/filter options.
- Clicking a header opens its dropdown. Clicking the same header closes it.

The presentation changes use neutral headers, selected rows and hierarchy borders,
clearer text, 16-pixel square checkboxes and white child panels. The model-driven
form supplies the surrounding section, so the extra inner frame is removed.
Row heights, column-width arithmetic, virtualization, native commands and data
queries are unchanged. The complete column-dropdown CSS block matches the baseline.

Design reference: Microsoft's [modern Fluent model-driven app guidance](https://learn.microsoft.com/en-us/power-apps/user/modern-fluent-design)
and [Power Apps grid control](https://learn.microsoft.com/en-us/power-apps/maker/model-driven-apps/the-power-apps-grid-control).
These inform the direction; this candidate is not a claim of pixel-perfect parity.

## Checks completed

- Added a regression that fails on the baseline's second header click and passes
  after the toggle fix. It also covers reopening with clean draft state, switching
  headers, condition-editor dismissal, outside clicks and child-header toggling.
- Full `npm run test:local` passes, including filters, editing, selection, native
  hierarchy, 10,000-row paging and the 42,000-record local fixture.
- Authoritative `dotnet build` of the solution in Release mode passes, including
  TypeScript, ESLint and unmanaged packaging. Reused existing npm dependencies;
  no lockfile or dependency changes. Three existing Webpack size/performance
  recommendations remain (730 KiB bundle).
- `git diff --check` passes.
- The interactive sample compiles and its HTTP endpoint serves the changed CSS.
  Start it with `npm run demo`; it binds only to `http://127.0.0.1:8183/` and uses
  disposable in-memory records. It does not contact Dataverse.

Package: `CMModelDrivenGridSolution/bin/Release/CMModelDrivenGridSolution.zip`.
Verified `Managed=0`, one PCF root, solution `1.0.97.0`, control `1.0.118`.
SHA-256: `3DD88DFE13C591CA199BDF821EC75B69FF2E62522A8304087D344C8B2898EDD6`.

## Remaining release work

Computer Use stopped because it could not determine the active browser URL with
enough confidence to enforce its policy. No further browser automation was attempted.
The preview's visual layout and the signed-in Dynamics runtime have not been
verified for this candidate. No Dataverse write was performed.

Before sandbox release, complete the local visual check, create and verify a fresh
exact rollback baseline, verify current sandbox identity/versions/publisher/pending
changes, and use the guarded unmanaged deployment workflow. Prefer targeted
component publication. Do not reuse a September rollback without fresh locks.
Verify installed versions and all three asset hashes, then smoke-test a fresh
signed-in QA form: parent/child header toggles, header switching, Escape/outside
dismissal, unchanged checklist/conditions, selection/native commands, expansion,
editing cancellation and paging. Export the final unmanaged package and evidence
after acceptance. No production or form/view/schema changes are included.
