# DrivenGrid presentation and scrolling — 16 November 2026

**Published sandbox candidate: control 1.0.119 / unmanaged solution 1.0.98.0.**
Local automated and browser checks pass. Installed versions and all three resource
hashes match the tested package. **Signed-in menu/command smoke acceptance remains
pending because Computer Use's connection became unavailable.** No production change.

## Requested behaviour and implemented scope

- Use a more native Power Apps grid appearance: neutral headers/selection,
  clearer text and separators, square checkboxes, white child panels and a neutral
  hierarchy rail. Keep row heights, column sizing, virtualization and native commands.
- Hide the alphabet bar by default and scroll automatically. The preview no longer
  forces the alphabet on; the shared Grid now defaults to infinite scrolling,
  matching the manifest and Dataverse host. Explicit alphabet/paged options remain.
- Refresh the header dropdown cards with rounded corners, lighter borders/shadows,
  clearer action/title styling and a separated footer. All dropdown callbacks,
  checklist/condition logic, sort/filter operations and Apply/Cancel functions remain.
- Click a header to open its card; click the same header to close it without
  applying drafts. The same handler is used by parent and child grids.

Design references: [modern Fluent model-driven apps](https://learn.microsoft.com/en-us/power-apps/user/modern-fluent-design)
and [Power Apps grid control](https://learn.microsoft.com/en-us/power-apps/maker/model-driven-apps/the-power-apps-grid-control).
This is a closer visual match, not a claim of pixel-perfect native parity.

## Local verification

- New header toggle and default-scrolling regressions fail against their previous
  implementations, then pass after correction. Explicit manual paging still passes.
- Full `npm run test:local` passes: filters, editing, selection, metadata guards,
  hierarchy, 10,000-row paging and the 42,000-record in-memory fixture.
- Authoritative unmanaged Release solution build passes, including TypeScript and
  ESLint. Existing dependencies were reused; no lockfile changes or npm installation.
  Three existing Webpack size/performance recommendations remain (730 KiB bundle).
- `git diff --check` and deployment-script PowerShell parsing pass.
- Browser: no alphabet or manual paging buttons; root rows loaded automatically
  from 50 to 75; child rows from 25 to 50. Parent and child headers open and close.
  Searching the checklist for Atlas then Cancel preserves 180 records and no filters.
  The number-condition editor retains all nine operators and Apply/Cancel controls.
  Refreshed cards render with 8-pixel corners and reachable checklist/footer.

Preview: `npm run demo`, loopback only at `http://127.0.0.1:8183/`, disposable
in-memory records and no Dataverse connection. Screenshot in CrmPackage:
`outputs/qa/drivengrid-native-look-20261116/preview-dropdown.jpg`.

Exact tested unmanaged ZIP SHA-256:
`FD4B1D337461856AFD464BF9130BBE0ABDE9651B9FE9A3834ED020A06BA6C2A7`.
Verified one PCF root, `Managed=0`, solution `1.0.98.0`, control `1.0.119`.

## Sandbox deployment evidence

Locked target: `https://capricornsandbox.crm4.dynamics.com`, organization
`27252def-2a75-f111-b27b-000d3ab8db73`, existing CraigMansell publisher and control.
Helper: `scripts/Deploy-DrivenGridNativeLookSandbox.ps1`.

Fresh exact rollback and receipts in CrmPackage:
`outputs/deployment-backups/drivengrid-native-look-sandbox-20261116-090813/`.
The helper verified the exported rollback's three assets against the published
baseline and its exact 1.0.117 bundle before any import.

- Before PCF unmanaged export SHA-256:
  `5D20AD796F675D45181114A845C254A3CE8E18AC4BED54B6E3C577CE2936AB1F`.
- Before full Capricorn release unmanaged export SHA-256:
  `1B0B9DEC4EF9900CAC431760A1C7F8E2E6745FFAC780E0D02C0C3ECFDDEED5B5`.
- Import completed successfully, 100%, at `2026-11-16T09:12:19Z`.
- Targeted publication activated resources but left the manifest at 1.0.117.
  Guarded Publish All completed at `2026-11-16T09:15:28Z` after repeating exact
  form/view/publisher checks and confirming the entire pending UI inventory empty.
- Installed 1.0.119 / 1.0.98.0 verified at `2026-11-16T09:15:31Z`. Published/draft
  Item Template forms, Item Template Option views and publisher remain unchanged.
- All three published resources match the tested package. Bundle SHA-256:
  `F14B04F6E4618F2310E90A0351D6575C5599E9AC6FDBF004E9CDF02BBA4FCA33`.
- `CMModelDrivenGridSolution.v119.tested.unmanaged.zip` preserves the imported bytes
  beside the rollback. `v119-candidate-package.json` records identity and status.

No business records, forms/views, schema, relationships, ribbon, security or
production customizations were changed by this batch.

## Remaining runtime acceptance

A genuinely fresh signed-in QA form loaded the existing disposable template with
33 records, no alphabet bar and no visible grid errors. During the first live menu
check, browser commands timed out. Subsequent tab/inventory checks also timed out;
the native-window fallback reported that Computer Use's pipe was unavailable.
No further browser automation was attempted after that connection failure.

This is not a pass or diagnosed product failure for the live menu/command path.
The deployment receipt retains `runtimeSmokeRequired=true`. The final after-state
unmanaged export is deferred until the following checks complete:

1. Open a fresh QA form using the link below and confirm the current 8-pixel cards.
2. Parent/child header toggle and switching; searched checklist Cancel/Apply;
   condition editor, sort/filter preservation and outside/Escape dismissal.
3. Automatic root/child scrolling, expansion/collapse and selection cleanup.
4. Existing native command targeting; open the exact selected child with Edit,
   cancel without saving, and cancel an inline editor without changing a record.
5. Re-run helper Verify against the preserved baseline and hash-locked input.
   If freshness/concurrent-change locks no longer pass, review the new state first.
6. After acceptance, run helper Export and record the final unmanaged hashes and
   runtime evidence. Never bypass the locks or mark this candidate accepted early.

[Open the sandbox QA form](https://capricornsandbox.crm4.dynamics.com/main.aspx?appid=f29c988c-f0cc-f011-8544-6045bde18fc1&pagetype=entityrecord&etn=cm_itemtemplate&id=79357ab7-61b3-4f6f-bc6b-946fa6a95dae&formid=e1c3c7c8-2d9d-4a11-8011-b1b9517f81f3).
