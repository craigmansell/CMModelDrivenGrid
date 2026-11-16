# DrivenGrid Power Apps styling — 16 November 2026

Installed in sandbox: control **1.0.122** / unmanaged solution **1.0.101.0**.
Local validation, the authoritative build and signed-in functional acceptance pass.

## Requested scope

Bring the appearance closer to the actual native Power Apps grid while retaining
the hierarchy arrows, full-header click to open/close, and every existing function.

The native sandbox reference uses underlined links in `rgb(17, 94, 163)`, white
headers, approximately 42-pixel headers/rows, faint `rgb(240, 240, 240)` cell
separators and `rgb(207, 228, 250)` selected rows. Its dropdown arrows sit beside
the labels and numeric headings align with their values.

The update changes CSS plus one presentation property in `Grid.tsx`: a numeric
header class. No event handler, data query, selection algorithm, filter, sort,
editing, hierarchy, paging, native-command or responsive-sizing logic changes.
The existing dropdown-card stylesheet and hierarchy-arrow dimensions remain intact.

## Validation and package

- Full `npm run test:local` passes, including parent/child header toggles,
  editing and lookup guards, filtering, selection, retained children, native
  commands, responsive sizing, 10,000-row paging and the 42,000-record fixture.
- Browser-only local checks confirm 42-pixel header/row geometry, white header,
  exact native selected color, numeric-heading alignment, adjacent carets, no
  default alphabet bar and one rounded outer frame. Clicking the blank right
  side of a header opens its menu; clicking the same point closes it.
- The authoritative Release build, TypeScript and ESLint pass. Three existing
  Webpack size recommendations remain. Dependencies/lockfile are unchanged.
- Exact package: `Managed=0`, one PCF root, eight expected files, control 1.0.122
  and solution 1.0.101.0. SHA-256:
  `5929CADECFF03A3BA4303ABDFB0BCAAF38190C64E6E6D5F38D2218E568586902`.
- Local preview and logs are in CrmPackage
  `outputs/qa/drivengrid-power-apps-style-20261116/`.

## Sandbox safeguards

Target: `https://capricornsandbox.crm4.dynamics.com`, organization
`27252def-2a75-f111-b27b-000d3ab8db73`.
Helper: CrmPackage `Migration/Deploy-DrivenGridPowerAppsStyleSandbox.ps1`.

Fresh exact rollback and evidence:
`outputs/deployment-backups/drivengrid-power-apps-style-sandbox-20261116-154121/`.
The prior installed 1.0.121 / 1.0.100.0 assets matched the rollback before any write.

- Before PCF unmanaged package SHA-256:
  `205BA08BEF295D92737348C8A17FD9430DD668F521D0DAE63F76472FCB57E2CD`.
- Before full CapricornRelease20260622 unmanaged package SHA-256:
  `2BF6D6C2E8F5CD8C7731A7AF386BEA7AFF60DA9EF0299C94F1B70D3053DD8A00`.
- The native full-tab QA form is locked to XML SHA-256
  `BDE9E95104C9515C3C0EE52B2F325C6BD02412FDDE4BF6430AFB2F160645B98A`.
  Every Item Template form and Item Template Option view is frozen, including drafts.
- Organization, control, publisher, unmanaged status, versions, package contents
  and asset hashes are locked. Unrelated pending customizations stop publication.
- Final exports require all signed-in parent/child toggle, hierarchy, paging,
  native Edit, inline Cancel, full-tab and style checks, with hashed browser evidence.
  Screenshots are hash-checked when available; capture failure is recorded explicitly.

No form, schema, relationship, solution-boundary, business-record or production
change is part of this PCF styling batch.

## Installed result and signed-in acceptance

- Imported the exact tested package. Targeted resource publication left the old
  manifest; audited Publish All completed at `2026-11-16T15:47:17.8745166Z` after
  confirming no unrelated pending UI customizations. Installed versions and all
  three assets match the candidate. Forms, views and publisher remain unchanged.
- Published bundle SHA-256:
  `D06C1083E03D18B7E4569776DEF6760E995620F06E9928E800A610E2E026DD4F`.
  Published stylesheet SHA-256:
  `709FF78BC2542F8444E0F3939EE5B0E5B3321C387FCBCA5FD12B693EE043C273`.
- Parent and child Name headers open and close on repeat clicks. The parent test
  also uses the blank right side of the full header hit target and its search field.
- Hierarchy expansion loads 50 children; internal scrolling automatically loads
  all 64. Selecting child 64 and using native Edit opens `DGQA child 64`, record
  `3f8257c1-4b8f-4320-b735-0b26050adb77`. Back returns to the QA form without saving.
- Inline Unit Cost opens at zero; Escape restores the unchanged `£0.00` display.
- Signed-in DOM measurements confirm a white header, blue selected row
  `rgb(207, 228, 250)`, underlined `rgb(17, 94, 163)` links, three right-aligned
  numeric headings, an 8-pixel outer frame, approximately 42-pixel header/rows,
  no alphabet bar, and a footer within one rounding pixel of the viewport edge.
- A normal browser reload was needed to replace the cached v121 presentation.
  The reloaded QA preview shows the new v122 classes and computed styles.
- Browser CDP screenshot capture repeatedly timed out. The real local screenshot
  is `outputs/qa/drivengrid-power-apps-style-20261116/local-preview.jpg`; it is a
  harness preview, not a sandbox screenshot. The earlier stale sandbox capture is
  labelled `sandbox-stale-before-refresh.jpg` and is excluded from acceptance.
  Runtime acceptance uses signed-in AX/DOM evidence, without a sandbox pixel claim.
- Hashed evidence: `browser-runtime-evidence.json`, SHA-256
  `A7646830129E1072CB4F3F05EE12814F56CD644F977180B348A659A053565237`.
  The nine passing checks are recorded in `runtime-smoke.json` beside the rollback.

## Final unmanaged exports

Both final exports passed identity, unmanaged, version, exact PCF asset,
form/view/draft, publisher and pending-customization checks after runtime acceptance.
The final receipt is `final-packages.json` beside the rollback and browser evidence.

- PCF solution 1.0.101.0 / control 1.0.122 SHA-256:
  `DE493407E8BAC13674BE37AF679EE23926B069193D99406DF72AB0C8BB2C821D`.
- CapricornRelease20260622 remains unmanaged at 1.0.0.6, SHA-256:
  `D3C82A99FF1048F4501F99CAD36FBA3958ADE283B8F237ADC28DCC9DC79F50C2`.
