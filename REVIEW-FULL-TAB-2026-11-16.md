# DrivenGrid full-tab sizing — 16 November 2026

**Installed in sandbox: control 1.0.121 / unmanaged solution 1.0.100.0.**
The full local suite, Release build and signed-in browser functional checks pass.
The owner explicitly chose **Continue without a screenshot**. Acceptance uses
browser accessibility actions and DOM geometry; no pixel-rendering sign-off is claimed.

## Behaviour

Support the native **Expand first component to full tab** option shown by the owner.
The form supplies the work area; the grid must occupy that area and resize with it.
The subgrid must be the first component, and its optional Sub Grid Height should
be blank for automatic sizing.

- A positive height supplied by Power Apps takes precedence over the embedded
  fixed-height fallback. Previously the fallback always overrode that allocation.
- With no explicit height, a component-owned wrapper measures its own available
  height and observes resizing. A short tab area can be smaller than the grid's
  previous 320-pixel automatic minimum without pushing the footer out of the area.
- Native subgrid hosts can include command-bar space in their height. Automatic
  content sizing also respects the visible viewport below the component's own top
  edge. Explicit allocations, fixed heights and offscreen embedded grids are retained.
- Window resizing is supported, with ResizeObserver observation and cleanup.
  No Dynamics ancestors, private DOM, form globals or tab internals are inspected.
- The same hierarchy component remains mounted, preserving selection and grid
  state. Menu logic, sorting/filtering, infinite scroll, native commands and child
  sizing are unchanged. No data refresh or write is initiated by resizing.
- The existing optional fixed-height fallback remains available for embedded grids.

The maker label/help now explain when to leave the fixed height blank. No new
manifest input or separate full-screen command was added.

Native option reference: [Microsoft form setup](https://learn.microsoft.com/en-us/dynamics365/customer-insights/journeys/contactability-control-custom-forms)
and [singleComponent tab API](https://learn.microsoft.com/en-us/power-apps/developer/model-driven-apps/clientapi/reference/formcontext-ui-tabs/setcontenttype).

## Validation

- `npm run test:grid -- --layout` and the full `npm run test:local` pass.
- Layout cases cover available heights 720, 900, 576 and 240; width at 100%; a
  window resize to 810; native allocation 850 taking precedence over fixed 520;
  ordinary fixed-height fallback 520; invalid/missing sizes; retained real record
  selection; observer/queued-frame cleanup; zero sort/filter/write callbacks.
- Existing editing, menus, filters, metadata, hierarchy, 10,000-row paging and
  42,000-record local fixture checks pass in the same integrated run.
- Release PCF compilation, TypeScript and ESLint pass. Three existing Webpack size
  recommendations remain (732 KiB bundle). Dependencies and lockfile are unchanged.
- A regression reproduced the actual 1.0.120 sandbox host: 935 pixels high at
  top 292.441 in an 1185-pixel viewport. It failed before the visible-area fix.
  The corrected suite also covers viewport shrink, selection retention and offscreen grids.
- The authoritative 1.0.121 Release build compiled and packaged once successfully.
  A dedicated output leaf avoids the packager's output-directory exclusion issue.
- Package asserts `Managed=0`, one PCF root, solution 1.0.100.0 and control 1.0.121.
  SHA-256: `B67D12E98333CC65B0A8D507F1767A4E6FC02C5133AF2C18B1BFD1041814D7FE`.
- The previous deployed 1.0.119 ZIP still matches its original hash. The local
  fixture endpoint serves the actual responsive component and the full-tab layout.

Local fixture: [full-tab preview](http://127.0.0.1:8183/?fulltab=1).
It retains the sample commands and grid functions while using the remaining work
area for the first component. It is an in-memory layout fixture, not a native
Dynamics form or proof that the checkbox has been accepted in the live host.

## Sandbox deployment and acceptance

Target: `https://capricornsandbox.crm4.dynamics.com`, organization
`27252def-2a75-f111-b27b-000d3ab8db73`. Fresh full unmanaged rollback and receipts:
CrmPackage `outputs/deployment-backups/drivengrid-full-tab-sandbox-20261116-125948/`.

The dedicated QA form `e1c3c7c8-2d9d-4a11-8011-b1b9517f81f3` now puts its grid first,
uses native `singleComponent` content, and has no fixed height in any of its three
bindings. The original Template Name control is preserved on Template details.
The reversible form test proves that only those changes occur. Published XML and
compiled form JSON match; other Item Template forms, option views and publisher
remain unchanged. The release solution remains unmanaged at 1.0.0.6.

1.0.120 was deployed first and exposed the command-bar offset. The PCF-only
1.0.121 correction reused the verified full rollback and exact installed 1.0.120
package, without further form changes. Both imports succeeded. Targeted resource
publication did not activate the manifest; guarded Publish All followed a repeated
audit with no pending forms, views, apps, charts, sitemaps or resources. Final
1.0.121 publication completed at 2026-11-16T13:56:29Z. All three resources match
the tested package; bundle SHA-256:
`2FF7C20CEC60FA1E637B9F658FE0EEBED6913C7DB00F4EAD94EC4A35E3EB4823`.

Signed-in Edge checks used browser-only controls after the owner's instruction.
All 13 checks in `browser-runtime-evidence.json` pass:

- Grid/footer fit viewport heights 1126, 1000 and 875 with no outer overflow.
  Footer bottoms are 1125.430, 999.434 and 874.434 respectively. Selection survives resizing.
- Parent and child headers open, then close on a second click. Search drafts
  cancel without filtering; the card has 8-pixel corners. No default alphabet bar.
- Scrolling stays inside the parent grid. Child scrolling automatically grows
  from 50 to all 64 records, without manual paging.
- Native Edit opens exactly selected child `DGQA child 64`
  (`3f8257c1-4b8f-4320-b735-0b26050adb77`); return without saving succeeds.
- Inline Unit Cost editing opens and cancels, retaining the original value.
  Template details retains the existing Template Name field and value.
- Final preview is restored to normal browser sizing, first records, no open
  menu and no selection. No business record was changed by these checks.

Screenshot capture timed out; the owner explicitly waived it. The export gate
retains exact version/package/form and hashed runtime-evidence checks, accepting
the recorded waiver instead of an image. Final unmanaged hashes are recorded in
`final-packages.json` after export verification. Final unmanaged PCF SHA-256:
`851696BF1D4A4D69381C4F38BD2D89A242C0D0BBCB355B6C8135ADC5C3809851`.
Final full CapricornRelease20260622 SHA-256:
`9238C0897E8BBCFEB8383F14B94E0875C8D6194902F6E64A264D3A0BC9BF4F44`.
The full release export contains the native QA tab. No production write occurred.

[Open the sandbox QA preview](https://capricornsandbox.crm4.dynamics.com/main.aspx?appid=f29c988c-f0cc-f011-8544-6045bde18fc1&pagetype=entityrecord&etn=cm_itemtemplate&id=79357ab7-61b3-4f6f-bc6b-946fa6a95dae&formid=e1c3c7c8-2d9d-4a11-8011-b1b9517f81f3).
