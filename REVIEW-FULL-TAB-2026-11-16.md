# DrivenGrid full-tab sizing — 16 November 2026

**Prepared locally: control 1.0.120 / unmanaged solution 1.0.99.0. Not deployed.**
The local suite, PCF Release compilation and unmanaged packaging pass. Actual
Dynamics form configuration and responsive browser acceptance remain pending
because the Computer Use connection is unavailable.

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
- Solution-only packaging then completed successfully from those compiled assets.
  A dedicated output leaf avoids excluding source metadata through the packager's
  output-directory filter; no duplicate PCF compilation was performed.
- Package asserts `Managed=0`, one PCF root, solution 1.0.99.0 and control 1.0.120.
  SHA-256: `4052F841EDBCCAB59B5D9D3EFFE48D4AE87A7D345048135613088A54D9FA2588`.
- The previous deployed 1.0.119 ZIP still matches its original hash. The local
  fixture endpoint serves the actual responsive component and the full-tab layout.

Local fixture: [full-tab preview](http://127.0.0.1:8183/?fulltab=1).
It retains the sample commands and grid functions while using the remaining work
area for the first component. It is an in-memory layout fixture, not a native
Dynamics form or proof that the checkbox has been accepted in the live host.

## Remaining work

Read-only verification confirms sandbox 1.0.119 / unmanaged 1.0.98.0. This task
performed no Dataverse import, publication, form edit or business-record write.
The 1.0.119 signed-in smoke check is still pending as described in the earlier
native-presentation review.

Computer Use reported its native connection unavailable at the start of this
work. No further UI automation was attempted. Do not call the full-tab candidate
visually accepted or installed.

When the connection is restored, use a fresh full rollback for the intended form
change and PCF deployment. On the dedicated QA form, place the grid first in its
own one-column tab, enable the native option and clear the fixed height. Do not
change business forms while preparing that test. The previous deployment helper
is locked to 1.0.119; do not use it to bypass the new target/version/form boundaries.

Verify a fresh signed-in form at multiple window sizes: grid/footer within the
available tab, internal scrolling, no unused bottom region or outer overflow,
parent/child menus and selection, native command targeting and editing Cancel.
Preserve the earlier header toggles, card styling, hidden alphabet and infinite
scroll defaults. Export the final unmanaged packages and evidence after acceptance.
