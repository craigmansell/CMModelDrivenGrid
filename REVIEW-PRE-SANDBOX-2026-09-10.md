# DrivenGrid pre-sandbox recheck — HOLD

10 September 2026. Reviewed product source at commit
`1db17b158c5feeb02f2520910ae2892d04d9a97a`: control **1.0.112**, unmanaged
solution **1.0.91.0**. Product source and versions were not changed during this
review. Only review evidence and opt-in test probes were added.

**Decision: fix the three reproduced defects below before accepting this candidate
for sandbox deployment.** The existing regression suite and release build pass;
the supplementary failure-path checks do not. Passing the original tests is not
a complete feature sign-off or proof of native-grid equivalence.

No Dataverse import, publish, form/view/ribbon change, server record mutation or
production action was performed. The installed sandbox version was not re-read in
this review; its last recorded verification was control 1.0.110 / solution 1.0.89.0.

## Reproduced findings

### P1 — A rejected ordinary filter can leave Loading records stuck

In the standard dataset path (native child selection disabled),
`CMModelDrivenGrid/index.ts` `applyRootFilter` changes the logical/UI filter before
the host accepts it, then catches a synchronous `setFilter` rejection without
restoring the previous filter or signalling completion back to the grid. The grid
has already set its own loading flag. An error-only render with unchanged records
and `itemsLoading=false` does not clear that flag.

The supplemental test uses the real controller/render path and injects a host
rejection when the A alphabet filter is applied. Actual result:

- Loading overlay remains present.
- UI filter is `name LIKE A%`, while the physical dataset filter is still absent.
- No refresh was started; there is no subsequent loading completion to clear it.

The same controller transaction is used by ordinary Excel-menu filters/sorts.
This is a newly reproduced failure case, **not proof that it caused an earlier
sandbox screenshot**. Restore both filter/sort state on rejection and explicitly
release UI loading; test a successful retry and unchanged-result completion too.

Source: `CMModelDrivenGrid/index.ts` lines 119–131 and `CMModelDrivenGrid/Grid.tsx`
lines 1197–1222, 2495–2501.

### P2 — Compact-layout End/range navigation targets undisplayed columns

Keyboard navigation uses the full grid column count, but compact cards only render
the configured first N fields. End therefore targets a cell that does not exist;
the focus retry expires while the range state points beyond the displayed fields.

Reproduced both in the actual browser (`layout=list`, six total columns, three
displayed) and in a three-column/two-displayed-field test. Starting at column zero,
End leaves focus at zero instead of the final displayed field. Restrict navigation
and range copying to displayed card fields, clear pending targets on reflow, and
test Home/End, Shift extension and arrow boundaries in both layouts.

Source: `CMModelDrivenGrid/Grid.tsx` lines 1290–1312, 1953–1954 and 2213–2229.

### P2 — Lookup save misses a column becoming read-only during metadata work

`onUpdateLookupCell` checks the column before awaiting navigation-property/entity
metadata. Immediately before the eventual write it rechecks grid/child read-only
state, but does not recheck the column's current flags.

The supplemental test changes the column to `isReadOnly=true` while resolving
metadata. The mock receives **one update**, where the required safe result is zero.
Revalidate the current editable column immediately before each permitted write.
This concerns the control's editing policy; it does **not** demonstrate a bypass
of Dataverse's server-enforced privileges or field security.

Source: `CMModelDrivenGrid/index.ts` lines 1247–1275.

## Fresh checks that passed

- Full `npm run test:local`: core logic, inline edits, cell/range selection and TSV,
  grid UI, filter semantics/evaluator/editor, child FetchXML scope/escaping/paging,
  filter candidate resolution/rollback, hierarchy command selection, fixture and
  development tooling tests. The ordinary grid suite was repeated after adding
  the opt-in review entry point and still passes.
- Control TypeScript check and focused TypeScript check of the harness, core,
  presentation, cell/grid tests and supplemental probes.
- Authoritative unmanaged Release build, including manifest validation, TypeScript,
  ESLint and solution packaging. Three Webpack bundle-size/performance warnings
  remain; no build or lint errors.
- Mock strict-host paging: 10,000 rows, 200 batches, no missing/duplicate rows or
  overlapping host page calls.
- Fixture enumeration: 2,000 parents, exactly 20 children each, 42,000 total, all
  80 parent batches and all 2,000 child branches checked. This is not a 42,000-row
  Dataverse load benchmark.
- Existing tests cover paging errors/timeouts/duplicates, filter scope and rollback,
  selection cancellation/query changes, lookup cancellation/wrong table/invalid ID,
  and configured lookup restrictions. The three new probes expose additional gaps.
- Runtime-only dependency audit: zero known reported vulnerabilities. Full audit:
  nine development-tool entries (five moderate, four high), through PCF telemetry
  and BrowserSync dependencies. No dependency upgrades were made in this review.
- `git diff --check`.

## Fresh browser checks (local sample data only)

- Initial 2,000-parent fixture: 50 loaded, 30 rendered row nodes, no alert.
- Select All: exactly 1,000 selected/loaded, 30 rendered row nodes at the checked
  viewport, zero child requests, matching command count and footer.
- Expanding Atlas project 001 returns 20 child rows. No child action toolbar;
  selecting Option 002 and the outer Open command identifies `sample-1-child-2`.
- Outer Activate changes that selected sample child to Yes, not the parent.
  Collapsing clears its selection and disables Open selected.
- Excel menu retains search, (Select All), dotted checklist guides, OK/Cancel and
  all ten requested text conditions. Screenshot inspection shows one value-list
  scrollbar within the menu at the checked viewport, with actions kept visible.
- Contains Atlas yields exactly 334 parents. Descending sort retains that filter
  and begins 1999, 1993, 1987, 1981; no stuck overlay on this successful path.
- Compact End defect reproduced as described above. Preview restored to the
  unfiltered, unselected 2,000-parent fixture afterwards.

These are sample commands, not proof that a real Dynamics ribbon command has the
same availability, target, business rules or save lifecycle.

## Rebuilt artifact identity

`CMModelDrivenGridSolution/bin/Release/CMModelDrivenGridSolution.zip`:

- Solution 1.0.91.0; Managed=0; publisher CraigMansell / cm.
- Exactly one custom control, one dataset, 39 direct inputs. No entity, form,
  workflow, security, relationship or other business component added.
- All manifest display/description keys found in the 91-key resource file.
- Bundle: 741,638 bytes; its SHA-256 is unchanged from the earlier candidate:
  `C41841934C71F1D8F72487BB6C361E6369A50B743CFAE29B6E39B19AD308FED3`.
- Rebuilt ZIP SHA-256:
  `1BF50E8FB0305E5383729735477A2BF0F4674146BD8100BD5E4A3314F2E88054`.
  This is the new archive identity; the earlier validation document retains its
  historical archive hash. Neither package is approved by this review for release.

## Running the supplementary release gates

```text
npx tsx scripts/local-grid-regression-test.tsx --pre-sandbox-review
```

The three probes in `scripts/local-pre-sandbox-review-probes.tsx` report all their
results and return exit code 1 while any release gate fails. They use local mock
host calls only. The normal regression command remains separate so its prior pass
status can be compared honestly; do not use that pass to override these failures.

## Still required after fixes

1. Make the three supplemental probes pass, promote their coverage into the normal
   regression suite, and repeat integrated checks/build/artifact validation.
2. Before any sandbox write, create a fresh exact unmanaged rollback and check the
   sandbox organisation, installed versions, component scope and publisher locks.
3. After an authorised deployment, test the candidate on the isolated signed-in QA
   form with normal business handlers enabled. Check parent/child command targeting,
   save/readback, lookup dialog, refresh/filter failures and actual server paging.
4. Capture the final unmanaged package and runtime readback. Do not claim full
   native parity: independent/cross-table child datasets, specialist editors,
   label-pattern lookup filters, native events/customizers, DateTime user-timezone
   semantics and broader device/accessibility/security acceptance remain gaps.

Microsoft documents a 1,000-row multi-selection limit and the native grid's
feature/configuration surface in [Power Apps grid control](https://learn.microsoft.com/en-us/power-apps/maker/model-driven-apps/the-power-apps-grid-control).
DrivenGrid's Select All is an intentional extension, not a claim of identical
header behaviour: that page currently says Select All is unavailable with native
infinite scroll. The supported PCF lookup dialog has no `filters` option, explaining
the need for separate result validation in DrivenGrid; see
[PCF lookupObjects](https://learn.microsoft.com/en-us/power-apps/developer/component-framework/reference/utility/lookupobjects).
