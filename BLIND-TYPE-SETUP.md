# Blind type setup

Place the grid inside the existing Item **Options** tab. Preserve the Item main
form, its other tabs, its existing Options controls and its handlers. The isolated
sandbox preview copies that main-form layout and adds one setup section to Options;
it does not replace the existing Item main form.

Native release status: published in sandbox; signed-in native acceptance remains
pending. CRM Hub registration and the documented API ribbon rebuild succeeded.
Browsers show the complete Item layout and control 1.0.7. Adapter v8 and command
revision 9 are published; the rebuild succeeded at 17:23:50 UTC. Edge still loads
form adapter v7/ribbon adapter v4 after reload, while the in-app browser previously
retained v5 rules. Sandbox site-data clearing is pending. The startup record
callback remains unverified. A rejected bound-column
trial was cleaned up; the clean V2 dataset uses public events/outputs, with no
bound signal property. The existing standalone 1.0.0.1 bookmark has
not been redirected yet.

The native editor uses the existing ResponsiveGrid and HierarchyGrid inside a
Dynamics dataset control. It preserves hierarchy arrows, repeated header clicks,
child filtering/sorting, retained expansions and infinite scrolling. The installed
general grid remains 1.0.122; the separate setup control candidate is V2 1.0.7.

## Editing in the native subgrid

- In **Options**, with nothing selected, choose the native **New Blind type setup
  row** command to add an option. Enter a unique name, choose a component type,
  select the components available for this blind type and an optional default,
  then **Save**. The form is inline and saves straight to Dynamics.
- Select an option and choose native **New** to insert a component line beneath
  it. Type an existing component code or description, choose an available match,
  optionally tick **Default**, and press **Save** or Enter. Cancel removes only
  the unsaved entry. Components already selected or outside this blind type/type
  are excluded. The command hides when no eligible choices remain.
- Select an option/component and use native **Edit**. Options use a non-modal
  panel; components edit in their own row. **Delete** appears only when the
  selected entry can be removed without breaking references. Catalogue options
  are retained and use **Activate/Deactivate** where the rule dependencies allow.
- In **Linked options**, native **New** creates a target when nothing is selected,
  adds a condition for a selected linked option, or inserts a result line for a
  selected condition/result. Edit and Delete follow the selected hierarchy level.
- Each entry's **Save** persists immediately. Save configuration, Undo and Reload
  are removed. The native **Refresh** command reloads confirmed state; commands
  are hidden during an open entry or save. Failed saves retain the entry.

Available components come from the catalogue stored on the selected Item and its
component type. Defaults must be selected available components. Removing a choice
used by conditions/results is blocked until references are changed. Validation
also prevents dependency cycles, invalid size ranges and duplicate results.

## Native host and data contract

The isolated Item preview form is `9f4ad010-6a1b-4a90-bf36-238d93d10001`, with
`Subgrid_BlindTypeSetup` related to real `cm_blindtypesetuprow` records. Six existing
Dynamics buttons are reused with context-sensitive rules and handlers. No second
PCF toolbar is rendered. Selection resolves to actual persisted Dataverse row IDs,
including nested children, through public dataset APIs.

The form sends its actual Item ID through Microsoft's documented PCF custom-event
callback. Dataset queries are explicitly scoped to that Item. Commands remain
hidden before the handshake, during native selection preparation and when editing
is disabled. The preview is available in CRM Hub and is not the default Item form.
Its Options-only main-form candidate has an exact recovery check: removing only
the added section, control description, library and handler recovers the original
complete Item main form XML. Main-form activation remains pending native acceptance.

`cm_item.cm_blindtypecataloguejson` stores the administrator-loaded catalogue;
`cm_item.cm_blindtypesetupjson` stores the versioned setup. A synchronous server
plug-in validates writes and projects the saved setup into stable related records
in the same transaction. Raw writes to projection rows are blocked. Setup edits
use the signed-in user's ordinary Item/column and native-row permissions, current
catalogue validation and If-Match concurrency. Setup users need read/create/write/
delete plus append permissions on the new table; this release does not grant roles
or elevate the caller. Native command privilege checks still apply.

Existing documents remain format 1. Custom options use format 2 with stable IDs,
names and immutable catalogue type references; removing the final custom option
returns to format 1. Preserve later configuration data before rollback. Unmanaged
solution imports do not delete newly added schema or restore row data.

The existing Item Option/Item Option Selection tables and quotation calculation
integration are unchanged. Catalogue component IDs retain their Blindata meaning;
they are not fabricated Dataverse component lookup IDs. Formula text is preserved
and never evaluated in the browser. Blindata and SQL remain read only.

The R200S example contains 15 options, 82 component choices, four linked groups,
32 conditions and 68 results (201 native rows). Two historical unavailable NONE-0
results remain labelled and cannot be selected for a new result.

## Build, preview and sandbox

- Local preview: `http://127.0.0.1:8183/?demo=r200s&setup=1`. Its isolated memory
  store never writes to Dynamics. The standalone toolbar is for local authoring.
- Native build: `scripts/build-native-blindtype-control.ps1`, using existing
  dependencies and a separate build directory. Package is unmanaged
  `CMBlindTypeSetup` 1.1.0.0 and includes only the dedicated setup PCF.
- Checks: `npm run test:local`; the native build runs authoritative lint/compile.
  Direct-save probes cover retained failures, concurrent/double saves, scoped
  inline entries, cancellation and native command capabilities.
- Deployment: CrmPackage `BlindTypeSetup/Deploy-NativeBlindTypeSetupSandbox.ps1`.
  Exact rollback/evidence: `outputs/deployment-backups/blindtype-native-subgrid-sandbox-20261117-113020/`.
  Schema/form changes receive full unmanaged rollback snapshots; refinements keep
  the original immutable rollback and separate exact candidate/receipt artifacts.
- Sandbox demo Item: `25a03e63-9b28-472f-8bd6-71586867a157`. Native form URL:
  `https://capricornsandbox.crm4.dynamics.com/main.aspx?appid=f29c988c-f0cc-f011-8544-6045bde18fc1&pagetype=entityrecord&etn=cm_item&id=25a03e63-9b28-472f-8bd6-71586867a157&formid=9f4ad010-6a1b-4a90-bf36-238d93d10001`.

## Accepted Add option update

Unmanaged `CMBlindTypeSetup` 1.0.0.1 updates the existing three resources and
solution version. The full local suite, scoped TypeScript and resource build
pass. Signed-in creation, scoped components/default, reload persistence, custom
linked-option/condition/result authoring, arrows and all nested header toggles
pass. Independent readback verifies preservation of the previous configuration.
Temporary test entries are removed. No generic PCF, schema or existing form/view
changes. Exact rollback and acceptance evidence are in CrmPackage
`outputs/deployment-backups/blindtype-add-option-sandbox-20261117-103516/`.
Screenshot capture timed out; signed-in AX/DOM and independent readback establish
functional acceptance. The source and deployment branches are
`codex/blind-type-add-option-20261117` in their respective repositories.

Final unmanaged 1.0.0.1 SHA-256:
`B9159867F9D764301E25B87A47C54C55B1F899BE91A279DD16AFA85FD2C9F550`.

## Initial sandbox release

Unmanaged solution `CMBlindTypeSetup` 1.0.0.0 is installed. All eight signed-in
checks pass, including real condition/result CRUD, default persistence after
reload, scope-restricted component lists, three hierarchy/header levels and a
stale save rejected after a separate same-data row-version change. The original
R200S record and existing forms/views/PCF remain unchanged. Temporary authoring
test entries were removed; the demo retains the original rule counts and the
white System 40 mechanism as its example default.

Final unmanaged package SHA-256:
`A55219DD8699BE5E251AA8CE27F1EC679DE817D7DB9EC3316D59A3ED17A7DE62`.
Rollback and exact receipts are in CrmPackage
`outputs/deployment-backups/blindtype-setup-sandbox-20261117-093826/`.
Browser screenshot capture timed out; acceptance uses hashed signed-in AX/DOM
evidence and independent Dataverse readback, without a pixel-rendering claim.

Microsoft references: [HTML web resource record parameters](https://learn.microsoft.com/en-us/power-apps/developer/model-driven-apps/webpage-html-web-resources),
[Dataverse metadata and solution context](https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/create-update-entity-definitions-using-web-api).
