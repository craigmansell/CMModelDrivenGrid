# Blind type setup

The editor uses the existing ResponsiveGrid and HierarchyGrid, including the
Power Apps styling, hierarchy arrows, repeat-click column menus, independent
child filtering/sorting and retained expansions. No generic PCF source or
manifest changes are needed.

## Editing

- **Options:** expand a group to inspect its components. Select a group or
  component and choose **Edit selected** to set its allowed components, default,
  order, enabled/visible status and free-text setting. **Set as default** is a
  shortcut for a selected component.
- **Linked options:** add or edit a target option, then select it and use
  **Add condition**. Set AND/OR choice tests, width/drop limits and optional
  chain length, panel count, area and formula values. Select a condition and
  choose **Add result**, or a result and **Choose component**, to edit its
  component and order.
- Apply each editor to the draft, then **Save configuration** to save the whole
  setup in Dynamics. Undo restores previous draft changes. Reload asks before
  discarding an unsaved draft. Server and permission errors retain the draft.

The available component list is specific to the catalogue on the selected Item
and its option group. Removing a component used by a condition/result is blocked
until those references are changed. Conditions cannot create dependency cycles,
invalid size ranges, duplicate results or references to disabled source options.

## Data and scope

Dedicated Item fields `cm_blindtypecataloguejson` and `cm_blindtypesetupjson` hold
versioned catalogue and configuration documents. The catalogue is administrator
loaded; the editor changes only the configuration field. Data access uses the
signed-in user's ordinary Dataverse permissions, current-catalogue validation,
and `If-Match` optimistic concurrency. A stale save cannot replace another user's
version. Validation is in the editor/save adapter; no server validation plug-in
is registered for clients that write the JSON column directly.

This release stores blind-type setup. It does not change the existing Item
Option/Item Option Selection tables or automatically apply these rules to quote
calculations. Component IDs in the catalogue are the source Blindata choice IDs;
they are scoped within the versioned catalogue rather than invented Dataverse
lookup references. Formula text is preserved and never evaluated in the browser.

The R200S seed has 15 option groups, 82 component choices, four linked groups,
32 conditions and 68 result links. Two existing `NONE-0` results are absent from
their direct allowed pool. They are retained as labelled historical references
and cannot be selected for a new result. No Blindata or SQL writes are performed.

## Preview and deployment

- Local authoring preview: `http://127.0.0.1:8183/?demo=r200s&setup=1`.
  Its Save demo action uses an isolated memory store, with no Dynamics write.
- Production resource build: `npm run build:blindtype-setup`.
- Tests: `npm run test:local` and the scoped TypeScript check recorded in the QA
  evidence. The new store/model tests cover scoping, defaults, cycles, invalid
  references, legacy result preservation and conditional save payloads.
- Sandbox deployment: CrmPackage
  `BlindTypeSetup/Deploy-BlindTypeSetupSandbox.ps1` and matching
  `CHANGELOG-CMBlindTypeSetup.md`. A fresh exact rollback, target/publisher locks,
  unchanged form/view/PCF audits and signed-in save/reload acceptance are required.
- The dedicated sandbox demo Item is `25a03e63-9b28-472f-8bd6-71586867a157`.
  Open `/WebResources/cm_blindtypesetup.html?data=<Item GUID>` in the target
  Dynamics environment. This can also be embedded as a web resource in an Item
  form with the supported record-context parameter; existing forms are untouched.

## Accepted sandbox release

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
