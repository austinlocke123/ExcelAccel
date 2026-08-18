# Named-range and external-link command contracts

Status: **Draft for review**  
Capabilities: CAP-NAME-001, CAP-LINK-001  
Earliest phase: gated after Phase 2  
Approved scope: read-only inventory, search, navigation, and export

## Common boundary

- Commands do not create, rename, delete, repoint, break, or repair anything.
- Inventory is explicit; there is no continuous background scan.
- Coverage is reported by object/reference category.
- Closed external sources are never opened or contacted automatically.
- Results are derived from a stable snapshot and marked stale after relevant
  workbook changes.

## `names.inventory.open`

- Version: 1
- Impact: read_only
- Parameters: workbook, optional worksheet scope, inclusion of hidden/built-in
  names
- Snapshot: workbook- and worksheet-scoped name metadata, visible/hidden/built-in
  classification, target expression, and approved usage-bearing objects
- Plan: normalize name/scope identity; classify broken target, duplicate display
  name across scopes, target kind, and usage coverage
- Execute: populate virtualized local pane with explicit counts and coverage
- Failure: unsupported name expressions remain visible with `unresolved`; they
  are not silently omitted
- Acceptance: AC-NAME-001 through AC-NAME-005

## `names.search`

- Version: 1
- Impact: read_only
- Parameters: local query and filters for scope/status/target/usage
- Execute: deterministic filtering over the completed inventory; no workbook
  rescan for each keystroke
- Acceptance: AC-NAME-006

## `names.navigate_target`

- Parameters: exact inventory item ID
- CanExecute: target resolves to a supported current workbook object and remains
  valid
- Execute: activate/select target with navigation history
- Failure: constants/formulas/external/unresolved names remain inspectable but
  may be non-navigable with a reason
- Acceptance: AC-NAME-007, AC-NAV-005

## `names.navigate_usage`

- Parameters: exact usage record ID
- Initial usage coverage: supported cell formulas, defined-name expressions,
  chart series formulas, data validation formulas, and print areas/titles;
  every category is separately qualified
- Execute: navigate when the usage maps to a selectable Excel object
- Acceptance: AC-NAME-008 through AC-NAME-010

## `names.inventory.export`

- Explicit local CSV/JSON export with manifest, coverage, and redaction choices.
- Default output includes name/scope/status/target category but excludes raw
  formulas and paths unless the user explicitly enables them in preview.
- Acceptance: AC-NAME-011, AC-SEC-004

## `links.inventory.open`

- Version: 1
- Impact: read_only
- Parameters: explicit workbook and approved link categories
- Snapshot categories: formula external references, defined names, chart series,
  queries/connections, validation, and other separately qualified link-bearing
  objects
- Plan: group by normalized source token without contacting the source; count
  usages and classify accessible/closed/broken/unsupported only from local Excel
  metadata and explicit safe checks
- Execute: virtualized pane with per-category coverage and unsupported counts
- Acceptance: AC-LINK-001 through AC-LINK-006

## `links.search`

- Deterministic local filtering by source, status, category, sheet, and supported
  usage type over a completed inventory.
- Acceptance: AC-LINK-007

## `links.navigate_usage`

- Revalidate and navigate to a supported local usage; never opens the external
  source.
- Non-cell objects expose inspectable metadata and a supported object-selection
  path where qualified.
- Acceptance: AC-LINK-008 through AC-LINK-010

## `links.inventory.export`

- Explicit local export with previewed inclusion of paths/source tokens.
- Paths are excluded/redacted by default.
- Acceptance: AC-LINK-011, AC-SEC-004

## Explicit non-commands

The following IDs MUST NOT exist until new requirements and ADRs are accepted:

- `names.rename`
- `names.delete`
- `links.repoint`
- `links.break`

Inventory results may describe these ideas in help text but may not expose
disabled or experimental mutation buttons.
