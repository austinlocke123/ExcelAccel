# Discovery, favorites, styles, and profile command contracts

Status: **Draft for review**  
Capabilities: CAP-SEARCH-001, CAP-FAV-001, CAP-STYLE-001, CAP-PROF-002  
Earliest phase: 1B

Implementation note (2026-08-19): WP-1B-01 search/favorites, WP-1B-02 style
recipes, and WP-1B-03 offline profile exchange are implemented and locally
qualified within their retained limits.

## Common constraints

- Search and favorites operate on registry metadata and local settings only.
- Style commands change formatting only; they never transfer values, formulas,
  notes, comments, validation, hyperlinks, names, or workbook structure.
- Profile packages contain declarative JSON and approved style assets only.
- Import is previewed and atomic. Invalid input changes nothing.
- No command in this file requires network access.

## `command.search.open`

- Version: 1
- Impact: read_only
- Parameters: optional initial query
- Supported context: Excel active; workbook optional
- Snapshot: immutable registry metadata and current `CanExecute` summaries; no
  workbook scan
- Plan: normalize query using invariant search rules, rank name/alias/category/
  description/shortcut matches deterministically, and cap results
- Preview: none
- Execute: open/focus the local search surface; selecting a result invokes that
  command through the normal command boundary
- Failure: registry unavailable -> `Refused`; individual command eligibility is
  shown without executing it
- Performance: first visible results within frozen `PERF-002`; UI remains
  responsive with the full registry
- Acceptance: AC-SEARCH-001 through AC-SEARCH-004

## `favorite.add`

- Version: 1
- Impact: low local-settings mutation
- Parameters: command ID and optional typed fixed arguments
- Supported context: registered command whose arguments are serializable and do
  not contain workbook-derived content
- Changed properties: user profile favorites only
- Plan: validate command/version/arguments and proposed position
- Preview: none
- Execute: atomic profile update; duplicate adds are idempotent
- Undo: none; explicit remove is available
- Failure: invalid/incompatible command leaves profile unchanged
- Acceptance: AC-FAV-001, AC-FAV-003, AC-PROF-002

## `favorite.remove`

- Same boundary as `favorite.add`.
- Removes only the selected favorite entry; does not alter command bindings or
  definitions.
- Removing an absent entry succeeds with no change.
- Acceptance: AC-FAV-001, AC-FAV-003

## `favorite.invoke`

- Version: 1
- Impact: inherited from referenced command
- Parameters: favorite ID
- CanExecute: resolve current command version and arguments, then call the
  referenced command's `CanExecute`
- Execute: route through the referenced command lifecycle; never call a feature
  implementation directly
- Failure: incompatible/missing command is visible and non-mutating
- Acceptance: AC-FAV-002 through AC-FAV-004, AC-CMD-002

## `style.capture`

- Version: 1
- Impact: low local-settings mutation; workbook read-only
- Parameters: style name, selected property set, overwrite flag
- Supported context: exactly one source cell in a supported workbook
- Snapshot: only selected formatting properties from the source cell
- Plan: create a normalized style recipe with explicit properties and invariant
  serialization; exclude all workbook content and unsupported theme/object data
- Preview: list captured properties and whether a local style will be replaced
- Execute: atomically save the recipe to the user profile
- Undo: none; style can be deleted or overwritten explicitly
- Failure: any unreadable/unsupported selected property refuses the capture
  rather than producing a partial style
- Acceptance: AC-STYLE-001 through AC-STYLE-003, AC-PROF-003

## `style.apply`

- Version: 1
- Impact: low or medium by target size
- Parameters: stable local style ID
- Supported context: cell/range selection supported by every property in the
  style recipe
- Changed properties: exactly the recipe's declared formatting properties
- Snapshot: target before-state for declared properties only
- Plan: property-level deltas, changed/skipped counts, warnings, receipt size
- Preview: threshold-based; mandatory when target/receipt cap is exceeded or
  when a recipe contains row/column dimensions
- Execute: block/group writes by property; verify postconditions
- Undo: session receipt when within qualified caps
- Failure: unsupported property/target combination is planned as a skip only if
  the recipe declares skip-safe behavior; otherwise refuse before mutation
- Acceptance: AC-STYLE-004 through AC-STYLE-007, AC-REL-005

## `style.apply_builtin`

- Parameters: one of the versioned built-in style IDs such as `major_header`,
  `minor_header`, `date_header`, `assumption`, `formula`, `linked_formula`,
  `output`, `warning`, or `total`
- All other behavior is identical to `style.apply`.
- Built-ins are profile data, not separate implementations.
- Acceptance: AC-STYLE-004 through AC-STYLE-007

## `style.delete_local`

- Impact: low local-settings mutation
- Parameters: local style ID
- CanExecute: built-in or organization-locked styles cannot be deleted
- Execute: atomic profile update; existing workbooks are not scanned or changed
- Acceptance: AC-STYLE-008, AC-PROF-002

## `profile.export`

- Version: 1
- Impact: local file write; workbook read-only
- Parameters: explicit destination, included local components
- Plan: manifest with schema/product compatibility, counts, approved asset list,
  and destination; reject existing destination unless overwrite is explicit
- Preview: mandatory manifest and destination
- Execute: write to a temporary file, validate/checksum, then atomic rename;
  never launch or transmit the package
- Failure: remove incomplete temporary output where safe; preserve existing file
- Acceptance: AC-PROF-003 through AC-PROF-006

## `profile.import.preview`

- Version: 1
- Impact: read_only
- Parameters: explicit local package path
- Snapshot: current effective profile and policy locks
- Plan: validate size, archive paths, manifest, checksum, schema, versions,
  command/style IDs, migrations, locked conflicts, additions, and replacements
- Preview: returned as a deterministic structured diff; no settings write
- Failure: invalid/untrusted content produces categorized refusal
- Acceptance: AC-PROF-002, AC-PROF-004, AC-PROF-007

## `profile.import.apply`

- Version: 1
- Impact: medium local-settings mutation
- Parameters: exact import-plan hash
- CanExecute: revalidate package identity, current profile version, and policy
  locks against the preview
- Execute: create bounded local backup, write/migrate to a temporary file,
  validate effective settings, atomically replace, and refresh registry bindings
- Undo: restore pre-import profile through an explicit one-session receipt
- Failure: prior profile remains active; no partially migrated settings
- Acceptance: AC-PROF-002, AC-PROF-004, AC-PROF-007 through AC-PROF-009

## `bindings.cheat_sheet.export`

- Version: 1
- Impact: local file write
- Parameters: destination and approved format (`html` or `csv` initially)
- Plan: active effective bindings, conflicts, command names/categories, and
  destination; no workbook data
- Execute: deterministic local export via temp-and-replace
- Acceptance: AC-PROF-005, AC-KEY-004

Competitor-specific keybinding import is not approved by these contracts.
