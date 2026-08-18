# Selection, row/column, and broad-format command contracts

Status: **Draft for review**  
Capabilities: CAP-SELECT-001, CAP-STRUCT-001, CAP-FMT-003  
Earliest phase: selection tools in 1B/2 by parser dependency; structural and
broad-format commands individually gated

## 1. Deterministic selection tools

Selection commands change Excel selection/focus only, not workbook content.
Initial scope is the current selection or the active worksheet's explicitly
confirmed used range. Cross-sheet matches use a results pane; ExcelAccel does not
attempt one cross-sheet multi-area selection.

Common parameters:

- explicit scope;
- whether hidden/filtered cells are included;
- maximum resulting cells/areas;
- parser coverage where formulas are inspected.

Common behavior:

- build a stable match set from one snapshot;
- preview when area/cell caps are exceeded;
- refuse rather than create an unstable/excessive multi-area selection;
- preserve workbook content and push the prior location to navigation history.

| Command ID | Exact predicate | Acceptance |
|---|---|---|
| `selection.select.formulas` | Cells whose underlying qualified formula property contains a formula | AC-SELECT-001..004 |
| `selection.select.constants` | Nonblank cells containing constants, excluding formulas | AC-SELECT-001..004 |
| `selection.select.blanks` | Truly empty cells under the approved blank predicate; formulas returning empty text are not blank | AC-SELECT-001..005 |
| `selection.select.errors` | Cells containing qualified Excel error values | AC-SELECT-001..004 |
| `selection.select.numeric_hardcodes` | Numeric constants in cells, excluding formula literals and numeric text unless separately enabled | AC-SELECT-001..004, AC-SELECT-006 |
| `selection.select.external_formulas` | Formula cells containing a supported external-workbook reference | AC-SELECT-001..004, AC-SELECT-007 |

## 2. Row/column visibility and outline commands

### `rows.hide` and `columns.hide`

- Version: 1
- Impact: medium
- Parameters: exact complete rows/columns resolved from selection
- Changed properties: hidden state only
- Preview: threshold-based and mandatory when hidden/filtered/table/outline
  interactions are present
- Undo: qualified property receipt
- Acceptance: AC-STRUCT-001 through AC-STRUCT-005

### `rows.unhide` and `columns.unhide`

- Same contract; changes hidden state only for exact explicit rows/columns.
- It does not infer surrounding hidden ranges beyond the selected bounds.
- Acceptance: AC-STRUCT-001 through AC-STRUCT-005

### `rows.group`, `columns.group`, `rows.ungroup`, `columns.ungroup`

- Version: 1
- Impact: medium
- Parameters: exact complete rows/columns and outline action
- Plan: current outline levels, proposed levels, existing/overlapping groups,
  affected bounds, and undo eligibility
- Changed properties: outline/group state only
- Preview: mandatory for overlaps or multi-level changes
- Failure: unsupported/protected/partial-outline contexts refuse the whole plan
- Acceptance: AC-STRUCT-006 through AC-STRUCT-010

### `rows.smart_hide` and `columns.smart_hide`

- Exact behavior: create/extend the explicit outline group under the approved
  overlap policy, then collapse it. It does not merely set `Hidden = true`.
- Preview: mandatory and shows group bounds/levels plus collapse state.
- Acceptance: AC-STRUCT-006 through AC-STRUCT-011

## 3. Structural insert/delete commands

These contracts are documented but explicitly deferred. They cannot use the
Phase 1A property-only undo mechanism.

### `rows.insert` and `columns.insert`

- Version: 1
- Impact: high
- Parameters: exact insertion bounds/count and approved Excel shift/copy-origin
  behavior
- Plan: affected sheets/ranges/formulas/names/tables/charts/spills/arrays,
  formatting origin, protection, file/coauthoring state, and rollback mechanism
- Preview: mandatory
- Execute: unavailable until a dedicated structural transaction/rollback ADR
  proves supported behavior across the full matrix
- Acceptance before registry release: AC-STRUCT-012 through AC-STRUCT-017

### `rows.delete` and `columns.delete`

- Version: 1
- Impact: high/destructive
- Plan must inventory removed values, formulas, formats, names, tables, objects,
  validation, comments, outline, and reference effects within declared coverage.
- Preview must identify data/formula-bearing targets and permanent-risk cases.
- Execute remains unavailable until exact rollback and post-delete reference
  behavior are proven; refusal is the default.
- Acceptance before registry release: AC-STRUCT-012 through AC-STRUCT-019

## 4. Broad formatting commands

### `format.sheet_style.apply`

- Version: 1
- Impact: high
- Parameters: one explicit worksheet and versioned enabled style components
- Changed properties: only listed cell/row/column/view formatting properties;
  values/formulas/names/VBA/comments/validation/hyperlinks are excluded
- Plan: per-component targets/counts/skips, samples, performance cost, and
  receipt eligibility
- Preview: mandatory
- Execute: component-batched exact plan with verification
- Undo: every enabled component requires a qualified bounded receipt; otherwise
  the command refuses or the component cannot be enabled
- Acceptance: AC-FMT-014 through AC-FMT-018

### `format.workbook_autoformat.apply`

- Version: 1
- Impact: high
- Parameters: explicit included worksheets and enabled versioned components
- Plan: complete sheet/component inventory and aggregate receipt/resource cost;
  no hidden sheet or workbook-wide expansion unless explicitly selected
- Preview: mandatory
- Execute: gated until workbook-scale rollback, AutoSave/coauthoring refusal, and
  performance evidence pass
- Acceptance: AC-FMT-014 through AC-FMT-020

### `format.auto_color.workbook`

- Version: 1
- Impact: high
- Parameters: explicit included sheets and profile classification recipe
- Changed properties: font color only
- Plan: per-sheet counts, unsupported cells, samples, workload, and undo caps
- Preview: mandatory
- Execute: gated until workbook-scale snapshot/write/rollback evidence passes;
  partial success is not an approved default
- Acceptance: AC-FMT-010..013, AC-FMT-019, AC-FMT-020

## 5. Explicitly unspecified view/print commands

The original draft referred generically to print/view settings. Only gridlines,
zoom, and freeze/unfreeze panes currently have approved detailed contracts.
Agents MUST NOT invent additional page setup, print area, view mode, heading,
formula-bar, or window commands. Each requires its own property-level contract.
