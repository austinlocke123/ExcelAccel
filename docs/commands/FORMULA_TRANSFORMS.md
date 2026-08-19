# Formula, paste, and fill command contracts

Status: **Approved contract; implementation qualification in progress**
Capability: CAP-FORM-001
Earliest phase: 1B, after ADR-0004 is accepted

Implementation note (2026-08-19): ADR-0004 now accepts a narrow A1-only
transformation boundary. Inspect-only syntax and R1C1 mutation remain
fail-closed. Formula-edit reference toggle remains host-disabled pending a safe
caret/edit-text API; its pure transformation contract is implemented and tested.
Smart Copy Down/Right, IFERROR, sign, and unit commands now use a transactional
typed-matrix adapter with exact revalidation, verification, compensation, and
bounded optimistic undo. See `docs/evidence/WP-1B-05_07_FORMULA_COMMANDS.md`.
Spacing, transpose, formula/value-from-above, and typed numeric/date fill are
registered with explicit parameter dialogs where required. Operations with a
separate source range use qualified selection-preserving off-selection
revalidation, including calculated-value fingerprints for value-only commands.
Transpose and formulas/values-only paste use the bounded internal source;
see `docs/evidence/WP-1B-06_08_ADVANCED_FORMULA_PLANNERS.md`.

## Common formula boundary

- Each command operates only on formula syntax marked `transform` in the
  accepted parser coverage matrix.
- Unsupported or partially parsed formulas refuse before mutation.
- Plans record original and proposed formula text, normalized syntax identity,
  target, locale/dialect, spill/array status, and exact changed references.
- Formula writes use the qualified `Formula`/`Formula2` policy.
- Legacy arrays, dynamic arrays, data tables, table calculated columns, shared
  formulas, and multi-cell spill owners have explicit command-specific rules.
- Raw string replacement is prohibited.
- Formula-transform preview is mandatory above the qualified target threshold
  and always shows representative before/after formulas.

## `formula.reference.toggle`

- Version: 1
- Impact: low
- Parameters: none; acts on the reference containing the caret
- Supported context: Excel formula edit mode only when the host spike proves a
  safe, exact caret/reference boundary API; otherwise this command remains
  disabled
- Changed properties: active edited formula text only
- Plan: cycle `A1 -> $A$1 -> A$1 -> $A1 -> A1` while preserving sheet/workbook
  qualifiers and range structure according to the accepted behavior matrix
- Preview: none; proposed formula must be known before replacing edit text
- Undo: native edit cancellation/undo only; add-in session receipt is not used
- Failure: ambiguous caret, structured reference, name, or unsupported syntax
  refuses without changing edit text or keyboard state
- Acceptance: AC-FORM-005 through AC-FORM-007, AC-KEY-001

## `formula.copy.down` and `formula.copy.right`

- Version: 1
- Impact: medium
- Parameters: source edge/row/column and explicit destination inferred only from
  the validated selection shape
- Supported context: rectangular cell selection with a qualified formula source
  row/column and no destructive overlap ambiguity
- Changed properties: formulas in destination cells only
- Plan: translate references using Excel-equivalent relative/absolute semantics;
  list source-to-target mapping and overwritten destination formulas/values
- Preview: mandatory when destination is nonblank or exceeds threshold
- Undo: session receipt only after formula-value receipt qualification; otherwise
  command cannot ship
- Failure: protected, array/spill, table, mixed unsupported syntax, or unsafe
  overlap refuses the whole plan
- Acceptance: AC-FORM-008 through AC-FORM-012

## `formula.spacing.rows` and `formula.spacing.columns`

- Version: 1
- Impact: medium
- Parameters: positive interval, source formula cell/range, explicit destination
- Plan: map source formulas to each destination position and translate parsed
  references by the exact row/column displacement
- Changed properties: formulas only; pre-existing nonblank destinations require
  mandatory preview and explicit overwrite confirmation
- Failure/undo: same as Smart Copy
- Acceptance: AC-FORM-008 through AC-FORM-012

## `formula.transpose`

- Version: 1
- Impact: medium
- Parameters: source range and explicit top-left destination
- Supported context: rectangular formula/value source within qualified size;
  initial contract supports formulas and constants only, not formatting
- Plan: transpose positions and transform each parsed relative reference as if
  the formula pattern were transposed; do not merely call textual transpose
- Preview: mandatory with complete destination/overwrite summary and samples
- Execute: write constants/formulas in a bounded block after revalidation
- Undo: qualified session receipt required
- Failure: any unsupported formula or destination conflict refuses the entire
  plan; no partial transpose
- Acceptance: AC-FORM-013 through AC-FORM-016

## `formula.iferror.toggle`

- Version: 1
- Impact: medium
- Parameters: profile-defined fallback expression/value
- Supported context: selected formula cells whose top-level syntax is qualified
- Plan: if the exact configured wrapper is present at the top level, remove one
  layer; otherwise wrap once. Never remove an unrelated IFERROR or double-wrap
  an equivalent configured wrapper
- Changed properties: formulas only
- Preview: threshold-based with before/after samples
- Undo: qualified session receipt required
- Failure: non-formulas are skipped and counted; unsupported formulas cause
  refusal according to the all-or-refuse parser policy
- Acceptance: AC-FORM-017 through AC-FORM-020

## `formula.sign.reverse`

- Version: 1
- Impact: medium
- Parameters: none
- Plan: apply/remove one canonical negation while preserving precedence and
  avoiding repeated wrapper growth; constants are handled only if the plan
  explicitly lists them
- Changed properties: formulas or numeric values according to selected target
  categories shown in preview; text/blanks/errors are skipped
- Preview: threshold-based; always state whether constants are included
- Undo: qualified session receipt required
- Acceptance: AC-FORM-021 through AC-FORM-024

## Unit-scale commands

| Command ID | Exact operation |
|---|---|
| `formula.units.to_thousands` | formula/value divided by 1,000 |
| `formula.units.from_thousands` | formula/value multiplied by 1,000 |
| `formula.units.to_millions` | formula/value divided by 1,000,000 |
| `formula.units.from_millions` | formula/value multiplied by 1,000,000 |

- Version: 1
- Impact: medium
- Plan: parsed wrapper preserving precedence; preview names the operator and
  scale and shows formula/value samples
- Constants are included only when explicitly selected in typed parameters.
- Reapplying is not treated as an automatic toggle; it applies the named
  operation again.
- Acceptance: AC-FORM-021 through AC-FORM-025

## `fill.formula_from_above`

- Version: 1
- Impact: medium
- Parameters: explicit destination selection; source is the immediately adjacent
  cell above each destination column
- CanExecute: each source contains a supported formula and every destination is
  within one rectangular qualified selection
- Plan: translate each source formula downward using the exact relative/
  absolute reference rules; list nonblank destination overwrites
- Preview: mandatory for nonblank destinations and above threshold
- Changed properties: destination formulas only
- Undo: qualified formula receipt required
- Acceptance: AC-FORM-035 through AC-FORM-038

## `fill.value_from_above`

- Version: 1
- Impact: medium
- Parameters: explicit destination selection; source is the immediately adjacent
  cell above each destination column
- Plan: copy the source's underlying current value, not its formula or displayed
  formatted text; formula sources therefore produce a value-only destination
- Preview: mandatory and explicitly identifies formula-source-to-value targets
- Changed properties: destination values only; no source or formatting changes
- This is an explicitly permitted formula-to-value operation at the destination.
- Undo: qualified value receipt required
- Acceptance: AC-FORM-035 through AC-FORM-038

## Paste commands

### `paste.formulas_only`

- Source: explicit internal clipboard snapshot captured by ExcelAccel or a
  qualified Excel clipboard path; the contract chosen in Phase 1B must be one
  implementation, not two inconsistent behaviors
- Changed properties: formulas only
- Values/formats/validation/comments/hyperlinks are not transferred
- Destination shape/repetition rules are explicit and previewed on mismatch
- Acceptance: AC-FORM-026 through AC-FORM-029

### `paste.values_only`

- Changed properties: cell values only
- This is an explicitly permitted formula-to-value operation at the destination.
- Preview is mandatory when destination formulas will be replaced.
- The plan lists formula-to-value target count and representative examples.
- Acceptance: AC-FORM-026 through AC-FORM-030

### `paste.formats_only`

- Changed properties: approved formatting property set only
- No values, formulas, validation, comments, hyperlinks, or dimensions unless a
  typed parameter explicitly includes row/column dimensions
- Acceptance: AC-FORM-026 through AC-FORM-029, AC-REL-005

Released v1 property set: number format; font name, size, bold, italic, and
underline; horizontal and vertical alignment; and indent level. Source and
destination are each capped at 100 cells and must follow the same exact-shape or
whole-multiple, nonoverlap rule as other internal paste commands. Font color,
fill, borders, validation, comments, hyperlinks, row/column dimensions, values,
and formulas are explicitly excluded and preserved. Capture, plan, execution,
verification, rollback, and optimistic undo use a serialized per-cell property
matrix; no Office or Windows clipboard is used.

## `fill.numeric_sequence` and `fill.date_sequence`

- Version: 1
- Impact: medium
- Parameters: explicit start, step, direction, and count/destination
- Plan: compute every output deterministically using invariant numeric/date
  arithmetic and the selected locale display policy
- Dates require an unambiguous typed date value; free-text date guessing is
  prohibited
- Preview: mandatory for nonblank overwrite and above threshold
- Acceptance: AC-FORM-031 through AC-FORM-034, AC-LOC-001

## Shared failure behavior

After mutation begins, an unexpected write or verification failure stops further
writes and performs the command's qualified rollback. The result reports exact
completed, rolled-back, and remaining targets. No formula command may report
success based only on the absence of a COM exception.
