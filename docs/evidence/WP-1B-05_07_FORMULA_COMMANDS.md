# WP-1B-05 and WP-1B-07 transactional formula commands

Date: 2026-08-19

## Implemented

- Smart Copy Down and Smart Copy Right over one bounded unmerged rectangle.
- Exact top-edge/left-edge formula ownership and A1 mixed-anchor translation.
- IFERROR toggle using the profile-defined qualified fallback expression.
- Canonical formula sign reversal.
- Explicit to/from thousands and millions operations. Ribbon routes affect
  formulas only; the typed application contract supports opt-in numeric
  constants and identifies formula/value changes separately.
- Dedicated Formula Ribbon menu and searchable registry metadata with unique
  KeyTips under `Alt, X, A, M`.

WP-1B-04's pure caret transform is qualified, but the live edit-mode command is
still intentionally not registered because no safe exact caret/edit-text host
API has passed qualification.

## Transaction boundary

The selected rectangle is captured as an immutable, typed matrix containing
blank, formula, text, finite number, or Boolean cells. Constant error values and
unknown COM values refuse. The matrix has a 10,000-cell and 1,000,000-character
serialized ceiling.

Each command:

1. captures and safety-checks the full rectangle;
2. creates exact before/after hashes, counts, and representative samples;
3. requires confirmation for a nonblank Smart Copy overwrite or threshold-sized
   operation;
4. re-captures and byte-compares the complete typed matrix immediately before
   writing;
5. writes through the Excel application-state guard on the owning thread;
6. re-captures and verifies every cell;
7. restores and verifies the complete before-matrix on any write/postcondition
   failure; and
8. records a bounded, case-sensitive optimistic receipt only after success.

Formula mutation refuses before writing when the undo receipt store is absent.
Workbook receipt memory is capped at 4,000,000 before/after characters and old
receipts are evicted deterministically.

## Profile migration

Profile schema v5 adds `formula_iferror_fallback`. V2-v4 profiles migrate to the
qualified default `0`. New v5 profiles require the field, reject a leading `=`,
enforce a 1,024-character ceiling, and refuse inspect-only formula syntax.

## Verification

- Release suite: **225 passed**, zero failed.
- Debug and Release solution builds: zero warnings, zero errors.
- Injected stale plan, missing receipt, partial write, rollback, exact
  postcondition, case-sensitive receipt, and undo tests pass.
- Packed-XLL hidden Excel smoke passes exact copy-down:
  - source `=B10+$C$1` remains unchanged;
  - destinations become `=B11+$C$1` and `=B12+$C$1`;
  - one ExcelAccel undo restores both destinations to blank;
  - existing style/format/navigation/fault/refusal checks still pass; and
  - workbook close and `Excel.Quit()` return with no surviving Excel process.

The first host run correctly refused because Excel represents blanks in a mixed
`Range.Formula` matrix as empty strings. The adapter now classifies an empty
formula plus empty/null calculated value as a blank typed cell; qualification
then passed. This behavior is covered in the persistent real-Excel harness.
