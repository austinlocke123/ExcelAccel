# WP-1B-11 deterministic selection tools evidence

Date: 2026-08-19

## Implemented commands

- `selection.select.formulas`
- `selection.select.constants`
- `selection.select.blanks`
- `selection.select.numeric_hardcodes`
- `selection.select.external_formulas`

All five commands are in the central registry and Command Search and have a
dedicated Ribbon Select menu with unique `Alt, X, A, L, ...` KeyTips.

## Safety and exactness boundary

- The source must be one unmerged rectangular range and is bounded by the
  existing 10,000-cell typed formula/value snapshot limit.
- Planning builds a stable Boolean match matrix from one snapshot. Horizontal
  runs are compressed and identical runs on consecutive rows are merged.
- Execution refuses more than 64 resulting areas or 4,096 address characters.
- The full source matrix, top-left position, workbook, worksheet, and address
  are revalidated immediately before selection.
- Excel is asked to select the exact comma-separated area plan on the planned
  worksheet. The adapter then captures, position-sorts, and compares every
  reported area address to the plan.
- This command family is read-only: it calls no value, formula, or formatting
  writer and creates no undo receipt.
- All owned Range, Areas, Worksheet, Worksheets, and Workbook COM wrappers are
  explicitly released. The borrowed root Application wrapper is not released.

## Predicate semantics

- Formulas use the typed underlying formula kind, not displayed results.
- Constants are nonblank nonformula cells, including text, numbers, and Boolean
  constants.
- Blanks are true empty cells only; formulas returning empty text do not match.
- Numeric hardcodes are numeric constants only; numeric text and formulas are
  excluded.
- External formulas require a successfully parsed reference whose qualifier
  contains an external-workbook bracket. Bracket-like text inside a formula
  string does not match.

## Verification

- Release suite: **258 passed**, zero failed.
- Debug solution build: zero warnings, zero errors.
- Unit coverage verifies every released predicate, true-blank and numeric-text
  distinctions, parsed external references, deterministic rectangle merging,
  fragmented-area refusal, stale-source refusal before selection, exact
  postcondition verification, and zero write calls.
- Packed-XLL hidden Excel smoke creates a mixed 3x4 range and selects four
  separated numeric constants. It verifies the exact reported address
  `A40,C40,B41,D42`, preserves every formula and constant, retains all previous
  smoke checks, closes the workbook, and observes Excel exit naturally.

## Explicitly gated

- `selection.select.errors`: the current typed capture intentionally has no
  error-cell kind. Guessing from display text would be lossy and is refused.
- used-range scope, hidden/filtered inclusion controls, and automatic navigation
  history remain unregistered until their typed contracts and host composition
  are qualified.
