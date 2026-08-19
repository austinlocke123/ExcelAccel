# WP-1B-06/08 advanced formula and fill evidence

Date: 2026-08-19

## Implemented planners and commands

- Formula row/column spacing uses an explicit interval, exact destination set,
  parsed A1 translation, overwrite detection, and bounded preview.
- Formula transpose maps positions and relative coordinate axes across exact
  source/destination dimensions while preserving constants and excluding
  formatting.
- Formula-from-above translates each exact adjacent source-column formula into
  every destination row.
- Value-from-above copies each adjacent source's underlying calculated value,
  never its formula or displayed text.
- Formulas-only and values-only paste use one internal captured-source model,
  refuse overlap and non-multiple shapes, and support exact-shape or explicit
  whole-multiple repetition.
- Numeric sequence fill uses finite invariant start/step and explicit row-major
  or column-major direction.
- Date sequence fill uses an exact `yyyy-MM-dd` date, integer day step, explicit
  direction, and the active workbook's detected Excel 1900/1904 date system.
  The 1900 leap-year bug is modeled intentionally (`1900-02-28 = 59`,
  `1900-03-01 = 61`).

## Host and transaction boundary

Spacing and numeric/date fill are registered Ribbon/Search commands. Accessible
modal parameter dialogs collect the typed values and validate them without
consulting neighboring cells or process locale. Cancellation occurs before a
plan or mutation. Required previews authorize the exact plan hash only.

The selection-preserving source adapter resolves an exact active-workbook,
worksheet, and address without selecting the source. It captures formula/value
identity plus a separate typed underlying `Value2` matrix. A value-only plan
fingerprints both; execution revalidates both immediately before any destination
write. Formula-result error values refuse the complete plan.

The internal source expires after 30 minutes, is cleared on add-in shutdown,
never uses the Windows/Office clipboard, and never persists workbook content.
Formula and value paste report repetitions and overwrites; value paste also
reports formula-source mappings and destination-formula replacements. All
commands use bounded full-block write, exact verification, compensation, and
optimistic session undo. Source cells and destination formatting are untouched.

Text constants are written with Excel's input apostrophe marker so numeric/date-
like text cannot be silently coerced. The marker is absent from `Value2` and
subsequent typed snapshots.

## Verification

- Release suite: **281 passed**, zero failed.
- Debug solution build: zero warnings and zero errors.
- Golden tests cover spacing sets, transpose position/reference mapping, mixed
  anchors, constants, exact source adjacency, both fill directions, Excel date
  epochs/leap behavior, time-component refusal, values-only repetition,
  formula-to-value counts, calculated error refusal, and stale calculated-value
  refusal before destination write.
- Packed-XLL hidden Excel smoke verifies mixed transpose, calculated formula
  results in values-only paste, exact formula and value fill from above,
  invariant numeric sequence, workbook-date-system sequence, source
  preservation, and exact reverse undo for each destination.
- All previous formatting/style/data/selection/fault checks pass, the workbook
  closes, and Excel exits naturally with no surviving process.

## Formats-only completion

`paste.formats_only` is now registered with a mandatory preview and a separate
property-scoped transaction. Its approved v1 set is number format; font name,
size, bold, italic, and underline; horizontal/vertical alignment; and indent.
It intentionally excludes font color, fill, borders, dimensions, validation,
comments, hyperlinks, values, and formulas. Both source and destination are
capped at 100 cells to keep per-cell COM work responsive and receipts bounded.

The format source is captured with the internal source command and expires with
it. Planning requires same-sheet, nonoverlapping exact-shape or whole-multiple
mapping. Execution revalidates every source and destination property, writes
under one state guard, captures every property again for postcondition proof,
and restores/verifies the complete destination matrix on failure. Undo is
optimistic and refuses if any planned post-state property changed.

Seven additional tests cover serialization, exact repetition, overlap/shape
refusal, verified execution plus undo, stale-source refusal, and verified
rollback when receipt storage fails for format or formula/value matrices. The current
Release suite is **288 passed**, zero failed. Hidden-Excel smoke verifies all
nine approved properties, unchanged values/formulas, preservation of excluded
font/fill colors, exact undo, all previous smoke assertions, and natural Excel
exit.

## Remaining phase work

- Run consolidated Phase 1B fault, locale, performance, and soak qualification
  in WP-1B-12.
