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

## Remaining WP-1B-08 work

- Add the approved property-scoped formatting snapshot and transaction for
  formats-only paste.
- Run consolidated Phase 1B fault, locale, performance, and soak qualification
  in WP-1B-12.
