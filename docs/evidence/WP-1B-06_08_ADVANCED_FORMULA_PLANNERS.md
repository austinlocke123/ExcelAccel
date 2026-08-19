# WP-1B-06/08 advanced formula planner checkpoint

Date: 2026-08-19

## Implemented and qualified in pure/application layers

- Formula row/column spacing with explicit interval, exact destination set,
  A1 translation, overwrite detection, preview threshold, and complete samples.
- Formula transpose position mapping across exact source/destination dimensions.
- Transpose of relative coordinate values and absolute/relative anchor kinds by
  axis, while preserving constants as constants and excluding formatting.
- Formula-from-above using one exact adjacent source row and per-destination row
  translation.
- Numeric sequence fill with finite typed start/step and explicit row-first or
  column-first direction.
- Date sequence fill from a typed `DateTime`, integer day step, explicit
  direction, and explicit Excel 1900/1904 date system. The 1900 leap-year bug is
  modeled intentionally (`1900-02-28 = 59`, `1900-03-01 = 61`).

## Execution status

Spacing and numeric/date plans operate on the selected destination snapshot and
can use the qualified transactional block executor. Ribbon parameter collection
has not yet been added, so they are not registered as runnable commands.

The selection-preserving source adapter is now qualified. It resolves the exact
active-workbook/worksheet/address plan without selecting or activating the
source, captures its current safety/content matrix on Excel's owning thread, and
compares it before destination mutation. A stale or unavailable source refuses
before writing. Targeted destination write, verification, compensation, and
undo likewise do not depend on the current selection.

The Ribbon exposes `Capture Formula Source` and `Transpose Captured Source Here`.
The in-memory source expires after 30 minutes, is cleared on add-in shutdown,
never uses the Windows/Office clipboard, and never persists workbook content.
Transpose always requires confirmation of the exact source/destination plan.
Formula-from-above uses the same qualified adapter semantics but still needs its
final Ribbon destination/source-range composition.

`Paste Formulas Only` is also registered against this internal source. It never
uses the system clipboard, transfers formula cells only, preserves destination
cells mapped from source constants, translates every supported formula by the
exact source/destination displacement, refuses overlap and non-multiple shapes,
and supports only explicit exact-shape or whole-multiple repetition. Repetition,
nonblank overwrite, or threshold-sized changes require exact-plan confirmation.

## Verification

- Release suite: **236 passed**, zero failed.
- Golden tests cover spacing destination sets, transpose position/reference
  mapping, mixed anchors, constants, exact source adjacency, two fill directions,
  Excel date epochs/leap behavior, time-component refusal, and the pre-write
  external-source gate.
- Debug and Release solution builds remain zero warnings and zero errors.
- Packed-XLL hidden Excel smoke transposes a mixed formula/value 2×2 source into
  an off-selection destination, verifies exact formulas/constants, verifies the
  destination selection never changes, restores all four destination cells with
  one undo, and exits Excel naturally with no surviving process.

## Remaining WP-1B-08 work

- add underlying-value capture for values-only paste/value-from-above;
- add the approved formatting snapshot for formats-only paste;
- add value-from-above using captured underlying calculated values;
- expose typed parameter UI only after exact-plan confirmation behavior is wired;
- rerun packed-XLL smoke, fault injection, performance, and soak.
