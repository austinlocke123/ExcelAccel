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

## Fail-closed execution status

Spacing and numeric/date plans operate on the selected destination snapshot and
can use the qualified transactional block executor. Ribbon parameter collection
has not yet been added, so they are not registered as runnable commands.

Transpose and formula-from-above depend on a source range outside the active
destination. Their plans bind both source and destination hashes and are marked
`RequiresExternalSourceRevalidation`. The shared executor explicitly refuses
these plans before any write until a main-thread adapter can recapture the exact
source range without changing selection/focus and compare it immediately before
mutation. This is an intentional safety gate, not silent partial support.

## Verification

- Release suite: **232 passed**, zero failed.
- Golden tests cover spacing destination sets, transpose position/reference
  mapping, mixed anchors, constants, exact source adjacency, two fill directions,
  Excel date epochs/leap behavior, time-component refusal, and the pre-write
  external-source gate.
- Release solution build remains zero warnings and zero errors.

## Remaining WP-1B-08 work

- qualify selection-preserving off-selection capture/revalidation;
- add internal clipboard snapshot and formulas/values/formats paste policies;
- add value-from-above using captured underlying calculated values;
- expose typed parameter UI only after exact-plan confirmation behavior is wired;
- rerun packed-XLL smoke, fault injection, performance, and soak.
