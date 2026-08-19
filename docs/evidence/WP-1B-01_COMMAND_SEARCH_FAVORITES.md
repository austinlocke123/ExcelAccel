# WP-1B-01 command search and favorites evidence

Date: **2026-08-19**

## Implemented boundaries

- `CommandSearchIndex` searches immutable registry metadata only. It normalizes
  invariant text, matches name/alias/category/description/shortcut/ID, requires
  every query term, caps the registry at 2,048 and results at 100, and resolves
  ties by display name then stable command ID.
- Availability is injected as `CanExecuteResult`; search displays refusal and
  remediation without invoking a command. The Excel host captures one bounded
  selection/context snapshot per refresh and performs no workbook scan.
- The modeless Excel-owned search window is keyboard complete: typing filters,
  Down moves to results, Enter invokes, Ctrl+D toggles a favorite, Escape closes,
  and every interactive control has an accessible name/description.
- Search releases the global callback/reentrancy lease immediately after opening.
  Later invocations re-enter the normal callback boundary and central dispatcher.
- Favorites are immutable, versioned command references with bounded invariant
  arguments. Missing/incompatible favorites remain visible with remediation.
- Favorite invocation resolves the current command/version and `CanExecute`, then
  can only route through the central dispatcher with `InvocationSource.Favorite`;
  it cannot call feature implementations directly or bypass preview/impact/undo.
- Add/remove is idempotent and atomically rewrites only the local profile. Profile
  schema v3 adds the favorites array; strict v2 profiles migrate to v3 with an
  empty list, while unknown fields and unsupported schema versions still refuse.
- Add-in shutdown closes the modeless search surface before clearing session state.

## Verification

- Debug and Release builds: zero warnings/errors.
- Unit suite: **152 passed**, zero failed.
- The maximum-size 2,048-command pure registry test returned the exact alias match
  within the 100 ms immediate-interaction budget on the qualification machine.
- Tests cover deterministic ranking, availability/refusal display, idempotent
  add/remove, stale/incompatible visibility, router-only favorite invocation,
  atomic persistence, deterministic serialization, and v2→v3 migration.
- Real hidden Excel smoke opened and closed the modeless search UI, then passed
  the existing mutation/refusal/undo/navigation/fault-restoration suite and
  observed natural workbook/application close with no remaining Excel process.

## Retained limits

- Favorites currently accept the bounded argument representation, but registered
  Phase 1A commands expose no fixed host arguments; nonempty arguments therefore
  refuse until a command-specific typed binder is implemented and tested.
- Live Quick Key interception remains separately gated. Command Search uses only
  its owned dialog keystrokes and does not replace built-in Excel shortcuts.
