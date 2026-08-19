# WP-1B-02 style recipe evidence

Date: **2026-08-19**

## Implemented scope

- Versioned declarative recipes contain only a strict 12-property formatting
  whitelist: borders, fill color, font bold/color/italic/name/size, horizontal
  and vertical alignment, indent, number format, and underline.
- Each property value has invariant validation and a bounded representation.
  Values, formulas, names, workbook/sheet paths, notes, comments, validation,
  hyperlinks, macros, and executable content have no representation in the API.
- Capture requires exactly one unmerged source cell and reads exactly the
  explicitly selected property IDs. One unsupported request refuses before any
  property read; there is no partial unlabeled capture.
- Profile schema v4 adds deterministic local-style storage and migrates v2/v3
  profiles with empty missing collections. Local writes use the existing
  validate/temp/atomic-replace profile store.
- Nine versioned built-ins are registered as immutable profile data: major and
  minor header, date header, assumption, formula, linked formula, output,
  warning, and total. Built-in IDs cannot be replaced or deleted.
- The modeless Excel-owned Style Library is keyboard accessible: Enter applies,
  Ctrl+N captures, Delete removes a local style, and Escape closes. The Ribbon
  also exposes every built-in through unique custom-tab KeyTips.

## Apply/rollback contract

- Planning captures one immutable target/context and exact before/after values
  for every recipe property. Mixed before-state refuses unless a recipe is
  explicitly skip-safe. Skips and changed property counts remain in the plan.
- Dynamic changed properties use a registry-declared whitelist/subset policy;
  the executable plan contains only the recipe properties it may actually write.
- Threshold preview confirms the exact canonical plan hash. Selection or
  before-state drift refuses before mutation.
- Every property write is immediately read back. Any exception or mismatch
  rolls all attempted properties back in reverse order and verifies the exact
  before-state. Complete rollback returns `Failed`; incomplete rollback returns
  `Partial` with the exact remaining changed property IDs and an error UI.
- Successful multi-property apply creates one bounded in-memory batch receipt.
  Undo validates every after-state before any write, restores all properties in
  reverse order, and attempts to return already-restored properties to their
  post-command state if an undo write fails.

## Verification

- Debug and Release builds: zero warnings/errors.
- Unit suite: **162 passed**, zero failed.
- Tests cover exact capture reads, unsupported-property zero-read refusal,
  built-in locking, deterministic persistence/migration, exact-plan preview,
  verified apply, complete rollback, injected incomplete rollback reporting,
  multi-property optimistic undo, and compensation for a write that reports
  failure after partially mutating its property.
- Real hidden Excel smoke opened/closed the Style Library, applied the four-
  property Major Header recipe, preserved the cell value, restored all four
  original properties with one undo, and exited with no remaining Excel process.

## Retained limits

- Style apply is capped at 50,000 cells and refuses mixed before-state for exact
  rollback. Broader/mixed formatting transactions require a dedicated snapshot
  representation and qualification rather than lossy best-effort behavior.
- Organization policy/locked recipes have a model representation but no policy
  distribution source yet; only built-in and local recipes are effective.
