# ADR-0003: Session-only optimistic property undo

- Status: **Proposed**
- Date: 2026-08-18
- Deciders: open

## Context

Programmatic Excel mutations may clear native Excel undo. The original draft
specified a persistent, encrypted, crash-recoverable journal that would avoid
overwriting unrelated later edits. Reliably tracking every change from users,
recalculation, VBA, coauthors, and other add-ins is not available as a complete
event log. Persistent recovery also materially expands privacy, corruption, and
startup risk.

## Proposed decision

For Phase 1A, implement add-in undo as:

- session-only and in memory;
- bounded to 20 eligible receipts per workbook by default;
- limited to explicitly qualified property-scoped mutations;
- cleared on workbook close, add-in disable, or process exit;
- based on optimistic validation of current property values against the exact
  post-state written by the command;
- all-or-refuse by default for a receipt unless a command contract explicitly
  defines safe property-level partial undo;
- unavailable when snapshot size or context exceeds qualified caps.

## Receipt data

- command ID and version;
- workbook/sheet/target identity valid only for the session;
- property scope;
- compact before-state;
- exact written post-state;
- canonical plan hash;
- completion and rollback status;
- creation order and expiry.

No formula/value content is written to disk by this feature.

## User-visible contract

- Add-in commands may clear Excel's native undo stack.
- Add-in undo covers only eligible ExcelAccel mutations in the current session.
- If a target or property changed after the command, undo refuses rather than
  overwriting the later state.
- Undo is not workbook history, version control, or crash recovery.

## Consequences

- smaller correctness and privacy surface;
- no undo after restart or workbook close;
- broad/formula/value mutations remain deferred until separately qualified;
- persistent undo requires a new ADR.

## Required evidence

AC-REL-011 and AC-REL-012, plus memory-cap, stale-property, deleted-target,
protected-target, AutoSave, coauthoring, and failure-injection cases.
