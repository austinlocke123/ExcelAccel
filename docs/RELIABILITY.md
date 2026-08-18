# Reliability, crash safety, and responsiveness

Status: **Draft for review**  
Applies to: every active capability and every Excel/Office integration point

## 1. Reliability position

ExcelAccel is an in-process extension. A defect can freeze or terminate Excel
and jeopardize unsaved work. Reliability therefore has veto authority over
feature scope: when safe behavior is not proven, the command refuses or remains
deferred.

The release target is zero add-in-attributable Excel crashes in qualification,
soak, and fault-injection testing. This is enforced as evidence, not marketed as
an impossible universal guarantee about the Excel process.

Any Excel crash or hang that occurs while an ExcelAccel callback, snapshot,
analysis completion, mutation, rollback, or shutdown path is active is presumed
add-in-attributable and release-blocking until reproducible evidence establishes
another cause. An unexplained crash is not waived as an "Excel issue."

## 2. Failure-containment model

Every externally initiated path MUST use the common boundary:

1. assign a correlation/diagnostic ID;
2. reject shutdown, reentrant, or invalid state;
3. invoke bounded application logic;
4. catch all managed exceptions before returning to Excel;
5. translate exceptions into stable failure categories;
6. restore owned Excel state;
7. emit sanitized local diagnostics;
8. return without showing recursive or modal error UI from an unsafe context.

This applies to command callbacks, Ribbon callbacks, Excel events, pane events,
timers, async completions, startup, shutdown, and optional PowerPoint callbacks.

No feature may create its own top-level exception policy.

## 3. Crash-path prevention rules

| ID | Rule |
|---|---|
| SAFE-001 | No COM proxy may cross a thread, queue, cache, or domain boundary. |
| SAFE-002 | No fire-and-forget task may retain mutation authority or call Excel. |
| SAFE-003 | No callback may block indefinitely on a task, lock, process, file, network, UI, or COM retry. |
| SAFE-004 | COM busy/rejected-call retries MUST be bounded, jitter-free or predictably delayed, cancellable where safe, and limited to documented transient cases. |
| SAFE-005 | Event recursion and overlapping workbook mutation MUST be prevented by a scoped, workbook-aware guard. |
| SAFE-006 | Background results MUST be discarded when workbook identity, sheet identity, structure, or relevant preconditions changed. |
| SAFE-007 | Native dependencies, unsafe code, hooks, and process injection are prohibited unless an accepted ADR and dedicated crash qualification approve them. |
| SAFE-008 | Forced process termination, `Environment.FailFast`, and equivalent behavior are prohibited. |
| SAFE-009 | Finalizers MUST NOT call Excel. Shutdown cleanup is explicit and bounded. |
| SAFE-010 | A feature with repeated failures MAY be quarantined locally; quarantine cannot mutate a workbook. |

## 4. Excel application-state safety

- State changes are scoped by the common guard.
- Each property has captured, changed, and restored markers.
- Restoration occurs in reverse ownership order.
- An unset or unreadable property is not guessed.
- External changes not owned by the guard are preserved where detectable.
- A restoration failure escalates the command result and triggers feature
  quarantine review.
- Tests inject failures after each state change and mutation boundary.

## 5. Reentrancy and concurrency

- At most one mutating command may run per workbook.
- A read-only operation may overlap only when it uses an immutable snapshot and
  cannot present stale information as current.
- Event handlers triggered by the add-in's own writes are suppressed only for
  the narrow mutation scope and restored immediately.
- Global suppression flags without nesting/workbook identity are prohibited.
- Modal Excel states, formula edit mode, drag/copy modes, calculation activity,
  and shutdown are explicit command preconditions.
- Lock ordering is documented. Code running on the Excel thread MUST NOT wait on
  a worker that needs to marshal back to the Excel thread.

## 6. Undo and rollback

Initial add-in undo is intentionally limited:

- current Excel session only;
- bounded to the latest 20 eligible receipts by default;
- property-scoped formatting and other explicitly qualified mutations only;
- stored in memory with no crash recovery;
- cleared on workbook close and add-in disable;
- validates target existence and current property values against the command's
  recorded post-state;
- refuses rather than overwriting a later change;
- never claims to preserve Excel's native undo stack.

Rollback is part of the active command, while undo is a later user action.
Rollback support must be proven independently for each mutation family.

## 7. AutoSave and coauthoring

- AutoSave state is detected and included in command context.
- The add-in does not silently change AutoSave.
- Coauthoring indicators and event coverage are treated as incomplete until
  Phase 0 proves the supported behavior.
- High-impact mutation refuses in an actively coauthored workbook unless a
  later ADR proves safe bounded semantics.
- Any plan created while AutoSave/coauthoring is active uses fresh property
  preconditions and a short validity window.
- Acceptance tests include remote edits between snapshot, preview, execution,
  and undo where the test environment permits.

## 8. Responsiveness model

Work is divided into three phases:

1. **Snapshot:** bounded reads on the Excel thread.
2. **Analyze/plan:** pure computation, eligible for a worker thread.
3. **Commit:** bounded precondition checks and writes on the Excel thread.

Rules:

- operations expected to exceed 500 ms show non-modal progress;
- cancellation is offered before mutation when safe;
- result panes virtualize large collections and never bind 100,000 objects
  directly to visual controls;
- snapshot and commit are chunked when qualification proves yielding does not
  invalidate transaction semantics;
- UI updates are throttled and coalesced;
- event handlers do not scan workbooks;
- startup performs no workbook scan and opens no pane automatically after an
  unclean session;
- all queues, caches, snapshots, receipts, and result sets have explicit caps.

## 9. Memory and resource safety

- Large two-dimensional COM arrays are converted once into compact immutable
  structures and released from the adapter scope.
- Snapshots have size estimates and reject workloads exceeding configured caps.
- Results use streaming or compact representations where possible, but partial
  results are not exposed as final.
- Cancellation tokens, event subscriptions, timers, file handles, and UI
  resources are owned by explicit scopes.
- Soak tests repeatedly open/close workbooks, show/hide panes, run/cancel
  commands, recalculate, and disable/re-enable the add-in while tracking working
  set, handles, threads, and COM-related symptoms.

## 10. Recovery and safe mode

The add-in writes a content-free operation marker immediately before a mutation
and clears it after state restoration and receipt finalization.

On next startup after an unclean marker:

- load the minimum host only;
- do not replay undo or mutation;
- do not restore panes or run scans;
- quarantine the implicated command ID until the user clears safe mode or a
  newer build supersedes the marker;
- offer local diagnostic export;
- allow the user to disable the add-in cleanly.

The marker is diagnostic evidence, not proof that Excel crashed because of the
add-in.

## 11. Diagnostic categories

- `HOST_CALLBACK_FAILURE`
- `COM_BUSY_TIMEOUT`
- `COM_UNEXPECTED_FAILURE`
- `INVALID_OR_STALE_CONTEXT`
- `PLAN_PRECONDITION_CHANGED`
- `MUTATION_PARTIAL`
- `ROLLBACK_FAILED`
- `STATE_RESTORE_FAILED`
- `BACKGROUND_RESULT_STALE`
- `RESOURCE_LIMIT_EXCEEDED`
- `PROFILE_INVALID`
- `UNSUPPORTED_EXCEL_CAPABILITY`

Messages shown to users remain safe and actionable; detailed exception data is
local and sanitized.

## 12. Required reliability test layers

1. Pure unit tests for state machines, plans, policies, and limits.
2. Static architecture tests for forbidden references and dependencies.
3. COM-boundary integration tests in real Excel processes.
4. Failure injection after every lifecycle transition and adapter write stage.
5. Reentrancy tests triggered by events, recalculation, panes, and callbacks.
6. AutoSave/coauthoring scenarios on supported builds.
7. Startup/shutdown/disable/re-enable loops.
8. Long-running memory, handle, and thread-count soak tests.
9. Process-level crash monitoring that records Excel exit codes and Windows
   failure evidence without automatically uploading workbook content.
10. Clean-VM install, update, rollback, and uninstall tests.

## 13. Release blockers

Any of the following blocks release:

- an add-in-attributable Excel crash, hang, or forced termination;
- an unhandled exception crossing a callback boundary;
- lost or corrupted workbook content in the supported corpus;
- incomplete restoration of required Excel state;
- silent partial mutation reported as success;
- a stale plan applied to changed targets;
- an unbounded wait, retry, queue, cache, history, log, or snapshot;
- core network traffic with licensing/update services disabled;
- a known path that steals edit-mode typing;
- a high-impact command executing under an unqualified AutoSave/coauthoring
  state.
