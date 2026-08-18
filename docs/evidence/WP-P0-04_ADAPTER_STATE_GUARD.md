# WP-P0-04 Excel adapter and state-guard evidence

- Date: 2026-08-18
- Scope: second Phase 0 implementation slice
- Status: local implementation passing; WP-P0-04 remains open for the gaps
  listed below
- Parent evidence: [`PHASE0_VERTICAL_SLICE.md`](PHASE0_VERTICAL_SLICE.md)

## Implemented behavior

### Safety-state snapshot and refusal

The Excel adapter captures plain immutable safety state with no COM object in
the command core:

- area count;
- merged-cell presence, conservatively treating mixed state as unsafe;
- worksheet protection;
- workbook read-only state;
- legacy array presence;
- dynamic-array spill presence and whether the Excel build exposes the required
  spill check.

`format.number.currency` performs side-effect-free `CanExecute` validation both
before planning and immediately before mutation. It returns stable refusal codes
for multi-area, merged, protected, read-only, array/spill, capability-missing,
resource-limit, stale-context, and quarantine conditions. Unsupported state is
never inferred as safe.

### Application-state ownership

The state guard is implemented in the pure core behind
`IApplicationStatePort`. For the qualified property mutation it owns only:

- `EnableEvents`, when it changed the property;
- `ScreenUpdating`, when it changed the property.

It applies state in a fixed order and restores in reverse order. It preserves
already-suppressed external state, restores after guard-entry failure and
mutation failure, and evaluates every owned restoration even when one restore
fails. A restoration failure raises `StateRestoreException`; the callback cannot
report success, writes stable `STATE_RESTORE_FAILED` diagnostics, and
quarantines the command for the rest of the Excel session.

### Bounded COM retry

Adapter operations use a deterministic policy of at most three attempts with
two 25 ms delays. Only these documented transient COM HRESULT classes retry:

- `RPC_E_CALL_REJECTED` (`0x80010001`);
- `RPC_E_SERVERCALL_RETRYLATER` (`0x8001010A`);
- Excel busy (`0x800AC472`).

Other exceptions return immediately. The maximum retry delay introduced by one
adapter operation is 50 ms; there is no jitter, unbounded loop, or worker wait.

## Automated evidence

Debug build and tests:

- build passed with zero warnings and zero errors;
- 20 unit/architecture tests passed, 0 failed;
- fault cases cover successful restoration, mutation failure, partial guard
  entry, restoration failure escalation, externally suppressed state, transient
  retry success, retry exhaustion, and non-transient no-retry behavior;
- command cases cover each safety refusal and safety-state revalidation before
  write.

Real Excel watchdog smoke test:

- packed XLL registered successfully;
- health UDF returned `1.0.0.0`;
- normal currency formatting changed only number format and preserved content;
- an injected failure inside the real Excel state guard restored screen updating
  and event enablement;
- protected, multi-area, and merged targets were refused without formatting;
- the temporary workbook closed, `Excel.Quit()` returned, and the Excel PID
  exited.

The testing also exercised recovery behavior: after a deliberately forced test
process termination left an unclean marker, the next Excel session refused the
mutation in safe mode. That session exited cleanly; the following session ran
normally and passed the full smoke suite.

## Harness correction found during qualification

An early failed automation scenario exited its PowerShell worker before calling
`Excel.Quit`, leaving a hidden Excel PID and locking the packed XLL. The harness
now closes the temporary workbook and calls `Excel.Quit` from its `finally`
path, waits an additional five seconds for the exact recorded Excel PID, and
terminates only that test PID if graceful shutdown fails.

## Remaining WP-P0-04 gaps

- qualify edit-mode detection and refusal without intercepting normal typing;
- qualify table, filtered-range, spill-parent/child, and legacy-array fixtures on
  the supported Excel build matrix;
- prove calculation, alerts, cursor, status-bar, cut/copy, focus, and selection
  ownership only for commands that actually change those properties;
- add deterministic block formula/value snapshot and mutation adapters before
  formula or data commands use them;
- exercise real COM busy/rejected-call injection rather than pure retry tests;
- add workbook-aware event-recursion tests and simultaneous mutation tests;
- run repeated open/close and mutation soak with handle, memory, and duration
  measurements.

No later feature family should bypass these refusals or implement its own state
guard/retry policy.
