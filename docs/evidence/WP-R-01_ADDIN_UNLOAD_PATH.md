# WP-R-01 Add-in unload path

Date: 2026-08-23
Status: Complete. The largest recorded reliability gap is closed.

## Contract

- **Capability:** CAP-REL-001
- **Acceptance:** AC-REL-005, AC-REL-011 (teardown resilience)
- **Allowed implementation:** the teardown sequence, its host wiring, the smoke
  hook that exercises it, and their tests.
- **Excluded:** everything else; no command behaviour changed.

## The gap, confirmed rather than assumed

`docs/PROJECT_STATUS.md` recorded that `AutoClose` was never invoked. Measured
again before starting: **290 `addin.open` events, 0 `addin.close`.** So the whole
unload path — eleven runtime resets, the quarantine clear, and the session-marker
deletion — had never executed once.

The cause is that Excel does not call `xlAutoClose` during a COM-automated quit.
On process exit the runtime's own `ProcessExit`/`DomainUnload` handlers delete the
marker, which is why no stale markers were ever observed and why nothing looked
wrong. The uncovered case is the one that matters: a user **disabling or removing
the add-in while Excel keeps running**, where the resets are what close modeless
dialogs and release state.

## The defect that was hiding there

`AutoClose` ran all eleven resets inline inside one `CallbackBoundary.RunLifecycle`
block, which catches and logs. A reset that threw therefore skipped **every later
reset and `RuntimeState.StopCleanly()` with it**.

`StopCleanly` is what deletes the `*.running` session marker. So a single failing
dialog reset would have left a stale marker, and the user's **next** Excel session
would have started in safe mode refusing every mutation command — the exact
failure mode already documented as looking like an unrelated formatting problem.

Nothing had ever run this path, so nothing had ever found it.

## What changed

`LifecycleTeardown.Run` performs a named step sequence where a failure in one step
cannot abandon the others, and a final step runs in a `finally` regardless. It
lives in the application layer specifically so it can be fault-injected: the host
project has no test project, which is why this logic was untestable where it was.

`AddInLifecycle.AutoClose` now delegates to it, passing the eleven resets and
`RuntimeState.StopCleanly` as the always-run step. Failures are logged per step
under `addin.close.<name>` and summarised as `partial:<n>`.

A reporter that throws is also contained, since a logging failure defeating the
teardown would reproduce the original defect by another route.

## Verification

```
build Release   0 warnings, 0 errors
build Debug     0 warnings, 0 errors
test  Release   614/614 passed (was 608/608)
smoke           scripts/Test-ExcelAddIn.ps1 PASS, including addin_unload=exercised
                Excel exited, 0 surviving processes, 0 stale session markers
```

Six fault-injection tests cover the sequence, including a step that throws, every
step throwing, the final step throwing, and a throwing reporter. The middle one is
the point: `AThrowingStepDoesNotStopTheOnesAfterItOrTheFinalStep` asserts the
order `first, last, always` when the middle step throws. Against the old inline
code that assertion fails.

In real Excel, the new smoke hook opens a modeless dialog, drives the full
teardown, and reopens:

```
smoke.addin.unload  marker_before=True|cleared=True|shutting_down=True|reopened=True
addin.close         normal
```

That is the first `addin.close` in 291 recorded sessions. The step runs last in
the smoke, because teardown resets the profile, undo, and view runtimes, so
anything after it would be exercising a reopened session.

## What this does not cover

Excel still does not call `AutoClose` on a COM-automated quit, and that is Excel's
behaviour rather than something the add-in can force. The path is now exercised
deliberately instead of never; a real user disabling the add-in follows the same
code. Marker cleanup on abrupt process exit continues to rely on the
`ProcessExit`/`DomainUnload` handlers, which is unchanged and correct.
