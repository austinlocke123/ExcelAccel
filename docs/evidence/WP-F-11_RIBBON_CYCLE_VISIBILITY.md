# WP-F-11 Ribbon buttons follow their cycles

Date: 2026-08-23
Status: Complete. AC-FMT-039 is now fully met.

## Contract

- **Capability:** CAP-FMT-004, CAP-UX-002
- **Acceptance:** AC-FMT-039
- **Allowed implementation:** the visibility rule, the ribbon `onLoad` and
  `getVisible` callbacks, invalidation on profile change, and their tests.
- **Excluded:** the settings editor, the AutoColor adapter, tracing.

## The gap this closes

WP-F-02 left one part of AC-FMT-039 unmet: deleting a built-in cycle such as
`date` left its ribbon button in place, refusing by name on every press. The
ribbon XML is static, so a button cannot remove itself.

## What changed

`CycleVisibility.IsVisible` decides, in the application layer where it can be
tested without a host. Fourteen cycle-bound controls carry
`getVisible='OnGetCycleVisible'`, and the callback is a one-liner over that rule.

**The decimals commands are deliberately excluded.** They operate on any number
format, including formats belonging to no cycle, so emptying the whole
`number_format` family must not hide them. A test asserts exactly that.

`onLoad` captures `IRibbonUI`, and `ProfileRuntime.Activate` invalidates it. The
ribbon caches `getVisible` results until something invalidates them, so an
imported profile that removes a cycle would otherwise leave the button visible
for the rest of the session.

A control whose visibility cannot be evaluated stays **visible**, and the failure
is logged. An invisible command is harder to diagnose than one that refuses with
a reason.

## Verifying a callback that no unit test can reach

Excel resolves ribbon callbacks by name and signature at load time, so a mismatch
breaks the ExcelAccel tab without failing any unit test. The existing smoke drives
macros, not ribbon callbacks, so it could not have caught it either.

A Debug-only smoke hook now calls `OnGetCycleVisible` the way Excel does, with a
real `IRibbonControl`, and checks the emitted XML carries both callback
attributes. In real Excel it reported:

```
smoke.ribbon.cycle_visibility  configured=True|decimals=True|wired=True
```

That is a configured cycle staying visible, a decimals command staying visible,
and both attributes present in the XML Excel actually receives.

## Verification

```
build Release   0 warnings, 0 errors
build Debug     0 warnings, 0 errors
test  Release   608/608 passed (was 593/593)
smoke           scripts/Test-ExcelAddIn.ps1 PASS, including ribbon_callbacks=invoked
                Excel exited, 0 surviving processes, 0 stale session markers
```

`EveryHideableButtonRespondsToItsFamilyBeingEmptied` walks all eight property
families and asserts each button hides when its family is emptied, so a control
carrying `getVisible` cannot be one the rule ignores.
