# WP-F-02 User-defined cycles

Date: 2026-08-20
Status: Cycle management and reachability complete. One part of AC-FMT-039 is
not met and is explained below.

## Contract

- **Capability:** CAP-FMT-004
- **Acceptance:** AC-FMT-026, AC-FMT-027, AC-FMT-028, AC-FMT-029 (met);
  AC-FMT-039 (partly); AC-FMT-040 (logic met, editor UI is WP-F-03)
- **Allowed implementation:** cycle add/remove/rename/reorder, the profile-derived
  command surface, its dispatcher arm, and Command Search indexing.
- **Excluded:** the settings dialog (WP-F-03), Excel round-trip validation
  (AC-FMT-044), the AutoColor adapter, and every tracing command.

## What changed

`ProfileCycleEditor` holds add, remove, rename, set-entries, and move as pure
operations returning a new collection. They live in the application layer, not in
a dialog, so they are testable without a host; WP-F-03's dialog is expected to be
a thin shell over them.

There is deliberately **no "clear the entries" operation**. An empty cycle is
exactly the phantom AC-FMT-039 forbids, and `ProfileCycle` refuses to construct
one. Deletion is the only way a cycle stops existing, and **removing the last
cycle in a family removes the family**, so nothing is left behind for a command
to find and refuse on.

`CycleCommandFactory` builds descriptors for cycles the profile defines that no
built-in command already covers, under ids shaped `format.cycle.<family>.<cycleId>`.
Command Search indexes the static registry plus these, which is what makes a
cycle invented after install reachable at all: the ribbon is a static XML string
and can never grow a button for it.

Those descriptors advertise `"Search Commands, then the cycle name"` rather than
an Alt sequence, because they have no ribbon path and claiming one would be a
lie. `RibbonRoutes.For` is deliberately not called for them; it now throws on an
unknown id, which is the correct behaviour and the reason this is explicit.

Splitting a generated id matches the family against the known closed set rather
than splitting on the last dot, because a cycle id may itself contain a dot.

## What is not met

**AC-FMT-039's "does not appear on the ribbon" is only partly satisfied.**

Since v6, an unconfigured cycle does not exist as an object, so there is no
ribbon button for a cycle that was never created. But the *built-in* buttons are
static XML: if a user deletes the `date` cycle, the Date button remains on the
ribbon and refuses by name when pressed. Making it disappear needs `getVisible`
on every cycle button plus `IRibbonUI.Invalidate` on profile change, which is a
ribbon-lifecycle change rather than a cycle change and did not belong in this
package.

The refusal it gives is at least honest and names the cycle, which is AC-FMT-028
and is tested. Recorded as an open item in `PROJECT_STATUS.md`.

**AC-FMT-040's editor is WP-F-03.** The operations it will call, including the
eight-per-family ceiling and the message naming the limit, exist and are tested
here.

## Verification

```
build Release   0 warnings, 0 errors
build Debug     0 warnings, 0 errors
test  Release   571/571 passed (was 556/556)
smoke           scripts/Test-ExcelAddIn.ps1 PASS
                Excel exited, 0 surviving processes, 0 stale session markers
```

Tests worth naming:

- `AUserCanAddACycleThatDidNotShipWithTheProduct`
- `AddingANinthCycleToAFamilyIsRefusedNamingTheLimit`
- `RemovingTheLastCycleInAFamilyRemovesTheFamily`
- `MovingACycleChangesWhichOneAnUnnamedCommandFollows` — proves slot zero is
  load-bearing for commands whose ribbon label names no particular cycle
- `AnInvalidEntryIsRefusedAndLeavesTheOriginalUntouched`
- `AUserAddedCycleBecomesSearchableUnderItsOwnName`
- `AGeneratedIdResolvesBackToARunnableCycle`
- `ADeletedCycleRefusesByNameAndWritesNothing`
- `GeneratedCommandsAdvertiseAnHonestRouteRatherThanAnAltSequence`
