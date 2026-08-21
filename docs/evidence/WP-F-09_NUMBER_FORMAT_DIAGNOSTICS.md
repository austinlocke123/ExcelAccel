# WP-F-09 Number-format diagnostics

Date: 2026-08-20
Status: The pure half of AC-FMT-044 is complete. The live Excel probe and its
harness corpus are not built; they belong with the settings editor.

## Contract

- **Capability:** CAP-FMT-004
- **Acceptance:** AC-FMT-044 (pure half)
- **Allowed implementation:** advisory validation of a number-format entry.
- **Excluded:** the live probe, the oracle harness, the settings editor, the
  AutoColor adapter, and every tracing command.

## What this guards

A cycle finds its position by matching the cell's **stored** number format
against its entries. If Excel rewrites a format on assignment, the stored string
never matches the entry that produced it, so the comparison always misses and the
cycle sticks on its first entry forever.

Excel does this to locale-qualified currency formats. Measured earlier by
writing, reading back, and comparing code points:

| Written | Stored back | Round-trips |
|---|---|---|
| `£#,##0_);(£#,##0)` | unchanged | yes |
| `€#,##0_);(€#,##0)` | unchanged | yes |
| `[$£-en-GB]#,##0_);(...)` | `[$£-809]#,##0_);(...)` | **no** |
| `[$€-x-euro2]#,##0_);(...)` | `[$€-2]#,##0_);(...)` | **no** |

The qualified forms are exactly what Excel's own currency dialog produces, which
makes this a trap rather than an obvious mistake.

`NumberFormatDiagnostics.Inspect` catches that family statically and suggests the
bare-symbol form. `EvaluateRoundTrip` is the general case: given what was written
and what Excel stored, it reports the mismatch and names the stored form.

It also rejects what is not a number format at all: empty, over 256 characters, a
leading `=`, control characters, unbalanced brackets, an unterminated quoted
string, and more than four sections.

## The important design decision

**These checks are deliberately not enforced in `ProfileCycle`'s constructor.**

Constructor validation also runs during schema migration. Tightening it would
make an existing profile containing a locale-qualified currency format fail to
load, and `ProfileRuntime` falls back to the embedded default when a load throws.
The user would silently lose every setting they had, to fix a formatting nuisance.

A validation improvement that discards a user's profile is worse than the problem
it solves. The checks are therefore advisory, and the settings editor will call
them at the point where the user is typing an entry, where a warning helps and
nothing is lost by refusing.

## Two defects in this package's own code, caught by its tests

The suggested replacement originally rewrote only the **first**
locale-qualified token, leaving the negative section still qualified, so the
suggestion itself would not have round-tripped. It now replaces every occurrence.

The two rewrite paths also worded the same conclusion differently, which a test
comparing them caught.

## Verification

```
build Release   0 warnings, 0 errors
build Debug     0 warnings, 0 errors
test  Release   591/591 passed (was 571/571)
```

No smoke run: nothing reachable from Excel changed.

`EveryDefaultNumberFormatEntryPassesInspection` asserts that every shipped
default passes the product's own check, so the product cannot ship an entry it
would refuse if the user typed it.

## What remains for AC-FMT-044

The live probe, `INumberFormatRoundTripProbe` implemented over a hidden
add-in-owned scratch workbook, and an oracle harness modelled on
`Test-ExcelFormulaOracle.ps1` that proves the shipped defaults round-trip on a
real build. Both need Excel and belong with the settings editor that will call
them.

The probe must not write to the user's workbook: it would pollute Excel's undo
stack and violate the no-unrequested-writes rule.
