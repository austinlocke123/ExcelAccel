# WP-F-08 AutoColor classification

Date: 2026-08-20
Status: Classification, planning, and gating complete. The commands are **not
registered**, because the adapter that writes colours does not exist yet.

## Contract

- **Capability:** CAP-FMT-002
- **Acceptance:** AC-FMT-034, AC-FMT-035 (met); AC-FMT-036, AC-FMT-037,
  AC-FMT-046 (not yet, no command surface exists to assert them against)
- **Allowed implementation:** `AutoColorPlanner` classification and gate.
- **Excluded:** the COM adapter, descriptors, ribbon entries, dispatcher arms,
  the settings editor, and every tracing command.

## What was wrong

`Classify` used two regular expressions and **never returned `NumericHardcode`
for a formula at all**. Every formula fell through to same-sheet, cross-sheet, or
external. `=A1*2` was black, when the approved specification says blue.

The regexes were also naive: `\[[^\]]+\]` matches a bracket inside a string
literal, and the sheet-reference pattern treats any identifier followed by `!` as
a sheet qualifier, so a defined name would be misread.

## What it does now

Precedence is Error, numeric hardcode, external, cross-sheet, same-sheet, text,
first match wins, built on `FormulaParser` and
`FormulaShape.ReadEmbeddedLiterals` instead of pattern matching.

**A hardcode outranks external and cross-sheet.**
`='[Other.xlsx]Sheet1'!A1+5` is a hardcode, not external. The red external signal
is given up on that one cell so the typed number stays findable, and Model
Check's external-reference rule still reports it independently.

**No allowlist.** Any numeric literal makes the cell a hardcode, so `=A1*2`,
`=A1+1`, `=SUM(A1:A9)/12` and `=DATE(2026,1,1)` are all blue. This deliberately
diverges from `check.formula.embedded_numeric_constant`, which allowlists
`0, 1, -1, 2, 100, 12, 365`. The divergence is asserted directly by
`AutoColorIsStricterThanModelCheckOnTheSameLiteral`, so nobody can quietly unify
them: the test fails if either side changes to agree with the other.

Expect more blue cells than Model Check findings on the same sheet. That is the
intended behaviour.

## Two care points, both handled

`ReadEmbeddedLiterals` returns an empty list on a **parse failure** as well as on
"no literals present", so treating empty as "not a hardcode" would silently
mis-colour every formula outside qualified parser coverage. Parse success is now
checked explicitly, and an unparseable formula classifies as `Unsupported`:
counted, and left exactly as it is. Guessing would either hide a hardcode or
invent one.

A sheet-qualified reference to the sheet the formula lives on is **not**
cross-sheet. `=Sheet1!A1` written on Sheet1 is same-sheet. No helper existed for
this; `Classify` now takes the origin worksheet, which the planner supplies from
the selection context.

## The gate is split

`ExecutionGate()` refused unconditionally, which blocked selection scope against
the approved specification. It now takes a scope: **selection is permitted**,
because it is bounded by the selection the user made, and **worksheet stays
refused** with `PERFORMANCE_QUALIFICATION_REQUIRED` until it has a transactional
adapter, rollback and fault-injection evidence, and a worksheet-scale preview.

## Why no commands are registered

Registering `format.auto_color.selection` and `.worksheet` now would put two
buttons on the ribbon that refuse the moment they are pressed, because there is
no port that reads per-cell font colours or writes them. `ExcelSelectionAdapter`
reads and writes the whole selection's colour at once and returns an aggregate
for a mixed range, so it cannot serve this.

Wiring waits for the adapter, which is out of scope for this run by decision:
COM code that writes thousands of cells with rollback cannot be verified without
Excel. AC-FMT-036, AC-FMT-037 and AC-FMT-046 stay unmet until then, and are not
claimed.

## Blocker restated

`PropertyBatchReceipt` caps at **32 changes**. A real selection exceeds that, so
AutoColor cannot record undo with today's receipt types. This does not block
classification, but it must be resolved before AutoColor executes. Options are a
new receipt kind, or a single coarse property in the style of the existing
`cell_format_block_v1`.

## Verification

```
build Release   0 warnings, 0 errors
build Debug     0 warnings, 0 errors
test  Release   556/556 passed (was 546/546)
```

No smoke run: nothing reachable from Excel changed in this package.

The regex implementation failed the new `=A1*2` and hardcode-precedence tests
before the rewrite, which is how the gap was confirmed rather than assumed.
