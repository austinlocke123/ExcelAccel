# WP-F-10 Basis points applies its number format

Date: 2026-08-23
Status: Complete. AC-FMT-032 is now fully met.

## Contract

- **Capability:** CAP-FORM-001, CAP-FMT-004
- **Acceptance:** AC-FMT-032
- **Allowed implementation:** one port member, an optional number format on the
  formula block plan and its execution, the basis-point cycle in the default
  profile, and their tests.
- **Excluded:** the settings editor, the AutoColor adapter, tracing.

## The problem this closes

WP-F-06 shipped the value transform without the format, and recorded why: the
formula block pipeline writes formula and value through a transactional adapter
with rollback and one undo receipt, and a number format is a different property.
Stapling a second command on afterwards would produce two receipts, so a single
Ctrl+Z would reverse only the format and leave the values multiplied by 10,000 —
silently misreading by exactly that factor.

That reasoning was right about stapling and wrong about the conclusion. A
`PropertyBatchReceipt` already carries **up to 32 changes under one receipt id**,
which is precisely the mechanism for "two properties, one undo". Nothing new was
needed.

## What changed

`FormulaBlockPlan` carries an optional `NumberFormat`. When present, `Execute`
captures the target's current formats, writes the contents, applies the format,
reads it back, and records **one batch receipt with two changes**:
`cell_contents_v1` and `cell_format_block_v1`. One Ctrl+Z reverses both.

Only one new port member was required, `ApplyNumberFormat`. Capture and restore
already existed: `IPropertyReceiptPort` reads and writes the whole format block
under `cell_format_block_v1`, which is what undo has always used for
formats-only paste.

Rollback covers the format too. If the contents write fails after the format was
applied, the prior format block is restored alongside the prior contents, and a
failure to restore either is reported as incomplete rather than as success.

If the prior formats cannot be captured at all, the command **refuses before
writing anything**, because there would be no way to make the change undoable.

## The format is data

`0" bps"_);(0" bps")` is not a literal in any C# file. A `basis_points` cycle
joined the `number_format` family in the embedded default profile, and the
dispatcher resolves entry 0 from the active profile.

Two consequences worth noting. The user can edit the bps format like any other
cycle. And because no ribbon button covers `basis_points`, it is automatically
picked up by `CycleCommandFactory` and reachable from Command Search by name — a
basis-point number-format cycle the user did not have before.

## Contract change

The descriptor now declares `cell_format_block_v1` alongside `formula` and
`value`. Without it, `CommandExecutionGate` refused the plan with
`CONTRACT_MISMATCH`, which is the gate doing its job: a command may not touch a
property it never declared. The policy stays `DeclaredSubset`, so a transform
that changes only formulas is still valid.

## Verification

```
build Release   0 warnings, 0 errors
build Debug     0 warnings, 0 errors
test  Release   593/593 passed (was 591/591)
smoke           scripts/Test-ExcelAddIn.ps1 PASS
                Excel exited, 0 surviving processes, 0 stale session markers
```

`BasisPointsAppliesTheFormatAndOneUndoReversesValueAndFormatTogether` is the test
that matters: it scales 0.0125 to 125, asserts the format was applied, then undoes
once and asserts the value is back to 0.0125 **and** the format is back to
General. A two-receipt design fails it.

`BasisPointsRefusesWhenThePriorFormatsCannotBeCaptured` asserts nothing is
written when the change could not be made undoable.
