# WP-F-06 Basis-point unit transform

Date: 2026-08-20
Status: Complete for the value transform; the number-format half is deferred and
explained below.

## Contract

- **Capability:** CAP-FORM-001
- **Acceptance:** AC-FMT-032 (partially; see "What is not delivered")
- **Allowed implementation:** one descriptor, one route, one ribbon button, one
  dispatcher arm, one widened scale allowlist, one test.
- **Excluded:** number-format application, the settings editor, tracing.

## What changed

`formula.units.to_basis_points` multiplies formulas and, when the user opts in,
numeric constants by 10,000, so a rate held as `0.0125` reads as `125`. It sits
in the Units menu beside the thousands and millions transforms and carries the
identical contract: `CommandImpact.Medium`, `PreviewPolicy.Threshold`,
`UndoPolicy.SessionPropertyReceipt`, changed properties `formula` and `value`
under `DeclaredSubset`.

`FormulaWrapperTransformer` previously refused any scale other than 1,000 and
1,000,000. That is an allowlist rather than a range, deliberately, so every
accepted scale is one a reviewer has reasoned about. 10,000 was added to it and
the refusal message updated. The guard remains an allowlist; arbitrary scales
are still refused.

## What is not delivered

AC-FMT-032 also requires applying a `0" bps"` number format. **That half is not
built**, and the acceptance criterion should not be marked complete.

The reason is structural rather than effort. The formula block pipeline writes
`formula` and `value` through a transactional adapter with rollback and a single
undo receipt. A number format is a different property written through a
different port. The two plausible ways to combine them are both worse than
waiting:

- Extending the formula block plan to carry a third property means teaching the
  transactional adapter and its rollback path to write number formats, in the
  one code path in the product with real rollback guarantees.
- Stapling a second command on after the first produces two undo receipts, so a
  single Ctrl+Z would reverse only the format and leave the values scaled — a
  worse outcome than doing nothing.

Recorded as an open decision in `PROJECT_STATUS.md`. In the meantime the user
applies a bps format from the number-format cycle, which is one keypress and
loses nothing.

## Verification

```
build Release   0 warnings, 0 errors
build Debug     0 warnings, 0 errors
test  Release   540/540 passed
smoke           scripts/Test-ExcelAddIn.ps1 PASS
                Excel exited, 0 surviving processes, 0 stale session markers
```

The smoke was run because the ribbon XML changed; a malformed ribbon fails
add-in load, which the harness would catch.

New test: `BasisPointsMultipliesByTenThousandAndCarriesTheUnitTransformContract`
asserts both the arithmetic and that the descriptor carries the same impact,
preview, and undo policy as its sibling transforms.
