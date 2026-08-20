# WP-2-02b workbook-scope dependents and Model Check

Date: 2026-08-20

Status: **Complete. Read-only workbook-scope scanning is delivered for both
dependents and Model Check, under explicit bounds.**

## The gate this opens

`IMPLEMENTATION_PLAN.md` §4 held workbook-scale performance closed;
`PROJECT_STATUS.md` no longer listed it. WP-2-02b sat behind that disagreement
from the moment it was written, and it was raised as a stop sign rather than
resolved by default.

**Resolved on 2026-08-20: opened, bounded.** WP-2-09 supplied the evidence the
gate was waiting for: a worksheet dependent scan at 200 ms P95 over 10,040
formulas, and a Model Check scan at 951 ms over 16,000 cells, both far under
budget and both measured on the slower Debug build. The two documents are now
reconciled, and §4's blanket statement is explicitly superseded for this
capability.

Only **read-only scanning** is opened. Workbook-scale mutation remains out of
scope.

## Bounds

`WorkbookScanPlan` applies every ceiling in pure code, over each worksheet's
untrusted reported used region:

- at most **64 worksheets** and **1,000,000 aggregate cells** per scan;
- each worksheet is planned through the existing `DependentScanRegion`, so the
  per-worksheet region and band ceilings still apply;
- a worksheet that cannot be bounded is **excluded with a stated reason** rather
  than failing the whole workbook, and an exclusion is a coverage gap that blocks
  any completeness claim;
- a plan with nothing left to read **refuses with the first exclusion reason**.
  Without that rule, scanning one over-large worksheet would have returned an
  empty plan that reads as "nothing found" — a regression the test suite caught
  during this package;
- the sheet inventory is **always** confirmed before a workbook scan reads
  anything, whatever its size, satisfying the always-required preview for
  workbook scope;
- the scan stays cancellable throughout, checked before every band.

## Registration

- `audit.dependents.workbook` on KeyTip route `Alt, X, A, A, DW`
- `model_check.run.workbook` on KeyTip route `Alt, X, A, K, MB`

Both are read-only, declare no changed property, and declare
`PreviewPolicy.Mandatory` because a workbook scan always confirms.

## Port reshaping

`IDependentScanPort` moved from scope-based to worksheet-name-based capture and
gained `CaptureWorksheetNames()`, matching `IModelCheckSnapshotPort`. Both
coordinators now share the same plan/band/confirm/cancel shape, so workbook scope
is one code path rather than two.

## Verification

- Release and Debug builds: **zero warnings, zero errors**.
- Release tests: **505 passed**, zero failed.
- New coverage: workbook scope reaching dependents on every worksheet, a
  worksheet scope still refusing an off-sheet target, excluded worksheets counted
  as coverage gaps, an unconfirmed workbook scan refusing without reading, a
  confirmed workbook scan reading every worksheet and showing its inventory, and
  one over-large worksheet refusing with its reason rather than returning empty.
- **Real Excel:** with a second worksheet added, the workbook scan returned
  `Complete|Sheet1!B200,Sheet1!C200,WorkbookScopeProbe!A1|workbook|0` — reaching
  dependents on **both** worksheets, declaring workbook scope, with no coverage
  gap. An unconfirmed workbook scan returned
  `Refused|AUDIT_PREVIEW_REQUIRED|0`. Workbook contents were unchanged and Excel
  exited naturally.

## Retained limitations

- The 64-worksheet and 1,000,000-cell ceilings are provisional, chosen to bound
  the work rather than measured. WP-2-09's corpus is single-worksheet, so no
  multi-worksheet workload has been measured.
- Defined-name capture for a workbook scan still anchors on the active
  worksheet's local names plus workbook names; per-worksheet local names on other
  sheets are not enumerated, so a name-bound edge on another sheet stays a
  coverage gap.
- Workbook-scale mutation remains out of scope and unaddressed.
