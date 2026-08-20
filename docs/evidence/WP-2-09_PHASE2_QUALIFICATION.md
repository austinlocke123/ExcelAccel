# WP-2-09 Phase 2 large-corpus, cancellation, privacy, and performance qualification

Date: 2026-08-20

Status: **Complete for the qualification profile on the reference machine.
Phase 2 now has a measured workload where it previously had only ceilings.**

## Contract

- Capability: CAP-AUD-001, CAP-CHECK-001
- Acceptance: AC-AUD-009, AC-AUD-010, AC-CHECK-006, AC-CHECK-007, AC-SEC-002
- Allowed implementation: a Phase 2 corpus definition, Debug-only measurement
  hooks, a qualification harness, and engineering evidence
- Excluded: workbook-scope scanning, which remains gated; frozen release budgets

## What this closes

Every earlier Phase 2 package recorded the same gap: responsiveness and bounded
resources rested on explicit ceilings and the short smoke, not on a measured
workload. This package supplies the workload.

## Corpus

[`benchmarks/phase2-corpus-v1.json`](../../benchmarks/phase2-corpus-v1.json)
defines a 2,000-row by 5-column formula grid over one input column, with
deliberate rule violations seeded at known rows: four pattern exceptions, a
divide-by-zero, a broken `#REF!`, two external references, and a 40-deep
precedent chain. The harness builds it through array assignment, because
cell-by-cell COM writes would dominate the run time and measure the harness
rather than the add-in.

The corpus produces **16,000 scanned cells and 10,040 formulas**, which is an
order of magnitude beyond anything the feature smoke exercised.

## Measurement method

`scripts/Test-ExcelPhase2Qualification.ps1` runs each workload with warmup and
measured iterations, times each operation **inside the add-in** so the reported
milliseconds exclude harness COM setup, and computes a P95 per workload. It runs
each iteration in a fresh hidden Excel process and verifies that process exited.

Measurements are taken against the **Debug** packed XLL, because the timed hooks
are Debug-only and must not ship in Release. Debug is slower than Release, so
every passing measurement here is a conservative upper bound.

## Results, qualification profile, three fresh processes

| Workload | P95 | Provisional budget | Scale |
|---|---:|---:|---|
| `audit_precedents_direct` | **1.5 ms** | 750 ms | 1 precedent, complete |
| `audit_dependents_worksheet` | **200.3 ms** | 20,000 ms | 10,040 formulas scanned, complete |
| `audit_precedents_indirect` | **8.0 ms** | 12,000 ms | depth cap reached at 8 |
| `model_check_worksheet` | **950.9 ms** | 30,000 ms | 16,000 cells, 7 rules, 0 rule failures |

- Cancellation P95: **1 ms**, refusing with `CHECK_SCAN_CANCELLED` and no findings.
- Handle count P95: **1,946**. Working set P95: **318,914,150 bytes**.
- Every iteration closed its workbook, exited Excel naturally, and left no
  surviving process.

Measured values are far below the provisional budgets. The budgets are therefore
loose ceilings rather than tight targets; their value now is regression
detection, and the measured numbers above are the reference point.

## Two bounded behaviours the corpus exercised for real

Both of these were designed and unit-tested earlier. This is the first time
either fired against a real workbook at scale.

- **The Model Check finding cap.** The corpus has 6,000 embedded-literal
  candidates, above the 5,000 finding ceiling, so the scan returned exactly
  5,000 findings as an explicit truncated `Partial` result rather than growing
  without bound. The timing above is therefore a truncated run, which is the
  intended behaviour at this scale.
- **The traversal depth cap.** The 40-deep precedent chain returned a `Partial`
  result at the default depth of 8 with an unexpanded frontier, rather than
  walking the whole chain.

## Cancellation (AC-CHECK-006)

A scan cancelled before it starts refuses in **1 ms** at corpus scale, carrying
`CHECK_SCAN_CANCELLED` and zero findings. No partial scan is presented as
complete, and the harness asserts both the refusal code and the time ceiling.

## Privacy (AC-SEC-002)

A unique marker is seeded into four places that a leak could plausibly travel
through: a formula, a cell value, a defined name, and the worksheet name. After
every Phase 2 operation has run against that workbook, the harness exports the
sanitized diagnostics through the production exporter and searches the written
file for the marker.

Result: **clean**, over a 215,466-byte export. The marker survived into none of
the diagnostics.

## Retained limitations

- **The reference machine is a single developer workstation.** These numbers are
  a regression baseline for this machine, not a cross-machine guarantee, and the
  performance-baseline protocol's reference-hardware requirement is unmet.
- Timings are Debug-build. Release would be faster; the gap is unmeasured.
- Budgets are provisional Phase 2 ceilings, not frozen release budgets. Freezing
  them belongs with the wider release qualification.
- The corpus is one shape: a dense rectangular formula grid on one worksheet. It
  does not cover sparse layouts, many-worksheet workbooks, or wide row-oriented
  models.
- Workbook scope is untested here because it remains refused, pending the
  unresolved workbook-scale performance gate.
- This is a bounded repeat run, not a long-duration soak. Handle count and
  working set are recorded per iteration, but three iterations cannot show slow
  leakage. The existing ten-iteration reliability soak still covers the Phase 1B
  feature suite only.

## Commands

```powershell
./scripts/Test-ExcelPhase2Qualification.ps1 -Profile Quick
./scripts/Test-ExcelPhase2Qualification.ps1 -Profile Qualification -Iterations 3
```
