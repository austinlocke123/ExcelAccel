# WP-P0-07 performance baseline evidence

- Date: 2026-08-19
- Branch: `agent/phase-0-performance-baseline`
- Status: **Initial exploratory slice; AC-P0-006 and AC-PERF-001 remain open**

## Implemented evidence mechanics

- Versioned synthetic corpus: `benchmarks/performance-corpus-v1.json`
- Cold/warm and reference-machine protocol:
  `docs/performance/BASELINE_PROTOCOL.md`
- Isolated real-Excel harness: `scripts/Test-ExcelPerformance.ps1`
- Deterministic nearest-rank percentile, sample-deviation, variance, and P95
  regression-gate logic under `ExcelAccel.Core.Performance`
- Unit coverage for sample validation, percentile behavior, zero baselines, and
  the provisional 15% regression tolerance

The harness refuses to start when any Excel process is already open. Each
workload runs against a generated unsaved workbook in a hidden Excel process,
records the exact PID, requires clean close/Quit/PID exit, and writes raw JSON
only under ignored `.tools/performance/`.

## Build and unit evidence

| Configuration | Build | Tests |
|---|---|---|
| Debug | zero warnings/errors | 84/84 passed |
| Release | zero warnings/errors | 84/84 passed |

## Exploratory Quick run

The first successful end-to-end run used:

- Windows: Microsoft Windows NT 10.0.26200.0 x64
- CPU: AMD64 Family 23 Model 96, 16 logical processors
- Physical memory: 33,528,926,208 bytes
- Excel: 16.0 build 20228 x64
- Fixture/profile: `excelaccel-performance-v1` / `quick`

| Measurement | Exploratory result | Provisional requirement |
|---|---:|---:|
| Cold startup proxy | 1,035.0159 ms | PERF-001: 750 ms P95 |
| Warm startup proxy P95 | 1,005.9107 ms | PERF-001: 750 ms P95 |
| 1,000-cell block read P95 | 1.5209 ms | PERF-002: 100 ms P95 |
| 10,000-cell property write P95 | 10.8824 ms | PERF-003: 500 ms P95 |
| 100,000-cell block read P95 | 75.1496 ms | PERF-004: 3,000 ms P95 |
| 250,000-cell/20-sheet block-read pass P95 | 120.5276 ms | PERF-005: 8,000 ms P95 |

The run completed all four workloads and produced a JSON report. It is not a
pass/fail judgment against the provisional targets: `quick` has too few samples,
startup currently includes Excel process launch plus packed-XLL registration,
and scan workloads measure snapshot mechanics rather than finished feature
analysis/progress behavior.

## Findings

1. Block reads and the property-only write are fast enough to justify continued
   block-based adapter work on this machine.
2. The startup proxy is above the provisional PERF-001 target and must be split
   into Excel launch, add-in registration/load, and ExcelAccel-owned callback
   cost before proposing a requirement change.
3. Maximum individual block-call duration is only a UI-blocking proxy. It does
   not satisfy the required UI heartbeat, progress, or cancellation evidence.
4. A complete Quick run takes several minutes because it uses seven isolated
   Excel sessions. This is acceptable for qualification safety but warrants a
   separate single-session developer diagnostic if iteration time becomes a
   problem.
5. One diagnostic rerun was interrupted by an external 184-second command
   timeout. Its benchmark-owned controller was terminated explicitly, and its
   child processes then exited. A subsequent 226.8-second end-to-end run of the
   exact current implementation passed with no surviving Excel process.

## Remaining exit gates

- record at least three clean `Qualification` runs;
- approve the reference machine and corpus;
- add UI heartbeat/progress/cancellation evidence;
- add repeated resource-retention and variance evidence;
- separate Excel launch time from add-in-owned startup cost;
- freeze Phase 1 budgets through explicit review.
