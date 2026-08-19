# WP-P0-07 performance baseline evidence

- Date: 2026-08-19
- Branch: `agent/phase-0-performance-baseline`
- Status: **Three-run workload/heartbeat qualification passed; PERF-001 startup attribution remains open**

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
| Debug | zero warnings/errors | 145/145 passed |
| Release | zero warnings/errors | 145/145 passed |

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

## Three-run Qualification record

Three separate clean Qualification runs completed on the recorded Windows 11 /
Excel 16.0 build 20228 x64 machine. Each run used 15 fresh-process startup
samples and 30 measured iterations per workload. All 57 isolated Excel sessions
per run closed naturally, every workload P95 gate passed, and no report left an
Excel process behind. Raw reports remain ignored under `.tools/performance/`.

| Workload | P95 range across runs | Frozen Phase 1 budget | Worst heartbeat | Heartbeat timeouts | Peak working-set delta |
|---|---:|---:|---:|---:|---:|
| 1,000-cell block read | 1.2515–2.4072 ms | 100 ms | 1 ms | 0 | 0.31 MB |
| 10,000-cell property write | 16.5812–28.2430 ms | 500 ms | 68 ms | 0 | 26.77 MB |
| 100,000-cell block read | 79.0769–112.3336 ms | 3,000 ms | 41 ms | 0 | 3.99 MB |
| 250,000-cell/20-sheet read | 164.5602–244.0693 ms | 8,000 ms | 22 ms | 0 | 15.60 MB |

Warm total-startup proxy P95 ranged from 1,058.3153 to 1,805.1962 ms. That is
not an add-in regression result: it contains Excel process launch and therefore
cannot be compared directly with PERF-001's 750 ms *added add-in cost*. A future
control run must subtract a matched Excel-only startup distribution.

The harness now fails when a workload exceeds its corpus P95 budget or when the
independent background `WM_NULL` probe records no samples, any 250 ms timeout,
or a response above 500 ms. One incomplete attempt was excluded after its outer
orchestration ceiling expired; its child worker and responsive Excel process
then exited naturally. The valid replacement run completed under a corrected
outer allowance.

## Findings

1. Block reads and the property-only write are fast enough to justify continued
   block-based adapter work on this machine.
2. The startup proxy is above the provisional PERF-001 target and must be split
   into Excel launch, add-in registration/load, and ExcelAccel-owned callback
   cost before proposing a requirement change.
3. The independent UI heartbeat now qualifies message-loop responsiveness for
   the synthetic mechanisms. Feature-specific progress and cancellation remain
   required for commands whose end-to-end work exceeds 500 ms.
4. A complete Quick run takes several minutes because it uses seven isolated
   Excel sessions. This is acceptable for qualification safety but warrants a
   separate single-session developer diagnostic if iteration time becomes a
   problem.
5. One diagnostic rerun was interrupted by an external 184-second command
   timeout. Its benchmark-owned controller was terminated explicitly, and its
   child processes then exited. A subsequent 226.8-second end-to-end run of the
   exact current implementation passed with no surviving Excel process.

## Remaining exit gates

- separate Excel launch time from add-in-owned startup cost;
- add feature-specific progress/cancellation evidence when an enabled command
  exceeds 500 ms;
- add long-duration retained-memory evidence for features that retain large
  snapshots;
- repeat qualification when the reference machine, Excel build, runtime, or
  corpus changes.
