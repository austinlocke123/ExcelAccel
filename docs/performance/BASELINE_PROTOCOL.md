# WP-P0-07 performance baseline protocol

Status: **Workload budgets frozen for Phase 1; startup attribution remains open**

## Scope and authority

This protocol covers PERF-001 through PERF-010, AC-P0-006, and AC-PERF-001.
It measures Phase 0 mechanisms and does not weaken the provisional requirements
or authorize Phase 1 feature work.

The versioned source fixture is
[`../../benchmarks/performance-corpus-v1.json`](../../benchmarks/performance-corpus-v1.json).
The harness generates temporary workbooks from that manifest; binary workbooks,
machine-specific paths, and raw result files are not source artifacts.

## Reference-machine record

Every run MUST record:

- Windows version and process architecture;
- CPU identifier and logical processor count;
- physical memory;
- Excel version/build/bitness;
- .NET runtime and PowerShell version;
- add-in path and commit under test;
- fixture ID/schema and selected run profile;
- UTC timestamp and explicit cold/warm classification.

The first approved qualification machine becomes the Phase 1 reference machine.
The 2026-08-19 qualification record accepts the recorded Windows/Excel x64
machine as the Phase 1 workload-mechanism reference. The corpus P95 limits are
now build gates for the measured block-read/property-write mechanisms. Startup
remains a proxy until Excel-only control measurements isolate add-in-owned cost.

## Cold and warm protocol

1. Build the requested configuration before measuring; build time is excluded.
2. Close all Excel processes and verify no `EXCEL` process exists.
3. Do not open customer workbooks. The harness creates unsaved synthetic books.
4. The first fresh-process startup measurement is `cold`. Subsequent
   fresh-process measurements in the same uninterrupted run are `warm`; each
   still requires clean `Workbook.Close`, `Application.Quit`, and PID exit.
5. Each non-startup workload performs unmeasured warmups followed by measured
   iterations in one isolated Excel process.
6. Record every sample. Do not discard outliers automatically.
7. Use nearest-rank P95, sample standard deviation, and coefficient of
   variation. Report failures and timeouts rather than replacing samples.

`quick` is a harness check only. `qualification` is the minimum profile eligible
for reviewed budgets. At least three qualification runs on separate clean
sessions are required before targets may be frozen.

## Workloads

| Workload | Mechanism | Cells | Current interpretation |
|---|---|---:|---|
| startup | fresh Excel plus packed-XLL registration | n/a | Added startup proxy; Excel launch is reported separately from registration where available. |
| `block_read_1000` | one `Range.Value2` block read | 1,000 | Immediate snapshot mechanism. |
| `property_write_10000` | one rectangular `NumberFormat` assignment | 10,000 | Property-only batched mutation mechanism. |
| `block_read_100000` | one `Range.Value2` block read | 100,000 | Large-snapshot mechanism; analysis/progress is not yet included. |
| `workbook_read_250000` | 20 bounded sheet-level block reads | 250,000 | Workbook snapshot mechanism; full scan analysis is not yet included. |

Synthetic cells contain deterministic numeric values. Fixture preparation is
excluded from measured intervals. The harness records working-set change as an
early resource signal, but WP-P0-07 does not claim managed-retention or COM-leak
qualification from a single run.

## Responsiveness interpretation

Alongside maximum Excel-thread call duration, the harness runs an independent
background Win32 `WM_NULL` probe against Excel's main window. Every workload
must produce at least one sample, zero 250 ms timeouts, and a maximum observed
heartbeat no greater than 500 ms. This proves bounded message-loop response for
the synthetic workload; it is not an end-user input-latency trace. Commands
whose end-to-end duration exceeds 500 ms still require their own progress and
cancellation evidence.

## Output and privacy

The default output is `.tools/performance/wp-p0-07-latest.json`, which is
already ignored by Git. Output contains timings and machine metadata only. It
MUST NOT contain cell values, formulas, workbook names, user paths beyond the
explicit add-in path, or workbook content.

## Exit gate

WP-P0-07 workload-mechanism qualification completed on 2026-08-19 with three
clean `qualification` runs. The corpus limits of 100/500/3,000/8,000 ms and the
250 MB incremental-memory requirement are frozen for Phase 1. Remaining gates
are deliberately narrower:

- isolate Excel-only launch from add-in-owned startup cost for PERF-001;
- retain three-run regression evidence when the reference machine, Excel build,
  add-in runtime, or corpus changes;
- add feature-specific progress/cancellation and retained-memory evidence for
  operations that actually exceed 500 ms or retain large snapshots.
