# WP-P0-07 performance baseline protocol

Status: **Initial implementation; budgets are not frozen**

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
Until a human accepts that record, results are exploratory and MUST NOT become a
CI regression gate.

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

The harness records the maximum individual Excel-thread block-call duration.
This is a conservative blocking proxy for PERF-007, not proof that Excel's UI
message loop remained interactive. Before AC-P0-006 can pass, a reviewed probe
must independently demonstrate UI heartbeat/progress/cancellation behavior for
operations whose end-to-end duration exceeds 500 ms.

## Output and privacy

The default output is `.tools/performance/wp-p0-07-latest.json`, which is
already ignored by Git. Output contains timings and machine metadata only. It
MUST NOT contain cell values, formulas, workbook names, user paths beyond the
explicit add-in path, or workbook content.

## Exit gate

WP-P0-07 remains open until:

- the harness produces repeatable `qualification` distributions;
- reference machine and corpus are approved;
- cold/warm variance is reviewed;
- UI responsiveness and cancellation evidence is recorded;
- working-set/resource behavior is bounded across repeated runs;
- Phase 1 P95/memory targets are explicitly frozen or requirements are changed
  through review.
