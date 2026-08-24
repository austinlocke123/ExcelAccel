# WP-R-06 Corpus shapes and in-process retention

Date: 2026-08-24
Status: Complete. The corpus covers more than one shape, and in-process retention
is measurable for the first time.

## Contract

- **Capability:** CAP-PERF-001, CAP-REL-001
- **Acceptance:** AC-PERF-002..006, AC-CHECK-007
- **Allowed implementation:** corpus shapes, the harness that measures them, a new
  in-process retention harness, and the status corrections they justify.
- **Excluded:** changing any budget for the existing dense shape.

## Part one: the corpus covered one shape

The Phase 2 corpus was a single dense rectangle, 2,000 rows by 5 formula columns.
Two shapes were added on their own worksheets, so the dense block and its four
budgets are untouched and remain comparable run to run.

**Sparse**: populated every 50th row across three columns over a 2,000-row span.
Its used range is 7,804 cells for roughly 160 populated ones, which is the
untrusted-used-range hazard the product documents and the dense block never
reaches.

**Wide**: 20 rows by 200 columns, the transpose of the dense block, because
column-major bounding is not the same code path as row-major.

They earned their place immediately. Measured:

```
workload=model_check_worksheet p95_ms=1096.6  detail=Partial;5000;16000;0
workload=model_check_sparse    p95_ms=114.7   detail=Complete;40;7804;0
workload=model_check_wide      p95_ms=386.4   detail=Complete;3920;4020;0
```

The dense shape always truncates at the 5,000-cell cap and reports **Partial**.
Both new shapes report **Complete**. So until now, every Phase 2 qualification run
exercised the truncated path and never the completing one.

## Part two: the soak could never answer the question asked of it

The status note said three iterations cannot show slow leakage and that the
ten-iteration soak covered Phase 1B only. Both halves needed correcting.

**The coverage claim was stale.** The soak loops `Test-ExcelAddIn.ps1`, and that
smoke has grown to cover direct and indirect precedents and dependents, workbook
dependents, the formula inspector, Model Check, the name and link inventories,
AutoColor, and the add-in unload path. It is not a Phase 1B smoke any more.

**More iterations could not have helped.** The soak launches a fresh Excel process
per iteration, which its own limitations already state proves "process cleanup and
cross-session stability rather than in-process long-duration retention". At any
iteration count, a fresh process cannot show whether a long-lived session grows.
The gap was the harness shape, not the iteration count.

A 12-iteration soak was run to confirm cross-session behaviour: all clean exits,
add-in unlocked every time, and first-three versus last-three drift of +0.8%
working set, +1.4% private memory, +2.0% handles, with a non-monotonic handle
series. No cross-session leak.

## Part three: measuring what nothing measured

`scripts/Test-ExcelInProcessRetention.ps1` holds **one** Excel process and one
workbook, repeats the four Phase 2 read-only operations, and samples working set,
private memory, and handles throughout. Drift compares the first and last thirds
after a warm-up, so one noisy reading cannot decide the outcome.

Over 80 measured cycles:

| Quarter | 1 | 2 | 3 | 4 |
|---|---|---|---|---|
| Handles | 1,847 | 1,970 | 1,972 | 1,982 |
| Working set (MB) | 298 | 309 | 308 | 309 |

Resources rise once during warm-up, then plateau. The second-half slope is
+0.46 handles per cycle against a 186-handle spread between the run's minimum and
maximum, so it is not distinguishable from noise.

**Finding: no evidence of unbounded in-process retention over 80 cycles.** The
honest limit of that statement is that a very slow leak and noise look alike at
this length; the harness now exists to ask the question again over a longer run.

An earlier 30-cycle run showed handles rising to 2,025 and then receding to 1,968,
which is what prompted the longer run rather than accepting the first number. A
naive linear slope over 30 cycles read as 6.95 handles per cycle, which the
80-cycle data shows was fitting a warm-up curve, not a leak.

## Verification

```
build Release   0 warnings, 0 errors
test  Release   740/740 passed
phase 2 qual    Quick profile, all six workloads inside budget,
                cancellation refused in 1 ms, diagnostics privacy clean,
                corpus unchanged
soak            12/12 iterations, clean exits, add-in unlocked every time
retention       80 cycles, drift 5.45% handles and 3.64% working set,
                both inside the 10% ceiling
```

Every harness targeted the reported Excel PID and never terminated Excel by name.
The retention harness refuses to start if any Excel is already open, since it
would otherwise measure a process it does not control.
