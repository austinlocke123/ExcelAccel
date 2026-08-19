# Phase 1A qualification checkpoint

Date: **2026-08-19**

## Decision

Phase 1A source implementation and local reference-machine qualification are
complete enough to begin dependency-safe Phase 1B packages. This is not an
end-user distribution approval. AutoColor execution, real workbook-close event
ownership, CA-signed distribution, clean-VM/support-matrix coverage, and
add-in-attributed startup measurement remain fail-closed gates.

## Verified checkpoint

- Release build: zero warnings and zero errors.
- Unit suite: 145/145 passed.
- Three full performance Qualification runs passed all frozen workload P95
  budgets, with zero independent UI-heartbeat timeouts and a worst response of
  68 ms.
- Peak observed workload working-set delta was 26.77 MB against the 250 MB
  incremental-memory requirement.
- Ten-session real-Excel reliability soak passed 10/10 with natural process
  shutdown.
- Real-Excel smoke covers command refusal, exact formatting postconditions,
  property undo, navigation, workbook close, and natural application exit.
- Workspace-confined unsigned package lifecycle passed install/load,
  upgrade/load, rollback/load, disable, uninstall, and sandbox removal.

## Audit fixes included

- The legacy currency command now verifies the exact postcondition and creates
  the same bounded session undo receipt as profile formatting.
- Failed command results are logged and surfaced as failures rather than being
  mislabeled as refusals; user messages include command identity, remediation,
  code, and diagnostic ID.
- Navigation exposes per-workbook cleanup and has unit coverage. A direct COM
  workbook-close subscription was tested and removed because it retained the
  Excel automation process; add-in shutdown clears all session state, while
  host-owned workbook-close cleanup remains unqualified.
- The performance harness independently probes Excel's message loop and fails
  on missing samples, timeouts, responses above 500 ms, or workload P95 budget
  regressions.
- The package lifecycle waits for each isolated Excel load to exit and allows a
  bounded native XLL-handle drain before owned sandbox deletion.

## Retained gates

- PERF-001 remains open because current startup data measure total Excel launch,
  not ExcelAccel-attributed added cost.
- AutoColor remains unregistered until its transactional Excel adapter, preview,
  rollback/fault injection, and command-specific UX evidence pass.
- Formula-mutating Phase 1B packages remain gated by ADR-0004. Non-formula
  Phase 1B packages may proceed independently.
- Production install remains blocked on CA-issued timestamped signing,
  clean-VM registry/allowlisting evidence, and supported Office/Windows matrix.
