# Phase 0 closure decision and gate ledger

- Decision date: 2026-08-19
- Closure scope: permission to begin the Phase 1A production foundation
- Release status: not qualified for distribution
- Governing revision: Acceptance §2.1

## Decision

Phase 0 is closed for Phase 1A foundation work. The architecture, runtime,
command lifecycle, adapter/state safety, conservative collaboration policy,
performance harness, and package lifecycle have enough evidence to stop writing
disposable spikes and begin consolidating production code.

Closure does not waive unproved behavior. External and feature-specific work is
converted into an explicit capability or release gate, and the corresponding
command/installer remains unavailable until that gate passes.

## Exit-gate ledger

| Gate | Phase 1A disposition | Evidence | Retained restriction |
|---|---|---|---|
| AC-P0-001 host/lifecycle | Start gate satisfied for local production-foundation work | Packed x64 XLL registration, health callback, real property mutation, injected-fault restoration, clean close, and ten-session soak | Task pane, coexistence, build matrix, disable/unload, and clean-VM behavior remain WP-1A-11/12 qualification |
| AC-P0-002 runtime | Architecture decision accepted | ADR-0002; deterministic net48/netstandard2.0 Debug and Release artifacts | Clean-machine and representative add-in coexistence remain release gates |
| AC-P0-003 command vertical slice | Passed | Canonical plan, validation, property precondition, mutation, postcondition, result, and refusal tests plus real Excel execution | New commands must reuse the lifecycle; no direct callback mutations |
| AC-P0-004 boundary enforcement | Passed for the current architecture | Core assembly-reference test plus public-API host/Office/COM exposure test; Excel calls remain in the host adapter | Any future worker subsystem must add payload inspection before use |
| AC-P0-005 formula strategy | Capability-gated | Lossless syntax prototype and refusal corpus exist; ADR-0004 remains open | No formula-mutating command may be registered or enabled |
| AC-P0-006 performance baseline | Capability-gated | Versioned corpus, deterministic distribution math, quick real-Excel baseline | WP-1A-07 and frozen budgets wait for accepted qualification runs, UI heartbeat, and reference-machine approval |
| AC-P0-007 collaboration | Passed only for accepted conservative policy | Exact-property fingerprint, revision/lease invalidation, stale-plan refusal, and read-only signal probe | Unknown collaboration state and medium/high-impact collaborative mutations remain refused |
| AC-P0-008 packaging/trust | Release-gated | Manifest/path/hash checks, signing mechanics, production signature refusal, load test, and sandbox lifecycle | No distribution or WP-1A-12 acceptance before CA/timestamp and clean-VM lifecycle |

## Closure verification

The closure slice adds automated architecture boundary checks and a repeated
real-Excel reliability soak.

Results on the reference development machine:

- Debug build: zero warnings and zero errors;
- Release build: zero warnings and zero errors;
- unit suite after boundary enforcement: 99 passed, 0 failed;
- ten fresh hidden Excel sessions: 10 passed, 0 failed;
- each session registered the packed XLL, invoked the health function, executed
  the property-only command, preserved cell content, restored owned application
  state after an injected failure, and refused stale/protected/multi-area/merged
  targets;
- every Excel process exited naturally within the bounded shutdown window;
- the packed XLL accepted an exclusive file handle after every iteration;
- observed soak P95: 6,642.0535 ms end-to-end, 277,413,888-byte working set,
  210,006,016-byte private memory, and 1,805 handles;
- handle-count range across fresh processes: 25.

The ignored machine-readable report is written to
`.tools/reliability/phase0-soak-latest.json`.

## Performance qualification finding

The existing Quick profile remains healthy: the recorded real-Excel sample had
approximately 1,006 ms warm startup P95, 1.5209 ms P95 for a 1,000-cell block
read, 10.8824 ms for a 10,000-cell property write, 75.1496 ms for a 100,000-cell
block read, and 120.5276 ms for a 250,000-cell workbook read.

The full Qualification profile did not produce a report inside a ten-minute
outer execution bound on this development machine. The exact hidden Excel PID
created by the interrupted harness was identified and removed; no user Excel
process was targeted. This is a failed practicality/qualification attempt, not
a passed distribution. Budgets therefore remain provisional and WP-1A-07 stays
gated while the harness gains progress reporting, an approved reference-machine
protocol, and repeatable completion evidence.

## Accepted architecture decisions

- ADR-0001: Excel-DNA is the sole production host.
- ADR-0002: net48 x64 host with netstandard2.0 pure core.
- ADR-0005: conservative collaboration policy with unknown-state refusal.
- ADR-0006: versioned, pack-then-sign, rollback-capable package lifecycle.

ADR-0003 remains proposed until session undo implementation. ADR-0004 remains
open and continues to block formula-mutating commands.

## Phase 1A start boundary

Phase 1A begins with WP-1A-01 production structure and architecture
enforcement. The first implementation slice may consolidate the command and
adapter foundations, but it must not:

- register formula mutation or AutoColor;
- enable medium/high-impact writes in unknown collaborative state;
- persist an undo journal;
- create installer, registry, Trust Center, or machine-wide changes;
- claim frozen performance budgets or release readiness.
