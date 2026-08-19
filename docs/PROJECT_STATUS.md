# Project status and restart guide

Snapshot date: **2026-08-19**

ExcelAccel has closed Phase 0 for the Phase 1A production foundation. The
repository contains a working Excel-DNA add-in vertical slice, progressively
stacked reliability spikes, and an explicit closure ledger. It is not qualified
for end-user distribution.

## Published branch stack

| Branch | Commit | Scope | Review |
|---|---|---|---|
| `main` | `88b41f7` | Agent-ready engineering specification | GitHub default branch |
| `agent/phase-0-excel-addin` | `6f9ad45` | Phase 0 Excel add-in vertical slice | Draft PR #1 |
| `agent/phase-0-adapter-hardening` | `a4949d2` | Excel adapter and state safety | Draft PR #2, stacked on PR #1 |
| `agent/phase-0-formula-strategy` | `ced68b1` | Formula strategy spike | Draft PR #3, stacked on PR #2 |
| `agent/phase-0-autosave-coauthoring` | `a85143b` | WP-P0-06 collaboration safety | Draft PR #4, stacked on PR #3 |
| `agent/phase-0-performance-baseline` | `29c91e8` | WP-P0-07 performance baseline | Draft PR #5, stacked on PR #4 |
| `agent/phase-0-packaging-trust` | `b30a4ad` | WP-P0-08 packaging and trust | Draft PR #6, stacked on PR #5 |
| `agent/phase-0-closure` | `b95be00` | Gate ledger, ADR acceptance, boundary enforcement, and reliability soak | Draft PR #7, stacked on PR #6 |
| `agent/phase-1a-foundation` | `e13d28e` | WP-1A-01 production solution boundaries | Draft PR #8, stacked on PR #7 |

## Completed engineering work

- WP-P0-01 through WP-P0-03 provide the Excel-DNA lifecycle boundary, isolated
  Excel harness, canonical command lifecycle, one read-only command, and one
  property-only mutation.
- WP-P0-04 provides Excel-thread dispatch, application-state restoration,
  reentrancy protection, bounded retries, stale-context refusal, and COM
  ownership boundaries.
- WP-P0-05 provides the formula corpus/oracle and lossless syntax prototype.
  Formula mutation remains disabled while ADR-0004 is open.
- WP-P0-06 provides conservative collaboration classification, revision and
  lease invalidation, exact-property fingerprints, and fail-closed impact
  policy. ADR-0005 is accepted only with unknown-state refusal.
- WP-P0-07 provides a versioned performance corpus, distribution math, Quick
  real-Excel evidence, and regression-gate mechanics. Full Qualification and
  frozen budgets remain open.
- WP-P0-08 provides package path/hash/length verification, signing mechanics,
  production refusal of untrusted artifacts, isolated package loading, and a
  sandboxed version/upgrade/rollback/disable/uninstall rehearsal.
- Phase 0 closure adds public-API architecture boundary tests and a passing
  ten-session real-Excel reliability soak with natural process exit and XLL
  unlock checks.

Current local verification is 100 unit tests with zero failures and Debug and
Release builds with zero warnings/errors. See
[`evidence/PHASE0_CLOSURE.md`](evidence/PHASE0_CLOSURE.md) for the gate-by-gate
decision and exact measured evidence, and
[`evidence/WP-1A-01_FOUNDATION.md`](evidence/WP-1A-01_FOUNDATION.md) for the
production boundary implementation.

## Accepted and open decisions

- Accepted: ADR-0001 Excel-DNA host.
- Accepted: ADR-0002 net48 x64 host plus netstandard2.0 pure core.
- Proposed: ADR-0003 session-only optimistic property undo.
- Open: ADR-0004 formula parser/semantic transform strategy.
- Accepted with fail-closed scope: ADR-0005 collaboration policy.
- Accepted design, not distribution approval: ADR-0006 package lifecycle.

## Retained gates

- Formula-mutating commands remain unregistered until AC-P0-005 and ADR-0004
  pass.
- WP-1A-07 and frozen performance budgets remain blocked by AC-P0-006,
  including repeatable Qualification completion and UI-heartbeat evidence.
- Unknown collaboration state and medium/high-impact collaborative mutation
  remain refused pending real cloud/build evidence.
- Distribution remains blocked until a CA-issued timestamped signature, signed
  installer, exact registry ownership, allowlisting behavior, and complete
  clean-VM lifecycle pass.
- Supported Excel/Windows build matrix, coexistence, accessibility, and
  long-duration in-process soak remain release qualification.

## Current work and restart point

WP-1A-01 is published on `agent/phase-1a-foundation`. Core, Application,
ExcelInterop, and ExcelAddIn are now separate acyclic projects; command
orchestration moved into Application, Excel adapters moved into ExcelInterop,
the host supplies its root-application/thread delegates, and the packed XLL
embeds every required assembly.

After WP-1A-01 review, continue with WP-1A-02 command context/registry/plan/result
hardening and then WP-1A-03 adapter/state boundaries.

WP-1A-02/03 implementation is now underway: registry metadata, deterministic
canonical plan hashing, exact-preview authorization, explicit result status,
and Excel readiness refusal have been added without registering new commands.

Do not start with formula transforms, AutoColor, installer mutation, or broad
workbook commands: those are still behind their retained gates. The first
Phase 1A slice should consolidate production interfaces, dependency direction,
and test enforcement without expanding workbook mutation authority.

## Local-worktree caution

`PRD/PRD_Review_Feedback.md` and `.claude/` are separate user work. They are not
part of implementation commits and must not be staged unless the user
explicitly requests it.
