# Project status and restart guide

Snapshot date: **2026-08-19**

ExcelAccel is in Phase 0 engineering qualification. The repository contains a
working Excel-DNA add-in vertical slice and progressively stacked reliability
spikes. Product feature families are specified but intentionally remain behind
their implementation gates.

## Published branch stack

| Branch | Commit | Scope | Review |
|---|---|---|---|
| `main` | `88b41f7` | Agent-ready engineering specification | GitHub default branch |
| `agent/phase-0-excel-addin` | `6f9ad45` | Phase 0 Excel add-in vertical slice | Draft PR #1 |
| `agent/phase-0-adapter-hardening` | `a4949d2` | Excel adapter and state safety | Draft PR #2, stacked on PR #1 |
| `agent/phase-0-formula-strategy` | `ced68b1` | Formula strategy spike | Draft PR #3, stacked on PR #2 |
| `agent/phase-0-autosave-coauthoring` | this status snapshot | WP-P0-06 collaboration safety | Approved for publication |

## Completed engineering work

- WP-P0-01 through WP-P0-03 are represented by the vertical slice: Excel-DNA
  lifecycle boundaries, a test/process harness, and read-only plus
  property-only commands.
- WP-P0-04 hardens Excel-thread dispatch, application-state restoration,
  reentrancy protection, bounded retries, stale-context refusal, and COM
  ownership boundaries.
- WP-P0-05 records the formula corpus/oracle evidence and the proposed hybrid
  parsing strategy. ADR-0004 remains proposed pending its acceptance gate.
- WP-P0-06 introduces conservative AutoSave/coauthoring classification,
  workbook revision tracking, expiring plan leases, precondition fingerprints,
  an impact-tier policy matrix, and stale-property refusal before mutation.
  ADR-0005 remains proposed pending real cloud/coauthor build evidence.

The current WP-P0-06 qualification result is:

- Debug and Release builds: zero warnings and zero errors;
- Debug and Release unit suites: 72 of 72 passing in each configuration;
- packed-XLL real-Excel smoke test: passing, including stale-property refusal;
- collaboration signal probe: passing with AutoSave unchanged;
- Excel process cleanup and packed-XLL unlock checks: passing;
- debug smoke entry point excluded from the Release assembly.

Detailed evidence is under [`evidence/`](evidence/), with the current
collaboration policy in [`collaboration/POLICY_MATRIX.md`](collaboration/POLICY_MATRIX.md).

## Open decisions and limitations

- Proposed ADRs are evidence-backed recommendations, not accepted architecture.
- WP-P0-06 does not yet qualify real concurrent Microsoft 365 cloud edits or a
  production remote-change event sink across supported Excel builds.
- The reference machine, benchmark workbook corpus, measurement protocol, and
  frozen Phase 1 performance budgets are not yet defined.
- Installer, signing, clean-VM compatibility, and update/rollback work remains
  in WP-P0-08.
- Feature-family implementation must not begin merely because a spike passed;
  the gates in [`IMPLEMENTATION_PLAN.md`](IMPLEMENTATION_PLAN.md) still apply.

## Recommended restart point

Start **WP-P0-07: performance baseline** on a new branch stacked from the
published WP-P0-06 commit.

The first slice should:

1. inventory PERF-001 through PERF-010 and AC-P0-006/AC-PERF-001;
2. define the reference-machine record and versioned benchmark-workbook corpus;
3. implement a repeatable cold/warm measurement harness for startup, block
   snapshot, property write, cancellation, UI responsiveness, and variance;
4. keep benchmarks outside production command paths and prevent generated
   workbooks or machine-specific output from being committed accidentally;
5. run Debug/Release unit validation plus real-Excel cleanup and lock checks;
6. publish evidence and proposed Phase 1 budgets without silently changing the
   normative performance requirements.

## Local-worktree caution

At this snapshot, `PRD/PRD_Review_Feedback.md` and `.claude/` are separate user
work. They are not part of the Phase 0 implementation commits and must not be
staged unless the user explicitly requests it.
