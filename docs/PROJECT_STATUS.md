# Project status and restart guide

Snapshot date: **2026-08-19**

Status: **Phase 1B engineering checkpoint passed; stacked review and external
release qualification remain.**

ExcelAccel is a native Windows desktop Excel-DNA add-in. The repository now
contains the Phase 0 safety foundation, Phase 1A command/format/navigation
runtime, and the Phase 1B daily-speed feature core. It is not qualified for
end-user distribution and is not an Excel for the web add-in.

## Published branch stack

The repository uses deliberately stacked draft PRs. Review and merge in order;
each PR's base is the branch immediately above it.

| PR | Branch / commit | Scope |
|---:|---|---|
| #1..#7 | Phase 0 stack through `agent/phase-0-closure` | Excel-DNA host, adapter/state safety, formula/collaboration/performance/package spikes, closure ledger |
| #8..#12 | Phase 1A stack through `agent/phase-1a-safety-runtime` / `1914e28` | architecture, command runtime, profiles/keys, formatting/navigation, safety/undo/recovery/installer source |
| #13 | `agent/phase-1b-search-favorites` / `71fbef5` | command search and favorites |
| #14 | `agent/phase-1b-style-recipes` / `316380e` | style recipes and batch undo |
| #15 | `agent/phase-1b-profile-exchange` / `641c77e` | offline profile/binding exchange |
| #16 | `agent/phase-1b-formula-foundation` / `03576c5` | narrow A1 formula foundation and ADR evidence |
| #17 | `agent/phase-1b-formula-commands` / `6986125` | transactional Smart Copy, IFERROR, sign, and unit commands |
| #18 | `agent/phase-1b-paste-fill` / `4163cbe` | transpose, source capture, formulas-only paste, advanced planners |
| #19 | `agent/phase-1b-data-cleaning` / `3c73a57` | Unicode cleaning and display conversions |
| #20 | `agent/phase-1b-selection-tools` / `40fe68c` | deterministic selection tools |
| #21 | `agent/phase-1b-typed-conversions` / `0b52466` | typed text/number/date conversions |
| #22 | `agent/phase-1b-paste-values-fill` / `efdc4ab` | values-only paste, formula/value above, spacing and sequences host routes |
| #23 | `agent/phase-1b-formats-paste` / `81acd6d` | bounded formats-only transaction and receipt-failure rollback |

Current local branch: `agent/phase-1b-qualification`. It contains WP-1B-12
soak/performance evidence and this restart guide; publication is the current
checkpoint action.

## Phase 1B delivered behavior

- Searchable registry, deterministic ranking, favorites, availability reasons,
  and Ribbon KeyTips for every registered command.
- Local/built-in style library with capture, apply, delete policy, exact preview,
  rollback, and optimistic undo.
- Bounded offline profile export/import preview/apply and binding export.
- A1-only formula transformations with parser-specific refusal: Smart Copy,
  row/column spacing, transpose, IFERROR, sign, and unit scaling.
- One internal, expiring source model for formulas-only, values-only, and the
  approved formats-only paste set; the Office/Windows clipboard is untouched.
- Formula/value from above and invariant numeric/date sequence fills.
- Unicode trim/collapse/control removal; explicit number/text/date grammars; and
  all blank/zero/N/A/NM/dash conversions.
- Read-only deterministic selection of formulas, constants, true blanks,
  numeric hardcodes, and parsed external formulas.

All enabled mutation paths are bounded and fail closed. Plans carry exact
fingerprints; execution revalidates, writes under a state guard, verifies the
post-state, compensates on failure, and records optimistic session undo. A
receipt-storage failure rolls the completed mutation back.

## Current verification

- **288/288 Release tests passed.**
- Debug and Release builds: **zero warnings, zero errors**.
- Complete packed-XLL hidden-Excel smoke: passed; bounded Phase 1B feature suite
  measured **1,286 ms**; workbook closed and Excel exited naturally.
- Ten fresh-process soak: **10/10 passed**, all Excel processes exited naturally,
  and the packed XLL unlocked after every run.
- Soak P95: feature suite **1,532 ms**; working set **322,646,016 bytes**;
  private memory **248,094,720 bytes**; handles **1,864** (range 20).
- Frozen Quick performance corpus: all four workloads passed their budgets;
  zero heartbeat timeouts; worst heartbeat **39 ms**.

Detailed evidence: `docs/evidence/WP-1B-12_QUALIFICATION.md`.

## Deliberate retained gates

- Live formula-edit reference toggle remains unregistered because no exact,
  crash-safe caret/edit-text API has been proven. No hooks or injected
  keystrokes are used.
- Calculated-error selection remains unavailable until the typed model can
  represent it without display-text guessing.
- Formats-only paste is capped at 100 cells and nine explicit properties.
- Unknown collaboration state and medium/high-impact collaborative mutation
  remain refused.
- AutoColor enablement, attributed startup cost, long-duration single-process
  retention, supported Office/coexistence/accessibility matrix, CA-signed
  installer, clean-VM lifecycle, and enterprise trust remain release gates.

## Recommended restart point

1. Review and merge the draft stack in order, especially PRs #13 through #23;
   do not merge a child before its base.
2. Review the WP-1B-12 qualification PR and the retained gates above.
3. After the stack is integrated, rerun Release tests, the full hidden-Excel
   smoke, and at least a three-session soak on the merged commit.
4. Then choose explicitly between external release qualification and Phase 2
   WP-2-01 (reference snapshot/index and direct precedents). Do not enable a
   retained gate implicitly while starting Phase 2.

## Local-worktree caution

`PRD/PRD_Review_Feedback.md` and `.claude/` are separate user work. They have
not been staged or modified by implementation commits.
