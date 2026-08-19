# Project status and restart guide

Snapshot date: **2026-08-19**

Status: **Phase 1B is integrated on `main`; lightweight post-merge verification
passed and the next implementation package is Phase 2 WP-2-01.**

ExcelAccel is a native Windows desktop Excel-DNA add-in. The repository now
contains the Phase 0 safety foundation, Phase 1A command/format/navigation
runtime, and the Phase 1B daily-speed feature core. It is not qualified for
end-user distribution and is not an Excel for the web add-in.

## Integrated baseline

PRs #1 through #24 were reviewed, retargeted, and merged into `main` in strict
parent-before-child order on 2026-08-19. The integrated baseline is merge commit
`9091d79`; its tree matches the qualified `agent/phase-1b-qualification` head.

| PRs | Integrated scope |
|---:|---|
| #1..#7 | Phase 0 Excel-DNA host, adapter/state safety, formula/collaboration/performance/package spikes, and closure ledger |
| #8..#12 | Phase 1A architecture, command runtime, profiles/keys, formatting/navigation, safety/undo/recovery, and installer source |
| #13..#24 | Phase 1B discovery, styles, profiles, formula/paste/fill commands, data cleaning, selection tools, bounded formats paste, and qualification evidence |

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

- Post-merge Release build: **zero warnings, zero errors**.
- Post-merge Release tests: **288/288 passed**.
- Post-merge Debug build: **zero warnings, zero errors**.
- Post-merge packed-XLL hidden-Excel smoke: passed; bounded Phase 1B feature
  suite measured **2,400 ms**; workbook closed and Excel exited naturally.
- Post-merge three-session check: **3/3 passed**; every Excel process exited
  naturally and the XLL unlocked after every run. Feature-suite P95 was
  **1,544 ms**, session P95 was **8,280.5087 ms**, handle-count P95 was **1,856**
  with range **8**, and no Excel process survived the check.
- Phase 1B qualification smoke measured the bounded feature suite at
  **1,286 ms**.
- Ten fresh-process soak: **10/10 passed**, all Excel processes exited naturally,
  and the packed XLL unlocked after every run.
- Soak P95: feature suite **1,532 ms**; working set **322,646,016 bytes**;
  private memory **248,094,720 bytes**; handles **1,864** (range 20).
- Frozen Quick performance corpus: all four workloads passed their budgets;
  zero heartbeat timeouts; worst heartbeat **39 ms**.

Detailed evidence: `docs/evidence/WP-1B-12_QUALIFICATION.md`.

## Deliberate retained gates

The following feature gates remain closed but do **not** block Phase 2:

- Live formula-edit reference toggle remains unregistered because no exact,
  crash-safe caret/edit-text API has been proven. No hooks or injected
  keystrokes are used.
- Calculated-error selection remains unavailable until the typed model can
  represent it without display-text guessing.
- Formats-only paste is capped at 100 cells and nine explicit properties.
- Unknown collaboration state and medium/high-impact collaborative mutation
  remain refused.
- AutoColor remains disabled.

Attributed startup cost, long-duration single-process retention, the supported
Office/coexistence/accessibility matrix, CA-signed installer, clean-VM lifecycle,
and enterprise trust remain deferred distribution gates. They are required when
broad distribution approaches, not before ordinary feature development.

## Recommended restart point

1. Review and merge the post-merge status update.
2. Start Phase 2 WP-2-01: reference snapshot/index and direct precedents.
3. Keep all retained gates above closed unless a dedicated work package supplies
   their missing evidence.
4. Continue the normal per-package Release tests and use the short real-Excel
   smoke whenever a package changes the Excel adapter, host, or command wiring.
5. Resume heavier release qualification only when distribution is approaching.

## Local-worktree caution

`PRD/PRD_Review_Feedback.md` and `.claude/` are separate user work. They have
not been staged or modified by implementation commits.
