# Project status and restart guide

Snapshot date: **2026-08-19**

Status: **Phase 2 is active; WP-2-01 and WP-2-02a are complete. WP-2-02b is gated; WP-2-03 is next.**

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

## WP-2-01 delivered behavior

- A1-only direct precedents for one selected formula cell, captured through a
  bounded plan, revalidated against the exact source formula, and refused as
  stale if it changed during capture.
- A read-only result view stating status, scan scope, parser coverage,
  limitation/refusal code, unresolved and external edge counts, and whether
  completeness can be claimed. Deduplicated nodes retain every source-edge span.
- Closed external references are listed and never opened. No Excel trace arrow
  or workbook annotation is used, and nothing is written to the workbook.
- `audit.precedents.direct` is registered through the central dispatcher,
  Command Search, and a new Ribbon `Audit` menu on KeyTip route
  `Alt, X, A, A, PD`. It is unavailable, with a stated reason, unless the
  selection is exactly one single-area cell.
- The view is discarded on explicit close, on add-in unload, and when its source
  workbook is no longer open. The workbook probe is read-only and subscribes to
  no Excel event.

## WP-2-02a delivered behavior

The bounded reverse index and the Excel worksheet scan boundary are implemented
for **worksheet scope only**, with progress and cancellation wired.

- Scan scope is an explicitly declared value. Workbook scope, an out-of-scope
  target, and unsupported target notation are each refused with a stable code,
  and an out-of-scope formula is counted as a coverage gap rather than read.
- Each formula is parsed once at build time; queries intersect rectangles and
  never re-parse. The index is capped at 20,000 formulas and truncates
  explicitly. An independent brute-force oracle proves AC-AUD-007 equivalence.
- **Excel's reported used range is untrusted and is never a resource bound.** A
  region above 250,000 cells, or wider than the 10,000-cell band ceiling, is
  refused before any block is read. Otherwise the region is banded by rows so no
  single read exceeds 10,000 cells, and the bands provably tile the region with
  no overlap or gap.
- `OperationProgressTracker` is wired to a real operation for the first time.
  Progress advances monotonically through Snapshot, Analyze, and Completed;
  cancellation is checked before every band; and a cancelled scan is refused
  rather than reported as a partial result.

- A planned region above 25,000 cells must be confirmed before any block is
  read; without confirmation the scan is refused with `AUDIT_PREVIEW_REQUIRED`
  and nothing is read.
- `audit.dependents.direct` is registered through the central dispatcher, Command
  Search, and the Ribbon `Audit` menu on KeyTip route `Alt, X, A, A, DD`. Its
  read-only view states scan scope, coverage gaps, truncation, and whether
  completeness is claimed, and is discarded on close, on unload, and when its
  source workbook closes.
- `AuditPresentationLabels` is the single definition of auditing wording, so the
  precedent and dependent views cannot describe the same state differently, and
  `TraceResultPresentation` plus one shared `TraceViewRuntime` mean the view
  lifecycle exists in one place. Both extractions were additive:
  `DirectPrecedentReport` kept its whole public surface and the ten WP-2-01
  presentation tests were never modified.

Workbook scope stays unqualified pending the workbook-scale performance gate
noted in the implementation plan. The scan has no performance corpus yet, so
AC-AUD-009 rests on the ceilings, the band tiling, and the live smoke rather
than a measured large-worksheet workload. The dependent view deliberately
duplicates the precedent view; see the restart point below.

A defect that shipped in WP-2-01 was found and fixed here: A1 column names were
wrong for every exact multiple of 26 (Z rendered as AZ), which made precedent
capture read and display the wrong cell. See
`docs/evidence/WP-2-02_DIRECT_DEPENDENTS.md`.

## Current verification

- WP-2-02a presentation/registration slice: **389/389 Release tests passed**;
  Release and Debug builds warning-free; the registered `audit.dependents.direct`
  route opened its read-only view in real Excel (`open|success`), preserved the
  selection, and released on explicit close.
- WP-2-02a scan-boundary slice: **376/376 Release tests passed**; Release and
  Debug builds warning-free; the live hidden-Excel dependent scan returned
  `Partial|B200,C200|16|1|Completed`, finding exactly the two direct dependents
  of `A200` and correctly excluding `D200`, which depends on `B200` rather than
  the target. A pre-cancelled scan through the same adapter returned
  `Refused|AUDIT_SCAN_CANCELLED|0`. The selection and workbook contents were
  unchanged and Excel exited naturally.
- WP-2-02a pure-core slice: **361/361 Release tests passed**.
- WP-2-01 presentation/registration slice: **315/315 Release tests passed**;
  Release and Debug builds are warning-free; the packed-XLL hidden-Excel smoke
  passed exact cell/name precedent classification, view open/retain/discard/close
  lifecycle, selection and content preservation, workbook close, and natural
  Excel exit with no surviving process. The bounded Phase 1B feature suite
  measured **1,895 ms**.
- WP-2-01 capture slice: **305/305 Release tests passed**.
- WP-2-01 pure-core slice: **302/302 Release tests passed**.
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

1. WP-2-03: indirect traversal, cycles, caps, and trace navigation. Every
   auditing result now projects into `TraceResultPresentation` and is rendered by
   one shared `TraceViewRuntime`, so a traversal view is a projection, not a new
   window. Lifecycle decisions live in `TraceViewSession` in the Application
   layer, which the test project can reference and which is unit-tested.
3. Do not use Excel trace arrows or workbook annotations.
4. Keep all retained gates above closed unless a dedicated work package supplies
   their missing evidence.
5. Continue the normal per-package Release tests and use the short real-Excel
   smoke whenever a package changes the Excel adapter, host, or command wiring.
6. Resume heavier release qualification only when distribution is approaching.

## Local-worktree caution

`PRD/PRD_Review_Feedback.md` and `.claude/` are separate user work. They have
not been staged or modified by implementation commits.
