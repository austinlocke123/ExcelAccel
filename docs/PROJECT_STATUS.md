# Project status and restart guide

Snapshot date: **2026-08-20**

Status: **Phase 2 is complete.** WP-2-01, WP-2-02a, WP-2-02b, WP-2-03, WP-2-04,
WP-2-05 through WP-2-08 (Model Check), and WP-2-09 qualification are all merged.
There is no Phase 3: the implementation plan deliberately treats the remaining
work as thirteen individually gated packages, each needing its own approval.

For what a fresh session would otherwise rediscover — including that the add-in
is installed and usable today — read
[`SESSION_HANDOFF.md`](SESSION_HANDOFF.md) alongside this file.

ExcelAccel is a native Windows desktop Excel-DNA add-in. The repository contains
the Phase 0 safety foundation, the Phase 1A command/format/navigation runtime,
the Phase 1B daily-speed feature core, and the complete Phase 2 formula auditing
and Model Check capability.

It **is installable and usable on a developer machine today** via
`scripts/New-ExcelAccelPackage.ps1` and `scripts/Install-ExcelAccel.ps1
-AllowUntrustedPrototype`; see the handoff for the exact commands. It is not
qualified for *distribution* to other people, and is not an Excel for the web
add-in. The deferred distribution gates govern shipping to others, not personal
use.

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

- Scan scope is an explicitly declared value. An out-of-scope target and
  unsupported target notation are each refused with a stable code, and an
  out-of-scope formula is counted as a coverage gap rather than read.
- Each formula is parsed once at build time; queries intersect rectangles and
  never re-parse. The index is capped at 20,000 formulas and truncates
  explicitly. An independent brute-force oracle proves AC-AUD-007 equivalence.
- Coverage gaps are counted **per cause**. Structured references, dynamic arrays,
  and intersections are gaps; external references, unions, and resolved names are
  not, because they cannot conceal an in-scope edge. This needed the parser to
  expose every coverage cause rather than only the first, added additively as
  `FormulaSyntaxDocument.LimitationCodes`. A worksheet containing an external
  link can now claim completeness, which it previously never could.
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

Workbook scope was delivered later, in WP-2-02b, once the workbook-scale gate
was resolved.

A defect that shipped in WP-2-01 was found and fixed here: A1 column names were
wrong for every exact multiple of 26 (Z rendered as AZ), which made precedent
capture read and display the wrong cell. See
`docs/evidence/WP-2-02_DIRECT_DEPENDENTS.md`.

## WP-2-03 delivered behavior

Bounded indirect traversal in both directions, with cycles, caps, and trace
navigation.

- `IndirectTraceEngine` is a deterministic breadth-first traversal, hard-ceilinged
  at depth 16 and 5,000 nodes. A node is expanded at most once, so a circular
  model terminates and the revisit is shown as a cycle edge. Reaching either cap
  produces an explicit truncated result with an unexpanded frontier, never a
  silent omission. Every node retains the direct evidence of the edge that reached
  it and the node it came from. A cancelled traversal is refused, so no partial
  traversal is reported as a trace.
- Precedent traversal reads each node's formula and reuses the qualified
  analyzer; a cell with no formula is a leaf, not a gap. Dependent traversal runs
  against one prebuilt worksheet index, so it never rescans per step.
- `audit.precedents.indirect` (`Alt, X, A, A, PI`) and `audit.dependents.indirect`
  (`Alt, X, A, A, DI`) are registered, read-only, and render through the shared
  trace view. The indirect dependent scan applies the same threshold preview as
  the direct one.
- Trace navigation revalidates the target, selects it, and records the prior
  location so session Back returns there. External, unresolved, and cycle rows
  are not navigable.

## Model Check delivered behavior (WP-2-05 to WP-2-08)

- A deterministic read-only rule engine over one immutable snapshot. Rules run in
  stable order, findings sort canonically, a rule failure names the rule and makes
  the scan partial rather than dropping it, and a cancelled scan is refused with
  the prior result left in place.
- Seven rules: pattern inconsistency, constant interrupting a formula region,
  embedded numeric constants, formula errors and broken references, external
  references, circular references, and number-format inconsistency.
- Findings carry rule, version, severity, target, evidence, coverage category, and
  a SHA-256 fingerprint containing no raw formula or value content. No finding
  declares correctness or carries a score, and a test asserts the vocabulary.
- Selection, worksheet, and workbook scopes. A worksheet scan confirms anything
  above 25,000 cells before reading; a workbook scan always confirms its sheet
  inventory.
- Navigation, local ignores by exact fingerprint, rescan against a fresh snapshot,
  and export behind a confirmed manifest that excludes formulas and values by
  default. Ignores live in their own atomic local file rather than the profile
  schema; see the evidence for why and what remains open.
- Six commands registered on the Ribbon Model Check menu, KeyTip `Alt, X, A, K`.

## WP-2-09 qualification

Phase 2 now has a measured workload where it previously had only ceilings. A
2,000-row by 5-column corpus produces 16,000 scanned cells and 10,040 formulas,
with rule violations and a 40-deep precedent chain seeded at known rows.

Measured P95 across three fresh processes, qualification profile, against the
Debug packed XLL (so a conservative upper bound):

| Workload | P95 | Provisional budget |
|---|---:|---:|
| Direct precedents | 1.5 ms | 750 ms |
| Worksheet dependent scan | 200.3 ms | 20,000 ms |
| Indirect precedents | 8.0 ms | 12,000 ms |
| Worksheet Model Check | 950.9 ms | 30,000 ms |

Cancellation refuses in 1 ms with `CHECK_SCAN_CANCELLED` and no findings. Handle
count P95 1,946; working set P95 318,914,150 bytes. Every iteration exited Excel
naturally with no surviving process.

The run also fired two bounded behaviours for the first time against a real
workbook: the Model Check finding cap truncated at 5,000 explicitly, and the
traversal depth cap stopped the 40-deep chain at depth 8 with an unexpanded
frontier.

Privacy: a marker seeded into a formula, a value, a defined name, and the
worksheet name survived into **none** of the exported diagnostics.

Measured values sit far below the provisional budgets, so those budgets are loose
ceilings whose value is regression detection. The numbers above are the reference
point. See `docs/evidence/WP-2-09_PHASE2_QUALIFICATION.md` for the retained
limitations, which include single-machine measurement and a single corpus shape.

## WP-2-02b workbook scope

The workbook-scale performance gate was **resolved on 2026-08-20: opened,
bounded**, on the strength of the WP-2-09 measurements. Read-only workbook-scope
scanning is now delivered for both dependents and Model Check. Workbook-scale
mutation remains out of scope.

- `WorkbookScanPlan` bounds a workbook scan at 64 worksheets and 1,000,000
  aggregate cells, applying every ceiling in pure code over each worksheet's
  untrusted reported used region.
- A worksheet that cannot be bounded is **excluded with a stated reason** rather
  than failing the whole workbook, and an exclusion is a coverage gap that blocks
  any completeness claim.
- A plan with nothing left to read refuses with the first exclusion reason, so
  one over-large worksheet can never read as "nothing found".
- A workbook scan **always** confirms its sheet inventory before reading
  anything, whatever its size, and stays cancellable throughout.
- `audit.dependents.workbook` (`Alt, X, A, A, DW`) and `model_check.run.workbook`
  (`Alt, X, A, K, MB`) are registered, read-only, and declare a mandatory preview.

## WP-2-04 Formula Inspector delivered behavior

- `FormulaTreeBuilder` builds the immutable syntax tree the qualified parser
  never had: precedence climbing over its token stream, right-associative `^`,
  unary and postfix operators, function argument lists, groups, and typed leaves
  for references, names, numbers, text, booleans, and error literals.
- The build is **additive**. The parser's own tokens, references, coverage
  disposition, and limitation codes are unchanged, so nothing that already worked
  moved.
- A construct the parser marks inspect-only for a structural reason yields an
  explicit limitation **with its exact span**, never a tree that looks complete.
- Every node carries the source span that quotes its own text, and pre-order
  flattening is the keyboard focus order.
- Nothing is evaluated, scored, or explained; a test asserts the vocabulary.
- `audit.formula.inspect` is registered on `Alt, X, A, A, FI` and renders through
  the shared trace view.

## Result presentation is no longer interrupting

A command that succeeds no longer opens a dialog. Successful outcomes go to the
Excel status bar, which needs no dismissing, so keyboard-driven work is never
blocked. Dialogs are reserved for refusals, failures, partial results, and
faults.

The internal `showResult` flag therefore governs **failure** dialogs only. It
stays suppressed solely where a result view already presents the refusal itself:
the three Model Check scan commands and the auditing commands. Model Check rescan
is deliberately not suppressed, because it can refuse before any view exists.

## Fixed: a dead exit check, and the process leak it was hiding

Two defects, found and fixed on 2026-08-20 while qualifying WP-2-04.

**1. The Excel-exit verification was dead across the whole harness.** Six scripts
matched the worker's reported process id with
`[regex]::Match($output, '(?m)^excel_pid=(\d+)$')`. Worker output uses CRLF, so
in .NET the `$` anchor never matches after the digits, the captured id was always
empty, and the entire "did Excel exit?" block was skipped. Five of the six
scripts predate Phase 2, so **every "Excel exited naturally, no surviving
process" claim this harness produced before 2026-08-20 was unverified**, not
wrong on purpose but never actually checked. Fixed by dropping the anchor. The
repaired check immediately caught the defect below, which is the point.

**2. A smoke hook leaked an Excel process.** With the check working, a full smoke
run left one hidden Excel alive behind an Excel-DNA diagnostic window.

Isolating it: the pre-Phase-2 smoke script run against the current add-in was
clean, which ruled out product code. Removing only the trace-navigation step
removed the leak. `ExcelAccel.Smoke.TraceNavigate` took `ActiveWorkbook` and
`Selection` into locals and never released them, ignoring the `ComRelease.Owned`
pattern every production adapter follows. An un-released COM reference keeps
Excel's COM server alive, so the process survived `Quit`.

Fixed by releasing both objects in a `finally`. Two consecutive smoke runs now
pass with the exit check active and zero surviving processes.

**Worth keeping in mind:** the leak was in a Debug-only test hook, never in
shipped code. But it went unnoticed for six merges because the check that should
have caught it was broken, and the same broken check means earlier
process-hygiene evidence rests on the unit and behavioural assertions rather than
on a verified clean exit.

A related note observed while debugging: `AutoClose` never runs in this smoke
setup. `addin.close` has been logged zero times against 263 `addin.open` events,
so the add-in's unload path is not exercised by the smoke at all. That is a real
coverage gap, recorded here rather than fixed.

## Current verification

- **Head of `main`: 528/528 Release tests pass**, Release and Debug builds are
  warning-free, and the hidden-Excel smoke passes with the process-exit check
  working and no surviving Excel process.
- WP-2-04 Formula Inspector: 528/528 Release tests; in real Excel the registered
  route returned `formula_inspector=open|success` with workbook contents
  unchanged.
- WP-2-02b workbook scope: **505/505 Release tests passed**; in real Excel the
  workbook scan returned `Complete|Sheet1!B200,Sheet1!C200,WorkbookScopeProbe!A1|workbook|0`,
  reaching dependents on both worksheets with no coverage gap, and an unconfirmed
  workbook scan failed closed with `AUDIT_PREVIEW_REQUIRED`.
- WP-2-09 Phase 2 qualification: **3/3 iterations passed** on the qualification
  profile over a 16,000-cell corpus; cancellation refused in 1 ms; diagnostics
  privacy clean; no Excel process survived.
- Model Check slice: **502/502 Release tests passed**; Release and Debug builds
  warning-free; in real Excel the registered selection route returned
  `open|success|2|0` against a seeded inconsistent formula with no rule failure,
  rescan repeated the scope, and workbook contents were unchanged.
- WP-2-03 traversal slice: **443/443 Release tests passed**; Release and Debug
  builds warning-free; in real Excel both indirect routes opened their read-only
  view over a live `A210 -> B210 -> C210` chain, trace navigation selected `C210`
  and recorded return history, and workbook contents were unchanged.
- Per-cause coverage slice: **417/417 Release tests passed**; the live worksheet
  scan moved from `Partial|B200,C200|16|1` to `Complete|B200,C200|16|0`.
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

## WP-F-01 delivered behavior

Profile schema v6 replaced nine cycle keys with one `cycles` object holding
family to ordered named cycles, and `ProfileDefinition` went from 18 constructor
parameters to 10. Cycle contents live entirely in the embedded default profile;
no format string, colour, or cycle order is written in C#.

Number-format commands cycle rather than applying once. `format.number.currency`
walks dollar, euro, and pound at zero and two decimals, and the real-Excel smoke
now asserts the advance rather than a fixed format.

`ApplyCurrencyFormatCommand` was deleted: it hard-coded a format string and could
not cycle. Its refusal codes changed as a result, documented in
`evidence/WP-F-01_PROFILE_CYCLES.md`; nothing that used to refuse now succeeds.

Two defects were caught before shipping. The specified default font colour cycle
would have oscillated between two colours, because several categories share a
colour under the default palette and the stateless advance matches the first
index by value; resolution now collapses duplicates by resolved value. And
ordering the default by classification precedence made a font-colour keypress
produce red where it had always produced black, which the smoke caught; the
default is now palette order and AC-FMT-041 was reworded to match.

Verification: Release and Debug builds warning-free, 539/539 Release tests,
`scripts/Test-ExcelAddIn.ps1` passing with Excel exiting cleanly and no stale
session markers.

## WP-F-06 delivered behavior

`formula.units.to_basis_points` multiplies by 10,000 so a rate held as 0.0125
reads as 125, sitting beside the thousands and millions transforms with the
identical impact, preview, and undo contract. The qualified scale allowlist in
`FormulaWrapperTransformer` gained 10,000 and stays an allowlist.

The number-format half of AC-FMT-032 is **not** delivered. The formula block
pipeline writes formula and value through a transactional adapter with rollback
and one undo receipt, and a number format is a different property on a different
port. Extending the plan means teaching the rollback path to write number
formats; stapling a second command on produces two receipts, so one Ctrl+Z would
reverse only the format and leave the values scaled. Both are worse than waiting
for a decision.

## WP-F-07 delivered behavior

`RibbonRouteTests` is the validator `commands/RIBBON_LAYOUT.md` had described
for weeks without it existing. It parses the ribbon XML from source and enforces
that every control has an action and a KeyTip, that no KeyTip duplicates or
prefixes another within the scope Excel resolves it in, that every tagged button
names a registered command, and that every descriptor's route equals both the
route table and the path its button actually has.

It found ten drifted descriptors, six of them previously unknown. Four Model
Check commands built their route by concatenating a different command's route
with a KeyTip fragment, and `navigate.cell.a1` named a Navigate menu KeyTip of
`V` when the menu is `VN`. All five advertised keyboard paths that did nothing,
printed verbatim by Command Search and the cheat sheet. Two style commands
carried labels naming a menu that does not exist, and three favorite commands
spelled the same route two different ways.

`RibbonRoutes.For` now throws on an unknown id. The old fallback returned the
Command Search route, so a typo produced a descriptor that looked routed. Every
registered command already had a real entry, so the fallback only ever hid
mistakes.

## Open design questions blocking WP-F

All three specifications were reviewed and approved on 2026-08-20, and WP-F-01
has landed. The remaining open items below still gate the work packages named
against them.

Settled 2026-08-20:

- Hardcode outranks cross-sheet and external; any numeric literal makes a cell
  a hardcode, with no allowlist.
- Error ranks above hardcode in classification precedence.
- Only numeric literals count as hardcodes; text literals stay black.
- Eight custom cycle slots per family is sufficient, and unconfigured slots must
  be invisible and unreachable rather than inert.
- AutoColor category colours are user-settable, and colour cycles reference
  those categories symbolically so the two never drift apart.

Still open, needed before the work packages they touch:

- **Should `to_basis_points` also apply a number format?** AC-FMT-032 says yes,
  but doing it properly means teaching the transactional formula adapter to
  write a second property type. Today the transform ships without it and the
  user applies a bps format from the number-format cycle. Options: extend the
  adapter, drop the format half from the criterion, or leave it as is.

- **Precedent and dependent tracing behaviour.** The user has flagged this as
  complicated and wants to specify it deliberately rather than have it inferred.
  Nothing in the tracing commands should be redesigned until they do. This is
  the largest open item.
- Underline cycle: whether accounting single and double variants belong in the
  default cycle.
- Border cycles: which edge combinations form the cycle, which line styles, and
  whether a separate border colour cycle exists.
- Currency cycle: sign placement, decimal defaults, and whether a no-symbol
  variant belongs in the default.

## Recommended restart point

Phase 2 is finished, so nothing here is blocking. The most useful next step is to
**use the add-in on a real model and let the friction set the backlog**, ahead of
any remaining plan row.

If engineering work is wanted instead, in rough order of value:

1. **Cover the add-in unload path.** `AutoClose` is never invoked by the smoke:
   `addin.close` has been logged zero times against 263 `addin.open` events, so
   the unload path, the runtime resets, and the recovery-marker cleanup are
   entirely unexercised. This is the largest remaining reliability gap.
2. **WP-G-02 external-link inventory** and **WP-G-01 named-range inventory**.
   Both are read-only and reuse the shared trace view, the registration pattern,
   and the export-with-manifest that already exist, so they are the cheapest
   genuine features left. **WP-G-03 compare** is the larger user win and is
   unblocked now that WP-2-04 has landed.
3. **Settle the WP-1A-12 dependency.** WP-G-04, G-09, G-11, and G-13 depend on
   it; the installer source exists but its GA gates are deferred, so "depends on
   WP-1A-12" is ambiguous. Same class of document conflict as the workbook-scale
   gate, and worth resolving before starting any of them.
4. **Decide AutoColor.** The planner is complete and tested but registered
   nowhere, and execution is hard-stopped pending a transactional adapter,
   rollback and fault-injection evidence, and worksheet-scale preview UI. Note
   that this document lists it as disabled while WP-G-13 assumes workbook
   AutoColor gets built.
5. Folding Model Check ignores into the profile schema, if the separate atomic
   ignore file is not acceptable long term.
6. Extending the Phase 2 corpus beyond one dense rectangular shape, and running a
   long-duration soak of the Phase 2 operations. Three iterations cannot show
   slow leakage, and the existing ten-iteration soak covers Phase 1B only.

Standing constraints for any of the above:

- Do not use Excel trace arrows or workbook annotations.
- Keep the retained gates above closed unless a dedicated work package supplies
  their missing evidence.
- Run the per-package Release tests, and the short real-Excel smoke whenever a
  package changes the Excel adapter, the host, or command wiring.
- Never kill Excel by name during harness work; target the reported PID. A
  force-kill leaves `*.running` markers in `%LOCALAPPDATA%\ExcelAccel\sessions`
  that put the next run into safe mode, which then looks like an unrelated
  formatting failure.

## Local-worktree caution

`PRD/PRD_Review_Feedback.md` and `.claude/` are separate user work and were never
staged or modified by implementation commits. The review feedback is now
untracked and ignored by intent: it stays on its author's machine and is absent
from a fresh clone.
