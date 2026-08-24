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

- **Head of `main`: 740/740 Release tests pass**, Release and Debug builds are
  warning-free, and the hidden-Excel smoke passes with the process-exit check
  working and no surviving Excel process. The rows below are historical
  per-package records and keep the counts current at the time each landed.
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
- AutoColor **worksheet** scope remains disabled, refusing at planning time until
  its performance qualification, worksheet-scale rollback evidence, and a preview
  built for thousands of rows exist. Selection scope shipped in WP-R-04.

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

## WP-F-08 delivered behavior

AutoColor classification now follows the approved precedence, built on the
formula parser instead of two regular expressions. The old implementation never
returned a hardcode for a formula at all, so `=A1*2` was black where the
specification says blue.

A hardcode outranks external and cross-sheet, and no allowlist applies. The
divergence from Model Check's allowlisted embedded-constant rule is asserted
directly by a test, so neither side can quietly drift into agreeing with the
other. Expect more blue cells than Model Check findings on the same sheet; that
is intended.

An unparseable formula classifies as unsupported and is left alone rather than
guessed at. `ReadEmbeddedLiterals` returns empty both for "no literals" and for a
parse failure, so parse success is checked explicitly.

The execution gate is split: selection scope is permitted, worksheet scope stays
behind `PERFORMANCE_QUALIFICATION_REQUIRED`.

**The two commands are deliberately not registered.** No port reads or writes
per-cell font colours, so registering them would put two buttons on the ribbon
that refuse the moment they are pressed. AC-FMT-036, AC-FMT-037 and AC-FMT-046
are not claimed.

## WP-F-02 delivered behavior

`ProfileCycleEditor` provides add, remove, rename, set-entries and move as pure
operations that return a new collection, so a rejected edit leaves the caller's
profile untouched. There is deliberately no way to empty a cycle: deletion is the
only way one stops existing, and removing the last cycle in a family removes the
family, leaving nothing for a command to find.

A cycle the user invents after install is reachable through Command Search, which
now indexes the profile's cycles alongside the static registry. The ribbon is a
static XML string and can never grow a button for such a cycle, so this is the
only mechanism that can work. Those commands advertise "Search Commands, then the
cycle name" rather than an Alt sequence they do not have.

One part of AC-FMT-039 is not met: deleting a built-in cycle such as `date`
leaves its ribbon button in place, refusing by name when pressed. Removing the
button needs `getVisible` on every cycle button plus `IRibbonUI.Invalidate` on
profile change, which is a ribbon-lifecycle change rather than a cycle change.

## WP-F-09 delivered behavior

`NumberFormatDiagnostics` inspects a cycle entry: structural rejection for what
is not a number format at all, and detection of the locale-qualified currency
tokens Excel rewrites on assignment. A rewritten entry can never match itself, so
its cycle would stick on the first entry forever; the diagnostic names the stored
form and suggests the bare-symbol form that round-trips.

These checks are deliberately **advisory, not constructor-enforced**. Constructor
validation also runs during schema migration, so tightening it would make an
existing profile holding such a format fail to load, and `ProfileRuntime` falls
back to the embedded default when a load throws. Silently discarding a user's
whole profile to fix a formatting nuisance is worse than the nuisance. The
settings editor calls these where the user is typing.

The live probe and the oracle harness remain, and belong with that editor.

## Overnight run, 2026-08-20 to 2026-08-21

Six work packages merged to `main`, each as its own PR with evidence:

| PR | Package | What landed |
|---|---|---|
| #39 | WP-F-01 | Profile schema v6, cycles as data |
| #40 | WP-F-06 | Basis-point unit transform |
| #41 | WP-F-07 | Ribbon route validator, and ten drifted routes |
| #42 | WP-F-08 | AutoColor classification |
| #43 | WP-F-02 | User-defined cycles |
| #44 | WP-F-09 | Number-format diagnostics |

Verification on `main` after the last merge: Release and Debug builds
warning-free, **591/591** Release tests (528 at the start), and
`scripts/Test-ExcelAddIn.ps1` passing with Excel exiting cleanly and no stale
session markers.

### Four defects found that were not on anyone's list

- **The specified default font colour cycle would have oscillated.** Several
  categories share a colour under the default palette, and the stateless advance
  matches the first index by value, so the cycle would have run red, blue, red
  forever with green and black unreachable.
- **Ten descriptors advertised keyboard routes that did nothing**, six of them
  previously unknown, including `navigate.cell.a1` naming a menu KeyTip that does
  not exist. Command Search and the cheat sheet printed them verbatim.
- **AutoColor never classified a formula as a hardcode at all**, so `=A1*2` was
  black where the specification says blue.
- **Ordering the font colour default by classification precedence** changed what a
  keypress had always done, from black to red. The real-Excel smoke caught it.

### Not started, and why

**WP-F-03, the settings editor**, is a WinForms dialog in a project with no test
coverage. The operations it will call exist and are tested; the dialog needs eyes
on it. **The AutoColor adapter** writes per-cell colours with rollback and cannot
be verified without Excel. Both were excluded from this run by decision.

**Tracing is untouched**, as agreed.

### The add-in on this machine is still 0.3.0-local

None of the above is installed. Nothing was installed overnight because it
changes the working environment and the installer requires every Excel closed.
To pick it up, with Excel closed:

```powershell
./scripts/New-ExcelAccelPackage.ps1 -Version "0.4.0-local"
./scripts/Install-ExcelAccel.ps1 -Action Upgrade `
  -PackageDirectory ".tools/packages/ExcelAccel-0.4.0-local-x64" -AllowUntrustedPrototype
```

`-Action Rollback` returns to 0.3.0-local.

Note that upgrading migrates the profile at `%LOCALAPPDATA%\ExcelAccel\profile.json`
from schema 5 to 6 on first read. The migration keeps every existing setting, but
a migrated profile shows one-entry cycles where a fresh install shows three;
"reset to default" is the way to adopt the new defaults.

## WP-F-10 delivered behavior

`formula.units.to_basis_points` now applies the basis-point format as well as
scaling the value, closing AC-FMT-032.

WP-F-06 deferred this on the grounds that a second command would mean a second
undo receipt. That was right about stapling and wrong about the conclusion:
`PropertyBatchReceipt` already carries up to 32 changes under one receipt id,
which is exactly the mechanism for two properties and one undo. Only one new port
member was needed, because `IPropertyReceiptPort` already reads and writes the
whole format block under `cell_format_block_v1`.

If the prior formats cannot be captured, the command refuses before writing
anything, because the change could not be made undoable. Rollback restores the
format alongside the contents.

The format is data: a `basis_points` cycle joined the `number_format` family in
the default profile. Since no ribbon button covers it, it is also reachable from
Command Search by name.

## WP-F-11 delivered behavior

Fourteen cycle-bound ribbon controls now carry `getVisible`, so deleting a cycle
hides its button instead of leaving one that refuses by name on every press. The
decimals commands are deliberately excluded, because they work on any number
format including formats belonging to no cycle.

`onLoad` captures `IRibbonUI` and `ProfileRuntime.Activate` invalidates it, since
the ribbon caches `getVisible` until told otherwise. A control whose visibility
cannot be evaluated stays visible and logs the failure; an invisible command is
harder to diagnose than one that refuses with a reason.

Ribbon callbacks are bound by name and signature at load time, so a mismatch
would break the tab without failing any unit test, and the smoke drives macros
rather than callbacks. A Debug-only hook now calls them the way Excel does; in
real Excel it reported `configured=True|decimals=True|wired=True`.

## Decisions raised 2026-08-20, all answered 2026-08-23

1. ~~**AC-FMT-041 was reworded.**~~ **Confirmed 2026-08-23:** the default font
   colour cycle stays in palette order, so a keypress still produces black first.
2. ~~**Should `to_basis_points` also apply a number format?**~~ **Answered: yes.**
   Delivered in WP-F-10 on one batch receipt, so a single undo reverses both.
3. **The 32-change undo receipt ceiling blocks AutoColor execution.** A real
   selection exceeds it. Options: a new receipt kind, or one coarse property in
   the style of the existing `cell_format_block_v1`.
4. ~~**Deleting a built-in cycle leaves its ribbon button in place.**~~
   **Answered: fix it.** Delivered in WP-F-11.

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

- ~~Ribbon buttons cannot hide themselves.~~ Closed by WP-F-11.

- **The undo receipt ceiling blocks AutoColor execution.** `PropertyBatchReceipt`
  caps at 32 changes, and a real selection exceeds that, so AutoColor cannot
  record undo with today's receipt types. Options: a new receipt kind, or one
  coarse property in the style of the existing `cell_format_block_v1`.

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

## WP-R-01 add-in unload path

Measured before starting: 290 `addin.open` events, 0 `addin.close`. The unload
path had never executed, because Excel does not call `xlAutoClose` during a
COM-automated quit. Marker cleanup on process exit is handled by the runtime's own
exit handlers, which is why nothing looked wrong.

Running it found a real defect. The eleven resets ran inline inside one catching
block, so a reset that threw skipped every later reset **and**
`RuntimeState.StopCleanly` with it. `StopCleanly` deletes the session marker, so
one failing dialog reset would have left a stale marker and put the next Excel
session into safe mode, refusing every mutation command for a reason that looks
unrelated. Nothing had run the path, so nothing had found it.

`LifecycleTeardown.Run` now guarantees each step runs and the marker cleanup runs
in a `finally`. It sits in the application layer so it can be fault-injected; the
host has no test project, which is why the logic was untestable where it was.

In real Excel the new smoke hook opens a modeless dialog, tears down, and reopens:
`marker_before=True|cleared=True|shutting_down=True|reopened=True`, with
`addin.close normal`. That is the first `addin.close` in 291 sessions.

## WP-G-01 named-range inventory

`names.inventory.open`, `names.navigate_target`, and `names.inventory.export`
list every qualified defined name with its scope, visibility, target category,
and navigability. Targets are classified from local metadata alone, so a name
pointing at a closed workbook is classified without opening it. The `#REF!` check
runs before the external check, so a broken external name reports as broken.

**No new window was written.** The inventory projects into
`TraceResultPresentation` and renders through the existing shared
`TraceViewRuntime`, so there is no second read-only view to drift.

Search takes no port at all: re-reading per keystroke would put a workbook scan
behind typing and let results change under a filter the user did not touch.

Export excludes name expressions by default, since a target can carry a path or a
business term, and writes through a temporary file so a failure cannot leave a
half-written export that looks complete.

Real Excel found what no fixture would have invented: a fresh workbook carries
`_xlfn.SINGLE`, a function shim Excel adds for itself, which appeared in a list
the user would expect to hold only their own names. Reserved-name recognition now
lives in `ReservedNames` in the core layer, where it is testable, covering the
classic reserved names and the `_xlfn.`, `_xlref`, `_xludf.`, `_xlchart.`
prefixes.

AC-NAME-008..010, usage navigation, is **not built and not claimed**. It needs
per-category qualification across cell formulas, name expressions, chart series,
validation, and print areas, and the spec says each is separately qualified.

## Session of 2026-08-23

Four packages merged, each as its own PR with evidence:

| PR | Package | What landed |
|---|---|---|
| #45 | WP-F-10 | Basis points applies its number format, on one receipt |
| #46 | WP-F-11 | Ribbon buttons hide when their cycle is deleted |
| #47 | WP-R-01 | Add-in unload path exercised, and the defect it exposed fixed |
| #48 | WP-G-01 | Named-range inventory |

Verification on `main` after the last merge: Release and Debug builds
warning-free, **675/675** Release tests (614 at the start of the day), and
`scripts/Test-ExcelAddIn.ps1` passing with Excel exiting cleanly and no stale
session markers.

**All four decisions raised on 2026-08-20 are now answered and implemented.**
AC-FMT-041 stays palette order; basis points applies its format; the ribbon
hides deleted cycles; and the undo receipt question was answered by finding that
`PropertyBatchReceipt` already carries up to 32 changes under one receipt id,
which is what made WP-F-10 possible without new receipt machinery.

Two defects were found by running code that had never run:

- **The unload path had never executed** in 290 sessions, and running it exposed
  a reset sequence where one throwing step skipped the session-marker cleanup.
  That would have put the user's next Excel session into safe mode.
- **A fresh workbook carries `_xlfn.SINGLE`**, an Excel-internal function shim,
  which appeared in a named-range inventory the user would expect to hold only
  their own names. No fixture would have invented it.

### Where the restart-point items stand

Item 1 is done. Item 2 is half done: named ranges have landed, external links
have not. Items 3 through 6 are untouched, and item 2's second half is the
natural next package.

## WP-G-02 external-link inventory

`links.inventory.open`, `links.navigate_usage`, and `links.inventory.export`
group every qualified external link by normalized source and list where each is
used. Sources normalize to their file name textually, so a formula token and a
full path to the same workbook land in one group without resolving any path.

**Status never claims a check that did not happen.** The vocabulary is about this
Excel session: "Open in Excel", "Not open (not checked)", "Broken reference",
"Unsupported source". No filesystem probe is performed, deliberately — a
`File.Exists` on a UNC path pointing at a dead share can block, and an inventory
that hangs Excel is worse than one that says "not checked". A test asserts the
label says "not checked" and never "missing", "inaccessible", or "unavailable".
The adapter never calls `UpdateLink`, `BreakLink`, `ChangeLink`, or
`Workbooks.Open`, and never triggers a recalculation.

Nothing is hidden: a source Excel reports but no scanned usage explains still
appears, since omitting it would contradict Excel's own Edit Links dialog.
Unscanned categories are named as coverage gaps rather than counted as zero.

The tests caught a real bug. Display strings were chosen by taking the longest
form seen, and a usage token is a raw formula, so a formula fragment would have
appeared in the column the user reads as a file path — and would have defeated the
export's path redaction by smuggling formula text into a redacted column. Only a
source Excel itself reported now contributes a display string.

AC-LINK-009 is half met: non-cell usages are visibly non-navigable with a stated
reason, but no qualified selection path exists for chart series or connections
because neither is scanned yet.

## WP-R-02 settled dependencies and one contract conflict

Two ambiguities that would have misled the next person are resolved.

**A WP-1A-12 dependency is a release gate, not a start gate.** WP-G-04, G-09,
G-11 and G-13 read as blocked on an installer whose GA gates are deferred, which
would have stalled them indefinitely. AC-P0-008 already said the opposite, that
it "does not block source implementation that cannot install itself or alter
Office trust"; the dependency column simply could not express the difference. The
four rows now name implementation dependencies separately from the release gate.
Every phase through Phase 2 shipped and installed locally without those GA gates,
which is the evidence the reading is right.

**Workbook AutoColor is withdrawn, not deferred.** WP-G-13 was scoped to include
it and AC-FMT-019 required it, but the approved `commands/AUTOCOLOR.md` defines
AutoColor as exactly two commands with no workbook scope, and `docs/README.md`
puts `commands/` above `ACCEPTANCE.md` for feature contracts. AC-FMT-019 keeps
its sheet-inventory requirement for workbook **AutoFormat**, which is a different
feature. Reinstating workbook AutoColor now needs a product decision reversing an
approved contract, not a plan edit.

## WP-R-03 font-colour block receipt

The undo ceiling that blocked AutoColor execution is resolved, and not by raising
it. `PropertyBatchReceipt` caps at 32 changes, which is a statement that a receipt
holds a few properties rather than many cells; the product already carries whole
blocks under one coarse property twice, as `cell_contents_v1` and
`cell_format_block_v1`. AutoColor now does the same with
`cell_font_color_block_v1`, so a recolour of any size is one change on one
receipt and the cap is irrelevant rather than raised.

The value groups addresses under their colour, is deterministic so two captures
of one state are byte-identical, and is compared ordinally like the other coarse
properties. The 50,000-cell ceiling is tied to the store's per-value character
limit by a test that fails if either constant moves without the other being
reconsidered.

Its own tests caught a bug: duplicate detection was per colour group, so one cell
listed under two colours passed, then deserialized as one cell with two
conflicting colours and undo would have written whichever it read last.

## WP-R-04 AutoColor execution

`format.auto_color.selection` recolours the selection by what each cell is, and a
single Ctrl+Z reverses the whole change. `format.auto_color.worksheet` is
registered and refuses at planning time, so it is visible and honest rather than
absent. Only font colour is written, declared as the one coarse property WP-R-03
introduced, so plan, write, and receipt all describe the same thing.

Wiring it found two defects. The planner advertised a 250,000-cell bound it could
never reach: the fingerprint was the raw concatenation of every cell and
`PreconditionFingerprint` caps at 1,000,000 characters, so any sizeable range
threw an unhandled `ArgumentOutOfRangeException` instead of refusing. It now
hashes the stream incrementally. And a recolour with more changes than the undo
value can hold is refused before writing, rather than producing a change the user
cannot reverse.

Real Excel, with a typed number, a same-sheet formula, and a text label all
starting from one colour none should end on: the number went blue, the formula
and label went black, and one undo restored all three exactly. The harness asserts
the restored colours, so a stuck recolour or a failed undo fails the smoke.

## WP-R-05 Model Check ignores stay in their own store

Settled rather than changed. `MODEL_CHECK.md` said ignore/unignore performs an
"atomic profile write" and AC-CHECK-031 said it "changes only the atomic local
profile", while the implementation writes a separate TSV beside the profile. The
spec and the code genuinely disagreed, which is what made this item look open.

The separate store is the better design and is kept. The deciding reason is blast
radius: `ProfileStore` refuses a profile it cannot parse whole and
`ProfileRuntime` falls back to the embedded default when a load throws, so
folding ignores in would mean a damaged suppression list costs the user every
cycle, colour, quick key, and favorite they have. It is also the wrong shape,
since an ignore fingerprints a finding in one model while a profile is user-wide
and portable, so an exported profile would carry entries that can never match
elsewhere. Secondary: 2,048 entries is roughly a fifth of the profile's 1 MiB
budget, and folding costs a schema bump with migration and package ripple for no
user benefit.

The spec and the criterion now describe what the code does, including where the
file lives. Tests lock the decision in: the profile carries no ignore data and
`ProfileDefinition` exposes no ignore member, so folding them in later fails
loudly rather than drifting.

The visible-and-removable half of AC-CHECK-033 was already met by the existing
manage-ignores dialog, and ignores are not portable by any unapproved route.

## Recommended restart point

Nothing here is blocking, and no decision is outstanding. The most useful next
step is still to **install and use the add-in on a real model**, letting the
friction set the backlog ahead of any remaining plan row; note that the installed
build is 0.3.0-local and predates everything below.

Remaining engineering work, in rough order of value:

1. ~~**Cover the add-in unload path.**~~ **Done (WP-R-01).** The path is now
   exercised by the smoke and its resilience is fault-injected; a defect that
   would have left a stale session marker was found and fixed.
2. **Both inventories are done** (WP-G-01, WP-G-02). What remains of this item:
   **WP-G-03 compare**, the larger user win, unblocked now that WP-2-04 has
   landed; name **usage** coverage (AC-NAME-008..010); and link scanning for
   chart series, queries/connections, and validation, each of which the spec
   requires to be separately qualified and each of which is currently reported as
   a coverage gap.
3. ~~**Settle the WP-1A-12 dependency.**~~ **Done (WP-R-02).** A WP-1A-12
   dependency gates release only, never a start, which is what AC-P0-008 already
   said; the four package rows now say so. Workbook AutoColor was withdrawn from
   WP-G-13 in the same pass, because the approved command contract rules it out.
4. ~~**Finish AutoColor.**~~ **Selection scope done** (WP-R-03, WP-R-04): it
   recolours, verifies, rolls back, and one undo reverses the whole change.
   **Worksheet scope remains gated** and needs its performance qualification,
   worksheet-scale rollback and fault-injection evidence, and a preview built for
   thousands of rows rather than a message box.
5. ~~Folding Model Check ignores into the profile schema.~~ **Settled (WP-R-05):
   they stay in their own atomic store.** Folding them in would let a damaged
   suppression list cost the user their whole profile, and an ignore fingerprints
   one model while a profile is user-wide. The spec and AC-CHECK-031 said
   "profile" and now describe the file the code actually writes.
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
