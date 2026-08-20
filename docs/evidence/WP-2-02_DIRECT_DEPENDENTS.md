# WP-2-02 direct dependents and bounded reverse index

Date: 2026-08-19

Status: **WP-2-02a complete. The reverse index, bounded Excel worksheet scan
boundary, progress/cancellation, threshold preview, read-only result
presentation, and command registration are implemented and verified.
WP-2-02b (workbook scope) remains gated.**

## Contract

- Capability: CAP-AUD-001
- Governing decision: ADR-0004
- Acceptance: AC-AUD-006 through AC-AUD-009
- Allowed implementation: pure Core dependent model, reverse index, bounded
  Excel worksheet scan boundary, read-only presentation, command registration,
  deterministic tests, and engineering evidence
- Excluded: workbook-scope scanning, indirect traversal, trace navigation, Model
  Check, workbook mutation, automatic external-workbook opening, parser
  expansion, and production command registration before the Excel scan boundary
  exists

## Scope slicing

Per the implementation plan, WP-2-02 is delivered in two scope slices.
**WP-2-02a is worksheet scope only.** Workbook scope is representable so that
refusing it is explicit and testable, and stays unqualified until WP-2-02b and
the unresolved workbook-scale performance gate are settled.

## Implemented foundation

- `DependentScanScope` makes the scanned region an explicitly declared value.
  A scan never widens itself: a workbook scope is refused with
  `AUDIT_SCOPE_UNSUPPORTED`, a target outside the declared scope is refused with
  `AUDIT_TARGET_OUTSIDE_SCOPE`, and an out-of-scope formula supplied by a caller
  is counted as a coverage gap rather than read.
- `ReverseReferenceIndex` parses each in-scope formula exactly once at build
  time and stores its resolved reference rectangles. Queries intersect a target
  rectangle against those footprints, so repeated queries never re-parse.
- Direct dependence is an **intersection** relation, not containment: a formula
  that reads any part of the target range reads the target.
- `AuditAddress` and `AuditRectangle` are the shared invariant A1 address model.
  Whole-column, whole-row, multi-area, and sheet-qualified target notation fails
  closed rather than being guessed.
- Cell, range, and defined-name edges each retain their source text, source span,
  and edge kind. Dependents are ordered deterministically by identity.
- A circular self-reference is reported as a dependent rather than hidden.
- The scan is capped at 20,000 formulas and returns an explicit truncated
  partial result with `AUDIT_SCAN_TRUNCATED` rather than running on.
- No COM types, workbook writes, network access, persistence, formula
  evaluation, or automatic opening of external workbooks.

## Bounded worksheet scan boundary

`DependentScanRegion` is the read plan, and every ceiling is applied in pure code
so the bound never depends on what Excel reports.

- **The reported used range is untrusted input, never a resource bound.** Excel
  routinely reports a used range far larger than the real content because of
  stray formatting. `UsedRegionBounds` is deliberately named and documented as
  untrusted, and the ceilings are applied to it before any read happens.
- A region above 250,000 cells is refused with `AUDIT_SCAN_REGION_TOO_LARGE`
  before a single block is read. A full-grid used range hits this immediately.
- A region wider than the 10,000-cell block ceiling is refused rather than split
  unsafely, because a single row band could not be read within the bounded
  formula-block limit the adapter already enforces.
- A region outside the addressable grid is refused with
  `AUDIT_SCAN_REGION_UNSUPPORTED`.
- An empty worksheet plans zero blocks and completes; it is not a refusal.
- Otherwise the region is banded by rows so that no single read exceeds 10,000
  cells. Bands are proven to tile the region exactly, with no overlap and no gap.

`ExcelDependentScanAdapter` reports what the worksheet says and reads exactly the
bands it is handed. It runs through the existing Excel-thread and COM-retry
boundaries, writes nothing, selects nothing, and opens no external workbook.
Defined names are enumerated up to 4,096 and only simple local single-target
definitions are bound; anything else is omitted so the index reports it as an
explicit coverage gap rather than guessing.

## Progress and cancellation

`OperationProgressTracker` is wired to a real operation for the first time. The
coordinator owns the policy so it is testable without Excel:

- progress advances monotonically through Snapshot, Analyze, and Completed, with
  the snapshot phase counting completed bands against the planned band count;
- cancellation is checked before every band read and once more before analysis;
- a cancelled scan is **refused**, never reported as a partial result, so no
  partial scan can be mistaken for a trace. It carries `AUDIT_SCAN_CANCELLED`
  and no dependents;
- because the scan never leaves the Snapshot or Analyze phase, it stays
  cancellable for its whole duration; the tracker's own rule then correctly
  refuses cancellation once the scan has completed.

## Threshold preview

The AUDITING contract requires a preview before a scan above the scan threshold.
A planned region above **25,000 cells** must be confirmed before any block is
read. Without confirmation the scan is refused with `AUDIT_PREVIEW_REQUIRED` and
nothing is read at all. The preview describes the planned read only: worksheet,
target, cell count, and block count.

The confirmation gate deliberately reports no progress.
`OperationPhase.AwaitingConfirmation` sorts *after* `Snapshot` in the phase enum
because that order models a mutation flow, and the tracker requires monotonic
phases; reporting it here would make the later snapshot reports throw. The gate
runs before any work, so it has no progress to report.

## Read-only result presentation

`DirectDependentReport` is a pure Core projection of an existing
`DirectDependentResult`. It formats what the index established and never scans,
resolves, evaluates, reorders, or reclassifies anything.

Every report states the target, workbook, worksheet, status, scan scope,
limitation/refusal code, dependent count, formulas scanned, coverage gaps,
truncation, and whether completeness is claimed. A dependent reached by more than
one kind of reference lists each kind once and retains every source edge with its
span. A refusal is presented in the same view with its categorized code and no
dependent rows.

`DirectDependentView` is a modeless read-only view with accessible names on every
control. It writes nothing, uses no Excel trace arrow or workbook annotation, and
is discarded on explicit close, on add-in unload, and when its source workbook is
no longer open, using the same read-only presence probe as the precedent view.

### Shared wording and, now, a shared view

`AuditPresentationLabels` is now the single definition of how auditing results
are worded: status, coverage, kind, and classification labels, count formatting,
worksheet-qualified locations, and evidence spans. Both presentations read from
it, so they cannot describe the same state in different words. This was extracted
**additively**: `DirectPrecedentReport` kept its entire public surface, and the
ten WP-2-01 presentation tests were not modified, so their passing is genuine
proof that precedent behavior did not change.

The view was duplicated at first, on purpose: the reports are protected by
exacting unit tests, but the view runtime had none, because its real logic is
WinForms over COM and the test project cannot reference either. Extracting a
shared view before both were proven would have moved untested lifecycle code with
no net under it.

Both are now shared. `TraceResultPresentation` is the display-ready shape every
auditing result projects into, and one `TraceViewRuntime` renders it, so the
lifecycle exists in exactly one place. The lifecycle decisions moved into
`TraceViewSession` in the Application layer, which the test project **can**
reference, so the state machine now has real unit coverage for the first time:
open, closed, unverifiable, probe failure, recovery, re-presentation, discard,
reentrancy, and empty-workbook cases. The view is left as a renderer that does
what the session decides.

## Registration

- `audit.dependents.direct` is registered in `AuditingCommandCatalog`, joined into
  `BuiltInCommandRegistry`, and routed through the central `CommandDispatcher`.
  It is read-only, declares no changed property, has no undo policy, and declares
  `PreviewPolicy.Threshold`.
- Command Search reports it unavailable, with a reason, unless the selection is a
  single area.
- The Ribbon exposes it in the `Audit` menu on KeyTip route `Alt, X, A, A, DD`,
  which does not collide with the precedent route `Alt, X, A, A, PD`.
  Registry-wide keyboard-route uniqueness is asserted by an existing test.

## Coverage-gap accounting

Any formula that cannot be fully resolved within qualified parser coverage is
counted as a gap, and any gap blocks a completeness claim. Resolvable edges from
a partially covered formula are still retained, so a gap reduces the claim
without discarding real edges.

Gaps are counted per **cause**, not per inspect-only disposition. A limitation
counts as a gap only when it could conceal a reference to an in-scope cell:

| Cause | Gap? | Why |
|---|---|---|
| Structured reference | yes | Resolves to real cells but produces no parsed reference at all - a genuine blind spot. |
| Dynamic array / implicit intersection | yes | The extent cannot be known without evaluating. |
| Intersection | yes | Operands are parsed, but the cells actually read are only their overlap, so treating the operands as read would over-report. |
| External reference | no | Addresses another workbook; it can never name a cell in this scope. |
| Union | no | Each operand is parsed and each is genuinely read, so recording them is exact. |
| Defined name | no | Resolved, or counted as unresolved, by this index directly. |

This required a parser change, because `FormulaSyntaxDocument` previously exposed
only the **first** limitation in a fixed precedence order - and external is
checked before intersection, so "external" never ruled out an intersection also
being present. `LimitationCodes` now carries every cause, added additively with
`LimitationCode` still returning the same first-match value.

The effect is not cosmetic. External workbook links are common in the models this
add-in targets, and the previous reading made a worksheet containing any external
link permanently unable to claim completeness. The live smoke worksheet moved
from `Partial|B200,C200|16|1` to **`Complete|B200,C200|16|0`** on exactly this
change.

Precedent analysis deliberately keeps the stricter rule. For a precedent trace an
external reference *is* a real edge whose contents cannot be classified, so it
correctly forces a partial result; for a dependent scan the same reference is
simply irrelevant.

## Defect found and fixed

`DirectPrecedentAnalyzer` rendered A1 column names with a loop that dropped the
remainder before dividing. Every column that is an exact multiple of 26 was
wrong: **Z became AZ, AZ became BZ, and ZZ became AAZ**.

This shipped in WP-2-01 (merge `9bf3fb6`). Its effect was not cosmetic. The
capture plan requested the shifted address from Excel, so a precedent in column
Z was classified from the contents of column AZ and displayed at the wrong
address. It escaped the WP-2-01 suite because every fixture and the hidden-Excel
smoke used low columns.

Both the analyzer and the new index now share `AuditAddress.Cell`, which is
verified against known values at 1, 25, 26, 27, 51, 52, 53, 702, 703, and 16384,
and by a round-trip through `AuditAddress.TryParse`. Two regression tests assert
that `=Z1` resolves to `Z1` in both the analysis result and the capture plan.

`AuditNameCandidates` likewise now holds the single definition of "this
identifier token is a defined-name reference", so precedent and dependent
analysis cannot drift apart.

## Verification

- Release and Debug builds: **zero warnings, zero errors**.
- Release tests: **417 passed**, zero failed.
- New coverage includes cell and range dependents with retained evidence,
  partial range overlap, name-bound dependents, unbound names, inspect-only
  formulas that keep their resolvable edge, spill references, external
  references, out-of-scope formulas, workbook-scope refusal, out-of-scope target
  refusal, unsupported target notation, cap truncation, deterministic ordering
  across build order, circular self-reference, and null-argument rejection.
- **AC-AUD-007 equivalence:** an independent brute-force oracle in the test suite
  expands every reference into an explicit set of cell addresses and tests set
  intersection, deliberately avoiding the production rectangle arithmetic. The
  indexed results match the oracle across a twelve-formula corpus and eleven
  targets, including the previously broken multiple-of-26 columns.
- Scan coverage adds exact band tiling with no overlap or gap, the block ceiling
  holding for every band, inflated used-range refusal without reading a block,
  over-wide region refusal, out-of-grid refusal, the region ceiling at its exact
  boundary, empty-worksheet completion, monotonic progress through all three
  phases, cancellation mid-scan and before any read, and scope that follows the
  target worksheet and is never widened.
- The packed Debug XLL passed the hidden-Excel smoke: exact precedent
  classification, full precedent-view lifecycle, workbook closed, and Excel
  exited naturally with no surviving process.
- **Real-Excel dependent scan:** against a live worksheet the scan returned
  `Complete|B200,C200|16|0|Completed`. It found exactly the two direct dependents
  of `A200` and correctly excluded `D200`, which depends on `B200` rather than on
  the target, scanned 16 formulas across the worksheet with no coverage gap, and
  finished in the `Completed` progress phase. The selection and every fixture
  formula were unchanged. Before per-cause gap accounting this same worksheet
  reported `Partial|...|1`, because of the external-reference formula an earlier
  Phase 1B smoke step leaves on the sheet.
- **Real-Excel cancellation:** a pre-cancelled scan through the same adapter
  returned `Refused|AUDIT_SCAN_CANCELLED|0`, proving the wiring fails closed end
  to end rather than only in unit tests.
- Presentation and preview coverage adds claimed and withheld completeness, named
  coverage gaps, multi-kind dependents retaining every edge, refusal rendering,
  repeat determinism, shared wording between the two presentations, the
  registered descriptor's read-only threshold-preview contract, and the preview
  gate below, above, and exactly at its threshold boundary.
- **Registered route in real Excel:** invoking `audit.dependents.direct` through
  the central dispatcher opened the read-only view (`open|success`), preserved
  the selection, and released on the explicit close path (`closed`).
- The ten WP-2-01 presentation tests were **not modified** while the shared label
  extraction landed, so their passing proves precedent behavior is unchanged.

## Retained limitations

- Worksheet scope in this package. Workbook scope followed in WP-2-02b.
- The trace view lifecycle is now shared and unit-tested, but the WinForms
  rendering itself is still covered only by the hidden-Excel smoke.
- A worksheet whose used region exceeds 250,000 cells, or spans more than 10,000
  columns, is refused outright rather than partially scanned. Both are deliberate
  fail-closed bounds, not partial results.
- Gaps are now counted per cause. Structured references, dynamic arrays, and
  intersections remain gaps by design; refining the intersection case would need
  the parser to model the overlap rather than the operands.
- The scan has no performance corpus yet. AC-AUD-009 responsiveness and bounded
  resources are demonstrated by the ceilings, the band tiling, and the live smoke,
  not by a measured large-worksheet workload. A dependent-scan workload belongs in
  WP-2-09 or alongside WP-2-02b.

## Next slice

WP-2-02a is complete. The next work is either:

1. **WP-2-03**, indirect traversal, cycles, caps, and trace navigation. The
   shared trace view is in place, so a traversal result only needs to project
   into `TraceResultPresentation`.
2. **WP-2-02b**, workbook scope, which stays blocked on the workbook-scale
   performance gate recorded in the implementation plan and needs a
   dependent-scan performance corpus.

No Excel trace arrows or workbook annotations are permitted.
