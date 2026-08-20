# WP-2-02 direct dependents and bounded reverse index

Date: 2026-08-19

Status: **WP-2-02a pure-core reverse index and bounded Excel worksheet scan
boundary implemented, with progress and cancellation wired; read-only result
presentation and command registration remain in progress**

## Contract

- Capability: CAP-AUD-001
- Governing decision: ADR-0004
- Acceptance: AC-AUD-006 through AC-AUD-009
- Allowed implementation: pure Core dependent model, reverse index, bounded
  Excel worksheet scan boundary, deterministic tests, and engineering evidence
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

## Coverage-gap accounting

Any formula that cannot be fully resolved within qualified parser coverage is
counted as a gap, and any gap blocks a completeness claim. Resolvable edges from
a partially covered formula are still retained, so a gap reduces the claim
without discarding real edges.

Gaps currently include every inspect-only disposition. Some causes — an external
reference, a union — cannot hide an in-scope edge, so in principle they need not
reduce confidence. The parser surfaces only its **first** coverage limitation in
a fixed precedence order, so a formula reported as external may also contain
another unqualified construct. Counting it as a gap is the only reading that
cannot over-claim. Refining this requires the parser to expose every coverage
cause rather than the first; that is a separate change to the qualified parser
and is deliberately not attempted here.

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
- Release tests: **376 passed**, zero failed.
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
  `Partial|B200,C200|16|1|Completed`. It found exactly the two direct dependents
  of `A200` and correctly excluded `D200`, which depends on `B200` rather than on
  the target. It scanned 16 formulas across the worksheet and reported one
  coverage gap, which is the external-reference formula an earlier Phase 1B smoke
  step leaves on the sheet, and it finished in the `Completed` progress phase.
  The selection and every fixture formula were unchanged.
- **Real-Excel cancellation:** a pre-cancelled scan through the same adapter
  returned `Refused|AUDIT_SCAN_CANCELLED|0`, proving the wiring fails closed end
  to end rather than only in unit tests.

## Retained limitations

- Worksheet scope only. Workbook scope is representable so refusing it is
  explicit and testable, and stays unqualified until WP-2-02b and the
  workbook-scale performance gate are resolved.
- A worksheet whose used region exceeds 250,000 cells, or spans more than 10,000
  columns, is refused outright rather than partially scanned. Both are deliberate
  fail-closed bounds, not partial results.
- Every inspect-only formula counts as a coverage gap, so a worksheet containing
  an external reference cannot currently claim completeness even though an
  external reference cannot hide an in-scope edge. Refining this needs the parser
  to expose every coverage cause rather than the first.
- The scan has no performance corpus yet. AC-AUD-009 responsiveness and bounded
  resources are demonstrated by the ceilings, the band tiling, and the live smoke,
  not by a measured large-worksheet workload. A dependent-scan workload belongs in
  WP-2-09 or alongside WP-2-02b.

## Next slice

Add the read-only dependent result presentation and register
`audit.dependents.direct` through the central dispatcher, Command Search, the
Ribbon `Audit` menu, and a non-conflicting KeyTip. Generalizing the WP-2-01
`DirectPrecedentReport` into a shared trace-result presentation is the natural
move, since WP-2-03 and WP-2-04 need the same shape; that is a public-contract
change to shipped code and should be a deliberate decision rather than a silent
refactor. The AUDITING contract also requires a mandatory preview above a scan
threshold, which is not implemented yet and belongs with registration.

No Excel trace arrows or workbook annotations are permitted.
