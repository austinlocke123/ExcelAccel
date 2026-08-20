# WP-2-02 direct dependents and bounded reverse index

Date: 2026-08-19

Status: **WP-2-02a pure-core reverse index implemented; the Excel scan boundary,
progress/cancellation wiring, and result presentation remain in progress**

## Contract

- Capability: CAP-AUD-001
- Governing decision: ADR-0004
- Acceptance: AC-AUD-006 through AC-AUD-009
- Allowed implementation: pure Core dependent model, reverse index, deterministic
  tests, and engineering evidence
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
- Release tests: **361 passed**, zero failed.
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
- The packed Debug XLL passed the hidden-Excel smoke unchanged after the shared
  address refactor: exact precedent classification, full view lifecycle, workbook
  closed, and Excel exited naturally with no surviving process.

## Next slice

Add the bounded Excel worksheet scan boundary, wire `OperationProgressTracker`
to it for progress and cancellation, and prove the used-range bound. An Excel
worksheet's reported used range is routinely far larger than its real content
because of stray formatting; it must not be trusted as a resource bound. Then
add result presentation and registration.
