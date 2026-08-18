# Same-shape comparison command contracts

Status: **Draft for review**  
Capability: CAP-CMP-001  
Earliest phase: gated after parser and Phase 2 reliability evidence  
Approved scope: read-only comparison of already-open sources with explicit,
same-shape pairing

## Comparison boundary

- Comparison never mutates, recalculates, saves, closes, or opens a workbook.
- Initial sources must already be open in the current Excel instance.
- There is no structural alignment, inferred row/column matching, fuzzy sheet
  pairing, or mechanical-shift classification.
- Shape/structure mismatch is reported and refused, not approximated.
- Formula comparison uses accepted parser coverage and distinguishes exact text,
  canonical semantic structure, and unsupported syntax.

## Difference model

Each difference contains:

- source/target workbook, sheet, and relative position identities;
- category: formula, constant, displayed value when requested, number format,
  approved formatting property, missing/unsupported;
- source and target type;
- canonical formula token/reference delta where supported;
- raw values/formulas only in memory and only when needed for display;
- navigation eligibility and stale markers;
- no claim that an unsupported or unexamined property is equal.

## `compare.ranges.same_shape`

- Version: 1
- Impact: read_only
- Parameters: two explicit open-workbook ranges, comparison categories, and
  value-comparison policy
- CanExecute: rectangular ranges have equal dimensions and are within resource
  caps; source/target are not ambiguous aliases of the same range unless an
  explicit self-comparison is allowed for testing
- Snapshot: requested formulas/values/formats in blocks, plus identity/version
  fingerprints
- Plan: position-by-position deterministic comparison; formula token/reference
  differences only within parser coverage
- Preview: required above workload threshold and shows source/target/dimensions/
  categories
- Execute: populate results only after both stable snapshots and completed
  analysis
- Cancellation: no completed result replaces the previous pane
- Acceptance: AC-CMP-001 through AC-CMP-008

## `compare.worksheets.same_shape`

- Version: 1
- Impact: read_only
- Parameters: two explicit worksheets and exact comparison rectangle policy
  (`used_range_intersection` is prohibited; choose equal explicit ranges or
  require equal approved used-range bounds)
- CanExecute: chosen bounds have the same address dimensions; hidden rows/
  columns and filtered state are included and reported
- Plan/execute: same as range comparison with sheet-level counts
- Failure: different shapes refuse with a summary; no alignment attempt
- Acceptance: AC-CMP-001 through AC-CMP-010

## `compare.workbooks.same_structure`

- Version: 1
- Impact: read_only
- Parameters: two explicit open workbooks, explicit sheet pairing, and
  comparison categories
- CanExecute: each paired sheet satisfies same-shape contract; unpaired sheets
  are reported as structure differences and block cell-level comparison unless
  the user explicitly limits scope to valid pairs
- Pairing: exact user mapping or exact sheet name only; no signatures/fuzzy match
- Execute: compare valid explicit pairs and present unpaired/shape mismatches as
  separate structure results
- Acceptance: AC-CMP-011 through AC-CMP-015

## `compare.result.navigate_source` and `compare.result.navigate_target`

- Revalidate exact result side and select the current target.
- Closed/changed/deleted target refuses or marks result stale.
- Acceptance: AC-CMP-016, AC-NAV-005

## `compare.results.export`

- Version: 1
- Impact: local file write; workbooks remain read-only
- Parameters: destination, approved format, included fields/redaction
- Preview: result count, sources, categories, raw content/path inclusion, and
  destination
- Execute: deterministic temp-write/validate/replace; no transmission
- Acceptance: AC-CMP-017 through AC-CMP-019, AC-SEC-004

## Performance and completeness

- Snapshot, comparison, and rendering timings are recorded separately.
- Proposed qualification workload is two open 250,000-cell workbooks with equal
  structure; its P95 target is frozen before this capability is approved.
- Result count, parser gaps, skipped property categories, and truncation are
  always visible.
- Large result panes are virtualized and bounded.
- Any source change invalidates affected results; navigation revalidates but does
  not pretend stale comparison evidence is current.
