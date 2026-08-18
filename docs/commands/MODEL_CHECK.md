# Model Check command and rule contracts

Status: **Draft for review**  
Capabilities: CAP-CHECK-001, CAP-CHECK-002  
Earliest phase: 2, after ADR-0004 is accepted

## Product and safety boundary

Model Check is a deterministic, read-only rule engine. A finding states a rule,
location, evidence, and suggested inspection action. It never declares that a
model is wrong, assigns a health/quality score, modifies a workbook, or creates
review-status workflow.

## Finding model

Each finding contains:

- stable rule ID/version and configured severity;
- current-session workbook/sheet/target identity;
- concise rule statement;
- evidence derived from the scan snapshot;
- peer/baseline context where applicable;
- parser/coverage confidence category, not a correctness probability;
- deterministic local fingerprint for ignore matching;
- navigation eligibility and stale marker;
- no persisted raw value/formula unless explicitly exported by the user.

## Scan commands

### `model_check.run.selection`

- Version: 1
- Impact: read_only
- Parameters: enabled rule IDs/versions and explicit selection
- Snapshot: only properties required by enabled rules
- Plan: validate coverage and estimate workload; stable rule/target evaluation
  order
- Preview: threshold-based
- Execute: run rules on immutable data and replace the pane's prior result set
  only when the scan completes successfully
- Cancellation: prior results remain visible and marked prior; partial results
  are not presented as a completed scan
- Acceptance: AC-CHECK-001 through AC-CHECK-006

### `model_check.run.worksheet`

- Same contract with worksheet used-range coverage and PERF-004.
- Preview required above the qualified cell/formula count.
- Acceptance: AC-CHECK-001 through AC-CHECK-007

### `model_check.run.workbook`

- Same contract with explicit included/excluded sheet inventory and PERF-005.
- Preview always required and shows sheet/cell/formula counts and unsupported
  content categories.
- Acceptance: AC-CHECK-001 through AC-CHECK-008

## Rule contracts

### `check.formula.pattern_inconsistency`

- Detect copied-formula peer regions using deterministic spatial grouping and
  normalized relative-reference structures.
- A finding identifies the peer group, expected normalized pattern, exceptional
  cell, and neighboring evidence.
- Minimum peer count, boundary behavior, blank treatment, and confidence
  categories are profile settings with approved defaults.
- A region with parser coverage gaps cannot claim complete consistency.
- Acceptance: AC-CHECK-009 through AC-CHECK-012

### `check.constant.interrupts_formula_region`

- Detect a constant cell inside a qualified otherwise formula-consistent peer
  region.
- Report separately from a numeric literal embedded inside a formula.
- Blanks, labels, intentional input rows/columns, totals, and boundaries follow
  explicit configuration; no semantic guessing.
- Acceptance: AC-CHECK-013, AC-CHECK-014

### `check.formula.embedded_numeric_constant`

- Traverse numeric literal AST nodes inside formulas.
- Apply versioned literal allowlists and contextual structural exclusions only
  when their rule is explicit.
- Evidence identifies the literal and source span without evaluating intent.
- Acceptance: AC-CHECK-015 through AC-CHECK-018

### `check.formula.error`

- Detect cells containing qualified Excel error values and formulas with broken
  `#REF!` references.
- Classify error kind; do not recalculate the workbook solely for this rule.
- Acceptance: AC-CHECK-019, AC-CHECK-020

### `check.reference.external`

- Detect external workbook references in supported formulas/names within the
  declared coverage matrix.
- Report source token/category and affected target; do not contact or open the
  source.
- Acceptance: AC-CHECK-021, AC-CHECK-022

### `check.reference.circular`

- Use qualified Excel-exposed circular-reference information and/or bounded
  graph-cycle detection from the stable snapshot.
- Distinguish declared configured circularity-switch regions from undeclared
  cycles only through explicit configuration, not inference.
- Acceptance: AC-CHECK-023 through AC-CHECK-025

### `check.format.number_inconsistency`

- Compare normalized number-format identity inside a deterministically defined
  peer region.
- Report the majority/baseline format and exceptions; never change formats.
- Values/formulas do not affect format equivalence except where the configured
  peer-region rule explicitly uses cell type.
- Acceptance: AC-CHECK-026 through AC-CHECK-028

## Result commands

### `model_check.finding.navigate`

- Revalidate exact target identity and select it; stale/missing targets refuse.
- Push prior location to audit navigation history.
- Acceptance: AC-CHECK-029

### `model_check.finding.ignore_local`

- Impact: low local-settings mutation
- Parameters: finding ID and rule-specific fingerprint scope
- Plan: show rule, normalized fingerprint inputs, and whether ignore is local
  profile or separately exported ignore set
- Execute: atomic profile write; rescan required to apply
- Must not persist raw formula/value content.
- Acceptance: AC-CHECK-030 through AC-CHECK-033

### `model_check.finding.unignore_local`

- Remove only the selected fingerprint from the local profile.
- Acceptance: AC-CHECK-031 through AC-CHECK-033

### `model_check.rescan`

- Re-run the exact prior scope and rule configuration only after creating a new
  snapshot; never reuse old findings as current evidence.
- Acceptance: AC-CHECK-034

### `model_check.results.export`

- Impact: explicit local file write
- Parameters: destination and approved format
- Preview: manifest stating included workbook-derived fields and redaction
  choices; default excludes formulas/values
- Execute: deterministic temp-write/validate/replace; never transmit
- Acceptance: AC-CHECK-035 through AC-CHECK-037, AC-SEC-004

## Resource and failure rules

- Rule execution, findings, peer groups, graph edges, and ignores are bounded.
- A rule failure is not silently dropped: the scan returns failed/incomplete
  with the rule ID and safe diagnostic.
- Unsupported coverage is reported per rule and sheet.
- Identical snapshot, rules, configuration, parser version, and profile produce
  canonically identical findings and fingerprints.
