# Formula auditing and inspector command contracts

Status: **Draft for review**  
Capabilities: CAP-AUD-001, CAP-AUD-002  
Earliest phase: 2, after ADR-0004 is accepted

## Common boundary

- Auditing commands are read-only with respect to workbook content.
- Results come from an immutable formula/reference snapshot and include explicit
  coverage and truncation metadata.
- Closed external workbooks are never opened automatically.
- Navigation is a separate action from analysis and revalidates its target.
- No command evaluates a selected formula subexpression, generates prose, or
  assigns a complexity/quality score.

## Finding/reference model

Every trace node contains:

- workbook/sheet/address identity for the current session;
- reference kind and direct/indirect depth;
- formula/value/error classification;
- displayed value only when explicitly requested and kept in memory;
- source edge evidence and parser coverage status;
- external/closed/unresolved flags;
- cycle and truncation markers;
- no persisted raw formula/value by default.

## `audit.precedents.direct`

- Version: 1
- Impact: read_only
- Parameters: active formula cell or explicit single-cell target
- Snapshot: target formula plus only references required to resolve direct edges
- Plan: parse direct reference nodes, resolve supported local/workbook names and
  ranges, deduplicate targets, and classify unresolved/external edges
- Execute: populate/focus the trace pane; no Excel arrows or workbook annotations
- Failure: non-formula or unsupported syntax returns categorized refusal; a
  partially resolvable formula may return results only when unresolved edges are
  prominently represented
- Performance: PERF-002 for ordinary formulas; bounded token/reference caps
- Acceptance: AC-AUD-001 through AC-AUD-005

## `audit.dependents.direct`

- Version: 1
- Impact: read_only
- Parameters: active cell/range and explicit scan scope (`worksheet` or
  `workbook`)
- Snapshot: formulas and name definitions within declared scope, captured in
  blocks; no hidden expansion beyond scope
- Plan: build a bounded reverse-reference index using qualified parser coverage
- Preview: required before a workbook scan above the scan threshold
- Execute: populate trace pane after a stable snapshot; findings are labeled by
  coverage and scan scope
- Failure: unsupported formulas are counted as coverage gaps; result cannot
  claim completeness when gaps exist
- Performance: PERF-004/005
- Acceptance: AC-AUD-001, AC-AUD-002, AC-AUD-006 through AC-AUD-009

## `audit.precedents.indirect` and `audit.dependents.indirect`

- Version: 1
- Impact: read_only
- Parameters: explicit maximum depth and result cap within approved bounds
- Plan: deterministic breadth-first traversal, stable node/edge order, visited
  set, cycle detection, and per-depth coverage/truncation
- Execute: show graph/list results only after the requested bounded traversal is
  complete; navigation remains separately revalidated
- Failure: resource cap returns a valid truncated result, never a hang or an
  implicit deeper scan
- Acceptance: AC-AUD-010 through AC-AUD-014

## `audit.trace.navigate`

- Version: 1
- Impact: read_only navigation
- Parameters: exact result node ID
- CanExecute: workbook and target still open/exist; external closed targets are
  non-navigable with a safe reason
- Execute: activate workbook/sheet and select target, then push the prior
  location to audit navigation history
- Acceptance: AC-AUD-015, AC-NAV-005

## `audit.trace.back` and `audit.trace.forward`

- Restore valid audit navigation locations without modifying workbook content.
- Deleted/closed targets are skipped with an explanation.
- This history is session-only and separate from mutation undo.
- Acceptance: AC-AUD-015, AC-NAV-005

## `formula_inspector.open`

- Version: 1
- Impact: read_only
- Parameters: active single formula cell
- Snapshot: exact formula text/dialect and minimal context needed to resolve
  names/reference display
- Plan: immutable syntax tree of functions, operators, constants, references,
  array constructs, and nesting with source spans
- Execute: render a virtualized keyboard-navigable tree; selecting a node does
  not alter formula edit state
- Failure: unsupported syntax shows the exact unsupported span/category and no
  misleading partial tree unless the parser explicitly marks it as partial
- Performance: bounded parse/render time and node count
- Acceptance: AC-AUD-016 through AC-AUD-020

## `formula_inspector.navigate_reference`

- Parameters: exact reference-node ID from the current inspector snapshot
- Plan: resolve supported local reference without evaluating it
- Execute: navigate with return history after revalidation
- External/ambiguous/dynamic references that cannot be resolved safely remain
  visible but non-navigable
- Acceptance: AC-AUD-019 through AC-AUD-021

## Shared safety behavior

- Worksheet/workbook scans do not run continuously or automatically.
- A workbook event invalidates affected cached indexes; stale results are marked
  stale and cannot navigate without revalidation.
- Cancellation during snapshot or analysis exposes no result as complete.
- Result export, if later enabled, is an explicit local command with a reviewed
  manifest and is never automatic.
