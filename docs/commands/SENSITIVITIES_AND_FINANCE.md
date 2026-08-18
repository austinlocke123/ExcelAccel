# Sensitivity, circularity, and finance-template command contracts

Status: **Draft for review**  
Capabilities: CAP-SENS-001, CAP-CIRC-001, CAP-TPL-001  
Earliest phase: individually gated after Phase 2

## 1. Sensitivity-table boundary

- Initial support uses native Excel What-If Data Tables only.
- A formula-grid fallback is not approved.
- Inputs, output, axes, destination, labels, and overwrite range are explicit.
- The add-in never silently changes calculation mode, iterative calculation, or
  model assumptions.
- Creation is high impact and cannot ship until rollback and session-undo
  behavior for native Data Tables is proven on the support matrix.

## `sensitivity.one_way.create`

- Version: 1
- Impact: high
- Parameters:
  - one explicit output cell;
  - one explicit model input cell;
  - explicit axis values/range and orientation;
  - explicit top-left destination;
  - optional labels and approved formatting recipe.
- Supported context: one workbook; qualified calculation state; unprotected
  destination; no spill/array/table overlap; input/output/destination do not
  overlap unsafely; destination fits worksheet bounds
- Snapshot: output/input identity, axis values, full destination before-state,
  calculation/iteration/AutoSave/coauthoring state, and affected formatting
- Plan: exact destination rectangle, overwritten cells/properties, native Data
  Table row/column input binding, labels, formula anchor, formatting, cost, and
  undo eligibility
- Preview: mandatory; show every reference, dimensions, overwrite counts,
  calculation state, warnings, and representative axis values
- Execute: revalidate all references/preconditions; create the native Data Table
  through the qualified adapter; apply only approved formatting; verify table
  identity/dimensions and restore application state
- Undo: qualified session receipt must remove the created table and restore the
  exact destination before-state only when the current created table still
  matches the receipt
- Failure: rollback created content and formatting; if rollback is incomplete,
  report exact remaining range and block further mutation
- Performance: capability-specific baseline required before approval
- Acceptance: AC-SENS-001 through AC-SENS-009

## `sensitivity.two_way.create`

- Version: 1
- Impact: high
- Parameters: output cell, explicit row input cell/axis, explicit column input
  cell/axis, destination, labels, and formatting
- Additional validation: row and column inputs are distinct unless an explicitly
  qualified case allows otherwise; both axes are nonempty and within caps
- Plan/preview/execute/undo/failure: same as one-way, with complete two-axis
  bindings and destination matrix
- Acceptance: AC-SENS-001 through AC-SENS-012

## `sensitivity.inspect`

- Version: 1
- Impact: read_only
- Parameters: selected native Data Table
- Output: source formula anchor, row/column input bindings when Excel exposes
  them, bounds, calculation state, and coverage limitations
- No recalculation or mutation is triggered solely for inspection.
- Acceptance: AC-SENS-013, AC-SENS-014

## 2. Circularity tools

## `circularity.inspect`

- Version: 1
- Impact: read_only
- Parameters: active workbook and explicit scope/depth cap
- Snapshot: Excel-exposed circular references, iterative calculation settings,
  and qualified formula graph needed for bounded cycle evidence
- Plan: stable cycles/components with coverage/truncation and declared switch
  configuration matches
- Execute: populate read-only pane; no setting or formula change
- Acceptance: AC-CIRC-001 through AC-CIRC-005

## `circularity.switch.preview`

- Version: 1
- Impact: read_only planning
- Parameters: versioned switch template, explicit switch cell, participating
  formula cells, on/off semantics, and typed input mappings
- Plan: exact proposed formulas/values/labels/formats and calculation setting
  implications; never infer participating cells
- Execute: return mandatory preview only
- Acceptance: AC-CIRC-006 through AC-CIRC-008

## `circularity.switch.insert`

- Version: 1
- Impact: high
- Parameters: exact preview plan hash
- CanExecute: revalidate all source formulas, destination cells, calculation
  settings, array/spill/table/protection, and AutoSave/coauthoring policy
- Changed properties: only formulas/values/labels/formats explicitly shown in
  the plan; iterative settings are not changed by this command
- Execute: bounded exact-plan mutation and postcondition verification
- Undo: qualified formula/value session receipt required before release
- Failure: all-or-rollback; no partially wired switch reported as success
- Acceptance: AC-CIRC-006 through AC-CIRC-012

## `circularity.iteration_settings.set`

- Version: 1
- Impact: high
- Parameters: explicit enabled flag, maximum iterations, and maximum change
- Preview: mandatory current/proposed application/workbook-visible effects and
  restoration policy
- Execute: only if the supported Excel build exposes qualified semantics and
  scope; save nothing automatically
- Undo: restore exact prior settings in-session after optimistic validation
- This command remains deferred even if switch insertion ships, until its scope
  and cross-workbook effects are proven.
- Acceptance: AC-CIRC-013 through AC-CIRC-016

## 3. Declarative finance-template system

Templates are versioned content interpreted by one engine, not separate hard-
coded command classes.

### Required template record

- stable template ID/version and localizable name/help;
- typed input slots with workbook/range/single-cell constraints;
- typed output slots and exact destination shape;
- invariant formula templates expressed in the accepted formula AST/model;
- optional labels and formatting recipe references;
- preconditions, impact, overwrite policy, and parser coverage;
- independent finance golden cases with cited convention/version;
- no arbitrary code, macros, callbacks, loops, workbook-content branching, or
  network access.

## `finance.template.preview`

- Version: 1
- Impact: read_only planning
- Parameters: template ID/version, explicit slot bindings, destination, and
  approved options
- Plan: exact formulas, labels, references, formatting, overwrites, assumptions
  supplied by the user, and golden-case/template version
- Validation: every required slot is bound and type/shape compatible; the engine
  does not invent assumptions or locate model inputs heuristically
- Acceptance: AC-TPL-001 through AC-TPL-006

## `finance.template.insert`

- Version: 1
- Impact: medium or high according to template destination/overwrite
- Parameters: exact preview plan hash
- Changed properties: exact formulas, values, labels, and formats declared in
  the template plan
- Preview: mandatory
- Execute: revalidate bindings/destination and apply exact plan
- Undo: qualified session receipt required
- Failure: all-or-rollback; never insert a partial finance schedule
- Acceptance: AC-TPL-001 through AC-TPL-010

## `finance.template.library.import`

- Version: 1
- Impact: medium local-settings/content mutation
- Parameters: explicit local declarative package
- Plan: schema/license/source manifest, template IDs/versions, formula coverage,
  conflicts, golden-test metadata, and prohibited-content scan
- Execute: atomic local import after validation; does not touch a workbook
- Acceptance: AC-TPL-011 through AC-TPL-014, AC-SEC-003

## Proposed built-in template content

The engine may ship templates for MOIC, IRR/XIRR, leverage, sources and uses,
purchase price, minimum cash, management rollover, entry/exit bridge, cash
interest, PIK interest, debt sweep, and fee amortization only after each has:

1. an approved typed slot schema;
2. an explicit finance convention and formula definition;
3. independent golden cases;
4. overwrite/format behavior;
5. its own content review.

The names above are candidate content, not permission for an agent to invent a
house convention.
