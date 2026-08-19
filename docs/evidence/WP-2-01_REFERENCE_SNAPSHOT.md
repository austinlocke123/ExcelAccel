# WP-2-01 reference snapshot and direct-precedent foundation

Date: 2026-08-19

Status: **Pure-core foundation implemented; Excel capture and presentation slice
remain in progress**

## Contract

- Capability: CAP-AUD-001
- Governing decision: ADR-0004
- Acceptance: AC-AUD-001 through AC-AUD-005
- Allowed implementation: pure Core auditing models/analyzer, deterministic
  tests, and engineering evidence
- Excluded: dependents, indirect traversal, trace navigation, Model Check,
  workbook mutation, automatic external-workbook opening, parser expansion, and
  production command registration before the Excel capture boundary exists

## Implemented foundation

- Immutable workbook/sheet/address identities, cell classifications, name
  bindings, bounded snapshot index, source-edge evidence, precedent nodes, and
  complete/partial/refused result states.
- A1-only direct-reference resolution for cells, normalized ranges, quoted local
  sheets, explicit workbook and worksheet name bindings, and closed external
  references.
- Deterministic target ordering and semantic deduplication. Equivalent cell,
  absolute/mixed, sheet-qualified, range-direction, and name-bound references
  share one node while retaining every source edge and edge kind.
- Explicit parser coverage, limitation/refusal codes, scan scope, unresolved and
  external counts, and completeness claims. Missing captured classification,
  external references, names outside the supplied index, and inspect-only syntax
  can never be reported as complete.
- No COM types, workbook writes, network access, persistence, formula evaluation,
  or automatic opening of external workbooks.

The snapshot index is capped at 4,096 cell/range classifications and 4,096 name
bindings. Existing FormulaParser token, length, and nesting ceilings remain the
formula-side resource boundary.

## Verification

- Release tests: **302 passed**, zero failed.
- New golden coverage includes local and quoted-sheet cells/ranges,
  formula/value/error/mixed classification, equivalent-reference deduplication,
  source spans and edge kinds, worksheet-name precedence, unresolved names,
  closed external references, structured-reference partial results, non-formula
  and invalid-formula refusal, R1C1 refusal, defensive copying, missing capture
  classification, and repeat determinism.

## Next slice

Add a bounded Excel main-thread capture adapter and Application coordinator that
capture exactly one formula cell plus the reference classifications/name
bindings requested by this model, revalidate the source formula, and hand the
immutable snapshot to the analyzer. Keep the production Ribbon/search command
unregistered until its read-only result view and cleanup path are available.
