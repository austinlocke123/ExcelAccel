# WP-2-01 reference snapshot and direct-precedent foundation

Date: 2026-08-19

Status: **Complete. The pure-core foundation, bounded Excel capture, read-only
result presentation, and central command registration are implemented and
verified.**

## Contract

- Capability: CAP-AUD-001
- Governing decision: ADR-0004
- Acceptance: AC-AUD-001 through AC-AUD-005
- Allowed implementation: pure Core auditing models/analyzer, deterministic
  tests, and engineering evidence
- Excluded: dependents, indirect traversal, trace navigation, Model Check,
  workbook mutation, automatic external-workbook opening, and parser expansion

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

## Read-only result presentation

- `DirectPrecedentReport` is a pure Core projection of an existing
  `DirectPrecedentResult`. It formats only what the analyzer already
  established: it resolves nothing, evaluates nothing, reorders nothing, and
  reclassifies nothing.
- Every report states the source, workbook, status, scan scope, parser coverage,
  limitation/refusal code, precedent count, unresolved and external edge counts,
  and whether completeness is claimed. A missing captured classification renders
  as `Not captured` with an `Unresolved` state and can never read as resolved.
- A closed external edge renders as `External (closed; never opened)`. The view
  never opens, contacts, or resolves an external workbook.
- Deduplicated nodes show their edge count and every retained source-edge span,
  so equivalent references remain one row without losing evidence.
- A refusal is presented in the same view with its categorized code and no
  precedent rows, so a refusal can never read as a complete trace.
- `DirectPrecedentView` is a modeless read-only WinForms view with accessible
  names on every control. It writes nothing to the workbook, uses no Excel trace
  arrow or workbook annotation, and holds the analysis in memory only. Nothing
  beyond status and counts reaches the diagnostic log.

## Cleanup path

The view is discarded on three explicit paths:

- the user closes it (Close button or Esc);
- `AddInLifecycle.AutoClose` calls `PrecedentViewRuntime.Reset()` on unload;
- the source workbook is no longer open.

Workbook-close cleanup is a read-only probe rather than a COM event handler. No
Excel event is subscribed anywhere in the add-in, and this package did not
introduce the first one. `ExcelReferenceSnapshotAdapter.Probe` enumerates only
already-open workbooks through the existing Excel-thread and COM-retry
boundaries and never opens one. The view re-probes on activation and on the
explicit revalidation path. A `Closed` result discards the analysis and closes
the view; a transient COM failure reports `Unknown`, which keeps the view open
under an explicit "state could not be verified" notice rather than claiming the
source is live. The limitation is deliberate: a closed source workbook is
detected when the view is next activated or revalidated, not at the instant of
closing.

## Registration

- `audit.precedents.direct` is registered in `AuditingCommandCatalog`, joined
  into `BuiltInCommandRegistry`, and routed through the central
  `CommandDispatcher`. It is read-only, declares no changed property, and has no
  undo or preview policy.
- Command Search reports it as unavailable, with a reason, unless the selection
  is exactly one single-area cell.
- The Ribbon exposes it under a new `Audit` menu with KeyTip route
  `Alt, X, A, A, PD`. The route does not collide with any existing control:
  `A` was unused among the group's KeyTips, and registry-wide keyboard-route
  uniqueness is asserted by `RegistryEntriesContainCompleteReleaseMetadata`.

## Verification

- Release tests: **315 passed**, zero failed.
- New golden coverage includes local and quoted-sheet cells/ranges,
  formula/value/error/mixed classification, equivalent-reference deduplication,
  source spans and edge kinds, worksheet-name precedence, unresolved names,
  closed external references, structured-reference partial results, non-formula
  and invalid-formula refusal, R1C1 refusal, defensive copying, missing capture
  classification, and repeat determinism.
- Presentation coverage adds complete/partial/refused projection, claimed and
  withheld completeness, retained deduplicated edge spans, closed-external
  labelling, unresolved names, missing capture classification, named parser
  coverage gaps, repeat determinism, and the registered descriptor's read-only
  contract and acceptance IDs.
- The packed Debug XLL passed the hidden-Excel smoke with exact local-cell and
  workbook-name precedent classification. The adapter preserved the selection
  and workbook contents, closed the workbook, released the XLL, and Excel exited
  naturally with no surviving process.
- The smoke also proved the view lifecycle end to end: it opened on the
  registered dispatcher route without touching the selection or workbook
  contents (`open|success`), retained a result whose workbook is still open
  (`retained|open`), discarded the analysis and closed the view after its source
  workbook was closed (`discarded|closed`), and released on the explicit close
  path (`closed`). The bounded Phase 1B feature suite measured **1,895 ms**.

## Excel capture boundary

- `DirectPrecedentCoordinator` captures one source formula, creates the bounded
  pure-Core capture plan, requests only declared local targets/names, revalidates
  the exact source formula, and publishes no analysis if it changed.
- `ExcelReferenceSnapshotAdapter` runs through the existing Excel-thread and COM
  retry boundaries. It reads exact target ranges without selecting them and has
  a 10,000-cell aggregate capture ceiling.
- Simple worksheet/workbook names are resolved from their local `RefersTo`
  definition without opening external workbooks. Unsupported, missing,
  multi-target, external, or oversized definitions remain explicit unresolved
  edges and therefore cannot produce a completeness claim.

## Defect found after merge

A1 column-name rendering was wrong for every column that is an exact multiple of
26: Z became AZ, AZ became BZ, ZZ became AAZ. The capture plan therefore
requested the shifted address from Excel, so a precedent in column Z was
classified from the contents of column AZ and displayed at the wrong address.
It escaped this package's suite because every fixture and the smoke used low
columns. Found and fixed during WP-2-02; see
[`WP-2-02_DIRECT_DEPENDENTS.md`](WP-2-02_DIRECT_DEPENDENTS.md). Address
formatting now lives in the shared `AuditAddress` with value-verified coverage
and two regression tests.

## Retained limitations

- Results are a point-in-time capture. Nothing refreshes them automatically, and
  a stale source is discarded rather than silently re-analyzed.
- A closed source workbook is detected on view activation or explicit
  revalidation, not at the instant of closing.
- Trace navigation from a result row is deliberately absent; it belongs to
  WP-2-03 and requires its own target revalidation.

## Next slice

WP-2-02: direct dependents and the bounded reverse index.
