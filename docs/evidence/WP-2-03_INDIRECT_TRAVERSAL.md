# WP-2-03 indirect traversal, cycles, caps, and trace navigation

Date: 2026-08-19

Status: **Complete. The bounded traversal engine, both Excel expansions,
presentation, registration, and trace navigation are implemented and verified.**

## Contract

- Capability: CAP-AUD-001
- Governing decision: ADR-0004
- Acceptance: AC-AUD-010 through AC-AUD-015
- Allowed implementation: pure Core traversal engine, Application expansions and
  coordinator, Excel expansion ports, presentation, registration, trace
  navigation, deterministic tests, and engineering evidence
- Excluded: Formula Inspector, Model Check, workbook-scope dependents, workbook
  mutation, automatic external-workbook opening, and parser expansion

## Bounded traversal

`IndirectTraceEngine` is a deterministic breadth-first traversal over an
expansion port, and it is the only place traversal policy lives.

- **Determinism (AC-AUD-010).** Traversal is breadth-first and preserves the
  expansion's own ordering, so an identical snapshot always yields an identical
  node list, depth, origin, and cycle marking.
- **Cycles (AC-AUD-011).** A node is expanded at most once. Revisiting one is
  represented as a cycle edge rather than expanded again, so a circular model
  terminates. A self-reference is a single cycle edge, and a diamond visits the
  shared node once while still recording both paths.
- **Caps (AC-AUD-012).** Depth and node caps are explicit, hard-ceilinged at 16
  and 5,000, and validated on construction. Reaching either produces an explicit
  truncated `Partial` result with `AUDIT_DEPTH_LIMIT_REACHED` or
  `AUDIT_SCAN_TRUNCATED`, never a hang or a silent omission. A node reached but
  not expanded is marked as an unexpanded frontier so the boundary is visible.
- **Evidence (AC-AUD-013).** Every node retains the direct evidence of the edge
  that reached it, plus the node it was reached from, at every depth.
- **Cancellation (AC-AUD-014).** A cancelled traversal is **refused**, carrying
  `AUDIT_SCAN_CANCELLED` and no nodes, so no partial traversal can be reported as
  a trace.

An incomplete expansion is counted as a coverage gap and blocks the completeness
claim, consistent with the direct analyses.

## Expansions

- `PrecedentTraceExpansion` walks upstream by reading each node's formula and
  reusing the qualified `DirectPrecedentAnalyzer`. A cell holding no formula is a
  **leaf, not a gap** — it genuinely has no precedents. A refused analysis is a
  gap.
- `DependentTraceExpansion` walks downstream against **one** prebuilt
  `ReverseReferenceIndex`, so a traversal never rescans the worksheet per step.
  The worksheet is read once, under the same bounded region plan, threshold
  preview, and ceilings as the direct dependent scan.

`ExcelReferenceSnapshotAdapter` gained `TryReadFormula` for the precedent walk. It
reads one cell through the existing Excel-thread and COM-retry boundaries, writes
nothing, and selects nothing.

## Trace navigation (AC-AUD-015)

Navigation is a separate action from analysis. `NavigationService.GoTo` was added
additively and reuses the existing navigation semantics exactly: it revalidates
the target through the navigation port, selects it, and records the prior
location so session Back returns there. A target that no longer resolves is
refused with `STALE_CONTEXT`. Nothing is written to the workbook.

A row is navigable only when it resolves to a real cell. External, unresolved,
and cycle rows carry no navigation target and do nothing, so navigation can never
be offered for something that cannot be selected.

## Presentation and registration

`IndirectTraceReport` projects into the shared `TraceResultPresentation`, so the
traversal reuses the one trace view rather than adding a third window. Rows show
depth, node, the node reached from, kind, state, and source reference; the
summary names both caps, the deepest depth reached, coverage gaps, truncation by
depth and by node cap, whether a cycle was found, and whether completeness is
claimed.

- `audit.precedents.indirect` on KeyTip route `Alt, X, A, A, PI`, unavailable
  unless exactly one cell is selected.
- `audit.dependents.indirect` on `Alt, X, A, A, DI`, declaring
  `PreviewPolicy.Threshold` and applying the same confirmation gate as the direct
  scan before reading a large worksheet.

Both are read-only, declare no changed property, and have no undo policy.

## Verification

- Release and Debug builds: **zero warnings, zero errors**.
- Release tests: **443 passed**, zero failed.
- Engine coverage includes breadth-first ordering with depth and origin, evidence
  retention at every depth, cycles, self-reference, diamonds, the depth cap and
  its frontier, the node cap, incomplete expansions, external and unresolved
  terminals, cancellation, repeat determinism, a 200-node chain terminating at the
  hard ceiling, and option validation.
- Coordinator coverage includes precedent chains through formula lookups,
  non-formula leaves, circular chains, dependent chains over one index, progress,
  cancellation, projection into the shared view shape, non-navigable cycle rows,
  cap reporting in the summary, and both registered descriptors.
- **Real Excel:** the registered `audit.precedents.indirect` and
  `audit.dependents.indirect` routes each opened their read-only view over a live
  `A210 -> B210 -> C210` chain (`open|success`). Trace navigation selected `C210`
  and recorded return history (`C210|recorded`). Workbook contents were unchanged,
  the view released on explicit close, the workbook closed, and Excel exited
  naturally with no surviving process.

## Retained limitations

- Indirect dependents are worksheet scope only, inheriting the WP-2-02b gate.
- Traversal has no performance corpus. AC-AUD-009-style responsiveness rests on
  the depth and node ceilings and the live smoke, not a measured deep-model
  workload. That belongs in WP-2-09.
- A precedent walk costs one Excel read per expanded node. It is bounded by the
  node cap, but a deep trace on a slow workbook will be correspondingly slower;
  no batching is attempted.
- Trace navigation shares the existing session history rather than keeping a
  separate audit history. `audit.trace.back` and `audit.trace.forward` are
  therefore served by the existing navigation Back and Forward commands.

## Next

WP-2-04 (Formula Inspector) requires a real parse tree inside the qualified
parser; see the implementation plan note. WP-2-05 onward is Model Check.
