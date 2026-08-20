# ExcelAccel requirements

Status: **Draft for review**  
Source baseline: original PRD v1.0 plus `PRD_Review_Feedback.md` (the review is intentionally untracked and not present in a fresh clone)

## 1. Purpose

ExcelAccel is a native Windows Excel add-in that accelerates recurring financial
modeling, formatting, navigation, formula review, and workbook inspection. It
is local-first, deterministic, keyboard-first, and designed to operate safely
inside Excel's process.

This is an engineering requirements specification, not a pricing, market, or
go-to-market plan.

## 2. Non-negotiable invariants

| ID | Requirement |
|---|---|
| INV-001 | Core commands MUST remain usable with outbound network access blocked. |
| INV-002 | Given equivalent supported workbook state, selection, settings, and Excel behavior, command planning MUST produce a canonically equivalent plan. |
| INV-003 | A command MUST change only the properties declared by its contract. |
| INV-004 | A formula-sensitive transformation MUST use parsed semantics or refuse; raw reference-sensitive string replacement is prohibited. |
| INV-005 | A mutation MUST apply the exact approved plan or stop and require replanning. |
| INV-006 | No add-in exception may escape an add-in callback boundary into Excel. |
| INV-007 | Excel COM and C API access MUST occur only through the approved host/adapter boundary and on the owning Excel thread. |
| INV-008 | COM proxies MUST NOT cross a worker-thread boundary or appear in domain-layer APIs. |
| INV-009 | A command MUST restore owned Excel application state on success, refusal, cancellation, and failure. |
| INV-010 | A command MUST NOT silently save, close, upload, email, publish, disable AutoSave, change calculation policy, or broaden its target scope. |
| INV-011 | Workbook-derived content MUST remain local unless the user explicitly exports it. |
| INV-012 | Unsupported, ambiguous, stale, protected, or unsafe state MUST cause a specific refusal rather than a best guess. |

`Deterministic` means equivalent inputs produce equivalent outputs. It does not
mean heuristic-free. A heuristic MAY be used when it is reproducible,
documented, bounded, and exposes uncertainty instead of silently asserting an
exact answer.

## 3. Supported operating envelope

| ID | Requirement |
|---|---|
| PLAT-001 | Windows desktop Excel is the only target. |
| PLAT-002 | x64 Excel is required. x86 is deferred unless separately approved and qualified. |
| PLAT-003 | Supported workbook formats are `.xlsx`, `.xlsm`, `.xltx`, and `.xltm`; existing VBA MUST remain enabled and unmodified. |
| PLAT-004 | The exact Microsoft 365/LTSC build matrix is a Phase 0 decision. |
| PLAT-005 | Dynamic-array support and the `Formula`/`Formula2` compatibility policy are Phase 0 decisions and MUST be explicit before formula mutation ships. |
| PLAT-006 | Protected, read-only, AutoSave, coauthored, edit-mode, multi-area, merged, table, array, and spilled-array states MUST have documented behavior per command class. |

## 4. Capability scope and sequencing

The phases express dependency and risk order, not calendar commitments.
Detailed user-visible behavior for every retained feature family is defined in
the [`commands/` contract index](commands/README.md). A capability listed here
without an approved phase remains gated even though its contract is reviewable.

### 4.1 Phase 0: architecture and risk retirement

- CAP-FOUND-001: Excel-DNA host proof, lifecycle, Ribbon callback, task-pane
  proof, and clean unload/disable behavior.
- CAP-FOUND-002: runtime and deployment selection on a clean reference VM.
- CAP-FOUND-003: command lifecycle vertical slice covering validation,
  snapshot, immutable plan, optional preview, mutation, receipt, result, and
  state restoration.
- CAP-FOUND-004: Excel-thread dispatcher, immutable snapshot boundary, COM
  proxy detector, and fault-injection harness.
- CAP-FOUND-005: formula parser/evaluator-adapter coverage spike with explicit
  supported and refused syntax.
- CAP-FOUND-006: compatibility and performance corpus with repeatable baselines.
- CAP-FOUND-007: AutoSave/coauthoring detection and conservative transaction
  policy proof.
- CAP-FOUND-008: signed packaging/install/update feasibility proof.

Phase 0 produces evidence and accepted ADRs; it does not produce a broadly
featured release.

Phase 0 closure on 2026-08-19 separates permission to build the production
foundation from permission to enable a gated capability or distribute the
add-in. Formula mutation, workbook-scale performance, unqualified collaborative
mutation, and signed clean-machine distribution retain their individual gates
under Acceptance §2.1.

### 4.2 Phase 1A: reliable daily-speed core

- CAP-CMD-001: unified command registry and contract.
- CAP-KEY-001: assignable Quick Keys, collision detection, reserved-key
  warnings, multi-stroke timeout, and Escape cancellation.
- CAP-PROF-001: versioned local profile with atomic load/save and deterministic
  effective-setting resolution.
- CAP-FMT-001: property-scoped formatting cycles and explicit formatting
  commands.
- CAP-FMT-002: AutoColor for selection and worksheet.
- CAP-NAV-001: worksheet and range navigation plus session navigation history.
- CAP-UNDO-001: bounded, session-only, property-scoped add-in undo.
- CAP-UX-001: non-modal status, actionable refusals, keyboard focus
  restoration, accessibility foundations, and no continuous scanning.

### 4.3 Phase 1B: expanded daily-speed core

- CAP-SEARCH-001: local command search.
- CAP-FAV-001: favorites/command bar.
- CAP-STYLE-001: named style library and capture/apply formatting-only styles.
- CAP-FORM-001: Smart Copy, formula spacing, transpose, IFERROR toggle, sign,
  and unit-scale transformations within qualified parser coverage.
- CAP-PROF-002: offline profile import/export with preview and validation.
- CAP-DATA-001: explicit deterministic cleaning commands.
- CAP-SELECT-001: deterministic selection tools for formulas, constants, blanks,
  errors, numeric hardcodes, and external-reference formulas.

### 4.4 Phase 2: inspection and model checking

- CAP-AUD-001: precedents/dependents, trace navigation, and return history.
- CAP-AUD-002: deterministic structural Formula Inspector.
- CAP-CHECK-001: rule-based Model Check with evidence and configurable rules.
- CAP-CHECK-002: local ignore fingerprints and explicit export/import format if
  later approved; no hidden reviewer workflow.

### 4.5 Later, gated capabilities

- CAP-NAME-001: read-only named-range inventory, search, navigation, and
  diagnostics.
- CAP-LINK-001: read-only external-link inventory and navigation.
- CAP-SENS-001: one-way/two-way sensitivity builder.
- CAP-CIRC-001: explicit circularity switch tools.
- CAP-TPL-001: typed, declarative finance formula templates.
- CAP-CHART-001: scoped native chart formatting.
- CAP-PPT-001: one-time PowerPoint snapshot through a separately qualified COM
  adapter.
- CAP-CMP-001: same-shape, read-only workbook/range comparison.
- CAP-STRUCT-001: explicit row/column visibility, grouping, Smart Hide, and
  separately gated structural insert/delete operations.
- CAP-FMT-003: worksheet/workbook formatting recipes and workbook AutoColor,
  gated by workbook-scale transaction and performance evidence.

### 4.6 Explicitly deferred

- structural workbook alignment and inferred row/column matching;
- named-range rename/delete with reference rewriting;
- external-link repointing and breaking;
- row/column insert/delete execution until a structural transaction ADR is
  accepted;
- print/view commands beyond the explicitly contracted gridline, zoom, and
  freeze-pane commands;
- persistent or crash-recoverable undo journals;
- workbook-persisted bookmarks;
- command chains;
- live Excel-PowerPoint links;
- proprietary chart, spreadsheet, scenario, or version-control engines;
- formula-quality scores, workbook-health scores, or opaque judgments;
- AI features in v1 and any nondeterministic workbook mutation in any phase;
- cloud processing or a server dependency in core command execution.

Moving a deferred capability into an active phase requires requirements,
architecture, reliability, command, and acceptance review.

## 5. Command behavior requirements

| ID | Requirement |
|---|---|
| CMD-001 | Every user-visible action MUST have one stable command ID and one registered implementation. |
| CMD-002 | Invocation through shortcut, Ribbon, search, favorite, or pane MUST enter the same command lifecycle. |
| CMD-003 | `CanExecute` MUST be fast, side-effect free, and return a specific refusal and remediation. |
| CMD-004 | `Plan` MUST be immutable, serializable in canonical form for testing, and contain targets, intended property changes, warnings, cost, and undo policy. |
| CMD-005 | Medium/high-impact commands MUST validate that relevant workbook state still matches the plan before mutation. |
| CMD-006 | High-impact commands MUST present a deterministic preview and require confirmation. |
| CMD-007 | Execution MUST report one of `Success`, `SuccessWithWarnings`, `Cancelled`, `Refused`, `RolledBack`, or `Failed`; success cannot hide incomplete planned work. |
| CMD-008 | A fixed-parameter command MUST be shortcut-assignable. A parameterized command MUST be fully reachable from the keyboard. |
| CMD-009 | Formula-to-value conversion is permitted only for commands whose contract explicitly declares it. The active phase contains no such command by default. |

## 6. Reliability requirements

Detailed rules are normative in [`RELIABILITY.md`](RELIABILITY.md).

| ID | Requirement |
|---|---|
| REL-001 | The qualification suite MUST record zero add-in-attributable Excel process terminations. |
| REL-002 | Every external callback, event handler, timer callback, background completion, and COM entry point MUST have an exception boundary. |
| REL-003 | Startup MUST perform only bounded essential work; feature loading and indexing MUST be lazy. |
| REL-004 | Reentrancy, event recursion, stale plans, and shutdown races MUST be explicitly guarded. |
| REL-005 | On an unclean prior session, the add-in MUST start in a conservative recovery mode and MUST NOT automatically mutate or reopen a workbook. |
| REL-006 | A failing optional feature MUST be locally quarantinable without disabling the whole add-in. |
| REL-007 | Resource usage, queues, caches, history, logs, and background work MUST be bounded. |

Zero crashes is a release objective and qualification gate, not a claim that any
in-process extension can mathematically guarantee that Excel will never fail.
The design MUST eliminate known add-in-controlled crash paths and refuse unsafe
work instead of risking process integrity.

## 7. Responsiveness and performance requirements

Targets are provisional until Phase 0 establishes the reference machine,
workbook corpus, cold/warm protocol, and benchmark variance.

| ID | Operation | Reference workload | Provisional P95 | Required behavior |
|---|---|---|---|---|
| PERF-001 | Added startup cost | panes closed | 750 ms | No modal UI; lazy nonessential work. |
| PERF-002 | Immediate command | up to 1,000 cells | 100 ms | No spinner. |
| PERF-003 | Batched mutation | 10,000 cells | 500 ms | No cell-by-cell repaint. |
| PERF-004 | Large scan | 100,000 used cells | 3 s | Progress after 500 ms; cancellable before commit. |
| PERF-005 | Workbook scan | 250,000 cells/20 sheets | 8 s | Responsive progress and bounded snapshot. |
| PERF-006 | Incremental memory | 250,000-cell workbook | 250 MB | Bounded retention and explicit release. |

- PERF-007: no operation may block the Excel UI thread for more than 100 ms
  without yielding or entering an explicit bounded host operation whose behavior
  was qualified in Phase 0.
- PERF-008: large data MUST be read and written in blocks; unapproved
  performance-sensitive cell-by-cell COM loops fail review.
- PERF-009: compute MAY run in the background only on immutable non-COM data.
- PERF-010: benchmark regression gates begin only after baseline variance is
  measured; the initial proposed threshold is a 15% P95 regression.

## 8. Security, privacy, and persistence

| ID | Requirement |
|---|---|
| SEC-001 | Core execution MUST make zero outbound network requests with licensing and updates disabled or stubbed. |
| SEC-002 | Logs MUST exclude formulas, values, workbook/sheet names, paths, user-defined names, images, and chart data by default. |
| SEC-003 | Profile import MUST enforce schema, size, path, content, and version validation and MUST be atomic. |
| SEC-004 | Settings and profile files MUST use least-privilege local storage and atomic replace semantics. |
| SEC-005 | User-approved diagnostics export MUST show a manifest before creation and MUST never transmit automatically. |
| SEC-006 | Updates and installers MUST be signed and rollback-capable before general availability. |

## 9. User-interface requirements relevant to engineering

- UX-001: shortcut handling MUST NOT steal normal typing in Excel edit mode
  unless the command explicitly supports formula editing.
- UX-002: multi-stroke sequences MUST time out or cancel without workbook change.
- UX-003: focus MUST return to the expected workbook object after command UI.
- UX-004: panes MUST be keyboard navigable, high-contrast compatible, and
  usable at 1024x768 effective resolution without clipped primary actions.
- UX-005: panes MUST NOT continuously scan workbooks.
- UX-006: long operations MUST show phase, progress where measurable,
  cancellation availability, and whether mutation has begun.
- UX-007: refusal and failure messages MUST identify the command, safe reason,
  remediation, and local diagnostic ID without exposing workbook content.
- UX-008: display strings MUST be localizable while IDs and serialized values
  remain invariant.

## 10. Non-goals

- Mac, web, and mobile parity;
- financial/market-data services;
- collaboration, issue tracking, model notes, or embedded review status;
- whole-model generation;
- background workbook surveillance;
- hidden workbook mutations for convenience features;
- arbitrary scripts, macros, or executable content in profiles or templates.

## 11. Open review decisions

No implementation agent may resolve these implicitly:

1. Runtime target and packaging model (`ADR-0002`).
2. Minimum Excel build and dynamic-array compatibility matrix.
3. Formula parser/library strategy and v1 syntax coverage (`ADR-0004`).
4. Exact AutoSave and coauthoring policy by impact tier (`ADR-0005`).
5. Phase 0 reference machine, workbook corpus, and benchmark protocol.
6. Signing certificate, installer/update technology, and managed deployment
   qualification path.
