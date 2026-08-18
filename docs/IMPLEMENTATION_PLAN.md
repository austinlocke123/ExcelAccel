# Agent-ready implementation plan

Status: **Draft for review; implementation not authorized**

This plan decomposes future work. It does not create permission to code. After
documentation approval, each work package should be assigned separately with a
bounded scope and required evidence.

## 1. Authorization gate

Before any production or spike code is created:

- the documentation baseline is approved;
- Phase 0 scope is accepted;
- proposed ADRs identify their decision owner and evidence plan;
- repository language/runtime/tooling conventions are approved;
- reviewers decide whether Phase 0 spikes live on the main branch or disposable
  branches;
- acceptance evidence storage and CI environment are defined.

## 2. Agent work-package contract

Every implementation assignment MUST include:

- one work-package ID;
- governing requirement, ADR, command, and acceptance IDs;
- explicit allowed files/projects;
- explicit exclusions;
- expected interfaces and artifacts;
- verification commands and evidence;
- a rollback/removal plan for experimental work;
- an instruction to stop on unresolved specification conflict.

Agents MUST NOT expand capability scope, add dependencies, change a public
contract, weaken a refusal, or resolve an open ADR implicitly.

## 3. Phase 0 work packages

### WP-P0-01: host and lifecycle spike

- Inputs: ADR-0001, ADR-0002, AC-P0-001/002
- Prove: Excel-DNA load/callback/pane/disable/unload, exception boundary,
  x64 clean VM, cold/warm startup, coexistence, shutdown cleanup
- Output: evidence report and accepted runtime ADR
- Excludes: feature commands and production UI

### WP-P0-02: test and process harness

- Inputs: REL-001..007, AC-REL-001/002/009
- Prove: launch/drive/close Excel, collect exit/hang evidence, inject managed
  failures, repeat sessions, measure process resources
- Output: repeatable local test commands and evidence schema
- Excludes: telemetry upload

### WP-P0-03: command vertical slice

- Inputs: CMD-001..009, architecture command lifecycle
- Prove: one read-only command and one property-only mutation through complete
  plan/preview/execute/result paths
- Output: interface recommendation, canonical plan fixture, failure matrix
- Excludes: general command catalog implementation

### WP-P0-04: Excel adapter and state guard

- Inputs: INV-006..010, AC-ARCH-003/004, AC-REL-003..008
- Prove: thread dispatcher, block read/write, state ownership, reentrancy guard,
  bounded retry, stale-context refusal, COM boundary enforcement
- Output: accepted adapter/state-guard design and test evidence

### WP-P0-05: formula strategy spike

- Inputs: ADR-0004, AC-P0-005, AC-FORM-001/002
- Prove: candidate parser strategies against the approved coverage/refusal corpus
- Output: accepted parser ADR, coverage matrix, canonical AST/reference model
- Excludes: production transforms, inspector, and Model Check

### WP-P0-06: AutoSave/coauthoring spike

- Inputs: ADR-0005, AC-P0-007, AC-SYNC-001..003
- Prove: detectable states/events, intervening edits, plan invalidation, refusal,
  undo implications across supported builds
- Output: accepted policy and command-context contract

### WP-P0-07: performance baseline

- Inputs: PERF-001..010, AC-P0-006, AC-PERF-001
- Prove: reference machine/corpus, cold/warm protocol, block snapshot/write cost,
  UI responsiveness, benchmark variance
- Output: frozen Phase 1 budgets or reviewed requirement changes

### WP-P0-08: packaging and trust spike

- Inputs: SEC-006, AC-P0-008
- Prove: signed artifact, clean-VM install, load, disable, upgrade, rollback,
  uninstall, trust/allowlisting constraints
- Output: installer/update ADR and qualification procedure

## 4. Phase 1A work packages

These begin only after all applicable Phase 0 gates pass.

| Work package | Scope | Depends on | Primary acceptance |
|---|---|---|---|
| WP-1A-01 | Production solution skeleton and architecture enforcement | P0 ADRs | AC-ARCH-001..004 |
| WP-1A-02 | Command registry, context, plans, results, impact policy | WP-1A-01, WP-P0-03 | AC-CMD-001..008 |
| WP-1A-03 | Excel adapter, dispatcher, state/reentrancy guards | WP-1A-01, WP-P0-04 | AC-REL-002..008 |
| WP-1A-04 | Profiles, schema, atomic storage, default-profile artifact | WP-1A-01 | AC-PROF-001..004 |
| WP-1A-05 | Quick Keys and collision/edit-mode safety | WP-1A-02, WP-1A-03, WP-1A-04 | AC-KEY-001..004 |
| WP-1A-06 | Formatting cycles and explicit formatting commands | WP-1A-02, WP-1A-03, WP-1A-04 | AC-FMT-001..009 |
| WP-1A-07 | Selection and worksheet AutoColor | WP-1A-02, WP-1A-03, WP-1A-04, WP-P0-07 | AC-FMT-010..013 |
| WP-1A-08 | Navigation and session history/bookmarks | WP-1A-02, WP-1A-03 | AC-NAV-001..006 |
| WP-1A-09 | Session undo receipts and optimistic validation | WP-1A-02, WP-1A-03, WP-1A-06, ADR-0003 | AC-REL-011/012 |
| WP-1A-10 | Recovery marker, quarantine, diagnostics export | WP-1A-02, WP-1A-03 | AC-REL-004/010, AC-SEC-002/004 |
| WP-1A-11 | Accessibility, focus, progress, and performance hardening | all above | AC-UX-001..005, AC-PERF-002..006 |
| WP-1A-12 | Installer, update/rollback, full qualification | all above | GA gates |

## 5. Phase 1B work packages

Phase 1B begins only after Phase 1A reliability evidence is reviewed and
ADR-0004 is accepted for formula work.

| Work package | Scope | Depends on | Primary acceptance |
|---|---|---|---|
| WP-1B-01 | Command search, favorites, registry ranking, keyboard/accessibility | WP-1A-02/04/05/11 | AC-SEARCH-001..004, AC-FAV-001..004 |
| WP-1B-02 | Style recipe schema, capture, apply, built-ins, delete | WP-1A-04/06/09 | AC-STYLE-001..008 |
| WP-1B-03 | Profile export/import preview/apply and binding export | WP-1A-04/10 | AC-PROF-005..009, AC-SEC-003/004 |
| WP-1B-04 | Formula reference toggle edit-mode spike and command | ADR-0004, WP-1A-05 | AC-FORM-005..007, AC-KEY-001 |
| WP-1B-05 | Smart Copy and formula spacing | ADR-0004, WP-1A-02/03/09 | AC-FORM-008..012 |
| WP-1B-06 | Formula transpose | WP-1B-05 | AC-FORM-013..016 |
| WP-1B-07 | IFERROR, sign, and unit transforms | WP-1B-05 | AC-FORM-017..025 |
| WP-1B-08 | Formulas/values/formats paste and deterministic fill | WP-1B-05 | AC-FORM-026..034 |
| WP-1B-09 | Text normalization and typed conversion framework | WP-1A-02/03/09 | AC-DATA-001..014 |
| WP-1B-10 | Explicit display-value conversions | WP-1B-09 | AC-DATA-015..019 |
| WP-1B-11 | Deterministic selection tools and multi-area limits | WP-1A-08, WP-1B-05 for parser-dependent predicates | AC-SELECT-001..007 |
| WP-1B-12 | Phase 1B fault, locale, performance, and soak qualification | all Phase 1B packages | Applicable GA gates |

## 6. Phase 2 work packages

| Work package | Scope | Depends on | Primary acceptance |
|---|---|---|---|
| WP-2-01 | Reference snapshot/index and direct precedents | ADR-0004, Phase 1B parser evidence | AC-AUD-001..005 |
| WP-2-02 | Direct dependents and bounded reverse index | WP-2-01 | AC-AUD-006..009 |
| WP-2-03 | Indirect traversal, cycles, caps, trace navigation | WP-2-01/02, WP-1A-08 | AC-AUD-010..015 |
| WP-2-04 | Formula Inspector tree and reference navigation | WP-2-01 | AC-AUD-016..021 |
| WP-2-05 | Model Check engine, finding schema, scan orchestration | WP-2-01/02 | AC-CHECK-001..008 |
| WP-2-06 | Pattern, constant, and embedded-hardcode rules | WP-2-05 | AC-CHECK-009..018 |
| WP-2-07 | Error, external, circular, and number-format rules | WP-2-05 | AC-CHECK-019..028 |
| WP-2-08 | Finding navigation, local ignores, rescan, and export | WP-2-05/06/07 | AC-CHECK-029..037 |
| WP-2-09 | Phase 2 large-corpus, cancellation, privacy, performance, and soak qualification | all Phase 2 packages | Applicable GA gates |

## 7. Individually gated feature work packages

These packages are not one combined phase. Each requires explicit approval after
its prerequisites and capability-specific performance/reliability corpus exist.

| Work package | Scope | Depends on | Primary acceptance |
|---|---|---|---|
| WP-G-01 | Read-only named-range inventory/search/navigation/export | WP-2-01/09 | AC-NAME-001..011 |
| WP-G-02 | Read-only external-link inventory/search/navigation/export | WP-2-01/09 | AC-LINK-001..011 |
| WP-G-03 | Same-shape range/worksheet/workbook compare and export | WP-2-01/04/09 | AC-CMP-001..019 |
| WP-G-04 | Native one/two-way sensitivity creation and inspection | WP-1A-03/09/12, ADR-0005 | AC-SENS-001..014 |
| WP-G-05 | Circularity inspection and switch insertion | WP-2-01/09, ADR-0004/0005 | AC-CIRC-001..012 |
| WP-G-06 | Iterative-calculation settings command | WP-G-05 plus separate side-effect approval | AC-CIRC-013..016 |
| WP-G-07 | Declarative finance-template engine/library import | WP-1B-05/11 | AC-TPL-001..005, AC-TPL-007..013 |
| WP-G-08 | Individually reviewed built-in finance template content | WP-G-07 | AC-TPL-006/014 |
| WP-G-09 | Selected native chart property commands and style recipes | WP-1A-03/09/12 | AC-CHART-001..027 |
| WP-G-10 | Optional PowerPoint image-snapshot adapter and commands | WP-G-09 plus separate COM ADR/spike | AC-PPT-001..009 |
| WP-G-11 | Row/column hide, unhide, group, ungroup, and Smart Hide | WP-1A-03/09/12 | AC-STRUCT-001..011 |
| WP-G-12 | Structural row/column insert/delete | WP-G-11 plus dedicated structural transaction ADR | AC-STRUCT-012..019 |
| WP-G-13 | Sheet/workbook style, AutoFormat, and workbook AutoColor | WP-1A-06/07/09/12 | AC-FMT-014..020 |

Every gated package ends with its own fault-injection, resource-soak,
compatibility, performance, accessibility, and privacy review before it may be
included in a release.

## 8. Pull-request slicing rules

- One architectural mechanism or coherent command family per PR.
- No PR mixes dependency upgrades with feature behavior unless required by an
  accepted ADR.
- Tests and architecture enforcement land with or before production behavior.
- Generated binaries, Office documents, workbooks, and benchmark artifacts have
  explicit repository/storage policy before check-in.
- A PR cannot mark an acceptance criterion complete without linked evidence.
- Reliability failures are fixed before adding the next feature family.

## 9. Definition of done for a work package

- governing docs and accepted ADRs are satisfied;
- public contracts are documented;
- unit, architecture, integration, failure-injection, and performance tests
  required by the package pass;
- no new warning, flaky test, resource leak, unbounded operation, or unexplained
  Excel exit appears;
- relevant traceability rows link to evidence;
- diagnostics contain no seeded sensitive data;
- a human can disable/remove the change and recover the environment;
- remaining limitations and deferrals are explicit.
