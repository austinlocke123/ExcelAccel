# Agent-ready implementation plan

Status: **Phase 2 is complete. WP-F feature packages are active; see §7b.**

This plan decomposes implementation work. The documentation baseline and Phase
0 start were approved on 2026-08-18. Phase 1A foundation work was authorized on
2026-08-19 under the capability-specific gate disposition in
[`ACCEPTANCE.md`](ACCEPTANCE.md). Gated feature packages still require their
stated prerequisites and evidence.

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

These begin only after all applicable Phase 0 gates pass. “Applicable” follows
the explicit closure disposition in Acceptance §2.1: formula mutation,
workbook-scale performance, unqualified collaboration, and distribution remain
disabled until their retained gates pass.

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

### WP-2-02 scope slicing

AC-AUD-006 makes scan scope an explicitly declared parameter that may never
expand silently, so WP-2-02 is delivered in two scope slices:

- **WP-2-02a, worksheet scope.** The declared scope is one worksheet. This is
  the first package to wire `OperationProgressTracker` to a real operation, and
  it establishes the reverse-index design, the brute-force equivalence oracle
  required by AC-AUD-007, the coverage-gap accounting required by AC-AUD-008,
  and the bounded-resource behavior required by AC-AUD-009.
- **WP-2-02b, workbook scope.** Delivered. The workbook-scale gate was resolved
  in favour of a bounded opening once WP-2-09 supplied the measured corpus; see
  the resolution below.

An Excel worksheet's reported used range is routinely far larger than its real
content because of stray formatting. A dependent scan MUST NOT trust it as a
resource bound; an oversized scan region returns an explicit bounded partial
result rather than a hang.

### Resolved: workbook-scale performance gate

Section 4 previously stated that workbook-scale performance remained disabled
until its retained gate passed, while the retained-gate list in
[`PROJECT_STATUS.md`](PROJECT_STATUS.md) no longer named it. The two documents
disagreed on a point that changed scope, and WP-2-02b was held for a human
decision.

**Resolved on 2026-08-20: opened, bounded.** WP-2-09 measured a worksheet
dependent scan at 200 ms P95 over 10,040 formulas and a Model Check scan at
951 ms over 16,000 cells, both far under budget and both on the slower Debug
build. Workbook scope is therefore permitted under explicit bounds rather than
left indefinitely deferred:

- at most 64 worksheets and 1,000,000 aggregate cells per workbook scan;
- a worksheet that cannot be bounded is excluded with a stated reason rather
  than failing the whole workbook, and an exclusion is a coverage gap that
  blocks any completeness claim;
- a plan with nothing left to read refuses with the first exclusion reason,
  rather than returning an empty result that would read as "nothing found";
- the sheet inventory is **always** confirmed before a workbook scan reads
  anything, whatever its size;
- the scan remains cancellable throughout.

Section 4's blanket statement is superseded for this capability. Workbook-scale
*mutation* remains out of scope; only read-only scanning is opened.

### WP-2-04 depends on parser work the table does not show

The dependency table lists WP-2-04 against WP-2-01 alone, which understates it.
`FormulaSyntaxDocument` exposes a token stream and a flat reference list; there
is no syntax tree, and every shipped formula transform rewrites at token level.
AC-AUD-016 requires an immutable tree of functions, operators, constants,
references, arrays, and nesting. WP-2-04 therefore requires building a real
parse tree inside the qualified parser and is ADR-0004 work, not view-layer
work. It must not be picked up as a cheap package.

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

## 7b. Approved feature work packages

These are approved product changes, not gated capability packages. They came
from using the add-in rather than from the original specification.

| Work package | Scope | Depends on | Primary acceptance | State |
|---|---|---|---|---|
| WP-F-01 | Profile schema 5 to 6: cycles as data, cycle semantics, defaults with parenthesised negatives and `_)` padding, migration | none | AC-FMT-021..025, AC-FMT-031, AC-FMT-043, AC-FMT-045 | |
| WP-F-02 | User-defined cycles: add/remove/reorder, validation, custom cycle slots, Command Search by cycle name, no phantom slots | WP-F-01 | AC-FMT-026..029, AC-FMT-039, AC-FMT-040 | |
| WP-F-03 | Settings editor for cycles, with positive and negative preview per entry | WP-F-01/02 | AC-FMT-030 | |
| WP-F-04 | Ribbon regrouping to the approved task taxonomy | none | AC-FMT-033 | **Delivered 2026-08-20** |
| WP-F-05 | Blue-black input/formula toggle over a selection | WP-F-04 | AC-FMT-034..038 | **Withdrawn** |
| WP-F-06 | `formula.units.to_basis_points` unit transform | none | AC-FMT-032 | |
| WP-F-07 | Ribbon route validator: KeyTip uniqueness, prefix rule, descriptor/ribbon agreement | none | AC-FMT-033 | |
| WP-F-08 | AutoColor classification, planning and gating to the approved precedence | WP-F-01 | AC-FMT-034..038, AC-FMT-041, AC-FMT-042, AC-FMT-046 | |
| WP-F-09 | Number-format entry diagnostics, including the round-trip trap | WP-F-01 | AC-FMT-044 | |

Sequencing note: WP-F-01 is the schema change and everything in WP-F-02,
WP-F-03, and WP-F-08 sits on it, so it goes first. WP-F-06 and WP-F-07 are
independent and can land in any order.

**WP-F-05 is withdrawn, not deferred.** The blue-black toggle applied the same
classification to a selection, which is what AutoColor Selection already does on
a one-cell selection, so it was a second name for one behaviour. AC-FMT-036 now
reads "no third command applies the classification". Setting a colour against
the classification is the font colour cycle's job. Its acceptance range moved to
WP-F-08.

**WP-F-07 exists because `commands/RIBBON_LAYOUT.md` describes a generator and a
validator that do not exist.** `RibbonRoutes.cs` is hand-maintained and
`RibbonRoutes.For()` falls back silently on an unknown id, so a descriptor can
advertise a route that does not work; four Model Check descriptors already do.
Either the code or the document has to change, and the document is right.

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
