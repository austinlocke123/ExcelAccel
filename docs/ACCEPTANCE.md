# Acceptance criteria and release gates

Status: **Draft for review**

Acceptance criteria are normative observable outcomes. Detailed test cases,
fixtures, and harness code will live beside the future tests and may evolve
without weakening these outcomes.

## 1. Evidence rules

- Each result records build, Excel version/channel/bitness, Windows version,
  runtime, machine profile, workbook fixture version, and warm/cold state.
- Automated evidence links to a repeatable command or CI artifact.
- Manual evidence identifies the operator and exact procedure.
- `Pass with waiver` is not allowed for data integrity, Excel process crashes,
  exception containment, state restoration, offline execution, signature
  validation, or privacy.
- A criterion marked for a later phase does not block an earlier phase unless
  that capability is included in the build.

## 2. Phase 0 exit gates

| ID | Criterion |
|---|---|
| AC-P0-001 | A minimal Excel-DNA add-in loads, invokes one callback, opens/closes one task pane, disables, and unloads across the proposed clean-VM support matrix without an Excel crash or orphaned UI/resource. |
| AC-P0-002 | The runtime ADR is accepted after cold-start, deployment, coexistence, unload, crash-isolation, and clean-machine evidence. |
| AC-P0-003 | A vertical-slice command completes validation, snapshot, pure planning, preview, state-guarded mutation, postcondition verification, receipt creation, and explicit result handling. |
| AC-P0-004 | Static and dynamic tests fail when an Excel/PowerPoint COM proxy enters a domain API or worker-thread payload. |
| AC-P0-005 | The formula strategy passes 100% of the approved v1 syntax/locale/reference corpus and refuses every explicitly unsupported case without mutation. |
| AC-P0-006 | The benchmark harness produces repeatable distributions on the approved reference machine/corpus and establishes variance plus frozen Phase 1 targets. |
| AC-P0-007 | AutoSave/coauthoring experiments produce an accepted impact-tier policy and demonstrate stale-plan refusal under representative intervening changes. |
| AC-P0-008 | A signed packaging prototype installs, loads, disables, upgrades, rolls back, and uninstalls on an approved clean VM with no development tools. |

No Phase 1 feature work begins until these gates pass or the governing
requirement is explicitly revised.

## 3. Architecture and boundary criteria

| ID | Criterion |
|---|---|
| AC-ARCH-001 | Domain assemblies reference no Excel-DNA, Office interop, UI, registry, network, or host-specific type. |
| AC-ARCH-002 | Every Excel callback routes through the shared exception and diagnostic boundary. |
| AC-ARCH-003 | All Excel object-model calls are made through the approved adapter on the qualified Excel thread. |
| AC-ARCH-004 | Background analysis accepts and returns only immutable non-COM data and discards stale results. |
| AC-ARCH-005 | Equivalent canonical inputs produce byte-identical canonical plan serialization. |
| AC-ARCH-006 | Execution refuses when workbook identity, structure, target, or relevant property preconditions differ from the approved plan. |
| AC-ARCH-007 | No active-phase package contains the deferred PowerPoint adapter, structural compare engine, command-chain engine, or persistent undo journal. |

## 4. Reliability and crash-safety criteria

| ID | Criterion |
|---|---|
| AC-REL-001 | Qualification, fault-injection, and soak suites produce zero add-in-attributable Excel process terminations or hangs. |
| AC-REL-002 | Injected exceptions at every command lifecycle transition return a non-success result and do not cross the Excel callback boundary. |
| AC-REL-003 | Calculation mode, events, screen updating, alerts, cursor/status ownership, and qualified focus/selection state are restored on success, refusal after guard entry, cancellation, mutation failure, and rollback failure. |
| AC-REL-004 | A state-restoration failure cannot be reported as success and produces a stable local diagnostic plus quarantine signal. |
| AC-REL-005 | A formatting command changes only its declared property; seeded values, formulas, unrelated formats, names, VBA, validation, comments, and hyperlinks remain unchanged. |
| AC-REL-006 | Reentrant events and simultaneous mutation attempts for one workbook are refused or serialized without deadlock, recursion, partial writes, or lost state. |
| AC-REL-007 | Excel-thread code never waits on a worker that requires Excel-thread completion. Deadlock-detection/fault tests terminate within the bounded timeout. |
| AC-REL-008 | COM busy/rejected-call behavior stops at the approved retry/time limit and returns an actionable failure without freezing Excel. |
| AC-REL-009 | Repeated workbook open/close, pane show/hide, command run/cancel, recalc, and add-in disable/re-enable remain within approved memory, handle, thread, and duration bounds. |
| AC-REL-010 | An unclean operation marker causes conservative next-start recovery with no automatic workbook mutation, pane restore, scan, or undo replay. |
| AC-REL-011 | Session undo restores only properties whose current value still matches the recorded command post-state and refuses changed properties. |
| AC-REL-012 | Closing a workbook or disabling the add-in clears its session receipts, bookmarks, pending work, and event ownership without touching workbook content. |

## 5. Command framework criteria

| ID | Criterion |
|---|---|
| AC-CMD-001 | Every released action appears exactly once in the registry with ID, version, capability, context, impact, changed properties, preview, undo, and acceptance metadata. |
| AC-CMD-002 | Equivalent invocation through shortcut, Ribbon, search, favorite, or pane produces the same canonical plan. |
| AC-CMD-003 | `CanExecute` performs no mutation and returns a specific refusal for unsupported selection, protection, edit mode, workbook state, resource limit, and quarantine. |
| AC-CMD-004 | A mandatory-preview command cannot execute without confirmation of the same plan hash. |
| AC-CMD-005 | A changed workbook/selection/precondition after preview produces `Refused` and offers refresh; no silent retarget occurs. |
| AC-CMD-006 | Changed, skipped, warning, rollback, and restoration results match observed postconditions; incomplete work never reports `Success`. |
| AC-CMD-007 | No command automatically saves, closes, uploads, emails, publishes, changes AutoSave, or changes iteration/calculation policy unless a future approved contract explicitly says so. |
| AC-CMD-008 | Canonical plan serialization contains no timestamp, random identifier, localized display string, or unstable enumeration order. |

## 6. Keyboard, invocation, and UX criteria

| ID | Criterion |
|---|---|
| AC-KEY-001 | Normal typing and Excel shortcuts in cell/formula edit mode are never intercepted except by an explicitly qualified edit-mode command. |
| AC-KEY-002 | Native Excel, Windows, accessibility-tool, add-in, duplicate, prefix, and reserved-key conflicts are detected to the documented coverage and shown before assignment. |
| AC-KEY-003 | An incomplete multi-stroke sequence times out or cancels with Escape without workbook mutation or stuck keyboard state. |
| AC-KEY-004 | Every fixed-parameter released command can be assigned a shortcut; every parameterized command is reachable and operable from the keyboard. |
| AC-UX-001 | Ribbon, panes, search, dialogs, and errors restore expected workbook focus unless the user explicitly retains pane focus. |
| AC-UX-002 | Every interactive control has an accessible name, deterministic focus order, high-contrast support, and non-color-only status. |
| AC-UX-003 | Primary pane actions remain visible and keyboard usable at 1024x768 effective resolution and qualified Windows scaling values. |
| AC-UX-004 | No pane scans continuously; scans are explicit and cancellation/progress behavior matches the command contract. |
| AC-UX-005 | Failure/refusal UI contains command, safe reason, remediation, and diagnostic ID without sensitive workbook data. |

## 7. Formatting, navigation, selection, and structure criteria

| ID | Criterion |
|---|---|
| AC-FMT-001 | Each cycle changes only its named property and advances according to the validated active profile and approved mixed-selection rule. |
| AC-FMT-002 | Named number-format commands apply the exact profile format without changing the underlying value/formula or another format property. |
| AC-FMT-003 | Decimal increase/decrease preserves the qualified number-format family and never changes the value. |
| AC-FMT-004 | Row-height and column-width cycles affect only the resolved selected rows/columns and refuse ambiguous unsupported contexts. |
| AC-FMT-005 | Center Across Selection never merges cells or changes values/formulas. |
| AC-FMT-006 | Sum-bar and remove-border commands change only declared border edges/styles. |
| AC-FMT-007 | AutoFit affects only target rows/columns and completes through qualified block/object operations without unapproved cell loops. |
| AC-FMT-008 | Gridline and zoom commands affect only the active qualified window/view properties. |
| AC-FMT-009 | Freeze/unfreeze commands use the explicit validated anchor and restore focus; they do not infer a financial-model layout. |
| AC-FMT-010 | AutoColor deterministically distinguishes text, numeric hardcodes, same-sheet formulas, cross-sheet formulas, external formulas, and errors on the golden corpus. |
| AC-FMT-011 | AutoColor changes font color only and preserves values, formulas, fill, number formats, borders, and workbook structure. |
| AC-FMT-012 | Worksheet AutoColor always previews scope/counts and refuses when size, protection, array/spill, AutoSave, or coauthoring policy cannot be satisfied. |
| AC-FMT-013 | AutoColor cancellation before commit changes nothing; qualified mutation failure rolls back or reports exact remaining targets. |
| AC-FMT-014 | Sheet/workbook formatting plans enumerate every enabled component, target, skip, property, sample, resource cost, and undo eligibility. |
| AC-FMT-015 | Broad formatting changes only explicitly enabled formatting/view properties and preserves values, formulas, names, VBA, comments, validation, hyperlinks, and workbook structure. |
| AC-FMT-016 | Sheet style requires mandatory preview and exact-plan postcondition verification. |
| AC-FMT-017 | Every enabled broad-format component has qualified rollback/receipt semantics or is unavailable. |
| AC-FMT-018 | Sheet-style failure rolls back qualified writes or reports exact remaining changed properties/targets and cannot report success. |
| AC-FMT-019 | Workbook AutoFormat/AutoColor requires an explicit included-sheet inventory and never silently expands to hidden/unselected sheets. |
| AC-FMT-020 | Workbook-scale formatting meets frozen resource/performance/AutoSave/coauthoring gates and cannot expose partial success as the default result. |
| AC-NAV-001 | Previous/next sheet navigation honors visible-sheet and wrap policy without altering workbook content. |
| AC-NAV-002 | A1 navigation selects A1 on the intended active sheet and preserves workbook content. |
| AC-NAV-003 | First/last used navigation matches the documented used-range semantics and exposes/refuses unsupported ambiguity. |
| AC-NAV-004 | Region-edge navigation matches the approved golden workbook cases for blanks, formulas returning empty text, hidden cells, and table boundaries. |
| AC-NAV-005 | Back/forward history restores valid locations, skips or reports closed/deleted targets, and never invokes mutation undo. |
| AC-NAV-006 | Session bookmarks do not dirty or alter the workbook and are cleared on workbook close/add-in disable. |
| AC-SELECT-001 | Selection commands match their exact predicates on the golden corpus and change only Excel selection/focus. |
| AC-SELECT-002 | Scope, hidden/filtered policy, maximum cells/areas, and parser coverage are explicit and never silently broaden. |
| AC-SELECT-003 | Excessive/unstable multi-area results preview/refuse rather than freezing Excel or producing a truncated unlabeled selection. |
| AC-SELECT-004 | Selection tools preserve workbook content and push the prior valid location to navigation history. |
| AC-SELECT-005 | Blank selection distinguishes truly empty cells from formulas returning empty text. |
| AC-SELECT-006 | Numeric-hardcode selection distinguishes numeric constants from numeric text and numeric literals inside formulas. |
| AC-SELECT-007 | External-formula selection matches only formulas with supported external-workbook references and reports parser gaps. |
| AC-STRUCT-001 | Hide/unhide resolves exact complete selected rows/columns and changes hidden state only. |
| AC-STRUCT-002 | Hide/unhide preview reports hidden/filtered/table/outline interactions and exact bounds. |
| AC-STRUCT-003 | Unsupported/protected/partial contexts refuse without changing visibility. |
| AC-STRUCT-004 | Qualified undo restores hidden state only after optimistic post-state validation. |
| AC-STRUCT-005 | Hide/unhide never changes values, formulas, dimensions, outline levels, filters, or sheet visibility. |
| AC-STRUCT-006 | Group/ungroup plans report exact bounds, current/proposed outline levels, overlaps, and collapse state. |
| AC-STRUCT-007 | Group/ungroup changes outline state only and follows the approved overlap policy. |
| AC-STRUCT-008 | Unsupported/protected/partial outline contexts refuse the entire plan. |
| AC-STRUCT-009 | Outline failure rolls back qualified changes or reports exact remaining bounds/levels. |
| AC-STRUCT-010 | Outline undo restores only unchanged written outline/collapse properties. |
| AC-STRUCT-011 | Smart Hide creates/extends the explicit approved group and collapses it; it does not merely hide rows/columns. |
| AC-STRUCT-012 | Insert/delete requires exact complete bounds/count and mandatory preview of affected workbook objects/reference categories. |
| AC-STRUCT-013 | Structural preview inventories values, formulas, formats, names, tables, charts, spills/arrays, validation, comments, outlines, protection, and unsupported coverage. |
| AC-STRUCT-014 | Insert/delete never ships on a support case lacking a dedicated proven structural transaction and rollback policy. |
| AC-STRUCT-015 | Stale bounds/content/structure/AutoSave/coauthoring state refuses before structural mutation. |
| AC-STRUCT-016 | Insert behavior matches native approved shift/copy-origin and reference-adjustment golden cases. |
| AC-STRUCT-017 | Insert failure restores exact prior structure or reports release-blocking incomplete rollback. |
| AC-STRUCT-018 | Delete preview explicitly identifies data/formula-bearing cells and irreversible/unsupported object effects. |
| AC-STRUCT-019 | Delete failure/undo restores the complete approved before-state and reference behavior or the command remains absent from the registry. |

## 8. Discovery, styles, profiles, formulas, cleaning, and locale criteria

| ID | Criterion |
|---|---|
| AC-SEARCH-001 | Registry search ranks name, alias, category, description, and shortcut matches deterministically for the same query and registry. |
| AC-SEARCH-002 | Opening or typing in search performs no workbook scan or mutation and returns initial results within the frozen immediate-interaction budget. |
| AC-SEARCH-003 | Search shows current executability/refusal without executing a command, and invocation routes through the normal command lifecycle. |
| AC-SEARCH-004 | Search remains keyboard complete, accessible, and responsive with the maximum qualified registry size. |
| AC-FAV-001 | Add/remove updates only the user-profile favorite list atomically and handles duplicate/absent entries idempotently. |
| AC-FAV-002 | Favorite invocation produces the same canonical plan as direct invocation of the referenced command and arguments. |
| AC-FAV-003 | Missing, incompatible, locked, or invalid favorite references remain visible and non-mutating with a remediation. |
| AC-FAV-004 | A favorite cannot bypass the referenced command's current `CanExecute`, preview, impact, undo, or acceptance policy. |
| AC-STYLE-001 | Capture reads exactly one source cell and stores only explicitly selected supported formatting properties. |
| AC-STYLE-002 | Captured recipes contain no values, formulas, names, paths, notes, comments, validation, hyperlinks, macros, or executable content. |
| AC-STYLE-003 | Unsupported requested capture properties refuse the capture rather than producing an unlabeled partial style. |
| AC-STYLE-004 | Applying a local or built-in style changes exactly the recipe properties and no workbook content property outside them. |
| AC-STYLE-005 | Style preview reports exact properties, targets, skips, overwrites, and undo eligibility from the same executable plan. |
| AC-STYLE-006 | Unsupported target/property combinations follow the recipe's explicit skip/refuse policy and never disappear from result counts. |
| AC-STYLE-007 | Style apply failure rolls back qualified writes or reports exact remaining changed targets and cannot report success after incomplete work. |
| AC-STYLE-008 | Built-in/organization-locked styles cannot be deleted; deleting a local style does not scan or modify workbooks. |
| AC-PROF-001 | Active-phase profile resolution follows defaults, optional organization policy, and user profile with locked-setting enforcement. A workbook layer is absent unless separately approved. |
| AC-PROF-002 | Profile write/import uses validate-then-atomic-replace and leaves the prior profile intact on parse, schema, migration, permission, or I/O failure. |
| AC-PROF-003 | Export contains no workbook-derived formulas, values, names, paths, screenshots, or executable content. |
| AC-PROF-004 | Unknown/incompatible command IDs and versions are reported and never silently reinterpreted. |
| AC-PROF-005 | Profile/binding export previews the destination and manifest, writes through temp/validate/replace, and never transmits or launches output. |
| AC-PROF-006 | Export refuses an existing destination unless overwrite is explicit and preserves the existing file on failure. |
| AC-PROF-007 | Import preview validates size, archive paths, checksum, schema, versions, migrations, IDs, assets, and policy locks without changing settings. |
| AC-PROF-008 | Import apply accepts only the exact still-valid preview plan, atomically activates the validated profile, and preserves a bounded pre-import backup. |
| AC-PROF-009 | Import failure leaves the prior effective profile and registry bindings active with no partially migrated settings. |
| AC-FORM-001 | Parser coverage passes the approved corpus for A1/R1C1, relative/absolute/mixed, quoted sheet, external workbook, structured, name, dynamic-array, spill, and locale forms explicitly marked supported. |
| AC-FORM-002 | Every unsupported formula construct refuses before mutation with a coverage-specific reason. |
| AC-FORM-003 | Smart Copy, spacing, transpose, IFERROR, sign, and unit transforms preserve formula semantics, array/spill policy, and operator precedence on golden cases. |
| AC-FORM-004 | Unit commands use intent-based names and preview the actual operator; `to_thousands` divides by 1,000 and `from_thousands` multiplies by 1,000. |
| AC-FORM-005 | Reference toggle acts only on the reference containing the qualified formula-edit caret and cycles through the approved absolute/relative states. |
| AC-FORM-006 | Ambiguous caret, unsupported syntax, name, or structured reference refuses without changing edit text, focus, selection, or keyboard state. |
| AC-FORM-007 | Reference toggle preserves sheet/workbook qualifiers, range structure, locale dialect, and every non-target token. |
| AC-FORM-008 | Smart Copy and spacing translate every supported reference by the exact source-to-destination displacement on golden cases. |
| AC-FORM-009 | Absolute, mixed, sheet/workbook-qualified, name, structured, dynamic-array, and spill cases follow the accepted transform/refusal matrix. |
| AC-FORM-010 | Nonblank destinations and threshold-sized writes require a preview with exact overwrite counts and representative formulas. |
| AC-FORM-011 | One unsupported formula, unsafe overlap, protected target, or unsafe array/spill/table context refuses the entire copy/spacing plan. |
| AC-FORM-012 | Copy/spacing writes formulas only, verifies postconditions, and supports qualified rollback/receipt before release. |
| AC-FORM-013 | Formula transpose maps every source position to the correct transposed destination and transforms relative references according to the approved transpose semantics. |
| AC-FORM-014 | Transpose preserves constants as constants and formulas as formulas and never copies formatting in the initial contract. |
| AC-FORM-015 | Any unsupported source formula, destination conflict, or invalid bound refuses the complete transpose before mutation. |
| AC-FORM-016 | Transpose preview and receipt cover the entire destination rectangle and exact overwritten before-state. |
| AC-FORM-017 | IFERROR toggle adds exactly one configured top-level wrapper when absent and removes exactly that wrapper when present. |
| AC-FORM-018 | IFERROR toggle never removes an unrelated wrapper, double-wraps an equivalent configured formula, or alters the interior AST. |
| AC-FORM-019 | Non-formulas are counted/skipped and unsupported formulas follow the all-or-refuse parser policy. |
| AC-FORM-020 | Reapplying IFERROR toggle restores the original supported formula canonically. |
| AC-FORM-021 | Sign/unit transforms preserve operator precedence and formula evaluation semantics across the golden corpus. |
| AC-FORM-022 | Preview identifies formula versus constant targets, exact operator/scale, skips, and representative before/after output. |
| AC-FORM-023 | Text, blanks, errors, arrays/spills, and unsupported formulas remain unchanged and are counted according to contract. |
| AC-FORM-024 | Qualified undo restores exact prior formulas/values only when optimistic post-state validation passes. |
| AC-FORM-025 | Reapplying a unit command applies its named operation again; no hidden toggle or unit inference occurs. |
| AC-FORM-026 | Paste uses one documented source/clipboard model and deterministic destination shape/repetition rules. |
| AC-FORM-027 | Formulas-only paste changes formulas only; formats-only paste changes only its approved formatting set. |
| AC-FORM-028 | Paste shape mismatch, unsupported destination, or stale source refuses or follows an explicitly previewed repetition rule. |
| AC-FORM-029 | Paste failure rolls back qualified writes or reports exact remaining destinations; unrelated clipboard/workbook properties are preserved. |
| AC-FORM-030 | Values-only paste explicitly counts destination formulas converted to values and requires confirmation before replacing them. |
| AC-FORM-031 | Numeric/date fill produces the exact typed start/step/direction/count sequence with no inference from neighboring workbook content. |
| AC-FORM-032 | Date fill rejects ambiguous free-text dates and follows the explicit calendar/locale policy. |
| AC-FORM-033 | Fill previews any nonblank overwrite and remains within exact destination bounds. |
| AC-FORM-034 | Fill cancellation before commit changes nothing and qualified undo restores the destination before-state. |
| AC-FORM-035 | Formula-from-above translates the immediately adjacent source formula for each destination using approved fill-down reference semantics. |
| AC-FORM-036 | Value-from-above copies the underlying current value only and explicitly previews formula-source-to-value destinations. |
| AC-FORM-037 | Both commands preserve source cells and destination formatting and require preview for nonblank overwrites. |
| AC-FORM-038 | Unsupported source formulas, stale sources, invalid shapes, or unavailable receipts refuse the full operation before mutation. |
| AC-DATA-001 | Text normalization operates only on text constants and skips formulas, numbers, errors, and blanks according to the declared source predicate. |
| AC-DATA-002 | Trim/collapse behavior matches the versioned whitespace character table on Unicode and locale fixtures. |
| AC-DATA-003 | Text normalization is idempotent and does not alter non-whitespace characters. |
| AC-DATA-004 | Mixed selections preview changed/skipped/error categories and representative exact outputs. |
| AC-DATA-005 | Nonprinting removal deletes only the approved control table and preserves international text plus configured tabs/newlines. |
| AC-DATA-006 | Text/number conversion accepts only strings or values matching the complete explicit grammar and never partially parses. |
| AC-DATA-007 | Decimal, thousands, sign, parentheses, currency, percent, and output-format policies match typed parameters and locale fixtures. |
| AC-DATA-008 | Number-to-text never introduces unintended scientific notation, rounding, or locale-dependent output outside the declared format. |
| AC-DATA-009 | Conversion changes only value and any separately previewed number format; formulas remain formulas and are skipped. |
| AC-DATA-010 | Invalid conversion leaves the source unchanged and appears in categorized result counts. |
| AC-DATA-011 | Date normalization accepts only the explicit input patterns and produces the exact declared serial/text output and format. |
| AC-DATA-012 | Ambiguous dates and undeclared two-digit-year pivots refuse without guessing. |
| AC-DATA-013 | Date conversion matches the approved locale/calendar/time fixtures and workbook date-system policy. |
| AC-DATA-014 | Reapplying a supported date normalization is idempotent or explicitly reported as already normalized. |
| AC-DATA-015 | Each blank/zero/NA/NM/dash conversion matches only its explicitly declared source value/type/spelling/case policy. |
| AC-DATA-016 | `zero_to_blank` does not convert formula results displaying zero or text zero unless a future explicit option is approved. |
| AC-DATA-017 | Display-value conversion requires a mandatory categorized preview and qualified value undo. |
| AC-DATA-018 | A stale source after preview refuses or requires a regenerated preview rather than becoming a silent skip. |
| AC-DATA-019 | Identical source snapshot and parameters produce identical changed/skipped/error sets and outputs. |
| AC-LOC-001 | Number formats, formula separators/function forms, dates, decimal/list separators, and serialization pass the approved locale matrix while IDs and canonical values remain invariant. |

## 9. Auditing and Model Check criteria

| ID | Criterion |
|---|---|
| AC-AUD-001 | Direct precedent/dependent results identify the correct workbook, sheet, address/range, edge kind, and formula/value/error classification on the golden corpus. |
| AC-AUD-002 | Every result reports scan scope, parser/reference coverage, unresolved/external edges, and whether completeness can be claimed. |
| AC-AUD-003 | Direct precedent analysis deduplicates equivalent references without losing distinct source-edge evidence. |
| AC-AUD-004 | Closed external sources are represented but never opened or contacted automatically. |
| AC-AUD-005 | A non-formula or unsupported formula returns a categorized refusal or explicitly partial result, never a misleading complete trace. |
| AC-AUD-006 | Dependent scans read only the declared worksheet/workbook scope and never silently expand it. |
| AC-AUD-007 | Reverse-reference indexing matches brute-force golden results within approved parser coverage. |
| AC-AUD-008 | Parser gaps are counted and prevent a false completeness claim. |
| AC-AUD-009 | Large dependent scans meet progress, cancellation, responsiveness, and bounded-resource requirements. |
| AC-AUD-010 | Indirect traversal is deterministic, breadth/depth bounded, and ordered stably for identical snapshots. |
| AC-AUD-011 | Cycle detection terminates and represents the cycle without infinite or duplicate expansion. |
| AC-AUD-012 | Depth and result caps produce an explicit truncated result rather than a hang, crash, or silent omission. |
| AC-AUD-013 | Each indirect edge retains the direct evidence from which it was derived. |
| AC-AUD-014 | Cancellation exposes no partial traversal as a completed trace. |
| AC-AUD-015 | Trace navigation revalidates targets, preserves session return history, and never mutates workbook content. |
| AC-AUD-016 | Formula Inspector renders the correct immutable tree of functions, operators, constants, references, arrays, and nesting on the parser corpus. |
| AC-AUD-017 | Inspector nodes retain correct source spans and keyboard focus order and remain responsive at the qualified node cap. |
| AC-AUD-018 | Unsupported syntax identifies the exact category/span and is not presented as a fully parsed formula. |
| AC-AUD-019 | Inspector performs no subexpression evaluation, AI explanation, complexity score, or workbook mutation. |
| AC-AUD-020 | Selecting an inspector node changes only inspector selection unless explicit reference navigation is invoked. |
| AC-AUD-021 | Reference navigation resolves only supported unambiguous nodes and refuses external/dynamic/ambiguous targets safely. |
| AC-CHECK-001 | Selection/worksheet/workbook scans evaluate exactly the enabled rule IDs/versions against one stable snapshot. |
| AC-CHECK-002 | Findings contain rule/version, severity, target, evidence, coverage, fingerprint, and stale/navigation state. |
| AC-CHECK-003 | Findings never declare correctness, assign a workbook/formula score, or create review-status workflow. |
| AC-CHECK-004 | Identical snapshot, rules, parser, and configuration produce canonically identical findings/fingerprints. |
| AC-CHECK-005 | A rule failure makes the scan failed/incomplete with the rule ID; it is never silently omitted. |
| AC-CHECK-006 | Cancellation preserves prior results and does not present partial results as a completed scan. |
| AC-CHECK-007 | Worksheet scan uses the documented used-range/coverage policy and meets its frozen scan budget. |
| AC-CHECK-008 | Workbook preview inventories included/excluded sheets, workload, unsupported categories, and resource limits. |
| AC-CHECK-009 | Pattern-inconsistency golden cases identify the correct deterministic peer group, baseline normalized formula, exception, and neighbors. |
| AC-CHECK-010 | Peer minimum, boundaries, blanks, labels, totals, and confidence categories follow the approved profile configuration. |
| AC-CHECK-011 | Parser gaps prevent a false complete-consistency result for the affected peer region. |
| AC-CHECK-012 | Reordering scan execution does not change canonical pattern findings. |
| AC-CHECK-013 | A constant interrupting a formula-consistent region is reported separately from a constant embedded inside a formula. |
| AC-CHECK-014 | Configured input/label/boundary exclusions are explicit and deterministic; no semantic intent is guessed. |
| AC-CHECK-015 | Embedded-hardcode detection identifies exact numeric-literal AST nodes and source spans. |
| AC-CHECK-016 | Versioned allowlists and contextual structural exclusions produce deterministic included/excluded results. |
| AC-CHECK-017 | Changing an allowlist produces the expected deterministic rescan delta without changing the workbook. |
| AC-CHECK-018 | Numeric text, cell constants, and calculated results are not misclassified as embedded formula literals. |
| AC-CHECK-019 | Formula-error findings classify supported Excel error values and broken `#REF!` references correctly without forcing recalculation. |
| AC-CHECK-020 | Error navigation selects the exact current cell or refuses when stale. |
| AC-CHECK-021 | External-reference findings cover every qualified formula/name form and identify the local affected target. |
| AC-CHECK-022 | External checks perform no source discovery, open, network access, link update, or workbook mutation. |
| AC-CHECK-023 | Circular-reference checks terminate, identify supported cycles/components, and expose coverage/truncation. |
| AC-CHECK-024 | Declared circularity switches are recognized only from explicit configuration, never inference. |
| AC-CHECK-025 | Inspecting circularity does not change iterative calculation settings or recalculate solely for the scan. |
| AC-CHECK-026 | Number-format inconsistency uses the approved normalized format identity and peer-region rule. |
| AC-CHECK-027 | Format findings identify baseline and exception without changing formatting. |
| AC-CHECK-028 | Values/formulas affect peer grouping only where the versioned rule explicitly declares it. |
| AC-CHECK-029 | Finding navigation revalidates exact target identity and pushes prior location to audit history. |
| AC-CHECK-030 | Local ignore stores a rule-specific normalized fingerprint with no raw formula/value content. |
| AC-CHECK-031 | Ignore/unignore changes only the atomic local profile and takes effect through an explicit rescan. |
| AC-CHECK-032 | An ignore suppresses only equivalent findings for the same rule/version/scope semantics. |
| AC-CHECK-033 | Active ignores are visible, removable, and portable only through an explicitly approved export/import action. |
| AC-CHECK-034 | Rescan captures a new snapshot of the exact prior scope/configuration and never relabels old findings as current. |
| AC-CHECK-035 | Results export requires an explicit destination and previewed manifest/redaction choices. |
| AC-CHECK-036 | Default export excludes formulas/values and writes deterministically through temp/validate/replace. |
| AC-CHECK-037 | Export failure preserves existing destination data and never transmits output. |

## 10. Names, links, and comparison criteria

| ID | Criterion |
|---|---|
| AC-NAME-001 | Inventory includes qualified workbook- and worksheet-scoped names with stable identity, visibility, built-in classification, and target category. |
| AC-NAME-002 | Broken, duplicate-across-scope, constant, formula, range, external, hidden, and unsupported names are classified without mutation. |
| AC-NAME-003 | Unsupported expressions remain visible as unresolved and are never silently omitted. |
| AC-NAME-004 | Inventory reports per-category usage coverage and cannot claim completeness outside it. |
| AC-NAME-005 | Opening/refreshing inventory never creates, edits, deletes, or recalculates a name. |
| AC-NAME-006 | Name search/filter is deterministic over the completed snapshot and performs no per-keystroke workbook scan. |
| AC-NAME-007 | Target navigation revalidates and selects supported current targets and explains non-navigable targets. |
| AC-NAME-008 | Usage detection matches golden formulas, names, chart series, validation, and print areas/titles within separately approved coverage. |
| AC-NAME-009 | Usage navigation selects the exact supported object/range and refuses stale or non-selectable records. |
| AC-NAME-010 | Coverage gaps and unsupported usage-bearing objects remain visible in counts. |
| AC-NAME-011 | Export previews raw expression/path inclusion, excludes sensitive content by default, and never mutates names. |
| AC-LINK-001 | Inventory classifies qualified formula, name, chart, query/connection, validation, and separately approved link-bearing objects. |
| AC-LINK-002 | Usages group deterministically by normalized local source token without contacting the source. |
| AC-LINK-003 | Status labels are based only on documented local Excel metadata/safe checks and never imply unperformed network discovery. |
| AC-LINK-004 | Closed, inaccessible, broken, unsupported, and ambiguous sources remain visible and categorized. |
| AC-LINK-005 | Inventory reports coverage and unsupported counts per object category. |
| AC-LINK-006 | Inventory never updates, opens, repoints, breaks, or recalculates a link. |
| AC-LINK-007 | Link search/filter is deterministic over completed results with no continuous rescan. |
| AC-LINK-008 | Usage navigation revalidates exact local targets and never opens the external source. |
| AC-LINK-009 | Supported non-cell object navigation follows a qualified selection path or is visibly non-navigable. |
| AC-LINK-010 | Stale/deleted usages refuse navigation and mark affected results stale. |
| AC-LINK-011 | Export excludes/redacts paths by default, requires manifest confirmation, and never transmits. |
| AC-CMP-001 | Comparison reads two explicit already-open sources and never opens, saves, closes, recalculates, or mutates either source. |
| AC-CMP-002 | Range comparison requires equal dimensions and refuses mismatch without alignment or truncation. |
| AC-CMP-003 | Formula, constant, optional displayed value, number format, and approved formatting categories match independent golden differences. |
| AC-CMP-004 | Formula results distinguish exact text, supported canonical structure/reference differences, and unsupported parser coverage. |
| AC-CMP-005 | Unsupported/unexamined properties are never reported as equal. |
| AC-CMP-006 | Both snapshots carry identity/precondition fingerprints and source change invalidates affected results. |
| AC-CMP-007 | Cancellation/analysis failure leaves prior results visible and does not present a partial comparison as complete. |
| AC-CMP-008 | Large result sets expose counts, coverage, truncation, and remain bounded/virtualized. |
| AC-CMP-009 | Worksheet comparison uses one documented equal-bounds policy and includes hidden/filtered state as specified. |
| AC-CMP-010 | Different worksheet shapes produce a structure mismatch and no inferred alignment. |
| AC-CMP-011 | Workbook comparison uses only explicit sheet mapping or exact-name pairing. |
| AC-CMP-012 | Unpaired sheets are reported as structure differences and never fuzzily paired. |
| AC-CMP-013 | Cell-level comparison runs only on valid explicit equal-shape pairs. |
| AC-CMP-014 | Same inputs/categories/parser/configuration produce canonically identical difference records. |
| AC-CMP-015 | Snapshot, analysis, and result-render timing are measured separately against the frozen comparison corpus. |
| AC-CMP-016 | Source/target navigation revalidates exact targets and refuses closed/changed/deleted objects. |
| AC-CMP-017 | Export preview states sources, categories, raw content/path inclusion, result count, and destination. |
| AC-CMP-018 | Export is deterministic, local, temp/validate/replace, and non-transmitting. |
| AC-CMP-019 | Export failure preserves existing destination data and comparison sources. |

## 11. Sensitivity, circularity, and finance-template criteria

| ID | Criterion |
|---|---|
| AC-SENS-001 | One/two-way creation requires explicit output, input cell(s), axis values, destination, and formatting; no model input is inferred. |
| AC-SENS-002 | Validation covers protection, bounds, overlap, spills/arrays/tables, calculation/iteration, AutoSave/coauthoring, and resource caps before preview. |
| AC-SENS-003 | Preview shows exact native Data Table bindings, destination rectangle, overwritten cells/properties, calculation state, warnings, and undo eligibility. |
| AC-SENS-004 | Creation uses a native Excel What-If Data Table and never silently substitutes a formula grid. |
| AC-SENS-005 | The command does not change model assumptions outside native Data Table calculation or silently change calculation/iteration settings. |
| AC-SENS-006 | Results match independently constructed native Data Tables on the approved financial/calculation corpus. |
| AC-SENS-007 | Source/destination/precondition change after preview refuses execution and requires refresh. |
| AC-SENS-008 | Failure removes partially created content or reports the exact remaining destination and blocks further mutation. |
| AC-SENS-009 | Qualified undo removes the unchanged created Data Table and restores the exact destination before-state. |
| AC-SENS-010 | Two-way row and column axes bind to the correct distinct explicit inputs and orientation. |
| AC-SENS-011 | Two-way output matrix matches native golden cases across qualified calculation modes and circularity states. |
| AC-SENS-012 | Invalid identical inputs, empty axes, or oversized destination refuse without partial creation. |
| AC-SENS-013 | Inspection reports qualified table bounds, formula anchor, input bindings, and calculation state without mutation/recalculation. |
| AC-SENS-014 | Unsupported native Data Table metadata is labeled unavailable rather than guessed. |
| AC-CIRC-001 | Circularity inspection reports qualified Excel-exposed references and bounded graph cycles with coverage/truncation. |
| AC-CIRC-002 | Cycle analysis terminates deterministically and does not duplicate infinite paths. |
| AC-CIRC-003 | Iterative calculation settings are displayed but not changed by inspection. |
| AC-CIRC-004 | Declared switch-controlled circularity is distinguished only through explicit versioned configuration. |
| AC-CIRC-005 | Inspection performs no formula mutation, setting change, save, or automatic recalculation solely for analysis. |
| AC-CIRC-006 | Switch preview requires explicit template, switch cell, participating formula cells, semantics, and typed mappings. |
| AC-CIRC-007 | Preview shows every proposed formula/value/label/format and calculation implication from one immutable plan. |
| AC-CIRC-008 | The engine never infers participating formulas or undeclared model assumptions. |
| AC-CIRC-009 | Switch insertion applies only the exact previewed cells/properties and never changes iterative settings. |
| AC-CIRC-010 | Stale formula/destination/settings/AutoSave/coauthoring state refuses insertion. |
| AC-CIRC-011 | Inserted switch behavior matches independent golden on/off cases and preserves unrelated formulas/settings. |
| AC-CIRC-012 | Failure rolls back the complete switch or reports exact remaining changes; partial wiring cannot report success. |
| AC-CIRC-013 | Iteration-settings command previews exact current/proposed enabled/max-iterations/max-change values and scope. |
| AC-CIRC-014 | Applying settings never saves a workbook and restores prior settings on qualified undo/failure. |
| AC-CIRC-015 | Unsupported build/scope/cross-workbook behavior refuses rather than approximates. |
| AC-CIRC-016 | Iteration-settings capability remains unavailable until its platform-wide side effects pass the support matrix. |
| AC-TPL-001 | Every template has stable ID/version, typed slots, exact destination shape, invariant formula model, preconditions, impact, and golden cases. |
| AC-TPL-002 | Templates contain no arbitrary code, macros, callbacks, loops, workbook-content branching, or network behavior. |
| AC-TPL-003 | Preview requires every slot binding and rejects type/shape/parser incompatibility before mutation. |
| AC-TPL-004 | Preview lists exact formulas, labels, formats, overwrites, user-supplied assumptions, and template/convention version. |
| AC-TPL-005 | The engine never invents an assumption or heuristically locates an input. |
| AC-TPL-006 | Formula output matches independently reviewed finance golden cases for the exact template version. |
| AC-TPL-007 | Insert accepts only the exact still-valid preview plan and changes only declared cells/properties. |
| AC-TPL-008 | Template insert failure is all-or-rollback and cannot leave a partial schedule reported as success. |
| AC-TPL-009 | Qualified undo validates written post-state and restores the exact destination before-state. |
| AC-TPL-010 | Reapplying/inserting an overlapping template requires explicit overwrite preview; no hidden merge occurs. |
| AC-TPL-011 | Library import validates schema, manifest, IDs/versions, formula coverage, conflicts, size/path limits, and prohibited content. |
| AC-TPL-012 | Imported content remains declarative/local and cannot execute during validation or use. |
| AC-TPL-013 | Import atomically activates all validated templates or none and preserves prior library on failure. |
| AC-TPL-014 | Candidate built-ins do not ship until their individual slot schema, convention, golden cases, and content review are approved. |

## 12. Chart and PowerPoint criteria

| ID | Criterion |
|---|---|
| AC-CHART-001 | Font commands target exactly one selected supported native chart and explicit text element set. |
| AC-CHART-002 | Font family/size commands change only the named font property and preserve chart data/type/source and all unrelated formatting. |
| AC-CHART-003 | Unsupported chart elements/types are refused or explicitly skipped with reasons and counts. |
| AC-CHART-004 | Whole-chart font preview lists every affected supported element and current/proposed value. |
| AC-CHART-005 | Qualified chart receipt restores only unchanged written properties and refuses later-modified properties. |
| AC-CHART-006 | Legend position changes only legend existence/position according to the explicit supported native value. |
| AC-CHART-007 | Legend change preserves series, plot dimensions, and all other legend formatting unless the exact native behavior necessarily reflows and is previewed. |
| AC-CHART-008 | Gridline toggle acts on the explicit axis and major/minor kind only. |
| AC-CHART-009 | Missing/unsupported axes or gridlines refuse without creating unrelated chart elements. |
| AC-CHART-010 | Axis number format changes only the selected tick-label format under the declared source-link policy. |
| AC-CHART-011 | Axis format matches locale/profile golden cases and preserves source data and bounds. |
| AC-CHART-012 | Ambiguous/missing axis selection refuses. |
| AC-CHART-013 | Axis-bounds preview shows exact axis, scale mode, current/proposed min/max/auto state. |
| AC-CHART-014 | Invalid min/max, log/date-axis, or unsupported chart combinations refuse before mutation. |
| AC-CHART-015 | Bounds apply changes only selected axis bound/auto properties. |
| AC-CHART-016 | Failure restores qualified prior bounds or reports exact remaining state. |
| AC-CHART-017 | Gap-width/overlap validates explicit series group and Excel-qualified bounds and changes only the named property. |
| AC-CHART-018 | Unsupported chart families/series groups refuse without changing series formulas/order/type. |
| AC-CHART-019 | Border removal changes only the explicit chart-area or plot-area border. |
| AC-CHART-020 | Cleanup/house-style recipes are versioned declarative property lists, not inferred appearance changes. |
| AC-CHART-021 | Recipe preview enumerates every supported change/skip and honors per-property opt-outs in the final plan. |
| AC-CHART-022 | Recipe apply uses the exact preview hash and never changes data, chart type, sources, or undeclared elements. |
| AC-CHART-023 | Recipe failure/undo follows exact chart-property post-state validation and reports partial state. |
| AC-CHART-024 | Standard size applies only explicit ChartObject dimensions and optional previewed position. |
| AC-CHART-025 | Sizing does not resize/move worksheet cells or alter chart aspect policy beyond typed parameters. |
| AC-CHART-026 | Chart sheets and embedded charts follow separately qualified geometry behavior. |
| AC-CHART-027 | Workbook-wide and inferred multi-chart targeting are absent from the released registry. |
| AC-PPT-001 | Snapshot requires one explicit Excel source and one explicit already-open PowerPoint presentation and slide. |
| AC-PPT-002 | Multiple instances/presentations or stale/closed targets require selection/refusal; the command never guesses. |
| AC-PPT-003 | Initial output is one image shape with no live link, embedded workbook, source mutation, presentation save, email, upload, or PowerPoint launch. |
| AC-PPT-004 | Preview states source, presentation/slide, image format, dimensions, placement, overlap warning, transfer mechanism implications, and undo eligibility. |
| AC-PPT-005 | Execute creates and verifies exactly one target shape and restores qualified Excel/PowerPoint focus/clipboard state. |
| AC-PPT-006 | Failure removes a partial shape and temporary artifacts or reports the exact remaining target state; Excel source remains unchanged. |
| AC-PPT-007 | Session undo deletes only the exact unchanged created shape and refuses if identity/properties no longer match. |
| AC-PPT-008 | Range/chart visual output and placement match golden cases across qualified Office builds and display scaling. |
| AC-PPT-009 | Repeated snapshots, cancellation, busy/rejected calls, target closure, and app shutdown pass COM/resource/crash soak gates. |

## 13. AutoSave, coauthoring, security, and offline criteria

| ID | Criterion |
|---|---|
| AC-SYNC-001 | Each command impact tier has documented behavior with AutoSave on/off/disabled and with detected coauthoring state. |
| AC-SYNC-002 | A remote/local intervening change to a planned property produces stale-plan refusal before mutation. |
| AC-SYNC-003 | High-impact commands refuse in unqualified coauthoring state and do not silently toggle AutoSave. |
| AC-SEC-001 | With licensing/update services disabled or stubbed and outbound traffic captured, the complete released core suite produces zero requests. |
| AC-SEC-002 | Seeded sensitive formulas, values, names, paths, sheet/workbook names, usernames, images, and chart data are absent from default logs and support manifests. |
| AC-SEC-003 | Malicious/fuzzed profile packages cannot execute code, traverse paths, exceed limits, partially modify settings, or create unbounded work. |
| AC-SEC-004 | User-approved support export is local, presents a manifest, and performs no automatic transmission. |

## 14. Performance criteria

| ID | Criterion |
|---|---|
| AC-PERF-001 | Phase 0 records cold/warm benchmark distributions and freezes P95 targets with measured variance. |
| AC-PERF-002 | Released operations meet the applicable frozen P95 and memory targets on the reference corpus. |
| AC-PERF-003 | Static/dynamic checks find no unapproved performance-sensitive cell-by-cell COM loop. |
| AC-PERF-004 | Operations exceeding 500 ms provide non-modal progress and remain cancellable before mutation where the contract permits. |
| AC-PERF-005 | Result panes remain keyboard responsive at 100,000 findings using bounded/virtualized presentation. |
| AC-PERF-006 | A regression exceeding the frozen tolerance fails the build unless a time-bounded human-approved record changes the baseline or requirement. |

## 15. General availability gates

- All criteria for included phases pass on the approved support matrix.
- Zero known workbook-corruption, add-in-attributable crash/hang,
  state-restoration, privacy, signing, or unrecoverable-settings defects.
- Zero open severity-1 defects.
- Installer, upgrade, rollback, disable, and uninstall pass on managed and
  unmanaged qualification profiles.
- Every released command has contract, help, limitations, shortcut/discovery
  path, preview/undo behavior, performance evidence, and automated coverage.
- The implementation traceability matrix contains no unexplained gap.
