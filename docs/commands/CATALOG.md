# Command catalog

Status: **Draft for review**

Command IDs become stable only when this catalog is approved. The catalog
defines scope; it is not authorization to implement later-phase commands.

Detailed retained-feature contracts are indexed in [`README.md`](README.md).

> **Number-format commands changed.** `format.number.*` now cycles a
> profile-defined list of formats rather than applying one fixed format, and the
> set of cycles is user-editable and extensible. The rows below still describe
> the property scope and impact correctly; for cycle semantics, defaults, the
> settings surface, and basis points see
> [`FORMAT_CYCLES.md`](FORMAT_CYCLES.md).

## 1. Phase 1A command groups

All Phase 1A mutations:

- operate only on the active user-selected target;
- change only declared properties;
- use block operations where the Excel API permits;
- refuse unsupported merged, protected, array/spill, multi-area, or object
  states unless the individual contract explicitly supports them;
- never save the workbook;
- use session receipts only after receipt semantics pass Phase 0 qualification.

### 1.1 Formatting cycles

Cycle values come from the active profile. Mixed selections advance from the
active cell's effective property unless the profile selects a separately tested
mixed-selection policy. An empty cycle or invalid profile causes refusal.

| Command ID | Changed property | Context | Impact | Preview | Acceptance |
|---|---|---|---|---|---|
| `format.font_color.cycle` | Font color only | Cell/range | low | none | AC-FMT-001, AC-REL-005 |
| `format.fill_color.cycle` | Fill color only | Cell/range | low | none | AC-FMT-001, AC-REL-005 |
| `format.number.general` | Number format cycle only | Cell/range | low | none | AC-FMT-002 |
| `format.number.currency` | Number format cycle only | Cell/range | low | none | AC-FMT-002, AC-LOC-001 |
| `format.number.percentage` | Number format cycle only | Cell/range | low | none | AC-FMT-002, AC-LOC-001 |
| `format.number.multiple` | Number format cycle only | Cell/range | low | none | AC-FMT-002 |
| `format.number.date` | Number format cycle only | Cell/range | low | none | AC-FMT-002, AC-LOC-001 |
| `format.number.boolean` | Number format cycle only | Cell/range | low | none | AC-FMT-002 |
| `format.number.decimals.increase` | Number format decimal precision only | Cell/range | low | none | AC-FMT-003 |
| `format.number.decimals.decrease` | Number format decimal precision only | Cell/range | low | none | AC-FMT-003 |
| `format.alignment.horizontal.cycle` | Horizontal alignment only | Cell/range | low | none | AC-FMT-001 |
| `format.alignment.vertical.cycle` | Vertical alignment only | Cell/range | low | none | AC-FMT-001 |
| `format.indent.increase` | Indent level only | Cell/range | low | none | AC-FMT-001 |
| `format.indent.decrease` | Indent level only | Cell/range | low | none | AC-FMT-001 |
| `format.underline.cycle` | Underline style only | Cell/range | low | none | AC-FMT-001 |
| `format.font_size.cycle` | Font size only | Cell/range | low | none | AC-FMT-001 |
| `format.row_height.cycle` | Row height only | Complete selected rows or cell-derived rows | low | none | AC-FMT-004 |
| `format.column_width.cycle` | Column width only | Complete selected columns or cell-derived columns | low | none | AC-FMT-004 |

### 1.2 Explicit formatting commands

| Command ID | Behavior | Impact | Acceptance |
|---|---|---|---|
| `format.center_across.apply` | Apply Center Across Selection without merging cells or changing values. | low | AC-FMT-005 |
| `format.border.sum_bar.apply` | Apply only the profile-defined sum-bar border edge/style. | low | AC-FMT-006 |
| `format.border.remove` | Remove borders only. | low | AC-FMT-006 |
| `format.rows.autofit` | AutoFit only rows intersecting the validated selection. | low | AC-FMT-007 |
| `format.columns.autofit` | AutoFit only columns intersecting the validated selection. | low | AC-FMT-007 |
| `view.gridlines.toggle` | Toggle gridline visibility for the active worksheet window only. | low | AC-FMT-008 |
| `view.zoom.set` | Set active-window zoom to a typed/profile value within qualified bounds. | low | AC-FMT-008 |
| `view.panes.freeze` | Freeze at an explicit validated anchor; no inferred layout. | medium | AC-FMT-009 |
| `view.panes.unfreeze` | Remove freeze panes in the active window. | low | AC-FMT-009 |

### 1.3 AutoColor

#### `format.auto_color.selection`

- Capability: CAP-FMT-002
- Phase: 1A
- Impact: medium
- Changed properties: font color only
- Context: cell/range selection within qualified size; formulas/values are read
  but never changed
- Plan: classify text, numeric hardcodes, same-sheet formulas, cross-sheet
  formulas, external-workbook formulas, and errors using deterministic rules
- Preview: threshold-based; mandatory when selection exceeds the Phase 0
  qualified immediate limit
- Undo: session receipt when qualified
- Failure: unsupported cells are skipped with classified counts; an adapter
  failure stops and rolls back the current bounded write set
- Acceptance: AC-FMT-010 through AC-FMT-013

#### `format.auto_color.worksheet`

Same contract as selection scope, but always requires a target/count preview and
is subject to `PERF-004`. Workbook-wide AutoColor is deferred.

### 1.4 Navigation

Navigation commands are read-only with respect to workbook content. Navigation
history is session-only and separate from mutation undo.

| Command ID | Behavior | Acceptance |
|---|---|---|
| `navigate.sheet.previous` | Activate the previous visible worksheet with profile-defined wrap behavior. | AC-NAV-001 |
| `navigate.sheet.next` | Activate the next visible worksheet with profile-defined wrap behavior. | AC-NAV-001 |
| `navigate.cell.a1` | Select A1 on the active worksheet. | AC-NAV-002 |
| `navigate.used.first` | Navigate to the first qualified used cell; refuse if no stable used range is available. | AC-NAV-003 |
| `navigate.used.last` | Navigate to the last qualified used cell; expose the used-range limitation. | AC-NAV-003 |
| `navigate.region.edge.up` | Navigate to the prior populated-region boundary using documented Excel-compatible rules. | AC-NAV-004 |
| `navigate.region.edge.down` | Navigate to the next populated-region boundary. | AC-NAV-004 |
| `navigate.region.edge.left` | Navigate to the prior populated-region boundary. | AC-NAV-004 |
| `navigate.region.edge.right` | Navigate to the next populated-region boundary. | AC-NAV-004 |
| `navigate.history.back` | Restore the prior valid workbook/sheet/range location without mutation. | AC-NAV-005 |
| `navigate.history.forward` | Restore the next valid history location without mutation. | AC-NAV-005 |
| `navigate.bookmark.add_session` | Add a session-only bookmark for the current location. | AC-NAV-006 |
| `navigate.bookmark.next_session` | Navigate to the next still-valid session bookmark. | AC-NAV-006 |
| `navigate.bookmark.previous_session` | Navigate to the previous still-valid session bookmark. | AC-NAV-006 |
| `navigate.bookmark.clear_session` | Clear session bookmarks without workbook mutation. | AC-NAV-006 |

## 2. Phase 1B catalog

Detailed contracts are defined in:

- [`DISCOVERY_STYLES_AND_PROFILES.md`](DISCOVERY_STYLES_AND_PROFILES.md)
- [`FORMULA_TRANSFORMS.md`](FORMULA_TRANSFORMS.md)
- [`DATA_CLEANING.md`](DATA_CLEANING.md)
- [`WORKBOOK_OPERATIONS.md`](WORKBOOK_OPERATIONS.md) for selection tools

| Capability | Proposed command families | Notes |
|---|---|---|
| Command discovery | `command.search.open`, favorite add/remove/invoke | Local-only, keyboard complete. |
| Style library | capture/apply named styles, major/minor/date header, assumption, formula, linked formula, output, warning, total | Formatting only; never values/formulas/comments/validation/hyperlinks. |
| Formula transform | reference toggle, smart copy down/right, spacing, transpose, IFERROR toggle, reverse sign | Parser-qualified syntax only. |
| Unit scale | `formula.units.to_thousands`, `from_thousands`, `to_millions`, `from_millions` | Intent-based names; operator shown before execution. |
| Paste/fill | formula/value from above, formulas only, values only, formats only, deterministic numeric/date fill | Formula-to-value permission explicit for value commands. |
| Data cleaning | whitespace, non-printing, typed text/number/date conversions, explicit display conversions | No catch-all clean command. |
| Profiles | import/export/preview, binding presets, printable active-binding export | No executable content. Competitor-specific imports remain unapproved. |
| Selection tools | formulas, constants, blanks, errors, numeric hardcodes, external formulas | Selection/focus only; bounded areas. |

## 3. Phase 2 catalog

Detailed contracts are defined in:

- [`AUDITING.md`](AUDITING.md)
- [`MODEL_CHECK.md`](MODEL_CHECK.md)

| Capability | Proposed command families | Safety boundary |
|---|---|---|
| Formula tracing | direct/indirect precedents and dependents, return navigation | Read-only; bounded depth and cycle detection. |
| Formula Inspector | parse tree, reference navigation | Structural explanation only; no evaluation or AI prose. |
| Model Check | selection/sheet/workbook scan, navigate/export findings | Read-only scan; findings are evidence, not declarations of error. |

Model Check rules include copied-formula pattern inconsistency, constants inside
formula-consistent ranges, embedded numeric hardcodes with allowlists, formula
errors, broken/external/circular references, and number-format inconsistency.

## 4. Later gated catalog

Detailed contracts are defined in:

- [`NAMES_AND_LINKS.md`](NAMES_AND_LINKS.md)
- [`COMPARE.md`](COMPARE.md)
- [`SENSITIVITIES_AND_FINANCE.md`](SENSITIVITIES_AND_FINANCE.md)
- [`CHARTS_AND_POWERPOINT.md`](CHARTS_AND_POWERPOINT.md)
- [`WORKBOOK_OPERATIONS.md`](WORKBOOK_OPERATIONS.md) for row/column and broad
  formatting commands

| Capability | Initial permitted scope | Explicitly excluded until separately approved |
|---|---|---|
| Named ranges | inventory, search, navigate, report broken/unused/conflicting | Rename/delete and reference rewriting. |
| External links | inventory, category, navigate, report inaccessible source | Repoint/break and formula-to-value conversion. |
| Compare | same-shape, read-only formula/value/selected-format diff | Structural alignment, confidence matching, mutation. |
| Sensitivity | explicit one/two-way inputs, output, destination, native Data Table where safe | Inferred inputs or silent fallback. |
| Circularity | inspect status and explicit switch templates | Silent iterative-calculation changes. |
| Finance templates | typed declarative formula templates with preview | Arbitrary code, scripts, whole-model generation. |
| Charts | selected native chart property formatting | Proprietary charts or source/type changes. |
| PowerPoint | explicit one-time snapshot to chosen open presentation/slide | Live links and ambiguous target selection. |
| Row/column outline | hide/unhide/group/ungroup/Smart Hide | Structural insert/delete require a separate transaction ADR. |
| Broad formatting | sheet style, workbook AutoFormat, workbook AutoColor | Requires workbook-scale preview, rollback, and performance proof. |

Command chains are not in the approved catalog.

## 5. Proposed profile defaults requiring review

Profiles own exact cycles, colors, number formats, sizes, widths, heights,
shortcuts, thresholds, and wrap behavior. The specification intentionally does
not hard-code bank-specific conventions before profile review.

An implementation agent MUST NOT invent these defaults. A default-profile
artifact and acceptance fixtures are required before Phase 1A feature coding.
