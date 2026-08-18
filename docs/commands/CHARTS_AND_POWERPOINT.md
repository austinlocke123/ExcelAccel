# Chart-formatting and PowerPoint snapshot command contracts

Status: **Draft for review**  
Capabilities: CAP-CHART-001, CAP-PPT-001  
Earliest phase: individually gated after Excel-only reliability evidence

## 1. Chart-formatting boundary

- Operate on Excel native charts only.
- Initial scope is exactly one explicitly selected embedded chart or chart sheet.
- Commands change only declared chart properties.
- Chart type, source data, series formulas, workbook values/formulas, and
  embedded objects never change unless a future separately approved command
  contract says so.
- Unsupported chart/axis/series types refuse or skip with exact reasons.
- Workbook-wide and inferred multi-chart targeting are deferred.

## `chart.font_family.set`

- Version: 1
- Impact: low
- Parameters: approved font family and explicit target elements (`chart`,
  `title`, `legend`, `axes`, `data_labels`, or selected element)
- Plan: exact supported text elements and current/proposed family
- Changed properties: font family only
- Preview: threshold-based for whole-chart element sets
- Undo: chart-property session receipt after qualification
- Acceptance: AC-CHART-001 through AC-CHART-005

## `chart.font_size.set`

- Same boundary as font family; changes font size only within approved bounds.
- Acceptance: AC-CHART-001 through AC-CHART-005

## `chart.legend_position.set`

- Parameters: explicit supported native position or `none`
- Changed properties: legend existence/position only; legend formatting and
  series remain unchanged
- Acceptance: AC-CHART-006, AC-CHART-007

## `chart.gridlines.toggle`

- Parameters: explicit axis and major/minor gridline kind
- CanExecute: the specified axis/gridline concept exists for the chart type
- Changed properties: selected gridline visibility only
- Acceptance: AC-CHART-008, AC-CHART-009

## `chart.axis.number_format.set`

- Parameters: explicit axis ID and profile number format
- Changed properties: axis tick-label number format only; no source-link change
  unless a typed option explicitly and safely controls it
- Locale behavior follows AC-LOC-001.
- Acceptance: AC-CHART-010 through AC-CHART-012

## `chart.axis.bounds.set`

- Version: 1
- Impact: medium
- Parameters: explicit axis ID and typed minimum/maximum or `auto`; log-scale and
  date-axis behavior require separately qualified typed options
- Plan: current/proposed bounds, scale mode, and validation that min < max
- Preview: mandatory
- Changed properties: selected axis bounds/auto flags only
- Acceptance: AC-CHART-013 through AC-CHART-016

## `chart.series.gap_width.set`

- Parameters: explicit series group and value within Excel-qualified bounds
- Changed properties: gap width only
- Unsupported chart families refuse.
- Acceptance: AC-CHART-017, AC-CHART-018

## `chart.series.overlap.set`

- Parameters: explicit series group and value within qualified bounds
- Changed properties: overlap only
- Acceptance: AC-CHART-017, AC-CHART-018

## `chart.border.remove`

- Parameters: exact selected border target (`chart_area` or `plot_area`)
- Changed properties: target border visibility/style only
- Acceptance: AC-CHART-019

## `chart.plot_area.cleanup`

- This is a versioned formatting recipe, not an inferred cleanup.
- Plan lists every property, such as fill/border/gridline, and every target.
- Preview is mandatory; per-property opt-out is supported before confirmation.
- It never changes data, chart type, titles, axes, series order, or labels unless
  those properties are explicitly present in the approved recipe version.
- Acceptance: AC-CHART-020 through AC-CHART-023

## `chart.size.apply_standard`

- Parameters: explicit approved width/height and anchor policy
- Changed properties: ChartObject dimensions and, only if explicit, position
- It does not resize or move worksheet cells.
- Acceptance: AC-CHART-024 through AC-CHART-026

## `chart.house_style.preview` and `chart.house_style.apply`

- Style is a versioned declarative recipe of the individual qualified chart
  properties above.
- Preview enumerates supported changes, skips, current/proposed values, and
  per-property opt-outs.
- Apply requires the exact preview hash and one chart target.
- Undo requires an exact chart-property receipt within size/capability limits.
- Acceptance: AC-CHART-020 through AC-CHART-023, AC-CHART-027

## 2. PowerPoint snapshot boundary

- PowerPoint integration is a separately packaged optional adapter.
- Initial release supports image snapshots only; editable embedded workbooks and
  live links are deferred.
- PowerPoint must already be open, and the user explicitly chooses the target
  presentation and slide.
- Multiple PowerPoint instances, presentations, or ambiguous targets require
  selection; the add-in never guesses.
- All Excel and PowerPoint COM calls remain on their qualified owning thread/
  apartment path and no proxy crosses a worker boundary.

## `powerpoint.snapshot.range_image`

- Version: 1
- Impact: medium mutation of the chosen PowerPoint presentation; Excel read-only
- Parameters: explicit Excel range, PowerPoint presentation/slide, approved
  image format, placement/size, and optional alt text
- Snapshot: source range visual representation through the Phase-specific
  qualified transfer mechanism; no live link or workbook embedding
- Plan: exact source, target presentation/slide, image format, dimensions,
  placement, overwrite/overlap warning, temporary/clipboard behavior, and undo
  eligibility
- Preview: mandatory target and placement summary
- Execute: revalidate source and target, create exactly one PowerPoint shape,
  verify it, assign approved alt text, restore clipboard/focus state where
  qualified, and delete temporary artifacts
- Undo: session receipt may delete the exact created shape only if its identity
  and relevant properties still match
- Failure: delete any partially created shape; Excel source is unchanged; report
  exact target state if cleanup fails
- Acceptance: AC-PPT-001 through AC-PPT-009

## `powerpoint.snapshot.chart_image`

- Same contract using one explicitly selected native Excel chart as source.
- Chart data/type/source are read-only.
- Acceptance: AC-PPT-001 through AC-PPT-009

## `powerpoint.snapshot.repeat_placement`

- Not approved. Reusing the last target or placement is deferred until users can
  see and explicitly confirm the target; it must never become a hidden default.

## PowerPoint qualification requirements

- clean source/target application state on success/failure/cancellation;
- ambiguous/multiple-instance behavior;
- target closed between preview and commit;
- clipboard contention and restoration if clipboard is used;
- PowerPoint busy/rejected-call bounded retry;
- partial shape creation and cleanup;
- presentation read-only/protected state;
- display scaling and placement units;
- repeated snapshot memory/handle/COM soak;
- PowerPoint crash/exit while the adapter is active;
- no link, embedded workbook, auto-save, email, upload, or presentation save.
