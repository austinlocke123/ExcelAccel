# WP-G-03 Same-shape comparison

Date: 2026-08-24
Status: Range comparison, target-side navigation, and export complete. Workbook
structure comparison and the frozen timing corpus are not built and are not
claimed.

## Contract

- **Capability:** CAP-CMP-001
- **Acceptance:** AC-CMP-001..008, AC-CMP-014, AC-CMP-016..019, AC-SEC-004 (met);
  AC-CMP-009/010 partially, AC-CMP-011..013 and AC-CMP-015 not built
- **Allowed implementation:** the comparison engine, its presentation, the
  coordinator, the read-only adapter, the exporter, and their host wiring.
- **Excluded:** workbook structure comparison, sheet pairing, and the frozen
  comparison performance corpus.

## What it does

`compare.source.capture` remembers the selected range as the source side.
`compare.ranges.same_shape` compares the current selection against it, position
by position. `compare.result.navigate_target` re-lists the same differences
against the target side so a result can be navigated on either workbook, and
`compare.results.export` writes them to a local CSV after a manifest.

Both sides must already be open. Nothing is opened, saved, closed, or
recalculated, and the source is stored as an **address only** and re-read when
the comparison runs, so a stale capture can never present old content as current.
If the captured workbook or worksheet is gone, the comparison refuses with
`STALE_CONTEXT` rather than reaching for something else.

## Same shape means same shape

Unequal dimensions are refused, with a message that says the comparison "does not
align or truncate". There is no row matching, no shift detection, and no
comparison of a common sub-rectangle. That is the contract, and the refusal
states it so nobody assumes the tool quietly did something cleverer.

## The distinction that makes results readable

Formula differences are classified into four kinds: identical, **equivalent
shape**, structure, and outside coverage.

Equivalent shape is the one that matters. `=J1*2` and `=K1*2` are different text
but the same formula copied to a different anchor. Reporting that as a structural
difference would bury the real ones under every copied column in the model. The
comparison normalizes token shape with references treated as interchangeable, so
a copied formula reads as "Same shape, different references".

A formula the parser does not cover is reported as **outside coverage**: the
texts differ, and nothing is claimed about why. It is never silently called a
structural difference.

## Nothing unexamined is called equal

A category that was not compared is named in the coverage gaps, so a clean result
with number formats unexamined reports `Partial` and says
`category_not_compared:number_format`. AC-CMP-005 is the reason: a result that
looks complete while quietly skipping a category is worse than one that admits
the gap.

Result and cell ceilings behave the same way — both are reported rather than
silently truncating, and the difference **counts** cover every difference found,
not only those retained for display.

## No new window

Like the two inventories, this projects into `TraceResultPresentation` and renders
through the existing shared `TraceViewRuntime`. No read-only view was written.

## Export redaction

Cell contents are **excluded by default**, because the differing cells are the
ones most likely to carry the numbers and formulas the model is about. The
manifest names both sides, the counts, the fields, and the coverage before
anything is written, and the file goes through a temporary file in the same
directory so a failure leaves any existing export exactly as it was.

## Verification

```
build Release   0 warnings, 0 errors
build Debug     0 warnings, 0 errors
test  Release   771/771 passed (was 740/740)
smoke           scripts/Test-ExcelAddIn.ps1 PASS, including compare=exercised
                Excel exited, 0 surviving processes, 0 stale session markers
```

Real Excel, comparing two three-cell blocks where row 1 matches, row 2 holds a
different value, and row 3 holds the same formula copied to another anchor:

```
smoke.compare.capture  captured
smoke.compare          differences=2|cells=3|complete=False|
                       r2:Constant:Identical+r3:Formula:EquivalentShape
```

The matching row produced nothing, the differing value was found, the copied
formula was classified as same-shape rather than a structural change, and
`complete=False` honestly reflects that number formats and displayed values were
not among the compared categories.

## Not built, not claimed

**AC-CMP-011..013, workbook structure comparison.** It needs explicit sheet
pairing, which is a mapping surface rather than a comparison change, and unpaired
sheets have to be presented as structure differences in their own right.

**AC-CMP-015, separately measured snapshot, analysis, and render timing against a
frozen comparison corpus.** The spec proposes two open 250,000-cell workbooks with
a frozen P95 target; neither the corpus nor the target exists.

**AC-CMP-009/010, worksheet comparison**, is partly covered: the engine and the
shape refusal are exactly what a worksheet comparison needs, but the documented
equal-bounds policy and the hidden/filtered reporting it requires are not
implemented, so no worksheet-level command is registered.
