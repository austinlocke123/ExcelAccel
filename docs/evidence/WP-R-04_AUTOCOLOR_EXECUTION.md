# WP-R-04 AutoColor execution

Date: 2026-08-24
Status: AutoColor Selection is registered and works end to end. Worksheet scope
remains gated.

## Contract

- **Capability:** CAP-FMT-002
- **Acceptance:** AC-FMT-034..037, AC-FMT-046 (met); worksheet scope still behind
  `PERFORMANCE_QUALIFICATION_REQUIRED`
- **Allowed implementation:** the command, its descriptors, the COM adapter, and
  the host wiring.
- **Excluded:** worksheet-scale qualification, and the AutoFormat family.

## What now works

`format.auto_color.selection` recolours the selection by what each cell is, and a
single Ctrl+Z reverses the whole change. `format.auto_color.worksheet` is
registered but refuses at planning time, so it is visible and honest rather than
absent.

Only font colour is written. The command declares exactly one changed property,
`cell_font_color_block_v1`, which is the block WP-R-03 introduced, so the plan,
the write, and the undo receipt all describe the same thing.

Execution follows the pattern the rest of the product uses: authorize, re-capture
and compare the full fingerprint, write, read back and verify, record the
receipt, and restore the exact prior colours on any failure. A rollback that
cannot be verified reports partial rather than failed, because "failed" would
imply nothing changed.

## Two defects found while wiring it

**The planner advertised a bound it could never reach.**
`MaximumPlannableCells` is 250,000, but the fingerprint was the raw concatenation
of every cell's address, kind, formula, and colour, and
`PreconditionFingerprint` caps at 1,000,000 characters. Any sizeable range threw
`ArgumentOutOfRangeException` from inside the planner instead of refusing
cleanly — an unhandled argument exception surfacing as a crash dialog rather than
a refusal. The fingerprint now hashes the cell stream incrementally, which has no
length limit, so the declared bound is the only thing deciding what is plannable.

**A recolour too large to undo is refused before writing.** The undo value holds
at most `FontColorBlock.MaximumCells`, so a plan with more changes than that
would produce a write the user could not reverse. It now refuses with
`RESOURCE_LIMIT`, naming both numbers.

## The adapter

Writes group by colour: every cell taking one colour is combined into a single
range and coloured in one COM call, in bounded batches, so a recolour costs
roughly one write per distinct colour per batch rather than one per cell.

Reads are still per cell, because kind, formula, and colour are needed together
and a cell's error state is only visible through its displayed text. That is the
slower half, it is bounded by the planner's ceiling, and the class comment says
so rather than claiming a bulk read the code does not do.

Undo reads and writes the block through the same port under
`cell_font_color_block_v1`; every other property defers to the existing selection
adapter.

## Verification

```
build Release   0 warnings, 0 errors
build Debug     0 warnings, 0 errors
test  Release   734/734 passed (was 723/723)
smoke           scripts/Test-ExcelAddIn.ps1 PASS, including autocolor=exercised
                Excel exited, 0 surviving processes, 0 stale session markers
```

Real Excel, with a typed number, a same-sheet formula, and a text label all
starting from one colour none of them should end on:

```
smoke.autocolor.selection  status=Success|changed=3|G1:#0000FF+G2:#000000+G3:#000000
smoke.autocolor.undo       outcome=Success|G1:#123456+G2:#123456+G3:#123456
```

The typed number went blue, the formula and the label went black, and **one undo
restored all three** to their exact prior colour. The harness now asserts the
restored colours, so a stuck recolour or a failed undo fails the smoke rather
than being noticed by eye.

## Still gated

Worksheet scope needs its performance qualification, rollback and fault-injection
evidence at worksheet scale, and a preview surface built for thousands of rows
rather than a message box. `AutoColorPlanner.ExecutionGate` refuses it at
planning time, and a test asserts that.
