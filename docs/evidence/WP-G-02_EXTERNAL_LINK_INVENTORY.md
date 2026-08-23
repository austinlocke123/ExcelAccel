# WP-G-02 External-link inventory

Date: 2026-08-23
Status: Inventory, grouping, status, search, navigation, and export complete.
Chart-series, query/connection, and validation scanning is not built and is
reported as a coverage gap rather than claimed.

## Contract

- **Capability:** CAP-LINK-001
- **Acceptance:** AC-LINK-001..008, AC-LINK-010, AC-LINK-011, AC-SEC-004 (met);
  AC-LINK-009 partially, see below
- **Allowed implementation:** the link model, its presentation, the coordinator,
  the read-only adapter, the exporter, and their host wiring.
- **Excluded:** WP-G-03 compare, name usage coverage, and every mutation id the
  spec forbids.

## What it does

`links.inventory.open` groups every qualified external link by normalized source
and lists where each is used. `links.inventory.export` writes it to a local CSV
after confirming a manifest. Search filters the captured snapshot.

Sources normalize to their file name, purely textually, so a formula token
`'[Budget.xlsx]Sheet1'` and a full path `C:\Models\Budget.xlsx` land in one
group. No path is resolved and no file is touched.

## Status never claims a check that did not happen

AC-LINK-003 forbids a status implying unperformed network discovery, so the
vocabulary is deliberately about **this Excel session**:

| Status | Shown as | Means |
|---|---|---|
| `OpenInSession` | Open in Excel | The source workbook is open here |
| `NotOpenInSession` | Not open (not checked) | Not open here; existence and reachability are unknown |
| `Broken` | Broken reference | A usage no longer resolves |
| `Unsupported` | Unsupported source | The token could not be normalized |

**No filesystem probe is performed**, deliberately. `File.Exists` on a UNC path
pointing at a dead share can block for a long time, and a link inventory that
hangs Excel is worse than one that says "not checked". A test asserts the label
contains "not checked" and contains none of "missing", "inaccessible", or
"unavailable".

The adapter never calls `UpdateLink`, `BreakLink`, `ChangeLink`, or
`Workbooks.Open`, and never triggers a recalculation, which is AC-LINK-006.

## Nothing is hidden

A source Excel reports but no scanned usage explains **still appears**, with "No
usage found in a scanned category" — omitting it would contradict what the user
sees in Excel's own Edit Links dialog. A usage whose source Excel did not report
also appears. Broken, unsupported, and ambiguous sources all stay listed and
categorized.

Unscanned categories are named as coverage gaps rather than counted as zero, so a
partial inventory cannot look complete.

## No new window

Like WP-G-01, this projects into `TraceResultPresentation` and renders through
the existing shared `TraceViewRuntime`. **No WinForms view was written** beyond
the export manifest dialog.

## The bug the tests caught

The first implementation chose each group's display string by taking the longest
form seen. A usage token is a **raw formula**, and
`='[BUDGET.XLSX]Sheet2'!$B$2` is longer than `C:\Models\Budget.xlsx`, so the
source column would have shown a formula fragment where the user expects a file.

Worse, it would have defeated the export redaction: the path column is redacted
to a file name by default, but a formula fragment smuggled into that column
carries whatever the formula carries.

Only a source Excel itself reported now contributes a display string; anything
else shows its normalized token.

## Export redaction

Source paths are **redacted to the file name by default**, because a link path
routinely names a drive, a client, or a deal. The manifest states counts, fields,
and coverage before anything is written, and the file goes through a temporary
file in the same directory so a failure cannot leave a half-written export.

## What is bounded

Formula scanning uses `SpecialCells` to read only formula cells, skips any
worksheet whose used range exceeds 250,000 cells, and stops at 5,000 usages and
512 sources. Each ceiling that bites is reported as a coverage gap.

## Not fully met

**AC-LINK-009**, qualified non-cell object navigation, is only half met. Non-cell
usages are visibly non-navigable with a stated reason, which is the criterion's
fallback, but no qualified selection path exists for chart series or connections
because neither is scanned yet.

## Verification

```
build Release   0 warnings, 0 errors
build Debug     0 warnings, 0 errors
test  Release   703/703 passed (was 675/675)
smoke           scripts/Test-ExcelAddIn.ps1 PASS, including link_inventory=exercised
                Excel exited, 0 surviving processes, 0 stale session markers
```

Real Excel, with an external formula, a broken external formula, and an external
defined name seeded:

```
smoke.links.inventory  sources=2|usages=4|navigable=2|complete=False|
                       book one.xlsx:NotOpenInSession:1+nosuchbook.xlsx:Broken:3
```

Two sources grouped correctly; the seeded workbook shows Broken because one of
its three usages carries `#REF!`; a formula usage and a defined-name usage
grouped under one source; two cell-anchored usages navigable and the broken one
not; and `complete=False` because chart, query, and validation are unscanned.

## A smoke-harness trap worth recording

Seeding the external formulas early made an **unrelated** assertion fail: the
bounded dependent scan counts scanned formulas exactly, and two new formulas took
it from 16 to 18. The seeding now runs after every scan assertion, with a comment
saying why. A shared worksheet in a long harness is a shared fixture, and adding
cells to it is not a local change.
