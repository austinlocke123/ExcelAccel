# WP-G-01 Named-range inventory

Date: 2026-08-23
Status: Inventory, search, navigation, and export complete. Usage coverage
(AC-NAME-008..010) is not built and is not claimed.

## Contract

- **Capability:** CAP-NAME-001
- **Acceptance:** AC-NAME-001..007, AC-NAME-011, AC-SEC-004 (met);
  AC-NAME-008..010 (not built, see below)
- **Allowed implementation:** the inventory model, its presentation, the
  coordinator, the read-only adapter, the exporter, and their host wiring.
- **Excluded:** external-link inventory (WP-G-02), compare (WP-G-03), and every
  mutation id the spec forbids.

## What it does

`names.inventory.open` lists every qualified defined name with its scope,
visibility, built-in status, target category, and whether it can be navigated.
`names.inventory.export` writes it to a local CSV after confirming a manifest.
Search filters the captured snapshot.

Targets are classified from local metadata alone, into Range, Constant, Formula,
External, Broken, or Unresolved. **Nothing opens or contacts an external source**,
so a name pointing at a closed workbook is classified without touching it.

The `#REF!` check runs before the external check, so a broken external name
reports as broken rather than external. Getting that order wrong would file the
most urgent category under the least urgent one.

Non-navigable names stay listed with a stated reason. Dropping them would hide
exactly the names a reviewer is looking for.

## No new window

The inventory renders through the existing shared `TraceViewRuntime` by
projecting into `TraceResultPresentation`. **No WinForms view was written**, so
there is no second read-only window to drift from the auditing ones. That is what
the restart note meant by reusing the shared trace view.

## Search never touches Excel

`NameInventoryCoordinator.Search` deliberately takes no port. Re-reading the
workbook per keystroke would put a scan behind typing, and would let results
change under a filter the user did not touch. A test asserts the capture count
stays at one across an open and a search.

## Export redaction

Name expressions are **excluded by default**, because a target can carry a file
path or a business term. The manifest states the destination, the count, the
exact fields, and the coverage statement before anything is written, and the
file is written through a temporary file in the same directory so a failure
cannot leave a half-written export that looks complete.

## What real Excel found that the unit tests could not

The first live run returned four names from a workbook seeded with three:

```
count=4|...|SmokeConst:Constant+SmokeExternal:External+SmokeRange:Range+_xlfn.SINGLE:Formula
```

`_xlfn.SINGLE` is a function shim Excel adds for itself. It has no business in a
list of the user's named ranges, and no unit test would have found it because no
fixture would have invented it.

Reserved-name recognition moved into `ReservedNames` in the core layer, covering
the classic reserved names and the `_xlfn.`, `_xlref`, `_xludf.`, `_xlchart.`
prefixes. It sits in Core rather than the adapter specifically so it is testable;
the interop project has no test project. The live run then returned exactly the
three seeded names.

## Not built, not claimed

**AC-NAME-008..010, usage navigation**, needs per-category qualification across
cell formulas, name expressions, chart series, data validation, and print areas.
The spec says each category is separately qualified, and none has been. The
inventory reports what a name *is*, not yet everywhere it is used.

`names.navigate_target` is registered and refuses with a clear reason when no
inventory is open; navigation itself happens by activating a row, which is where
the selected name is known.

## Verification

```
build Release   0 warnings, 0 errors
build Debug     0 warnings, 0 errors
test  Release   675/675 passed (was 650/650)
smoke           scripts/Test-ExcelAddIn.ps1 PASS, including name_inventory=exercised
                Excel exited, 0 surviving processes, 0 stale session markers
```

Real Excel, with a range, a constant, and an external name seeded:

```
smoke.names.inventory  count=3|complete=True|navigable=1|
                       SmokeConst:Constant+SmokeExternal:External+SmokeRange:Range
```

Each classified correctly, and exactly one navigable.

`MutationCommandsDoNotExist` asserts `names.rename`, `names.delete`,
`links.repoint`, and `links.break` are absent from the registry, since the spec
forbids those ids existing at all until new requirements and ADRs are accepted.
