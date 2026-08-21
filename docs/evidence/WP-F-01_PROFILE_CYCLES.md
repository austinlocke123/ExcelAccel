# WP-F-01 Profile schema v6: cycles as data

Date: 2026-08-20
Status: Complete

## Contract

- **Capability:** CAP-FMT-004
- **Acceptance:** AC-FMT-021..025, AC-FMT-031, AC-FMT-041, AC-FMT-043, AC-FMT-045
- **Allowed implementation:** profile schema and migration, cycle model, number
  and property cycle resolution, default profile contents, the tests and smoke
  assertions covering them.
- **Excluded:** the settings editor (WP-F-03), user-defined cycle management
  (WP-F-02), AutoColor classification (WP-F-08), any COM adapter work, and every
  tracing command.

## What changed

Nine top-level profile keys (`font_color_cycle` through `column_width_cycle`,
plus `number_formats`) collapsed into one `cycles` key holding family → ordered
array of named cycles. `ProfileDefinition`'s constructor went from 18 positional
parameters to 10.

Family identifiers are the property identifiers commands already declare in
`ChangedProperties`, so no command-to-family lookup table exists. That is what
keeps AC-FMT-045 honest rather than aspirational.

Every entry is stored as a string, including the numeric families. The
formatting port already returns strings and the old `NextNumber` helper already
stringified with `"0.####"` before comparing, so storing strings deleted that
helper and left one validator instead of two.

Style-shaped families delegate to `StylePropertyCatalog.Normalize` rather than
growing a parallel validator; only the two dimension families and the colour
reference sigil are handled in `CycleFamilyCatalog`.

## Two defects found and fixed before they shipped

**The default font colour cycle would have oscillated.** The approved spec said
the default cycle is the six category references in precedence order. Under the
default palette `@error` and `@external` both resolve to `#FF0000`, and
`@same_sheet` and `@text` both to `#000000`. The stateless advance matches the
first index by value, so the cycle would have run red → blue → red forever with
green and black unreachable. `ResolveCycle` now collapses entries by resolved
value, keeping the first occurrence, and the collapse is asserted directly.
Giving a category its own colour makes its entry reappear automatically.

**Precedence order was the wrong default order.** The first version of the
default cycle led with red, because precedence puts Error first. The real-Excel
smoke caught it: a font-colour keypress that had always produced black now
produced red. Precedence decides which category a cell *is*; it is not an order
to walk colours in. The default is now the six references in palette order,
resolving to black, blue, green, red, which is byte-identical in effect to the
v5 palette. AC-FMT-041 was reworded to match, and the reasoning is recorded in
`commands/AUTOCOLOR.md`.

## Deliberate behaviour changes

**`ApplyCurrencyFormatCommand` was deleted.** It hard-coded
`"$#,##0.00;($#,##0.00);-"`, which AC-FMT-045 forbids, and applied one format
rather than cycling. `format.number.currency` is now an ordinary entry in
`Phase1AFormattingCatalog`; the command id, ribbon KeyTip, and route are
unchanged, so no profile or `RibbonRoutes` edit was needed.

**Its refusal codes changed**, because the surviving pipeline uses different
ones for the same conditions. This is recorded rather than glossed:

| Condition | Was | Now |
|---|---|---|
| Multi-area selection | `MULTI_AREA_UNSUPPORTED` | `SELECTION_UNSUPPORTED` |
| Spill check unavailable | `EXCEL_CAPABILITY_MISSING` | `ARRAY_OR_SPILL_UNSAFE` |
| Oversized selection | refused at `CanExecute` | refused at `Plan`, `RESOURCE_LIMIT` |

No condition became permitted; only the code and the refusal point moved.
`CommandExecutionTests` was rewritten against the surviving contract rather than
deleted, preserving the safety, ceiling, revalidation, and stale-selection
coverage that suite existed to provide.

**A migrated profile is not a fresh one.** A v5 profile carrying
`"date": "m/d/yyyy"` migrates to a one-entry cycle holding exactly that, not to
the three-entry default a fresh install gets. AC-FMT-031 requires losing no
setting, so this is correct, but it means an existing user sees shorter cycles
than a new user until they reset.

## Verification

```
build Release   0 warnings, 0 errors
build Debug     0 warnings, 0 errors
test  Release   539/539 passed (was 528/528)
smoke           scripts/Test-ExcelAddIn.ps1 PASS
                Excel exited, 0 surviving processes, 0 stale session markers
```

The smoke now proves the cycle rather than a fixed format: a General cell
matches no entry, so the first press lands on entry 0
(`$#,##0_);($#,##0)`) and a second press must advance to entry 1
(`$#,##0.00_);($#,##0.00)`). A stuck cycle fails the harness.

New tests worth naming:

- `ProfileV5MigratesNumberFormatsToOneEntryCyclesAndLiftsPropertyCycles`
- `DefaultFontColorCycleCollapsesCategoriesThatShareAColour`
- `ColorReferencesTrackTheCategoryWhileLiteralsStayPinned`
- `AnEmptyCycleAndAnEmptyFamilyAreBothUnrepresentable`
- `NinthCycleInAFamilyIsRefusedNamingTheLimit`
- `SerializedProfileEmitsNoLegacyCycleFieldAndNoNullMember`
- `DefaultCurrencyCycleWalksDollarEuroPoundAtZeroAndTwoDecimals`

The three legacy migration tests were rebased onto a checked-in
`Fixtures/profile-v5.json`, because they previously derived the old shape by
string-replacing current serializer output and v6 no longer emits those keys.

## Traps recorded for whoever comes next

- The legacy DTO members must stay declared forever. `MissingMemberHandling.Error`
  means removing them refuses every existing profile on disk, and
  `NullValueHandling.Ignore` on each is load-bearing: without it every saved v6
  profile carries `"font_color_cycle": null` and fails the unknown-field check on
  the next read.
- An unconfigured family is absent from the JSON rather than present and empty,
  and `ProfileCycle` refuses zero entries. AC-FMT-039's "no phantom slot" is
  therefore unrepresentable rather than filtered at invocation time.
- `€` and `£` are stored as literal characters in the default profile. The
  round-trip trap they belong to is a validation concern, not a storage one; see
  `commands/FORMAT_CYCLES.md`.

## Not done here

AC-FMT-044, the Excel round-trip validation of cycle entries, is untouched. It
needs a live probe and a harness corpus, and belongs with the settings editor
that will call it. AC-FMT-022, AC-FMT-023, and AC-FMT-026..030 belong to WP-F-02
and WP-F-03.
