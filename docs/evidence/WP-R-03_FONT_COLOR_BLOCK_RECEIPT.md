# WP-R-03 Font-colour block receipt

Date: 2026-08-23
Status: Complete. The undo ceiling that blocked AutoColor execution is resolved.

## Contract

- **Capability:** CAP-REL-001, CAP-FMT-002
- **Acceptance:** AC-REL-011, AC-REL-012 (undo semantics for a bounded recolour)
- **Allowed implementation:** the block value, its store recognition, and tests.
- **Excluded:** the AutoColor adapter and command registration, which follow.

## The blocker, and why the obvious fix was wrong

`PropertyBatchReceipt` caps at 32 changes. A recolour touches thousands of cells,
so one change per cell cannot work, and AutoColor could not record undo at all.

Raising the cap was the wrong instinct. The product already solves this twice:
`cell_contents_v1` carries a whole formula block and `cell_format_block_v1`
carries a whole format block, each as **one coarse property**. The cap is not an
obstacle to work around; it is a statement that a receipt holds a few properties,
not many cells.

AutoColor now follows the same shape with `cell_font_color_block_v1`. A recolour
of any size is one change on one receipt, so a single undo restores all of it or
none, and the 32-change limit becomes irrelevant rather than raised.

## The value

Addresses are grouped under their colour, `#RRGGBB=A1,A2;#000000=B7`, because a
real model holds a handful of distinct font colours across thousands of cells.
5,000 cells sharing two colours serialize into two groups.

It is exact, not approximate: no colour is rounded and no cell is dropped.
Colours are written in ordinal order and addresses in ordinal order within each,
so two captures of one state are byte-identical. That matters because undo
compares the stored value against what it reads back, and the block is compared
**ordinally** like the other two coarse properties, since its exact bytes are the
value.

## The ceiling is reasoned, not guessed

`FontColorBlock.MaximumCells` is 50,000, matching the existing formatting
ceiling. Worst case is every cell carrying a distinct colour, about 16 characters
each, so 800,000 characters against the store's 1,000,000-character per-value
limit.

`WorstCaseCharacters` exists so the two constants stay tied together, and
`TheCellCeilingKeepsTheWorstCaseInsideTheStoreLimit` fails if either moves
without the other being reconsidered. A separate test serializes 20,000
genuinely distinct colours and asserts it still fits.

## A bug its own tests caught

Duplicate-address detection was per colour group. The same cell listed under two
different colours went into two different groups, so neither noticed, and the
block would have serialized that cell twice and deserialized it as one cell with
two conflicting colours. Undo would then have written whichever it read last.

Addresses are now tracked across every colour. Separators inside an address are
refused for the same class of reason: an address containing `;`, `,`, or `=`
would silently split into two cells on the way back.

## Verification

```
build Release   0 warnings, 0 errors
build Debug     0 warnings, 0 errors
test  Release   723/723 passed (was 703/703)
```

No smoke run: nothing reachable from Excel changed yet.

`UndoRestoresTheWholeBlockAndDetectsAnInterveningChange` and
`UndoRefusesWhenTheBlockChangedSincePlanning` drive the value through the real
`SessionUndoStore`, so the stale check is exercised rather than assumed.
