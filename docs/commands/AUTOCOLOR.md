# Cell classification and AutoColor

Status: **Reviewed and approved 2026-08-20, including the command surface.
Not implemented: the planner exists and execution is hard-stopped.**
Capability: CAP-FMT-002
Related: [`FORMAT_CYCLES.md`](FORMAT_CYCLES.md),
[`MODEL_CHECK.md`](MODEL_CHECK.md)

## What this governs

One classification rule feeds **two commands, and only two**:

- **AutoColor Selection** — recolours the current selection, whether that is a
  single cell or a contiguous block.
- **AutoColor Worksheet** — recolours the used range of the active worksheet.

Both are invoked by the user, deliberately, one press at a time. They share one
rule so a cell gets the same colour whichever command reaches it.

There is no workbook scope. `AutoColorScope` in the planner is exactly
`{ Selection, Worksheet }`, and that is the whole surface.

### There is no third toggle command

An earlier draft of this document described a separate "blue-black toggle" as
the single-cell case of the rule. That is precisely what AutoColor Selection
already does on a one-cell selection, so it is **not** a separate command. Two
commands that classify identically and differ only in name would be a
maintenance trap and a second thing to keep in sync.

A user wanting to set a colour **against** the classification — marking a cell
blue that AutoColor would call black — uses the font colour cycle, whose entries
include the category references described below. That is the deliberate override
path, and it is manual by construction.

## Categories and colours

The default palette follows the convention these models are read with. Every
colour is a profile setting.

| Category | Default | Meaning |
|---|---|---|
| Numeric hardcode | blue `#0000FF` | The cell contributes a typed number |
| Same-sheet formula | black `#000000` | Derived from this worksheet only |
| Cross-sheet formula | green `#008000` | Pulls from another worksheet |
| External formula | red `#FF0000` | Pulls from another workbook |
| Error | red `#FF0000` | The cell holds an Excel error |
| Text | black `#000000` | Text constant |

Every one of these is a **default, not a rule**. The user may set any category to
any colour from the settings editor, including making two categories share a
colour or setting one to the same black as formulas to switch it off visually.
The product hard-codes no colour; the table above is what ships in the default
profile and nothing more.

## The colour cycle and AutoColor share a palette

The font colour cycle and the AutoColor categories draw on one palette, because
a user who recolours a cell by hand is usually reaching for the same meanings
AutoColor assigns automatically.

A cycle entry is therefore one of two things:

- **A category reference** — `@hardcode`, `@same_sheet`, `@cross_sheet`,
  `@external`, `@error`, `@text`. It resolves to whatever that category's colour
  currently is.
- **A literal colour** — a hex value belonging to no category, for colours the
  user wants on the cycle but that carry no classification meaning.

The default font colour cycle is the six category references in precedence
order, and the user may add literals or reorder freely.

**References track, literals do not.** Changing the hardcode colour from blue to
navy changes it everywhere at once: both AutoColor commands and the cycle
entry. This is the reason for references rather than copied hex values —
a palette that silently drifts out of step with AutoColor would produce
hand-coloured cells that look like classified ones but aren't, which is the
exact confusion the colouring exists to prevent.

A user who wants a colour pinned against future category changes uses a literal.

## Classification precedence

Evaluated in order; the first match wins.

1. **Error** — the cell holds an Excel error value.
2. **Numeric hardcode** — the cell is a numeric constant, **or** its formula
   embeds any numeric literal.
3. **External formula** — references another workbook.
4. **Cross-sheet formula** — references another worksheet.
5. **Same-sheet formula** — references only this worksheet.
6. **Text** — a text constant.

### Hardcode outranks everything below it

A hardcode is blue **even when the same formula is also cross-sheet or
external**. `='[Other.xlsx]Sheet1'!A1+5` is **blue**, not red.

This is deliberate and was chosen explicitly. The point of the colouring is to
make typed numbers findable; a number buried inside an external reference is
exactly the kind that hides. Losing the red external signal on that one cell is
an accepted cost, and the external-reference Model Check rule still reports it
independently.

### No allowlist

**Any** numeric literal in a formula makes the cell a hardcode. `=A1*2` is
**blue**.

This is a deliberate divergence from Model Check. `check.formula.embedded_numeric_constant`
applies a versioned allowlist (`0, 1, -1, 2, 100, 12, 365`) and structural
exclusions, because its job is to raise findings worth a person's attention and
an allowlist keeps the noise down. AutoColor's job is different: it is a visual
map, where a missed hardcode is worse than an over-coloured one.

**Do not unify these two rules.** They answer different questions and are
correct to disagree. The allowlist belongs to the finding; the strictness
belongs to the colour.

Consequence to expect: on a typical model, more cells turn blue than a Model
Check scan reports as findings. That is the intended behaviour, not a defect.

## Both commands are user-invoked, always

**Nothing in this feature watches cells.** There is no `Worksheet_Change`
handler, no `SheetChange` or `SheetCalculate` subscription, and no timer that
recolours in the background. A cell's colour changes when, and only when, the
user invokes one of the two commands.

This is a firm product constraint, not an implementation convenience:

- Live recolouring would mutate the workbook while the user is typing, making
  every keystroke a potential undo entry and defeating the single custom undo
  slot the product relies on.
- A model in progress is *full* of cells that are temporarily hardcodes. Colour
  that reacts instantly turns normal editing into a flicker of changing colours
  and trains the user to ignore it, which destroys the signal.
- Unrequested writes to a workbook the user did not ask to modify is exactly
  the behaviour that makes add-ins untrustworthy.

Both commands write only the font colour, record an undo receipt, and never
touch the value, the formula, or any other format property.

## Two decisions, confirmed 2026-08-20

Both were raised for review and approved as written:

1. **Error is placed above hardcode.** A cell holding `#REF!` is red even if its
   formula also contains a literal, on the grounds that a broken cell needs
   fixing before its inputs matter.
2. **Only numeric literals count as hardcodes.** A text literal does not, so
   `=IF(A1>0,"Yes","No")` is a formula and stays black. This does mean hardcoded
   labels and thresholds expressed as text go uncoloured; that was accepted.

Note that under the no-allowlist rule, `=DATE(2026,1,1)` is blue, because its
arguments are numeric literals.

## Scope boundary

The two commands carry very different risk, and are gated differently.

**AutoColor Selection is bounded** by the existing selection ceiling and can
ship now. A user selecting a block has already scoped the change themselves.

**AutoColor Worksheet changes thousands of cells in one action.** The planner is
deterministic, address ordered, bounded at 250,000 cells, carries a complete
precondition fingerprint, and requires preview at this scope, but execution
stays refused with `PERFORMANCE_QUALIFICATION_REQUIRED` until it has a
transactional adapter, rollback and fault-injection evidence, and a
worksheet-scale preview UI. Nothing else in the product mutates that many cells
at once, and a half-applied recolour is exactly the partial success the
architecture refuses everywhere else.

Shipping the selection command first is therefore not a compromise: it is the
whole feature for anyone working a block at a time, available while the
worksheet command earns its qualification.

## Acceptance

| ID | Criterion |
|---|---|
| AC-FMT-034 | Classification follows the stated precedence, and a cell containing any numeric literal is a hardcode regardless of cross-sheet or external references in the same formula. |
| AC-FMT-035 | No allowlist is applied; `=A1*2` classifies as a hardcode while Model Check continues to exclude the same literal from its findings. |
| AC-FMT-036 | AutoColor Selection and AutoColor Worksheet assign the same colour to the same cell, and no third command applies the classification. |
| AC-FMT-037 | Both commands write only the font colour, record an undo receipt, and leave value, formula, and other format properties unchanged. |
| AC-FMT-046 | Colouring occurs only on explicit user invocation; the add-in registers no change, calculate, or timer handler that recolours cells. |
| AC-FMT-038 | Every category colour resolves from the active profile, with no colour hard-coded in the product, and each is editable per category from the settings editor. |
| AC-FMT-041 | A colour cycle entry may be a category reference or a literal colour, and the default font colour cycle is the six category references in precedence order. |
| AC-FMT-042 | Changing a category's colour updates both AutoColor commands and every cycle entry referencing it, while literal entries are unaffected. |
