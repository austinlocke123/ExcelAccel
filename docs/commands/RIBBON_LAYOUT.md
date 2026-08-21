# Ribbon layout and keyboard routing

Status: **Reviewed and approved 2026-08-20. Regrouping to the task taxonomy
below is approved and not yet implemented.**
Capability: CAP-UX-002
Acceptance: AC-FMT-033

## Why this document exists

The ribbon was assembled incrementally, one work package at a time, and its
structure was never specified anywhere. Everything ended up in a single group
called "Safe Tools" — a Phase 0 placeholder — holding 110 buttons behind nine
menus. Layout is a product decision and needs a home, or it drifts again.

## Principles

1. **Group by modeling task, not by implementation.** Users think in
   "numbers", "colors", "borders", not in "the formatting catalog".
2. **Shallow beats tidy.** A keyboard-first tool should reach a frequent command
   in as few keystrokes as possible, so frequent commands are direct buttons and
   only infrequent ones live in menus. More groups is an acceptable price.
3. **Frequency orders the tab, left to right**, because leftmost is nearest the
   tab's own entry point.
4. **Read-only and mutating commands are not mixed** within a group where the
   distinction matters to the user.
5. **Every command remains reachable from Command Search**, so the ribbon never
   has to carry everything.

## Group taxonomy

Derived from how modeling tools in this category organise the same territory,
and confirmed against a reference product's shortcut taxonomy.

| Group | Holds |
|---|---|
| Find | Search Commands, Inspect Selection, Undo ExcelAccel |
| Modeling | Smart Copy down/right, spacing, fills, IFERROR, reverse sign, unit transforms |
| Paste | Capture source, paste formulas/values/formats, transpose |
| Numbers | The number-format cycles, increase/decrease decimals |
| Colors | Font colour cycle, fill colour cycle, AutoColor Selection, AutoColor Worksheet when qualified |
| Fonts | Font size cycle, underline cycle |
| Alignment | Horizontal and vertical cycles, indent, center across |
| Borders | Sum bar, remove borders, and the border cycles when built |
| Rows and Columns | Row height, column width, autofit, and group/ungroup/expand/collapse when built |
| Auditing | Precedents, dependents, formula inspector, Model Check |
| Clean and Select | Data cleaning, deterministic selection |
| View and Navigate | Navigation, freeze, gridlines, zoom |
| Styles | Built-in styles and the style library |
| Settings | Profile export/import, cheat sheet, diagnostics, cycle editor |

**Superseded structure.** The eight-group layout shipped on 2026-08-20 fused
Colors, Fonts, Alignment, Borders, and Rows and Columns into one "Cell Format"
group, and fused Modeling with Paste into "Quick Formulas". That collapsed
distinctions users hold separately and buried rows and columns two levels deep.
The taxonomy above replaces it.

## Keyboard routing

KeyTips are assigned under one hard constraint: **Excel resolves a KeyTip as
soon as it is unambiguous, so a single-letter KeyTip may never be a prefix of a
longer one.** If `S` is assigned, no `S?` KeyTip is reachable.

The scheme is therefore:

- frequent commands take single letters;
- menus take reserved two-letter prefixes whose first letter is never used
  alone;
- sub-items within a menu need only be unique within that menu.

A validator parses the ribbon XML and fails on any duplicate KeyTip, any prefix
collision, and any button without an action. It lives in `RibbonRouteTests` and
was built in WP-F-07; before that this paragraph described behaviour that did
not exist.

### Routes are hand-maintained, and checked against the ribbon

`RibbonRoutes` is written by hand, not generated. Every command descriptor must
read its route from there rather than composing one from a KeyTip fragment;
catalogs that composed their own let a descriptor disagree with the ribbon
silently, so Command Search and the cheat sheet printed routes that did nothing.
Ten descriptors had drifted that way before WP-F-07 checked.

`RibbonRoutes.For` throws on an unknown id. It used to fall back to the Command
Search route, which meant a typo produced a descriptor that looked routed and
advertised a path that did nothing.

Update `RibbonRoutes` whenever the ribbon changes; `RibbonRouteTests` fails if
you forget, and compares each descriptor's route against the path its button
actually has.

Commands hosted in a dialog rather than on the ribbon carry an honest route
describing how they are actually reached, for example
"Search Commands, then Ctrl+D".

## Extending the ribbon

A new command must state which group it belongs to. If it fits none, that is
evidence the taxonomy needs a new group, not that the command should be dropped
into the nearest one.

Ribbon XML is static, so anything user-defined after install cannot get a
purpose-built button. Such features need fixed slots plus Command Search; see
[`FORMAT_CYCLES.md`](FORMAT_CYCLES.md) for how the user-defined number-format
cycles handle this.
