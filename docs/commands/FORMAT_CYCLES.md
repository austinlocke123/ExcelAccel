# Number format cycles and user-defined cycles

Status: **Approved for implementation, 2026-08-20**
Capability: CAP-FMT-004
Supersedes: the one-shot behaviour of `format.number.*` in
[`CATALOG.md`](CATALOG.md) §1.1

## Why this changes

Named number-format commands currently apply one profile format and stop.
Pressing Percentage twice does nothing the second time. The reference tools in
this category instead **cycle** a family of related formats, so repeated presses
walk decimal and presentation variants without the user leaving the keyboard.

This also removes hard-coded formats from the product. Every cycle, including
which cycles exist at all, becomes a profile setting the user can edit and
extend.

## Cycle semantics

A cycle is an ordered list of number-format strings. Advancing is **stateless**:
position is derived from what the cell currently holds, never from remembered
state.

1. Read the selection's current number format.
2. Find its index in the cycle, compared case-insensitively against the exact
   stored format string.
3. Apply the entry at `index + 1`, wrapping to the start at the end.
4. If the current format is **not** in the cycle, apply entry 0.

Statelessness is the point: there is no cursor to desynchronise across sessions,
selections, undo, or two windows on the same workbook. The cell is the state.

### Family independence

Each cycle is independent. Invoking Percent on a currency-formatted cell enters
the percent cycle at entry 0; it does not attempt to carry decimal position
across families. First press lands in the family, subsequent presses walk it.

### Mixed selections

When the selected cells do not share one number format there is no single current
value, so the selection is treated as **no match** and entry 0 is applied to
every cell. The first press normalizes the block; the second advances it
together. This is the approved mixed-selection rule of AC-FMT-001 applied to
number formats.

### Relationship to the decimals commands

`format.number.decimals.increase` and `.decrease` are retained. Cycles cover the
common path; the decimals commands operate on **any** format, including formats
that belong to no cycle, and remain the only way to adjust precision on an
arbitrary custom format.

## Negative numbers and horizontal alignment

Negatives are shown in parentheses. Default cycle entries carry an explicit
negative section rather than relying on the locale's default negative
presentation.

Parenthesised negatives create an alignment problem: the closing paren occupies
a character width that the positive form does not, so a column of mixed signs
does not line up at the right edge. Every default entry in the General,
Currency, Percent, and Multiple cycles therefore ends its positive section with
`_)`, which reserves exactly the width of a closing paren without printing one.

Measured, rather than assumed:

| Format | Value | Displays as |
|---|---|---|
| `#,##0;(#,##0)` | 1234 | `1,234` |
| `#,##0_);(#,##0)` | 1234 | `1,234 ` |
| `#,##0_);(#,##0)` | -1234 | `(1,234)` |
| `$#,##0_);($#,##0)` | -1234 | `($1,234)` |
| `0.0%_);(0.0%)` | 0.125 | `12.5% ` |
| `0.0%_);(0.0%)` | -0.125 | `(12.5%)` |
| `0.0x_);(0.0x)` | -2.5 | `(2.5x)` |

The four families align on the same principle, so switching a column between
General, Currency, Percent, and Multiple does not shift the digits.

**The alignment is delivered by the format string, never by setting the cell's
horizontal alignment property.** A number-format command changes the number
format and nothing else, per AC-FMT-024. Reaching across to mutate a second
property would break the property-scope discipline the whole command set rests
on, and would silently overwrite a deliberate alignment choice the user made.
Date and Binary carry no padding, having no parenthesised negative form.

## Basis points

**A scaling basis-point display format is not achievable in Excel number
formatting, and the product will not pretend otherwise.** This was measured
rather than assumed:

| Format | Value 0.0125 displays as |
|---|---|
| `0.00%` | `1.25%` |
| `0%%` | `125%%` |
| `0.00%%` | `125.00%%` |
| `0"bps"` | `0bps` |
| `0%%" bps"` | `125%% bps` |

Each `%` multiplies the displayed value by 100 **and** prints a `%` character.
The two are inseparable, so a format can scale by 10,000 or print a clean
`bps` suffix, but not both.

Basis points are therefore supported two ways, both honest:

- **As a unit transform.** A `formula.units.to_basis_points` command multiplies
  the formula by 10,000 and applies a `0" bps"` format, matching the existing
  to/from thousands and millions commands. This changes the value, so it carries
  the same impact tier, preview, and undo receipt as the other unit transforms.
- **As a cycle entry for values already stored in basis points.** A model that
  holds a spread as `125` can use a `0" bps"` cycle entry directly, because no
  scaling is required.

The percent cycle itself does not include a scaling bps entry, because none
exists that displays correctly.

## Default cycles

Defaults only. Every entry is editable, and cycles may be added or removed.

| Cycle | Entries |
|---|---|
| General | `#,##0_);(#,##0)` → `#,##0.0_);(#,##0.0)` → `#,##0.00_);(#,##0.00)` |
| Currency | `$#,##0_);($#,##0)` → `$#,##0.00_);($#,##0.00)` → `€#,##0_);(€#,##0)` → `€#,##0.00_);(€#,##0.00)` → `£#,##0_);(£#,##0)` → `£#,##0.00_);(£#,##0.00)` |
| Percent | `0.0%_);(0.0%)` → `0.00%_);(0.00%)` → `0%_);(0%)` |
| Multiple | `0.0x_);(0.0x)` → `0.00x_);(0.00x)` → `0x_);(0x)` |
| Date | `m/d/yyyy` → `mmm-yy` → `mmmm d, yyyy` |
| Binary | `"TRUE";"TRUE";"FALSE"` → `"YES";"YES";"NO"` → `0;0;1` |

The Currency cycle walks decimals within a symbol, then moves to the next
symbol: dollar zero-decimal, dollar two-decimal, euro zero, euro two, pound
zero, pound two. A user working in one currency stays in the first two entries;
a user working across currencies reaches all six from one key.

This replaces the previous third entry, a `$#,##0.0,,"m"` millions format. Unit
scaling belongs to the `formula.units.*` transforms, which change the value and
carry preview and undo, not to a display format that silently divides what the
reader sees by a million.

### Cycle entries must survive Excel's round-trip

A cycle finds its position by comparing the cell's **stored** number format
against its entries. If Excel rewrites a format string on assignment, the stored
string never matches the entry that produced it, the comparison always misses,
and the cycle sticks on entry 0 forever.

Excel does exactly this to locale-qualified currency formats. Measured by
writing a format, reading it back, and comparing code points:

| Written | Stored back | Round-trips |
|---|---|---|
| `£#,##0_);(£#,##0)` | unchanged | yes |
| `€#,##0_);(€#,##0)` | unchanged | yes |
| `[$£-en-GB]#,##0_);(...)` | `[$£-809]#,##0_);(...)` | **no** |
| `[$€-x-euro2]#,##0_);(...)` | `[$€-2]#,##0_);(...)` | **no** |

**The defaults therefore use bare currency symbols.** The locale-qualified forms
are the ones a user is likely to reach for, and are what Excel's own currency
dialog produces, which makes this a trap rather than an obvious mistake.

Two consequences for implementation:

- Cycle entry validation must write each format to a scratch cell, read it back,
  and reject any entry Excel does not store verbatim, naming the rewritten form
  so the user can adopt it instead. Validating syntax alone would accept these.
- Bare symbols are display-only and do not carry locale currency behaviour. That
  is acceptable here because the cycle's job is presentation; a model needing
  true locale currency semantics sets the format directly.

## Where cycle defaults live

Defaults ship as **data, in `config/default-profile.json`**, which is embedded as
a resource and read by `ProfileStore`. This is already how the existing colour,
font-size, alignment, underline, row-height, and column-width cycles are
defined, so the starting point exists and needs extending rather than inventing.

The rule, stated so it does not erode:

- **The default profile is the only place a cycle's contents are written.** No
  C# file contains a format string, a colour, or a cycle order.
- **The settings editor edits the active profile**, never the embedded default.
- **Reset restores from the embedded default**, which is what makes it a
  meaningful starting point rather than a one-time seed.

So the answer to "settings feature or codebase" is both, with a clean split:
the codebase ships the starting point as data, and the settings feature is the
only way to change it. What must not happen is a cycle defined in code, because
then a user editing it in settings is overriding something they cannot see, and
the two drift apart.

## User-defined cycles

Cycles are **data, not code**. The profile owns an ordered, named collection;
the product ships defaults and nothing more.

- A user may reorder, edit, add, or remove entries within any cycle.
- A user may **add a cycle that did not exist when the product shipped**, and
  remove one that did.
- A cycle with no entries **does not exist as far as the user is concerned**.
  See "No phantom slots" below.
- Every format string is validated before it is stored; an invalid string is
  refused with its position in the cycle, and the prior profile is preserved.

### Invoking a user-added cycle

Ribbon XML is static, so a cycle invented after install cannot receive a
purpose-built button. Two mechanisms cover it, and both are required:

- **Custom cycle slots.** Up to eight per family, each bound to a named cycle in
  the profile. Eight is the ceiling, not the count: a family with two configured
  cycles has two, not two plus six empties.
- **Command Search.** Every configured cycle is searchable by its user-given
  name, so a ninth cycle is still reachable even with no free slot.

The built-in six are simply the cycles the default profile happens to define;
they carry no privilege beyond having dedicated ribbon buttons.

### No phantom slots

An unconfigured slot must be **invisible and unreachable**, not merely inert.
Pressing a key and getting "nothing is configured" is the same interruption the
no-dialogs-on-success rule exists to remove, and cycling through empty
placeholders to reach a real one defeats the point of a cycle.

Concretely, a slot with no entries:

- does not appear on the ribbon;
- is not offered by Command Search;
- is skipped entirely when a cycle advances, never occupying a position;
- refuses with a message naming the slot, and changes nothing, if reached by a
  stored route or cheat-sheet entry that predates its deletion.

The last case is the only one where a message is correct, because the user
invoked something specific by name and deserves to know why nothing happened.

### Applies to every cycle family

This is not a number-format rule. The profile already defines cycles for font
colour, fill colour, font size, horizontal and vertical alignment, underline,
row height, and column width, and border cycles are still to be built. All of
them take the same treatment: user-definable, up to eight custom slots per
family, managed from the same settings editor, with no phantom slots.

A family's built-in cycle is just its default; it may be edited or removed like
any other.

## Settings surface

Editing cycles by hand-writing profile JSON is not acceptable for a setting this
central, so this work package includes a settings editor.

- Lists every family, and within it every configured cycle with its entries in
  order.
- Add, remove, rename, and reorder both cycles and entries, in any family.
- Enforces the eight-per-family ceiling, refusing a ninth with a message naming
  the limit rather than silently dropping it.
- Makes deletion the way a slot becomes unconfigured, so the user never has to
  leave an empty placeholder behind to get rid of a cycle.
- Shows a live preview of a representative positive and negative value for each
  entry, so the effect is visible before saving.
- Validates every format string on save and refuses invalid input with its
  position.
- Writes through the existing validate-then-atomic-replace profile path, leaving
  the prior profile intact on any failure.
- Reachable from the Settings group and from Command Search.

## Profile schema

`ProfileDefinition.NumberFormats`, today a single format per family, becomes an
ordered named collection of ordered format lists. This is a breaking schema
change: version 5 to 6, with a migration that lifts each existing single format
into a one-entry cycle so no user loses a setting.

Profile export and import must carry cycles, and an import whose cycles fail
validation is refused whole, per AC-PROF-007 and AC-PROF-009.

## Acceptance

| ID | Criterion |
|---|---|
| AC-FMT-021 | A cycle advances from the selection's current format to the next entry and wraps at the end, deriving position from the cell rather than stored state. |
| AC-FMT-022 | A current format absent from the cycle applies entry 0, and a mixed selection applies entry 0 to every cell. |
| AC-FMT-023 | Cycles are independent; entering a cycle from a different family starts at entry 0 and carries no position across families. |
| AC-FMT-024 | Cycling changes only the number format, never the underlying value or formula, and records an undo receipt. |
| AC-FMT-025 | Default cycle entries present negatives in parentheses, and General, Currency, Percent, and Multiple entries pad the positive section with `_)` so mixed-sign columns align. |
| AC-FMT-043 | The default Currency cycle walks dollar, euro, and pound at zero and two decimals, in that order. |
| AC-FMT-044 | Cycle entry validation writes each format to Excel and reads it back, rejecting any entry Excel does not store verbatim and naming the rewritten form. |
| AC-FMT-045 | No cycle contents are defined in code; every default resolves from the embedded default profile, and reset restores from it. |
| AC-FMT-026 | A user may add, remove, reorder, and edit cycles and their entries, including cycles that did not ship with the product. |
| AC-FMT-027 | An invalid format string is refused with its cycle and position, and the prior profile remains active. |
| AC-FMT-028 | A slot reached by a stored route or cheat-sheet entry after its cycle was deleted refuses with a message naming the slot and changes nothing. |
| AC-FMT-029 | Custom cycle slots and Command Search both invoke a user-added cycle by its user-given name. |
| AC-FMT-030 | The settings editor previews a positive and negative sample per entry and writes through validate-then-atomic-replace. |
| AC-FMT-031 | Migration from schema 5 lifts each existing single number format into a one-entry cycle with no setting lost. |
| AC-FMT-032 | `formula.units.to_basis_points` multiplies by 10,000, applies a basis-point format, and carries the same impact tier, preview, and undo as the other unit transforms. |
| AC-FMT-039 | An unconfigured cycle slot does not appear on the ribbon, is not offered by Command Search, and is skipped when a cycle advances, so no press ever lands on an empty slot. |
| AC-FMT-040 | The settings editor adds, deletes, renames, and defines cycles in every family, enforcing eight per family and refusing a ninth with a message naming the limit. |

## Explicitly out of scope

- A scaling basis-point *number format*, which Excel cannot express.
- Conditional or rule-driven formats; a cycle is an ordered list, nothing more.
- Conditional or rule-driven cycle entries; an entry is a value, not a rule.

Note that cycling properties beyond the number format is **in** scope as of
2026-08-20; see "Applies to every cycle family" above. The earlier statement
that the colour, font, alignment, row-height, and column-width cycles were
unchanged no longer holds — they become user-definable on the same terms.
