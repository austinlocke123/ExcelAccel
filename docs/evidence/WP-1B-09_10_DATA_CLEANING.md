# WP-1B-09/10 deterministic data cleaning evidence

Date: 2026-08-19

## Implemented commands

- `clean.text.trim_outer`
- `clean.text.collapse_whitespace`
- `clean.text.remove_nonprinting` (Ribbon policy preserves tab/CR/LF)
- all eight `clean.display.*` blank/zero/N/A/NM/dash directions

All commands are searchable and exposed in a dedicated Data Cleaning Ribbon
menu with unique KeyTips. Display conversions always require exact-plan preview;
text normalization requires preview for mixed source kinds or threshold-sized
changes.

## Exact source policies

- Text normalization changes text constants only. Formulas, numbers, Booleans,
  blanks, and already-normalized text are skipped and counted.
- The v1 whitespace table is explicit: U+0009..000D, U+0020, U+0085, U+00A0,
  U+1680, U+2000..200A, U+2028, U+2029, U+202F, U+205F, and U+3000.
- Nonprinting removal targets only C0 U+0000..001F and C1 U+007F..009F;
  international text and surrogate pairs remain untouched. Tab/CR/LF retention
  is a typed policy and is enabled by the Ribbon command.
- `zero_to_blank` matches finite numeric zero constants, including negative zero.
  It skips formula results, blank cells, text `0`, and nonzero numbers.
- Text-to-blank commands match only exact case-sensitive `N/A`, `NM`, or ASCII
  `-`; lookalikes and alternative casing are skipped.

## Transaction behavior

The commands reuse the qualified formula/value block transaction: bounded typed
snapshot, exact before/after hashes, stale-source comparison, state-guarded
write, complete postcondition verification, full-block compensation, and
case-sensitive optimistic undo. A receipt store is required before mutation.

The Excel adapter now clears the bounded range before applying its complete
planned matrix. This makes true blanks deterministic; if either the clear or
matrix write fails, the executor restores and verifies the complete before
matrix. Formatting is not changed.

## Verification

- Release suite: **251 passed**, zero failed.
- Debug and Release solution builds: zero warnings, zero errors.
- Unit coverage includes every whitespace family, NBSP behavior, international
  text, surrogate pairs, control-table policy, idempotence, mixed source types,
  formula preservation, numeric-vs-text zero, exact casing/spelling, mandatory
  preview, transaction execution, and undo.
- Packed-XLL hidden Excel smoke uses one mixed row containing Unicode-spaced
  text, a formula, numeric zero, text zero, and a true blank. It verifies:
  - only the text constant is trimmed;
  - the formula is byte-identical;
  - numeric zero becomes a true blank while text zero remains text;
  - two receipts undo the zero conversion and trim in reverse order exactly;
  - the existing formula/format/style/navigation/fault checks still pass; and
  - Excel exits naturally with no surviving process.

## Host findings closed during qualification

- A mixed Excel `Formula` matrix can expose a null formula element with a
  populated `Value2`; capture now uses the populated value rather than calling
  the cell blank.
- Excel can expose a numeric constant's `Formula` element as text such as `0`;
  `Value2` type is authoritative for distinguishing numeric zero from text zero.
- Intermediate worksheet-collection RCWs in off-selection lookup are now owned
  and released explicitly; the borrowed root Application RCW remains untouched.

## Typed-conversion completion checkpoint

The remaining WP-1B-09 slice now registers three fixed, mandatory-preview
commands:

- `clean.convert.text_to_number` uses a complete-string invariant financial
  grammar: decimal `.`, exact `,` grouping, optional sign or negative
  parentheses, optional leading `$`, optional trailing percent, and no
  whitespace/partial parsing;
- `clean.convert.number_to_text` uses the declared invariant
  `0.################` format and forces output to an Excel text constant; and
- `clean.convert.date_normalize` accepts only `yyyy-MM-dd`, `yyyy/MM/dd`, or
  `yyyyMMdd` Gregorian date-only text and outputs canonical `yyyy-MM-dd` text.

The pure parameter contracts also have dot/comma European separator and
`dd.MM.yyyy` fixtures, proving that parsing follows typed policy rather than
process locale. Formulas, unsupported types, invalid grouping, ambiguous or
undeclared dates, partial parses, and already-normalized cells are skipped and
counted. All three commands use the same stale-check/write/verify/compensate and
optimistic-undo transaction.

The mixed-matrix writer now prefixes every planned text constant with Excel's
input apostrophe marker. The marker is not part of `Value2`; it prevents Excel
from silently coercing planned numeric/date-like text during `Formula` array
assignment. Hidden-Excel smoke verifies output CLR types are `System.String`,
formulas/nonmatches remain exact, three reverse-order receipts restore the full
source matrix, and Excel exits naturally.

Current Release suite after this slice: **278 passed**, zero failed. Debug and
Release builds have zero warnings/errors.
