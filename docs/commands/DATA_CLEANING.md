# Deterministic data-cleaning command contracts

Status: **Approved contract; implementation qualification in progress**
Capability: CAP-DATA-001
Earliest phase: 1B

Implementation note (2026-08-19): text trim/collapse/control removal, fixed
typed text/number/date conversions, and all eight explicit display-value
conversions are registered and qualified through the transactional typed-matrix
adapter. The Ribbon presets expose their separators, grammar, and formats in
preview rather than inferring them from process locale or workbook neighbors. See
`docs/evidence/WP-1B-09_10_DATA_CLEANING.md`.

## Common boundary

- There is no generic `clean` command.
- Each command declares source predicates, exact output, formula policy, locale
  policy, changed/skipped/error categories, and idempotence expectation.
- Formulas are skipped unless a future command explicitly targets formula text.
- Mixed selections require a preview grouped by source type.
- Existing number formats are preserved unless the command explicitly declares
  a number-format change.
- Errors and unsupported conversions remain unchanged and are reported.

## Text normalization

### `clean.text.trim_outer`

- Changed properties: text values only
- Output: remove leading/trailing Unicode whitespace according to the approved
  character table; internal whitespace is unchanged
- Idempotent: yes
- Acceptance: AC-DATA-001 through AC-DATA-004

### `clean.text.collapse_whitespace`

- Output: trim outer whitespace and replace each internal run from the approved
  whitespace table with one ordinary space
- Nonbreaking-space treatment must be explicit in the approved table
- Idempotent: yes
- Acceptance: AC-DATA-001 through AC-DATA-004

### `clean.text.remove_nonprinting`

- Output: remove only code points in the versioned disallowed/control table;
  preserve tabs/newlines only when the selected policy says so
- It MUST NOT remove arbitrary non-ASCII or international characters.
- Idempotent: yes
- Acceptance: AC-DATA-001 through AC-DATA-005

## Typed conversions

### `clean.convert.text_to_number`

- Parameters: explicit decimal separator, thousands separator, sign/parentheses
  policy, currency-symbol policy, percent policy, and blank policy
- Changed properties: values only; optional number format is a separate typed
  parameter shown in preview
- Output: numeric value only when the entire source string matches the declared
  grammar; no partial parse
- Acceptance: AC-DATA-006 through AC-DATA-010, AC-LOC-001

Released Ribbon preset: decimal `.`, thousands `,`, optional leading sign or
negative parentheses, optional leading `$`, optional trailing percent (divides
by 100), and no whitespace. Grouping must be exact and the complete string must
match. The pure planner accepts other explicitly constructed separator/symbol
policies for future typed UI without consulting current culture.

### `clean.convert.number_to_text`

- Parameters: explicit invariant/display format and whether to preserve shown
  decimals
- Output: text value computed from the explicit format; no implicit scientific
  notation or locale guess
- Acceptance: AC-DATA-006 through AC-DATA-010, AC-LOC-001

Released Ribbon preset: invariant `0.################`; output is forced to an
Excel text constant so Excel cannot reinterpret numeric/date-like text.

### `clean.convert.date_normalize`

- Parameters: explicit input pattern(s), calendar/time-zone policy if relevant,
  output as Excel date serial or text, and output format
- Ambiguous free-text dates refuse. Two-digit-year pivot is prohibited unless
  explicitly configured and shown.
- Acceptance: AC-DATA-011 through AC-DATA-014, AC-LOC-001

Released Ribbon preset: accept only `yyyy-MM-dd`, `yyyy/MM/dd`, and `yyyyMMdd`
Gregorian date-only text and output `yyyy-MM-dd` text. It has no time zone and
no two-digit-year pivot. Already canonical text is reported as skipped.

## Explicit display-value conversions

Separate commands are used for each direction:

- `clean.display.blank_to_zero`
- `clean.display.zero_to_blank`
- `clean.display.blank_to_na_text`
- `clean.display.blank_to_nm_text`
- `clean.display.blank_to_dash_text`
- `clean.display.na_text_to_blank`
- `clean.display.nm_text_to_blank`
- `clean.display.dash_text_to_blank`

Contracts:

- Version: 1
- Impact: medium
- Parameters: exact accepted source spellings/case policy and output value
- Changed properties: values only
- Formulas, error values, and nonmatching text are skipped
- `zero_to_blank` matches numeric zero only unless a separately approved option
  includes text zero; it never converts a formula whose result displays zero
- Preview: mandatory, with each source/output category and examples
- Undo: qualified value receipt required before release
- Acceptance: AC-DATA-015 through AC-DATA-019

## Execution and failure behavior

- Snapshot source values and destination properties in a bounded block.
- Build a complete categorized plan before mutation.
- Revalidate source values immediately before write.
- Write changed values in blocks without touching skipped cells.
- Verify output categories and counts.
- Any stale source causes refusal or a fully regenerated preview; it is not
  silently skipped after confirmation.
- Any write failure follows qualified rollback and reports exact remaining cells.
- Performance: PERF-003 for 10,000 cells; larger scopes use PERF-004 behavior.
