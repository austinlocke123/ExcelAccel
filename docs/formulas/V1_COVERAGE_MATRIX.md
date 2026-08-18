# Proposed v1 formula coverage matrix

- Work package: WP-P0-05
- Status: **Prototype evidence; not accepted for production mutation**
- Corpus: [`formula-v1-corpus.json`](../../tests/ExcelAccel.Core.Tests/Fixtures/formula-v1-corpus.json)
- Governing decision: [ADR-0004](../decisions/ADR-0004-formula-strategy.md)

## Dispositions

| Disposition | Meaning |
|---|---|
| `transform` | The prototype produces a lossless syntax document and explicit reference model. Production transformation is still prohibited until ADR-0004 is accepted. |
| `round_trip` | The syntax is preserved exactly, but the formula contains no reference currently eligible for transformation. |
| `inspect_only` | Parsing and exact serialization are allowed; any mutating consumer must refuse with the recorded limitation code. |
| `refuse` | Parsing returns no document and a stable coverage-specific reason. Partial syntax is never exposed. |

## Current matrix

| Syntax dimension | Prototype disposition | Stable limitation/refusal |
|---|---|---|
| A1 relative, absolute, and mixed cells | `transform` | — |
| A1 rectangular ranges | `transform` | — |
| R1C1 current, absolute, and bracket-relative cells/ranges | `transform` in the pure parser | native lifecycle qualification failed on Excel 16.0 build 20228; no production use |
| Unquoted and escaped quoted worksheet qualifiers | `transform` | — |
| String literals, numeric literals, error literals, operators, functions, and exact whitespace | `round_trip` or `transform` when supported references are present | — |
| Configured list/decimal separators | `transform` in the pure parser | real-Excel locale qualification pending |
| External workbook qualifiers | `inspect_only` | `FORMULA_EXTERNAL_REFERENCE_INSPECT_ONLY` |
| Workbook/worksheet names | `inspect_only` | `FORMULA_NAME_INSPECT_ONLY` |
| Structured table references | `inspect_only` | `FORMULA_STRUCTURED_REFERENCE_INSPECT_ONLY` |
| Spill and implicit-intersection operators | `inspect_only` | `FORMULA_DYNAMIC_ARRAY_INSPECT_ONLY` |
| Reference intersections | `inspect_only` | `FORMULA_INTERSECTION_INSPECT_ONLY` |
| Top-level reference unions | `inspect_only` | `FORMULA_UNION_INSPECT_ONLY` |
| Array constants and legacy array wrappers | `refuse` | `FORMULA_ARRAY_SYNTAX_UNSUPPORTED` |
| Wrong locale separator | `refuse` | `FORMULA_DIALECT_MISMATCH` |
| A1 reference beyond XFD/1,048,576 or invalid R1C1 coordinate | `refuse` | `FORMULA_INVALID_REFERENCE` |
| Unbalanced, incomplete, ambiguous, or unknown token sequence | `refuse` | structured syntax/resource reason |

## Resource and immutability contract

- Default formula length: at most 8,192 UTF-16 code units.
- Default token count: at most 4,096 lossless tokens.
- Default parenthesis nesting: at most 64 levels.
- Callers may lower limits. Hard ceilings are 32,768 characters, 16,384 tokens,
  and 256 nesting levels, so configuration cannot make the parser unbounded.
- Limits are checked during the single forward scan; there is no recursive
  descent, COM access, dependency loading, or mutation.
- A successful document contains immutable token and reference collections,
  source spans, dialect, disposition, and one stable limitation code when it is
  inspect-only.
- A refusal contains no document. Callers cannot accidentally operate on a
  partial parse.
- Serialization returns the exact original source. Canonical normalization is
  intentionally separate and not implemented by this spike.

## Required work before ADR acceptance

- expand the approved corpus across real workbook formulas and hostile fuzz
  minimization;
- validate `Formula`/`Formula2`, `FormulaR1C1`, and local dialect behavior on the
  supported Microsoft 365/LTSC build and locale matrix;
- decide and qualify the expression AST used by semantic transforms;
- compare any proposed third-party parser for license, maintenance, security,
  offline availability, grammar coverage, and round-trip fidelity;
- benchmark cold/warm distributions and memory on the reference corpus;
- prove every inspect-only/refuse disposition at the command mutation boundary.
