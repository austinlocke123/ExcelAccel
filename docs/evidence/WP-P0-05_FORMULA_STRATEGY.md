# WP-P0-05 formula strategy spike evidence

- Date: 2026-08-18
- Scope: first WP-P0-05 implementation slice
- Status: prototype passing; ADR-0004 and AC-P0-005 remain open
- Coverage matrix: [`V1_COVERAGE_MATRIX.md`](../formulas/V1_COVERAGE_MATRIX.md)

## Implemented prototype

The pure `ExcelAccel.Core.Formulas` boundary now provides:

- explicit A1/R1C1 dialect and locale separator inputs;
- a lossless token stream with source spans;
- explicit reference nodes with qualifier, range, spill/intersection flags, and
  absolute/relative/current row and column coordinates;
- exact deterministic serialization;
- stable transform, round-trip, inspect-only, and refusal outcomes;
- strict formula length, token-count, and nesting limits;
- structural validation that rejects incomplete or ambiguous token sequences;
- no Excel COM types, host callbacks, added packages, or mutation APIs.

## Corpus and automated evidence

The versioned JSON corpus currently covers 29 named golden/refusal cases across
A1, R1C1, mixed/absolute references, ranges, quoted sheets, external workbooks,
names, structured references, spills, implicit intersection, intersections,
unions, locale separators, errors, strings, array syntax, invalid references,
and malformed expressions.

Tests additionally cover:

- 500 deterministically generated supported formulas;
- exact token-span continuity and serialization;
- A1 and R1C1 coordinate semantics;
- concurrent pure-core parsing;
- hostile length, token, nesting, string, character, and expression inputs;
- defensive-copy immutability and non-overridable hard resource ceilings;
- the invariant that refusals expose no partial document.

The real-Excel oracle harness loads the same corpus and currently qualifies
eight non-mutating temporary-workbook A1 cases through native `Formula`. It
requires exact formula round-trip, workbook closure, `Excel.Quit`, and exit of
the recorded Excel PID. The clean run recorded Excel version 16.0, build 20228.

An early oracle run proved formula fidelity but exposed test-harness COM
ownership gaps: `Workbooks` and `Worksheets` collection wrappers had not been
retained for deterministic release. Only the exact two test-owned Excel PIDs
were terminated. The harness now explicitly releases cell, worksheet,
worksheets, workbook, workbooks, and application wrappers in child-to-root
order; the corrected A1 run exits cleanly.

A separate non-circular R1C1 case (`=R[-1]C[2]+R1C1`) round-tripped exactly
through `FormulaR1C1`, but Excel 16.0 build 20228 did not terminate after
`Excel.Quit` even after the collection-wrapper correction. The exact test-owned
PID was terminated. R1C1 therefore remains unqualified for native adapter use;
the case stays in the pure corpus but is excluded from the passing Excel oracle
set until lifecycle isolation explains and resolves the behavior.

## Strategy finding

The prototype supports a hybrid recommendation: keep lossless syntax and
reference ownership in a deterministic pure-core model, add a separately
qualified semantic expression AST, and use Excel only as a qualification oracle
or narrowly bounded adapter where native translation is demonstrably safer.

A regex-only transform strategy is rejected: pattern matching is used only to
recognize bounded reference tokens and is followed by coordinate validation,
lossless tokenization, and structural validation. Raw string replacement remains
prohibited.

## Qualification result

- Debug solution build: passed with zero warnings and zero errors.
- Release solution build: passed with zero warnings and zero errors.
- Debug tests: 37 passed, 0 failed.
- Release tests: 37 passed, 0 failed.
- A1 native oracle: eight exact round trips on Excel 16.0 build 20228 with
  workbook close, `Excel.Quit`, and clean PID exit.
- Packed debug XLL regression: registration, health UDF, guarded formatting,
  refusal cases, injected-fault restoration, workbook close, and clean PID exit
  all passed after embedding the formula prototype in the core assembly.

## Remaining exit-gate work

- expand and formally approve the v1 corpus;
- add real-Excel build/locale oracle runs and Formula/Formula2 policy evidence;
- isolate the native `FormulaR1C1` shutdown failure before any R1C1 adapter use;
- implement or select the semantic expression AST candidate;
- complete third-party dependency and security comparison;
- add reference-corpus performance distributions and fuzz minimization;
- prove coverage-specific refusal at an actual formula-command mutation boundary.

No formula-mutating command is registered or enabled by this slice.
