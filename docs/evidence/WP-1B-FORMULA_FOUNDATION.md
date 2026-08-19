# Phase 1B formula foundation evidence

Date: 2026-08-19
Scope: ADR-0004 decision gate and pure-core transform boundary

## Outcome

ADR-0004 is accepted only for a bounded A1 subset. This is permission to build
transactional commands on the qualified core, not a claim of broad Excel formula
support. Unsupported constructs refuse before planning or mutation.

## Implemented boundary

- Lossless parser admission with hard length/token/nesting ceilings.
- A1 reference translation using parsed coordinates and exact source spans.
- Formula-pattern transpose with relative coordinate offsets and anchor kinds
  exchanged by axis.
- Caret-scoped A1 endpoint toggle as a pure function.
- Exact top-level IFERROR toggle with dialect-aware separator.
- Canonical precedence-safe sign reversal.
- Explicit divide/multiply operations for thousands and millions.
- Stable structured refusal for R1C1 mutation, names, external/structured
  references, dynamic arrays, unions/intersections, arrays, invalid syntax, and
  out-of-bounds results.

No transformer calls Excel, owns a COM value, performs I/O, starts a thread, or
returns partial output on refusal.

## Verification

- Release suite: **208 passed**, zero failed.
- Generated reference translation cases: **1,000**, each verified by inverse
  displacement back to byte-identical input.
- Concurrent pure-core transforms: **500**, with no shared-state failures.
- Representative bounded throughput check: 500 transforms of a 200-reference
  formula complete inside a conservative five-second test ceiling.
- Existing parser generation: 500 formulas; concurrent parser generation: 250.
- Existing native A1 oracle: eight representative formulas, exact round-trip,
  clean Excel exit.

## Safety decisions

- R1C1 is not enabled for mutation because the native oracle lifecycle did not
  close Excel cleanly.
- The live formula-edit reference toggle is not enabled because standard Excel
  COM does not yet provide qualified caret/edit-text ownership. No global
  keyboard hook, injected code, simulated keystroke, clipboard dependency, or
  focus-stealing workaround is authorized.
- Excel command adapters must still add bounded snapshots, stale-plan checks,
  main-thread writes, immediate postcondition verification, compensation, and
  optimistic undo receipts before workbook mutation ships.

## Residual qualification

- Run formula-command golden workbooks through the installed add-in after each
  transactional adapter lands.
- Add explicitly supported Excel builds/locales to the evidence matrix; a locale
  is not inferred from parser acceptance.
- Any syntax expansion requires a coverage-matrix and ADR amendment first.
