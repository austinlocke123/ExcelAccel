# ADR-0004: Formula parser strategy and coverage

- Status: **Accepted (narrow v1 boundary)**
- Date: 2026-08-19
- Deciders: ExcelAccel engineering

## Context

Smart Copy, transpose, reference transforms, IFERROR handling, Formula
Inspector, Model Check normalization, and later comparison all depend on a
correct formula representation. Silent formula corruption is unacceptable.

The implementation approach is unresolved:

1. custom lexer/parser and AST;
2. an approved existing parser/library;
3. Excel-assisted parsing/translation where stable APIs suffice;
4. a hybrid with explicit ownership and normalization boundaries.

## Decision required

Phase 0 must select the strategy only after testing candidates against a
versioned corpus. Broad incomplete support is worse than narrow explicit
support.

## Minimum coverage dimensions to classify

- A1 and R1C1 notation;
- relative, absolute, and mixed references;
- ranges, unions, intersections, and multi-area references;
- quoted and escaped worksheet/workbook names;
- external workbook references;
- workbook- and worksheet-scoped names;
- structured table references;
- dynamic arrays, spill operator, and implicit intersection;
- legacy CSE array formulas;
- locale list/decimal separators and localized function behavior;
- error literals, strings, dates/numbers, and whitespace preservation;
- functions with reference/value context differences;
- unsupported future syntax and forward-compatible refusal.

Each case is marked `parse`, `round_trip`, `transform`, `inspect_only`, or
`refuse`. A construct may be safely inspectable while not safely transformable.

## Required interface properties

- immutable syntax representation;
- source spans and round-trip fidelity policy;
- explicit reference nodes and qualifier model;
- canonical normalization separate from user-visible formatting;
- deterministic serialization;
- no Excel COM types;
- bounded parse time, nesting, token count, and memory;
- structured failure with no partial transformation.

## Required evidence

- AC-P0-005 and AC-FORM-001/002;
- fuzz/property testing for parse/serialize/transform invariants;
- hostile depth/size inputs;
- supported locale/build matrix;
- dependency license, maintenance, security, and offline review;
- performance on representative formula corpora.

No formula-mutating production command may begin until this ADR is accepted.

## Phase 0 spike finding — 2026-08-18

The first WP-P0-05 slice implemented a dependency-free, pure-core lossless
syntax/reference prototype and a versioned refusal corpus. The current proposed
direction is a hybrid with strict ownership boundaries:

1. a pure-core lossless concrete syntax representation owns source spans,
   deterministic round-trip, resource limits, and explicit reference nodes;
2. a subsequently qualified expression AST owns semantic transforms only for
   syntax explicitly marked `transform`;
3. Excel may act as a build/locale qualification oracle or narrowly scoped
   translation adapter, but never as the domain representation and never from a
   worker thread;
4. names, structured references, external links, dynamic-array operators,
   unions, and intersections remain inspect-only until their individual
   transform rules pass corpus qualification.

At that checkpoint this finding was not yet an accepted decision. The bounded
transform work and decision below close the architectural gate without
expanding the supported syntax set.

## Decision — 2026-08-19

Accept the hybrid strategy for a deliberately narrow v1 mutation boundary:

1. `FormulaParser` owns bounded admission, exact source/tokens/spans, explicit
   reference nodes, coverage classification, and deterministic refusal.
2. `A1FormulaTransformer` owns only source-span-directed A1 reference edits.
   It supports ordinary A1 references/ranges, mixed anchors, and local worksheet
   qualifiers. Copy translation and transpose are coordinate operations over
   parsed reference nodes; no command may search-and-replace formula text.
3. `FormulaWrapperTransformer` owns only whole-expression IFERROR, sign, and
   unit wrappers. Parenthesized expression ownership preserves precedence and
   parsed tokens identify canonical top-level wrappers.
4. Production mutation remains A1-only. R1C1 remains available to the pure
   parser for inspection/corpus work, but no R1C1 mutation adapter may ship
   until its real-Excel lifecycle problem is resolved.
5. Names, structured references, external references, dynamic-array operators,
   unions, intersections, array constants/CSE arrays, and unknown syntax remain
   inspect-only or refused before a mutation plan exists.
6. Excel COM owns only bounded capture, revalidation, write, verification, and
   rollback on Excel's main thread. It does not own parsing or transformation.
7. Formula-edit caret automation is not part of this acceptance. The reference
   toggle command remains disabled unless a separate host spike proves an exact,
   supported caret/edit-text API without hooks, injection, focus theft, or
   retained COM objects.

## Candidate comparison

| Candidate | Decision | Reason |
|---|---|---|
| Custom bounded concrete syntax plus narrow transforms | Selected for v1 | Already passes the versioned corpus, has explicit resource ceilings and refusal codes, introduces no deployment dependency, and keeps mutation scope auditable. |
| General third-party formula parser | Deferred | A broader grammar would still require Excel-specific ownership rules, coverage classification, lifecycle qualification, and dependency/security review. It does not reduce immediate risk enough for the narrow v1 set. |
| Excel/COM as parser or transformation engine | Oracle/adapter only | COM is stateful and host-sensitive, does not provide the required immutable source model, and the R1C1 oracle experiment did not exit cleanly. |
| Raw string transformation | Rejected | It cannot safely distinguish references, strings, names, qualifiers, precedence, or future syntax. |

This decision can be revisited when a broader feature needs syntax outside the
accepted matrix. Expansion requires new corpus cases and real-Excel evidence;
it is never inferred from a successful parse.

## Acceptance evidence

- The versioned parser corpus covers A1/R1C1, mixed anchors, ranges, quoted
  sheets, external/structured/name/dynamic-array/locale cases, and stable
  inspect/refusal outcomes.
- Hostile size, token, nesting, delimiter, character, deterministic round-trip,
  and concurrent pure-core tests pass.
- A generated 1,000-formula translation corpus passes the forward/inverse
  displacement invariant; a 500-operation concurrent transform corpus passes.
- Golden tests cover exact changed spans, mixed-anchor translation, transpose
  axis/anchor exchange, caret endpoint selection, IFERROR identity, canonical
  sign reversal, unit operations, semicolon/comma dialects, bounds refusal, and
  every inspect-only mutation refusal.
- Eight representative A1 cases previously passed the native Excel formula
  oracle with clean process exit. R1C1 did not and therefore remains closed.
- The complete Release unit suite passes with warnings treated as errors. See
  `docs/evidence/WP-1B-FORMULA_FOUNDATION.md`.
