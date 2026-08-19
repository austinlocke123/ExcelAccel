# ADR-0004: Formula parser strategy and coverage

- Status: **Open**
- Date: 2026-08-18
- Deciders: open

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

This finding is not an accepted decision. The prototype deliberately stops
short of a complete expression AST, real-Excel locale/build qualification,
third-party dependency comparison, and production transformation. Those gaps
must close before the deciders change this ADR to `Accepted`.
