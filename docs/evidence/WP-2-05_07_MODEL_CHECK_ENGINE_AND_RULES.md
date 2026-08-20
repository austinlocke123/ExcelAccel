# WP-2-05 to WP-2-07 Model Check engine, finding schema, and rules

Date: 2026-08-19

Status: **Engine, finding schema, and all seven rules implemented in pure Core.
Result commands, Excel wiring, and registration are WP-2-08 and remain in
progress.**

## Contract

- Capability: CAP-CHECK-001, CAP-CHECK-002
- Governing decision: ADR-0004
- Acceptance: AC-CHECK-001 through AC-CHECK-028
- Allowed implementation: pure Core scan model, engine, rule set, deterministic
  tests, and engineering evidence
- Excluded: result presentation, navigation, ignores, rescan, export, Excel
  wiring, and command registration until WP-2-08

## Product boundary

Model Check is a deterministic read-only rule engine. A finding states a rule, a
location, evidence, and an inspection prompt. It never declares that a model is
wrong, assigns a health or quality score, modifies a workbook, or creates review
workflow. A test asserts that no rule statement uses verdict or score vocabulary.

## Engine

`ModelCheckEngine` runs the enabled rules against one immutable
`ModelCheckSnapshot`.

- **AC-CHECK-001** — exactly the enabled rule IDs and versions run, against one
  snapshot captured before any rule sees it.
- **AC-CHECK-002** — every finding carries rule ID and version, severity, target,
  evidence, coverage category, fingerprint, and navigation state.
- **AC-CHECK-003** — no finding declares correctness or carries a score.
- **AC-CHECK-004** — rules run in stable ID order and findings are sorted by a
  canonical key, so reordering execution cannot change the output. Identical
  inputs produce identical findings and fingerprints.
- **AC-CHECK-005** — a rule that throws is recorded as a failure with its rule ID
  and a safe diagnostic (the exception type name, never workbook content), and
  the scan becomes partial. A rule is never silently dropped.
- **AC-CHECK-006** — a cancelled scan is refused, carrying no findings, so a
  partial run can never be presented as a completed scan.

Findings are capped at 5,000 with explicit truncation. Any rule failure, partial
coverage, or truncation blocks the completeness claim.

Fingerprints are SHA-256 over normalized rule and location inputs, reusing the
existing `PreconditionFingerprint`. A test asserts the fingerprint is a 64-character
hex digest containing no raw formula or value content, which is what makes
`AC-CHECK-030` satisfiable when ignores land.

## Rules

| Rule | Behaviour | Acceptance |
|---|---|---|
| `check.formula.pattern_inconsistency` | Normalizes each formula to its relative shape and reports cells differing from the peer baseline. | AC-CHECK-009..012 |
| `check.constant.interrupts_formula_region` | Reports a constant inside an otherwise formula-consistent region, separately from an embedded literal. | AC-CHECK-013, 014 |
| `check.formula.embedded_numeric_constant` | Reports numeric literals with exact source spans, minus versioned allowlists and structural exclusions. | AC-CHECK-015..018 |
| `check.formula.error` | Reports captured error values and broken `#REF!` references without recalculating. | AC-CHECK-019, 020 |
| `check.reference.external` | Reports external references without opening, contacting, or refreshing the source. | AC-CHECK-021, 022 |
| `check.reference.circular` | Bounded graph cycle detection over the snapshot, reporting each cycle once. | AC-CHECK-023..025 |
| `check.format.number_inconsistency` | Reports number formats differing from the peer baseline, never changing formatting. | AC-CHECK-026..028 |

### Peer regions and shape

`PeerRegion` groups cells into contiguous column runs. Grouping is purely
spatial, so it does not depend on rule order and never infers meaning from
labels. A blank breaks a run under the approved default, and a row gap always
breaks one.

`FormulaShape` normalizes a formula into the structure two copied cells share:
each reference becomes an offset from its own cell, so `=A1*2` in B1 and `=A2*2`
in B2 share a shape while `=A1*3` does not. Absolute and relative coordinates
normalize differently, so they are correctly reported as distinct shapes.

**A formula the qualified parser cannot cover has no shape at all**, rather than
a guessed one. A peer region containing one is reported as partial coverage and
produces no findings, which is what prevents a false claim of consistency
(AC-CHECK-011).

### Literal detection without a syntax tree

The contract describes traversing numeric literal AST nodes. The qualified parser
exposes a token stream, not a tree, so literals are identified from `Number`
tokens with their exact source spans, and the enclosing function is tracked with
a parenthesis stack. This gives the exact literal and span AC-CHECK-015 requires.

The limitation is real and worth stating: the enclosing-function context is one
level deep and positional argument indices are not modelled, so a structural
exclusion is per-function rather than per-argument-position. Narrowing that
requires the parse tree that WP-2-04 needs.

Numeric text and cell constants are not misclassified as embedded literals
(AC-CHECK-018), because only `Number` tokens inside a formula are considered.

## Verification

- Release build: **zero warnings, zero errors**.
- Release tests: **481 passed**, zero failed.
- Engine coverage: enabled-rule selection, the full finding schema, fingerprints
  free of raw content, reordering invariance, identical-input determinism, rule
  failure naming the rule, cancellation, ignore suppression by exact fingerprint,
  partial coverage blocking completeness, a clean scan claiming completeness, and
  null and duplicate-cell rejection.
- Rule coverage: copied-formula baselines and exceptions, absolute versus
  relative normalization, consistent regions, sub-minimum regions, parser gaps
  preventing false consistency, blanks breaking regions, constants reported
  separately from embedded literals, literal spans, allowlists, structural
  exclusions, allowlist rescan deltas, numeric text exclusion, `#REF!` and
  captured errors, external references, direct, self, and longer cycles, acyclic
  chains, format baselines and exceptions, case and whitespace normalization,
  mixed formula and value format regions, catalog identity, and the no-verdict
  vocabulary assertion.

## Retained limitations

- Peer regions are column runs only. Row-oriented regions are not grouped, so a
  model laid out across a row is not judged for consistency.
- Structural literal exclusions are per-function, not per-argument-position, as
  described above.
- Circular detection sees only the snapshot's captured cells; a cycle leaving the
  scanned scope is not followed, and a name-bound edge marks coverage partial.
- The engine has no performance corpus yet. AC-CHECK-007 and AC-CHECK-008 budgets
  are not measured; the scan is bounded by the snapshot and finding caps.

## Next

WP-2-08: finding navigation, local ignores, rescan, and export, plus the Excel
snapshot capture, scan scoping with preview, presentation, and registration.
