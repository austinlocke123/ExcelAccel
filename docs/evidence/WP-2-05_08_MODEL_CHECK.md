# WP-2-05 to WP-2-08 Model Check engine, rules, and results

Date: 2026-08-19

Status: **Complete, together with WP-2-08. The engine, finding schema, seven
rules, result commands, Excel wiring, and registration are implemented and
verified. Workbook scope remains gated.**

## Contract

- Capability: CAP-CHECK-001, CAP-CHECK-002
- Governing decision: ADR-0004
- Acceptance: AC-CHECK-001 through AC-CHECK-028
- Allowed implementation: pure Core scan model, engine, rule set, deterministic
  tests, and engineering evidence
- Acceptance also covers AC-CHECK-029 through AC-CHECK-037 via WP-2-08
- Excluded: workbook-scope scanning, which remains gated

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
- Release tests: **502 passed**, zero failed.
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

## WP-2-08 results, wiring, and registration

### Scan scoping

`ModelCheckCoordinator` scans a selection or a worksheet. **Workbook scope is
refused** with `CHECK_SCOPE_TOO_LARGE`, inheriting the same unresolved
workbook-scale performance gate as WP-2-02b.

A worksheet scan reuses the bounded `DependentScanRegion` plan, so the reported
used range is untrusted here too: an over-large or over-wide region is refused
before anything is read, and the rest is banded so no single read exceeds the
block ceiling. A planned scan above 25,000 cells must be confirmed first; without
confirmation nothing is read.

Cancellation is checked before every band and inside the engine. A cancelled scan
is refused and **the prior result is not replaced**, so previous findings stay on
screen rather than being overwritten by a partial run (AC-CHECK-006).

### Navigation, ignores, rescan, export

- **Navigation (AC-CHECK-029)** reuses the shared trace view's Go To, which
  revalidates the target through the navigation port and records the prior
  location in session history. Every finding row is navigable because findings
  always point at a real scanned cell.
- **Local ignores (AC-CHECK-030..033)** store rule ID, rule version, and the
  normalized fingerprint only. A test asserts the written file contains no
  formula prefix. An ignore suppresses only an exactly matching fingerprint, so
  it can never hide a different rule or location. The confirm dialog states
  exactly what is stored and that a rescan is required. Active ignores are listed
  and removable.
- **Rescan (AC-CHECK-034)** repeats the exact prior scope and rule set against a
  newly captured snapshot, refreshing only the ignore configuration. Prior
  findings are never reused as current evidence.
- **Export (AC-CHECK-035..037)** requires an explicit destination and a confirmed
  manifest naming every included and excluded field. Evidence is **opt-in and off
  by default**, so formulas and values are excluded by default. The write is
  temp, validate, then replace, and a failure leaves existing destination data
  intact. Nothing is transmitted.

### Ignore storage deviation

The contract says an ignore writes the local profile. Ignores are instead held in
their own atomic file beside the profile, `model-check-ignores.tsv`.

Adding them to `ProfileDefinition` would bump the shipped profile schema and
ripple through profile export and import, which is a larger change than this
package should make. The properties the acceptance criteria actually require -
atomic local write, no raw content, exact-fingerprint matching, visibility,
removal, and portability only through a deliberate action - are all provided.
The `MODEL_CHECK.md` contract itself anticipates a local profile *or a separately
exported ignore set*. Folding it into the profile schema remains open.

### Registration

Six commands are registered through the central dispatcher, Command Search, and a
new Ribbon Model Check menu on KeyTip route `Alt, X, A, K`:
`model_check.run.selection` (MS), `model_check.run.worksheet` (MW),
`model_check.rescan` (MR), `model_check.finding.ignore_local` (MI),
`model_check.finding.unignore_local` (MU), and `model_check.results.export` (ME).
The three scan commands are read-only; ignore, unignore, and export declare their
local writes and a mandatory preview.

### Real-Excel verification

Against a live worksheet with a seeded inconsistent formula, the registered
selection route returned `open|success|2|0`: the view opened, two findings were
reported (the pattern exception and its embedded literal), and **no rule failed**.
Rescan repeated the scope and returned the same two findings. Workbook contents
were unchanged, the view released on explicit close, and Excel exited naturally
with no surviving process.

## Next

- WP-2-09: Phase 2 large-corpus, cancellation, privacy, performance, and soak
  qualification. This is where the missing performance corpus belongs.
- WP-2-04 (Formula Inspector) still needs the parse tree.
- Workbook scope for both dependents and Model Check stays blocked on the
  workbook-scale performance gate.
