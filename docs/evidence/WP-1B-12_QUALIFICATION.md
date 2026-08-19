# WP-1B-12 Phase 1B qualification checkpoint

Date: 2026-08-19

Status: **Engineering checkpoint passed; not end-user distribution approval**

## Qualified implementation scope

- WP-1B-01: local Command Search, deterministic ranking, favorites, and
  availability-aware invocation.
- WP-1B-02: versioned local/built-in style recipes, capture/apply/delete, exact
  preview, property-batch rollback, and optimistic undo.
- WP-1B-03: offline profile package export/import preview/apply plus binding
  cheat-sheet export with bounded validation and atomic replacement.
- WP-1B-04..08: narrow A1 formula parser/transforms; Smart Copy; row/column
  spacing; transpose; IFERROR, sign, and unit wrappers; formula/value/formats
  paste; formula/value from above; and typed numeric/date sequences.
- WP-1B-09/10: Unicode text normalization, complete-string typed number/text/date
  conversion, and all explicit blank/zero/N/A/NM/dash directions.
- WP-1B-11: bounded selection of formulas, constants, true blanks, numeric
  hardcodes, and parsed external formulas.

Every released mutation uses an immutable bounded plan, stale-state
revalidation, explicit changed-property scope, verified write, compensating
rollback, and bounded optimistic session receipt. Receipt-store failure itself
is fault-tested and rolls the completed mutation back before returning failure.

## Build and deterministic tests

- Release tests: **288 passed**, zero failed.
- Debug build: zero warnings and zero errors.
- Release build: zero warnings and zero errors.
- Coverage includes parser/golden transforms, full-string numeric grammars,
  explicit date patterns and Excel date systems, Unicode/locale fixtures,
  source/destination shape and overlap rules, calculated-value fingerprints,
  exact format matrices, stale plans, injected partial writes, postcondition
  mismatch, receipt failure, rollback, optimistic undo, command metadata,
  profile migration/import, package validation, and architecture boundaries.

## Real Excel functional/fault smoke

The packed Debug XLL completed the expanded hidden-Excel harness on Excel 16.0
x64. The harness verifies exact formula/value/format outputs and undo across all
Phase 1B command families, selection-only behavior, content/excluded-property
preservation, protected/multi-area/merged refusal, stale-plan refusal, state
restoration after an injected exception, UI open/close cleanup, workbook close,
and natural Excel process exit.

The measured bounded Phase 1B feature suite completed in **1,286 ms** against a
5,000 ms checkpoint ceiling. Its output contained no workbook-derived
diagnostics, no network access, and no surviving Excel process.

## Ten-session reliability soak

Ten fresh hidden Excel processes each ran the complete smoke suite. All 10:

- registered the packed XLL;
- passed every exactness, refusal, fault, and undo assertion;
- closed the workbook and returned from `Application.Quit`;
- exited naturally with no process left behind; and
- released the packed XLL for an exclusive file-open check.

| Metric | Result |
|---|---:|
| Phase 1B bounded feature-suite P95 | 1,532 ms |
| Complete session P95 | 7,982.7452 ms |
| Working-set P95 | 322,646,016 bytes |
| Private-memory P95 | 248,094,720 bytes |
| Handle-count P95 | 1,864 |
| Handle-count range across fresh processes | 20 |

This proves cross-session cleanup and stability. It is not a long-duration
single-process retained-memory test.

## Frozen performance/heartbeat regression run

The current packed build passed the versioned `Quick` corpus in isolated Excel
processes:

| Workload | Current P95 | Frozen budget | Worst heartbeat | Timeouts |
|---|---:|---:|---:|---:|
| 1,000-cell block read | 1.4877 ms | 100 ms | 0 ms | 0 |
| 10,000-cell property write | 11.1308 ms | 500 ms | 2 ms | 0 |
| 100,000-cell block read | 111.9316 ms | 3,000 ms | 39 ms | 0 |
| 250,000-cell / 20-sheet read | 116.7785 ms | 8,000 ms | 7 ms | 0 |

Cold total-startup proxy was 1,243.7181 ms and warm proxy P95 was 972.1703 ms.
Those include Excel process launch and are not compared with the 750 ms
add-in-owned startup target; matched Excel-only attribution remains open.

## Deliberate retained gates

- Live formula-edit reference toggle is not registered: Excel-DNA/COM has not
  provided a proven exact caret/edit-text boundary that avoids hooks, injected
  keystrokes, or keyboard-state risk. The pure cycle is tested.
- Selection of calculated error values is not registered because the current
  selection model intentionally refuses lossy error/display coercion.
- Formats-only paste is deliberately limited to 100 cells and nine properties;
  colors, fills, borders, dimensions, validation, comments, hyperlinks, values,
  and formulas are excluded.
- Unknown collaboration state and medium/high-impact collaborative mutation
  remain fail-closed.
- AutoColor enablement, long-duration in-process retention, supported Office
  build/coexistence/accessibility matrix, CA-signed installer, clean-VM
  lifecycle, enterprise trust/allowlisting, and attributed startup cost remain
  release/distribution gates.

These gates are intentional reliability decisions. This checkpoint authorizes
review of the Phase 1B engineering stack; it does not authorize public release
or silently broaden any gated command.
