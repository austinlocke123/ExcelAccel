# Command specification

Status: **Draft for review**

This folder owns user-visible command behavior. A command is the only supported
path for a user action that reads or mutates an Excel or PowerPoint document.
Ribbon buttons, Quick Keys, search, favorites, and panes are invocation surfaces,
not separate implementations.

## 1. Required command record

Every command specification MUST define:

| Field | Meaning |
|---|---|
| `id` | Stable lowercase dotted ID; never localized or reused. |
| `version` | Contract version for serialized references and migrations. |
| `capability` | Owning requirement ID. |
| `phase` | Earliest approved phase. |
| `category` | Stable grouping key. |
| `display_name` | Localizable name. |
| `aliases` | Localizable search terms; not identity. |
| `parameter_model` | Fixed parameters or typed parameter schema. |
| `supported_context` | Workbook, sheet, selection, object, edit-mode, protection, and capability requirements. |
| `impact` | `read_only`, `low`, `medium`, or `high`. |
| `changed_properties` | Exhaustive property/value categories the command may change. Empty for read-only commands. |
| `can_execute` | Side-effect-free validation and refusal codes. |
| `snapshot` | Exact properties read and maximum supported scope. |
| `plan` | Deterministic target and change calculation. |
| `preview` | `none`, `threshold`, or `mandatory`; content shown. |
| `execute` | Exact mutation semantics and postconditions. |
| `undo` | `none`, `session_receipt`, or later approved policy. |
| `failure_behavior` | Refusal, rollback, skip, and partial-failure policy. |
| `performance_class` | Applicable `PERF-*` target. |
| `acceptance` | Required `AC-*` identifiers. |

No implementation may infer omitted mutation behavior from a display name.

## 2. Impact policy

| Impact | Examples | Default execution policy |
|---|---|---|
| `read_only` | navigate, inspect, inventory | No mutation; may run from immutable snapshot. |
| `low` | one-property formatting, navigation state | Immediate after validation; session receipt when eligible. |
| `medium` | broad selection transform, worksheet AutoColor | Preview above qualified thresholds; precondition recheck and receipt required. |
| `high` | workbook-wide or formula-to-value mutation | Mandatory preview and confirmation; deferred unless rollback and AutoSave/coauthoring semantics are qualified. |

Impact is based on consequence, not implementation effort.

## 3. Common refusal codes

- `NO_WORKBOOK`
- `EDIT_MODE_UNSUPPORTED`
- `SELECTION_UNSUPPORTED`
- `MULTI_AREA_UNSUPPORTED`
- `PROTECTED_TARGET`
- `READ_ONLY_WORKBOOK`
- `AUTOSAVE_STATE_UNSAFE`
- `COAUTHORING_STATE_UNSAFE`
- `ARRAY_OR_SPILL_UNSAFE`
- `TABLE_CONTEXT_UNSAFE`
- `STALE_CONTEXT`
- `RESOURCE_LIMIT`
- `EXCEL_CAPABILITY_MISSING`
- `COMMAND_QUARANTINED`

Each refusal includes a safe user message and remediation. Refusal is not an
exception and does not enter the mutation state guard.

## 4. Canonical planning

A canonical plan orders targets and property deltas deterministically, excludes
timestamps and random IDs from its canonical hash, records the command contract
version, and normalizes locale-independent serialized values.

Byte-for-byte equality is required only for canonical serialization. UI preview
text may be localized and is not part of the plan hash.

## 5. Command review rules

A command change requires architecture/reliability review when it:

- adds a changed property or target type;
- increases scope or impact tier;
- adds formula-to-value behavior;
- adds a COM API surface;
- changes preview, transaction, rollback, or undo behavior;
- changes behavior under protection, AutoSave, or coauthoring;
- adds persistence, background work, network access, or another Office app.

Aliases, examples, help text, and proposed default bindings are lightweight if
behavior is unchanged.

## 6. Contract index

- [`CATALOG.md`](CATALOG.md) — phased index and Phase 1A contracts
- [`DISCOVERY_STYLES_AND_PROFILES.md`](DISCOVERY_STYLES_AND_PROFILES.md) —
  search, favorites, styles, and profile exchange
- [`FORMULA_TRANSFORMS.md`](FORMULA_TRANSFORMS.md) — reference editing, copy,
  transpose, wrappers, sign/scale, paste, and fill
- [`DATA_CLEANING.md`](DATA_CLEANING.md) — explicit text, type, date, and display
  conversions
- [`WORKBOOK_OPERATIONS.md`](WORKBOOK_OPERATIONS.md) — deterministic selection,
  row/column operations, and broad formatting
- [`AUDITING.md`](AUDITING.md) — precedents, dependents, navigation, and Formula
  Inspector
- [`MODEL_CHECK.md`](MODEL_CHECK.md) — scans, rule contracts, findings, ignores,
  and export
- [`NAMES_AND_LINKS.md`](NAMES_AND_LINKS.md) — read-only inventory, usage, search,
  navigation, and export
- [`COMPARE.md`](COMPARE.md) — same-shape range, worksheet, and workbook
  comparison
- [`SENSITIVITIES_AND_FINANCE.md`](SENSITIVITIES_AND_FINANCE.md) — native Data
  Tables, circularity tools, and declarative finance templates
- [`CHARTS_AND_POWERPOINT.md`](CHARTS_AND_POWERPOINT.md) — selected-chart
  formatting and explicit image snapshots
- [`../FEATURE_COVERAGE.md`](../FEATURE_COVERAGE.md) — original feature-to-
  contract coverage and deliberate scope changes

## 7. Specification template

```markdown
### `category.command_name`

- Version:
- Capability:
- Phase:
- Impact:
- Parameters:
- Supported context:
- Changed properties:
- CanExecute/refusals:
- Snapshot:
- Plan:
- Preview:
- Execute/postconditions:
- Undo:
- Failure behavior:
- Performance:
- Acceptance:
```
