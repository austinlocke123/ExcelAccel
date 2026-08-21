# Original feature coverage map

Status: **Draft for review**

This map shows where every original functional-requirements family is now
specified and whether its broadest behavior is active, gated, or deferred. It
prevents a feature from disappearing during document restructuring.

| Original area | Detailed contract | Current disposition |
|---|---|---|
| §6.1 formatting cycles | `commands/CATALOG.md` §1.1/1.2 | Phase 1A |
| Number-format cycles, user-defined | `commands/FORMAT_CYCLES.md` | Schema and cycle semantics implemented 2026-08-20 (WP-F-01); user management pending WP-F-02 |
| Settings editor for cycles | `commands/FORMAT_CYCLES.md` | Approved, not implemented (WP-F-03) |
| Blue-black input/formula toggle | `commands/AUTOCOLOR.md` | Approved, not implemented |
| Cell classification rules | `commands/AUTOCOLOR.md` | Implemented 2026-08-20 (WP-F-08); commands unregistered pending an adapter |
| Ribbon task grouping | `commands/RIBBON_LAYOUT.md` | Implemented 2026-08-20 |
| §6.1 AutoColor selection/worksheet | `commands/CATALOG.md` §1.3 | Phase 1A |
| §6.1 workbook AutoColor | `commands/WORKBOOK_OPERATIONS.md` §4 | Gated high-impact |
| §6.1 named styles and capture/apply | `commands/DISCOVERY_STYLES_AND_PROFILES.md` | Phase 1B |
| §6.1 sheet style/workbook AutoFormat | `commands/WORKBOOK_OPERATIONS.md` §4 | Gated high-impact |
| §6.1 gridlines/freeze/zoom/AutoFit | `commands/CATALOG.md` §1.2 | Phase 1A |
| §6.1 unspecified print/view settings | `commands/WORKBOOK_OPERATIONS.md` §5 | Unspecified/deferred pending property catalog |
| §6.2 formula reference toggle | `commands/FORMULA_TRANSFORMS.md` | Phase 1B after edit-mode/parser proof |
| §6.2 Smart Copy/spacing/transpose | `commands/FORMULA_TRANSFORMS.md` | Phase 1B after ADR-0004 |
| §6.2 copy formula/value from above | `commands/FORMULA_TRANSFORMS.md` | Phase 1B after receipt proof |
| §6.2 formulas/values/formats paste and fills | `commands/FORMULA_TRANSFORMS.md` | Phase 1B |
| §6.2 IFERROR/sign/unit transforms | `commands/FORMULA_TRANSFORMS.md` | Phase 1B |
| §6.3 precedents/dependents/trace history | `commands/AUDITING.md` | Phase 2 |
| §6.3 Formula Inspector | `commands/AUDITING.md` | Phase 2 |
| §6.3 selection tools | `commands/WORKBOOK_OPERATIONS.md` §1 | Phase 1B/2 by parser dependency |
| §6.4 Model Check | `commands/MODEL_CHECK.md` | Phase 2 |
| §6.5 workbook/worksheet compare | `commands/COMPARE.md` | Gated; same-shape/open-workbook/read-only only |
| §6.5 structural alignment and shift inference | `REQUIREMENTS.md` §4.6 | Deferred |
| §6.6 Named Range Manager | `commands/NAMES_AND_LINKS.md` | Gated read-only inventory/navigation |
| §6.6 name rename/delete | `REQUIREMENTS.md` §4.6 and `commands/NAMES_AND_LINKS.md` | Deferred |
| §6.7 External Link Manager | `commands/NAMES_AND_LINKS.md` | Gated read-only inventory/navigation |
| §6.7 link repoint/break | `REQUIREMENTS.md` §4.6 and `commands/NAMES_AND_LINKS.md` | Deferred |
| §6.8 one/two-way sensitivity | `commands/SENSITIVITIES_AND_FINANCE.md` §1 | Individually gated native Data Tables |
| §6.9 LBO/PE helpers | `commands/SENSITIVITIES_AND_FINANCE.md` §3 | Declarative template engine; content individually reviewed |
| §6.9 circularity tools | `commands/SENSITIVITIES_AND_FINANCE.md` §2 | Inspect/switch gated; settings separately gated |
| §6.10 deterministic cleaning | `commands/DATA_CLEANING.md` | Phase 1B |
| §6.11 navigation/bookmarks | `commands/CATALOG.md` §1.4 | Phase 1A, session-only bookmarks |
| §6.11 row/column hide/group/Smart Hide | `commands/WORKBOOK_OPERATIONS.md` §2 | Individually gated |
| §6.11 row/column insert/delete | `commands/WORKBOOK_OPERATIONS.md` §3 | Explicitly deferred pending structural transaction ADR |
| §6.12 native chart formatting | `commands/CHARTS_AND_POWERPOINT.md` §1 | Individually gated selected-chart scope |
| §6.13 Snapshot to PowerPoint | `commands/CHARTS_AND_POWERPOINT.md` §2 | Individually gated image-only adapter |
| §4 command search/favorites | `commands/DISCOVERY_STYLES_AND_PROFILES.md` | Phase 1B |
| §4 custom command chains | `REQUIREMENTS.md` §4.6 | Deferred |
| §8 profiles/import/export | `commands/DISCOVERY_STYLES_AND_PROFILES.md` | Phase 1A local profile; Phase 1B exchange |

## Deliberate scope changes from the original draft

- Comparison is initially same-shape and read-only; there is no inferred
  structural alignment.
- Names and links are initially read-only.
- Finance helpers are declarative reviewed templates, not hard-coded logic.
- Persistent workbook bookmarks and persistent undo are removed from active
  scope.
- PowerPoint output is image-only; embedded objects and live links are deferred.
- Structural row/column deletion and workbook-wide formatting are documented but
  require stronger transaction evidence before release.
- Generic print/view settings are not considered a feature contract.
