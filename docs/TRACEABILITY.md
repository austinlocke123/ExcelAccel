# Specification traceability

Status: **Draft for review**

This matrix is the routing index. Detailed authority remains in the linked
documents.

## 1. Foundation and cross-cutting requirements

| Requirement/capability | Architecture/decision | Work package | Acceptance |
|---|---|---|---|
| INV-001, SEC-001 | Architecture §12; ADR-0002 | WP-P0-01/08 | AC-SEC-001, AC-P0-002/008 |
| INV-002, CMD-004 | Architecture §4/8; command canonical planning | WP-P0-03 | AC-ARCH-005, AC-CMD-008 |
| INV-003/005/009 | Architecture §4/7/8; ADR-0003 | WP-P0-03/04 | AC-REL-003/005, AC-CMD-005/006 |
| INV-006..008 | Architecture §3/5; Reliability §2/3 | WP-P0-02/04 | AC-ARCH-001..004, AC-REL-001/002 |
| REL-001..007 | Reliability | WP-P0-02/04 | AC-REL-001..012 |
| PERF-001..010 | Architecture §5/6; Reliability §8/9 | WP-P0-07 | AC-PERF-001..006 |
| PLAT-001..006 | ADR-0001/0002/0004/0005 | WP-P0-01/05/06/08 | AC-P0-001/002/005/007/008 |

## 2. Active command capabilities

| Capability | Command source | Work package | Acceptance |
|---|---|---|---|
| CAP-CMD-001 | commands/README | WP-P0-03, WP-1A-02 | AC-CMD-001..008 |
| CAP-KEY-001 | commands/CATALOG §1 | WP-1A-05 | AC-KEY-001..004 |
| CAP-PROF-001 | Requirements §8; commands/CATALOG §5 | WP-1A-04 | AC-PROF-001..004 |
| CAP-FMT-001 | commands/CATALOG §1.1/1.2 | WP-1A-06 | AC-FMT-001..009, AC-REL-005 |
| CAP-FMT-002 | commands/CATALOG §1.3 | WP-1A-07 | AC-FMT-010..013 |
| CAP-NAV-001 | commands/CATALOG §1.4 | WP-1A-08 | AC-NAV-001..006 |
| CAP-UNDO-001 | ADR-0003; Reliability §6 | WP-1A-09 | AC-REL-011/012 |
| CAP-UX-001 | Requirements §9; Reliability §8 | WP-1A-11 | AC-UX-001..005 |

## 3. Phase 1B capabilities

| Capability | Command source | Work package | Acceptance |
|---|---|---|---|
| CAP-SEARCH-001, CAP-FAV-001 | commands/DISCOVERY_STYLES_AND_PROFILES | WP-1B-01 | AC-SEARCH-001..004, AC-FAV-001..004 |
| CAP-STYLE-001 | commands/DISCOVERY_STYLES_AND_PROFILES | WP-1B-02 | AC-STYLE-001..008 |
| CAP-PROF-002 | commands/DISCOVERY_STYLES_AND_PROFILES | WP-1B-03 | AC-PROF-005..009 |
| CAP-FORM-001 | commands/FORMULA_TRANSFORMS; ADR-0004 | WP-1B-04..08 | AC-FORM-001..038 |
| CAP-DATA-001 | commands/DATA_CLEANING | WP-1B-09/10 | AC-DATA-001..019 |
| CAP-SELECT-001 | commands/WORKBOOK_OPERATIONS §1 | WP-1B-11 | AC-SELECT-001..007 |

## 4. Phase 2 capabilities

| Capability | Command source | Work package | Acceptance |
|---|---|---|---|
| CAP-AUD-001, CAP-AUD-002 | commands/AUDITING | WP-2-01..04 | AC-AUD-001..021 |
| CAP-CHECK-001, CAP-CHECK-002 | commands/MODEL_CHECK | WP-2-05..08 | AC-CHECK-001..037 |

## 5. Individually gated capabilities

| Capability | Command source | Work package | Acceptance |
|---|---|---|---|
| CAP-NAME-001 | commands/NAMES_AND_LINKS | WP-G-01 | AC-NAME-001..011 |
| CAP-LINK-001 | commands/NAMES_AND_LINKS | WP-G-02 | AC-LINK-001..011 |
| CAP-CMP-001 | commands/COMPARE | WP-G-03 | AC-CMP-001..019 |
| CAP-SENS-001 | commands/SENSITIVITIES_AND_FINANCE | WP-G-04 | AC-SENS-001..014 |
| CAP-CIRC-001 | commands/SENSITIVITIES_AND_FINANCE | WP-G-05/06 | AC-CIRC-001..016 |
| CAP-TPL-001 | commands/SENSITIVITIES_AND_FINANCE | WP-G-07/08 | AC-TPL-001..014 |
| CAP-CHART-001 | commands/CHARTS_AND_POWERPOINT | WP-G-09 | AC-CHART-001..027 |
| CAP-PPT-001 | commands/CHARTS_AND_POWERPOINT | WP-G-10 | AC-PPT-001..009 |
| CAP-STRUCT-001 | commands/WORKBOOK_OPERATIONS §2/3 | WP-G-11/12 | AC-STRUCT-001..019 |
| CAP-FMT-003 | commands/WORKBOOK_OPERATIONS §4 | WP-G-13 | AC-FMT-014..020 |

## 6. Known review gaps

- Default formatting profile values and proposed shortcuts are intentionally
  unspecified.
- Runtime and minimum Excel build are unresolved.
- The parser choice remains unresolved; the proposed formula coverage matrix
  and hybrid recommendation are WP-P0-05 evidence, not an accepted ADR.
- AutoSave/coauthoring impact-tier behavior has a tested proposed matrix, but
  cloud/remote-event build evidence and ADR-0005 acceptance remain open.
- WP-P0-07 now has an initial synthetic corpus, distribution math, and isolated
  real-Excel Quick harness. Reference-machine acceptance, qualification
  distributions, UI-heartbeat evidence, and frozen budgets remain open.
- Installer/update technology and signing process are open.

These gaps are deliberate stop signs, not invitations for implementation agents
to choose defaults.
