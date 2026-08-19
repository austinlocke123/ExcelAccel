# WP-P0-06 AutoSave and coauthoring evidence

- Date: 2026-08-18
- Scope: first WP-P0-06 implementation slice
- Status: prototype passing; ADR-0005 accepted on 2026-08-19 with unknown-state
  refusal; broader AC-P0-007 cloud/build qualification remains open
- Policy: [`POLICY_MATRIX.md`](../collaboration/POLICY_MATRIX.md)

## Implemented behavior

The pure core now contains:

- explicit `On`, `OffOrDisabled`, and `Unknown` AutoSave states;
- distinct modern-risk, legacy-shared, remote-observed, remote-in-progress,
  not-detected, and unknown coauthoring states;
- a checked thread-safe workbook revision tracker;
- immutable execution stamps with workbook identity and opaque property
  fingerprint;
- session-only plan leases with caller-supplied tier lifetimes and a hard
  five-minute ceiling;
- explicit impact-tier decisions and stable refusal codes;
- deterministic locale-invariant SHA-256 fingerprints with component and size
  limits;
- stale-property refusal in the existing currency-format vertical slice.

The Excel adapter captures AutoSave, legacy-sharing, and path signals through
read-only property access. It never sets `AutoSaveOn`, calls `ExclusiveAccess`,
saves, or infers modern exclusivity from `MultiUserEditing`.

## Automated evidence

Pure tests cover:

- every impact tier across local, AutoSave-on, modern-risk, legacy, unknown, and
  remote-in-progress states;
- receipt and remote-event requirements for medium-impact plans;
- high-impact collaborative refusal;
- local and remote intervening-event invalidation;
- changed fingerprints when no event was observed;
- plan expiration, backward clocks, workbook identity change, close, and
  concurrent invalidation;
- classifier ambiguity, fingerprint locale invariance, hard resource limits,
  and separation of lease time from canonical plans.

## Real Excel signal probe

The read-only signal harness passed on Excel 16.0 build 20228:

- `AutoSaveOn` was readable as false and remained unchanged;
- `MultiUserEditing` was readable as false;
- the new workbook path was empty, qualifying the fixture as unsaved/local;
- a temporary number-format change was detected as an intervening planned-
  property change;
- the workbook closed, `Excel.Quit` returned, and the recorded PID exited.

The harness changes only a cell in its disposable unsaved workbook. It does not
toggle AutoSave or create/alter a shared workbook.

## Qualification result

- Debug solution build: passed with zero warnings and zero errors.
- Release solution build: passed with zero warnings and zero errors.
- Debug tests: 72 passed, 0 failed.
- Release tests: 72 passed, 0 failed.
- Packed debug XLL: registration, health, state restoration, normal formatting,
  protected/multi-area/merged refusals, and clean process exit passed.
- Real stale-property injection: the command planned against `General`, the
  debug harness changed the property to `0.00`, and execution refused without
  replacing it with the command's currency format.
- Collaboration-signal probe: passed with AutoSave unchanged and clean process
  exit on Excel 16.0 build 20228.

## Platform finding

Microsoft documents `AutoSaveOn` as a read/write Boolean whose false state may
also represent a disabled toggle, and documents `MultiUserEditing` as the legacy
shared-list state. Microsoft separately documents before/after remote-change
events for modern coauthor merges. These are complementary signals, not a
complete version log:

- [Workbook.AutoSaveOn](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.autosaveon)
- [Workbook.MultiUserEditing](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.multiuserediting)
- [About coauthoring in Excel](https://learn.microsoft.com/en-us/office/vba/excel/concepts/about-coauthoring-in-excel)
- [Workbook.BeforeRemoteChange](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.beforeremotechange)
- [Workbook.AfterRemoteChange](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.afterremotechange)

## Remaining exit-gate work

- qualify and wire remote-change events in the Excel-DNA host;
- run real cloud/coauthor pairs and supported build/channel fixtures;
- distinguish only those AutoSave states that can be identified without a write;
- exercise preview/execute/undo races, recalculation, external-link refresh,
  protection/read-only transitions, cancellation, close, and shutdown;
- review and accept exact production lease lifetimes for broader authority.

No command changes AutoSave or gains new collaborative mutation authority in
this slice.
