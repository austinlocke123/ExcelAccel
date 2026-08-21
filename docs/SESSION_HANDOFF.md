# Session handoff

Written: 2026-08-20. Read [`PROJECT_STATUS.md`](PROJECT_STATUS.md) first for the
engineering snapshot; this file covers what a fresh session would otherwise have
to rediscover.

## Repository state

- Branch `main`, clean tree, everything pushed. Head is `9bca976`.
- **528/528 Release tests pass.** Release and Debug builds are warning-free.
- The hidden-Excel smoke passes with the process-exit check **actually working**
  (see the defect note below) and leaves no surviving Excel process.

## The app is installed and usable right now

This is the thing most easily missed. ExcelAccel is not waiting on anything to be
used on this machine:

```powershell
# last done at version 0.3.0-local, which is what is installed today
./scripts/New-ExcelAccelPackage.ps1 -Version "0.1.0-local"
./scripts/Install-ExcelAccel.ps1 -Action Install `
  -PackageDirectory ".tools/packages/ExcelAccel-0.1.0-local-x64" -AllowUntrustedPrototype
```

It registers per-user under `HKCU:\Software\Microsoft\Office\16.0\Excel\Options`
and installs to `%LOCALAPPDATA%\ExcelAccel\app\versions\<version>\`. Excel shows
an **ExcelAccel** tab; everything is reachable via `Alt, X, A`.
`-Action Uninstall` removes it. **Every Excel process must be closed** before any
install, upgrade, rollback, or uninstall — the installer refuses otherwise.

The deferred distribution gates (CA-signed installer, clean-VM lifecycle,
enterprise trust) govern giving it to *other people*. They never blocked personal
use, and `-AllowUntrustedPrototype` exists for exactly this.

### Pending reinstall

`main` contains a usability change that is **not yet in the installed copy**:
successful commands no longer pop a dialog. To pick it up, close Excel and repeat
the package and install commands above with a new version string.

## Phase 2 is complete

WP-2-01, 2-02a, 2-02b, 2-03, 2-04, 2-05 through 2-08 (Model Check), and 2-09
qualification are all merged. There is **no Phase 3**: the implementation plan
§7 deliberately says the remaining work is thirteen individually gated packages
(WP-G-01…13), each needing its own approval, not one combined phase.

## Decisions taken this session

- **Workbook-scale performance gate: opened, bounded.** Previously a stop sign
  where `IMPLEMENTATION_PLAN` §4 and `PROJECT_STATUS` disagreed. Resolved on the
  strength of WP-2-09 measurements. Read-only workbook scanning only; workbook
  scale *mutation* stays out of scope. Documents are reconciled.
- **No dialog on success.** A command that works must not interrupt the user;
  outcomes go to the Excel status bar. Dialogs are reserved for refusals,
  failures, partial results, and faults. This is a standing UI principle, not a
  one-off change.

## Two defects found and fixed, worth remembering

1. **The Excel-exit check was dead across six harness scripts.** They matched
   `(?m)^excel_pid=(\d+)$` against CRLF output; in .NET the `$` anchor never
   matches after the digits, so the captured id was always empty and the whole
   "did Excel exit?" block was skipped. Five of the six predate Phase 2, so
   **every "Excel exited naturally" claim in evidence written before 2026-08-20
   was unverified**, not false on purpose but never checked.
2. **A Debug-only smoke hook leaked an Excel process.**
   `ExcelAccel.Smoke.TraceNavigate` took `ActiveWorkbook` and `Selection` into
   locals and never released them, ignoring the `ComRelease.Owned` pattern every
   production adapter follows. Never in shipped code, but it survived six merges
   because the check above could not see it.

The lesson that generalises: when a verification step has never once failed,
confirm it can fail before trusting it.

## Known open items

- **`AutoClose` is never exercised.** `addin.close` has been logged zero times
  against 263 `addin.open` events, so the unload path, runtime resets, and
  recovery-marker cleanup are entirely untested. Highest-value reliability gap.
- **AutoColor is planner-only.** `AutoColorPlanner` is complete and tested but
  registered nowhere; execution is hard-stopped with
  `PERFORMANCE_QUALIFICATION_REQUIRED`. It still needs a transactional adapter,
  rollback and fault-injection evidence, and worksheet-scale preview UI. Note the
  document conflict: `PROJECT_STATUS` lists AutoColor as disabled while WP-G-13
  assumes workbook AutoColor gets built.
- **WP-1A-12 dependency is ambiguous.** WP-G-04, G-09, G-11, and G-13 depend on
  it; the installer *source* exists but its GA gates are deferred. Settle whether
  "depends on WP-1A-12" means the source or the qualification before starting any
  of them. Same class of conflict as the workbook-scale gate.
- No reference-hardware performance baseline; WP-2-09 numbers are one developer
  workstation, Debug build, single corpus shape.
- Model Check ignores live in their own atomic file rather than the profile
  schema; peer regions are column runs only; structural literal exclusions are
  per-function rather than per-argument-position.

## Operational notes for whoever runs the harness

- Force-killing Excel leaves `*.running` markers in
  `%LOCALAPPDATA%\ExcelAccel\sessions`, which puts the **next** run into safe
  mode and makes mutation commands refuse. That is the recovery design working,
  but it looks like an unrelated formatting failure. Clear the markers before
  re-running.
- Never kill Excel by name. The user keeps real workbooks open; target the
  specific PID the harness reports.
- A hung Excel holding the packed `.xll` will fail the next build with
  "could not be deleted. (Perhaps loaded in Excel?)".

## Suggested next step

Use the add-in on a real model and let what irritates you set the backlog. If a
package is wanted instead, the unblocked read-only ones are WP-G-02 external-link
inventory and WP-G-01 named-range inventory, both of which reuse the shared trace
view, registration pattern, and export-with-manifest that already exist.
