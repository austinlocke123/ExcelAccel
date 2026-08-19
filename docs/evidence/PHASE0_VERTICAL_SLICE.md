# Phase 0 command vertical-slice evidence

- Date: 2026-08-18
- Scope: partial WP-P0-01, WP-P0-02, WP-P0-03, and WP-P0-04 evidence
- Status: implemented and locally passing; no ADR or Phase 0 gate is accepted by
  this evidence alone

## Implemented slice

- Excel-DNA 1.9.0, SDK-style `net48`, x64-only Excel host
- `netstandard2.0` command core with no Excel-DNA or Office interop reference
- packed single-file debug and Release XLL output
- Ribbon tab and keyboard KeyTips
- stable command metadata with required keyboard route
- read-only `inspect.selection.summary` command
- property-scoped `format.number.currency` command
- selection identity revalidation and 50,000-cell Phase 0 limit
- single-command reentrancy gate
- top-level Ribbon/lifecycle exception containment
- Excel-thread ownership check before COM access
- owned-child COM release policy; borrowed Excel application RCW is not released
- scoped screen-updating guard with no-throw restoration
- unclean-session marker and mutation safe mode
- local content-free diagnostic log
- debug-only automated Excel integration hook excluded from Release builds

## Reference environment

- Windows x64
- Microsoft Excel x64 at `C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE`
- Excel version `16.0.20228.20190`
- local .NET SDK `8.0.424`
- add-in runtime target `.NET Framework 4.8`

## Automated results

`dotnet build ExcelAccel.sln --configuration Debug --no-restore`

- passed with zero warnings and zero errors
- produced `ExcelAccel.ExcelAddIn-AddIn64-packed.xll`
- pack log confirmed both `ExcelAccel.ExcelAddIn.dll` and
  `ExcelAccel.Core.dll` were embedded

`dotnet test ExcelAccel.sln --configuration Debug --no-build --no-restore`

- 10 passed, 0 failed
- covers command-plan mutation invariants, stale-selection refusal, bounded
  selection refusal, exact property write, reentrancy, unique command IDs,
  keyboard-route presence, and pure-core dependency enforcement

`./scripts/Test-ExcelAddIn.ps1`

- packed XLL registration returned `True`
- `EXCELACCEL.VERSION()` returned `1.0.0.0`
- the actual Excel-hosted currency command produced
  `$#,##0.00;($#,##0.00);-`
- the selected cell's value and formula were unchanged
- the temporary workbook closed
- `Excel.Quit()` returned and the test Excel PID exited

## Crash found and corrected

The first packed-XLL smoke attempt caused an Excel APPCRASH with managed
exception code `0xe0434352`. The unpacked XLL loaded and calculated correctly,
which isolated the failure to packaging. The pack log showed that the host DLL
was embedded but `ExcelAccel.Core.dll` was not.

The host project now explicitly includes `ExcelAccel.Core.dll` in
`ExcelAddInInclude`. Subsequent pack logs show both assemblies embedded, and
repeated watchdog-controlled packed-XLL smoke tests pass without a new Windows
Error Reporting crash or orphaned Excel process.

This is a regression condition: a build whose pack evidence omits a required
project dependency must fail qualification and must not be opened in a user's
Excel session.

## Remaining Phase 0 gaps

- clean-VM and signed-artifact installation
- repeated open/close and command soak runs
- representative add-in coexistence matrix
- Ribbon visual/accessibility inspection
- edit-mode and shortcut collision detection beyond safe Ribbon KeyTips
- protected, merged, multi-area, array, spill, AutoSave, and coauthoring policy
  qualification
- fault injection at every Excel adapter operation
- session receipt/undo qualification
- startup, latency, memory, handle, and COM-proxy budgets
- task-pane feasibility and full disable/unload behavior

Until those gaps close, the XLL is an engineering artifact, not a release.
