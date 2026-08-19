# Phase 1A safety-runtime and distribution-source evidence

Date: **2026-08-19**

## WP-1A-07 AutoColor

The deterministic pure planner classifies text, numeric hardcodes, same-sheet
formulas, cross-sheet formulas, external-workbook formulas, errors, and
unsupported cells. Plans are address ordered, contain category/count/color
changes without values, have a complete precondition fingerprint, enforce a
250,000-cell planning bound, and always require preview for worksheet scope.

Execution is deliberately fail-closed with
`PERFORMANCE_QUALIFICATION_REQUIRED`. The command is not registered or exposed
because AC-P0-006 has not passed full Qualification/UI-heartbeat review. This
implements the reviewable planner without granting unqualified Excel mutation.

## WP-1A-09 session undo

- Receipts are in-memory only, expire after eight hours, contain exact
  target/property/before/after/plan identity, and are capped at 20 per workbook.
- A later property/target change consumes the receipt and refuses without a
  write. Successful undo verifies the exact restored postcondition.
- Profile formatting creates a receipt only after successful postcondition
  verification. The Ribbon exposes one owned `Undo ExcelAccel` action.
- All receipts clear on add-in close/disable. `ClearWorkbook` is implemented;
  direct workbook-close event ownership remains support-matrix qualification.
- Real Excel smoke applied a font color, created a receipt, restored the exact
  original OLE color, closed the workbook, and observed natural process exit.

## WP-1A-10 recovery and diagnostics

- Unclean-process markers start the next session in conservative safe mode
  with no automatic workbook mutation, reopen, pane restore, scan, or replay.
- State-restoration failures quarantine the command and show a correlation ID.
- Diagnostics contain bounded operation/outcome/timing/exception-type fields;
  individual fields are capped and the log rotates at 1 MiB.
- User-initiated support export requires an explicit local destination and
  exact manifest confirmation, revalidates unchanged content, and uses
  temp/verify/replace with restoration on failure. It never transmits.

## WP-1A-11 UX, responsiveness, and keyboard safety

- Every registered action has a unique custom-tab Ribbon KeyTip route; tests
  reject duplicates. No built-in Excel shortcut is replaced.
- All Phase 1A formatting/navigation commands are discoverable in
  keyboard-accessible Ribbon menus.
- Dialogs are owned by Excel's window to preserve focus and z-order.
- The progress tracker is bounded and monotonic, supports cancellation before
  commit, and refuses cancellation once an atomic commit begins.
- No enabled operation is expected to exceed 500 ms. AutoColor and other broad
  work remain gated until non-modal progress and UI-heartbeat evidence pass.
- Live Quick Key interception remains disabled because edit-mode pass-through
  and cleanup have not been qualified in the host; its pure engine is ready.

## WP-1A-12 installer/update source

`scripts/Install-ExcelAccel.ps1` implements a per-user, versioned, side-by-side
install/upgrade/disable/enable/rollback/uninstall protocol. It refuses while
Excel is running, requires a Windows-valid signature by default, owns only a
child of `%LOCALAPPDATA%\ExcelAccel`, allocates one free HKCU Excel `OPEN` value
without overwriting another add-in, and changes/removes that value only if its
exact prior content still matches. An owner marker and resolved path boundary
guard recursive uninstall. It never changes Trusted Locations, certificate
stores, Office policy, HKLM, or unrelated values and supports `-WhatIf`.

The script was not run against the user's real Office registration.
Distribution remains blocked by CA-issued timestamped signing, allowlisting,
clean-VM lifecycle, and support-matrix gates; source completion is not
distribution approval.

## Automated verification

- Debug build: zero warnings/errors.
- Unit tests: **142 passed**, zero failed in both Debug and Release.
- Real hidden Excel smoke passed, including exact property undo and natural
  Excel process exit.
- Installer PowerShell AST is valid; static scan finds no HKLM, trust-store,
  Trusted Location, or process-termination behavior.
