# ADR-0001: Excel-DNA host

- Status: **Accepted**
- Date: 2026-08-18
- Accepted: 2026-08-19
- Decider: project owner

## Context

The add-in requires native Windows keyboard integration, Ribbon and task panes,
large-range performance, Excel C API and COM access, offline operation, and
event/lifecycle control. Office.js does not expose the required integration
surface. The original draft selected Excel-DNA and retained VSTO as a fallback.

## Decision

Use Excel-DNA as the only designed production host. Do not build a dual-host
abstraction or implement VSTO in parallel.

Retain one time-bounded contingency trigger during Phase 0: reconsider the host
only if Excel-DNA fails an approved must-pass deployment, lifecycle, task-pane,
keyboard, signing, support-matrix, or crash-safety gate and the failure cannot be
resolved within the Phase 0 evidence budget.

## Consequences

- Host-specific code is isolated in `ExcelAccel.AddIn` and
  `ExcelAccel.ExcelInterop`.
- The pure domain/application boundary remains useful for testing, not as a
  speculative promise to support another host.
- Phase 0 focuses on proving Excel-DNA rather than repeating a broad comparison.
- If a contingency gate fails, a new ADR compares concrete evidence and costs;
  agents do not silently fall back to VSTO.

## Required evidence

- AC-P0-001, AC-P0-002, and AC-P0-008;
- callback exception containment and clean disable/unload;
- task-pane and Quick Key feasibility;
- supported x64 Excel clean-VM matrix;
- signed packaging feasibility;
- cold/warm startup and soak results.

## Acceptance note

The host decision is accepted for Phase 1A source implementation. The packed
x64 XLL, callback boundary, property mutation, fault restoration, repeated
fresh-process soak, and local signing mechanics passed. Clean-VM, coexistence,
task-pane, full disable/unload, and signed-distribution qualification remain
release gates; failure of one reopens this ADR rather than silently introducing
a second host.
