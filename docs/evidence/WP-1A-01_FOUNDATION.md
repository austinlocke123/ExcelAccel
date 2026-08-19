# WP-1A-01 production foundation evidence

- Date: 2026-08-19
- Branch: `agent/phase-1a-foundation`
- Scope: production solution skeleton and architecture enforcement
- Capability expansion: none

## Implemented boundary

The Phase 0 combined assembly has been split at the production dependency
seams:

- `ExcelAccel.Core` remains `netstandard2.0` and owns immutable models, pure
  policies, formula/parser primitives, reliability primitives, performance
  math, and package-integrity primitives;
- `ExcelAccel.Application` is `netstandard2.0`, references only Core, and owns
  command contracts, registry, planning, results, and port abstractions;
- `ExcelAccel.ExcelInterop` is x64 `net48`, references Application and Core,
  and owns Excel snapshot/mutation adapters, COM retry, owned-proxy cleanup, and
  Excel application-state adaptation;
- `ExcelAccel.ExcelAddIn` is the thin x64 `net48` Excel-DNA host and composes
  the root `ExcelDnaUtil.Application` provider plus the owning-thread verifier
  into ExcelInterop.

ExcelInterop does not reference ExcelAddIn. The host owns lifecycle and passes
two delegates into the adapter; no Excel application RCW enters Application or
Core.

## Enforced dependency graph

An automated test reads the production project files and requires exactly:

```text
ExcelAccel.Core
  ^
  |
ExcelAccel.Application
  ^             ^
  |             |
ExcelAccel.ExcelInterop
  ^
  |
ExcelAccel.ExcelAddIn
```

The host also references Core directly for the existing callback/reliability
boundary. Cycles, hidden extra project references, and a premature
`ExcelAccel.PresentationInterop` project fail the test.

Both pure assemblies are inspected for forbidden Excel-DNA/Office references
and public COM exposure. The packed-XLL build explicitly embeds Core,
Application, ExcelInterop, and the host assembly; omission of any dependency is
a release-blocking regression.

## Verification

- Debug build: zero warnings and zero errors;
- Release build: zero warnings and zero errors;
- test suite: 100 passed, 0 failed;
- pack log: Core, Application, ExcelInterop, and ExcelAddIn embedded;
- first real-Excel attempt: mutation correctly refused because a prior forced
  harness stop left an unclean-session marker; content remained unchanged;
- clean recovery restart: packed XLL registered, health function returned
  `1.0.0.0`, property-only formatting passed, content was preserved, injected
  state-guard failure restored state, unsafe targets were refused, workbook
  closed, and Excel exited naturally.

## Explicit exclusions

- no new command or shortcut;
- no formula mutation or AutoColor;
- no profile persistence;
- no preview UI or undo receipt;
- no installer, registry, Trust Center, or update mutation;
- no expansion of collaborative mutation authority.

WP-1A-02 may now build the production command registry/context/plan/result
pipeline on these dependencies. WP-1A-03 may extend the adapter without creating
a reverse reference into the host.
