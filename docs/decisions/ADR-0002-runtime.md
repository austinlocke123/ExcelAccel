# ADR-0002: Runtime and target frameworks

- Status: **Accepted**
- Date: 2026-08-18
- Accepted: 2026-08-19
- Decider: project owner

## Context

Excel-DNA supports .NET Framework and modern .NET. Reliability and deployment
inside another application's process are more important than using the newest
runtime features. Runtime presence, add-in coexistence, isolation, startup,
signing, and clean-machine behavior must be considered together.

Excel-DNA's current runtime guidance notes that .NET Framework is present on
Windows and provides AppDomain isolation, while modern .NET requires a matching
desktop runtime and only one modern .NET runtime version can be loaded in an
Excel process. See <https://excel-dna.net/docs/guides-basic/dotnet-runtime-support/>.

## Decision

Use:

- `.NET Framework 4.8` for the Excel-DNA host and Excel interop projects;
- `netstandard2.0` for pure domain/application libraries where doing so does not
  compromise required APIs, correctness, or testability;
- SDK-style project files and the latest approved C# compiler/tooling compatible
  with those targets.

This is a recommendation for Phase 0 validation, not an accepted implementation
choice.

## Rationale

- minimizes external runtime prerequisites;
- favors isolation and coexistence inside Excel;
- supports a broad Windows/Excel estate;
- keeps pure libraries portable for fast tests;
- reduces deployment variables while core crash-safety mechanisms mature.

## Costs and risks

- fewer modern runtime APIs and runtime performance features;
- possible library constraints;
- additional care when sharing DTO projects across targets;
- potential future migration cost.

## Alternatives to measure

- modern .NET LTS with framework-dependent deployment;
- modern .NET LTS with an approved packaging model;
- a single-target modern .NET solution if deployment/coexistence evidence is
  materially better than expected.

## Required evidence

- clean VM with no developer tools;
- x64 Excel across proposed supported builds;
- coexistence with representative add-ins;
- cold/warm startup and memory;
- pane/Ribbon/event lifecycle;
- disable/unload and repeated open/close soak;
- package signing, update, and rollback;
- dependency availability and security support.

## Acceptance note

The `.NET Framework 4.8` x64 host plus `netstandard2.0` pure-core split is
accepted for Phase 1A. It has produced deterministic Debug/Release builds,
packed XLLs, isolated Excel runs, and a pure-core test suite without an
additional runtime deployment prerequisite. Clean-machine, coexistence, and
signed installer evidence remain WP-1A-12 release qualification.
