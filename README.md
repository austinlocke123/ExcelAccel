# ExcelAccel

ExcelAccel is a native Windows Excel add-in for fast, deterministic,
keyboard-first financial-modeling workflows.

The repository has **closed Phase 0, Phase 1A/1B, and Phase 2**.
It contains a real 64-bit Excel-DNA `.xll`, a pure command core, reliability
boundaries, profile-driven property-scoped formatting, read-only navigation,
optimistic session undo, recovery/diagnostics, per-user installer source, the
Phase 1B daily-speed feature core, Phase 2 formula auditing through indirect
traversal, and the Model Check rule engine.
Core, Application, ExcelInterop, and the thin Excel-DNA host are separated by an
enforced acyclic dependency graph.

It **is installable and usable on a developer machine today**; see
[Session handoff](docs/SESSION_HANDOFF.md) for the commands. It is not qualified
for distribution to other people, and retained formula, performance,
collaboration, and clean-machine gates remain fail-closed. See
[Project status](docs/PROJECT_STATUS.md) for the current snapshot, verification
numbers, and the restart point.

## Start here

1. [Documentation map](docs/README.md)
2. [Requirements](docs/REQUIREMENTS.md)
3. [Architecture](docs/ARCHITECTURE.md)
4. [Reliability and responsiveness](docs/RELIABILITY.md)
5. [Complete feature and command specification](docs/commands/README.md)
6. [Number format cycles](docs/commands/FORMAT_CYCLES.md)
7. [Ribbon layout and keyboard routing](docs/commands/RIBBON_LAYOUT.md)
6. [Acceptance criteria](docs/ACCEPTANCE.md)
7. [Implementation plan](docs/IMPLEMENTATION_PLAN.md)
8. [Traceability matrix](docs/TRACEABILITY.md)
9. [Original feature coverage](docs/FEATURE_COVERAGE.md)
10. [Architecture decisions](docs/decisions/README.md)
11. [Phase 0 implementation evidence](docs/evidence/PHASE0_VERTICAL_SLICE.md)
12. [Excel adapter and state-guard evidence](docs/evidence/WP-P0-04_ADAPTER_STATE_GUARD.md)
13. [Formula strategy spike evidence](docs/evidence/WP-P0-05_FORMULA_STRATEGY.md)
14. [AutoSave and coauthoring spike evidence](docs/evidence/WP-P0-06_AUTOSAVE_COAUTHORING.md)
15. [Performance baseline protocol](docs/performance/BASELINE_PROTOCOL.md)
16. [WP-P0-07 performance evidence](docs/evidence/WP-P0-07_PERFORMANCE_BASELINE.md)
17. [Package and trust protocol](docs/packaging/PACKAGE_AND_TRUST_PROTOCOL.md)
18. [WP-P0-08 packaging evidence](docs/evidence/WP-P0-08_PACKAGING_TRUST.md)
19. [Phase 0 closure and gate ledger](docs/evidence/PHASE0_CLOSURE.md)
20. [WP-1A-01 production foundation evidence](docs/evidence/WP-1A-01_FOUNDATION.md)
21. [WP-1A-02/03 command runtime evidence](docs/evidence/WP-1A-02_03_COMMAND_RUNTIME.md)
22. [WP-1A-04/05 profiles and Quick Keys evidence](docs/evidence/WP-1A-04_05_PROFILES_QUICK_KEYS.md)
23. [WP-1A-06/08 formatting and navigation evidence](docs/evidence/WP-1A-06_08_FORMATTING_NAVIGATION.md)
24. [WP-1A safety runtime and distribution-source evidence](docs/evidence/WP-1A-07_09_10_11_12_SAFETY_RUNTIME.md)
25. [WP-1B command search and favorites evidence](docs/evidence/WP-1B-01_COMMAND_SEARCH_FAVORITES.md)
26. [WP-1B style recipe evidence](docs/evidence/WP-1B-02_STYLE_RECIPES.md)
27. [WP-1B profile exchange evidence](docs/evidence/WP-1B-03_PROFILE_EXCHANGE.md)
28. [WP-1B formula foundation evidence](docs/evidence/WP-1B-FORMULA_FOUNDATION.md)
29. [WP-1B formula command evidence](docs/evidence/WP-1B-05_07_FORMULA_COMMANDS.md)
30. [WP-1B advanced formula planner evidence](docs/evidence/WP-1B-06_08_ADVANCED_FORMULA_PLANNERS.md)
31. [WP-1B data cleaning evidence](docs/evidence/WP-1B-09_10_DATA_CLEANING.md)
32. [WP-1B selection tools evidence](docs/evidence/WP-1B-11_SELECTION_TOOLS.md)
33. [WP-1B qualification evidence](docs/evidence/WP-1B-12_QUALIFICATION.md)
34. [WP-2-01 reference snapshot and direct precedents evidence](docs/evidence/WP-2-01_REFERENCE_SNAPSHOT.md)
35. [WP-2-02 direct dependents evidence](docs/evidence/WP-2-02_DIRECT_DEPENDENTS.md)
36. [WP-2-03 indirect traversal evidence](docs/evidence/WP-2-03_INDIRECT_TRAVERSAL.md)
37. [WP-2-05 to WP-2-08 Model Check evidence](docs/evidence/WP-2-05_08_MODEL_CHECK.md)
38. [WP-2-02b workbook scope evidence](docs/evidence/WP-2-02B_WORKBOOK_SCOPE.md)
39. [WP-2-09 Phase 2 qualification evidence](docs/evidence/WP-2-09_PHASE2_QUALIFICATION.md)
40. [Project status and restart guide](docs/PROJECT_STATUS.md)
41. [Session handoff](docs/SESSION_HANDOFF.md)

## Build and test

Prerequisites are Windows, 64-bit desktop Excel, and a .NET 8 SDK. The shipped
add-in targets .NET Framework 4.8; reference assemblies are restored as a build
dependency, so the developer targeting pack is not required.

```powershell
dotnet restore ExcelAccel.sln
dotnet build ExcelAccel.sln --configuration Debug --no-restore
dotnet test ExcelAccel.sln --configuration Debug --no-build --no-restore
./scripts/Test-ExcelAddIn.ps1
./scripts/Test-ExcelFormulaOracle.ps1
./scripts/Test-ExcelCollaborationSignals.ps1
./scripts/Test-ExcelPerformance.ps1 -Profile Quick
./scripts/Test-ExcelPhase2Qualification.ps1 -Profile Quick
./scripts/Test-ExcelReliabilitySoak.ps1 -Iterations 10
./scripts/New-ExcelAccelPackage.ps1 -Version 0.0.0-local
./scripts/Test-ExcelAccelPackage.ps1 -PackageDirectory ./.tools/packages/ExcelAccel-0.0.0-local-x64 -LoadInExcel
```

The packed debug add-in is produced at
`src/ExcelAccel.ExcelAddIn/bin/Debug/net48/publish/ExcelAccel.ExcelAddIn-AddIn64-packed.xll`.
The smoke script launches a hidden temporary Excel process, verifies XLL
registration and the health UDF, exercises the real currency-format command,
checks content preservation, and requires clean Excel shutdown.
The formula-oracle script drives a temporary workbook from the versioned parser
corpus and requires exact qualified native `Formula` round-trip plus clean
process exit. R1C1 remains excluded after a recorded lifecycle failure; the
script does not register a formula command.
The collaboration-signal script performs read-only AutoSave and legacy-sharing
probes in a temporary workbook, demonstrates intervening-property detection,
and requires clean process exit.
The performance script uses only generated temporary workbooks and writes
machine-specific JSON under ignored `.tools/performance/`; `Quick` validates the
harness, while `Qualification` is required before budgets can be proposed.
The packaging scripts create and verify only ignored local package copies by
default. Production qualification additionally requires a valid CA-issued,
timestamped signature and the clean-VM lifecycle in the packaging protocol.

## Source material

The original draft and its review are retained under [`PRD/`](PRD/). They are
inputs to the Markdown specification, not the normative implementation source.

- `PRD/Native_Windows_Excel_Financial_Modeling_AddIn_PRD.docx`
- `PRD/PRD_Review_Feedback.md` — intentionally untracked, so it is not present in
  a fresh clone

## Current boundary

Phase 0 vertical-slice, reliability, formula-strategy, collaboration-safety,
performance-baseline, and packaging/trust prototype work is implemented. Phase
1A begins with the production foundation. Feature families remain governed by
their command contracts and retained gates; Phase 0 closure does not imply
release readiness.
