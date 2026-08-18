# ExcelAccel

ExcelAccel is a native Windows Excel add-in for fast, deterministic,
keyboard-first financial-modeling workflows.

The repository is in **Phase 0 implementation**. The first vertical slice is a
real 64-bit Excel-DNA `.xll` with a pure command core, a reliability boundary,
one read-only command, and one property-scoped formatting command.

## Start here

1. [Documentation map](docs/README.md)
2. [Requirements](docs/REQUIREMENTS.md)
3. [Architecture](docs/ARCHITECTURE.md)
4. [Reliability and responsiveness](docs/RELIABILITY.md)
5. [Complete feature and command specification](docs/commands/README.md)
6. [Acceptance criteria](docs/ACCEPTANCE.md)
7. [Implementation plan](docs/IMPLEMENTATION_PLAN.md)
8. [Traceability matrix](docs/TRACEABILITY.md)
9. [Original feature coverage](docs/FEATURE_COVERAGE.md)
10. [Architecture decisions](docs/decisions/README.md)
11. [Phase 0 implementation evidence](docs/evidence/PHASE0_VERTICAL_SLICE.md)
12. [Excel adapter and state-guard evidence](docs/evidence/WP-P0-04_ADAPTER_STATE_GUARD.md)
13. [Formula strategy spike evidence](docs/evidence/WP-P0-05_FORMULA_STRATEGY.md)

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

## Source material

The original draft and its review are retained under [`PRD/`](PRD/). They are
inputs to the Markdown specification, not the normative implementation source.

- `PRD/Native_Windows_Excel_Financial_Modeling_AddIn_PRD.docx`
- `PRD/PRD_Review_Feedback.md`

## Current boundary

Only Phase 0 vertical-slice, reliability, and formula-strategy prototype work is
implemented. Feature families remain governed by their command contracts and
phase gates; successful spikes do not mark proposed ADRs or the broader
acceptance matrix complete.
