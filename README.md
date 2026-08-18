# ExcelAccel

ExcelAccel is a proposed native Windows Excel add-in for fast, deterministic,
keyboard-first financial-modeling workflows.

The repository is currently in **specification review**. No implementation has
been authorized or started. The goal of the current documentation is to make
the feature set and engineering constraints precise enough for human review
before agents begin coding.

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

## Source material

The original draft and its review are retained under [`PRD/`](PRD/). They are
inputs to the Markdown specification, not the normative implementation source.

- `PRD/Native_Windows_Excel_Financial_Modeling_AddIn_PRD.docx`
- `PRD/PRD_Review_Feedback.md`

## Current boundary

This repository intentionally contains documentation only. Creating a Visual
Studio solution, selecting dependencies, writing production code, or publishing
an add-in requires an explicit post-review decision.
