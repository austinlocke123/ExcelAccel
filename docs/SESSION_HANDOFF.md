# Session handoff

Written 2026-08-24 for the agent picking this up next.

Read [`PROJECT_STATUS.md`](PROJECT_STATUS.md) for delivered behaviour per work
package and [`IMPLEMENTATION_PLAN.md`](IMPLEMENTATION_PLAN.md) §2 for the work
package contract you are expected to follow. This document covers what those two
do not: who did what, where the seams are, and what will bite you.

## Repository state

- Branch `main`, clean tree, everything merged and pushed.
- **771/771 Release tests**, Release and Debug builds warning-free.
- `scripts/Test-ExcelAddIn.ps1` passing, Excel exiting cleanly, no stale session
  markers.
- Phase 2 complete; WP-F, WP-R, and WP-G packages delivered as listed in
  `PROJECT_STATUS.md`.

Verify before you touch anything:

```bash
./.tools/dotnet/dotnet.exe build ExcelAccel.sln --configuration Release --no-restore
./.tools/dotnet/dotnet.exe test  ExcelAccel.sln --configuration Release --no-build --no-restore
./.tools/dotnet/dotnet.exe build ExcelAccel.sln --configuration Debug   --no-restore
```

`TreatWarningsAsErrors` is on, so any warning is a build failure. There is no CI;
these commands and the PowerShell harness are the entire gate.

---

## What Claude Opus 5 did, and how to review it

Everything from commit **`72bbe8f`** (exclusive) to `HEAD` was written by Claude
Opus 5, starting 2026-08-19.

| | |
|---|---|
| Last commit before this work | `72bbe8f` — *Capture WP-2-01 direct precedents* |
| First commit of this work | `168597e` — *Complete WP-2-01 direct-precedent presentation and registration* |
| Commits in the repo | 137 total, **49 are this work** |
| Attribution | **49 of 49** non-merge commits in range carry `Co-Authored-By: Claude Opus 5 (1M context)`. No commit outside the range carries any AI attribution. |
| Pull requests | **#27 through #55** |
| Direct to `main` | 18 non-merge commits, mostly specification and status; one implementation commit (`9bca976`, no-dialogs-on-success) broke the PR convention deliberately after the user declined a reset |
| Size | 190 files changed, +26,903 / −1,286 |

Scope it with `git log 72bbe8f..HEAD` or `git diff 72bbe8f..HEAD`.

**WP-2-01 straddles the boundary.** Its capture layer predates this work; the
presentation and registration on top of it do not. If you sample that package,
half of it is not this author's.

### Where to look first if you are reviewing

Ordered by where real problems are most likely, not by size:

1. **COM interop adapters** — `ExcelSelectionAdapter`, `ExcelAutoColorAdapter`,
   `ExcelNameInventoryAdapter`, `ExcelLinkInventoryAdapter`,
   `ExcelComparisonAdapter`. The test project targets `net8.0` and cannot
   reference these `net48` projects, so **none of this code has a unit test**.
   Its only coverage is the real-Excel smoke. Check COM release discipline: every
   acquired object needs `ComRelease.Owned` in a `finally`, and a leaked RCW keeps
   Excel alive past `Quit`.
2. **Profile schema v5 → v6 migration** (`ProfileStore.ToDefinition`). It runs
   against real user profiles on upgrade, and can only be exercised through
   `ProfileStore.Parse` — `ProfileDefinition` cannot represent a v5 profile at
   all. A mistake here silently costs a user their settings.
3. **WinForms dialogs** in `ExcelAccel.ExcelAddIn`. No test project exists for
   that assembly. The logic behind them was deliberately pushed into
   `ExcelAccel.Application`, so the dialogs should be thin; check that they are.
4. **The decisions listed below.** Those are judgement calls, and judgement is the
   thing most worth a second opinion.

### Decisions a reviewer should second-guess

These were reasoned about and recorded, not stumbled into. Each could defensibly
have gone the other way:

- **AutoColor applies no allowlist**, so `=A1*2` is blue, while Model Check
  allowlists the same literal. The two deliberately disagree and a test asserts
  they keep disagreeing. See `commands/AUTOCOLOR.md`.
- **A hardcode outranks external and cross-sheet.** `='[Other.xlsx]Sheet1'!A1+5`
  is blue, not red. This trades away the external signal on that one cell.
- **The default font colour cycle is palette order, not classification
  precedence.** Precedence would have led with red and changed what a keypress
  had always done. `AC-FMT-041` was reworded to match.
- **Model Check ignores stay in their own file**, not the profile. Folding them in
  would let a damaged suppression list cost the user every setting they have.
  `AC-CHECK-031` was reworded to match (WP-R-05).
- **Workbook AutoColor was withdrawn** from WP-G-13, because the approved command
  contract rules out workbook scope and `commands/` outranks `ACCEPTANCE.md`
  (WP-R-02).
- **Number-format validation is advisory, not constructor-enforced**, for the same
  blast-radius reason as the ignore store (WP-F-09).

Where an acceptance criterion was reworded rather than met, the criterion says so
inline and names the work package. Search `ACCEPTANCE.md` for "Reworded",
"Partially met", and "withdrawn".

### What was not claimed

Every partial delivery is marked partial in `ACCEPTANCE.md` and explained in its
`docs/evidence/WP-*.md`. The ones that matter:

- AutoColor **worksheet** scope refuses at planning time, pending performance
  qualification, worksheet-scale rollback evidence, and a real preview surface.
- Name **usage** coverage (AC-NAME-008..010) is unbuilt; each usage category is
  separately qualified by its spec and none has been.
- Link scanning covers formulas and defined names only. Chart series, queries,
  and validation are reported as coverage gaps, never counted as zero.
- Comparison covers ranges. Workbook structure with sheet pairing
  (AC-CMP-011..013) and the frozen timing corpus (AC-CMP-015) are unbuilt.

---

## The add-in on this machine is stale

**Installed: `0.3.0-local`.** It predates everything from 2026-08-20 onward —
cycles, AutoColor, the inventories, comparison, all of it.

To pick it up, with every Excel closed:

```powershell
./scripts/New-ExcelAccelPackage.ps1 -Version "0.4.0-local"
./scripts/Install-ExcelAccel.ps1 -Action Upgrade `
  -PackageDirectory ".tools/packages/ExcelAccel-0.4.0-local-x64" -AllowUntrustedPrototype
```

`-Action Rollback` returns to 0.3.0-local.

Upgrading migrates `%LOCALAPPDATA%\ExcelAccel\profile.json` from schema 5 to 6 on
first read. The migration keeps every setting, but a migrated profile shows
one-entry cycles where a fresh install shows three; reset is how to adopt the new
defaults. **That migration has never run against a real user profile** — only
against the embedded default and test fixtures.

---

## Traps that have already cost time

- **Never kill Excel by name.** `Stop-Process -Name EXCEL` was run once here and
  killed the user's open workbook. Target the reported PID. A force-kill also
  leaves `*.running` markers in `%LOCALAPPDATA%\ExcelAccel\sessions`, which puts
  the **next** Excel session into safe mode where every mutation command refuses —
  and that looks like an unrelated formatting bug. Clear the markers before
  re-running.
- **A hung Excel holding the packed `.xll` fails the next build** with "could not
  be deleted. (Perhaps loaded in Excel?)".
- **The smoke worksheet is a shared fixture.** Assertions count scanned formulas
  exactly, so adding cells anywhere on `Sheet1` breaks unrelated checks. Seed late,
  after the scan assertions, and say why in a comment.
- **Excel rewrites locale-qualified currency formats** on assignment:
  `[$£-en-GB]` becomes `[$£-809]`. Such an entry can never match itself, so its
  cycle sticks on entry one. Bare symbols round-trip; see `commands/FORMAT_CYCLES.md`.
- **Worksheet names cap at 31 characters.** The Phase 2 privacy needle is 26, so
  appending anything to it overflows.
- **When a verification step has never once failed, prove it can fail.** A dead
  regex anchor meant the harness's Excel-exit check silently skipped for 263
  sessions, and `AutoClose` had never run in 290. Both hid real defects.

---

## How work is done here

- One work package per PR, branch `agent/<slug>`, merged with a merge commit.
- Every package lands a `docs/evidence/WP-*.md` and updates `TRACEABILITY.md`,
  `FEATURE_COVERAGE.md`, and `PROJECT_STATUS.md`. A package is not done without
  this; the user asked for it explicitly so future agents can see intent.
- Commit subjects are imperative, sentence case, no prefix. Bodies are prose
  explaining *why* and what was rejected, wrapped near 72 columns. Spec-only
  commits end with `No code changes.`
- Run the Release tests every package, and the real-Excel smoke whenever the Excel
  adapter, the host, or command wiring changes.
- Put testable logic in `Core` or `Application`. `ArchitectureBoundaryTests`
  asserts the exact project reference graph and forbids COM types on any public
  Core or Application API, so it will stop you adding even a harmless-looking
  reference.
- `RibbonRoutes.For` throws on an unknown id, and `RibbonRouteTests` compares
  every descriptor against the ribbon XML. Add the route when you add the button.

## Harnesses

| Script | What it proves |
|---|---|
| `Test-ExcelAddIn.ps1` | The main smoke, ~80 asserted evidence lines |
| `Test-ExcelPhase2Qualification.ps1` | Phase 2 budgets across three corpus shapes |
| `Test-ExcelReliabilitySoak.ps1` | Process cleanup across fresh sessions |
| `Test-ExcelInProcessRetention.ps1` | Retention **inside** one long-lived session |
| `Test-ExcelFormulaOracle.ps1` | Formula round-trip against the versioned corpus |

The soak and the retention harness answer different questions and neither
substitutes for the other: the soak uses a fresh process per iteration, so it can
never show in-process growth at any iteration count.

## Suggested next step

Install `0.4.0-local` and use it on a real model. Nothing built since 2026-08-20
has been used in anger, and the friction from real use is worth more than the next
plan row. `PROJECT_STATUS.md` § "Recommended restart point" lists the engineering
work if that is wanted instead.
