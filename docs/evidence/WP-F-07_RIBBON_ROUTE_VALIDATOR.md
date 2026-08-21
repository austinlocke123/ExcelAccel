# WP-F-07 Ribbon route validator

Date: 2026-08-20
Status: Complete

## Contract

- **Capability:** CAP-UX-002
- **Acceptance:** AC-FMT-033
- **Allowed implementation:** a validating test suite, the route corrections it
  found, and the removal of the silent fallback that hid them.
- **Excluded:** ribbon regrouping (delivered in WP-F-04), any command behaviour.

## Why this existed

`docs/commands/RIBBON_LAYOUT.md` stated that routes are generated from the
ribbon definition and that "a validator parses the ribbon XML and fails on any
duplicate KeyTip, any prefix collision, and any button without an action."

Neither was true. `RibbonRoutes.cs` is hand-maintained, and no test read the
ribbon XML at all. The document described the right behaviour, so the code was
brought to meet it rather than the document weakened.

The reason nothing caught the drift is that `RibbonRoutes.For` fell back to the
Command Search route on an unknown id. A typo'd or renamed id therefore produced
a descriptor that looked routed and advertised a keyboard path that did nothing,
and the existing route-uniqueness assertion still passed because the wrong
strings were still distinct.

## What the validator found

Ten descriptors were advertising routes that disagreed with the ribbon. Four
were known; six were not.

| Command | Advertised | Actual | Effect |
|---|---|---|---|
| `model_check.run.selection` | `Alt, X, A, AM, EMS` | `Alt, X, A, AM, S` | dead route |
| `model_check.run.worksheet` | `Alt, X, A, AM, EMW` | `Alt, X, A, AM, W` | dead route |
| `model_check.run.workbook` | `Alt, X, A, AM, EMB` | `Alt, X, A, AM, B` | dead route |
| `model_check.rescan` | `Alt, X, A, AM, EMR` | `Alt, X, A, AM, R` | dead route |
| `navigate.cell.a1` | `Alt, X, A, V, A` | `Alt, X, A, VN, A` | dead route |
| `style.apply` | label `Alt, X, A, Y, L` | `Alt, X, A, EY, L` | wrong label |
| `style.apply_builtin` | label `Alt, X, A, Y, B` | `Alt, X, A, EY` | wrong label |
| `favorite.add` | `Command Search: Ctrl+D` | `Alt, X, A, Q, then Ctrl+D` | two spellings |
| `favorite.remove` | `Command Search: Ctrl+Shift+D` | `Alt, X, A, Q, then Ctrl+Shift+D` | two spellings |
| `favorite.invoke` | `Command Search: Enter` | `Alt, X, A, Q, then Enter on a favorite` | two spellings |

The four Model Check descriptors built their route by concatenating a *different*
command's route with a KeyTip fragment. `navigate.cell.a1` named a menu KeyTip
`V` that does not exist; the Navigate menu is `VN`. Every one of these was
printed verbatim by Command Search and the shortcut cheat sheet, so a user
following them pressed keys that did nothing.

All ten now read from `RibbonRoutes`, which is the single source. The dead
`keytip` parameter left behind on the Model Check factory was removed.

## The fallback is gone

`RibbonRoutes.For` now throws `KeyNotFoundException` naming the id instead of
returning a plausible-looking wrong answer. Every registered command already had
a real entry, so the fallback was pure dead weight that existed only to hide
mistakes. A command genuinely reached some other way still belongs in the table
with an honest route, which is how the three favorite commands are handled.

## What is enforced now

- Every ribbon control has an `onAction` and a `keytip`.
- No KeyTip duplicates another, or is a prefix of a longer one, within the scope
  Excel resolves it in. Excel resolves a KeyTip the moment it is unambiguous, so
  a single-letter KeyTip makes every longer one starting with that letter
  unreachable.
- Every tagged button names a registered command.
- Every route in the table belongs to a registered command or a built-in style.
- Every descriptor's route and shortcut label equal the table's entry.
- Every ribbon-hosted command's table entry equals the path its button actually
  has, derived by walking the XML.

The ribbon XML is read from source rather than from `ExcelAccelRibbon`, because
the test project deliberately does not reference the net48 host and
`ArchitectureBoundaryTests` asserts that exact project graph. Reading a source
file to check a repository invariant is the pattern that test already uses for
the `.csproj` dependency graph.

## Verification

```
build Release   0 warnings, 0 errors
build Debug     0 warnings, 0 errors
test  Release   546/546 passed (was 540/540)
smoke           scripts/Test-ExcelAddIn.ps1 PASS
                Excel exited, 0 surviving processes, 0 stale session markers
```

Falsification: the suite was written before the fixes and failed on all ten
descriptors, which is how they were found. Reverting any single route
reintroduces its failure.
