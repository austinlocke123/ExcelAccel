# WP-R-05 Model Check ignore storage

Date: 2026-08-24
Status: Settled. Ignores stay in their own atomic store; the spec and the
acceptance criterion now describe what the code does.

## Contract

- **Capability:** CAP-CHECK-001
- **Acceptance:** AC-CHECK-030..033
- **Allowed implementation:** the decision, the documents that state it, and
  tests that lock it in.
- **Excluded:** any change to how ignores are matched or applied.

## The real conflict

`MODEL_CHECK.md` said ignore/unignore performs an "atomic profile write" and that
unignore removes "from the local profile". AC-CHECK-031 said it "changes only the
atomic local profile". The implementation writes a separate TSV beside the
profile, at `%LOCALAPPDATA%\ExcelAccel\model-check-ignores.tsv`.

So this item was not a vague preference. The spec and the code disagreed, and one
of them had to move.

## Why the separate store wins

**Blast radius, which decides it.** `ProfileStore` refuses a profile it cannot
parse whole, and `ProfileRuntime` falls back to the embedded default when a load
throws. Folding ignores into the profile means a damaged suppression list costs
the user every cycle, colour, quick key, and favorite they have. That is the same
reasoning that kept number-format validation advisory in WP-F-09: a feature must
not be able to destroy unrelated settings.

**Wrong shape.** An ignore is a fingerprint of a finding in one model. A profile
is user-wide and portable. An exported profile would carry entries that can never
match on the other machine, which is noise at best.

**Budget.** 2,048 entries is roughly 200 KB against the profile's 1 MiB ceiling —
a fifth of the budget for data that is not settings.

**Cost.** Folding means a schema bump to 7, a migration, and a ripple through the
profile package manifest, for no user-visible benefit.

## What changed

The spec and AC-CHECK-031 now describe the ignore store, name the file, and give
the reason. Nothing about matching, bounding, or application changed.

Six tests lock the decision in. Two of them assert the profile carries no ignore
data and that `ProfileDefinition` exposes no ignore member, so folding them in
later fails loudly instead of drifting in. The rest cover what the separate store
provides: atomic replace with validation, no temporary file left behind,
duplicate fingerprints stored once, a note with tabs or newlines unable to
reshape its row, and a missing file reading as an empty set rather than an error.

## Already met

AC-CHECK-033's visible-and-removable half is satisfied by the existing
manage-ignores dialog, and ignores are not portable by any unapproved route,
which is what "portable only through an explicitly approved export/import action"
restricts.

## Verification

```
build Release   0 warnings, 0 errors
test  Release   740/740 passed (was 734/734)
```

No smoke run: nothing reachable from Excel changed.
