# WP-1B-03 profile exchange evidence

Date: **2026-08-19**

## Package boundary

- Export uses a strict, non-archive JSON envelope. It contains only product and
  schema identifiers, component counts, an explicit zero asset count, payload
  length/SHA-256, and base64 of the already validated declarative profile JSON.
- The package therefore has no archive paths, scripts, macros, binary assets,
  workbook formulas/values/names, screenshots, or executable content. Unknown
  envelope/profile fields and any nonzero asset count refuse.
- Fixed favorite arguments are refused from export until a typed binder can
  prove they contain no workbook-derived content. Current host commands already
  refuse nonempty fixed arguments.
- Package and profile sizes are bounded at 2 MiB and 1 MiB respectively. Payload
  length, hash, schema, product, component counts, command IDs, command contract
  versions, Quick Key IDs, style schema, and profile schema all validate before
  an import plan exists.

## Export/import transactions

- Export preview binds the explicit destination, overwrite decision, existing
  destination hash, package hash, counts, and manifest into one plan hash.
  Existing files require explicit overwrite. Destination drift refuses.
- Export writes an owned same-directory temporary file, reparses it, atomically
  moves/replaces, verifies the final hash, and restores the prior destination if
  post-write verification fails. Temporary backup names are unique and removed.
- Import preview is read-only and reports deterministic Quick Key, favorite, and
  local-style additions/replacements/removals. It binds the source hash, current
  profile hash, payload hash, and path into one plan hash.
- Import apply revalidates the exact source bytes, active profile bytes, package,
  command references, and confirmed plan. It writes one bounded `.import.bak`,
  then uses validate/temp/atomic-replace activation. Source/current drift changes
  nothing. The in-memory effective profile changes only after storage succeeds.
- The Excel host uses explicit owned Open/Save dialogs and manifest/diff
  confirmation. It never launches or transmits the package.

## Shortcut cheat sheet

- CSV and accessible HTML exports contain only current normalized Quick Keys,
  command IDs/names/categories, and deterministic conflict text.
- Both formats use explicit destination/overwrite preview, exact plan hash,
  same-directory temporary write, atomic replacement, and final hash verification.

## Verification

- Debug and Release builds: zero warnings/errors.
- Unit suite: **169 passed**, zero failed.
- Tests cover deterministic package bytes, exact-plan confirmation, round-trip,
  existing-file refusal, destination drift preservation, source/current drift,
  bounded backup, unknown command/version handling, payload tamper, asset refusal,
  unqualified fixed-argument refusal, and deterministic CSV/HTML output.
- Real hidden Excel smoke passed the complete search/style/formatting/undo/
  navigation/fault-restoration suite after the new Ribbon callbacks were added,
  with natural workbook/application close and no remaining Excel process.

## Retained limits

- Organization policy/locked-setting distribution is not implemented; there is
  currently no policy layer whose conflicts could be imported. The importer
  must gain explicit lock-diff/refusal behavior before such a layer is enabled.
- Profile packages are local trust-neutral configuration, not signed executable
  distribution artifacts. Production add-in signing gates remain unchanged.
