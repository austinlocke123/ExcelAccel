# WP-1A-04/05 profiles and Quick Keys evidence

- Date: 2026-08-19
- Scope: versioned local profile, atomic persistence, default artifact, and
  deterministic Quick Key validation/state machine
- Live keyboard hook status: not enabled

## Profiles

- schema version and profile ID are mandatory;
- color cycles require normalized `#RRGGBB`; font sizes and thresholds are
  bounded; number-format keys are unique and ordinal;
- JSON rejects missing and unknown fields, excessive depth, unsupported schema,
  invalid cycles, duplicate formats, and files larger than 1 MiB;
- serialization order is deterministic;
- save writes a same-directory temporary file, reparses it, atomically replaces
  the target, and keeps the prior file as `.bak`;
- absence of a user profile loads the embedded default without an implicit
  filesystem write;
- the versioned `config/default-profile.json` defines reviewed conservative
  colors, number formats, preview threshold, navigation wrap, and bindings.

## Quick Keys

- assignments normalize case/spacing and allow one through three strokes;
- validation reports known native Excel/Windows/accessibility reservations,
  duplicates, and prefix ambiguity before activation;
- the state machine never handles a stroke while Excel edit mode is true;
- pending sequences cancel with Escape, expire within a bounded timeout, and
  reset on any unrecognized continuation;
- default bindings use the non-native `Ctrl+Shift+K` prefix and reference only
  registered fixed-parameter commands;
- the Ribbon KeyTip route remains the active production keyboard path.

Live `Application.OnKey` or a Windows keyboard hook is intentionally not
installed by this slice. Intercepting the first stroke before independently
qualifying edit-mode pass-through, timeout cleanup, add-in disable, and conflict
restoration would violate AC-KEY-001/003. The pure engine and profile contract
are ready; host activation remains part of WP-1A-11 qualification.

## Verification

- embedded default validates and references only registered commands;
- parse/serialize/parse is byte deterministic;
- repeated atomic save retains a valid backup and no temporary residue;
- unknown JSON is refused;
- reserved, duplicate, and prefix conflicts are detected;
- edit-mode pass-through, Escape, timeout, and successful multi-stroke routing
  are tested;
- Debug/Release build and 110-test suite pass;
- the expanded packed XLL embeds Persistence and Newtonsoft.Json and passes the
  real-Excel smoke test with clean process exit.
