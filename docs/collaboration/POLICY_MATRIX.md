# Proposed AutoSave and coauthoring policy matrix

- Work package: WP-P0-06
- Status: **Prototype evidence; ADR-0005 remains proposed**
- Governing decision: [ADR-0005](../decisions/ADR-0005-autosave-coauthoring.md)

## Detected states

### AutoSave

| State | Meaning |
|---|---|
| `On` | The read-only `Workbook.AutoSaveOn` probe returned true. |
| `OffOrDisabled` | The probe returned false. The add-in does not write the property to distinguish off from disabled. |
| `Unknown` | The property was absent or the read failed; mutation refuses. |

### Coauthoring

| State | Meaning |
|---|---|
| `NotDetected` | Qualified only for an unsaved workbook with readable AutoSave/legacy signals and no legacy sharing. |
| `PotentialModern` | AutoSave is on or a cloud URL is visible; this does not assert another user is currently present. |
| `LegacyShared` | `Workbook.MultiUserEditing` is true. This is not treated as modern coauthoring. |
| `RemoteChangeObserved` | A qualified remote-change event occurred during the session. |
| `RemoteChangeInProgress` | A before-remote-change signal has not yet been paired with completion. Mutation refuses. |
| `Unknown` | Exclusivity cannot be proven. A local-looking path may still be synchronized cloud storage. |

## Impact-tier policy

All mutations require exact workbook identity and property-fingerprint
revalidation immediately before execution. Any changed revision, fingerprint,
identity, or closed workbook refuses. The lease clock is session-only and is not
part of canonical plan serialization.

| Impact | Local unsaved/not detected | AutoSave on or potential modern | Legacy/unknown/in-progress |
|---|---|---|---|
| `read_only` | permitted | permitted | permitted, except a closed/different workbook |
| `low` | permitted with fresh fingerprint | permitted with fresh fingerprint and bounded lease | refused |
| `medium` | requires eligible receipt and bounded lease | additionally requires qualified remote-event invalidation | refused |
| `high` | sync policy permits only the proven local state; command-specific rollback/preview gates still apply | refused | refused |

Production lifetime values are intentionally not chosen by this spike.
`CollaborationPolicyLimits` requires an explicit caller-supplied value for each
mutation tier and enforces a five-minute hard ceiling. Tests use short fixture
values only.

## Version and fingerprint contract

- Each relevant local edit, remote-change boundary, structure change,
  recalculation, save completion, protection/read-only transition, or AutoSave
  signal increments a checked monotonic revision.
- Event coverage is never considered complete. A current opaque SHA-256
  precondition fingerprint is required even when the revision did not change.
- Fingerprints use invariant length-prefixed components so concatenation and
  locale cannot alter identity. Inputs are limited to 4,096 components and
  1,000,000 characters.
- Fingerprints, leases, and event state remain in memory; they are not logged,
  persisted, or presented as cloud version history.
- Workbook close invalidates the tracker permanently. Later events are rejected.

## Current adapter boundary

`ExcelSelectionAdapter` now includes a collaboration snapshot captured through
read-only `AutoSaveOn`, `MultiUserEditing`, and `Path` probes. Remote-change
events are deliberately reported unsupported until an event sink is qualified.
The existing low-impact currency command also hashes and revalidates its planned
number-format precondition, so an intervening property change refuses before the
write.

The general collaboration policy is not yet wired to production command
dispatch. That remains gated on ADR acceptance and the evidence below.

## Required work before ADR acceptance

- wire and fault-test application/workbook before/after remote-change events;
- test actual OneDrive, OneDrive for Business, and SharePoint coauthor sessions;
- qualify AutoSave on, manually off, disabled, read-failure, and state-transition
  cases across the supported build/channel matrix;
- inject local and remote changes during snapshot, preview, execution, rollback,
  and undo;
- prove medium-impact receipt eligibility and high-impact refusal in the real
  host;
- exercise recalculation, external-link, protection, read-only, cancellation,
  close, and add-in shutdown races.
