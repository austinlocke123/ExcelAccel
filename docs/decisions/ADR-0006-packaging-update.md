# ADR-0006: Packaging, signing, installation, and rollback

- Status: **Accepted**
- Date: 2026-08-19
- Accepted: 2026-08-19
- Decider: project owner

## Context

ExcelAccel must ship a packed x64 Excel-DNA XLL that can be authenticated,
installed without development tools, disabled, upgraded, rolled back, and
removed without leaving Excel unstable. Office trust policy, Mark of the Web,
publisher trust, enterprise allowlisting, and locked in-process binaries are
separate concerns and must not be collapsed into "the file loaded once."

The Phase 0 source guidance is:

- Excel-DNA produces a single packed XLL and recommends signing the packed XLL;
- Microsoft recommends SHA-256 Authenticode plus an RFC 3161 SHA-256 timestamp;
- Office Trusted Locations bypass multiple protection layers and should be
  narrowly scoped and centrally controlled when used;
- persistent XLL startup registration can use Excel's `OPEN`, `OPEN1`, ...
  values, but an installer must own and restore only its exact entry.

## Decision

1. Pack first, then Authenticode-sign the final XLL with SHA-256 and an RFC 3161
   SHA-256 timestamp.
2. Production installation refuses an unsigned, untrusted, expired,
   wrong-publisher, hash-mismatched, size-mismatched, or Mark-of-the-Web-blocked
   artifact. A test override cannot exist in the production installer.
3. Install per user by default under a versioned directory such as
   `%LOCALAPPDATA%\ExcelAccel\versions\<version>\`. Never replace an XLL loaded
   in a running Excel process.
4. Keep the prior qualified version until the new version loads and passes a
   bounded health check. Upgrade switches one owned startup-registration entry;
   rollback restores that exact prior entry.
5. Installation, upgrade, rollback, disable, and uninstall require Excel to be
   closed. The installer does not terminate user Excel processes.
6. Do not create a broad Office Trusted Location automatically. Enterprise
   deployment uses a trusted publisher/allowlisting policy or an explicitly
   approved narrow location with least-privilege ACLs.
7. Record a signed installer/container decision separately. The current JSON
   manifest supplies integrity and audit metadata but is not itself an
   authenticity boundary; authenticity currently comes from the signed XLL.
8. Persistent registry mutation is qualified only on a disposable clean VM.
   Developer-workstation tests use `RegisterXLL` in an isolated Excel process
   and make no Trust Center or startup-registration change.

## Consequences

- Updates are side-by-side and recoverable rather than in-place.
- A real CA-issued certificate, protected signing service/key, and timestamp
  service are required before external qualification.
- Managed environments may still refuse the add-in until administrators approve
  its publisher/path/hash under their policy.
- Clean-VM lifecycle testing is mandatory before accepting this ADR.

## Evidence required for acceptance

- valid CA-issued and timestamped signature verified with Windows policy;
- expected publisher/thumbprint policy and certificate rollover procedure;
- clean-VM install and first load without developer tools;
- persistent registration, disable, upgrade, rollback, and uninstall with exact
  registry before/after evidence;
- Mark-of-the-Web and enterprise allowlisting cases;
- x64 supported-build matrix and coexistence testing;
- no orphaned file, registry entry, process, or trusted-location exception;
- accepted installer/container technology and privilege model.

## Acceptance note

The versioned, pack-then-sign, exact-entry ownership, no-forced-Excel-termination
design is accepted for implementation. This is not approval to distribute the
prototype: CA-issued signing, timestamping, installer selection, registry
ownership, allowlisting, rollover, and clean-VM lifecycle evidence remain hard
WP-1A-12 gates.
