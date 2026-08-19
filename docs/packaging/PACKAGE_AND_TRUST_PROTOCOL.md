# WP-P0-08 package and trust qualification protocol

Status: **Initial implementation; clean-VM lifecycle remains open**

## Safety boundary

The developer-workstation harness may:

- copy the already-packed Release XLL into ignored `.tools/packages/`;
- optionally Authenticode-sign only that copied artifact;
- generate a content-free manifest;
- verify length, SHA-256, embedded signer identity, and signature status;
- load the copied XLL with `RegisterXLL` in one hidden, temporary Excel process.

It MUST NOT:

- write Excel `OPEN` startup-registration values;
- add a Trusted Location or change Trust Center policy;
- add a certificate to Trusted Root or Trusted Publishers;
- overwrite a package directory;
- install under Program Files or LocalAppData;
- terminate a pre-existing Excel process;
- claim that a self-signed/untrusted signature passes production trust.

## Package contract

`New-ExcelAccelPackage.ps1` creates a versioned directory containing:

- `ExcelAccel-AddIn64-packed.xll` — the final packed artifact, signed after
  packing when a certificate is supplied;
- `package-manifest.json` — schema, product/version/runtime/architecture,
  artifact length/SHA-256, and observed Authenticode metadata.

The package builder refuses to overwrite an existing directory. The manifest is
bounded metadata, not an authenticity boundary. Production packaging must place
it inside a signed installer/container or use an approved signed catalog.

## Verification order

1. Resolve and validate the package root.
2. Validate manifest schema/product and relative artifact path.
3. Enforce the configured artifact-size cap.
4. Compare exact length and SHA-256.
5. Inspect Authenticode and compare the embedded signer thumbprint to the
   manifest when present.
6. When production mode is requested, require Windows to report `Valid`.
7. Report Mark of the Web; never remove it implicitly.
8. Only after verification, optionally load in an isolated Excel process.
9. Require workbook close, `Application.Quit`, exact PID exit, and unlocked XLL.

## Production signing profile

- digest: SHA-256;
- timestamp: RFC 3161 with SHA-256;
- certificate: CA-issued code-signing certificate with protected private key;
- target: final packed XLL and the selected installer/container;
- verification: Windows Authenticode policy plus explicit expected publisher;
- rollover: overlap old/new trusted publishers and retain rollback artifacts.

## Clean-VM lifecycle still required

On a disposable VM with supported x64 Excel and no development tools:

1. snapshot Excel/Office policy, startup-registration values, and processes;
2. install version A under a versioned per-user path;
3. register one owned startup entry and verify automatic load/health;
4. disable and verify Excel starts without ExcelAccel;
5. re-enable A, then install and switch to version B;
6. inject a B health failure and verify rollback to A;
7. uninstall and prove exact cleanup without changing unrelated entries;
8. repeat under representative signature/allowlisting/MOTW policies;
9. record crash, event-log, process, file-lock, and registry evidence.

No clean-VM step may kill Excel or replace a loaded XLL. If Excel is running,
the lifecycle operation refuses and asks the operator to close it.

`Test-ExcelAccelInstallLifecycle.ps1` rehearses the side-by-side version,
pointer, load, rollback, disable, and file-removal mechanics only under ignored
`.tools/install-sandbox/`. Its explicit `-AllowUntrustedPrototype` switch is
test-only. It never writes Excel startup registration, user installation
directories, certificate trust, or Office policy and therefore does not satisfy
the clean-VM lifecycle gate.

`Install-ExcelAccel.ps1` is the Phase 1A per-user installer source. Production
mode requires a Windows-valid package signature, refuses while Excel is open,
uses versioned LocalAppData directories, and owns one exact HKCU Excel `OPEN`
value. It includes install, upgrade, disable, enable, rollback, and uninstall
actions with `-WhatIf`. This source MUST NOT be treated as qualified until the
clean-VM procedure above passes with the selected signed installer/container.
