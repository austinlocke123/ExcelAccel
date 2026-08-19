# WP-P0-08 packaging and trust evidence

- Date: 2026-08-19
- Branch: `agent/phase-0-packaging-trust`
- Status: **Initial local feasibility slice; AC-P0-008 remains open**

## Implemented

- bounded, path-safe artifact descriptor and SHA-256/length verifier;
- deterministic versioned package directory and manifest generator;
- optional post-pack Authenticode signing through Windows SignTool;
- signature status and signer-thumbprint verification;
- production gate requiring Windows `Valid` signature status;
- Mark-of-the-Web detection without implicit removal;
- isolated real-Excel package load/health/close/Quit/PID verification;
- sandboxed side-by-side lifecycle rehearsal that is structurally limited to
  ignored `.tools/install-sandbox/`;
- proposed ADR-0006 and a clean-VM qualification protocol.

## Automated evidence

| Configuration | Build | Tests |
|---|---|---|
| Debug | zero warnings/errors | 97/97 passed |
| Release | zero warnings/errors | 97/97 passed |

The unsigned control package:

- contained one 692,736-byte packed x64 XLL;
- matched manifest SHA-256
  `1BB2D7A1A6C1D478DAB0CCE46662C5A90AD309B1C14C421A3F8DFDC95EC60DAC`;
- was explicitly reported `NotSigned`;
- loaded in isolated Excel 16.0, returned `1.0.0.0`, and exited cleanly;
- was refused by `-RequireValidSignature`.

The signing-mechanics package:

- used a seven-day, non-exportable, self-signed Current User test certificate;
- signed the copied packed XLL successfully with SHA-256;
- embedded signer thumbprint
  `AABD64D4E4119E93C721B882FFB5896D118C6E61`;
- matched its post-sign manifest hash;
- loaded and exited cleanly;
- was correctly reported `UnknownError`/untrusted and refused by the production
  signature gate;
- did not add the certificate to Trusted Root or Trusted Publishers;
- removed the test certificate from Current User personal storage after use,
  with zero matching certificates remaining.

Both package trees and their raw manifests are under ignored `.tools/packages/`
and are not release artifacts.

The test-only side-by-side lifecycle rehearsal also passed:

- installed and loaded `0.0.8-phase0-unsigned` as version A;
- installed and loaded `0.0.8-phase0-selfsigned` as version B;
- switched back to and loaded version A;
- disabled by atomically removing the active-version pointer;
- removed both owned version directories, pointer metadata, and sandbox root;
- left no Excel process or install sandbox behind.

The first rehearsals exposed transient post-load XLL file-lock and post-Quit
process-observation races. Cleanup now uses bounded 200 ms retries, and final
shutdown uses a bounded ten-second observation window. It does not terminate
Excel; failure to exit remains an explicit test failure.

## What this proves

- the packed XLL can be signed after packing without preventing Excel load;
- package integrity and signer identity can be checked before Excel sees the XLL;
- unsigned and untrusted artifacts can be refused deterministically;
- the local spike does not need Trust Center or startup-registry mutation.
- side-by-side file/pointer mechanics can support upgrade, rollback, disable,
  and uninstall without overwriting a loaded XLL.

## Remaining gates

- CA-issued production certificate and protected-key workflow;
- RFC 3161 timestamp and Windows `Valid` verification;
- signed installer/container selection;
- clean-VM per-user install and automatic startup load;
- disable, version upgrade, injected-failure rollback, and uninstall;
- exact registry ownership/restoration evidence;
- Mark-of-the-Web, trusted-publisher, enterprise allowlisting, and certificate
  rollover qualification;
- supported Excel/Windows build matrix.
