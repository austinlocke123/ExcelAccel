# WP-1A-02/03 command runtime evidence

- Date: 2026-08-19
- Scope: command contracts, canonical plans/results, execution authorization,
  and Excel readiness/adapter boundary
- Capability expansion: none

## Implemented

- registry metadata now requires stable ID/version, capability, context flags,
  impact, sorted changed properties, preview policy, undo policy, keyboard route,
  and acceptance IDs;
- command plans normalize changed properties and typed arguments into a stable
  ordinal order;
- canonical serialization is invariant, excludes localized summaries and all
  time/random data, escapes JSON deterministically, and produces a SHA-256 plan
  hash;
- execution authorization rejects registry/plan contract drift and requires the
  exact plan hash when preview is mandatory;
- results have an explicit success/refused/cancelled/failed/partial status plus
  changed/skipped counts and diagnostic/receipt fields;
- both implemented commands pass through the shared registry/plan authorization
  gate before revalidation and execution;
- ExcelInterop verifies the owning thread through the host-provided delegate and
  refuses while Excel reports it is not ready, covering edit/calculation/busy
  states without keyboard interception or speculative writes;
- COM access remains isolated in ExcelInterop and bounded retry remains limited
  to the approved transient HRESULT set.

## Verification

- canonical plans are byte-identical across argument/property enumeration order;
- exact preview hash is required and stale/wrong confirmation is refused;
- contract drift is refused before port access;
- every registered command has complete release metadata;
- project dependency and public COM-exposure tests continue to pass;
- Debug and Release builds/tests and packed-XLL real-Excel smoke are required
  before publication of this interval.

## Explicit exclusions

This interval does not register a new command, implement shortcuts, create
profiles, enable AutoColor, or grant any new collaborative mutation authority.
