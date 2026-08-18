# ADR-0005: AutoSave and coauthoring safety policy

- Status: **Proposed**
- Date: 2026-08-18
- Deciders: open

## Context

Cloud-hosted workbooks commonly use AutoSave, and shared workbook changes may
merge into the local copy. Plans, cached snapshots, receipts, and rollback
assumptions can become stale between capture and mutation. Available event
coverage must not be assumed to be a complete transaction/version log.

## Proposed decision

1. Detect AutoSave and all qualified coauthoring indicators in command context.
2. Never silently change AutoSave.
3. Use immediate property/structure precondition revalidation for every
   mutation, regardless of impact.
4. Permit low-impact, property-scoped mutations only after Phase 0 proves their
   postconditions and stale-plan refusal.
5. Permit medium-impact mutation only with fresh snapshot, bounded plan lifetime,
   preview policy, and eligible session receipt.
6. Refuse high-impact mutation during AutoSave/coauthoring until a later ADR
   proves a bounded safe transaction.
7. Invalidate caches and pending plans conservatively when relevant events or
   identity changes occur.
8. Treat inability to determine a safe state as a refusal, not as AutoSave off.

## Rejected default behavior

- automatically turning AutoSave off and back on;
- assuming a user prompt makes an unsafe transaction safe;
- relying on events as a complete external-change log;
- keeping a preview executable after intervening workbook change;
- persisting undo as a substitute for cloud version history.

## Consequences

- some commands may refuse in shared/cloud workbooks;
- lower feature reach is accepted to protect workbook and process integrity;
- capability may expand after real build-specific evidence;
- UI must explain the refusal and safe alternatives without implying the
  workbook is corrupt.

## Required evidence

- AC-P0-007 and AC-SYNC-001..003;
- AutoSave on/off/disabled states;
- local and remote edits between snapshot/preview/commit/undo;
- recalculation and external-link changes;
- supported build/channel matrix;
- protection, read-only, and shared-file transitions;
- command cancellation and add-in shutdown during relevant states.
