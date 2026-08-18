# Documentation map and authority

Status: **Draft for review**

This folder is the reviewable engineering baseline for ExcelAccel. It separates
facts that change at different rates and prevents an implementation agent from
treating an idea, a design option, and an accepted requirement as equivalent.

## Document authority

When documents disagree, use this order and open a review issue rather than
silently choosing an interpretation:

1. [`REQUIREMENTS.md`](REQUIREMENTS.md) owns supported behavior, scope,
   invariants, and non-functional requirements.
2. Accepted records under [`decisions/`](decisions/) own explicit architectural
   decisions. A proposed ADR is not yet authoritative.
3. [`ARCHITECTURE.md`](ARCHITECTURE.md) owns component boundaries, dependency
   direction, execution flow, and platform rules.
4. [`RELIABILITY.md`](RELIABILITY.md) owns crash-safety, recovery,
   responsiveness, resource, and observability rules.
5. [`commands/`](commands/) owns precise user-visible feature and command
   contracts for every retained phase.
6. [`ACCEPTANCE.md`](ACCEPTANCE.md) owns pass/fail evidence for release.
7. [`IMPLEMENTATION_PLAN.md`](IMPLEMENTATION_PLAN.md) owns sequencing only. It
   cannot weaken a requirement or acceptance criterion.
8. [`TRACEABILITY.md`](TRACEABILITY.md) maps requirements to commands,
   decisions, work packages, and acceptance evidence.
9. [`FEATURE_COVERAGE.md`](FEATURE_COVERAGE.md) maps each original functional
   area to its detailed contract and current disposition.

The files under [`../PRD/`](../PRD/) are retained source material and review
history. They are non-normative after this specification is approved.

## Normative language

- **MUST / MUST NOT**: required for the stated phase.
- **SHOULD / SHOULD NOT**: expected unless an ADR records a justified exception.
- **MAY**: optional behavior that cannot be assumed by callers.
- **Proposed**: a recommendation awaiting review.
- **Deferred**: explicitly outside the current implementation phase.
- **Open**: unresolved; agents must not invent the answer.

## Change control

- Requirements, accepted ADRs, architecture boundaries, and release criteria
  require explicit human review.
- Command descriptions may evolve through normal pull requests if they do not
  add a capability, weaken safety, or change a release gate.
- Examples, aliases, help text, and default shortcut proposals are lightweight
  changes.
- Any new workbook mutation, network dependency, COM integration surface,
  persistence mechanism, or background execution path requires an architecture
  and reliability review.

## Review checklist

Before implementation is authorized, reviewers should confirm:

- scope and deferrals are intentional;
- every Phase 0 decision has a falsifiable exit gate;
- no proposed ADR is being treated as accepted;
- command mutation, preview, undo, and refusal behavior is unambiguous;
- crash-safety and responsiveness requirements are testable;
- every normative requirement has acceptance evidence or a recorded gap;
- no work package requires a hidden architectural choice.
