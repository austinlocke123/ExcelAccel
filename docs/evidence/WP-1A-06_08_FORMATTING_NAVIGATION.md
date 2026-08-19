# WP-1A-06/08 formatting and navigation evidence

Date: **2026-08-19**  
Branch: `agent/phase-1a-formatting-navigation`

## Implemented scope

- The formatting catalog has explicit versioned descriptors for every Phase
  1A command in Catalog sections 1.1 and 1.2.
- Profile schema v2 supplies deterministic font/fill/font-size, alignment,
  underline, row-height, column-width, and named number-format recipes.
- Each executable formatting plan declares exactly one property, fingerprints
  its before-state, revalidates selection/safety/property state immediately
  before mutation, writes through the Excel state guard, and verifies the
  postcondition. A mismatch cannot report success.
- Decimal changes use a conservative invariant transform and refuse
  scientific, conditional/color, fractional, or otherwise unqualified
  formats rather than guessing.
- Freeze panes is medium impact and requires confirmation of the exact plan
  hash. The compact Ribbon route therefore refuses it until a preview UI can
  provide confirmation; unfreeze remains directly executable.
- Navigation covers visible-sheet previous/next, A1, used-range endpoints,
  populated-region edges, bounded back/forward history, and bounded
  session-only bookmarks. It does not write workbook content or persist state.
- Excel adapters enforce readiness and Excel-thread affinity, use bounded COM
  retry, and explicitly release owned child proxies.

## Verification

- Debug solution build: zero warnings and zero errors.
- Unit tests: **123 passed**, zero failed.
- Real hidden Excel smoke: XLL registered; the profile font-color cycle changed
  only font color; A1 navigation selected the exact target; and all original
  state-restoration, stale-plan, protection, unsupported-selection,
  content-preservation, resource, workbook-close, and process-exit checks
  passed.

## Reliability disposition

The commands operate only on a validated selection and never save. The profile
is loaded once per add-in session; an invalid local profile is logged without
raw content and the embedded default is used. Live Quick Key capture remains
disabled; actual routes are custom-tab Ribbon KeyTips and do not replace
Excel's built-in key bindings.

Session property receipts are not claimed here. Descriptors declare the
WP-1A-09 policy; the next work package supplies the bounded optimistic receipt
store before formatting undo is exposed.
