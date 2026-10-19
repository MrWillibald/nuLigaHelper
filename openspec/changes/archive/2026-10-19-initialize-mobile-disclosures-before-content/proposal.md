# Proposal

## Why

On mobile, navigating between pages briefly shows expanded Hinweise and Filter panels before the main JavaScript bundle closes them. Applying their native disclosure state before their contents are parsed prevents the visible expansion and resulting page movement, including when the main bundle loads slowly.

## What Changes

- Initialize shared responsive disclosures as their headers are parsed, using the existing 700 CSS pixel breakpoint.
- Use a small same-origin external script compatible with the current production Content Security Policy.
- Preserve desktop expansion, native keyboard/touch operation, expanded introduction/filter fallbacks when JavaScript is unavailable, and validation-error visibility.
- Make early and later initialization share one controller so the main bundle cannot reset an earlier user choice, unsaved edit or focused form.
- Verify initial rendering while the main bundle is deliberately delayed, rather than only checking disclosure state after page initialization.

## Capabilities

### New Capabilities

None.

### Modified Capabilities

- `public-page-presentation`: Add a shared requirement that existing responsive disclosure defaults take effect before their contents can first appear, with safe fallback and initialization continuity.

## Impact

The implementation affects the shared disclosure initializer in `static/app.js`, a small new external asset, and disclosure markup in `templates/schedule.html`, `templates/persons.html` and `templates/statistik.html`. Shared layout asset preparation may also affect `templates/base.html`. Existing tests and `README.MD` need to cover the earlier initialization. No API, domain, database or provider changes are required.

This proposal requires the existing `improve-mobile-usability` implementation and synchronized specifications from commit `119b9908f7b70721efe65d51a0a1ebccbe802624`, now included in `master` through merge commit `f4e1ef0`. Apply this follow-up on that baseline or a descendant; integrating the prerequisite is not part of this change.

All existing responsive disclosures share the initializer, including statistics and person maintenance. Their documented default states, permissions, data and presentation remain the baseline; this change adjusts initialization timing and preserves interaction state.
