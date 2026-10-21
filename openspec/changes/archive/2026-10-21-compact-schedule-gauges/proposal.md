# Proposal

## Why

Collapsed schedule cards use a wide progress bar and several stacked lines while leaving the actual missing duties hidden. A small ring gauge and a compact group of red vacancy badges will make staffing needs immediately readable and reduce vertical scrolling, especially on phones.

## What Changes

- Replace schedule progress bars and separate percentage lines with a small ring gauge containing an occupied/required count such as `5/8`.
- Show one red badge per missing duty type beside the ring, grouping multiple free positions as `Verkauf (2×)` and using the existing statistics badge styling and vacancy semantics.
- Apply the presentation consistently to games, preparation, cleanup and configured cake delivery; retain explicit setup-needed and zero-cake states.
- Place content-sized badges directly beside each other, wrapping to a new line only when the available width is filled. Reduce collapsed-card spacing and keep the existing `Details` caption and triangle in the summary without a separate full-width action row.
- Update gauges and badges together from saved assignment/configuration state, preserving separate eligibility warnings and the occupied status of duties seeking replacement.
- Keep the initial presentation useful without JavaScript and retain keyboard, touch, guest privacy, filtering and candidate-loading behavior.

## Capabilities

### New Capabilities

None.

### Modified Capabilities

- `schedule-overview`: Replace bar/percentage requirements with accessible ring gauges, add grouped vacancy badges and a compact responsive summary layout, and preserve eligibility and cake special states.

## Impact

- Presentation and initial view data: `templates/schedule.html`, `static/style.css`, and schedule/block view builders in `webapp.py`.
- Saved-state updates: gauge, game staffing, block assignment and cake rendering paths in `static/app.js`. Reuse existing `db.staffing_status()` vacancy data and block capacities; add public count/label fields to existing block responses only where needed for consistent saved-state refreshes.
- Existing schedule, mobile, cake, privacy and JavaScript regressions need targeted updates and responsive visual verification; document the overview in `README.MD`.
- No schema migration, new runtime dependency, notification change or change to binding assignments, eligibility, authorization or statistics behavior.
