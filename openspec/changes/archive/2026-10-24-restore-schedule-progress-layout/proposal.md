# Proposal

## Why

The recent overview changes replaced familiar horizontal staffing bars with rings and vacancy badges and compressed the card layout. Restore the earlier presentation shown in the user's reference screenshot so occupancy and percentage are readable across the card again.

## What Changes

- Restore horizontal progress bars with a descriptive German occupied/required count on the left and percentage on the right for games, preparation, cleanup and configured positive cake quantities.
- Remove staffing rings and vacancy badges from the schedule summaries.
- Restore the card spacing and responsive arrangement from immediately before commit `98694d3`: prominent time above title and metadata, responsible team for games, information above progress on mobile, and information beside progress on desktop. Keep `Details` aligned right beneath the progress/status area.
- Keep saved-state updates, accessible disclosure and progress semantics, eligibility warnings, cake setup/zero states, guest privacy and current assignment behavior.

## Capabilities

### New Capabilities

None.

### Modified Capabilities

- `schedule-overview`: Replace ring/badge requirements with horizontal occupancy bars and descriptive counts/percentages, restore the earlier responsive card layout, and retain independent eligibility and configuration states.

## Impact

- Presentation changes in `templates/schedule.html`, `static/style.css` and progress rendering in `static/app.js`; adjust existing Python/JavaScript regressions and browser verification fixtures accordingly.
- Reuse current authoritative game staffing and block snapshots, including candidate-refresh responses; no API removal, database migration, new dependency or notification change is needed.
- Update the overview description in `README.MD` and relevant test documentation during implementation.
- Current age rules, seven required adult positions including one Wischer, binding assignments, replacement requests, locks, filtering and statistics remain in force. Historical commits serve as presentation references, not whole-application rollback targets.
