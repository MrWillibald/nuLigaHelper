## Why

The home-game overview shows every assignment field at once, making a full game day difficult to scan. Visitors need to narrow the view to a date and see staffing progress before opening individual games or day tasks.

## What Changes

- Add a single-date dropdown to the existing overview filters, with an all-dates option. The chosen date limits games and their preparation and cleanup blocks and combines with the existing filters.
- Apply the person filter to each game and day-task block independently. Hide day-task blocks whenever a responsible-team filter is selected; a matching block alone can keep its date visible.
- Start every game, preparation block, and cleanup block in a compact state. Each compact card shows its identifying time and title plus a labeled staffing progress bar. Compact game cards also show the responsible team, or an open placeholder when none is assigned; opening a card reveals its assignment fields and, for games, the responsible-team field.
- Count the five required game slots toward game progress. Keep optional `Unterstützung` assignable in the expanded card without counting it toward completion. Count all three slots for each day-task block.
- Make a selected past date visible inside the past-games section without an extra expand action.

## Capabilities

### New Capabilities

None.

### Modified Capabilities

- `schedule-overview`: Add date filtering and compact, expandable game and day-task presentation with staffing progress.

## Impact

The schedule builder and route in `webapp.py`, `templates/schedule.html`, `static/style.css`, and assignment interactions in `static/app.js` will change. Schedule and day-block web tests will cover filtering, default collapsed state, progress counts, and guest-safe rendering. No schema migration or new dependency is expected.
