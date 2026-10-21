# Proposal

## Why

The overview currently places the start time beside the matchup or day-task label when space permits, making individual card details harder to scan. A consistent vertical order with a slightly larger start time makes the schedule easier to read on phones and desktop screens.

## What Changes

- Give every game summary four separate information rows, in order: start time; matchup (or existing Spielfest identity); age class, hall and game number together; responsible team or the existing open placeholder.
- Give preparation, cleanup and cake-delivery summaries three separate information rows: time or existing time placeholder (including any date qualifier); task label; existing task metadata such as "Tagesdienst" or cake quantity. Blocks continue to have no responsible team or empty fourth row.
- Increase every card's start time or time placeholder slightly from its current size and make it larger than the matchup or task-label text.
- Apply the same order to upcoming and past cards for guests and signed-in viewers, including expanded summaries.
- Allow long text to wrap within its information row rather than clipping it. Keep classification and staffing warnings visible as supplemental content.
- Preserve the existing staffing area, disclosure controls, block icons and cake configuration states.

## Capabilities

### New Capabilities

None.

### Modified Capabilities

- `schedule-overview`: Replace the shared time/matchup or time/task-label row with ordered information rows and a more prominent start time for games and all day-task cards, while retaining responsive staffing summaries.

## Impact

- Primarily `static/style.css`, using the shared summary structure in `templates/schedule.html` for games and day-task blocks; adjust markup only if needed to keep content in the specified rows.
- Update the overview description in `README.MD` and verify the layout using synthetic schedule data at mobile and desktop widths.
- No changes to database schema, APIs, assignment rules, notifications or dependencies.
