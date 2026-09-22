## Why

Task-assignment dropdowns currently present eligible people without prioritizing the most suitable helpers for a game. Grouping candidates by their relationship to the game makes suggested choices easier to find while keeping all permitted assignments available.

## What Changes

- Order each task-assignment dropdown into four consecutive groups: members of the responsible team, members of the Supporter team, members of other teams, and members of the team currently playing the game.
- Sort people alphabetically by display name within each group.
- Preserve the existing hints for people from other teams and people whose team is playing the game.
- Continue omitting people who already hold another task for the same game.
- Preserve this ordering when a released person is dynamically restored to other task dropdowns.

## Capabilities

### New Capabilities

None.

### Modified Capabilities

- `task-self-service`: Define the ordering and filtering behavior of the task-assignment candidate lists while retaining advisory team warnings.

## Impact

- Affects schedule construction and task dropdown rendering in `webapp.py` and `templates/schedule.html`.
- Affects client-side option restoration after assignment changes in `static/app.js`.
- Adds or updates schedule and assignment UI tests; no database schema, API contract, or dependency changes are expected.
