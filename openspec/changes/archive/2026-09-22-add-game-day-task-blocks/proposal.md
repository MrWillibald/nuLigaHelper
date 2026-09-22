## Why

Preparation before a home-game day and cleanup after it are day-wide duties, but the
current model can assign work only to individual games. Treating those duties as tasks
of the first or last game would give them the wrong identity and make their assignments
unstable when the schedule changes.

## What Changes

- Add one preparation block before the first home game and one cleanup block after the
  last home game of each date.
- Give each block three independently assignable slots. Like the two `Verkauf` slots,
  their numbers distinguish display positions only: all preparation slots have the task
  `Vorbereitung`, and all cleanup slots have the task `Aufräumen`.
- Remove both blocks and their current assignments when their season/date no longer
  contains a game. Record system-generated assignment-removal audits for occupied slots,
  log every removed block, and retain the resulting audit snapshots after deletion.
- Calculate the preparation time as 90 minutes before the first game's start and the
  cleanup time as 60 minutes after the last game's start.
- Allow signed-in people to claim or release their own block task and admins to manage
  any active person's block task. MVs receive no team-scoped authority for these blocks,
  and blocks have no responsible team.
- Keep the one-task limit scoped independently to each block and each game, so a person
  can hold a preparation task, game tasks in different games, and a cleanup task on the
  same date, but never two tasks in one block or game.
- Move the special one-week early preparation reminder from the first game's `Verkauf`
  helpers to the preparation-block assignees. Game-level `Verkauf` remains unchanged and
  uses the ordinary per-game reminder flow.
- Rename the per-game `Reinigung` role to `Unterstützung`. The slot remains assignable
  and notified when occupied, but becomes optional: an empty slot is excluded from open-
  task statistics and MV missing-task reminders.
- Preserve historical audit snapshots while migrating current `Reinigung` assignments
  to `Unterstützung` through an explicit Alembic schema revision.

## Capabilities

### New Capabilities

- `game-day-task-blocks`: Defines automatic day-level preparation and cleanup blocks,
  their timing, slots, lifecycle, notifications, and statistics behavior.

### Modified Capabilities

- `schedule-overview`: Shows day-level blocks around the games of a date and defines how
  they interact with public visibility and schedule filtering.
- `task-self-service`: Extends compare-and-swap self-service to block slots, scopes the
  one-task rule per container, and makes `Unterstützung` an optional game task.
- `assignment-audit`: Records block assignment changes durably while retaining readable
  snapshots independently of games and people.

## Impact

- Database models and an explicit Alembic migration for day blocks, block assignments,
  block audit references, and the current-assignment role rename.
- Assignment helpers, authorization checks, JSON endpoints, schedule view construction,
  JavaScript synchronization, statistics, and notification dispatch.
- Schedule, statistics, audit, notification, database, concurrency, migration, and access-
  control tests.
- README task descriptions and notification configuration templates.
