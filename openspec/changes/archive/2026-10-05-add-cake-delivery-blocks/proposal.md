# Proposal

## Why

Cake contributions belong to the entire game day rather than to one game. Administrators need to set when cakes are delivered and how many are needed, with one volunteer position per cake.

## What Changes

- Add one cake-delivery block per season/home-game date alongside preparation and cleanup.
- Let administrators set the delivery time and cake quantity for each date. The confirmed quantity defines the number of independently assignable volunteer positions, one per cake.
- Start new and migrated cake blocks without a guessed time or quantity; show administrator setup is needed before they become claimable.
- Preserve cake assignments and configured delivery details when games change on the same date.
- Prevent quantity reductions from silently discarding occupied positions; those assignments must be explicitly released before their positions are removed.
- Apply existing day-block self-service, privacy, compare-and-swap, reminders, statistics and durable assignment-audit behavior to cake duties.
- Remove a vanished date's cake block through the existing audited block-removal lifecycle.

## Capabilities

### New Capabilities

### Modified Capabilities

- `game-day-task-blocks`: Adds the cake block, administrator delivery settings, variable position count, lifecycle and reporting.
- `task-self-service`: Extends day-block slot operations and per-container assignment limits to cake positions.
- `schedule-overview`: Adds a compact cake card, delivery settings and existing filter/privacy behavior without replacing game-duty requirements.
- `assignment-audit`: Includes cake phases and named cake positions in durable day-block entry descriptions.

## Impact

Day-block models and constraints, a guarded Alembic revision, reconciliation, claim/release validation, administrator endpoints, schedule rendering and JavaScript, notification dispatch, statistics, audit display, migration fixtures and README documentation. Message templates can later adopt the independently proposed central Python message catalog.
