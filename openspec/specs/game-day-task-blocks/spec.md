# Game Day Task Blocks Specification

## Purpose

Defines day-wide preparation and cleanup duties around a home-game date without
misrepresenting those duties as work belonging to one particular game.

## Requirements

### Requirement: Every home-game date has two automatic task blocks

The system SHALL provide one preparation block and one cleanup block for every season
date containing at least one home game. The blocks SHALL have no responsible team and
SHALL not require an administrator to create them.

The preparation block SHALL contain exactly three slots displayed as `Vorbereitung 1`,
`Vorbereitung 2` and `Vorbereitung 3`. The cleanup block SHALL contain exactly three slots
displayed as `Aufräumen 1`, `Aufräumen 2` and `Aufräumen 3`. Their numbers SHALL identify
display positions only. All three preparation slots SHALL use the semantic task role
`Vorbereitung`, and all three cleanup slots SHALL use `Aufräumen`, analogous to the two
numbered display slots of the single `Verkauf` role.

When synchronization leaves a season/date without any home game, the system SHALL remove
both blocks and their current assignments. It SHALL record a system-generated assignment-
removal audit for every occupied slot before deletion and SHALL write an application log
entry for each removed block, including an empty block. Existing audit snapshots SHALL
remain readable after their block is deleted.

#### Scenario: Date with several games

- **WHEN** a season date contains one or more home games
- **THEN** the schedule contains one preparation block before the games
- **AND** it contains one cleanup block after the games
- **AND** each block contains its three defined task slots exactly once

#### Scenario: Boundary game changes within a date

- **WHEN** an earlier or later game is added to a date that already has block assignments
- **THEN** the same two blocks and their assignments remain attached to that date
- **AND** their displayed times are recalculated from the new boundary games

#### Scenario: Last game leaves a date

- **WHEN** synchronization leaves a season/date without any home game
- **THEN** its preparation and cleanup blocks and their current assignments are removed
- **AND** each occupied slot receives a system-generated assignment-removal audit
- **AND** each removed block produces an application log entry
- **AND** the audit snapshots remain readable after deletion

### Requirement: Block times follow the boundary game starts

The preparation block time SHALL be 90 minutes before the earliest valid game start on
the date. The cleanup block time SHALL be 60 minutes after the latest valid game start on
the date. Calculations SHALL handle crossing a calendar-day boundary without changing the
home-game date to which the block belongs.

#### Scenario: Normal preparation and cleanup times

- **WHEN** the first game starts at 10:00 and the last game starts at 18:00
- **THEN** the preparation block shows 08:30
- **AND** the cleanup block shows 19:00

#### Scenario: Offset crosses midnight

- **WHEN** applying an offset crosses midnight
- **THEN** the calculated clock time and adjacent calendar date are represented correctly
- **AND** the block remains grouped with its home-game date

#### Scenario: Boundary time is unavailable

- **WHEN** no game in the relevant boundary position has a parseable start time
- **THEN** the block remains available without an invented clock time
- **AND** schedule rendering and notification processing continue without failure

### Requirement: Block assignments receive day-level reminders

Assigned preparation and cleanup helpers SHALL receive reminders through the ordinary
contact preference rules. The preparation assignees SHALL receive the special early
preparation reminder one week before the game date, using the preparation block time.
The former special reminder SHALL no longer be sent to the first game's `Verkauf`
assignees; those assignees SHALL receive the ordinary per-game reminder flow. Assigned
block helpers SHALL also receive a reminder on the day before their duty.

#### Scenario: Weekly preparation reminder

- **WHEN** the daily job runs one week before a game date with occupied preparation slots
- **THEN** each preparation assignee receives the early preparation reminder
- **AND** the reminder identifies the preparation time

#### Scenario: First-game sale reminders stay game-specific

- **WHEN** the first game has assigned `Verkauf` helpers
- **THEN** those helpers receive the ordinary per-game reminder
- **AND** they do not receive the special preparation reminder merely because their game
  is first

#### Scenario: Day-before block reminder

- **WHEN** the daily job runs one day before a date with occupied block slots
- **THEN** every preparation and cleanup assignee receives a reminder for their block task

#### Scenario: Empty block slot

- **WHEN** a block slot is empty at reminder time
- **THEN** no helper reminder is attempted for that slot

### Requirement: Block duties participate in statistics

Every occupied block slot SHALL count as one duty in the assigned person's season
statistics and SHALL be grouped under its semantic task role. Empty slots of an upcoming
preparation or cleanup block SHALL be aggregated under that role with their missing count,
as for multi-slot `Verkauf`. Block gaps SHALL not produce an MV reminder because blocks
have no responsible team.

#### Scenario: Assigned preparation duty counted

- **WHEN** a person holds the second displayed preparation slot
- **THEN** their season duty total increases by one
- **AND** the role breakdown identifies the task as `Vorbereitung`

#### Scenario: Empty block task reported

- **WHEN** an upcoming preparation or cleanup block has an empty slot
- **THEN** the statistics page lists its semantic task role and missing count for the
  game date
- **AND** no MV is notified about the gap
