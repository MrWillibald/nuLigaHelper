# Game Day Task Blocks Specification

## Purpose

Defines day-wide preparation and cleanup duties around a home-game date without
misrepresenting those duties as work belonging to one particular game.

## Requirements

### Requirement: Every home-game date has automatic task blocks

The system SHALL provide one preparation block, one cake-delivery block and one cleanup block for every season
date containing at least one home game. The blocks SHALL have no responsible team and
SHALL not require an administrator to create them.

The preparation block SHALL contain exactly three slots displayed as `Vorbereitung 1`,
`Vorbereitung 2` and `Vorbereitung 3`. The cleanup block SHALL contain exactly three slots
displayed as `Aufräumen 1`, `Aufräumen 2` and `Aufräumen 3`. Their numbers SHALL identify
display positions only. All three preparation slots SHALL use the semantic task role
`Vorbereitung`, and all three cleanup slots SHALL use `Aufräumen`, analogous to the two
numbered display slots of the single `Verkauf` role.

When synchronization would leave a season/date without any home game, removal of preparation, cake-delivery and cleanup blocks SHALL be refused atomically if it would cancel a binding assignment. The system SHALL retain the affected games/date/blocks, occupants and replacement state and report the failed synchronization without successful removal audits or notifications. Empty uncommitted blocks MAY be removed under existing rules with an application log entry; existing audit snapshots SHALL remain readable. Administrator or system status SHALL NOT bypass this protection.

#### Scenario: Date with several games

- **WHEN** a season date contains one or more home games
- **THEN** the schedule contains one preparation block before the games
- **AND** it contains one cleanup block after the games
- **AND** each preparation and cleanup block contains its three defined task slots exactly once

#### Scenario: Boundary game changes within a date

- **WHEN** an earlier or later game is added to a date that already has block assignments
- **THEN** the same preparation and cleanup blocks and their assignments remain attached to that date
- **AND** their displayed times are recalculated from the new boundary games

#### Scenario: Last game leaves a date

- **WHEN** synchronization leaves a season/date without any home game
- **THEN** removal is refused if any affected block has a binding commitment
- **AND** the affected games/date/blocks, occupants and replacement state remain saved without partial removal
- **AND** the failure is reported without a successful removal audit or notification
- **AND** empty uncommitted removals retain application logging and readable historical snapshots

#### Scenario: Cake block follows the date

- **WHEN** several games share a season/date and cake-delivery settings or assignments have been saved
- **THEN** exactly one cake-delivery block exists for that date
- **AND** adding or changing a game on the same date preserves that block, its settings and its assignments

#### Scenario: Occupied cancelled date is protected

- **WHEN** synchronization would remove an occupied block because its last home game no longer exists on that season/date
- **THEN** cancellation is refused before destructive changes whether before or after cutoff
- **AND** the responsible helper and scheduled container remain saved with an actionable failure

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
block helpers SHALL also receive a reminder on the day before their duty. Weekly reminders SHALL NOT lock assignments or close replacement requests. At 09:00 Europe/Berlin on the day before the duty, occupied preparation and cleanup duties SHALL be fixed and open replacement requests SHALL close with their responsible helpers retained. This cutoff SHALL apply even when the reminder job is missed, helpers have no contact data or delivery fails. Authoritative reads/writes SHALL enforce the clock without depending on a previously saved batch marker; due lock transitions SHALL persist idempotently under serialization, including before reminder dispatch. Contact preference and skipped-delivery reporting SHALL remain unchanged.

#### Scenario: Weekly preparation reminder

- **WHEN** the daily job runs one week before a game date with occupied preparation slots
- **THEN** each preparation assignee receives the early preparation reminder
- **AND** the reminder identifies the preparation time
- **AND** this weekly reminder does not lock the assignments or close replacement requests

#### Scenario: First-game sale reminders stay game-specific

- **WHEN** the first game has assigned `Verkauf` helpers
- **THEN** those helpers receive the ordinary per-game reminder
- **AND** they do not receive the special preparation reminder merely because their game
  is first

#### Scenario: Day-before block reminder

- **WHEN** the daily job runs one day before a date with occupied block slots
- **THEN** every preparation and cleanup assignee receives a reminder for their block task through the ordinary contact rules
- **AND** occupied block assignments are protected by the common day-before lock and any remaining replacement requests close without vacating their helpers

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

### Requirement: Administrators configure the cake delivery

Administrators SHALL be able to set the delivery time and requested cake quantity for each cake block. The quantity SHALL be a nonnegative whole number and SHALL define one volunteer position per requested cake. Zero SHALL mean no cakes are requested. A new block SHALL begin without a guessed time or quantity, SHALL show that setup is needed and SHALL not accept claims until both settings are configured. Other viewers SHALL not change those settings. Invalid time or quantity SHALL be refused without changing saved settings.

Each occupied cake position SHALL represent one cake contribution by its volunteer. Increasing quantity SHALL add empty positions while preserving existing positions. A decrease that would remove an occupied position SHALL be refused; configuration SHALL NOT release, replace or silently retire its helper. No new force-release or cake-duty retirement path SHALL be added to make such a decrease possible; remaining positions SHALL not be silently renumbered. A stale configuration update SHALL be refused with the current settings rather than overwrite a newer edit. Claims SHALL be validated against the currently saved capacity.

#### Scenario: One volunteer per cake

- **WHEN** an administrator configures four requested cakes and a valid delivery time
- **THEN** the block provides exactly four numbered volunteer positions
- **AND** each occupied position represents one cake

#### Scenario: Delivery time is independent of games

- **WHEN** game start times change after cake delivery has been configured
- **THEN** the administrator's saved delivery time remains unchanged

#### Scenario: Increase requested quantity

- **WHEN** the requested quantity increases from two to four
- **THEN** the first two positions and any occupants remain unchanged
- **AND** two empty positions are added

#### Scenario: Occupied position would be removed

- **WHEN** quantity is reduced below the number of an occupied position
- **THEN** the update is refused with an explanation
- **AND** settings, positions and assignments remain unchanged

#### Scenario: Reduction removes only empty positions

- **WHEN** an administrator reduces quantity and every removed position is genuinely empty
- **THEN** the reduction succeeds subject to current-settings validation
- **AND** retained positions keep their occupants and original position numbers
- **AND** no occupied assignment is released or compacted

#### Scenario: Explicit release permits reduction

- **WHEN** an administrator attempts an ordinary release of an occupied cake position to permit a quantity reduction
- **THEN** the release does not bypass binding protection and the reduction remains refused while the position is occupied
- **AND** no force-release or separate offered-cake retirement operation is available to make the reduction succeed

#### Scenario: Explicitly request no cakes

- **WHEN** an administrator configures zero cakes and no assignments would be removed
- **THEN** the block shows that no cakes are requested
- **AND** it exposes no volunteer positions
- **AND** this state is distinct from unknown quantity

#### Scenario: Configuration update races a claim

- **WHEN** a claim targets a position removed by a committed quantity reduction
- **THEN** the claim is refused without creating an assignment or a successful-change audit

#### Scenario: Stale administrator edit

- **WHEN** an administrator submits settings based on an older saved configuration
- **THEN** the stale update is refused
- **AND** the response identifies the current settings

#### Scenario: Non-admin changes settings

- **WHEN** a member or MV attempts to change delivery time or cake quantity
- **THEN** the update is refused on the server

#### Scenario: Replacement request does not permit reduction

- **WHEN** a quantity decrease would remove a cake position whose helper is seeking a replacement
- **THEN** the decrease is refused because that position remains occupied
- **AND** the helper, responsibility and request remain saved

#### Scenario: Locked cake occupant cannot be removed by configuration

- **WHEN** an administrator decreases quantity to remove an occupied locked cake position
- **THEN** the update is refused without releasing the helper or providing an administrator override

#### Scenario: Empty capacity is added after the boundary

- **WHEN** an administrator increases cake quantity after the day-before boundary without changing an existing occupant
- **THEN** any new configured empty positions follow the ordinary empty-duty claim rules
- **AND** a successful late claim is immediately locked

### Requirement: Cake duties receive day-level reporting and reminders

Every occupied cake position SHALL count as one duty under the semantic cake-delivery role in the volunteer's season statistics. Empty configured cake positions SHALL be reported as missing duties with their count. An unconfigured block SHALL be reported as needing administrator setup without inventing a missing cake count or presenting it as fully staffed. No cake gap SHALL generate an MV reminder because the block has no responsible team.

Assigned cake volunteers SHALL receive ordinary day-level reminders one week before and one day before the game date through established contact preference rules. Each reminder SHALL identify the cake task, game date, saved delivery time and one-cake contribution. Empty positions SHALL not receive helper reminders. A cake position seeking a replacement SHALL remain occupied for reporting, statistics and reminders. Weekly reminders SHALL NOT lock cake assignments. The common cutoff SHALL be 09:00 Europe/Berlin on the day before the duty. It SHALL lock occupied cake duties and close remaining requests with each responsible volunteer retained, even when the reminder job is missed or delivery fails. Authoritative reads/writes SHALL enforce this deadline without waiting for cron, and due dated transitions SHALL persist idempotently under serialization before dispatch. The same contact and delivery rules SHALL remain in force.

#### Scenario: Missing configured cake positions

- **WHEN** two of four configured cake positions are empty
- **THEN** open-duty reporting identifies two missing cake duties
- **AND** no MV reminder is generated for them

#### Scenario: Cake contribution counted

- **WHEN** a person occupies one cake position
- **THEN** their season duty count increases by one under the cake-delivery role

#### Scenario: Cake reminders use saved settings

- **WHEN** the daily reminder job runs one week before or one day before a configured date with cake assignees
- **THEN** each cake assignee receives the applicable reminder with the saved delivery time
- **AND** empty positions receive no helper reminder

#### Scenario: Cake request remains a covered position

- **WHEN** a cake volunteer marks their position `Ersatz gesucht` before the lock
- **THEN** the position still counts as their occupied cake duty and receives ordinary reminders
- **AND** it is not reported as a missing cake

#### Scenario: Cake request is open at the boundary

- **WHEN** the day-before boundary is reached while a cake replacement request is still open
- **THEN** the request closes and the original volunteer remains responsible in the locked position
- **AND** no automatic release, return timer or replacement broadcast occurs

### Requirement: All day-block duties share the binding assignment lifecycle

Preparation, cleanup and every configured cake-delivery position SHALL use the same pre-save binding confirmation, replacement-request, atomic takeover and day-before occupied-duty lock rules as offered game duties. The lock boundary SHALL be 09:00 Europe/Berlin on the day before the duty independently of the reminder job or delivery. The complete dated cohort, including empty positions, SHALL use durable idempotent cutoff state and authoritative clock checks after acquiring write access and after retries; a missing cron marker SHALL NOT postpone locking. The current helper SHALL remain assigned and responsible while `Ersatz gesucht` is open. A valid authorized takeover before the boundary SHALL transfer the occupied position atomically, preserving per-block one-duty rules, current capacity, compare-and-swap conflicts and audit history. No new individual age rule SHALL be introduced for a day block. Block authority SHALL remain team independent: ordinary members and MVs act only for themselves, while an administrator acts within existing person-assignment rights without a lock override. A truly empty configured position SHALL remain claimable after the boundary and SHALL become immediately locked when successfully claimed.

#### Scenario: Preparation helper seeks a replacement

- **WHEN** an assigned helper opens a permitted replacement request for a preparation position before the boundary
- **THEN** the helper stays assigned and responsible while `Ersatz gesucht` is visible
- **AND** no empty position or second preparation duty for that person is created

#### Scenario: Cleanup takeover conflicts

- **WHEN** a cleanup takeover names an occupant or replacement request that has already changed
- **THEN** the takeover is refused with the current authoritative state
- **AND** the stored helper and successful-transfer audit remain unchanged

#### Scenario: Cake duty is genuinely empty after the boundary

- **WHEN** an authorized person confirms a claim of a configured empty cake position after the boundary
- **THEN** the claim is accepted if all current block rules permit it
- **AND** the new responsible helper is immediately locked without an in-service replacement or release flow

#### Scenario: Administrator attempts a locked cleanup change

- **WHEN** an administrator attempts to release, replace or change an occupied cleanup position after the boundary
- **THEN** the web service refuses the change and retains its responsible helper

#### Scenario: Reminder transport fails after batch transition

- **WHEN** the actual day-before batch has committed its dated cutoff and subsequent block reminder delivery fails or is skipped
- **THEN** every committed occupied-duty lock remains in force, including contactless helpers
- **AND** existing failure or skip reporting does not reopen a request or release responsibility
