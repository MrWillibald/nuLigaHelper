# Spec Delta

## RENAMED Requirements

- FROM: `### Requirement: Every home-game date has two automatic task blocks`
- TO: `### Requirement: Every home-game date has automatic task blocks`

## MODIFIED Requirements

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

When synchronization leaves a season/date without any home game, the system SHALL remove
all three blocks and their current assignments. It SHALL record a system-generated assignment-
removal audit for every occupied slot before deletion and SHALL write an application log
entry for each removed block, including an empty block. Existing audit snapshots SHALL
remain readable after their block is deleted.

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
- **THEN** its preparation, cake-delivery and cleanup blocks and their current assignments are removed
- **AND** each occupied slot receives a system-generated assignment-removal audit
- **AND** each removed block produces an application log entry
- **AND** the audit snapshots remain readable after deletion

#### Scenario: Cake block follows the date

- **WHEN** several games share a season/date and cake-delivery settings or assignments have been saved
- **THEN** exactly one cake-delivery block exists for that date
- **AND** adding or changing a game on the same date preserves that block, its settings and its assignments

## ADDED Requirements

### Requirement: Administrators configure the cake delivery

Administrators SHALL be able to set the delivery time and requested cake quantity for each cake block. The quantity SHALL be a nonnegative whole number and SHALL define one volunteer position per requested cake. Zero SHALL mean no cakes are requested. A new block SHALL begin without a guessed time or quantity, SHALL show that setup is needed and SHALL not accept claims until both settings are configured. Other viewers SHALL not change those settings. Invalid time or quantity SHALL be refused without changing saved settings.

Each occupied cake position SHALL represent one cake contribution by its volunteer. Increasing quantity SHALL add empty positions while preserving existing positions. A decrease that would remove an occupied position SHALL be refused until that assignment has been explicitly released; remaining positions SHALL not be silently renumbered. A stale configuration update SHALL be refused with the current settings rather than overwrite a newer edit. Claims SHALL be validated against the currently saved capacity.

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

#### Scenario: Explicit release permits reduction

- **WHEN** an authorized administrator explicitly releases every occupied position that a reduction would remove and then saves the reduction
- **THEN** the reduction succeeds
- **AND** the releases remain in assignment history
- **AND** retained positions keep their occupants and position numbers

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

### Requirement: Cake duties receive day-level reporting and reminders

Every occupied cake position SHALL count as one duty under the semantic cake-delivery role in the volunteer's season statistics. Empty configured cake positions SHALL be reported as missing duties with their count. An unconfigured block SHALL be reported as needing administrator setup without inventing a missing cake count or presenting it as fully staffed. No cake gap SHALL generate an MV reminder because the block has no responsible team.

Assigned cake volunteers SHALL receive ordinary day-level reminders one week before and one day before the game date through established contact preference rules. Each reminder SHALL identify the cake task, game date, saved delivery time and one-cake contribution. Empty positions SHALL not receive helper reminders.

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
