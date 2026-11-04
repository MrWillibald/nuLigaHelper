# Spec Delta

## ADDED Requirements

### Requirement: Occupied open-date blocks remain until emptied

An existing block with an empty, placeholder or unparseable date SHALL be removed by synchronization while it has no assignment. While occupied it SHALL be retained without failing synchronization, with its helpers still assigned and responsible. It SHALL stay visible to administrators even if no game has that date, and SHALL be removed by the first synchronization after it is empty.

#### Scenario: Empty open-date blocks are removed

- **WHEN** synchronization finds preparation, cake-delivery or cleanup blocks with the date "offen" and no assignments
- **THEN** those blocks are removed
- **AND** synchronization completes and existing audit snapshots stay readable

#### Scenario: Occupied open-date block does not stop synchronization

- **WHEN** synchronization finds an occupied block with the date "offen"
- **THEN** the block, its helpers and any replacement request are retained
- **AND** all other games and blocks are synchronized normally

#### Scenario: Retained block outlives its games

- **WHEN** no game of the season has the date of a retained open-date block
- **THEN** the block remains visible on the schedule
- **AND** an administrator can still reopen its occupied positions

#### Scenario: Emptied block is removed

- **WHEN** the last helper of a retained open-date block has been removed
- **THEN** the next synchronization removes the block
- **AND** a repeated synchronization changes nothing further

### Requirement: Cancelled home-game dates release their block duties

When a synchronization contains no game for a current or future calendar date, the system SHALL remove that date's blocks and their assignments atomically, including locked duties and open replacement requests, and SHALL notify each removed helper after saving. Dates before today SHALL keep occupied blocks without failing synchronization. An empty scrape SHALL fail before any change. Later games on that date SHALL NOT restore removed assignments.

#### Scenario: Last game moves away

- **WHEN** synchronization moves the only game of a future date with occupied blocks to another date
- **THEN** the game is updated and that date's blocks and assignments are removed
- **AND** synchronization completes
- **AND** the helpers are not moved to the blocks of the new date

#### Scenario: Last game vanishes from the plan

- **WHEN** the only game of a future date is missing from the scrape
- **THEN** that date's blocks and assignments are removed
- **AND** the stored game remains and is reported as no longer contained in the plan

#### Scenario: Locked duty on a cancelled date

- **WHEN** a removed date's occupied block duty is locked or seeking a replacement
- **THEN** the assignment is removed and its request ends with it

#### Scenario: Helpers are notified

- **WHEN** block assignments were removed for a cancelled date
- **THEN** each helper receives one German message per removed duty naming the task and date through the ordinary channel preference
- **AND** a helper without contact data is skipped with a warning and the removal stays saved

#### Scenario: Notification stage fails

- **WHEN** the removal has been saved and sending fails
- **THEN** the removal stays saved and the failure is reported by the daily job
- **AND** a later synchronization does not resend the message

#### Scenario: Past date keeps its record

- **WHEN** a date before today has occupied blocks and no game in the scrape
- **THEN** the blocks and assignments are retained and synchronization completes
- **AND** personal statistics are unchanged and no helper is notified

#### Scenario: Empty scrape

- **WHEN** a synchronization receives no games for a season that has stored games
- **THEN** it fails before any game, block or assignment changes and nobody is notified

#### Scenario: Date returns

- **WHEN** a game later receives a date whose blocks were removed
- **THEN** that date gets new empty blocks and the removed assignments stay removed

### Requirement: Retained open-date blocks offer no new work

A retained open-date block SHALL NOT accept new claims or cake configuration, and its free positions SHALL NOT be offered or reported as missing duties. Replacement requests and takeovers of its occupied positions SHALL keep their existing rules, and occupied positions SHALL keep counting in personal statistics.

#### Scenario: Claim of a free position refused

- **WHEN** any tier attempts to claim a free position of a retained open-date block
- **THEN** the claim is refused and no assignment or audit entry is added

#### Scenario: Cake configuration refused

- **WHEN** an administrator attempts to configure a retained open-date cake block
- **THEN** the change is refused and the saved settings remain

#### Scenario: Free positions are not vacancies

- **WHEN** a retained open-date block has free positions
- **THEN** the schedule offers no claim for them
- **AND** statistics report neither missing duties nor needed setup for that block

#### Scenario: Assigned helper seeks a replacement

- **WHEN** the helper of a retained open-date position opens a replacement request and an eligible person takes over
- **THEN** responsibility transfers under the existing atomic takeover rules

## MODIFIED Requirements

### Requirement: Every home-game date has automatic task blocks

The system SHALL provide one preparation block, one cake-delivery block and one cleanup block for every season
calendar date containing at least one home game. A game with an empty, placeholder or otherwise unparseable date SHALL NOT produce blocks. The blocks SHALL have no responsible team and
SHALL not require an administrator to create them.

The preparation block SHALL contain exactly three slots displayed as `Vorbereitung 1`,
`Vorbereitung 2` and `Vorbereitung 3`. The cleanup block SHALL contain exactly three slots
displayed as `Aufräumen 1`, `Aufräumen 2` and `Aufräumen 3`. Their numbers SHALL identify
display positions only. All three preparation slots SHALL use the semantic task role
`Vorbereitung`, and all three cleanup slots SHALL use `Aufräumen`, analogous to the two
numbered display slots of the single `Verkauf` role.

When synchronization would leave a current or future season calendar date without any home game, its preparation, cake-delivery and cleanup blocks SHALL be removed, including occupied ones, under the requirement for cancelled home-game dates. Empty blocks SHALL be removed with an application log entry; existing audit snapshots SHALL remain readable.

#### Scenario: Date with several games

- **WHEN** a season date contains one or more home games
- **THEN** the schedule contains one preparation block before the games
- **AND** it contains one cleanup block after the games
- **AND** each preparation and cleanup block contains its three defined task slots exactly once

#### Scenario: Game has an open date

- **WHEN** synchronization processes games whose date is empty, "offen" or otherwise not a calendar date
- **THEN** no preparation, cake-delivery or cleanup block is created for that value
- **AND** the games themselves stay on the schedule with their game duties

#### Scenario: Open date receives a calendar date

- **WHEN** a game with an open date later receives a valid date
- **THEN** that date has its three blocks under the ordinary rules

#### Scenario: Boundary game changes within a date

- **WHEN** an earlier or later game is added to a date that already has block assignments
- **THEN** the same preparation and cleanup blocks and their assignments remain attached to that date
- **AND** their displayed times are recalculated from the new boundary games

#### Scenario: Last game leaves a date

- **WHEN** synchronization leaves a current or future season calendar date without any home game
- **THEN** its blocks are removed under the requirement for cancelled home-game dates
- **AND** removals retain application logging and readable historical snapshots

#### Scenario: Cake block follows the date

- **WHEN** several games share a season/date and cake-delivery settings or assignments have been saved
- **THEN** exactly one cake-delivery block exists for that date
- **AND** adding or changing a game on the same date preserves that block, its settings and its assignments

#### Scenario: Occupied cancelled date is protected

- **WHEN** a date before today has occupied blocks and its last home game no longer exists on that season/date
- **THEN** the responsible helpers and blocks remain saved
- **AND** only current and future dates are released under the requirement for cancelled home-game dates
