# Game Identity Specification

## Purpose

Defines season-scoped canonical game-number identity, Spielfest aggregation, safe
lifecycle handling and unambiguous administrative selection for task-relevant games.

## Requirements

### Requirement: Canonical game number defines game identity

The system SHALL identify each ordinary game within a season exclusively by its
scraped game number. Changes to nuLiga meeting identifiers, links, matchup text,
scheduling fields, hall or score MUST NOT create another stored game while that
season and game number remain unchanged.

#### Scenario: Source metadata changes

- **WHEN** a known ordinary game is scraped with the same season and game number but a
  changed or missing nuLiga meeting identifier
- **THEN** the existing stored game is updated
- **AND** no additional game is created

#### Scenario: Ordinary game shifts

- **WHEN** a known ordinary game is scraped with the same season and game number but a
  changed date or time
- **THEN** its existing entry, responsible team, assignments and audit references remain
  associated with that game
- **AND** the shift is reported for that game

#### Scenario: Number is reused in another season

- **WHEN** the same ordinary game number occurs in two different seasons
- **THEN** each season retains its own game entry

### Requirement: Spielfest matches collapse into one task-relevant game

The system SHALL recognize SPF age groups case-insensitively and collapse all scraped
SPF matches with the same date and full normalized age group into one task-relevant
Spielfest game in the plan. The individual SPF matchups SHALL NOT create separate task
slots or separate plan entries.

#### Scenario: Several SPF matches form one Spielfest

- **WHEN** a scrape contains multiple SPF matches on one date with the same full age
  group
- **THEN** the plan contains exactly one Spielfest game for that date and age group
- **AND** it has one shared set of helper tasks

#### Scenario: Two Spielfeste use the same match numbers

- **WHEN** SPF match numbers are reused at Spielfeste on different dates
- **THEN** each date produces one distinct Spielfest game
- **AND** neither date produces one entry per individual match

#### Scenario: Different SPF age groups share a date

- **WHEN** two full SPF age groups have matches on the same date
- **THEN** each age group produces its own Spielfest game and task set

### Requirement: Spielfest game number is deterministic

Each collapsed Spielfest SHALL receive a textual pseudo game number derived
deterministically from its date and full normalized age group. The pseudo number SHALL
serve as that Spielfest's season-scoped game identity.

#### Scenario: Repeated scrape of a Spielfest

- **WHEN** the same SPF date and age group are scraped again with changed individual
  match numbers, opponents, scores or ordering
- **THEN** the same Spielfest game is updated
- **AND** its assignments and responsible team remain attached

#### Scenario: Spielfest date changes

- **WHEN** SPF matches for an age group move to a different date
- **THEN** the new date produces a new pseudo game number and a new game identity
- **AND** assignments from the previous identity are not transferred automatically

### Requirement: Spielfest aggregation fails safely on inconsistent scheduling data

Rows combined into one Spielfest SHALL agree on fields that describe the shared event.
The system SHALL reject an inconsistent group before synchronizing any games rather
than selecting arbitrary values.

#### Scenario: SPF group contains conflicting halls

- **WHEN** SPF rows with the same date and age group contain different halls
- **THEN** synchronization is refused with a diagnostic identifying the conflicting
  Spielfest rows
- **AND** no partial game synchronization is committed

#### Scenario: SPF group has several start times

- **WHEN** a valid SPF group contains matches at different times
- **THEN** the collapsed Spielfest uses the earliest match time as its plan time

### Requirement: Existing databases migrate without silent task-data loss

The system SHALL provide a backup-first migration from source-key game identity to
season-scoped canonical game-number identity. The migration SHALL preserve game IDs
where a survivor can be selected safely, retain unaffected relationships and audit
history, and stop for manual reconciliation when merging would be ambiguous.

#### Scenario: Existing SPF rows have no conflicting task data

- **WHEN** several stored SPF rows belong to one date and age group and their responsible
  team and assignments can be combined without conflict
- **THEN** they are migrated to one Spielfest game with a pseudo game number
- **AND** retained assignments and audit references point to that game

#### Scenario: SPF rows contain conflicting task data

- **WHEN** stored SPF rows to be collapsed have incompatible responsible teams,
  duplicate task slots or assignments that violate one-task-per-person
- **THEN** migration stops with a reconciliation report
- **AND** it does not silently discard or overwrite task data

#### Scenario: Ordinary stored rows share a game number

- **WHEN** multiple ordinary stored rows in one season have the same game number
- **THEN** migration stops and identifies those rows for manual survivor selection
- **AND** the canonical uniqueness constraint is not installed over unresolved data

### Requirement: Lifecycle events use canonical local identity

The system SHALL carry the affected local game identity through new-game, shift,
missing-referee and removed-game events. Event handling SHALL NOT resolve an existing
game from mutable source metadata, and an SPF event SHALL refer to its single collapsed
Spielfest game.

#### Scenario: One ordinary game shifts

- **WHEN** a known ordinary game changes date or time
- **THEN** the shift event refers to that exact stored game
- **AND** notifications are sent only to helpers assigned to it

#### Scenario: One ordinary game disappears

- **WHEN** an ordinary game number is absent from a later complete scrape for its season
- **THEN** that game is reported as removed

#### Scenario: Spielfest rows change without changing the aggregate identity

- **WHEN** individual SPF rows change but their date and full age group do not
- **THEN** lifecycle processing continues to refer to the existing collapsed Spielfest
  game

### Requirement: Administrative selection uses internal game identity

Administrative interfaces SHALL mutate an existing game by its stable internal ID.
Lists and filters SHALL show ordinary game numbers with enough scheduling context and
SHALL present a collapsed SPF group as one Spielfest game rather than as its individual
matches.

#### Scenario: CLI selects an ordinary game

- **WHEN** an administrator lists or searches games before changing assignments or the
  responsible team
- **THEN** each ordinary result includes its internal ID, canonical number, date, time,
  age group and matchup
- **AND** the mutation command accepts the selected internal ID

#### Scenario: Administrative picker contains a Spielfest

- **WHEN** an administrator opens a game picker for a date containing SPF matches
- **THEN** it contains one clearly labelled Spielfest option per full SPF age group
- **AND** selecting it manages the shared Spielfest task slots
