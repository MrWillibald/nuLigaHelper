# Spec Delta

## ADDED Requirements

### Requirement: Postponing a game to a later date releases its duty locks

On every synchronization, the system SHALL release a stored game's saved duty lock when its valid current date has no saved cutoff and a day-before cutoff that is not yet due and later than the lock time. Helpers SHALL stay assigned and responsible under ordinary pre-cutoff binding rules; nothing SHALL be vacated, reassigned, released or reopened. The new date's cutoff SHALL lock the duties again.

#### Scenario: Postponed game after its original cutoff

- **WHEN** a game whose occupied duties were locked at its original day-before cutoff is moved by synchronization to a later date whose cutoff has not yet been reached
- **THEN** those duties are no longer locked
- **AND** each original helper remains assigned and responsible for the same position
- **AND** the helper may open a replacement request and an eligible replacement may take over under the ordinary binding rules

#### Scenario: Request closed by the original lock stays closed

- **WHEN** a duty whose replacement request was closed by the original lock is unlocked by a postponement
- **THEN** no replacement request is reopened automatically
- **AND** the helper must open a new request to seek a replacement

#### Scenario: Late claim on the original date is also released

- **WHEN** a duty claimed after the original cutoff, and therefore immediately locked, belongs to a game postponed to a later date that is not yet due
- **THEN** that duty's lock is released in the same way

#### Scenario: New date cutoff locks again

- **WHEN** 09:00 Europe/Berlin on the day before the new date is reached
- **THEN** the occupied duties lock again under the existing day-before rules

#### Scenario: Postponement to a date already past its cutoff

- **WHEN** a game is moved to a later date whose day-before cutoff is already due or already saved
- **THEN** the occupied duties remain locked

#### Scenario: Game moved earlier or only its time changes

- **WHEN** synchronization moves a locked game to an earlier date or changes only its start time on the same date
- **THEN** the saved locks remain unchanged

#### Scenario: Open date becomes a valid future date

- **WHEN** a game with an open or unparseable date receives a valid date whose cutoff is not yet due and any of its duties still carries a saved lock
- **THEN** those locks are released in the same way

#### Scenario: Game postponed before this release

- **WHEN** the first synchronization after this release processes a stored game that was already moved to a later, not-yet-due date and still carries locks from its original date
- **THEN** those locks are released in the same way with one unlock entry per duty
- **AND** games whose current date is already due keep their locks

#### Scenario: Stored game no longer in the scrape

- **WHEN** a falsely locked stored game of the synchronized season is absent from the current scrape
- **THEN** its locks are evaluated by the same rule using its stored date

#### Scenario: Day blocks are not affected

- **WHEN** a game postponement releases game-duty locks
- **THEN** locks of the preparation, cake-delivery and cleanup blocks of any date remain unchanged

### Requirement: A game with an open or unparseable date has no duty lock

On every synchronization, the system SHALL release saved locks of occupied duties of every stored game whose current date is empty, a placeholder such as "offen" or otherwise not a calendar date, including games that already had such a date before. Helpers SHALL stay assigned and responsible under ordinary binding and age rules; nothing SHALL be vacated or reopened. Once a valid date is known, its cutoff SHALL apply again.

#### Scenario: Postponed without a new date

- **WHEN** synchronization replaces a locked game's date with an open or unparseable value
- **THEN** its occupied duties are no longer locked
- **AND** each original helper remains assigned and responsible for the same position
- **AND** the helper may open a replacement request under the ordinary binding rules

#### Scenario: Game already stored with an open date

- **WHEN** synchronization processes a game whose stored and scraped date are both open or unparseable and one of its duties carries a saved lock
- **THEN** that lock is released
- **AND** a repeated synchronization releases nothing further and records no further entry

#### Scenario: Age rules for an unresolved date still apply

- **WHEN** an unlocked duty of a game with an unparseable date is offered for takeover
- **THEN** the existing age rules for an invalid game date still decide whether the takeover is permitted

#### Scenario: Valid date already due locks again

- **WHEN** a game with an open date later receives a valid date whose day-before cutoff is already due
- **THEN** its occupied duties lock again under the existing day-before rules
