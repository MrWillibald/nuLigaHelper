# Assignment Audit Specification

## Purpose

Keeps a durable, append-only record of who assigned or released whom, for which task of
which game and when, so the club can answer after the fact why a task slot changed hands.

## Requirements

### Requirement: Every assignment change is recorded

The system SHALL record one entry for every change to a game or day-block task
assignment, whether it was made by the person themselves, by an MV where permitted, by
an admin or by an administrative deletion.

#### Scenario: Self-service claim recorded

- **WHEN** a member claims a game or day-block task slot
- **THEN** an entry is recorded naming that member as both the actor and the affected
  person

#### Scenario: Assignment by someone else recorded

- **WHEN** an authorized MV or admin assigns or releases a task for another person
- **THEN** an entry is recorded naming the actor and the affected person separately

#### Scenario: Cascading removal recorded

- **WHEN** a person is deleted and their game or day-block assignments are removed with
  them
- **THEN** an entry is recorded for each removed assignment

#### Scenario: Block removal recorded

- **WHEN** a day block is removed because its season/date no longer contains a game
- **THEN** a system-generated removal entry is recorded for each occupied slot before
  its assignment and block are deleted

#### Scenario: Refused change is not recorded as a change

- **WHEN** a game or day-block change is refused for lack of rights, a conflict, or a
  rule violation
- **THEN** no entry claiming the assignment changed is recorded

### Requirement: An entry describes the change completely

Each entry SHALL record the time of the change, the acting person, the tier the actor
acted as, the kind of change, the affected person, and the target task slot. A game entry
SHALL identify the game, role, and slot. A day-block entry SHALL identify the season,
home-game date, preparation or cleanup phase, displayed task name, and slot.

#### Scenario: Entry content

- **WHEN** a game assignment change is recorded
- **THEN** the entry states when it happened, who made it, in which tier, what kind of
  change it was, whom it affected, and which task slot of which game was involved

#### Scenario: Day-block entry content

- **WHEN** a preparation or cleanup assignment change is recorded
- **THEN** the entry states when it happened, who made it, in which tier, what kind of
  change it was, whom it affected, and which named slot of which dated block was involved

#### Scenario: Timestamps are precise

- **WHEN** entries are recorded
- **THEN** each carries a full date and time
- **AND** entries can be ordered by that time without relying on the German date format
  used for game and block dates

### Requirement: The record survives the people and games it refers to

Each entry SHALL retain a readable description of the acting person, the affected person,
and the game or day block, including the names and labels in force at the time of the
change, so that the entry stays meaningful after those records are renamed or deleted.

#### Scenario: Person deleted afterwards

- **WHEN** a person named in an entry is deleted from the roster
- **THEN** the entry remains and still shows the name that person had at the time

#### Scenario: Person renamed afterwards

- **WHEN** a person named in an entry changes their name
- **THEN** the entry still shows the name in force when the change was made

#### Scenario: Role renamed afterwards

- **WHEN** a task role is renamed after an entry was recorded
- **THEN** the entry retains the role name recorded at the time

#### Scenario: Day block deleted after its games disappear

- **WHEN** a dated block is deleted because its date no longer has any home game
- **THEN** its prior and system-generated removal audit entries remain readable

### Requirement: The record is append-only

The system SHALL NOT offer any way to edit or delete an entry through the web interface
or the management CLI.

#### Scenario: No way to alter history

- **WHEN** any user, including an admin, is signed in
- **THEN** the interface offers no action that changes or removes an existing entry

### Requirement: Admins can review the record

The system SHALL show the recorded entries to admins in the web interface, most recent
first, and SHALL let them narrow the list down to a single game or a single person. The
review interface SHALL use the application's established visual language and SHALL present
entries in a readable form on both desktop and mobile screens. It SHALL remain read-only.

#### Scenario: Admin reviews recent changes

- **WHEN** an admin opens the record
- **THEN** the most recent changes are listed first with time, actor, affected person,
  game and task
- **AND** the page uses the same section, card, filter, button and typography treatment as
  the rest of the web interface

#### Scenario: Admin investigates one game

- **WHEN** an admin filters the record by a game
- **THEN** only the changes concerning that game are listed

#### Scenario: Admin investigates one person

- **WHEN** an admin filters the record by a person
- **THEN** only the changes concerning that person as actor or affected person are listed

#### Scenario: Activity entries are readable on mobile

- **WHEN** an admin opens the record on a narrow screen
- **THEN** each entry is presented as a readable stacked item or card rather than requiring
  horizontal scrolling across the full desktop table
- **AND** all audit fields remain available

#### Scenario: Activity review remains read-only

- **WHEN** an admin reviews the record
- **THEN** no action is offered that edits or deletes an existing entry

#### Scenario: Non-admin has no access

- **WHEN** a guest, a member or an MV requests the record
- **THEN** access is refused
