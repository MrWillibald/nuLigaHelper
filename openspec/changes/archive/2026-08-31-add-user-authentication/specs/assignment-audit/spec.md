## Purpose

Keeps a durable, append-only record of who assigned or released whom, for which task of
which game and when, so the club can answer after the fact why a task slot changed hands.

## ADDED Requirements

### Requirement: Every assignment change is recorded

The system SHALL record one entry for every change to a task assignment, whether it was
made by the person themselves, by an MV, by an admin or by an administrative deletion.

#### Scenario: Self-service claim recorded

- **WHEN** a member claims a task slot
- **THEN** an entry is recorded naming that member as both the actor and the affected
  person

#### Scenario: Assignment by someone else recorded

- **WHEN** an MV or an admin assigns or releases a task for another person
- **THEN** an entry is recorded naming the actor and the affected person separately

#### Scenario: Cascading removal recorded

- **WHEN** a person is deleted and their assignments are removed with them
- **THEN** an entry is recorded for each removed assignment

#### Scenario: Refused change is not recorded as a change

- **WHEN** a change is refused, whether for lack of rights, a conflict or a rule violation
- **THEN** no entry claiming the assignment changed is recorded

### Requirement: An entry describes the change completely

Each entry SHALL record the time of the change, the acting person, the tier the actor
acted as, the kind of change, the affected person, and the game, task and slot concerned.

#### Scenario: Entry content

- **WHEN** an assignment change is recorded
- **THEN** the entry states when it happened, who made it, in which tier, what kind of
  change it was, whom it affected, and which task slot of which game was involved

#### Scenario: Timestamps are precise

- **WHEN** entries are recorded
- **THEN** each carries a full date and time
- **AND** entries can be ordered by that time without relying on the German date format
  used for scraped game dates

### Requirement: The record survives the people and games it refers to

Each entry SHALL retain a readable description of the acting person, the affected person
and the game, including the names in force at the time of the change, so that the entry
stays meaningful after those records are renamed or deleted.

#### Scenario: Person deleted afterwards

- **WHEN** a person named in an entry is deleted from the roster
- **THEN** the entry remains and still shows the name that person had at the time

#### Scenario: Person renamed afterwards

- **WHEN** a person named in an entry changes their name
- **THEN** the entry still shows the name in force when the change was made

### Requirement: The record is append-only

The system SHALL NOT offer any way to edit or delete an entry through the web interface
or the management CLI.

#### Scenario: No way to alter history

- **WHEN** any user, including an admin, is signed in
- **THEN** the interface offers no action that changes or removes an existing entry

### Requirement: Admins can review the record

The system SHALL show the recorded entries to admins in the web interface, most recent
first, and SHALL let them narrow the list down to a single game or a single person.

#### Scenario: Admin reviews recent changes

- **WHEN** an admin opens the record
- **THEN** the most recent changes are listed first with time, actor, affected person,
  game and task

#### Scenario: Admin investigates one game

- **WHEN** an admin filters the record by a game
- **THEN** only the changes concerning that game are listed

#### Scenario: Non-admin has no access

- **WHEN** a guest, a member or an MV requests the record
- **THEN** access is refused
