## Purpose

Lets helpers sign themselves up for the tasks of a home game and withdraw again, and
defines the claim and release semantics that keep the schedule consistent when several
people are editing the same game at the same time.

## ADDED Requirements

### Requirement: Claiming and releasing a single task slot

The system SHALL provide two operations on one task slot of one game: claim, which puts a
person into a free slot, and release, which empties a slot that person holds. Each
operation SHALL carry the occupant the caller expects the slot to have, and SHALL be
refused when the stored occupant differs.

#### Scenario: Slot claimed

- **WHEN** a caller claims a slot they are entitled to fill and the slot is free
- **THEN** the person is recorded in that slot

#### Scenario: Slot released

- **WHEN** a caller releases a slot held by the person they named
- **THEN** the slot becomes free

#### Scenario: Concurrent claim of the same slot

- **WHEN** two callers claim the same free slot and the second request arrives after the
  first has been stored
- **THEN** the second request is refused with a conflict
- **AND** the response reports the current occupant so the interface can correct itself
- **AND** the first claim remains in place

#### Scenario: Release of a slot someone else now holds

- **WHEN** a caller releases a slot whose stored occupant is not the person they named
- **THEN** the request is refused with a conflict and nothing is changed

#### Scenario: Claim of a slot filled in the meantime

- **WHEN** a caller claims a slot that has been filled since their page was rendered
- **THEN** the request is refused with a conflict and the existing assignment is kept

### Requirement: Members act only on their own assignments

The system SHALL allow a member to claim a free slot for themselves and to release a slot
they hold. It SHALL refuse any attempt by a member to place another person in a slot or to
release a slot held by another person.

#### Scenario: Member claims a task

- **WHEN** a member claims a free slot for themselves
- **THEN** the assignment is recorded

#### Scenario: Member releases their own task

- **WHEN** a member releases a slot they hold
- **THEN** the slot becomes free

#### Scenario: Member tries to assign someone else

- **WHEN** a member attempts to claim a slot on behalf of another person
- **THEN** the request is refused

#### Scenario: Member tries to release another person's task

- **WHEN** a member attempts to release a slot held by another person
- **THEN** the request is refused

### Requirement: MVs staff the games their team is responsible for

The system SHALL allow an MV to claim and release slots for members of their own team, on
games whose responsible team is that same team. Both conditions SHALL hold.

#### Scenario: MV assigns a team member

- **WHEN** an MV claims a slot for a member of their team on a game their team is
  responsible for
- **THEN** the assignment is recorded

#### Scenario: MV and a game owned by another team

- **WHEN** an MV attempts to claim a slot on a game whose responsible team is not theirs
- **THEN** the request is refused

#### Scenario: MV and a person from another team

- **WHEN** an MV attempts to claim a slot for a person who is not a member of their team
- **THEN** the request is refused

### Requirement: Admins assign anyone

The system SHALL allow an admin to claim and release any slot for any person on the
roster, subject to the same assignment rules that apply to everyone else.

#### Scenario: Admin reassigns a task

- **WHEN** an admin releases a slot held by one person and claims it for another
- **THEN** both changes are applied

### Requirement: Existing assignment rules apply to every tier

The system SHALL enforce that a person holds at most one task per game, whoever makes the
change, and SHALL refuse a claim that would give a person a second task in the same game.
The warnings that a person plays in the game themselves, or belongs to neither the
responsible team nor the support team, SHALL remain advisory: they are shown but do not
block the assignment.

#### Scenario: Second task in the same game refused

- **WHEN** a claim would give a person a second task in a game they already have a task in
- **THEN** the request is refused with an explanatory message

#### Scenario: Self-service claim by a player of the game

- **WHEN** a member claims a slot in a game their own team is playing
- **THEN** the claim succeeds
- **AND** the interface marks the assignment as one where the person plays themselves

#### Scenario: Self-service claim from outside the responsible team

- **WHEN** a member of neither the responsible team nor the support team claims a slot in
  a game that has a responsible team
- **THEN** the claim succeeds
- **AND** the interface marks the assignment as coming from outside that team

#### Scenario: Unapproved or deactivated person is never assignable

- **WHEN** a claim names a person whose registration is not approved, or who has been
  deactivated
- **THEN** the request is refused

### Requirement: No release cutoff

The system SHALL allow a release at any time before the game, including after
notifications for that game have been sent. The freed slot SHALL be reported as missing
again by the statistics and by the notification that chases open slots.

#### Scenario: Release after the reminder went out

- **WHEN** a member releases a slot after the notification for that game has been sent
- **THEN** the release succeeds
- **AND** the game appears again among the games with missing assignments

### Requirement: Self-service is limited to games that are still ahead

The system SHALL refuse claims and releases by members and MVs for games whose date has
passed, so the record of who served cannot be rewritten by the people it describes. An
admin SHALL still be able to correct a past game, and such a correction SHALL be recorded
like any other change.

#### Scenario: Member claims a past game

- **WHEN** a member or an MV claims or releases a slot on a game whose date lies in the
  past
- **THEN** the request is refused

#### Scenario: Admin corrects a past game

- **WHEN** an admin changes an assignment on a past game
- **THEN** the change is applied and recorded
