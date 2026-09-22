# Task Self-Service Specification

## Purpose

Lets helpers sign themselves up for the tasks of a home game and withdraw again, and
defines the claim and release semantics that keep the schedule consistent when several
people are editing the same game at the same time.

## Requirements

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

The system SHALL allow an MV to claim and release slots for any active person who belongs to the MV's team, on games whose responsible team is that same team. Both conditions SHALL hold, and membership in additional teams SHALL neither remove nor broaden that team-scoped authority.

#### Scenario: MV assigns a team member

- **WHEN** an MV claims a slot for a person whose membership set includes the responsible team managed by that MV
- **THEN** the assignment is recorded

#### Scenario: MV and a game owned by another team

- **WHEN** an MV attempts to claim a slot on a game whose responsible team is not among the teams they manage
- **THEN** the request is refused

#### Scenario: MV and a person from another team

- **WHEN** an MV attempts to claim a slot for a person whose membership set does not include the game's responsible team
- **THEN** the request is refused

#### Scenario: Person also belongs to another team

- **WHEN** the selected person belongs to the responsible team and one or more additional teams
- **THEN** the additional memberships do not prevent the MV from assigning that person

### Requirement: Admins assign anyone

The system SHALL allow an admin to claim and release any slot for any person on the
roster, subject to the same assignment rules that apply to everyone else.

#### Scenario: Admin reassigns a task

- **WHEN** an admin releases a slot held by one person and claims it for another
- **THEN** both changes are applied

### Requirement: Existing assignment rules apply to every tier

The system SHALL enforce that a person holds at most one task per game, whoever makes the change, and SHALL refuse a claim that would give a person a second task in the same game. Membership in the team playing the game, or membership in neither the responsible team nor the support team, SHALL produce advisory warnings only and SHALL NOT block the assignment. A person SHALL receive the playing-team warning when that team occurs anywhere in their membership set.

#### Scenario: Second task in the same game refused

- **WHEN** a claim would give a person a second task in a game they already have a task in
- **THEN** the request is refused with an explanatory message

#### Scenario: Self-service claim by a player of the game

- **WHEN** a member claims a slot in a game and their membership set includes the team playing that game
- **THEN** the claim succeeds
- **AND** the interface marks the assignment as one where the person's team plays itself

#### Scenario: Person belongs to playing and responsible teams

- **WHEN** a person's memberships include both the playing team and the responsible team for a game
- **THEN** the assignment remains permitted
- **AND** the playing-team warning takes precedence

#### Scenario: Self-service claim from outside the responsible team

- **WHEN** a member whose memberships include neither the responsible team nor the support team claims a slot in a game that has a responsible team
- **THEN** the claim succeeds
- **AND** the interface marks the assignment as coming from outside that team

#### Scenario: Unapproved or deactivated person is never assignable

- **WHEN** a claim names a person whose registration is not approved, or who has been deactivated
- **THEN** the request is refused

### Requirement: Task candidate lists prioritize suitable teams

The system SHALL present the people available in each task-assignment selection control as four consecutive, mutually exclusive categories in this order: people whose memberships include the game's responsible team, people whose memberships include the Supporter team, people belonging only to other teams or to no team, and people whose memberships include the team currently playing the game. People SHALL be ordered alphabetically by display name within each category, with a deterministic order for duplicate names.

Membership in the playing team SHALL place a person in the final category even when their memberships also include the responsible or Supporter team. Otherwise responsible-team membership SHALL take precedence over Supporter membership. An absent category, including the responsible-team category when no responsible team is selected, SHALL simply contribute no people. This ordering SHALL only arrange people whom the viewer is already authorized to assign; it SHALL NOT expand their assignment rights. A person already assigned to another task of the same game SHALL be omitted. The currently selected person SHALL remain visible in their own task control. Existing advisory hints for playing-team and outside-team candidates SHALL be retained and SHALL use the complete membership set.

#### Scenario: Full roster is grouped and sorted

- **WHEN** an administrator opens a task selection for a game with a responsible team
- **THEN** available people appear in responsible-team, Supporter, other-team, and playing-team order according to their complete membership sets
- **AND** people within each category are ordered alphabetically by display name
- **AND** playing-team and outside-team candidates retain their advisory hints

#### Scenario: Membership categories overlap

- **WHEN** a candidate belongs to several teams that would otherwise place them in several categories
- **THEN** the candidate appears exactly once
- **AND** playing-team membership has first precedence, followed by responsible-team membership, Supporter membership, and the other category

#### Scenario: Already assigned person is excluded

- **WHEN** a person holds one task for a game
- **THEN** that person is absent from every other task selection for that game
- **AND** the person remains visible as the selected value for their current task

#### Scenario: Released person returns in the correct position

- **WHEN** a task assignment is released without reloading the schedule
- **THEN** the released person is restored to the other task selections for that game in the correct membership category and alphabetical position
- **AND** the person's complete team label and advisory hint remain unchanged

#### Scenario: Limited assignment scope remains limited

- **WHEN** a member or MV opens a task selection whose candidates are restricted by their access tier
- **THEN** only the candidates already permitted for that viewer appear
- **AND** those candidates follow the applicable membership category and alphabetical ordering

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

### Requirement: Day-block slots use compare-and-swap assignment

The system SHALL provide claim and release operations for one named slot of one
preparation or cleanup block. Each operation SHALL carry the occupant the caller expects,
and SHALL be refused without overwriting stored data when the current occupant differs.

#### Scenario: Concurrent block claim

- **WHEN** two callers claim the same empty block slot and the second request is processed
  after the first claim is stored
- **THEN** the second request is refused with a conflict and the current occupant
- **AND** the first claim remains stored

#### Scenario: Stale block release

- **WHEN** a caller releases a block slot whose occupant differs from the expected person
- **THEN** the release is refused with a conflict
- **AND** the current assignment remains stored

### Requirement: Block assignment authority has no team scope

An active member, including an MV, SHALL be allowed to claim an empty block slot only for
themselves and release only a block slot they hold. An admin SHALL be allowed to claim or
release any block slot for any active person. A block SHALL have no responsible team, and
MV status or team membership SHALL neither broaden nor restrict block assignment rights.
Non-admin changes SHALL be refused after the block's home-game date has passed, while an
admin SHALL remain able to correct past block assignments.

#### Scenario: Member claims own preparation slot

- **WHEN** an active member claims an empty preparation slot for themselves on a current
  or future game date
- **THEN** the assignment is recorded

#### Scenario: MV tries to assign another person

- **WHEN** an MV who is not an admin attempts to claim a block slot for another person
- **THEN** the request is refused regardless of either person's team memberships

#### Scenario: Admin staffs a cleanup block

- **WHEN** an admin claims an empty cleanup slot for an active person
- **THEN** the assignment is recorded without requiring a responsible team

#### Scenario: Member changes a past block

- **WHEN** a member or MV attempts to claim or release a block slot after the block's
  home-game date
- **THEN** the request is refused

### Requirement: One-task limits are scoped per assignment container

A person SHALL hold at most one task in a particular preparation block, cleanup block, or
game. Holding a task in one container SHALL not prevent that person from holding one task
in another container on the same date.

#### Scenario: Two tasks in one preparation block

- **WHEN** a person who already holds one preparation slot is claimed for a second slot
  in the same preparation block
- **THEN** the second claim is refused

#### Scenario: Preparation and game task on the same date

- **WHEN** a preparation assignee is claimed for a task in a game later that date
- **THEN** the game claim is permitted if its slot and other game-level rules allow it

#### Scenario: Preparation and cleanup tasks on the same date

- **WHEN** a preparation assignee is claimed for one cleanup slot on the same date
- **THEN** the cleanup claim is permitted

### Requirement: Unterstützung is an optional per-game task

The per-game role formerly named `Reinigung` SHALL be named `Unterstützung`. It SHALL
remain assignable, auditable, included in assigned-person statistics, and included in
helper notifications when occupied. An empty `Unterstützung` slot SHALL not make a game
incomplete, appear as an open duty, or cause an MV missing-task reminder.

#### Scenario: Empty Unterstützung slot

- **WHEN** every required task of an upcoming game is occupied but `Unterstützung` is
  empty
- **THEN** the game is treated as completely staffed
- **AND** statistics and MV reminders do not report the optional slot as missing

#### Scenario: Occupied Unterstützung slot

- **WHEN** a person is assigned to `Unterstützung`
- **THEN** the assignment appears on the schedule and in that person's statistics
- **AND** the person receives the applicable per-game reminders

#### Scenario: Existing Reinigung assignment migrated

- **WHEN** the schema migration encounters a current `Reinigung` assignment
- **THEN** it preserves the same game, slot, and person under `Unterstützung`
- **AND** historical audit snapshots retain their recorded role text
