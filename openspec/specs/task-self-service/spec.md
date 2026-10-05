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

Every new claim SHALL also satisfy the applicable individual age minimum or collective Verkauf age rule, based on the scheduled game date and current stored data. These restrictions SHALL apply to every tier and entry point without an administrator override. Authorized releases SHALL remain possible under the existing release rules even if an occupant is ineligible or is the only qualifying adult seller.

#### Scenario: Second task in the same game refused

- **WHEN** a claim would give a person a second task in a game they already have a task in
- **THEN** the request is refused with an explanatory message

#### Scenario: Self-service claim by a player of the game

- **WHEN** a member claims an age-eligible slot in a game and their membership set includes the team playing that game
- **THEN** the claim succeeds
- **AND** the interface marks the assignment as one where the person's team plays itself

#### Scenario: Person belongs to playing and responsible teams

- **WHEN** a person's memberships include both the playing team and the responsible team for a game and the age rules permit the claim
- **THEN** the assignment remains permitted
- **AND** the playing-team warning takes precedence

#### Scenario: Self-service claim from outside the responsible team

- **WHEN** an age-eligible member whose memberships include neither the responsible team nor the support team claims a slot in a game that has a responsible team
- **THEN** the claim succeeds
- **AND** the interface marks the assignment as coming from outside that team

#### Scenario: Unapproved or deactivated person is never assignable

- **WHEN** a claim names a person whose registration is not approved, or who has been deactivated
- **THEN** the request is refused

#### Scenario: Authorized caller selects an underage helper

- **WHEN** a member, MV, administrator, or CLI caller claims an individually age-restricted task for a person below its minimum
- **THEN** the claim is refused and no successful assignment audit is added

#### Scenario: Sole adult seller releases their task

- **WHEN** an authorized caller releases the only qualifying adult Verkauf seller before the game
- **THEN** the release succeeds under the existing release and compare-and-swap rules
- **AND** adult coverage becomes outstanding

### Requirement: Task candidate lists prioritize suitable teams

The system SHALL present the people available in each task-assignment selection control as four consecutive, mutually exclusive categories in this order: people whose memberships include the game's responsible team, people whose memberships include the Supporter team, people belonging only to other teams or to no team, and people whose memberships include the team currently playing the game. People SHALL be ordered alphabetically by display name within each category, with a deterministic order for duplicate names.

Membership in the playing team SHALL place a person in the final category even when their memberships also include the responsible or Supporter team. Otherwise responsible-team membership SHALL take precedence over Supporter membership. An absent category, including the responsible-team category when no responsible team is selected, SHALL simply contribute no people. This ordering SHALL only arrange people whom the viewer is already authorized to assign; it SHALL NOT expand their assignment rights. A person already assigned to another task of the same game SHALL be omitted. The currently selected person SHALL remain visible in their own task control. Existing advisory hints for playing-team and outside-team candidates SHALL be retained and SHALL use the complete membership set.

For a free age-restricted game slot, selectable candidates SHALL be limited to claims permitted by the applicable individual or post-claim Verkauf group rule. A younger or unknown-date seller SHALL remain a candidate while the group can remain incomplete or already contains a qualifying adult seller. Candidates SHALL be refreshed when sale-group occupancy or eligibility changes. An existing occupant SHALL remain visible and releasable even when no longer eligible, with an explanatory deficiency rather than silently disappearing.

#### Scenario: Full roster is grouped and sorted

- **WHEN** an administrator opens a task selection for a game with a responsible team
- **THEN** eligible available people appear in responsible-team, Supporter, other-team, and playing-team order according to their complete membership sets
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
- **THEN** the released person is restored to other task selections for that game wherever current age rules permit their claim, in the correct membership category and alphabetical position
- **AND** the person's complete team label and advisory hint remain unchanged

#### Scenario: Limited assignment scope remains limited

- **WHEN** a member or MV opens a task selection whose candidates are restricted by their access tier
- **THEN** only candidates permitted for that viewer and the slot's age rule are selectable
- **AND** those candidates follow the applicable membership category and alphabetical ordering

#### Scenario: Timing candidate has an unknown birth date

- **WHEN** a free Zeitnehmer or Sekretär slot is opened
- **THEN** a person with an unknown birth date is not offered as an eligible candidate
- **AND** the interface explains unresolved eligibility without exposing dates or exact ages

#### Scenario: Adult seller is removed

- **WHEN** the only qualifying adult seller is released while a younger seller remains assigned
- **THEN** candidates for the last free sale slot are limited to people who can restore adult coverage

#### Scenario: Stored occupant becomes ineligible

- **WHEN** a date or birth-date change makes an existing occupant ineligible
- **THEN** the occupant remains selected in their own control
- **AND** an authorized caller can release them

### Requirement: No release cutoff

The system SHALL allow a release at any time before the game, including after
notifications for that game have been sent. A freed required position SHALL be reported
as missing again by the statistics and by the notification that chases open slots.
Releasing a retained removed-duty assignment SHALL not create a required vacancy or
change required-duty progress or MV follow-up.

#### Scenario: Release after the reminder went out

- **WHEN** a member releases a required position after the notification for that game has been sent
- **THEN** the release succeeds
- **AND** the game appears again among the games with missing assignments

#### Scenario: Retained removed duty released after the reminder

- **WHEN** a member releases their retained removed-duty assignment before the game after its reminder has been sent
- **THEN** the release succeeds and appends an audit
- **AND** required-duty progress, missing-duty statistics and MV follow-up are unchanged

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
preparation, cake-delivery or cleanup block. Each operation SHALL carry the occupant the caller expects,
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

A person SHALL hold at most one task in a particular preparation block, cake-delivery block, cleanup block, or
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

#### Scenario: Two cakes in one block

- **WHEN** a person who already holds one cake position is claimed for another position in that same cake block
- **THEN** the second claim is refused

#### Scenario: Cake and game duties on the same date

- **WHEN** a cake volunteer claims a position in another block or game on that date
- **THEN** the claim is permitted if the other container and its rules allow it

### Requirement: Kasse replaces the per-game Unterstützung task

The current per-game role named Unterstützung SHALL be named Kasse. It SHALL remain assignable, auditable, included in assigned-person statistics and included in helper notifications when occupied. It SHALL also carry the additional ordering/security duties of the rejected second Ordner while Ordnungsdienst remains a separate required singleton. Kasse SHALL be offered and required only for adult games. Youth and unknown-category games SHALL not offer new Kasse claims. Retained Kasse assignments SHALL stay visible and releasable under existing rights, remain in personal statistics and helper reminders, and SHALL not affect required progress, open duties or MV follow-up. No optional game duties SHALL be offered.

The migration SHALL preserve every current Unterstützung assignment's game, person and position under Kasse. It SHALL retain historical audit snapshots with the role text recorded at the time. The newly introduced Reinigung positions SHALL not receive those migrated assignments.

#### Scenario: Empty Unterstützung slot

- **WHEN** every required position of an upcoming youth game is occupied and there is no retained Kasse assignment
- **THEN** the game is treated as completely staffed
- **AND** statistics and MV reminders do not report that removed position as missing

#### Scenario: Empty required Kasse position

- **WHEN** an adult game's Kasse position is empty
- **THEN** that position is reported as a missing required duty

#### Scenario: Occupied Unterstützung slot

- **WHEN** a person is assigned to Kasse
- **THEN** the assignment appears on the schedule and in that person's statistics
- **AND** the person receives the applicable per-game reminders

#### Scenario: Existing Unterstützung assignment migrated

- **WHEN** migration encounters a current Unterstützung assignment
- **THEN** it preserves the same game, position and person under Kasse
- **AND** historical audit snapshots retain their recorded role text
- **AND** the new Reinigung positions begin empty

#### Scenario: Role collision is refused

- **WHEN** migration would overwrite a pre-existing Kasse assignment
- **THEN** migration refuses the collision without silently discarding either assignment

#### Scenario: Existing Reinigung assignment migrated

- **WHEN** a recognized older database still contains original Reinigung assignments and is upgraded through the reviewed migration chain
- **THEN** the existing Reinigung-to-Unterstützung revision and the new Unterstützung-to-Kasse revision preserve the same game, position and person under Kasse
- **AND** historical audit snapshots retain their recorded role text
- **AND** the newly introduced Reinigung positions do not inherit those legacy assignments

#### Scenario: Migrated youth Kasse is release-only

- **WHEN** a youth game has a migrated Kasse assignment
- **THEN** its occupant is shown as an existing assignment and may be released according to existing rights
- **AND** no new candidate, replacement claim or empty Kasse field is offered
- **AND** release appends an audit without changing historical snapshots

### Requirement: On-demand candidates preserve assignment selection behavior

Loading candidates after card expansion SHALL preserve the applicable candidate ordering, complete team labels, playing-team and outside-team hints, and the assignment scope of each slot. A person already assigned to another task in the same game or day-task block SHALL be excluded; the current occupant SHALL remain visible in their own slot even if they are no longer an active or age-eligible candidate. Free game slots SHALL offer only claims permitted by the current individual or collective age rules. Successful claims and releases SHALL keep sibling controls and affected Verkauf eligibility consistent without loading unrelated cards. Assignment writes SHALL retain their per-slot compare-and-swap checks and server-side age revalidation. This change SHALL introduce no new age limits for day-task blocks.

#### Scenario: Expand a staffed game

- **WHEN** a viewer opens a game with one or more occupied slots
- **THEN** each occupant remains selected in their own slot
- **AND** no occupant is offered in another slot of that game
- **AND** available age-eligible candidates retain the existing category order and hints

#### Scenario: Expand a staffed day-task block

- **WHEN** a viewer opens a preparation or cleanup block with an occupied slot
- **THEN** that occupant remains selected in their own slot and is absent from its other slots
- **AND** a task in another game or block does not exclude that person

#### Scenario: Change an assignment on an open card

- **WHEN** a claim or release succeeds
- **THEN** the other controls in that same card update their candidate availability and affected sale-group eligibility
- **AND** the stored occupant remains authoritative if a later write detects a stale expectation

#### Scenario: Loaded age eligibility becomes stale

- **WHEN** a caller submits a claim after another change invalidates loaded eligibility
- **THEN** the server refuses the claim using current stored data
- **AND** the interface retains saved occupants and refreshes the affected candidates
