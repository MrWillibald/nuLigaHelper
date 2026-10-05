# Spec Delta

## MODIFIED Requirements

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
