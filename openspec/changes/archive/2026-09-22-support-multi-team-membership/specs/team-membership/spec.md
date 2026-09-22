## Purpose

Defines how one person can belong to several automatically managed club teams while preserving unambiguous identity, MV invariants, roster administration, and lifecycle behavior.

## ADDED Requirements

### Requirement: Persons can belong to multiple teams

The system SHALL represent team membership as a set of relationships between a person and existing teams. A person MAY belong to no team, one team, or several teams, and the same person/team membership SHALL NOT occur more than once. Membership SHALL NOT distinguish players, coaches, staff, or other reasons for belonging to a team.

#### Scenario: Person belongs to two teams

- **WHEN** an administrator assigns a person to two existing teams
- **THEN** both memberships are stored for the same person identity
- **AND** neither membership replaces the other

#### Scenario: Duplicate membership is submitted

- **WHEN** a write attempts to add a team that is already among the person's memberships
- **THEN** the resulting membership set contains that team exactly once

#### Scenario: Person has no team

- **WHEN** an administrator removes a person's final membership
- **THEN** the person remains on the roster with no team membership
- **AND** their account, assignments, and audit history remain attached to the same identity

### Requirement: Administrators manage membership sets

The system SHALL allow an administrator to view and replace a person's membership set using existing automatically managed teams. An MV SHALL be able to add an active roster person to, or remove an active roster person from, a team they currently manage, without changing that person's memberships in any other team. Ordinary members SHALL NOT change membership sets. A verified self-registration SHALL retain the registrant's freely selected memberships in an inactive state until an administrator approves the user registration.

#### Scenario: Administrator updates several memberships

- **WHEN** an administrator saves a valid set of teams for a person
- **THEN** memberships in that set are retained or added
- **AND** memberships omitted from that set are removed

#### Scenario: Forged membership change by an MV

- **WHEN** an MV submits a membership change for a team they do not manage or attempts to replace a person's complete membership set
- **THEN** the request is refused
- **AND** none of the person's memberships change

#### Scenario: MV adds a person to a managed team

- **WHEN** an MV adds an active roster person to a team they manage
- **THEN** that membership is added
- **AND** all of the person's other memberships remain unchanged

#### Scenario: MV removes a person from a managed team

- **WHEN** an MV removes an active roster person other than themselves from a team they manage
- **THEN** that membership is removed
- **AND** all of the person's other memberships remain unchanged

#### Scenario: MV attempts to remove their qualifying membership

- **WHEN** an MV attempts to remove themselves from a team for which they are the appointed MV
- **THEN** the request is refused
- **AND** an administrator must change the membership or MV appointment

#### Scenario: MV targets an unapproved registration

- **WHEN** an MV attempts to change the selected teams of a registration that is not active
- **THEN** the request is refused
- **AND** only an administrator may decide the user registration

#### Scenario: Unknown team is submitted

- **WHEN** a membership write names a team that does not exist
- **THEN** the complete membership write is rejected
- **AND** the person's previous memberships remain unchanged

### Requirement: MV appointments require membership

Each team SHALL continue to have at most one appointed MV, and the appointed person SHALL be an active member of that team. Removing the applicable membership or deactivating the person SHALL clear that team's MV appointment atomically.

#### Scenario: Member is appointed MV

- **WHEN** an administrator appoints an active person who belongs to the team
- **THEN** that person becomes the team's MV

#### Scenario: Non-member is proposed as MV

- **WHEN** an administrator attempts to appoint a person who does not belong to the team
- **THEN** the appointment is rejected
- **AND** the existing MV appointment is unchanged

#### Scenario: MV membership is removed

- **WHEN** an administrator removes the membership that qualifies a person as a team's MV
- **THEN** that team is left without an MV in the same committed change

### Requirement: Memberships have deterministic public labels

Every protected roster, task selection, assignment label, and statistic that identifies a person's team context SHALL show all of that person's team names in a deterministic order. A person with no memberships SHALL receive the established no-team label. Contact data and internal identifiers SHALL remain subject to their existing visibility rules.

#### Scenario: Duplicate names have different membership sets

- **WHEN** two people share a display name but have different team memberships
- **THEN** selection and roster views show the complete ordered team labels beside each name
- **AND** the people remain distinguishable without exposing contact data or internal identifiers

#### Scenario: Person belongs to no team

- **WHEN** a roster or assignment view displays a person without memberships
- **THEN** it shows the established no-team label instead of an empty or ambiguous team value

### Requirement: Membership lifecycle preserves person history

Deactivation SHALL retain a person's memberships while making the person inactive and clearing every MV appointment they hold. Reactivation SHALL restore roster use of the retained memberships without restoring released assignments or cleared MV appointments. Deletion of an erroneous person record SHALL remove its membership relationships while preserving readable audit history under the existing deletion rules.

#### Scenario: Person is deactivated and reactivated

- **WHEN** an administrator deactivates and later reactivates a person with several memberships
- **THEN** the same membership set is available after reactivation
- **AND** no previous MV appointment is restored automatically

#### Scenario: Erroneous person is deleted

- **WHEN** an administrator deletes a person record under the existing deletion policy
- **THEN** every membership relationship for that person is removed
- **AND** no team or other person's membership is removed
