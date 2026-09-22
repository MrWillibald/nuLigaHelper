# Access Control Specification

## Purpose

Defines who may read and change what in the web interface, so the schedule can be shown
to the public while every modification stays restricted to the people entitled to make
it.

## Requirements

### Requirement: Four access tiers

The system SHALL recognise exactly four tiers — guest, member, MV and admin — and SHALL
derive the tier from stored facts rather than from a role chosen at login. A person is a
member when their account is approved and active, an MV when they are recorded as the
MV of at least one team, and an admin when they are marked as such.

#### Scenario: Tier follows the MV record

- **WHEN** a person is made the MV of a team
- **THEN** their session gains MV rights for that team without any further action

#### Scenario: MV rights removed with the record

- **WHEN** a person stops being the MV of a team
- **THEN** their MV rights for that team end
- **AND** any existing session reflects this on its next request

#### Scenario: Combined tiers

- **WHEN** a person is both an admin and the MV of a team
- **THEN** they hold the union of both sets of rights

### Requirement: The public sees the schedule and nothing else

The system SHALL serve the game schedule, including the names of assigned helpers, to
unauthenticated visitors in read-only form. It SHALL NOT serve the person management page
or the statistics page to them.

#### Scenario: Guest views the schedule

- **WHEN** an unauthenticated visitor opens the schedule
- **THEN** the games, their responsible teams and the names of assigned helpers are shown
- **AND** no control for changing an assignment or a responsible team is offered

#### Scenario: Guest is refused the protected pages

- **WHEN** an unauthenticated visitor requests the person management page or the
  statistics page
- **THEN** access is refused and they are directed to sign in

#### Scenario: Guest page carries no contact data and no roster

- **WHEN** the schedule is rendered for an unauthenticated visitor
- **THEN** the response contains no e-mail address or phone number
- **AND** it contains no list of persons beyond the names actually assigned to the
  displayed games

### Requirement: Authorization is enforced on every request

The system SHALL check the tier and the ownership rules on the server for every request
that reads protected data or writes data, independently of what the rendered page
offered. Absence of a control in the interface SHALL NOT be the only thing preventing an
action.

#### Scenario: Forged request from a lower tier

- **WHEN** a member sends a request that only an admin may make, bypassing the interface
- **THEN** the request is refused

#### Scenario: New endpoint is protected by default

- **WHEN** an endpoint is added without being explicitly declared public
- **THEN** unauthenticated requests to it are refused

### Requirement: Members see the roster without contact data

The system SHALL show a signed-in member the list of persons and every team to which each person belongs, and SHALL show contact data only for the member's own record.

#### Scenario: Member opens the roster

- **WHEN** a member opens the person management page
- **THEN** every visible person's name and complete team membership set is listed
- **AND** e-mail addresses and phone numbers of other persons are not shown

#### Scenario: Member sees own contact data

- **WHEN** a member views their own entry
- **THEN** their own e-mail address and phone number are shown and can be edited
- **AND** their team memberships are visible but cannot be changed by that member

### Requirement: Members and MVs may not perform administration

The system SHALL restrict approving or rejecting user registrations, deactivating, reactivating and deleting persons, changing another person's profile or complete membership set, setting the responsible team of a game, and appointing a team MV to admins. Members SHALL remain unable to create persons or change memberships. An MV SHALL be allowed to create an active person with one initial membership in a team they manage and to add or remove active roster persons for each team they manage. An MV SHALL NOT decide user registrations, change memberships in another team, replace a person's complete membership set, or remove their own membership in a team for which they are the appointed MV. All other person-management administration SHALL remain admin-only.

#### Scenario: Member attempts administration

- **WHEN** a member attempts to create or approve a person, change memberships, deactivate or delete a person, set a game's responsible team or appoint an MV
- **THEN** the request is refused

#### Scenario: MV creates a person for a managed team

- **WHEN** an MV creates a person and selects one team they manage
- **THEN** the request succeeds with that initial membership

#### Scenario: MV attempts to create a person for another team

- **WHEN** an MV attempts to create a person for a team they do not manage
- **THEN** the request is refused

#### Scenario: MV approves a managed-team registration

- **WHEN** an MV attempts to approve or reject a verified registration, including one that selected a team they manage
- **THEN** the request is refused
- **AND** the registration remains pending for an administrator

#### Scenario: MV edits a managed roster

- **WHEN** an MV adds or removes an active person for a team they manage
- **THEN** the request changes only that team membership
- **AND** all other memberships remain unchanged

#### Scenario: MV attempts to edit another roster

- **WHEN** an MV adds or removes a person for a team they do not manage
- **THEN** the request is refused

#### Scenario: MV attempts administration

- **WHEN** an MV attempts to replace an existing person's complete memberships, remove their own qualifying MV membership, deactivate, reactivate or delete a person, change another person's profile, set a game's responsible team, or appoint an MV
- **THEN** the request is refused

#### Scenario: Admin retains unrestricted person creation

- **WHEN** an admin creates a person with a valid set of existing teams
- **THEN** the request succeeds with those initial memberships

### Requirement: Roster supports safe member filtering

The person-management page SHALL allow signed-in users to narrow the displayed roster by name and team membership. A person SHALL match a team filter when that team is anywhere in their membership set. Admins MAY additionally filter by account status. Filtering SHALL not expose contact data, internal person identifiers, or roster entries that the viewer is otherwise not entitled to see.

#### Scenario: Member searches the roster

- **WHEN** a signed-in member enters a name or selects a team filter
- **THEN** only matching visible roster entries are shown, including every person who belongs to the selected team
- **AND** the member's own contact data remains available only on their own entry

#### Scenario: Person has several memberships

- **WHEN** a person belongs to two teams and either team is selected as the roster filter
- **THEN** that person appears in the filtered result

#### Scenario: Admin filters by account status

- **WHEN** an admin selects an account status filter
- **THEN** the roster shows only entries with that status

#### Scenario: Filter does not bypass visibility

- **WHEN** a viewer submits a filter that could match a hidden or unauthorized record
- **THEN** that record is not included in the response

### Requirement: Statistics require a session

The system SHALL make the statistics page available to members, MVs and admins, and SHALL
withhold it from unauthenticated visitors.

#### Scenario: Member opens statistics

- **WHEN** a signed-in member opens the statistics page
- **THEN** the team coverage, the per-person job counts and the list of games with missing
  assignments are shown

### Requirement: An expired session is reported as such

The system SHALL answer a request from an expired or absent session with a distinguishable
authentication failure, and the interface SHALL tell the person their session has ended
and offer to sign in again rather than reporting a generic error.

#### Scenario: Save attempted from a stale page

- **WHEN** a page has been open past the session lifetime and the person changes an
  assignment
- **THEN** the interface reports that the session has expired and prompts for a new sign-in
- **AND** it does not report a connection problem

#### Scenario: Work is not silently lost

- **WHEN** a change is refused because the session expired
- **THEN** the displayed schedule is not updated to suggest the change was saved
