## MODIFIED Requirements

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
