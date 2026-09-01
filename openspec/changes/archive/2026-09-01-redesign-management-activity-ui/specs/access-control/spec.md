## MODIFIED Requirements

### Requirement: Members and MVs may not perform administration

The system SHALL restrict deactivating, reactivating and deleting persons, changing another
person's profile, setting the responsible team of a game, and appointing a team MV to admins.
Members SHALL remain unable to create persons. An MV SHALL be allowed to create an active
person only for a team they manage and SHALL be allowed to approve or reject a verified
registration for a team they manage. All other person-management administration SHALL remain
admin-only.

#### Scenario: Member attempts administration

- **WHEN** a member attempts to create a person, deactivate or delete a person, set a game's
  responsible team or appoint an MV
- **THEN** the request is refused

#### Scenario: MV creates a person for a managed team

- **WHEN** an MV creates a person and selects a team they manage
- **THEN** the request succeeds

#### Scenario: MV attempts to create a person for another team

- **WHEN** an MV attempts to create a person for a team they do not manage
- **THEN** the request is refused

#### Scenario: MV approves a managed-team registration

- **WHEN** an MV approves or rejects a verified registration for a team they manage
- **THEN** the request succeeds

#### Scenario: MV attempts administration

- **WHEN** an MV attempts to deactivate, reactivate or delete a person, change another
  person's profile, set a game's responsible team, or appoint an MV
- **THEN** the request is refused

#### Scenario: Admin retains unrestricted person creation

- **WHEN** an admin creates a person for any existing team
- **THEN** the request succeeds

## ADDED Requirements

### Requirement: Roster supports safe member filtering

The person-management page SHALL allow signed-in users to narrow the displayed roster by
name and team. Admins MAY additionally filter by account status. Filtering SHALL not expose
contact data, internal person identifiers, or roster entries that the viewer is otherwise
not entitled to see.

#### Scenario: Member searches the roster

- **WHEN** a signed-in member enters a name or selects a team filter
- **THEN** only matching visible roster entries are shown
- **AND** the member's own contact data remains available only on their own entry

#### Scenario: Admin filters by account status

- **WHEN** an admin selects an account status filter
- **THEN** the roster shows only entries with that status

#### Scenario: Filter does not bypass visibility

- **WHEN** a viewer submits a filter that could match a hidden or unauthorized record
- **THEN** that record is not included in the response
