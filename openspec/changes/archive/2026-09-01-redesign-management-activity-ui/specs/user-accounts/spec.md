## ADDED Requirements

### Requirement: MVs can create users for managed teams

The system SHALL allow an MV to create an active roster person for any team for which
that person is currently the MV. The MV SHALL be able to choose among all of their managed
teams, but SHALL NOT create a person for another team. The existing separate e-mail and
phone fields, including the ability to leave either or both empty, SHALL remain unchanged.

#### Scenario: MV creates a user for a managed team

- **WHEN** an MV submits a valid new-person form with one of their managed teams selected
- **THEN** an active person is created in that team
- **AND** the person appears in the roster and becomes assignable

#### Scenario: MV manages multiple teams

- **WHEN** an MV manages more than one team
- **THEN** the new-person form offers every team they manage
- **AND** the MV may choose any one of those teams

#### Scenario: MV creates a user without contact data

- **WHEN** an MV creates a person without an e-mail address or phone number
- **THEN** the person is created as an active, assignable roster person
- **AND** the person has no way to log in or receive notifications, as for an admin-created
  contactless person

#### Scenario: MV attempts to create a user for another team

- **WHEN** an MV submits or forges a team choice outside their managed teams
- **THEN** the request is refused
- **AND** no person is created
