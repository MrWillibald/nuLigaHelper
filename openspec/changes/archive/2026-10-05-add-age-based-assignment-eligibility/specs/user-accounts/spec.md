# Spec Delta

## MODIFIED Requirements

### Requirement: Self-registration with one or more contact routes and teams

The system SHALL let an unauthenticated visitor register by supplying a name and valid birth date, selecting at least one existing team, and supplying at least one valid contact route from e-mail and phone. The visitor MAY select several teams and MAY supply both contact routes. Every supplied contact, birth date, and selected team SHALL be validated server-side; an invalid supplied value SHALL block registration until corrected. The system SHALL obtain the registrant's consent to publish their name on the public schedule before issuing the registration code. The selected valid route SHALL receive the verification code, while the birth date, all supplied valid contacts, and the complete team selection SHALL be stored on the pending registration. The birth date SHALL be preserved across the two-step verification flow.

#### Scenario: Registration submitted with both contacts

- **WHEN** a visitor submits a valid name, birth date, one or more teams, e-mail address and phone number, confirms consent, and selects either E-Mail or SMS
- **THEN** the system records the birth date, both canonical contacts, and every selected team on a pending registration
- **AND** sends a six-digit verification code only through the selected route
- **AND** the registrant is not yet on the roster

#### Scenario: Registration submitted with one contact

- **WHEN** a visitor submits a valid name, birth date, one or more teams and exactly one valid contact, confirms consent, and selects that contact's route
- **THEN** the system records the birth date, supplied canonical contact, and complete team selection on a pending registration
- **AND** sends a six-digit verification code through that route

#### Scenario: Registration without a contact

- **WHEN** a visitor requests a registration code without supplying either contact
- **THEN** the registration is rejected because a contact is required to prove control of the identity

#### Scenario: Registration with an invalid supplied contact

- **WHEN** a visitor submits an invalid e-mail address or phone number in a non-empty contact field
- **THEN** the registration is rejected with an explanatory validation error
- **AND** no authentication message is sent
- **AND** no person record is created or changed

#### Scenario: Registration without consent

- **WHEN** a visitor requests a registration code without confirming the consent notice
- **THEN** the registration is rejected with an explanatory message

#### Scenario: Registration for a contact already in use

- **WHEN** a visitor requests registration with either supplied contact already belonging to a person on the roster
- **THEN** the system presents exactly the same code-entry state as it does for an unused contact
- **AND** no second person is created or existing person modified
- **AND** any account-exists message is sent only through the selected route

#### Scenario: Registration selects several teams

- **WHEN** a visitor selects several valid teams
- **THEN** the pending registration retains every selected team without seeking approval from those teams' MVs

#### Scenario: Registration selects no team

- **WHEN** a visitor requests a registration code without selecting any team
- **THEN** the registration is rejected with an explanatory validation error
- **AND** no person record is created or changed

#### Scenario: Registration birth date is missing or invalid

- **WHEN** a visitor requests a registration code without a valid birth date
- **THEN** the request is refused before creating a person or sending a code
- **AND** the form explains the required correction

#### Scenario: Birth date survives code verification

- **WHEN** the registrant completes the verification step
- **THEN** the pending person retains the validated birth date from the registration request

### Requirement: Members maintain their own profile

The system SHALL let a signed-in member change their own name, e-mail address, phone number, and birth date, including clearing the contact fields. Birth-date changes SHALL follow shared calendar validation and SHALL allow completion of an unknown legacy date but not clearing a known date. Clearing every contact channel SHALL leave the person assignable subject to normal task eligibility, but unable to log in until an admin restores a channel.

#### Scenario: Member updates their own data

- **WHEN** a member changes their own name or contact data
- **THEN** the change is saved and used for future notifications

#### Scenario: Member cannot edit another person

- **WHEN** a member attempts to change another person's data
- **THEN** the request is rejected

#### Scenario: Member removes their last contact channel

- **WHEN** a member clears both e-mail and phone
- **THEN** the change is saved with a warning that they will no longer be able to log in
- **AND** their existing assignments are unaffected

#### Scenario: Member completes a missing birth date

- **WHEN** a member supplies a valid birth date for their own legacy record
- **THEN** the date is stored for that same person
- **AND** affected current/future assignment eligibility is reevaluated

#### Scenario: Member corrects a known birth date

- **WHEN** a member supplies a valid correction to their own birth date
- **THEN** the correction is stored and affected eligibility is reevaluated without deleting assignments

#### Scenario: Known birth date is cleared

- **WHEN** a member submits a blank replacement for an already known birth date
- **THEN** the correction is refused and the stored date remains unchanged

## ADDED Requirements

### Requirement: Administrative person creation collects birth dates

Admin and MV person-creation forms and CLI creation SHALL require a validated birth date, including when both contact fields are empty. Existing creation authority, managed-team restrictions, and contactless-account behavior SHALL remain unchanged. Administrators SHALL be able to complete and correct existing persons' birth dates through authorized person maintenance.

#### Scenario: Contactless person is created

- **WHEN** an authorized creator supplies a name and valid birth date but neither contact channel
- **THEN** the person is created without a login or notification channel
- **AND** assignments remain subject to the applicable age rules

#### Scenario: MV creates a managed-team person

- **WHEN** an MV supplies a valid birth date and otherwise valid creation data for a team they manage
- **THEN** the person is created within the existing one-managed-team scope
- **AND** MV creation does not grant permission to edit that person's profile later

#### Scenario: CLI creation lacks a valid birth date

- **WHEN** CLI creation omits or supplies an invalid birth date
- **THEN** creation fails without inserting a partial person record

#### Scenario: Administrator completes legacy data

- **WHEN** an administrator supplies a valid birth date for an existing person
- **THEN** that date is attached to the existing identity
- **AND** current/future assignment eligibility is reevaluated
