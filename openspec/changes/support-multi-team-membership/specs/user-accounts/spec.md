## MODIFIED Requirements

### Requirement: MVs can create users for managed teams

The system SHALL allow an MV to create an active roster person with an initial membership in any team for which the creator is currently the MV. The MV SHALL be able to choose among all of their managed teams, but SHALL NOT create a person for another team or grant additional memberships during creation. The existing separate e-mail and phone fields, including the ability to leave either or both empty, SHALL remain unchanged. This administrative creation of an active roster record SHALL remain separate from public self-registration, which requires admin approval.

#### Scenario: MV creates a user for a managed team

- **WHEN** an MV submits a valid new-person form with one of their managed teams selected
- **THEN** an active person is created with that team as their initial membership
- **AND** the person appears in the roster and becomes assignable

#### Scenario: MV manages multiple teams

- **WHEN** an MV manages more than one team
- **THEN** the new-person form offers every team they manage
- **AND** the MV may choose any one of those teams as the initial membership

#### Scenario: MV creates a user without contact data

- **WHEN** an MV creates a person without an e-mail address or phone number
- **THEN** the person is created as an active, assignable roster person
- **AND** the person has no way to log in or receive notifications, as for an admin-created contactless person

#### Scenario: MV attempts to create a user for another team

- **WHEN** an MV submits or forges a team choice outside their managed teams
- **THEN** the request is refused
- **AND** no person or membership is created

### Requirement: Person identity is an internal identifier

The system SHALL identify every person by an internal identifier that is stable for the lifetime of the record and is never displayed in the user interface. Names SHALL be mutable display data and SHALL NOT be required to be unique.

#### Scenario: Person renamed without losing history

- **WHEN** a person's name is changed
- **THEN** all existing task assignments, team memberships, account state and audit records for that person remain attached to them
- **AND** the schedule shows the new name

#### Scenario: Two persons share a name

- **WHEN** a second person is created or approved with a name that already exists on the roster
- **THEN** the system accepts it
- **AND** every place that lists or offers a person for selection shows all of the person's teams alongside the name so the two can be told apart

#### Scenario: Name is not an identity key

- **WHEN** a request refers to a person
- **THEN** the reference is the internal identifier
- **AND** a request that identifies a person only by name is rejected

### Requirement: Authentication forms present one guided two-step flow

The system SHALL present registration and login as matching, responsive two-step forms that keep requesting and entering a code on the same page. Every input SHALL have a visible German label and a short description of the expected value or its purpose. The registration form SHALL allow one or more existing teams to be selected and SHALL show the e-mail and SMS contact fields before the route selection. The route choices SHALL be available only for contact values that are present and valid.

#### Scenario: Visitor opens registration

- **WHEN** an unauthenticated visitor opens the registration page
- **THEN** the page shows fields for name, multiple team selection, e-mail address and SMS number
- **AND** the e-mail and SMS fields appear before the contact-route selection
- **AND** the route selection explicitly offers E-Mail and SMS
- **AND** SMS shows a country-calling-code selector and national-number field
- **AND** the page shows the consent control, code-request action, registration-code field and final registration action in their intended order

#### Scenario: Route choices reflect contact validity

- **WHEN** a visitor enters no value or an invalid value for a contact field
- **THEN** the corresponding route cannot be selected
- **AND** the visitor must correct or clear the invalid field before submitting registration

#### Scenario: Visitor opens login

- **WHEN** an unauthenticated visitor opens the login page
- **THEN** the page uses the same layout and contact controls as registration
- **AND** shows the code-request action before the login-code field and final login action

#### Scenario: Code is requested

- **WHEN** a visitor submits valid step-one data to request a code
- **THEN** the same page presents the code-confirmation step
- **AND** preserves the chosen contact and complete team-selection context
- **AND** allows the visitor to return to and change the step-one data

#### Scenario: Action hierarchy is shown

- **WHEN** either authentication form is displayed
- **THEN** the code-request button uses a less saturated secondary treatment
- **AND** the final registration or login button uses the established saturated primary treatment

#### Scenario: Form is used without client-side scripting

- **WHEN** client-side JavaScript is unavailable
- **THEN** both code request and code confirmation remain usable through server-rendered form submissions
- **AND** server validation still rejects invalid or incomplete contact or team-selection data

### Requirement: Self-registration with one or more contact routes and teams

The system SHALL let an unauthenticated visitor register by supplying a name, selecting at least one existing team, and supplying at least one valid contact route from e-mail and phone. The visitor MAY select several teams and MAY supply both contact routes. Every supplied contact and selected team SHALL be validated server-side; an invalid supplied value SHALL block registration until corrected. The system SHALL obtain the registrant's consent to publish their name on the public schedule before issuing the registration code. The selected valid route SHALL receive the verification code, while all supplied valid contacts and the complete team selection SHALL be stored on the pending registration.

#### Scenario: Registration submitted with both contacts

- **WHEN** a visitor submits a valid name, one or more teams, e-mail address and phone number, confirms consent, and selects either E-Mail or SMS
- **THEN** the system records both canonical contacts and every selected team on a pending registration
- **AND** sends a six-digit verification code only through the selected route
- **AND** the registrant is not yet on the roster

#### Scenario: Registration submitted with one contact

- **WHEN** a visitor submits a valid name, one or more teams and exactly one valid contact, confirms consent, and selects that contact's route
- **THEN** the system records the supplied canonical contact and complete team selection on a pending registration
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

### Requirement: Verification proves the channel, approval grants roster membership

The system SHALL treat channel verification and admin approval of the user registration as separate gates. Team selections SHALL require no separate approval, but SHALL remain inactive while the user registration is pending. A verified registrant SHALL be able to log in and see their own registration status, and SHALL NOT appear on the roster, in any person selection, or in any assignment until an administrator approves the registration.

#### Scenario: Channel verified

- **WHEN** the registrant enters the registration code within its validity period
- **THEN** the account becomes able to log in
- **AND** the registration and all selected teams are queued for admin approval as one user decision

#### Scenario: Verified but unapproved person is not assignable

- **WHEN** a registration has been verified but not yet approved
- **THEN** the person and their selected teams do not appear on the roster page, in any person selection, or in the statistics
- **AND** any attempt to assign that person or edit their team memberships is rejected

#### Scenario: Verified but unapproved person signs in

- **WHEN** a verified but unapproved person logs in
- **THEN** they see that their registration is awaiting admin approval
- **AND** they have no more access than a guest to the rest of the interface

### Requirement: Registrations are approved by an admin

The system SHALL make every verified user registration visible for approval to admins only. An admin SHALL approve or reject the person as one decision; approval SHALL activate every freely selected team membership together, without any team or MV confirmation, and rejection SHALL activate none. The approval action SHALL NOT substitute or add teams. MVs SHALL NOT view, approve, or reject pending user registrations.

#### Scenario: MV approves a registration for their own team

- **WHEN** an MV attempts to approve a pending registration that selected their team
- **THEN** the request is refused
- **AND** the registration remains pending for an admin

#### Scenario: MV cannot act on another team's registration

- **WHEN** an MV attempts to approve or reject any pending registration
- **THEN** the request is rejected regardless of the selected teams

#### Scenario: Requested team has no MV

- **WHEN** a registration selects a team with no MV or the Supporter team
- **THEN** the same admin-only approval flow applies
- **AND** no team-specific fallback or exception is needed

#### Scenario: Registration rejected

- **WHEN** an admin rejects a registration
- **THEN** none of the selected memberships becomes active and the person does not join the roster
- **AND** the account cannot be used to log in

#### Scenario: Admin approves all selected teams

- **WHEN** an admin approves a verified registration with several selected teams
- **THEN** the person becomes active, appears on the roster and becomes assignable
- **AND** every selected team membership becomes active in the same committed decision

## RENAMED Requirements

- FROM: `Self-registration with one or more contact routes and a desired team`
- TO: `Self-registration with one or more contact routes and teams`
- FROM: `Registrations are approved by the MV of the requested team or by an admin`
- TO: `Registrations are approved by an admin`
