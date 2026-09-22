## MODIFIED Requirements

### Requirement: Authentication forms present one guided two-step flow

The system SHALL present registration and login as matching, responsive two-step forms
that keep requesting and entering a code on the same page. Every input SHALL have a
visible German label and a short description of the expected value or its purpose.
The registration form SHALL show the e-mail and SMS contact fields before the route
selection. The route choices SHALL be available only for contact values that are
present and valid.

#### Scenario: Visitor opens registration

- **WHEN** an unauthenticated visitor opens the registration page
- **THEN** the page shows fields for name, desired team, e-mail address and SMS number
- **AND** the e-mail and SMS fields appear before the contact-route selection
- **AND** the route selection explicitly offers E-Mail and SMS
- **AND** SMS shows a country-calling-code selector and national-number field
- **AND** the page shows the consent control, code-request action, registration-code field
  and final registration action in their intended order

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
- **AND** preserves the chosen contact context
- **AND** allows the visitor to return to and change the step-one data

#### Scenario: Action hierarchy is shown

- **WHEN** either authentication form is displayed
- **THEN** the code-request button uses a less saturated secondary treatment
- **AND** the final registration or login button uses the established saturated primary treatment

#### Scenario: Form is used without client-side scripting

- **WHEN** client-side JavaScript is unavailable
- **THEN** both code request and code confirmation remain usable through server-rendered form submissions
- **AND** server validation still rejects invalid or incomplete contact data

### Requirement: Self-registration with one or more contact routes and a desired team

The system SHALL let an unauthenticated visitor register by supplying a name, a desired
team, and at least one valid contact route from e-mail and phone. The visitor MAY supply
both routes. Every supplied contact SHALL be validated server-side; an invalid supplied
contact SHALL block registration until it is corrected or cleared. The system SHALL obtain
the registrant's consent to publish their name on the public schedule before issuing the
registration code. The selected valid route SHALL receive the verification code, while all
supplied valid contacts SHALL be stored canonically on the pending registration.

#### Scenario: Registration submitted with both contacts

- **WHEN** a visitor submits a valid name, desired team, e-mail address and phone number,
  confirms consent, and selects either E-Mail or SMS
- **THEN** the system records both canonical contacts on a pending registration
- **AND** sends a six-digit verification code only through the selected route
- **AND** the registrant is not yet on the roster

#### Scenario: Registration submitted with one contact

- **WHEN** a visitor submits a valid name, desired team and exactly one valid contact,
  confirms consent, and selects that contact's route
- **THEN** the system records the supplied canonical contact on a pending registration
- **AND** sends a six-digit verification code through that route

#### Scenario: Registration without a contact

- **WHEN** a visitor requests a registration code without supplying either contact
- **THEN** the registration is rejected because a contact is required to prove control of
  the identity

#### Scenario: Registration with an invalid supplied contact

- **WHEN** a visitor submits an invalid e-mail address or phone number in a non-empty
  contact field
- **THEN** the registration is rejected with an explanatory validation error
- **AND** no authentication message is sent
- **AND** no person record is created or changed

#### Scenario: Registration without consent

- **WHEN** a visitor requests a registration code without confirming the consent notice
- **THEN** the registration is rejected with an explanatory message

#### Scenario: Registration for a contact already in use

- **WHEN** a visitor requests registration with either supplied contact already belonging
  to a person on the roster
- **THEN** the system presents exactly the same code-entry state as it does for an unused
  contact
- **AND** no second person is created or existing person modified
- **AND** any account-exists message is sent only through the selected route

### Requirement: Passwordless login by e-mail or SMS code

The system SHALL authenticate a person by proving control of any selected contact route
already stored on their record, using a single-use six-digit numeric code sent by e-mail or
SMS. A person with both stored routes SHALL be able to request login through either route.
The selected route SHALL determine delivery even when the person has both contact
channels. The system SHALL NOT store passwords.

#### Scenario: Login by e-mail code

- **WHEN** a person requests a login code for an e-mail address belonging to an account
  that may log in
- **THEN** a six-digit single-use code is sent to that address
- **AND** entering it within its validity period signs the person in

#### Scenario: Login by SMS code

- **WHEN** a person requests a login code for a phone number belonging to an account that
  may log in
- **THEN** a six-digit single-use code is sent by SMS
- **AND** entering it within its validity period signs the person in

#### Scenario: Person with both contacts can use either route

- **WHEN** an eligible person has both an e-mail address and phone number stored
- **THEN** the person can authenticate by requesting and entering a code through either
  route independently

#### Scenario: Active person completes login

- **WHEN** an active person enters a valid login code
- **THEN** the system creates their session
- **AND** redirects them to the Heimspielplan

#### Scenario: Login code is single use

- **WHEN** a login code that has already been used is presented again
- **THEN** the sign-in is refused and a new one must be requested

#### Scenario: Login code expires

- **WHEN** a login code is presented after its validity period has passed
- **THEN** the sign-in is refused and a new one must be requested

#### Scenario: Person without any contact channel cannot log in

- **WHEN** a login is requested for a person who has neither e-mail nor phone
- **THEN** no message is sent and no session is created
- **AND** the person remains assignable by others

## ADDED Requirements

### Requirement: Automatic notifications prefer e-mail

The system SHALL use a person's stored e-mail address for automatic notifications when
one is available and valid. It SHALL use the stored phone number as fallback only when no
usable e-mail address is available. This preference SHALL apply independently of the
contact route selected for registration or login.

#### Scenario: Person has both contacts for an automatic notification

- **WHEN** an automatic notification is sent to a person with both a valid e-mail address
  and phone number
- **THEN** the notification is sent by e-mail
- **AND** no SMS is sent for that notification

#### Scenario: Person has only phone for an automatic notification

- **WHEN** an automatic notification is sent to a person without a usable e-mail address
  but with a valid phone number
- **THEN** the notification is sent by SMS
