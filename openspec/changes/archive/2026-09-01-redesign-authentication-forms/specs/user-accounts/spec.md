## ADDED Requirements

### Requirement: Contact addresses are validated and canonicalized

The system SHALL validate every e-mail address and phone number on the server before it
is used for registration, authentication, or a person contact-data write. The system
SHALL store and compare accepted contact addresses in a canonical form. Browser-side
validation MAY provide earlier feedback but SHALL NOT replace server validation.

#### Scenario: Valid e-mail address submitted

- **WHEN** a visitor or member submits a syntactically valid e-mail address
- **THEN** the system accepts it
- **AND** stores or looks it up using the same canonical representation

#### Scenario: Invalid e-mail address submitted

- **WHEN** an e-mail address does not pass server-side syntax validation
- **THEN** the system shows an explanatory validation error
- **AND** sends no authentication message
- **AND** creates or changes no person record

#### Scenario: Valid international phone number submitted

- **WHEN** a visitor selects a country calling code and enters a valid national phone number
- **THEN** the system accepts the number
- **AND** combines and stores it in E.164 form

#### Scenario: Invalid phone number submitted

- **WHEN** the submitted country calling code and national number do not form a possible
  and valid telephone number
- **THEN** the system shows an explanatory validation error
- **AND** sends no authentication message
- **AND** creates or changes no person record

#### Scenario: Existing contact data is edited to an invalid value

- **WHEN** an admin or member submits invalid e-mail or phone contact data while editing
  a person
- **THEN** the write is rejected
- **AND** the previously stored contact data remains unchanged

#### Scenario: Canonically equivalent contact is reused

- **WHEN** a submitted address is equivalent to an existing contact after canonicalization
- **THEN** uniqueness and account lookup treat both representations as the same contact

### Requirement: Authentication forms present one guided two-step flow

The system SHALL present registration and login as matching, responsive two-step forms
that keep requesting and entering a code on the same page. Every input SHALL have a
visible German label and a short description of the expected value or its purpose.

#### Scenario: Visitor opens registration

- **WHEN** an unauthenticated visitor opens the registration page
- **THEN** the page shows fields for name, desired team, contact route and contact address
- **AND** the contact route is an explicit choice between E-Mail and SMS
- **AND** SMS shows a country-calling-code selector and national-number field
- **AND** the page shows the consent control, code-request action, registration-code field
  and final registration action in their intended order

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

## MODIFIED Requirements

### Requirement: Self-registration with a contact channel and a desired team

The system SHALL let an unauthenticated visitor register by supplying a name, exactly one
selected contact channel (e-mail address or phone number) and the team they wish to join.
The system SHALL obtain the registrant's consent to publish their name on the public
schedule before issuing the registration code. Registration SHALL be completed by
entering the code delivered to the selected contact channel.

#### Scenario: Registration submitted

- **WHEN** a visitor submits a valid name, contact channel and desired team, and confirms
  the consent notice
- **THEN** the system records a pending registration
- **AND** sends a six-digit verification code to the supplied channel
- **AND** the registrant is not yet on the roster

#### Scenario: Registration without consent

- **WHEN** a visitor requests a registration code without confirming the consent notice
- **THEN** the registration is rejected with an explanatory message

#### Scenario: Registration without a contact channel

- **WHEN** a visitor requests a registration code without selecting and supplying one
  contact channel
- **THEN** the registration is rejected, because a channel is the only way to prove
  control of the identity

#### Scenario: Registration for a contact already in use

- **WHEN** a visitor requests registration with a contact channel that already belongs
  to a person on the roster
- **THEN** the system presents exactly the same code-entry state as it does for an unused channel
- **AND** no second person is created
- **AND** the message sent to that channel explains that an account already exists

### Requirement: Verification proves the channel, approval grants roster membership

The system SHALL treat channel verification and club membership as separate gates. A
verified registrant SHALL be able to log in and see their own registration status, and
SHALL NOT appear on the roster, in any person selection, or in any assignment until the
registration has been approved.

#### Scenario: Channel verified

- **WHEN** the registrant enters the registration code within its validity period
- **THEN** the account becomes able to log in
- **AND** the registration is queued for approval

#### Scenario: Verified but unapproved person is not assignable

- **WHEN** a registration has been verified but not yet approved
- **THEN** the person does not appear on the roster page, in any person selection, or in
  the statistics
- **AND** any attempt to assign that person to a task is rejected

#### Scenario: Verified but unapproved person signs in

- **WHEN** a verified but unapproved person logs in
- **THEN** they see that their registration is awaiting approval
- **AND** they have no more access than a guest to the rest of the interface

### Requirement: Passwordless login by e-mail or SMS code

The system SHALL authenticate a person by proving control of a selected contact channel
already stored on their record, using a single-use six-digit numeric code sent by e-mail
or SMS. The selected route SHALL determine delivery even when the person has both contact
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

## RENAMED Requirements

- FROM: `Passwordless login by e-mail link or SMS code`
- TO: `Passwordless login by e-mail or SMS code`
