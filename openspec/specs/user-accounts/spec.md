# User Accounts Specification

## Purpose

Registration, approval and passwordless login for the people who use the web interface,
so that every write to the schedule can be attributed to a known member of the club
without the system ever storing a password.

## Requirements

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

### Requirement: Person identity is an internal identifier

The system SHALL identify every person by an internal identifier that is stable for the
lifetime of the record and is never displayed in the user interface. Names SHALL be
mutable display data and SHALL NOT be required to be unique.

#### Scenario: Person renamed without losing history

- **WHEN** a person's name is changed
- **THEN** all existing task assignments, team membership, account state and audit
  records for that person remain attached to them
- **AND** the schedule shows the new name

#### Scenario: Two persons share a name

- **WHEN** a second person is created or approved with a name that already exists on the
  roster
- **THEN** the system accepts it
- **AND** every place that lists or offers a person for selection shows the person's team
  alongside the name so the two can be told apart

#### Scenario: Name is not an identity key

- **WHEN** a request refers to a person
- **THEN** the reference is the internal identifier
- **AND** a request that identifies a person only by name is rejected

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

### Requirement: Registrations are approved by the MV of the requested team or by an admin

The system SHALL route a registration to the MV of the team named in the registration.
An MV SHALL be able to approve or reject registrations for their own team only. An admin
SHALL be able to approve or reject any registration. Approval SHALL place the person in
the team named in the registration and SHALL NOT allow the approver to choose a different
team.

#### Scenario: MV approves a registration for their own team

- **WHEN** the MV of the requested team approves a pending registration
- **THEN** the person joins that team, appears on the roster and becomes assignable

#### Scenario: MV cannot act on another team's registration

- **WHEN** an MV attempts to approve or reject a registration for a team they are not MV
  of
- **THEN** the request is rejected

#### Scenario: Requested team has no MV

- **WHEN** a registration names a team that has no MV, or names the support team
- **THEN** only an admin can approve or reject it

#### Scenario: Registration rejected

- **WHEN** an approver rejects a registration
- **THEN** the person does not join the roster
- **AND** the account cannot be used to log in

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

### Requirement: Login and registration do not disclose who is registered

The system SHALL return the same visible response whether or not the supplied contact
channel is known, so the forms cannot be used to test membership.

#### Scenario: Unknown contact submitted

- **WHEN** a login is requested for a contact channel that belongs to nobody
- **THEN** the response is indistinguishable from the response for a known contact
- **AND** no message is sent

### Requirement: Delivery of login and approval messages is rate limited

The system SHALL limit how often login and registration messages can be triggered, per
person and per originating client, and SHALL apply a stricter limit to SMS than to
e-mail because each message has a monetary cost.

#### Scenario: Repeated login requests for the same person

- **WHEN** login messages for the same person are requested more often than the limit
  allows
- **THEN** further requests are refused for a cooling-off period
- **AND** no additional message is sent

#### Scenario: Many login requests from one client

- **WHEN** a single client requests logins for many different contacts in quick
  succession
- **THEN** further requests from that client are refused for a cooling-off period

### Requirement: Sessions expire after one hour of inactivity

The system SHALL end a session after one hour without a request from it, and SHALL extend
the session on each request. A person SHALL be able to end their session explicitly.

#### Scenario: Session kept alive by use

- **WHEN** a signed-in person keeps using the interface
- **THEN** the session remains valid

#### Scenario: Session times out

- **WHEN** a signed-in person makes no request for more than one hour
- **THEN** the next request is treated as unauthenticated

#### Scenario: Logout

- **WHEN** a signed-in person logs out
- **THEN** the session is invalid immediately

### Requirement: Persons without contact data remain fully usable

The system SHALL allow an admin to create and maintain persons that have neither e-mail
nor phone. Such persons SHALL be assignable and SHALL be shown on the schedule like any
other person, SHALL receive no notifications, and SHALL have no way to log in.

#### Scenario: Admin creates a person without contact data

- **WHEN** an admin creates a person with no e-mail and no phone
- **THEN** the person joins the roster and can be assigned to tasks
- **AND** the schedule shows their name for the tasks they hold

#### Scenario: Contactless person is skipped by notifications

- **WHEN** notifications are sent for a game a contactless person is assigned to
- **THEN** that person is skipped and the skip is recorded in the run log

### Requirement: Persons are deactivated rather than deleted

The system SHALL let an admin deactivate a person so that the club's history stays
intact. A deactivated person SHALL keep every past assignment and audit entry, SHALL NOT
appear in any person selection, SHALL NOT be assignable, and SHALL NOT be able to log in.
An admin SHALL be able to reactivate them. Deactivation SHALL apply to persons with and
without contact data alike.

#### Scenario: Admin deactivates a person who left the club

- **WHEN** an admin deactivates a person
- **THEN** the person can no longer log in and is offered in no person selection
- **AND** their assignments on games that have already taken place remain recorded and
  keep showing their name
- **AND** the statistics keep counting those past assignments

#### Scenario: Future assignments are freed

- **WHEN** a person holding tasks on games that have not yet taken place is deactivated
- **THEN** those slots are released so the schedule does not show them as covered
- **AND** each release is recorded as a change made by the deactivating admin

#### Scenario: Deactivated MV is stood down

- **WHEN** the person being deactivated is recorded as the MV of a team
- **THEN** that team is left without an MV
- **AND** the team's registrations fall back to admin approval

#### Scenario: Deactivated person attempts to log in

- **WHEN** a deactivated person requests a login
- **THEN** no session is created
- **AND** the response does not reveal whether the account exists

#### Scenario: Reactivation restores the person

- **WHEN** an admin reactivates a deactivated person
- **THEN** the person appears on the roster again, is assignable again and can log in
  again
- **AND** the assignments freed at deactivation are not restored

#### Scenario: Deactivation is not deletion

- **WHEN** an admin deactivates a person
- **THEN** no audit entry and no past assignment is removed

### Requirement: Deletion stays available for records created in error

The system SHALL keep an admin-only delete action for roster entries that should never
have existed, SHALL warn that it removes the person's assignments, and SHALL point to
deactivation as the way to retire someone who did serve. Audit entries SHALL survive the
deletion.

#### Scenario: Admin deletes a mistaken entry

- **WHEN** an admin deletes a person
- **THEN** the interface first states that assignments will be removed and that
  deactivation preserves them
- **AND** on confirmation the person and their assignments are removed
- **AND** the audit entries naming that person remain readable

#### Scenario: Only admins may delete

- **WHEN** a member or an MV attempts to delete a person
- **THEN** the request is refused

### Requirement: Members maintain their own profile

The system SHALL let a signed-in member change their own name, e-mail address and phone
number, including clearing the contact fields. Clearing every contact channel SHALL leave
the person assignable but unable to log in until an admin restores a channel.

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
