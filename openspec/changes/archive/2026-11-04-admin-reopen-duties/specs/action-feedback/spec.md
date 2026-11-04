# Spec Delta

## ADDED Requirements

### Requirement: Administrator reopening requires confirmation

Reopening a duty SHALL present a German pre-save confirmation with `Bestätigen` and `Abbrechen` that identifies the duty and helper and explains that the helper loses the duty and is notified. Opening, cancelling or dismissing it SHALL NOT write, audit or notify. The confirmation SHALL remain usable with keyboard and without JavaScript.

#### Scenario: Administrator confirms reopening

- **WHEN** an administrator chooses `Auf offen setzen` for an occupied duty, including a locked one
- **THEN** the confirmation names the duty and helper before any write
- **AND** only `Bestätigen` submits the request for server validation

#### Scenario: Administrator cancels reopening

- **WHEN** an administrator cancels, presses Escape or dismisses the confirmation
- **THEN** the helper remains assigned and no audit or notification is produced

#### Scenario: Reopening without JavaScript

- **WHEN** an administrator reopens a duty without JavaScript
- **THEN** the duty form shows the same confirmation before submitting

## MODIFIED Requirements

### Requirement: Contextual validation and persistent status remain available

Field-specific errors, authentication-code instructions, candidate loading/retry/sign-in controls, ongoing account approval status, staffing deficiencies, replacement-request status and binding/locked assignment state SHALL remain beside the relevant content. They SHALL NOT be replaced by transient bubbles. Ordinary form submissions SHALL produce readable HTML feedback or an appropriate HTML error response instead of exposing raw API JSON. Required pre-deletion warning and confirmation SHALL remain separate from the post-deletion result. Server-rendered action feedback and authentication forms SHALL remain usable without JavaScript.

#### Scenario: Authentication field validation

- **WHEN** a login or registration submission contains an invalid contact, birth date, team selection or code
- **THEN** explanatory field errors remain associated with the affected controls
- **AND** the user can correct and resubmit the form without relying on a transient bubble

#### Scenario: Persistent status

- **WHEN** a user verifies a registration contact and still awaits administrative approval
- **THEN** the action can be confirmed through a bubble while the awaiting-approval instructions remain on the page
- **AND** the bubble does not imply the account is already approved

#### Scenario: Browser form refusal

- **WHEN** a normal management form submission is refused
- **THEN** the browser receives readable HTML feedback identifying the refusal
- **AND** API consumers retain their existing structured error contract and applicable error status

#### Scenario: Management field validation

- **WHEN** a person-creation or profile-edit submission has an invalid birth date or contact field
- **THEN** the explanatory validation error is presented beside the affected authorized form control
- **AND** it is not available only as a transient action-result bubble

#### Scenario: JavaScript is unavailable

- **WHEN** a server-submitted action returns feedback with JavaScript unavailable
- **THEN** its escaped feedback is readable on the destination page in a non-obstructing shared presentation within the document flow
- **AND** code requests, code entry and field-error correction remain usable
- **AND** no close button or permanently obstructing overlay is required

#### Scenario: Replacement request remains visible

- **WHEN** a saved assignment is seeking a replacement
- **THEN** the German status `Ersatz gesucht` remains beside that assignment after any result bubble disappears
- **AND** the assigned person's continued responsibility remains clear

#### Scenario: Occupied duty becomes locked

- **WHEN** the day-before boundary locks an occupied duty and closes its replacement request
- **THEN** the persistent display reflects the saved locked state and retained helper
- **AND** no replacement, release or occupant-change control remains available for any tier, apart from the administrator's `Auf offen setzen` control

### Requirement: Assignment confirmation precedes saving

Every user-interface flow that creates a game or day-block assignment, including a takeover and assignments made by an MV or administrator for another person, SHALL present a German pre-save confirmation with `Bestätigen` and `Abbrechen`. The confirmation SHALL identify the duty and intended helper and explain that saving is binding: the helper remains responsible unless an eligible replacement takes over before the day-before lock. It SHALL explain that helpers and MVs cannot release or change occupied duties through the web service after that boundary, and that a late claim of a truly empty duty is immediately locked. Selecting a helper, opening the confirmation or dismissing it SHALL NOT save an assignment, change responsibility, open a replacement request or append a mutation audit. Only explicit confirmation SHALL submit the requested assignment change, subject to authoritative server checks. The confirmation SHALL remain usable with keyboard and without JavaScript. The stated boundary SHALL be 09:00 Europe/Berlin on the day before the duty, even if the reminder job is missed or delivery fails.

#### Scenario: Helper confirms an empty duty

- **WHEN** a helper selects a free offered duty for themselves
- **THEN** they see its binding-responsibility explanation and `Bestätigen` and `Abbrechen` before any write
- **AND** only choosing `Bestätigen` submits the claim for server validation

#### Scenario: Assignment is cancelled or dismissed

- **WHEN** a user selects `Abbrechen` or dismisses the pre-save confirmation
- **THEN** no assignment, replacement status, responsibility or mutation audit changes
- **AND** the interface retains the authoritative saved assignment state

#### Scenario: Administrator confirms for someone else

- **WHEN** an administrator or authorized MV selects another person for an assignment
- **THEN** the same pre-save confirmation identifies that person and the binding duty
- **AND** tier authority does not bypass confirmation or the occupied-duty lock

#### Scenario: Late empty-slot claim is confirmed

- **WHEN** a user starts a permitted claim of a truly empty duty after the day-before boundary
- **THEN** the pre-save confirmation explains that a successful claim is immediately binding and locked
- **AND** cancelling still produces no write or audit

#### Scenario: Confirmation is stale

- **WHEN** the occupant, request, eligibility, authority or cutoff state changes while confirmation is open
- **THEN** confirming does not override current saved state
- **AND** the interface reports the refusal or conflict accurately and refreshes the affected state
