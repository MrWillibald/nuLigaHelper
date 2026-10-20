# Spec Delta

## MODIFIED Requirements

### Requirement: Website action results use shared bottom bubbles

With JavaScript enabled, the website SHALL present action-result feedback through the same bubble component at the bottom center of the current viewport, independent of page scroll position. This SHALL cover game and day-duty claims, replacement-request changes and atomic takeovers, and permitted pre-cutoff obsolete-duty retirement; responsible-team and MV changes; cake settings; person creation, profile edits, membership changes, deactivation, reactivation and deletion; registration approval/rejection; authentication-code requests, contact verification, and successful login/logout. Messages SHALL use German wording that identifies the action and its result. A successful replacement request SHALL explicitly state that the current helper remains assigned and responsible. A successful takeover SHALL be distinguishable from opening a request or claiming an empty position. Feedback SHALL NOT describe a replacement request as a completed release. Each logical result SHALL appear once rather than simultaneously as a bubble and a separate transient page banner or inline confirmation. Without JavaScript, server-submitted outcomes SHALL use a readable, non-obstructing shared fallback.

#### Scenario: Assignment and release

- **WHEN** a game or day-duty claim is confirmed as saved, or an ordinary release is attempted
- **THEN** the bottom bubble identifies the confirmed assignment or refused release accurately
- **AND** the affected card continues to display the saved occupancy and progress

#### Scenario: Helper maintenance

- **WHEN** person creation, profile editing, membership maintenance, deactivation, reactivation or deletion succeeds
- **THEN** the destination management page shows the shared bottom bubble identifying the completed action
- **AND** it does not show a duplicate transient banner at the page top

#### Scenario: Cake and team settings

- **WHEN** cake quantity/time, a responsible team or a team MV is saved
- **THEN** the result appears through the shared bottom bubble
- **AND** the relevant saved settings remain visible in their page context

#### Scenario: Registration decision and authentication

- **WHEN** an administrator approves or rejects a registration, or a user successfully signs in or signs out
- **THEN** the destination page explicitly confirms that action through the shared bubble
- **AND** successfully rejecting a registration is presented as a completed action rather than an error

### Requirement: Notifications describe confirmed and partial outcomes accurately

Success feedback SHALL depend on a confirmed result, and advisory feedback SHALL use the authoritative saved result where available. Responsibility takeover SHALL be one atomic outcome. Feedback SHALL NOT present an independently successful release as part of an ordinary takeover or imply that a refused takeover made the original duty empty. Loss of a response SHALL NOT be described as proof that no write occurred. A failure to refresh candidates after a confirmed save SHALL NOT reclassify the saved action as failed or undo its displayed saved state.

#### Scenario: Atomic takeover fails

- **WHEN** a takeover is refused for authorization, eligibility, stale state or the occupied-duty lock
- **THEN** the error bubble explains that the takeover was not saved
- **AND** the interface retains or reconciles the authoritative original occupant and replacement-request state without displaying an intermediate vacancy

#### Scenario: Release succeeds and replacement claim fails

- **WHEN** a caller attempts the former release-then-claim replacement sequence
- **THEN** no independently successful ordinary release is permitted
- **AND** feedback cannot report a partial replacement that freed the original duty
- **AND** only an atomic successful takeover changes the responsible helper

#### Scenario: Response is lost

- **WHEN** an action response cannot be received or interpreted sufficiently to confirm the result
- **THEN** the feedback explains that the result could not be confirmed and directs the user to refresh or check the saved state
- **AND** it does not assert that the database remained unchanged

#### Scenario: Candidate refresh fails after save

- **WHEN** an assignment is confirmed but its subsequent candidate refresh fails
- **THEN** the bubble still reports the saved assignment accurately
- **AND** local candidate recovery controls remain available without offering incomplete candidate choices

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
- **AND** no replacement, release or occupant-change control remains available for any tier

## ADDED Requirements

### Requirement: Assignment confirmation precedes saving

Every user-interface flow that creates a game or day-block assignment, including a takeover and assignments made by an MV or administrator for another person, SHALL present a German pre-save confirmation with `Bestätigen` and `Abbrechen`. The confirmation SHALL identify the duty and intended helper and explain that saving is binding: the helper remains responsible unless an eligible replacement takes over before the day-before lock. It SHALL explain that occupied duties cannot be released or changed through the web service after that boundary, and that a late claim of a truly empty duty is immediately locked. Selecting a helper, opening the confirmation or dismissing it SHALL NOT save an assignment, change responsibility, open a replacement request or append a mutation audit. Only explicit confirmation SHALL submit the requested assignment change, subject to authoritative server checks. The confirmation SHALL remain usable with keyboard and without JavaScript. The stated boundary SHALL be 09:00 Europe/Berlin on the day before the duty, even if the reminder job is missed or delivery fails.

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
