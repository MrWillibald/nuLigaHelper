# Action Feedback Specification

## Purpose

Provides consistent, readable confirmation of website actions through shared bubbles at the bottom of the current viewport, with truthful outcomes and distinct success, information, warning and error feedback.

## Requirements

### Requirement: Website action results use shared bottom bubbles

With JavaScript enabled, the website SHALL present action-result feedback through the same bubble component at the bottom center of the current viewport, independent of page scroll position. This SHALL cover game and day-duty claims, releases and replacements; responsible-team and MV changes; cake settings; person creation, profile edits, membership changes, deactivation, reactivation and deletion; registration approval/rejection; authentication-code requests, contact verification, and successful login/logout. Messages SHALL use German wording that identifies the action and its result. A successful release SHALL be distinguishable from a successful assignment. Each logical result SHALL appear once rather than simultaneously as a bubble and a separate transient page banner or inline confirmation. Without JavaScript, server-submitted outcomes SHALL use a readable, non-obstructing shared fallback.

#### Scenario: Assignment and release

- **WHEN** a game or day-duty assignment or release is confirmed
- **THEN** the bottom bubble identifies the assignment or release as successful
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

### Requirement: Warnings are orange and errors are red

Feedback SHALL distinguish success, information, warning and error through visible text or symbols as well as color. Warning feedback SHALL use orange styling, and red notification styling SHALL be reserved for errors. A successful operation with an advisory SHALL explicitly confirm completion and explain the advisory in the same warning notification. Advisories SHALL NOT imply that a saved action was refused. Neutral information and uncomplicated success SHALL use the site's existing compatible visual treatment without error styling.

#### Scenario: Saved assignment with a membership advisory

- **WHEN** an assignment succeeds with an advisory that the person plays in the game or is outside the responsible team
- **THEN** an orange bubble states that the assignment was saved and explains the advisory
- **AND** the saved assignment remains visible

#### Scenario: Last contact is removed

- **WHEN** a permitted profile edit saves a person without a contact route
- **THEN** orange feedback confirms the save and warns about the loss of future login through a contact route
- **AND** the notification does not label the successful save as an error

#### Scenario: A write is refused

- **WHEN** an action is refused by validation, authorization, age eligibility or a stale expectation
- **THEN** its result feedback uses red error styling and an explanatory message
- **AND** it does not present an uncommitted change as successful

### Requirement: Navigation delivers feedback once with a complete reading period

An action outcome SHALL remain available across any same-site redirect or reload required to finish that action. The destination page SHALL display it once with its full reading period beginning when it becomes visible there. Further reloads, unrelated navigation, a later login or a different account session SHALL NOT replay consumed or obsolete feedback. Routine bubbles SHALL NOT be erased before users can read them merely because an action requires navigation.

#### Scenario: Responsible team or MV save reloads

- **WHEN** a successful responsible-team or MV change requires a page reload
- **THEN** the reloaded page displays the corresponding confirmation with a complete visible lifetime
- **AND** a subsequent manual reload does not replay it

#### Scenario: A conflict requires refreshed state

- **WHEN** an assignment or cake-settings conflict refreshes the displayed state or reloads the page
- **THEN** the error feedback remains readable after reconciliation
- **AND** the interface displays the current saved state without reporting the refused change as successful

#### Scenario: Session expiry

- **WHEN** any authenticated action is refused because the session has expired
- **THEN** the shared feedback explicitly identifies session expiry and offers a route to sign in again
- **AND** any automatic navigation to sign-in preserves that explanation on the destination page

### Requirement: Notifications describe confirmed and partial outcomes accurately

Success feedback SHALL depend on a confirmed result, and advisory feedback SHALL use the authoritative saved result where available. Feedback for an interrupted multi-step replacement SHALL distinguish steps already confirmed from steps refused or not confirmed. Loss of a response SHALL NOT be described as proof that no write occurred. A failure to refresh candidates after a confirmed save SHALL NOT reclassify the saved action as failed or undo its displayed saved state.

#### Scenario: Release succeeds and replacement claim fails

- **WHEN** replacing a helper confirms release of the previous helper but refuses the new claim
- **THEN** the error bubble explains that the previous assignment was released and the new assignment could not be made
- **AND** the interface retains the confirmed release and reconciles any newer saved occupancy

#### Scenario: Response is lost

- **WHEN** an action response cannot be received or interpreted sufficiently to confirm the result
- **THEN** the feedback explains that the result could not be confirmed and directs the user to refresh or check the saved state
- **AND** it does not assert that the database remained unchanged

#### Scenario: Candidate refresh fails after save

- **WHEN** an assignment is confirmed but its subsequent candidate refresh fails
- **THEN** the bubble still reports the saved assignment accurately
- **AND** local candidate recovery controls remain available without offering incomplete candidate choices

### Requirement: Contextual validation and persistent status remain available

Field-specific errors, authentication-code instructions, candidate loading/retry/sign-in controls, ongoing account approval status and staffing deficiencies SHALL remain beside the relevant content. They SHALL NOT be replaced by transient bubbles. Ordinary form submissions SHALL produce readable HTML feedback or an appropriate HTML error response instead of exposing raw API JSON. Required pre-deletion warning and confirmation SHALL remain separate from the post-deletion result. Server-rendered action feedback and authentication forms SHALL remain usable without JavaScript.

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

### Requirement: Bubble lifetime and presentation are accessible

With JavaScript enabled, success, information, warning and error bubbles SHALL disappear automatically after five seconds of unpaused visible reading time. Their timers SHALL pause while hovered, containing keyboard focus, outside the stack's visible area, or while the document is hidden. Bubbles SHALL NOT provide a close button. New outcomes SHALL NOT silently replace unread messages; multiple outcomes SHALL be managed within a bounded viewport area, and earlier unread messages SHALL NOT prevent a new result from being seen. Notifications SHALL be announced to assistive technology without automatically moving focus, and error feedback SHALL be distinguishable from ordinary status feedback. Bubbles SHALL wrap long text within narrow screens, respect the visible viewport and safe-area space, and honor reduced-motion preferences. Temporary task-help overlays SHALL NOT hide a newly displayed action outcome, and displaying it SHALL preserve focus within guidance the user is currently reading.

#### Scenario: Uniform automatic lifetime

- **WHEN** a success, information, warning or error bubble becomes visible
- **THEN** it remains for five seconds of unpaused visible time
- **AND** it disappears automatically when that reading period has elapsed
- **AND** it has no close button

#### Scenario: Reading is paused

- **WHEN** a bubble is hovered, contains keyboard focus, is outside the stack's visible area, or the document is hidden
- **THEN** its reading timer is paused
- **AND** returning to unpaused visible reading resumes its remaining time

#### Scenario: Multiple action results

- **WHEN** several action results arrive before earlier feedback expires
- **THEN** unread outcomes are retained and presented without silently replacing one another
- **AND** the feedback region remains bounded within the current viewport

#### Scenario: A new success follows several unread errors

- **WHEN** multiple unread errors already occupy the feedback region and a later action succeeds
- **THEN** the new success confirmation becomes visible immediately
- **AND** the earlier unread errors remain accessible until their visible reading periods expire or deliberate navigation

#### Scenario: A result arrives while task help has keyboard focus

- **WHEN** an asynchronous action result arrives while the user has focus within a task description
- **THEN** its feedback is visible and announced without closing the focused guidance or moving the user's focus

#### Scenario: Narrow viewport and assistive technology

- **WHEN** an action result appears on a narrow or zoomed view, including with an on-screen keyboard
- **THEN** its text and any recovery link fit the visible viewport and remain readable
- **AND** it is announced without moving focus from the user's current control
- **AND** reduced-motion preferences suppress unnecessary slide/fade animation

### Requirement: Feedback preserves privacy and authentication response shape

Feedback SHALL render escaped plain text and SHALL NOT expose additional contact details, full birth dates, exact personal ages, authentication codes, signed challenges or session/security tokens. Code-request feedback SHALL preserve the same generic wording and appearance for known, unknown, ineligible and throttled contacts. Navigation transfer SHALL retain only the minimum short-lived feedback necessary for the immediate destination, rather than creating durable notification history or retaining form/API payloads.

#### Scenario: Generic authentication-code request

- **WHEN** a login or registration code is requested for a known, unknown, ineligible or throttled contact
- **THEN** the same generic informational bubble is shown for equivalent requests
- **AND** it does not disclose account existence or guarantee actual message delivery

#### Scenario: Safe feedback text and transfer

- **WHEN** feedback includes a permitted display name containing markup-like characters or crosses an internal reload
- **THEN** it is displayed as plain text rather than interpreted markup
- **AND** transferred data contains only the minimum feedback and lifecycle information, without copied form or API payloads
