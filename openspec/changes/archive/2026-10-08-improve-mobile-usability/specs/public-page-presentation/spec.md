# Spec Delta

## ADDED Requirements

### Requirement: Shared pages offer return-to-top navigation

Pages using the shared layout SHALL provide an upward-arrow action with the German accessible name "Nach oben". With JavaScript available, the action SHALL appear at the lower right after scrolling at least one viewport height and SHALL disappear near the top. Activating it SHALL return the viewer to the page top without submitting a form or changing data. It SHALL support keyboard and touch, visible focus and a touch target of at least 44 by 44 CSS pixels. Its placement SHALL respect device safe areas and SHALL NOT obstruct active form controls, dialogs, task-help popovers or action feedback. Animated scrolling SHALL be disabled when reduced motion is requested. Without JavaScript, a usable return-to-top link SHALL remain available in the footer.

#### Scenario: Return from a long page

- **WHEN** a viewer scrolls beyond one viewport height on a long schedule, roster or statistics page
- **THEN** the floating arrow is available
- **AND** activating it returns to the page top without changing filters, assignments or form values

#### Scenario: Keyboard and reduced motion

- **WHEN** a keyboard user with reduced motion enabled activates "Nach oben"
- **THEN** the page returns to the top without animated scrolling
- **AND** keyboard focus remains usable and is not lost when the floating control is hidden

#### Scenario: Other floating surfaces are active

- **WHEN** an action message, task-help popover or modal dialog is displayed on a narrow screen
- **THEN** the return-to-top action does not obscure or intercept interactions with that surface

#### Scenario: JavaScript is unavailable

- **WHEN** a viewer reaches the footer without JavaScript
- **THEN** a return-to-top link remains available and works

### Requirement: General page explanations use responsive disclosure

General introductory explanations on the schedule, helper-management and statistics pages SHALL start collapsed on viewports at most 700 CSS pixels wide when responsive enhancement is available. A German control labeled "Hinweise" SHALL expose the complete explanation. On wider initial viewports the explanations SHALL start expanded. The disclosures SHALL support keyboard and touch and SHALL not require JavaScript to be opened manually. JavaScript failure SHALL leave the text available. Validation errors, actionable staffing or eligibility warnings, setup-needed states, missing-birth-date follow-up notices and per-control instructions SHALL NOT be hidden by this general explanation disclosure. Existing task-help popovers, authentication guidance and legal content SHALL retain their existing presentation.

#### Scenario: Mobile introduction

- **WHEN** a viewer opens one of the three pages on a mobile viewport with responsive enhancement available
- **THEN** its general introduction is collapsed behind "Hinweise"
- **AND** opening the disclosure reveals the full text

#### Scenario: Desktop introduction

- **WHEN** a viewer opens one of the three pages on a viewport wider than 700 CSS pixels
- **THEN** its general introduction starts expanded

#### Scenario: An actionable issue exists

- **WHEN** a game has an age deficiency or cake setup is needed, or a form has a validation error
- **THEN** the relevant warning or error remains available in its existing context independently of the introduction disclosure

### Requirement: Disclosure captions identify content consistently

Disclosure captions SHALL remain the same in their expanded and collapsed states and SHALL identify content without action verbs. General explanations SHALL use "Hinweise", filter forms "Filter", person maintenance "Bearbeitung", game/day-card details "Details", and the past schedule group "Vergangene Spieltage". Schedule game/day cards SHALL show a triangle beside "Details", pointing right when closed and down when open, including without JavaScript. Statistics SHALL retain its section titles and counts as captions. Native disclosure state SHALL remain available to keyboard and assistive-technology users; decorative triangles SHALL NOT replace that state.

#### Scenario: Caption remains concise after expansion

- **WHEN** a viewer opens or closes a disclosure
- **THEN** its content caption remains unchanged
- **AND** native disclosure state still identifies whether the area is expanded or collapsed

#### Scenario: Game and day-card indicator follows disclosure state

- **WHEN** a viewer opens or closes a schedule game or day-task card, with or without JavaScript
- **THEN** the triangle beside "Details" points down for an open card and right for a closed card
- **AND** the "Details" caption remains unchanged

### Requirement: Responsive defaults preserve deliberate user choices

Responsive introductory, filter, helper-maintenance and statistics disclosures SHALL allow independent manual opening and closing. Responsive changes SHALL NOT override a disclosure the viewer has deliberately toggled during the current page visit or close a form containing unsaved edits or validation errors. A fresh page load SHALL restore the documented initial defaults, except for an affected validation-error form.

#### Scenario: Resize after opening a disclosure

- **WHEN** a viewer manually opens a disclosure and then rotates the device or changes viewport width
- **THEN** the disclosure remains open and entered form values remain intact
