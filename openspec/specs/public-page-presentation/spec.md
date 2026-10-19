# Public Page Presentation Specification

## Purpose

Defines the shared visual treatment of public page footer links so legal information remains easy to find and read across the site.

## Requirements

### Requirement: Legal footer links fit the shared design

The shared footer SHALL present the Impressum and Datenschutzerklärung links in a color and style that is readable against the footer background and consistent with the site's visual language. Both links SHALL remain visually identifiable as links and SHALL have visible hover and keyboard focus states on desktop and narrow screens.

#### Scenario: Navigate legal links from the footer

- **WHEN** a visitor views the footer on the overview or another public page
- **THEN** both legal links are readable and identifiable as interactive links
- **AND** activating each link opens its corresponding legal page

#### Scenario: Keyboard focus

- **WHEN** a visitor tabs to either legal footer link
- **THEN** its focus state is clearly visible against the footer background

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

### Requirement: Responsive disclosure defaults precede content presentation

Each shared responsive disclosure on the schedule, helper-management and statistics pages SHALL apply its initial native expanded or collapsed state before its contents can first be displayed when its early responsive enhancement succeeds. On viewports at most 700 CSS pixels wide, general Hinweise panels, schedule and roster Filter panels, and statistics sections SHALL start collapsed without briefly displaying their expanded contents while the main application scripts load. Wider initial viewports SHALL retain the existing expanded introduction/filter/team-statistics defaults and closed defaults for other statistics sections and person maintenance. The affected person-maintenance form returned with validation errors SHALL remain expanded regardless of viewport width.

This behavior SHALL apply to fresh navigation, reloads and filter-submission responses. It SHALL preserve the existing German captions, native keyboard/touch operation, independent expansion and permission boundaries. Active-filter summaries and clear/reset actions SHALL remain available outside collapsed filters. When JavaScript is disabled or early enhancement fails, native disclosures SHALL remain manually usable, with the existing server-rendered expanded fallback for introductions, filters and team statistics. The enhancement SHALL work with the existing production script policy without permitting inline or third-party scripts.

#### Scenario: Mobile navigation while application scripts are delayed

- **WHEN** a viewer navigates to the schedule, helper-management or statistics page at a viewport width of at most 700 CSS pixels, early enhancement succeeds, and the main application scripts have not loaded
- **THEN** the Hinweise and any Filter panels first appear collapsed without displaying their expanded contents
- **AND** the viewer can open and close them through their native controls

#### Scenario: Mobile filter response retains visible context

- **WHEN** a viewer submits a permitted schedule or roster filter on mobile and early enhancement succeeds on the resulting page
- **THEN** the filter's contents are collapsed before they can first appear
- **AND** the selected values remain in the form and the permitted active values and clear/reset action remain visible outside it
- **AND** the result set and visibility rules retain their existing behavior

#### Scenario: Mobile team statistics do not expand while loading

- **WHEN** a signed-in viewer opens statistics at a viewport width of at most 700 CSS pixels with early enhancement available
- **THEN** all statistics sections first appear collapsed, including the team section with its expanded server fallback
- **AND** the existing section order, captions and counts remain visible

#### Scenario: Desktop initial presentation

- **WHEN** a viewer opens an affected page at a viewport width greater than 700 CSS pixels
- **THEN** introductions, filters and team statistics start expanded
- **AND** the other statistics sections and ordinary person-maintenance areas start closed

#### Scenario: Validation response is available immediately

- **WHEN** authorized person maintenance returns an affected form with validation errors
- **THEN** that form remains expanded while early and later scripts initialize on mobile or desktop
- **AND** its errors and preserved submitted values remain accessible without expanding unrelated person-maintenance areas

#### Scenario: Early enhancement is unavailable

- **WHEN** JavaScript is disabled or the early enhancement asset fails to load
- **THEN** introductions, filters and team statistics retain their expanded server fallback
- **AND** native controls still allow manual opening and closing and existing form submission remains available
- **AND** no stylesheet hides the contents of an otherwise expanded disclosure

#### Scenario: Only some panels receive early enhancement

- **WHEN** early enhancement fails for one disclosure but succeeds for another on the same page, and the main application scripts subsequently load
- **THEN** the failed disclosure retains its current native fallback or manually chosen state instead of receiving a late initial-default reset
- **AND** successfully initialized disclosures retain their existing enhancement and independent interaction state

### Requirement: Later initialization preserves early disclosure interaction

After early disclosure enhancement initializes, later application initialization SHALL NOT reset a disclosure's state or discard interaction protection. Deliberate native opening or closing SHALL be preserved for the current page visit. Viewport changes SHALL NOT override deliberate choices or close a form containing unsaved edits, active focus or validation errors. Untouched disclosures SHALL continue to follow the existing responsive defaults on breakpoint changes. Shared disclosure initialization SHALL NOT start assignment-candidate requests or change saved data.

#### Scenario: Native interaction before the main bundle loads

- **WHEN** early enhancement has initialized, a viewer deliberately opens or closes a disclosure, and the main application scripts finish loading afterwards
- **THEN** the viewer's chosen state remains in effect after later initialization and a subsequent breakpoint change
- **AND** independent disclosures retain their own states

#### Scenario: Edits or focus precede later initialization

- **WHEN** a viewer edits a field or focuses a form inside an early-initialized disclosure before the main application scripts load and then changes viewport width
- **THEN** later initialization and resizing leave that form open and preserve its entered values and usable focus

#### Scenario: Untouched disclosure crosses the breakpoint

- **WHEN** an initialized disclosure without deliberate interaction, edits, focus or validation errors crosses the 700 CSS pixel breakpoint
- **THEN** it follows its existing mobile or desktop default
- **AND** initializing or toggling a general disclosure starts no assignment-candidate request and performs no data mutation
