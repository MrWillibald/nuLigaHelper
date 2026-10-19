# Spec Delta

## ADDED Requirements

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
