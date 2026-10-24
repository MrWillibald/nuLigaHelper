# Schedule Overview Delta

## ADDED Requirements

### Requirement: Administrators manage a durable optional visibility boundary

The schedule SHALL offer current active administrators a setting labeled "Heimspieltage ausblenden ab". A saved date SHALL be the first hidden calendar date, inclusive; an empty value SHALL mean unrestricted visibility. The setting SHALL be shared across viewers and application processes and survive restarts and season changes until an administrator changes or clears it. A newly chosen or changed nonempty date MUST be a valid calendar date strictly after the application's effective current day. An unchanged saved value SHALL remain valid as a no-op even after time has advanced past it. Invalid or stale edits SHALL leave the saved setting unchanged and provide German correction or conflict feedback. Concurrent edits based on an outdated saved value SHALL NOT overwrite a newer value. Saving, clearing and validation SHALL work without JavaScript.

#### Scenario: Administrator saves a future boundary

- **WHEN** an active administrator saves 01.12.2026 while the effective current day is 24.10.2026
- **THEN** that date becomes the shared first hidden date after a successful explicit save
- **AND** a new request from another viewer or application process uses the saved value
- **AND** restarting the application does not clear it

#### Scenario: Administrator clears the boundary

- **WHEN** an active administrator explicitly saves an empty value against the current saved value
- **THEN** the visibility boundary is removed and all current-season dates become available under ordinary filters

#### Scenario: Invalid proposed date

- **WHEN** an administrator submits a malformed date or changes the boundary to today or an earlier date
- **THEN** the request is refused with German field feedback and the saved boundary is unchanged
- **AND** submitting the identical already-saved date remains a no-op even if it is now in the past

#### Scenario: Two administrators edit the setting

- **WHEN** one administrator saves a new boundary after a second administrator loaded the old value
- **THEN** saving the second administrator's stale form reports a conflict without overwriting the first change
- **AND** the current saved value is available for review

#### Scenario: Time and season advance

- **WHEN** a saved boundary has been reached or a new season begins
- **THEN** the saved date continues to apply without automatically advancing, clearing or rolling forward
- **AND** administrators retain the ability to view all dates or clear/change the boundary

### Requirement: Visibility omits complete dated content without changing responsibility

Unless the current administrator explicitly selects the full-date view, the schedule SHALL omit every game and preparation, cake-delivery and cleanup block whose valid date is on or after the saved boundary. Date/month headings SHALL exist only for remaining matching content. Content with an unparseable or missing date SHALL retain existing unresolved-date presentation instead of being silently hidden by an unprovable comparison. With no boundary, existing visibility SHALL remain unchanged.

The boundary SHALL affect presentation only. It SHALL NOT remove stored games, blocks or assignments, change responsible teams, locks or replacement requests, create assignment-audit transitions, suppress scraping/synchronization/notifications, or introduce a restriction on otherwise authorized direct candidate and assignment operations. All existing authentication, authorization, confirmation, age, stale-write, capacity and cutoff rules SHALL continue to apply to hidden dates.

#### Scenario: Boundary is inclusive for every block phase

- **WHEN** the boundary is 01.12.2026 and a viewer uses the ordinary schedule view
- **THEN** a matching day on 30.11.2026 remains visible
- **AND** no game, preparation, cake-delivery or cleanup card from 01.12.2026 or later appears in the response
- **AND** headings with no remaining cards are absent

#### Scenario: Hidden commitment remains binding and notified

- **WHEN** an administrator hides a date containing assignments, open replacement requests or saved locks
- **THEN** those saved records and their history remain unchanged
- **AND** existing notification selection and due-lock enforcement still process those duties according to their ordinary rules

#### Scenario: Direct authorized assignment remains possible

- **WHEN** a signed-in viewer directly requests candidates or confirms a claim or takeover for a hidden date
- **THEN** the request is evaluated under the existing current-state assignment rules without a visibility-based refusal
- **AND** the ordinary non-JavaScript confirmation route obeys the same policy
- **AND** hidden status grants no additional permission or lock exception

#### Scenario: Date cannot be parsed

- **WHEN** a stored game has a missing or unparseable date while a boundary is active
- **THEN** the boundary does not silently exclude that game
- **AND** existing unresolved-date handling and assignment restrictions remain applicable

### Requirement: The current display scope is understandable and read-only

An active restriction SHALL be explained in German on the schedule with its inclusive date, for example "Heimspieltage ab dem 01.12.2026 sind derzeit ausgeblendet." Administrators in full-date view SHALL instead see that hidden dates are currently included and retain the saved boundary in the explanation. When the season contains games but no cards remain within the current scope, the schedule SHALL explain the restriction rather than present the season as having no games. Opening pages, changing the full-date view, opening or dismissing the setting form, and applying/clearing ordinary filters SHALL perform no setting or duty write, successful assignment audit, lock persistence or notification. All controls and notices SHALL remain usable by keyboard, at narrow viewport widths and without JavaScript.

#### Scenario: Boundary leaves no visible cards

- **WHEN** the season contains games but the boundary hides all cards available in the current view
- **THEN** the response explains the active date scope without instructing the viewer to import a supposedly empty season
- **AND** an administrator can still reach the setting and full-date view

#### Scenario: View actions do not save

- **WHEN** a viewer loads the schedule, changes display filters or an administrator previews hidden dates without submitting a setting change
- **THEN** settings, assignments, locks, requests, audit records and notifications remain unchanged

## MODIFIED Requirements

### Requirement: Overview filters by a selected game date

The overview SHALL offer a single-date dropdown containing the dates of all home games in the current season that are within the viewer's effective visibility scope, in chronological order and with an all-dates option selected by default. The administrator full-date view SHALL include dates hidden by the shared boundary. The date filter SHALL combine with the playing-team, responsible-team, and assigned-person filters. When a date is selected, only games and preparation or cleanup blocks on that date that satisfy the other active filters and the effective visibility scope SHALL be visible. A nonempty date value that does not identify an available current-season game date SHALL match no cards rather than show all dates; a manually supplied hidden date SHALL NOT bypass the boundary or reveal its cards. The date filter SHALL work for guests and signed-in viewers without JavaScript. Clearing filters SHALL restore the all-dates option within the effective visibility scope and preserve an authorized administrator full-date view.

#### Scenario: Select a date

- **WHEN** a visitor selects an available date with multiple home games
- **THEN** only that date's games and its preparation and cleanup blocks are shown
- **AND** the selected date remains selected after the page loads

#### Scenario: Combine date and existing filters

- **WHEN** a visitor selects an available date and a playing team
- **THEN** only games on that date matching the playing team are shown
- **AND** the date's blocks bookend those matching games

#### Scenario: Date with no remaining match

- **WHEN** the selected date has no games or blocks matching the other active filters, or the date value is not an available current-season game date
- **THEN** the ordinary no-results message and clear-filter action are shown
- **AND** no other date or block is shown

#### Scenario: Hidden date is supplied manually

- **WHEN** a guest, member or MV sends a date filter matching a date on or after the saved boundary
- **THEN** no cards for that date are returned and it is absent from the offered date choices
- **AND** clearing the ordinary filter does not disable the saved boundary

### Requirement: Filters are usable by all schedule viewers

The overview SHALL allow guests and signed-in viewers to apply and clear filters within their effective date-visibility scope. The person filter SHALL be a case-insensitive name search over assigned people visible when game or day-task block cards are expanded. The guest response SHALL contain no roster payload, person IDs, or contact data. When no cards match, the overview SHALL show a clear German empty result message and a way to clear the ordinary filters while retaining the central date boundary and any authorized administrator full-date view. When the season has no games at all, it SHALL retain a distinct no-games message; an empty limited view SHALL not be presented as an empty season.

#### Scenario: Guest filters by assigned name

- **WHEN** a guest searches for an assigned person's name
- **THEN** matching game or day-task block cards within the effective visibility scope are shown without a roster, person IDs, or contact data in the response

#### Scenario: No matching games

- **WHEN** filters exclude every game and day-task block available in the effective visibility scope
- **THEN** the overview explains that no results match and offers a clear-filter action

#### Scenario: Clear filters

- **WHEN** a visitor clears the active ordinary filters
- **THEN** all games and day-task blocks for the current season within the effective visibility scope appear in their normal upcoming and past sections
- **AND** an authorized administrator full-date view remains active if it was selected

### Requirement: Past games remain available in a muted expandable section

Games within the viewer's effective date-visibility scope whose date is before the application's effective current day SHALL appear in a separate past-games section that is collapsed by default and can be expanded and collapsed by the visitor. When a specific available past date is selected, that section SHALL start expanded so the selected date is immediately visible. Individual game and block cards SHALL still start collapsed. Visible games on the current day SHALL remain in the upcoming section. Past games SHALL have a visually muted but readable treatment consistent with the rest of the overview. Filtering SHALL apply to both sections, and the past section's label SHALL reflect the number of matching visible past game days. A retained boundary SHALL apply even when its date has become past; the administrator full-date view SHALL restore those dates to the normal past grouping.

#### Scenario: Open and close past games

- **WHEN** a visitor opens the overview with visible past games and no date selected
- **THEN** past game cards are initially hidden behind a labeled expand control
- **AND** activating the control reveals them and activating it again hides them

#### Scenario: Effective date boundary

- **WHEN** a visible game's date is the effective current day
- **THEN** that game is shown with upcoming games

#### Scenario: Filter past games

- **WHEN** a visitor applies a filter other than a date selection that matches only visible past games
- **THEN** the collapsed past section indicates the number of matching visible past game days
- **AND** expanding it reveals those matching games without unrelated dates or month headings

#### Scenario: Select a past date

- **WHEN** a visitor selects an available past date from the date dropdown
- **THEN** the past-games section starts expanded with only matching content from that date
- **AND** its game and block cards start collapsed

#### Scenario: Fixed boundary has moved into the past

- **WHEN** the effective day is later than the saved boundary and some hidden dates are now past
- **THEN** those dates remain excluded from the ordinary past section and its day count
- **AND** an administrator can show them in the full-date view without changing the setting
