# Schedule Overview Delta

## MODIFIED Requirements

### Requirement: Every game and day-task card starts compact

Every game, preparation block, cleanup block and cake-delivery block SHALL start collapsed on page load, regardless of responsible-team selection or occupied slots. A collapsed game card SHALL show its time, matchup or Spielfest identity, existing game metadata, responsible-team name or an open placeholder when unset, and the ring gauge and vacancy badges defined below. A collapsed preparation or cleanup card SHALL show its calculated time, block label, ring gauge and vacancy badges. Cake cards SHALL follow their configuration-specific presentation below.

Each card SHALL have a labeled control that expands and collapses that card independently. Its visible caption SHALL remain "Details" with a right-pointing triangle when closed and a downward-pointing triangle when open. The caption and indicator SHALL sit in the lower right corner of the summary on mobile and desktop, with enough reserved space to avoid overlapping staffing or warnings and without a separate full-width control. The expanded game SHALL reveal the responsible team and all existing task fields according to the game's offered set, plus occupied removed duties marked as existing assignments subject to binding protection and the explicit legacy retirement policy. Expanded preparation and cleanup cards SHALL reveal all three numbered assignment fields; expanded cake cards SHALL reveal positions according to their configured quantity. All viewers SHALL be able to expand cards, while assignment controls and private data SHALL continue to depend on existing access rights. Expansion SHALL be usable with a keyboard and touch, including without JavaScript.

#### Scenario: Open and close a game

- **WHEN** a visitor opens the overview containing a game without a responsible team
- **THEN** that game starts collapsed like every other game and shows the open responsible-team placeholder
- **AND** expanding it reveals the responsible-team field and every task field
- **AND** collapsing it hides those fields again

#### Scenario: Responsible team remains visible in a collapsed game

- **WHEN** a visitor opens the overview containing a game with a responsible team
- **THEN** its collapsed summary shows that team's name without opening the card
- **AND** expanding it reveals the responsible-team field and every task field

#### Scenario: Open a day-task block

- **WHEN** a visitor expands a preparation or cleanup card
- **THEN** all three numbered assignment fields and their visible occupants are shown
- **AND** other cards retain their current expansion states

#### Scenario: Youth duties and retained assignments

- **WHEN** a viewer expands a youth game
- **THEN** only the five offered baseline positions appear as game duty fields
- **AND** empty Kasse and Reinigung fields and optional markers are absent
- **AND** any retained removed-duty occupant appears separately as an existing assignment with only controls allowed by binding protection and the explicit legacy retirement policy

#### Scenario: Compact disclosure stays accessible

- **WHEN** a visitor opens or closes any game or day-task card by keyboard or touch, with or without JavaScript
- **THEN** the "Details" caption remains visible and its triangle reflects the expanded state
- **AND** keyboard focus and native disclosure state remain usable
- **AND** opening or closing the card performs no assignment or replacement-request mutation

### Requirement: Compact cards show accurate staffing progress

Each game SHALL display a ring gauge whose arc represents occupied positions in that game's offered/required-position set and whose center shows a readable occupied/required fraction such as `5/8`. Retained removed-duty positions and responsible-team selection SHALL be excluded from numerator and denominator. Adult games SHALL have eight required positions; youth games SHALL have five. Preparation and cleanup gauges SHALL count occupied slots among their respective three slots. Configured cake quantities greater than zero SHALL use the same ring presentation with the requested quantity as denominator. The overview SHALL replace the horizontal progress bar and separate percentage line with this ring and fraction, including at zero and full coverage.

Each gauge SHALL expose its current value, total and a German description identifying occupied duties, block positions or promised cakes to assistive technology. Meaning SHALL NOT depend on color alone. Initial counts and arcs SHALL be available without JavaScript. A successful permitted occupancy change SHALL update the affected ring and vacancy badges together without requiring a page reload; an unsuccessful, cancelled or stale operation SHALL not show uncommitted progress or vacancies.

#### Scenario: Optional support does not complete a game

- **WHEN** a youth game has three occupied required slots and an occupied retained removed Kasse slot
- **THEN** its compact card shows `3/5` in the ring with a 60-percent arc
- **AND** no separate percentage line is displayed

#### Scenario: Block coverage

- **WHEN** one of a preparation block's three slots is occupied
- **THEN** its compact card shows `1/3` in the ring with an approximately one-third arc

#### Scenario: Assignment changes progress

- **WHEN** a viewer successfully claims a position or completes an atomic takeover on an expanded card
- **THEN** that card's fraction, arc and vacancy badges reflect saved staffing
- **AND** a rejected assignment leaves the displayed progress and vacancies at their saved values

#### Scenario: Adult additional duties count

- **WHEN** an adult game has five of its eight required positions occupied
- **THEN** its compact card shows `5/8` required slots filled

#### Scenario: Retained removed duty released

- **WHEN** a retained Kasse or Reinigung assignment of a youth game is explicitly retired by an authorized caller before the fixed day-before cutoff
- **THEN** its required-duty gauge and vacancy badges are unchanged and the removed duty field disappears
- **AND** changing its sole Ordnungsdienst position updates the gauge and matching badge

#### Scenario: Replacement wanted is still occupied

- **WHEN** a replacement request opens or closes without takeover
- **THEN** physical occupancy, the ring and vacancy badges stay unchanged
- **AND** an atomic takeover changes responsibility without reducing progress or creating a vacancy badge

#### Scenario: Empty and full gauges

- **WHEN** a card with a positive capacity has no occupied positions or every position occupied
- **THEN** its ring shows an empty arc with `0/total` or a full arc with `total/total` respectively
- **AND** its count and German accessible description remain available without JavaScript

### Requirement: Age compliance is distinct from physical staffing progress

Game cards SHALL distinguish physical required-slot occupancy from unresolved or invalid age eligibility. An age deficiency SHALL NOT invent an empty slot or vacancy badge, conceal an assigned occupant, or alter the physical occupancy count. A game with an outstanding age deficiency SHALL not be presented as having complete valid staffing solely because its occupancy ring is full.

Current/future game cards SHALL provide a readable German eligibility indication identifying the affected duty or adult-coverage requirement without exposing a full birth date or exact personal age. This indication SHALL remain visible while the card is collapsed and distinguishable from vacancy badges. Saved assignment changes and refreshed game data SHALL update the affected indication.

#### Scenario: Full game lacks an adult seller

- **WHEN** every required slot of a current/future game is occupied but Verkauf has no qualifying adult seller
- **THEN** the occupancy ring remains full with an accurate fraction and no invented Verkauf vacancy
- **AND** the collapsed card visibly reports outstanding adult coverage instead of complete valid staffing

#### Scenario: Timing eligibility is unresolved

- **WHEN** a current/future timing assignment cannot be verified because of missing birth-date or game information
- **THEN** the assigned helper remains visible
- **AND** the duty carries an unresolved-eligibility indication

#### Scenario: Stored data changes compliance

- **WHEN** a successful correction or assignment mutation changes a game's eligibility status
- **THEN** the card reflects saved occupancy and eligibility state
- **AND** a rejected mutation does not display an uncommitted improvement

#### Scenario: Guest sees eligibility status

- **WHEN** a guest views a deficient game's public card
- **THEN** the card can explain the outstanding duty requirement
- **AND** its response contains no full birth date, exact personal age, private roster payload, or assignment controls

### Requirement: Cake delivery appears as a compact dated task card

The overview SHALL show the date's cake-delivery block alongside its existing preparation, game and cleanup cards. The cake card SHALL start collapsed and SHALL show its label, configured delivery time or setup-needed state, requested cake quantity and staffing status. Expanding it SHALL reveal one numbered volunteer position per requested cake and each position's assigned helper name. Each position SHALL represent one cake contribution. For a configured positive quantity, the ring SHALL count occupied positions out of that quantity and a grouped red vacancy badge SHALL identify remaining cake contributions. Ring, badge, time and quantity SHALL update together after successful assignments or configuration edits. A refused or stale operation SHALL not show uncommitted progress or configuration.

Configured zero SHALL show "Keine Kuchen angefragt" without a ring, vacancy badge, division-by-zero or invented staffing percentages. An unconfigured block SHALL visibly require administrator setup, show no ring or invented vacancy count and SHALL not offer claim controls.

The cake card SHALL use the established plain card treatment, a small 🍰 emoji beside its label and a labeled keyboard-operable expansion control. It SHALL have no responsible-team field. Administrators SHALL be offered controls for the date's delivery time and quantity; delivery time entry SHALL consistently use 24-hour `HH:MM` regardless of the browser locale. Other viewers SHALL see those settings without edit controls.

#### Scenario: Configured cake block is expanded

- **WHEN** a viewer expands a cake block configured for four cakes at 10:00
- **THEN** the card shows that delivery time and quantity
- **AND** exactly four independently displayed cake positions are shown

#### Scenario: Cake block needs setup

- **WHEN** a cake block has no delivery time or requested quantity
- **THEN** it shows that administrator setup is needed
- **AND** no invented time, quantity, gauge, vacancy count or claimable volunteer positions are shown

#### Scenario: Administrator adjusts delivery settings

- **WHEN** an administrator saves valid new delivery settings
- **THEN** the card shows the saved time and requested quantity
- **AND** its position count, ring and vacancy badge reflect the saved configuration

#### Scenario: No cakes requested

- **WHEN** a cake block is configured with a quantity of zero
- **THEN** its collapsed summary shows "Keine Kuchen angefragt"
- **AND** it shows no ring or vacancy badge

#### Scenario: Remaining cake contributions are grouped

- **WHEN** two of four requested cakes have been promised
- **THEN** the ring shows `2/4` and one red badge identifies two missing cake contributions
- **AND** a successful further claim updates both to `3/4` and one missing contribution

## ADDED Requirements

### Requirement: Collapsed cards identify grouped staffing vacancies

Each collapsed game card SHALL show red vacancy badges beside its ring, using the visual language and vacancy meaning of the statistics page. Each missing offered duty type SHALL appear once in a stable duty order, with its readable German name and a count when more than one position is vacant, such as "Verkauf (2×)" or "Reinigung (2×)". Occupied roles, retained removed-duty roles and responsible-team selection SHALL NOT produce vacancy badges. Preparation and cleanup SHALL each show one grouped badge for their remaining positions; configured positive cake quantities SHALL show one grouped badge for missing cake contributions. Fully occupied containers SHALL show no vacancy badges or empty badge placeholders.

The same vacancy information SHALL be available to guests and signed-in viewers on initial rendering without JavaScript. New summary content SHALL contain no person identifiers, candidate lists, contacts, full birth dates or exact personal ages. Badge updates SHALL reflect saved state from successful mutations and staffing refreshes, including candidate refreshes that return current staffing. Cancelled or dismissed confirmation, failed writes and stale responses SHALL NOT remove badges based on an unsaved selection. Existing replacement and fixed-duty states SHALL remain available in the expanded duty fields.

#### Scenario: Several roles are vacant

- **WHEN** both Verkauf positions and the single Kasse position are vacant in an adult game
- **THEN** its summary shows one red "Verkauf (2×)" badge and one red "Kasse" badge beside the ring
- **AND** the badges agree with the same game's vacancies on statistics

#### Scenario: Game type determines offered vacancies

- **WHEN** a youth or Spielfest card has empty Kasse and Reinigung positions outside its offered set
- **THEN** those roles produce no vacancy badges
- **AND** its ring continues to use its five required positions

#### Scenario: Fully occupied container

- **WHEN** every offered game duty or configured block position is occupied
- **THEN** its full ring remains visible without vacancy badges or reserved badge rows
- **AND** any applicable eligibility warning remains independently visible

#### Scenario: Cancellation and failed claim

- **WHEN** a viewer selects a candidate and cancels confirmation, or a confirmed claim is rejected
- **THEN** the ring and badges do not display the unsaved selection as an occupied position

#### Scenario: Staffing refresh reflects another saved change

- **WHEN** a permitted refresh returns staffing changed by another saved assignment
- **THEN** the affected card's ring and vacancy badges reflect that returned state together

#### Scenario: Guest sees public staffing needs

- **WHEN** a guest opens the schedule without JavaScript
- **THEN** the rings and grouped vacancy badges are present before any card is expanded
- **AND** their summary content exposes no private roster or person data

### Requirement: Staffing summaries use a compact responsive layout

Collapsed game and day-task summaries SHALL prioritize low vertical height while keeping identity, staffing needs and actionable warnings readable. The ring and vacancy badges SHALL form one shared status area with the ring beside content-sized badges placed directly next to each other. Badges SHALL flow in reading order with small gaps and wrap onto a new line only when the next badge does not fit; they SHALL NOT reserve equal-width columns or stretch to fill the row; long labels and enlarged text SHALL reflow without clipping or horizontal page scrolling. Absent badges SHALL NOT reserve width or additional rows. Card padding and line spacing SHALL be compact, and the "Details" control SHALL share existing summary space.

On wide screens the status area SHALL sit beside the game/block information. On narrow screens it SHALL move below that information when necessary, keeping the ring beside the badge group as space permits. Time and matchup/block label SHALL share a row where readable. Full matchup names, game metadata, responsible-team identity, eligibility warnings and setup-needed states SHALL remain available without truncation or opening the card. Ring counts and labels SHALL remain readable at 320, 360 and 390 CSS pixel viewport widths and with enlarged text. Readability and keyboard/touch access SHALL take precedence over a fixed card height or row count.

#### Scenario: Phone with several missing roles

- **WHEN** the schedule is viewed at 360 CSS pixels wide with several vacant game roles
- **THEN** the ring appears beside content-sized badges placed directly next to each other, wrapping when the next badge no longer fits
- **AND** the card requires no horizontal scrolling and the disclosure control has no separate full-width action row

#### Scenario: Wide display

- **WHEN** a summary has sufficient width for game information and the badge group
- **THEN** the status area appears beside the information and the badges fill each line in reading order before wrapping, without equal-width column gaps

#### Scenario: Long names and text enlargement

- **WHEN** a 320 CSS pixel viewport displays a long matchup, "Ordnungsdienst", a long responsible-team name or enlarged text
- **THEN** the text and badges wrap as needed without clipping, overlap or horizontal page scrolling
- **AND** ring counts, warnings and disclosure focus remain readable and usable

#### Scenario: Few or no vacancies

- **WHEN** a card has one or no vacancy badges
- **THEN** the status area reserves no height for unused badge rows or a repeated percentage/count line
