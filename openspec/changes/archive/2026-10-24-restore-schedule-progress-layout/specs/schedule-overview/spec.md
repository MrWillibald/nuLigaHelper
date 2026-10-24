# Schedule Overview Spec Delta

## MODIFIED Requirements

### Requirement: Every game and day-task card starts compact

Every game, preparation block, cleanup block and cake-delivery block SHALL start collapsed on page load, regardless of responsible-team selection or occupied slots. A collapsed game card SHALL show its time, matchup or Spielfest identity, existing game metadata, responsible-team name or an open placeholder when unset, and the horizontal progress bar, descriptive occupancy count and percentage defined below. A collapsed preparation or cleanup card SHALL show its calculated time, block label, horizontal progress bar, descriptive occupancy count and percentage. Cake cards SHALL follow their configuration-specific presentation below.

Each card SHALL have a labeled control that expands and collapses that card independently. Its visible caption SHALL remain "Details" with a right-pointing triangle when closed and a downward-pointing triangle when open. The caption and indicator SHALL sit in the lower right corner of the summary on mobile and desktop, right-aligned beneath the progress/status area with enough space to avoid overlapping staffing or warnings. The expanded game SHALL reveal the responsible team and all existing task fields according to the game's offered set, plus occupied removed duties marked as existing assignments subject to binding protection and the explicit legacy retirement policy. Expanded preparation and cleanup cards SHALL reveal all three numbered assignment fields; expanded cake cards SHALL reveal positions according to their configured quantity. All viewers SHALL be able to expand cards, while assignment controls and private data SHALL continue to depend on existing access rights. Expansion SHALL be usable with a keyboard and touch, including without JavaScript.

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
- **AND** empty Kasse and Wischer fields and optional markers are absent
- **AND** any retained removed-duty occupant appears separately as an existing assignment with only controls allowed by binding protection and the explicit legacy retirement policy

#### Scenario: Compact disclosure stays accessible

- **WHEN** a visitor opens or closes any game or day-task card by keyboard or touch, with or without JavaScript
- **THEN** the "Details" caption remains visible and its triangle reflects the expanded state
- **AND** keyboard focus and native disclosure state remain usable
- **AND** opening or closing the card performs no assignment or replacement-request mutation

### Requirement: Compact cards show accurate staffing progress

Each game SHALL display a horizontal progress bar representing occupied positions in that game's offered/required-position set. Above the bar, a descriptive German occupied/required count SHALL appear on the left, such as "5 von 7 Pflichtdiensten besetzt", with the percentage on the right. Bar fill and percentage SHALL use the same ratio, rounded to the nearest whole percent. Retained removed-duty positions and responsible-team selection SHALL be excluded from numerator and denominator. Adult games SHALL have seven required positions; youth and SPF games SHALL have five. Preparation and cleanup SHALL count occupied positions among their respective three positions and use wording such as "1 von 3 Plätzen besetzt". Configured positive cake quantities SHALL use the requested quantity as denominator and wording such as "2 von 4 Kuchen zugesagt". Schedule summaries SHALL show no staffing rings, central fractions, vacancy badges or reserved badge placeholders.

Each bar SHALL expose its current value, total and a contextual German description to assistive technology. Meaning SHALL NOT depend on color alone. Counts, percentages and fills SHALL be present on initial rendering without JavaScript. Successful permitted occupancy changes and authoritative staffing refreshes, including candidate refreshes, SHALL update the count, percentage, fill and accessible values together without a page reload. Selection, opening, cancellation, dismissal, rejected writes and stale responses SHALL NOT display unsaved or older state as current progress.

#### Scenario: Optional support does not complete a game

- **WHEN** a youth game has three occupied required slots and an occupied retained removed Kasse slot
- **THEN** its compact card shows "3 von 5 Pflichtdiensten besetzt", "60 %" and a 60-percent bar fill
- **AND** the retained assignment does not increase required progress

#### Scenario: Block coverage

- **WHEN** one of a preparation block's three slots is occupied
- **THEN** its compact card shows "1 von 3 Plätzen besetzt", "33 %" and a matching bar fill

#### Scenario: Assignment changes progress

- **WHEN** a viewer successfully claims a position or completes an atomic takeover on an expanded card
- **THEN** that card's descriptive count, percentage, bar fill and accessible values reflect saved staffing
- **AND** a rejected assignment leaves the displayed progress at its saved values

#### Scenario: Adult additional duties count

- **WHEN** an adult game has five of its seven required positions occupied
- **THEN** its compact card shows "5 von 7 Pflichtdiensten besetzt", "71 %" and a matching bar fill

#### Scenario: Retained removed duty released

- **WHEN** a retained Kasse or Wischer assignment of a youth game is explicitly retired by an authorized caller before the fixed day-before cutoff
- **THEN** its required-duty progress display is unchanged and the removed duty field disappears
- **AND** successfully claiming its empty offered Ordnungsdienst position updates required progress

#### Scenario: Replacement wanted is still occupied

- **WHEN** a replacement request opens or closes without takeover
- **THEN** physical occupancy and the progress display stay unchanged
- **AND** an atomic takeover changes responsibility without reducing progress

#### Scenario: Empty and full gauges

- **WHEN** a card with a positive capacity has no occupied positions or every position occupied
- **THEN** its bar is empty with a zero occupied count and "0 %", or full with all positions counted and "100 %", respectively
- **AND** its count and German accessible description remain available without JavaScript

#### Scenario: Cancelled confirmation and stale responses

- **WHEN** a viewer cancels or dismisses confirmation, or an older candidate response arrives after a newer saved assignment
- **THEN** the card does not display the attempted selection or older response as current progress
- **AND** cancellation and dismissal perform no assignment or replacement-request mutation

#### Scenario: Staffing refresh reflects another saved change

- **WHEN** an authorized candidate refresh returns staffing changed by another saved assignment
- **THEN** the affected game's or block's count, percentage, fill and accessible values reflect that same saved snapshot

#### Scenario: Initial public staffing display

- **WHEN** a guest opens the schedule without JavaScript
- **THEN** descriptive counts, percentages and horizontal bars are visible before expanding cards
- **AND** no staffing rings, vacancy badges, private identifiers, candidate lists, contacts, full birth dates or exact personal ages are exposed in the summaries

### Requirement: Age compliance is distinct from physical staffing progress

Game cards SHALL distinguish physical required-slot occupancy from unresolved or invalid age eligibility. An age deficiency SHALL NOT invent an empty slot, conceal an assigned occupant, or alter the physical occupancy count. A game with an outstanding age deficiency SHALL not be presented as having complete valid staffing solely because its occupancy bar shows 100 percent.

Current/future game cards SHALL provide a readable German eligibility indication identifying the affected duty or adult-coverage requirement without exposing a full birth date or exact personal age. This indication SHALL remain visible while the card is collapsed and separate from the occupancy count and bar. Saved assignment changes and refreshed game data SHALL update the affected indication.

The displayed guidance SHALL explain that Ordnungsdienst requires 18 for every game, adult-game Zeitnehmer and Sekretär require 16, and youth/SPF timing duties require 14. It SHALL explain that youth/SPF Verkauf has no age requirement while adult and unresolved-category games retain collective adult coverage. Guidance SHALL remain consistent with saved-state eligibility indications.

#### Scenario: Full game lacks an adult seller

- **WHEN** every required slot of a current/future adult or unresolved-category game is occupied but Verkauf has no qualifying adult seller
- **THEN** the occupancy bar remains full with an accurate descriptive count and "100 %" and no invented Verkauf vacancy
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

#### Scenario: Full youth-sale group has no adult

- **WHEN** a recognized youth or SPF game's Verkauf positions contain only younger helpers or legacy helpers with unknown birth dates
- **THEN** its card shows no sale adult-coverage or sale unknown-birth-date deficiency
- **AND** its physical occupancy bar and unrelated eligibility warnings remain accurate

#### Scenario: Retained Ordnungsdienst fails the new rule

- **WHEN** a current/future game's Ordnungsdienst is occupied by a helper below 18 or with an unknown birth date
- **THEN** the collapsed card reports the Ordnungsdienst eligibility deficiency in German while retaining the occupant and physical occupancy
- **AND** it exposes no full birth date or exact personal age

#### Scenario: Adult Zeitnehmer meets the new threshold

- **WHEN** a current/future adult-game Zeitnehmer is at least 16 on its scheduled date
- **THEN** its card shows no deficiency based on the former minimum of 18

#### Scenario: Guidance without JavaScript

- **WHEN** a signed-in viewer opens the schedule without JavaScript
- **THEN** the displayed German age guidance describes the revised thresholds and youth-sale exemption

### Requirement: Cake delivery appears as a compact dated task card

The overview SHALL show the date's cake-delivery block alongside its existing preparation, game and cleanup cards. The cake card SHALL start collapsed and SHALL show its label, configured delivery time or setup-needed state, requested cake quantity and staffing status. Expanding it SHALL reveal one numbered volunteer position per requested cake and each position's assigned helper name. Each position SHALL represent one cake contribution. For a configured positive quantity, a horizontal bar, descriptive German count and percentage SHALL show promised cakes out of that quantity. Count, percentage, fill, accessible values, time and quantity SHALL update together after successful assignments or configuration edits. A refused or stale operation SHALL not show uncommitted progress or configuration.

Configured zero SHALL show "Keine Kuchen angefragt" without a progress bar, occupancy count, percentage, division-by-zero or vacancy badge. An unconfigured block SHALL visibly require administrator setup, show no progress bar, occupancy count, percentage or invented vacancy count and SHALL not offer claim controls.

The cake card SHALL use the established plain card treatment, a small 🍰 emoji beside its label and a labeled keyboard-operable expansion control. It SHALL have no responsible-team field. Administrators SHALL be offered controls for the date's delivery time and quantity; delivery time entry SHALL consistently use 24-hour `HH:MM` regardless of the browser locale. Other viewers SHALL see those settings without edit controls.

#### Scenario: Configured cake block is expanded

- **WHEN** a viewer expands a cake block configured for four cakes at 10:00
- **THEN** the card shows that delivery time and quantity
- **AND** exactly four independently displayed cake positions are shown

#### Scenario: Cake block needs setup

- **WHEN** a cake block has no delivery time or requested quantity
- **THEN** it shows that administrator setup is needed
- **AND** no invented time, quantity, progress bar, percentage, vacancy count or claimable volunteer positions are shown

#### Scenario: Administrator adjusts delivery settings

- **WHEN** an administrator saves valid new delivery settings
- **THEN** the card shows the saved time and requested quantity
- **AND** its position count, descriptive count, percentage and bar reflect the saved configuration

#### Scenario: No cakes requested

- **WHEN** a cake block is configured with a quantity of zero
- **THEN** its collapsed summary shows "Keine Kuchen angefragt"
- **AND** it shows no progress bar, occupancy count, percentage or vacancy badge

#### Scenario: Remaining cake contributions are grouped

- **WHEN** two of four requested cakes have been promised
- **THEN** the summary shows "2 von 4 Kuchen zugesagt", "50 %" and a half-filled horizontal bar
- **AND** a successful further claim updates the count to "3 von 4 Kuchen zugesagt", percentage to "75 %" and bar fill accordingly

### Requirement: Staffing summaries use a compact responsive layout

Game and day-task summaries SHALL restore the roomier card presentation shown in the agreed reference screenshot, retaining the established navy/neutral palette, plain card surfaces, date grouping and task icons. On viewports at most 900 CSS pixels wide, game/block information SHALL appear above the progress/status area, with the horizontal bar spanning the available inner width. On wider viewports, information SHALL appear on the left and progress/status on the right. The "Details" caption and triangle SHALL be right-aligned beneath progress and associated warnings in normal reading order. No ring-and-badge status area SHALL appear.

Each game summary SHALL display four primary information rows: prominent start time; full matchup or existing Spielfest identity; age class, hall and game number together; responsible-team name prefixed by "Verantwortlich:" or the existing open placeholder. Preparation, cleanup and cake-delivery summaries SHALL display three primary information rows: calculated/configured time or existing time placeholder with any date qualifier; task label with its existing icon; existing task metadata such as "Tagesdienst" or requested cake quantity. Day-task cards SHALL have no responsible-team row or blank fourth row. Time SHALL remain larger than the matchup/task label, using the earlier presentation's emphasis rather than an additional size increase.

Full names, metadata, descriptive counts, percentages, eligibility/classification warnings and setup-needed states SHALL remain readable without truncation or opening the card. Labels and percentages SHALL wrap when necessary while retaining reading order and an unobstructed bar. Content SHALL NOT clip, overlap the disclosure control or require horizontal page scrolling at 320, 360 or 390 CSS pixels or with 200-percent text enlargement. Card height SHALL follow content. This presentation SHALL apply to upcoming and past cards, guests and signed-in viewers, collapsed and expanded summaries, and initial rendering without JavaScript.

#### Scenario: Phone with several missing roles

- **WHEN** the schedule is viewed at 360 CSS pixels wide with several vacant roles
- **THEN** game information is followed by a descriptive count and percentage above a full-width horizontal bar
- **AND** "Details" appears right-aligned below progress/status without badges or horizontal scrolling

#### Scenario: Wide display

- **WHEN** a summary is viewed above 900 CSS pixels wide
- **THEN** information appears beside the progress/status column
- **AND** time remains above the title and "Details" appears beneath progress/status

#### Scenario: Long names and text enlargement

- **WHEN** a 320 CSS pixel viewport displays long matchup/responsible-team names or text is enlarged to 200 percent
- **THEN** names, counts and percentages wrap without clipping, overlap or horizontal page scrolling
- **AND** warnings, the bar and disclosure focus remain readable and usable

#### Scenario: Few or no vacancies

- **WHEN** a positive-capacity card has one or no vacant positions
- **THEN** descriptive count, percentage and bar remain visible without reserved badge space

#### Scenario: Four game information rows

- **WHEN** a guest or signed-in viewer views an upcoming or past game at mobile or desktop width, with or without JavaScript
- **THEN** time appears above the full matchup, metadata and responsible-team row in that order
- **AND** opening the card preserves that structure

#### Scenario: Preparation and cleanup follow the same hierarchy

- **WHEN** a viewer views a preparation or cleanup card
- **THEN** its prominent time or "Zeit offen" placeholder retains any date qualifier above the task label and "Tagesdienst" metadata
- **AND** no responsible-team row or blank fourth row is reserved

#### Scenario: Cake cards preserve configuration states

- **WHEN** a viewer views a cake card with a positive quantity, explicit zero or incomplete configuration
- **THEN** delivery time or "Lieferzeit offen" appears above the cake label and quantity or "Anzahl offen" metadata
- **AND** the applicable progress display, setup-needed indication or "Keine Kuchen angefragt" status remains visible without a responsible-team row

#### Scenario: Warnings remain visible beside ordered information

- **WHEN** a game has unresolved classification or a staffing eligibility deficiency
- **THEN** all four primary information rows remain visible in order
- **AND** the supplemental warning stays readable without overlapping time, responsible team, progress or "Details"

## REMOVED Requirements

### Requirement: Collapsed cards identify grouped staffing vacancies

**Reason**: The agreed overview restores descriptive occupancy counts, percentages and horizontal bars in place of staffing rings and red vacancy badges.

**Migration**: Remove schedule-summary vacancy badge rendering and updates. Use the modified progress and responsive-layout requirements for all game and day-task cards. Keep vacancy calculations, statistics badges, saved-state refreshes, expanded duty fields and separate eligibility warnings intact.
