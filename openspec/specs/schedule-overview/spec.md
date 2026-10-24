# Schedule Overview Specification

## Purpose

Defines how visitors find relevant home games and distinguish completed dates from upcoming dates on the public game overview.

## Requirements

### Requirement: Overview filters by a selected game date

The overview SHALL offer a single-date dropdown containing the dates of all home games
in the current season in chronological order and an all-dates option selected by default.
The date filter SHALL combine with the playing-team, responsible-team, and assigned-person
filters. When a date is selected, only games and preparation or cleanup blocks on that
date that satisfy the other active filters SHALL be visible. A nonempty date value that
does not identify a current-season game date SHALL match no cards rather than show all dates.
The date filter SHALL work for guests and signed-in viewers without JavaScript. Clearing
filters SHALL restore the all-dates option.

#### Scenario: Select a date

- **WHEN** a visitor selects a date with multiple home games
- **THEN** only that date's games and its preparation and cleanup blocks are shown
- **AND** the selected date remains selected after the page loads

#### Scenario: Combine date and existing filters

- **WHEN** a visitor selects a date and a playing team
- **THEN** only games on that date matching the playing team are shown
- **AND** the date's blocks bookend those matching games

#### Scenario: Date with no remaining match

- **WHEN** the selected date has no games or blocks matching the other active filters, or the date
  value is not a current-season game date
- **THEN** the ordinary no-results message and clear-filter action are shown
- **AND** no other date or block is shown

### Requirement: Overview filters games by team and assigned helper

The overview SHALL offer filters for playing team, responsible team, and assigned person.
Selected team filters SHALL apply to games, and a date SHALL remain visible only when at
least one game on that date satisfies every selected team filter. An unset responsible
team SHALL not match a selected responsible team. A selected responsible team SHALL
hide all preparation and cleanup blocks because blocks have no responsible team.

The person filter SHALL match the names of people assigned to each game or day-level
preparation or cleanup block independently. It SHALL retain a game only when that game
has a matching assignee, and retain a block only when that block has a matching assignee.
A matching block SHALL be allowed to retain its date without any displayed game, provided
the date satisfies the selected team filters. Filtering SHALL not search the roster or
empty slots. It SHALL retain season ordering and SHALL show date and month headings only
when they contain matching content.

#### Scenario: Combine filters

- **WHEN** a visitor selects a playing team and responsible team and enters an assigned
  person's name
- **THEN** the overview shows only games satisfying both team filters and the person filter
- **AND** no day-task blocks are shown
- **AND** the remaining games stay in chronological order within their date groups

#### Scenario: Responsible team is open

- **WHEN** a visitor selects a responsible team
- **THEN** games with no responsible team are excluded
- **AND** preparation and cleanup blocks are excluded

#### Scenario: Duplicate names

- **WHEN** multiple people with the same displayed name hold assignments on different
  games or day blocks
- **THEN** a name search can match each assignment without treating the name as a unique
  person identity

#### Scenario: Person matches a day block

- **WHEN** the person filter matches a preparation or cleanup assignee but no game
  assignee on a qualifying date, and no responsible-team filter is selected
- **THEN** the date and only the matching block or blocks remain visible
- **AND** no game is shown merely because its date contains a matching block

#### Scenario: Person matches a game only

- **WHEN** the person filter matches a game's assignee but no day-block assignee
- **THEN** the matching game remains visible without either day-task block

### Requirement: Day task blocks bookend each visible game date

When visible, the overview SHALL render the preparation block before displayed games of
its date and the cleanup block after displayed games of its date. Filtering SHALL omit a
block that does not match and SHALL allow a matching block to appear without a displayed game.
Each block SHALL show its calculated time and staffing progress while collapsed, and SHALL show its
numbered display slots and assigned helper names when expanded. The numbers SHALL
distinguish UI positions without creating separate task roles. It SHALL not show or ask
for a responsible team. A guest response SHALL expose no roster payload, person IDs,
contact data, or assignment controls for the blocks. Block cards SHALL use the same plain
surface and border treatment as game cards, without a phase-colored background fade or side
bar. A `#a0e656` `↗` icon SHALL identify preparation, and a `#ffb752` `↘` icon SHALL
identify cleanup.

#### Scenario: Unfiltered date presentation

- **WHEN** a visitor opens a date containing several games
- **THEN** the preparation block appears before all game cards
- **AND** the cleanup block appears after all game cards

#### Scenario: Filter leaves only a middle game

- **WHEN** filtering retains only a middle game without a person or responsible-team filter
- **THEN** both blocks still bookend the displayed games
- **AND** their times remain derived from the actual first and last games of the full date

#### Scenario: Filter leaves only a day-task block

- **WHEN** a person filter matches only one block on a date with no matching game assignments
- **THEN** that block appears under the date heading without unrelated games or the other block
- **AND** its time remains derived from the full date's games

#### Scenario: Guest views block assignments

- **WHEN** an unauthenticated visitor expands a day-task block
- **THEN** assigned block helper names are visible
- **AND** the response contains no block person IDs, roster payload, contact data, or
  assignment controls

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
- **AND** successfully claiming its empty offered Ordner position updates required progress

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

### Requirement: Filters are usable by all schedule viewers

The overview SHALL allow guests and signed-in viewers to apply and clear filters. The
person filter SHALL be a case-insensitive name search over assigned people visible when
game or day-task block cards are expanded. The guest response SHALL contain no roster
payload, person IDs, or contact data. When no cards match, the overview SHALL show a
clear German empty result message and a way to clear the filters; when the season has
no games at all, it SHALL retain a distinct no-games message.

#### Scenario: Guest filters by assigned name

- **WHEN** a guest searches for an assigned person's name
- **THEN** matching game or day-task block cards are shown without a roster, person IDs,
  or contact data in the response

#### Scenario: No matching games

- **WHEN** filters exclude every game and day-task block in the season
- **THEN** the overview explains that no results match and offers a clear-filter action

#### Scenario: Clear filters

- **WHEN** a visitor clears the active filters
- **THEN** all games and day-task blocks for the current season appear in their normal
  upcoming and past sections

### Requirement: Past games remain available in a muted expandable section

Games whose date is before the application's effective current day SHALL appear in a
separate past-games section that is collapsed by default and can be expanded and collapsed
by the visitor. When a specific past date is selected, that section SHALL start expanded
so the selected date is immediately visible. Individual game and block cards SHALL still
start collapsed. Games on the current day SHALL remain in the upcoming section. Past games
SHALL have a visually muted but readable treatment consistent with the rest of the
overview. Filtering SHALL apply to both sections, and the past section's label SHALL
reflect the number of matching past game days.

#### Scenario: Open and close past games

- **WHEN** a visitor opens the overview with past games and no date selected
- **THEN** past game cards are initially hidden behind a labeled expand control
- **AND** activating the control reveals them and activating it again hides them

#### Scenario: Effective date boundary

- **WHEN** a game's date is the effective current day
- **THEN** that game is shown with upcoming games

#### Scenario: Filter past games

- **WHEN** a visitor applies a filter other than a date selection that matches only past games
- **THEN** the collapsed past section indicates the number of matching past game days
- **AND** expanding it reveals those matching games without unrelated dates or month headings

#### Scenario: Select a past date

- **WHEN** a visitor selects a past date from the date dropdown
- **THEN** the past-games section starts expanded with only matching content from that date
- **AND** its game and block cards start collapsed

### Requirement: Editable cards load candidates when opened

The initial schedule response SHALL show the current assignment state without including lists of unassigned candidates in collapsed or expanded card markup. When a signed-in viewer opens an editable game or day-task card, its candidate controls SHALL become usable after the permitted candidates have loaded. Opening one card SHALL not load candidates for other cards. The summary and assigned helper names SHALL remain available before candidate loading completes.

#### Scenario: Initial overview with a growing roster

- **WHEN** active, unassigned people are added to the roster and the same schedule is requested
- **THEN** the initial response does not repeat those people across task controls
- **AND** its candidate payload remains absent regardless of roster size

#### Scenario: Open one editable game

- **WHEN** a signed-in viewer opens an editable game card
- **THEN** that card loads its permitted candidates and makes its task controls usable
- **AND** unopened cards do not load their candidates

#### Scenario: Open an editable day-task block

- **WHEN** a signed-in viewer opens a preparation or cleanup card with editable slots
- **THEN** that block loads its permitted candidates independently of games and other blocks

#### Scenario: Candidate loading fails

- **WHEN** a candidate request fails or the session has expired
- **THEN** the affected controls remain unable to submit an assignment based on an incomplete list
- **AND** the viewer sees a German error or sign-in message and can retry after recovery
- **AND** current assignments remain visible and unchanged

### Requirement: Age compliance is distinct from physical staffing progress

Game cards SHALL distinguish physical required-slot occupancy from unresolved or invalid age eligibility. An age deficiency SHALL NOT invent an empty slot, conceal an assigned occupant, or alter the physical occupancy count. A game with an outstanding age deficiency SHALL not be presented as having complete valid staffing solely because its occupancy bar shows 100 percent.

Current/future game cards SHALL provide a readable German eligibility indication identifying the affected duty or adult-coverage requirement without exposing a full birth date or exact personal age. This indication SHALL remain visible while the card is collapsed and separate from the occupancy count and bar. Saved assignment changes and refreshed game data SHALL update the affected indication.

The displayed guidance SHALL explain that Ordner requires 18 for every game, adult-game Zeitnehmer and Sekretär require 16, and youth/SPF timing duties require 14. It SHALL explain that youth/SPF Verkauf has no age requirement while adult and unresolved-category games retain collective adult coverage. Guidance SHALL remain consistent with saved-state eligibility indications.

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

- **WHEN** a current/future game's Ordner is occupied by a helper below 18 or with an unknown birth date
- **THEN** the collapsed card reports the Ordner eligibility deficiency in German while retaining the occupant and physical occupancy
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

### Requirement: Cake cards follow day-block filtering and privacy

Cake cards SHALL participate in date and assigned-person filtering under the same rules as other day-task cards. A playing-team filter SHALL retain a cake card only for a date with a qualifying game. A responsible-team filter SHALL hide the cake card because it has no responsible team. A person-name filter SHALL retain the cake card only when one of its own assignments matches, and SHALL allow it to retain a qualifying date without a matching game card. Clearing filters SHALL restore cake cards in their normal upcoming or past date groups.

Guests SHALL see assigned names without roster payloads, person identifiers, contact data or assignment/configuration controls. Editable cake cards SHALL load their permitted candidates only when opened, retaining the existing conflict and candidate-failure behavior.

#### Scenario: Find a cake volunteer

- **WHEN** a person-name filter matches only a cake volunteer on a qualifying date
- **THEN** the date and cake card remain visible
- **AND** unrelated games and other day blocks are not shown merely because they share that date

#### Scenario: Responsible-team filter is selected

- **WHEN** a viewer selects a responsible-team filter
- **THEN** cake cards are hidden

#### Scenario: Guest views cake assignments

- **WHEN** a guest expands a configured cake card
- **THEN** the assigned helper names are visible
- **AND** the response contains no roster payload, person identifiers, contacts or editing controls

#### Scenario: Open an editable cake card

- **WHEN** a signed-in viewer opens an editable cake card
- **THEN** only that card loads its permitted candidates
- **AND** a failed request leaves its controls unable to submit claims based on an incomplete candidate list

### Requirement: Schedule filters have a compact mobile presentation

The existing schedule filter form SHALL start collapsed behind "Filter" on viewports at most 700 CSS pixels wide when responsive enhancement is available and SHALL start expanded on wider initial viewports. Applied non-default filter values SHALL remain visible as a readable German summary outside the collapsed form, with an available clear-filter action. The summary SHALL use human-readable date, team and entered-name values rather than exposing internal team or person identifiers. Selected values SHALL remain in the form after submission. Opening and closing the form SHALL NOT change results or submit filters. Existing combined filtering, past-date expansion, empty results and guest privacy guarantees SHALL continue to apply. Filtering and clearing filters SHALL remain usable without JavaScript, with the form accessible through native disclosure or an expanded fallback.

#### Scenario: Unfiltered mobile schedule

- **WHEN** a guest or signed-in viewer opens the unfiltered schedule on mobile with responsive enhancement available
- **THEN** the filter form starts collapsed behind "Filter"
- **AND** the viewer can open it and submit any currently offered filter

#### Scenario: Filters remain apparent after submission

- **WHEN** a viewer applies a date, team or assigned-person filter and the mobile page reloads
- **THEN** the filter form starts collapsed with its selected values retained
- **AND** an outside summary identifies the active values and offers clearing the filters
- **AND** the resulting games and blocks match the existing filtering rules

#### Scenario: Desktop and no-JavaScript filtering

- **WHEN** the schedule is opened on desktop, or filters are used without JavaScript
- **THEN** the existing GET submission and clear-filter behavior remain usable
- **AND** desktop filters start expanded

### Requirement: Assignment commitment and replacement state are visible

Every game and day-block duty SHALL show its assigned helper and persistent `Ersatz gesucht` status while an unlocked request is active, and SHALL explain when an occupied assignment is fixed. Guests SHALL receive only public names and read-only duty status without person IDs, roster payloads, contacts, full birth dates, exact ages or editing controls. Authorized claim and takeover flows SHALL provide the pre-save commitment confirmation defined by task self-service. Locked occupied duties SHALL not expose a release or replacement action; genuinely empty offered positions SHALL retain permitted claim actions. Existing filters, expansion behavior, responsive presentation and staffing-deficiency indications SHALL remain intact.

#### Scenario: Guest sees a replacement request

- **WHEN** a guest views an assignment seeking replacement
- **THEN** its helper name and `Ersatz gesucht` status remain visible
- **AND** no private identifiers, candidates or editing controls appear

#### Scenario: Cancel confirmation on any card

- **WHEN** a user cancels a claim confirmation in a game or any block phase
- **THEN** the card displays authoritative saved state without suggesting a claim succeeded

#### Scenario: Fixed and empty positions coexist

- **WHEN** a date has reached cutoff with some occupied and some genuinely empty duties
- **THEN** occupied duties show fixed state and permitted empty duties remain claimable

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
