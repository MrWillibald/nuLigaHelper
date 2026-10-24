# Spec Delta

## MODIFIED Requirements

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
