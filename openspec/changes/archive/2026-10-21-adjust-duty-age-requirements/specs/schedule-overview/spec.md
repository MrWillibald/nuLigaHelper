# Schedule Overview Delta

## MODIFIED Requirements

### Requirement: Age compliance is distinct from physical staffing progress

Game cards SHALL distinguish physical required-slot occupancy from unresolved or invalid age eligibility. An age deficiency SHALL NOT invent an empty slot or vacancy badge, conceal an assigned occupant, or alter the physical occupancy count. A game with an outstanding age deficiency SHALL not be presented as having complete valid staffing solely because its occupancy ring is full.

Current/future game cards SHALL provide a readable German eligibility indication identifying the affected duty or adult-coverage requirement without exposing a full birth date or exact personal age. This indication SHALL remain visible while the card is collapsed and distinguishable from vacancy badges. Saved assignment changes and refreshed game data SHALL update the affected indication.

The displayed guidance SHALL explain that Ordnungsdienst requires 18 for every game, adult-game Zeitnehmer and Sekretär require 16, and youth/SPF timing duties require 14. It SHALL explain that youth/SPF Verkauf has no age requirement while adult and unresolved-category games retain collective adult coverage. Guidance SHALL remain consistent with saved-state eligibility indications.

#### Scenario: Full game lacks an adult seller

- **WHEN** every required slot of a current/future adult or unresolved-category game is occupied but Verkauf has no qualifying adult seller
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

#### Scenario: Full youth-sale group has no adult

- **WHEN** a recognized youth or SPF game's Verkauf positions contain only younger helpers or legacy helpers with unknown birth dates
- **THEN** its card shows no sale adult-coverage or sale unknown-birth-date deficiency
- **AND** its physical occupancy ring and unrelated eligibility warnings remain accurate

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
