# Spec Delta

## MODIFIED Requirements

### Requirement: Staffing summaries use a compact responsive layout

Collapsed game and day-task summaries SHALL keep identity, staffing needs and actionable warnings readable with compact padding and line spacing around their separate information rows. The ring and vacancy badges SHALL form one shared status area with the ring beside content-sized badges placed directly next to each other. Badges SHALL flow in reading order with small gaps and wrap onto a new line only when the next badge does not fit; they SHALL NOT reserve equal-width columns or stretch to fill the row; long labels and enlarged text SHALL reflow without clipping or horizontal page scrolling. Absent badges SHALL NOT reserve width or additional rows. The "Details" control SHALL share existing summary space.

On wide screens the status area SHALL sit beside the game/block information. On narrow screens it SHALL move below that information when necessary, keeping the ring beside the badge group as space permits. Each game summary SHALL display four separate primary information rows in this order: start time; full matchup or existing Spielfest identity; age class, hall and game number together; responsible-team name prefixed by "Verantwortlich:" or the existing open placeholder when unset. The time SHALL occupy its own row even when the viewport has room to place it beside the matchup. Preparation, cleanup and cake-delivery summaries SHALL display three separate primary information rows in this order: calculated/configured time or the existing time placeholder, with any associated date qualifier; task label with its existing icon; existing task metadata such as "Tagesdienst" or requested cake quantity. Day-task cards SHALL NOT show a responsible-team row or reserve a blank fourth row.

The start time or time placeholder on every game and day-task card SHALL be slightly larger than its previous presentation and larger than the matchup or task-label text. The time row SHALL have additional space before the matchup or task label to separate it visually from the remaining information. Full matchup names, game metadata, responsible-team identity, eligibility and classification warnings and setup-needed states SHALL remain available without truncation or opening the card. Supplemental warnings SHALL NOT replace the primary information rows. Each primary row SHALL start below the preceding row, but its content SHALL wrap naturally when necessary. The row order and time emphasis SHALL apply at mobile and desktop widths, to upcoming and past cards, to guests and signed-in viewers, and while a card is collapsed or expanded, including without JavaScript. Ring counts and labels SHALL remain readable at 320, 360 and 390 CSS pixel viewport widths and with enlarged text. Readability and keyboard/touch access SHALL take precedence over a fixed card height or physical line count.

#### Scenario: Phone with several missing roles

- **WHEN** the schedule is viewed at 360 CSS pixels wide with several vacant game roles
- **THEN** the ring appears beside content-sized badges placed directly next to each other, wrapping when the next badge no longer fits
- **AND** the card requires no horizontal scrolling and the disclosure control has no separate full-width action row

#### Scenario: Wide display

- **WHEN** a summary has sufficient width for game information and the badge group
- **THEN** the status area appears beside the information and the badges fill each line in reading order before wrapping, without equal-width column gaps
- **AND** the time remains in its own information row above the matchup or task label

#### Scenario: Long names and text enlargement

- **WHEN** a 320 CSS pixel viewport displays a long matchup, "Ordnungsdienst", a long responsible-team name or enlarged text
- **THEN** the text and badges wrap as needed without clipping, overlap or horizontal page scrolling
- **AND** ring counts, warnings and disclosure focus remain readable and usable
- **AND** wrapping preserves the order of the primary information rows without forcing ellipses or a fixed card height

#### Scenario: Few or no vacancies

- **WHEN** a card has one or no vacancy badges
- **THEN** the status area reserves no height for unused badge rows or a repeated percentage/count line

#### Scenario: Four game information rows

- **WHEN** a guest or signed-in viewer views an upcoming or past game at mobile or desktop width, with or without JavaScript
- **THEN** its summary shows the start time in a separate first row with slightly larger text than before and larger text than the matchup
- **AND** the full matchup or existing Spielfest identity starts in the second row
- **AND** age class, hall and game number share the third information row, allowing natural wrapping
- **AND** "Verantwortlich:" and the responsible team or existing open placeholder start in the fourth row
- **AND** opening the card preserves that summary structure

#### Scenario: Preparation and cleanup follow the same hierarchy

- **WHEN** a viewer views a preparation or cleanup card
- **THEN** its calculated time or existing "Zeit offen" placeholder starts in a separate, slightly enlarged first row, retaining any date qualifier
- **AND** the task label with its existing icon starts in the second row and "Tagesdienst" starts in the third row
- **AND** the summary has no responsible-team row or blank row reserved for one

#### Scenario: Cake cards preserve configuration states

- **WHEN** a viewer views a cake-delivery card with a positive quantity, an explicit zero quantity or incomplete configuration
- **THEN** its delivery time or existing "Lieferzeit offen" placeholder starts in a separate, slightly enlarged first row
- **AND** the cake label with its icon starts in the second row and the configured quantity or existing "Anzahl offen" metadata starts in the third row
- **AND** setup-needed and "Keine Kuchen angefragt" states and the applicable presence or absence of ring and badges remain visible according to the existing cake-card requirements
- **AND** no responsible-team row or blank fourth row appears

#### Scenario: Warnings remain visible beside ordered information

- **WHEN** a game has an unresolved classification or staffing eligibility deficiency
- **THEN** all four primary information rows remain visible in order
- **AND** the existing warning remains visible as supplemental information without overlap with the time, responsible team, staffing area or "Details" control
