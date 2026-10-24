# Statistics Overview Specification

## Purpose

Defines how signed-in viewers browse season statistics through independently expandable sections, useful counts and responsive initial presentation.

## Requirements

### Requirement: Statistics preserve section order with responsive defaults

The statistics page SHALL retain the section order "Spiele pro Mannschaft", "Dienste pro Person", then "Offene Dienste, Altersanforderungen und Einrichtung". Section titles and counts SHALL remain the captions in both expanded and collapsed states, without alternate action labels. On viewports at most 700 CSS pixels wide with responsive enhancement available, all three sections SHALL start collapsed. On wider initial viewports, "Spiele pro Mannschaft" SHALL start expanded while the other two sections SHALL start collapsed. Each section SHALL support independent keyboard and touch expansion, and opening one SHALL NOT close another. Without JavaScript all section content SHALL remain reachable through native disclosure; the team section SHALL start expanded as the desktop fallback.

#### Scenario: Mobile statistics entry

- **WHEN** a signed-in viewer opens statistics on mobile with responsive enhancement available
- **THEN** the three sections appear collapsed in the existing order
- **AND** the viewer can independently open any combination of them

#### Scenario: Desktop statistics entry

- **WHEN** a signed-in viewer opens statistics on a viewport wider than 700 CSS pixels
- **THEN** "Spiele pro Mannschaft" is first and expanded
- **AND** the per-person and outstanding-staffing sections start collapsed

### Requirement: Statistics summaries identify counts and empty states

Section headers SHALL show readable German counts identifying the number of displayed teams, people with counted duties, and affected upcoming games or day blocks respectively. The outstanding count SHALL count affected containers within the effective date-visibility scope once, even when one has multiple vacancies, age deficiencies or setup needs; it SHALL NOT be labeled as a total number of missing helpers. Headers SHALL remain visible when a section is empty and SHALL show a zero count. Expanded empty sections SHALL explain that no corresponding data or outstanding issues exist within their applicable scope. With an active restriction and no visible outstanding rows, the explanation SHALL describe the displayed period and SHALL NOT claim that every season date is fully staffed or configured. Team and person counts SHALL remain season-wide. Existing retained-duty counting, eligibility reporting and access/privacy guarantees SHALL remain unchanged.

#### Scenario: Outstanding container has several issues

- **WHEN** one visible upcoming game has multiple empty positions and an age deficiency
- **THEN** it contributes one affected game to the section-header count
- **AND** expanding the section shows its vacancies and age deficiency separately

#### Scenario: Empty statistics

- **WHEN** there are no games or no assigned duties or no outstanding issues in the applicable scope
- **THEN** each corresponding section remains identifiable by its header and count
- **AND** expanding it shows a clear German empty-state explanation describing that scope

#### Scenario: Existing counts are preserved

- **WHEN** a person has game duties, day-block duties or retained duties outside a game's offered positions, including on hidden dates
- **THEN** expanding "Dienste pro Person" shows the existing authoritative season-wide duty counts
- **AND** no full birth dates or exact personal ages are added to statistics

#### Scenario: All remaining issues are hidden

- **WHEN** the only outstanding current/future games or blocks are on hidden dates
- **THEN** the ordinary outstanding section shows zero affected containers
- **AND** its expanded explanation states that the displayed period has no outstanding issues without asserting that the whole season is complete

### Requirement: Outstanding statistics follow the shared date visibility scope

The statistics section "Offene Dienste, Altersanforderungen und Einrichtung" SHALL include only current/future games and day blocks within the same effective date-visibility scope as the schedule. A saved boundary SHALL exclude dates on or after it, inclusive, from vacancy, eligibility-deficiency and setup-needed rows. The rule SHALL cover preparation, cleanup and cake delivery, including unconfigured cake blocks. Administrators SHALL be able to select or leave the full-date view on statistics without changing the shared setting; this SHALL restore hidden current/future outstanding containers without adding past issues. Other tiers SHALL NOT activate that override. An unparseable game date SHALL retain its existing outstanding-issue behavior rather than be silently hidden by the boundary. Existing handling of invalid block dates SHALL remain unchanged.

"Spiele pro Mannschaft", "Dienste pro Person", their denominators and all other season distribution/personal totals SHALL continue to include the entire current season, including duties on hidden dates and retained removed duties. The page SHALL explain this difference in scope whenever a boundary is active. Full-date mode SHALL be preserved when an administrator navigates between the schedule and statistics from that mode.

#### Scenario: Member views hidden staffing needs

- **WHEN** a member or MV opens statistics with a boundary of 01.12.2026
- **THEN** outstanding games and preparation, cleanup and cake blocks dated 01.12.2026 or later are absent from the rows and affected-container count
- **AND** visible current/future containers still show their vacancies, eligibility deficiencies and setup needs under the existing rules

#### Scenario: Hidden assignments still count toward the season

- **WHEN** a person has duties on dates before and after the boundary
- **THEN** both sets of duties still contribute to "Dienste pro Person" and its role breakdown
- **AND** team coverage, shares and season denominators remain based on all season games
- **AND** the page explains that these totals cover the entire season

#### Scenario: Administrator inspects all upcoming staffing needs

- **WHEN** an administrator enables the full-date statistics view
- **THEN** outstanding hidden current/future games and blocks are included without altering the setting
- **AND** past games and blocks remain excluded by existing outstanding-date rules
- **AND** navigating to the schedule retains the authorized full-date mode
