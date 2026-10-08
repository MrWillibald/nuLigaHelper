# Spec Delta

## Purpose

Defines how signed-in viewers browse season statistics through independently expandable sections, useful counts and responsive initial presentation.

## ADDED Requirements

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

Section headers SHALL show readable German counts identifying the number of displayed teams, people with counted duties, and affected upcoming games or day blocks respectively. The outstanding count SHALL count affected containers once, even when one has multiple vacancies, age deficiencies or setup needs; it SHALL NOT be labeled as a total number of missing helpers. Headers SHALL remain visible when a section is empty and SHALL show a zero count. Expanded empty sections SHALL explain that no corresponding data or outstanding issues exist. Existing statistics calculations, retained-duty counting, eligibility reporting and access/privacy guarantees SHALL remain unchanged.

#### Scenario: Outstanding container has several issues

- **WHEN** one upcoming game has multiple empty positions and an age deficiency
- **THEN** it contributes one affected game to the section-header count
- **AND** expanding the section shows its vacancies and age deficiency separately

#### Scenario: Empty statistics

- **WHEN** there are no games or no assigned duties or no outstanding issues
- **THEN** each corresponding section remains identifiable by its header and count
- **AND** expanding it shows a clear German empty-state explanation

#### Scenario: Existing counts are preserved

- **WHEN** a person has game duties, day-block duties or retained duties outside a game's offered positions
- **THEN** expanding "Dienste pro Person" shows the existing authoritative duty counts
- **AND** no full birth dates or exact personal ages are added to statistics
