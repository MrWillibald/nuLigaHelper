# Spec Delta

## ADDED Requirements

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
