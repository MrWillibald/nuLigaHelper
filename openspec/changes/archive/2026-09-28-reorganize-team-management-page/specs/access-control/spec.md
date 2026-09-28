## ADDED Requirements

### Requirement: Available team management precedes the roster

The person-management page SHALL display the team-management section before the roster when the viewer has team-management actions available. The section SHALL contain only controls the viewer is already authorized to use. The roster heading, filters, and person entries SHALL follow the section. The order SHALL remain the same when roster filters are applied, and changing the order SHALL NOT change action permissions or roster-data visibility.

#### Scenario: Administrator opens person management

- **WHEN** an administrator opens the person-management page
- **THEN** the controls to create users, decide pending registrations, and appoint team MVs appear before the roster heading, filters, and person entries
- **AND** those actions retain their existing authorization rules

#### Scenario: MV opens person management

- **WHEN** an MV who is not an administrator opens the person-management page
- **THEN** the control to create a user for a managed team appears before the roster heading, filters, and person entries
- **AND** registration-decision and MV-appointment controls are not shown

#### Scenario: Member opens person management

- **WHEN** a member without MV or administrator rights opens the person-management page
- **THEN** the roster is shown without a team-management section
- **AND** the member's existing contact-data visibility and self-edit rights are unchanged

#### Scenario: Viewer filters the roster

- **WHEN** a signed-in viewer applies a name, team, or permitted status filter on the person-management page
- **THEN** the filter changes only the roster results
- **AND** any team-management section available to that viewer remains before the filtered roster
