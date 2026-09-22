## ADDED Requirements

### Requirement: Task candidate lists prioritize suitable teams

The system SHALL present the people available in each task-assignment selection control as four consecutive, mutually exclusive categories in this order: members of the game's responsible team, members of the Supporter team, members of other teams or no team, and members of the team currently playing the game. People SHALL be ordered alphabetically by display name within each category, with a deterministic order for duplicate names. A member of the playing team SHALL belong to the final category even if that team is also selected as responsible; otherwise responsible-team membership SHALL take precedence over Supporter membership. An absent category, including the responsible-team category when no responsible team is selected, SHALL simply contribute no people.

This ordering SHALL only arrange people whom the viewer is already authorized to assign; it SHALL NOT expand their assignment rights. A person already assigned to another task of the same game SHALL be omitted. The currently selected person SHALL remain visible in their own task control. Existing advisory hints for members of the playing team and members outside the responsible and Supporter teams SHALL be retained.

#### Scenario: Full roster is grouped and sorted

- **WHEN** an administrator opens a task selection for a game with a responsible team
- **THEN** available people appear in responsible-team, Supporter, other-team, and playing-team order
- **AND** people within each category are ordered alphabetically by display name
- **AND** playing-team and outside-team candidates retain their advisory hints

#### Scenario: Already assigned person is excluded

- **WHEN** a person holds one task for a game
- **THEN** that person is absent from every other task selection for that game
- **AND** the person remains visible as the selected value for their current task

#### Scenario: Released person returns in the correct position

- **WHEN** a task assignment is released without reloading the schedule
- **THEN** the released person is restored to the other task selections for that game in the correct team category and alphabetical position
- **AND** the person's advisory hint remains unchanged

#### Scenario: Limited assignment scope remains limited

- **WHEN** a member or MV opens a task selection whose candidates are restricted by their access tier
- **THEN** only the candidates already permitted for that viewer appear
- **AND** those candidates follow the applicable category and alphabetical ordering
