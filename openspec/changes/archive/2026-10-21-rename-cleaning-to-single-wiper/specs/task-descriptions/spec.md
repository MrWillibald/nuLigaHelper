# Spec Delta

## MODIFIED Requirements

### Requirement: Every supported task has a description

The system SHALL provide a concise German description for every supported game and day-level assignment task. Each numbered position SHALL use the description of its semantic role. Descriptions SHALL explain the task's purpose or work without inventing requirements that are not part of the club's task definition. A task introduced or renamed by another change SHALL retain complete description coverage when that change is adopted.

#### Scenario: Current tasks have coverage

- **WHEN** the current game and day-level tasks are displayed
- **THEN** Zeitnehmer, Sekretär, Verkauf, Ordnungsdienst, Kasse, Wischer, Vorbereitung, Aufräumen and Kuchenlieferung each have a description

#### Scenario: Added or renamed duties have coverage

- **WHEN** a supported duty is introduced or renamed
- **THEN** its description is included in the actual supported role catalog
- **AND** obsolete labels are not presented as separate current tasks

#### Scenario: Numbered positions share meaning

- **WHEN** a viewer reads descriptions for Verkauf 1 and Verkauf 2
- **THEN** both positions explain the same Verkauf duty
- **AND** the same rule applies to other numbered roles

#### Scenario: Wischer help uses the current singular duty

- **WHEN** a viewer opens help for the single Wischer position
- **THEN** the German description explains the existing during-play floor-wiping duty using singular Wischer wording
- **AND** it does not imply that two Wischer are required or change the meaning of day-wide Aufräumen
