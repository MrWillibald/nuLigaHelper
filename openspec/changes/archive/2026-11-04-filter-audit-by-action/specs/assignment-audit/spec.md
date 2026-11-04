# Spec Delta

## ADDED Requirements

### Requirement: Administrators filter the audit by action

The audit page SHALL offer an `Aktion` filter listing every action present in the audit with a German label, plus `Alle Aktionen`. Selecting an action SHALL show only entries with that action, combined with the target and person filters. The filter SHALL be a read-only GET and SHALL remain administrator-only.

#### Scenario: Filtering by one action

- **WHEN** an administrator selects `Sperre` and submits the filter
- **THEN** only entries with the stored action `lock` are listed
- **AND** `Sperre` stays selected in the filter

#### Scenario: Action combined with other filters

- **WHEN** an administrator filters by an action together with a target and a person
- **THEN** only entries matching all three filters are listed

#### Scenario: Action list is independent of other filters

- **WHEN** a target or person filter is active
- **THEN** the `Aktion` list still offers every action stored in the audit

#### Scenario: Historical action without a label

- **WHEN** the audit contains an action without a German label
- **THEN** the filter offers it under its stored name and it can be selected

#### Scenario: Unknown action value

- **WHEN** a request names an action that is not stored in the audit
- **THEN** the page shows no entries and no error

#### Scenario: Filter is reset

- **WHEN** an administrator chooses `Zurücksetzen`
- **THEN** all entries are listed and `Alle Aktionen` is selected

#### Scenario: Non-administrator requests a filtered audit

- **WHEN** a guest, member or MV opens the audit page with an action filter
- **THEN** access is refused as for the unfiltered page
