# Production Operations Delta

## ADDED Requirements

### Requirement: Development overrides come only from the environment

Local testing MAY pin the effective date through `NULIGAHELPER_TODAY` and suppress outbound mail and SMS through `NULIGAHELPER_DRY_RUN`. Both overrides SHALL be read from the environment when they are used, SHALL be validated strictly and SHALL NOT be stored in committed source code.

#### Scenario: Local run pins the date without sending

- **WHEN** a local daily job runs with `NULIGAHELPER_TODAY=2025-11-21` and `NULIGAHELPER_DRY_RUN=1`
- **THEN** every date-dependent decision uses 21.11.2025
- **AND** no e-mail or SMS is handed to a provider while dispatch counts are still reported

#### Scenario: Date override alone still sends

- **WHEN** a local run sets only `NULIGAHELPER_TODAY`
- **THEN** the pinned date is used and notifications are delivered normally

#### Scenario: Invalid override value

- **WHEN** `NULIGAHELPER_TODAY` is not a valid ISO calendar date or `NULIGAHELPER_DRY_RUN` has a value other than `1`
- **THEN** the process stops with a German error naming the variable
- **AND** it does not fall back to the real date or to sending

### Requirement: Production refuses development overrides

When `NULIGAHELPER_ENV` is `production`, web startup, the daily job and the production preflight and launch check SHALL refuse to run while `NULIGAHELPER_TODAY` or `NULIGAHELPER_DRY_RUN` is set, before serving requests, scraping, synchronizing, notifying or backing up.

#### Scenario: Override set in production

- **WHEN** `NULIGAHELPER_ENV=production` and `NULIGAHELPER_TODAY` or `NULIGAHELPER_DRY_RUN` is set
- **THEN** web startup, the daily job, preflight and launch check fail before any work
- **AND** the diagnostic names the offending variable without printing other environment values

#### Scenario: Production without overrides

- **WHEN** `NULIGAHELPER_ENV=production` and neither override is set
- **THEN** the effective date is the current Europe/Berlin date and notifications are delivered normally
