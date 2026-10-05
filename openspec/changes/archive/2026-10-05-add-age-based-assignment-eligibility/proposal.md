# Proposal

## Why

The roster contains no birth dates, so task assignments cannot enforce age requirements. Collecting birth dates and applying game-date eligibility checks will prevent underage timing appointments while allowing younger helpers to work in Verkauf alongside an adult.

## What Changes

- Store a validated birth date for every newly created person, including registrations, admin-created persons, MV-created persons, and CLI creation.
- Preserve existing people with unknown birth dates and provide a path to complete their data without inventing dates or deleting assignments.
- Evaluate age on the game's scheduled date.
- Require Zeitnehmer to be at least 14 for youth games and 18 for adult games; require Sekretär to be at least 14 for youth games and 16 for adult games.
- Require at least one explicitly assigned Verkauf helper to be **18 or older**, eligible from their 18th birthday. The other seller has no individual minimum introduced by this change.
- Enforce the rules in every assignment entry point and refresh candidate eligibility when sibling slots change.
- Allow incomplete Verkauf staffing and existing releases, while preventing a new claim from filling the final Verkauf slot without a qualifying adult seller.
- Preserve existing assignments after migration, birth-date corrections, and game changes; report unresolved or invalid age eligibility separately from physical slot vacancies.
- Keep full birth dates out of schedule responses and expose them only in authorized person-maintenance contexts.
- **BREAKING**: new-person creation and registration require a valid birth date, and previously accepted age-ineligible assignment claims are refused.

## Capabilities

### New Capabilities

- `age-eligibility`: Birth-date validation, game-date age calculation, individual task minimums, the collective Verkauf rule, missing-data handling, and revalidation.

### Modified Capabilities

- `user-accounts`: Registration and person creation collect birth dates; authorized profile maintenance supports correcting them.
- `access-control`: Birth-date visibility is limited to the person themselves and administrators.
- `schema-migrations`: A guarded migration adds birth dates without inventing existing data or altering assignment history.
- `task-self-service`: Claims and candidates obey age rules while retaining access scope, compare-and-swap, withdrawals, and current-occupant visibility.
- `schedule-overview`: Age-compliance deficiencies are visible separately from physical staffing progress.

## Impact

- Database and operations: person records, assignment helpers, a reviewed Alembic revision, migration postflight, CLI person creation and correction, and operator documentation.
- Web and browser: registration, roster forms, per-game candidate responses, claim validation, sibling-control refresh, and eligibility notices.
- Staffing and notifications: shared staffing-status evaluation for overview, statistics, and existing responsible-team MV reminders.
- Integration: use the same explicit adult/youth classification as `update-game-duty-staffing`; coordinate German feedback with `centralize-message-templates` without requiring a particular implementation order.
- Tests: synthetic birth-date and birthday boundaries, assignment authority, collective-sale concurrency, missing data, migrations, notification routing, and private-data exclusion.
- Dependencies: no new external runtime dependency is required.
