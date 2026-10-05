# Tasks

## 1. Domain and migration

- [x] 1.1 Add or reuse the shared explicit adult/youth/unknown classifier and verify M/F, mA–mE/wA–wE, SPF, GE and malformed labels with synthetic fixtures.
- [x] 1.2 Define Kasse capacity one, exactly one Ordnungsdienst and Reinigung capacity two, plus per-game required-position sets; verify adult games require eight positions and youth/unsupported classes retain the baseline five.
- [x] 1.3 Add a guarded Alembic revision to rename current Unterstützung assignments to Kasse, preflight collisions and preserve old audits; verify populated migration, collision refusal, identity and snapshot recovery.
- [x] 1.4 Update schema fingerprints, fixtures and role compatibility handling; verify supported startup revisions, no model/head drift and no misrouting through the old ROLE_CLEANING alias.
- [x] 1.5 Document role names, capacities, category rules, Kasse's additional ordering/security work and retained-assignment migration behavior in README.MD; verify unsupported classes preserve baseline staffing and old audit labels remain readable.

## 2. Assignments and schedule

- [x] 2.1 Generalize role/position lookup, validation, CLI handling and on-demand candidates beyond Verkauf; verify both Reinigung positions support independent claim/release, conflicts and one-task-per-game enforcement; refuse new adult-only duties in youth/unknown games using serialized saved category while preserving authorized releases.
- [x] 2.2 Render eight adult or five youth/unknown offered positions plus release-only retained assignments, without optional markers; verify category markup, exact position labels, guest redaction, release-only candidates and access-tier refusals.
- [x] 2.3 Drive server/client progress from required positions rather than a hard-coded Unterstützung exclusion; verify eight-position adult and five-position youth progress, retained removed-duty releases and stale-write behavior.
- [x] 2.4 Update README schedule and CLI examples for repeated roles, category availability and retained removed assignments; verify documented role inputs and displayed controls match implementation.

## 3. Reporting and notifications

- [x] 3.1 Apply per-game required positions to open-duty statistics, completeness and MV notifications; verify adult-only gaps, no removed youth duties in gaps and the required sole youth Ordnungsdienst.
- [x] 3.2 Keep all retained removed-duty assignments in personal statistics and ordinary helper reminders and resolve repeated-role recipients by actual assignment; verify both position occupants receive the correct single duty context without dropped or duplicated recipients.
- [x] 3.3 Document adult/youth progress, reminders and personal-count behavior in README.MD; verify the examples agree across schedule, statistics and notifications.

## 4. Integration verification

- [x] 4.1 Run `openspec validate update-game-duty-staffing --strict` and `test/run_tests.sh`; resolve integration failures and verify both debug date switches remain false.
