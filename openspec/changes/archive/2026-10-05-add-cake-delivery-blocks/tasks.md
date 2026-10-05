# Tasks

## 1. Cake block storage and lifecycle

- [x] 1.1 Extend dated block kinds and per-kind capacity with nullable cake quantity/delivery time; verify model tests retain three preparation/cleanup positions and permit configured variable cake positions.
- [x] 1.2 Add a reviewed Alembic revision and update schema fingerprints/fixtures, seeding one unconfigured cake block for each existing dated game group; verify migration identity, uniqueness, unchanged existing assignments/audits and snapshot recovery.
- [x] 1.3 Reconcile cake blocks with successfully synchronized season/date groups, preserving same-date settings and assignments and auditing/logging vanished-date removals; verify repeated sync, moved/removed dates and failed-scrape safeguards.
- [x] 1.4 Document administrator setup and the guarded migration in README.MD; verify the described initial unset settings and stopped-service migration match implementation.

## 2. Configuration and assignments

- [x] 2.1 Add server-enforced admin-only, CSRF-protected time/quantity edits with whole-number validation, current-settings stale checks and atomic capacity updates; verify every tier, invalid values and concurrent edit/claim races.
- [x] 2.2 Append empty positions on increase and reject decreases that would remove occupied positions until explicit audited releases; verify retained position identities, no silent compaction and safe zero-quantity handling.
- [x] 2.3 Extend block claim/release and candidate validation to cake capacity, active people and existing team-independent authority; verify conflicts, one person per cake block, past-date restrictions and duties across different containers.
- [x] 2.4 Document quantity changes and explicit release-before-reduction behavior in README.MD; verify the examples retain existing occupants and history.

## 3. Schedule and public behavior

- [x] 3.1 Add a compact cake card with saved settings, setup-needed/zero states, variable numbered positions and accessible admin configuration; verify four-cake rendering, current progress and keyboard-operable expansion.
- [x] 3.2 Include cake cards in existing date/team/person filters and past sections and retain candidate loading on expansion; verify block-only person matches, responsible-team hiding and candidate-failure behavior.
- [x] 3.3 Preserve guest names-only output and update saved progress after successful position/configuration changes; verify redacted guest responses, no leaked roster data and unchanged progress after refused writes.
- [x] 3.4 Document cake cards and filtering in README.MD; verify documented controls and setup status match the rendered schedule.

## 4. Reminders statistics and history

- [x] 4.1 Add ordinary weekly/day-before cake reminders using saved delivery settings and existing contact/debug rules; verify one-cake meaning, recipients, empty positions and skipped contacts using synthetic providers.
- [x] 4.2 Count occupied cake positions and configured gaps by the semantic role, with a distinct unconfigured setup status and no MV gap reminder; verify personal totals, variable missing counts and zero/unconfigured states.
- [x] 4.3 Extend audit descriptions/review to named cake positions, including release-before-reduction and deletion snapshots; verify readable history after quantity changes and vanished-date removal.
- [x] 4.4 Document reminder timing, contribution counts and setup/gap reporting in README.MD; verify the text matches observable reporting.

## 5. Integration verification

- [x] 5.1 Run `openspec validate add-cake-delivery-blocks --strict` and `test/run_tests.sh`; resolve integration failures and verify both debug date switches remain false.
