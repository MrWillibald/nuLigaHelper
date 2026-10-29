# Tests

Automated tests for the nuLigaHelper database, notification and web layers.
All tests run offline against throwaway SQLite databases (sample game plan from
`helpers.py`, settings from `config_template.json`, wording from `messages.py`) – no real mails/SMS are
sent and neither `config.json` nor `nuliga_helper.db` are touched.

## Run

```bash
# all tests, standalone runner (no extra dependencies)
./venv/bin/python test/test_db.py
./venv/bin/python test/test_notifier.py
./venv/bin/python test/test_webapp.py

# or everything at once via the helper script
test/run_tests.sh

# or with pytest (nicer output, selective runs)
./venv/bin/pip install pytest
./venv/bin/python -m pytest test/ -v
```

The optional `browser_disclosures.mjs` check uses an already installed Playwright
module and Chromium browser; it installs no packages. It serves only a synthetic
app on loopback and checks delayed scripts, native fallbacks and a larger roster:

```bash
PLAYWRIGHT_MODULE=/absolute/path/to/playwright/index.mjs \
  BROWSER_EXECUTABLE=/absolute/path/to/chromium node test/browser_disclosures.mjs
```

## Test groups

Each `test_*.py` file is self-contained and named after the behavior it covers; the
groups below describe where to look. `helpers.py` provides import paths, temporary
databases, the synthetic game plan and the mini runner; `schema_fixtures.py`,
`schedule_progress_fixture.py`, `disclosure_markup.py` and `fixtures/` hold shared
synthetic data.

| Area | Files |
| --- | --- |
| Domain and sync | `test_db.py`, `test_scraper.py`, `test_game_duties.py`, `test_binding_assignments.py`, `test_replacement_cancellation.py`, `test_day_blocks.py`, `test_cake_blocks.py`, `test_memberships.py`, `test_contacts.py`, `test_age_*.py`, `test_person_birth_dates.py`, `test_spielfest_presentation.py`, `test_task_descriptions.py` |
| Concurrency and SQLite runtime | `test_concurrency.py`, `test_cake_concurrency.py`, `test_sqlite_runtime.py`, `test_daily_lock.py` |
| Web interface and access | `test_webapp.py`, `test_auth.py`, `test_auth_abuse.py`, `test_refusals.py`, `test_management_ui.py`, `test_schedule_*.py`, `test_sticky_web.py`, `test_day_blocks_web.py`, `test_cake_blocks_web.py`, `test_mobile_*.py`, `test_action_feedback.py`, `test_javascript.py` (runs the `js_*.mjs` checks) |
| Daily job and notifications | `test_main.py`, `test_development_overrides.py`, `test_notifier.py`, `test_cake_reminders.py`, `test_messages.py`, `test_message_config.py`, `test_message_accounts.py`, `test_provider_env.py` |
| Schema migrations | `test_schema_migrations.py`, `test_*_migration.py` |
| Backup, CLI and operations | `test_backup.py`, `test_cli.py`, `test_operations*.py`, `test_host_operations.py`, `test_privacy.py`, `test_production_runtime.py`, `test_recovery_check.py`, `test_deployment_assets.py`, `test_release_prepare.py` |
| Optional browser checks | `browser_*.mjs` (need a preinstalled Playwright/Chromium, see below) |

## Notes

- The webapp tests build on each other and run top to bottom within their file
  (like a user clicking through the interface once).
- Databases live in the system temp directory and are recreated on every run;
  migration tests build only synthetic baseline and versioned databases.
- Eligible baseline helpers carry an explicit synthetic adult birth date from
  `helpers.ADULT_BIRTH_DATE`; legacy migration and unknown-date scenarios retain
  null dates. New registration/creation payloads supply dates without reading
  production configuration.
- Dropbox, scraper, mail and SMS behavior is faked; backup tests use only local
  synthetic SQLite files and fake Dropbox clients.
- File-backed test databases use the same WAL/foreign-key/busy-timeout profile as
  the application so contention tests exercise the supported runtime behavior.


The optional `browser_binding_assignments.mjs` check uses the same preinstalled
Playwright/Chromium runtime and a synthetic loopback-only Flask fixture. It verifies
native confirmation, Cancel/Escape/focus return, every duty container, replacement
requests and atomic takeovers, concurrent claims, immediate late locking and the
non-JavaScript confirmation path. It installs nothing and never uses real providers:

```bash
PLAYWRIGHT_MODULE=/absolute/path/to/playwright/index.mjs \
  BROWSER_EXECUTABLE=/absolute/path/to/chromium node test/browser_binding_assignments.mjs
```

The optional `browser_schedule_progress.mjs` check covers horizontal schedule bars,
German occupied/required counts, rounded percentages, separate warnings and native
keyboard/touch disclosures at 320, 360, 390, 700, 900, 901 and 1280 CSS pixels,
with normal and doubled text. It uses `test/schedule_progress_fixture.py`, a
throwaway synthetic database and a loopback server; no production settings or
providers are used. Partial/empty/full staffing, long names, classification and
eligibility warnings, cake setup/zero/positive states and past cards are included.
It also checks delayed scripts, reduced motion, disclosure defaults and statistics
badges. Like the checks above, it requires preinstalled Playwright and Chromium:

```bash
PLAYWRIGHT_MODULE=/absolute/path/to/playwright/index.mjs \
  BROWSER_EXECUTABLE=/absolute/path/to/chromium \
  BROWSER_REPORT_PATH=/tmp/progress-layout.json \
  BROWSER_SCREENSHOT_DIR=/tmp/progress-screens \
  node test/browser_schedule_progress.mjs
```

Optional `PROGRESS_REFERENCE_DIR` points to captured reference `schedule.html` and
`style.css` files (for example from `1859deb`, before the ring layout). Those assets
render against the same saved data for comparison; there is no height reduction
target. Implementation evidence is retained in
`openspec/changes/archive/2026-10-24-restore-schedule-progress-layout/verification.md` and `verification/`.

The affected dependency-free JavaScript regressions also run in `test/run_tests.sh`:

```bash
node test/js_assignment_options.mjs
node test/js_schedule_progress_events.mjs
node test/js_candidate_loading.mjs
node test/js_candidate_toggle.mjs
node test/js_cake_blocks.mjs
node test/js_assignment_confirmation.mjs
node test/js_age_replacement.mjs
node test/js_replacement_cancellation.mjs
```


The optional `browser_schedule_visibility.mjs` check uses a synthetic database
and loopback server to verify visibility settings, native form saves/clearing,
keyboard links/disclosures, and schedule/statistics navigation with JavaScript
both disabled and enabled. It checks 320/360 CSS pixels and doubled text at 640
CSS pixels, including page overflow and control bounds. No provider calls or
production configuration are used. Reports and screenshots default to
`/tmp/nuliga-visibility-browser` (`BROWSER_REPORT_DIR` overrides this).

```bash
PLAYWRIGHT_MODULE=/absolute/path/to/playwright/index.mjs \
  BROWSER_EXECUTABLE=/absolute/path/to/chromium \
  node test/browser_schedule_visibility.mjs
```
