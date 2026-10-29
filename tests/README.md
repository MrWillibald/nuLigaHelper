# Tests

Automated tests for the nuLigaHelper database, notification and web layers.
All tests run offline against throwaway SQLite databases (sample game plan from
`helpers.py`, settings from `config_template.json`, wording from `nuligahelper/jobs/messages.py`) – no real mails/SMS are
sent and neither `config.json` nor `nuliga_helper.db` are touched.

## Run

`tests/` is a Python package; run test modules from the project root with `-m`:

```bash
tests/run_tests.sh                               # everything (standalone runner)
./venv/bin/python -m tests.db.test_db            # one file
./venv/bin/python -m pytest tests/ -v            # pytest (nicer output, -k selection)
./venv/bin/python -m pytest tests/web -k login   # one area / pattern
```

`test/run_tests.sh` only forwards to `tests/run_tests.sh` for deploy commands installed
before the package layout; remove it together with the root shims.

The optional `browser_disclosures.mjs` check uses an already installed Playwright
module and Chromium browser; it installs no packages. It serves only a synthetic
app on loopback and checks delayed scripts, native fallbacks and a larger roster:

```bash
PLAYWRIGHT_MODULE=/absolute/path/to/playwright/index.mjs \
  BROWSER_EXECUTABLE=/absolute/path/to/chromium node tests/js/browser_disclosures.mjs
```

## Layout

The folders mirror the application package. Each `test_*.py` file is self-contained
and named after the behavior it covers.

| Folder | Covers |
| --- | --- |
| `db/` | Domain rules, sync, duties, day blocks, memberships, contacts, age rules, concurrency and SQLite runtime, db package structure |
| `web/` | Pages, JSON APIs, authentication, refusals, access tiers, mobile presentation, web package structure; `test_javascript.py` runs the `js/js_*.mjs` checks |
| `jobs/` | Daily job orchestration, daily lock, scraper, notifications, message catalog and legacy wording, development overrides, provider settings |
| `migrations/` | Guarded schema upgrades and every revision's retention rules |
| `ops/` | Backup/restore, CLI, monitoring, privacy cleanup, production runtime, recovery check, deployment assets, release preparation, project layout guards |
| `js/` | Dependency-free JavaScript checks (`js_*.mjs`) and optional browser checks (`browser_*.mjs`, need a preinstalled Playwright/Chromium, see below) |

Shared helpers stay at the top: `helpers.py` (synthetic databases, game plan, test
clock, mini runner; import as `from tests import helpers as h`), `schema_fixtures.py`,
`schedule_progress_fixture.py`, `disclosure_markup.py` and `fixtures/`. Tests patch
package modules, e.g. `nuligahelper.common.effective_today` and
`nuligahelper.db.clock.duty_now`; `ops/test_project_layout.py` rejects old module names.

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
  BROWSER_EXECUTABLE=/absolute/path/to/chromium node tests/js/browser_binding_assignments.mjs
```

The optional `browser_schedule_progress.mjs` check covers horizontal schedule bars,
German occupied/required counts, rounded percentages, separate warnings and native
keyboard/touch disclosures at 320, 360, 390, 700, 900, 901 and 1280 CSS pixels,
with normal and doubled text. It uses `tests/schedule_progress_fixture.py`, a
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
  node tests/js/browser_schedule_progress.mjs
```

Optional `PROGRESS_REFERENCE_DIR` points to captured reference `schedule.html` and
`style.css` files (for example from `1859deb`, before the ring layout). Those assets
render against the same saved data for comparison; there is no height reduction
target. Implementation evidence is retained in
`openspec/changes/archive/2026-10-24-restore-schedule-progress-layout/verification.md` and `verification/`.

The affected dependency-free JavaScript regressions also run in `tests/run_tests.sh`:

```bash
node tests/js/js_assignment_options.mjs
node tests/js/js_schedule_progress_events.mjs
node tests/js/js_candidate_loading.mjs
node tests/js/js_candidate_toggle.mjs
node tests/js/js_cake_blocks.mjs
node tests/js/js_assignment_confirmation.mjs
node tests/js/js_age_replacement.mjs
node tests/js/js_replacement_cancellation.mjs
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
  node tests/js/browser_schedule_visibility.mjs
```
