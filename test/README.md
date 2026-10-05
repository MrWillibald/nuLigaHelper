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

## Files

| File               | Scope                                                                    |
|--------------------|--------------------------------------------------------------------------|
| `helpers.py`       | Shared setup: import paths, temp databases, sample games, mini runner    |
| `test_db.py`       | Bootstrap, sync events (new/shift/§77/removed), ordering, cascades       |
| `test_day_blocks.py` | Day-block lifecycle, timing, CAS, audit and optional game duties       |
| `test_day_blocks_web.py` | Block authorization, privacy, filtering, statistics and audit UI |
| `test_cake_blocks.py` | Variable cake capacity, setup/zero settings, safe reductions and durable dated history |
| `test_cake_concurrency.py` | Cake claim/configuration races, stale snapshots and bounded write retries |
| `test_cake_blocks_web.py` | Admin settings, cake authorization, guest privacy, filtering, candidates and reporting |
| `test_cake_block_migration.py` | Cake seeding, exact schema fingerprints, retained data and snapshot recovery |
| `test_cake_reminders.py` | Weekly/day-before one-cake messages, contact preference and debug suppression |
| `test_notifier.py` | Mail/SMS dispatch counts and texts (recorded, never sent)                |
| `test_messages.py` | Named contracts, channel variants, literal values and safe rendering failures |
| `test_message_config.py` | Non-mutating legacy preflight, safe warnings and recipient precedence |
| `test_webapp.py`   | Schedule rendering, inline assignment API, persons CRUD, statistics      |
| `test_sqlite_runtime.py` | SQLite WAL, foreign-key, timeout and startup invariants             |
| `test_production_runtime.py` | Production config and proxy/host/cookie/body/header behavior |
| `test_concurrency.py` | Game/block assignment CAS, audit atomicity and bounded WAL contention  |
| `test_schema_migrations.py` | Fingerprints, snapshots, Alembic upgrades and migration preflights |
| `test_age_eligibility.py` | Calendar validation, game-date birthdays, timing minimums and collective sale enforcement |
| `test_person_birth_dates.py` | Required collection and private self/admin correction |
| `test_birth_date_migration.py` | Guarded migration, unknown legacy dates, retained roster/history and failure recovery |
| `test_age_reporting.py` | Saved-state eligibility warnings, truthful occupancy, statistics and existing MV follow-up |
| `test_daily_lock.py` | Per-database process lock for non-overlapping daily runs                 |
| `test_backup.py`   | Online snapshots, Dropbox retention, staged failures and safe restore     |
| `test_main.py`     | Daily orchestration order, transaction boundaries and failure status      |
| `test_cli.py`      | Management commands including guarded snapshot restoration                |

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
