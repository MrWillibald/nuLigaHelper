# Tests

Automated tests for the nuLigaHelper database, notification and web layers.
All tests run offline against throwaway SQLite databases (sample game plan from
`helpers.py`, sample texts from `config_template.json`) – no real mails/SMS are
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
| `test_notifier.py` | Mail/SMS dispatch counts and texts (recorded, never sent)                |
| `test_webapp.py`   | Schedule rendering, inline assignment API, persons CRUD, statistics      |
| `test_sqlite_runtime.py` | SQLite WAL, foreign-key, timeout and startup invariants             |
| `test_production_runtime.py` | Production config, proxy/host/cookie/body/header behavior and deployment assets |
| `test_concurrency.py` | Assignment CAS, audit atomicity and bounded WAL contention             |
| `test_daily_lock.py` | Per-database process lock for non-overlapping daily runs                 |
| `test_backup.py`   | Online snapshots, Dropbox retention, staged failures and safe restore     |
| `test_main.py`     | Daily orchestration order, transaction boundaries and failure status      |
| `test_cli.py`      | Management commands including guarded snapshot restoration                |

## Notes

- The webapp tests build on each other and run top to bottom within their file
  (like a user clicking through the interface once).
- Databases live in the system temp directory and are recreated on every run;
  there is no schema-migration logic to test by design (see main README).
- Dropbox, scraper, mail and SMS behavior is faked; backup tests use only local
  synthetic SQLite files and fake Dropbox clients.
- File-backed test databases use the same WAL/foreign-key/busy-timeout profile as
  the application so contention tests exercise the supported runtime behavior.
