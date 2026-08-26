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

## Notes

- The webapp tests build on each other and run top to bottom within their file
  (like a user clicking through the interface once).
- Databases live in the system temp directory and are recreated on every run;
  there is no migration logic to test by design (see main README).
