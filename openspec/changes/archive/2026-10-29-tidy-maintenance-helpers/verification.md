# Implementation verification

Verified on 2026-10-29.

- `backup.readonly_uri(path, *, immutable=False)` is the single read-only SQLite URI
  helper; `schema_migrations.py` uses it and its private copy is removed.
  `grep -rn "_readonly_uri" *.py` is empty. `release-assets/recovery_check.py` was left
  standalone as intended.
- `db.ROLE_MV` and `db.ROLE_SUPPORT` are removed; no application module, template or
  script used them. The four tests now assert the intent directly: `db.ROLE_CASH` for the
  historical Unterstützung → Kasse migration and Kasse lookups, a literal `"MV"` for
  "no MV task description", and Wischer ≠ Kasse.
- `requirements.txt` is `-r requirements-test.txt`. A fresh CPython 3.13 venv installed
  it without errors (SQLAlchemy 2.0.52, pandas 3.0.5, Flask, Alembic, phonenumbers,
  email-validator, Twilio, Dropbox, pytest importable). Install commands stay unchanged;
  README and `AGENTS.md` explain the pinned set.
- `test_monitor_detects_staleness_missing_timers_clock_disk_and_tls` uses a synthetic
  `statvfs` result (50 % free) for healthy cases and now also asserts critical (5 %) and
  warning (15 %) disk/inode results, so it no longer depends on the host's free space.
- `Notifier.send_article`, its commented call in `main.py` and the dormant newspaper
  templates in `messages.py` are unchanged (`git diff` empty for those files).
- `test/run_tests.sh`: exit 0, 488 passing checks, no failures.
- `openspec validate tidy-maintenance-helpers --strict`: valid.
