# Tasks

## 1. Code

- [ ] 1.1 Rename `backup._readonly_uri` to public `readonly_uri`, update its internal callers and replace `schema_migrations._readonly_uri` with it; verify `grep -n "_readonly_uri" *.py` is empty and backup/migration tests pass.
- [ ] 1.2 Remove `ROLE_MV` and `ROLE_SUPPORT` from `db.py`; replace their uses in tests with local literals or `db.ROLE_CASH`, keeping the intent of each assertion (no MV task description, historical Unterstützung → Kasse migration).

## 2. Dependencies and tests

- [ ] 2.1 Replace the contents of `requirements.txt` with `-r requirements-test.txt`; update install commands in `README.MD`, `AGENTS.md`, `CLAUDE.md` and the hint in `test/run_tests.sh`; verify a fresh `python3 -m venv venv && ./venv/bin/pip install -r requirements.txt` on CPython 3.13 succeeds.
- [ ] 2.2 Inject a fixed healthy `statvfs` result into the healthy cases of `test_monitor_detects_staleness_missing_timers_clock_disk_and_tls` while keeping the low-disk assertions; verify the test passes independent of the host's free space.

## 3. Verification

- [ ] 3.1 Run `test/run_tests.sh` (expect 0 failures) and `openspec validate tidy-maintenance-helpers`; confirm `send_article` and the dormant templates are unchanged.
