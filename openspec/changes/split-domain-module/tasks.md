# Tasks

## 1. Safety net

- [x] 1.1 Record the current public names of `db` and the table set of `db.Base.metadata` in a test that must still pass after the split.
- [x] 1.2 `git mv db.py db/__init__.py`; verify `test/run_tests.sh` (including the schema-drift tests) and `manage_db.py --help` are unchanged.

## 2. Extraction (one commit per step, full suite green after each)

- [x] 2.1 `roles.py` (role/block constants, slot counts, labels, `TASK_DESCRIPTIONS`, offered/required positions); keep `task_labels.py` standalone to avoid an `age_eligibility` ↔ `db` import cycle.
- [x] 2.2 `engine.py` (path resolution, PRAGMAs, `make_engine`, `initialize_db`, `verify_db`, `database_ready`, startup errors).
- [x] 2.3 `models.py` (all ORM classes and the `person_teams` table).
- [x] 2.4 `visibility.py` and `sync.py` (incl. `SyncEvents` dataclasses).
- [x] 2.5 `people.py` (persons, teams, membership, registration, deactivation/deletion, MV).
- [x] 2.6 `clock.py` and `audit.py`; move all `patch("db.duty_now"…)` targets to `db.clock.duty_now` in the same commit and add the guard test from the design.
- [x] 2.7 `day_blocks.py`, `duties.py` and `staffing.py`.

## 3. Finish

- [x] 3.1 Reduce `db/__init__.py` to imports and `__all__`; verify no submodule imports a higher-level one and none exceeds ~600 lines.
- [x] 3.2 Update `AGENTS.md` and `CLAUDE.md` module descriptions and the patching rule.
- [x] 3.3 Run `test/run_tests.sh` and `openspec validate split-domain-module`.
