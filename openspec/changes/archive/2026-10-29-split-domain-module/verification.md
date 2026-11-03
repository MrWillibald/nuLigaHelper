# Implementation verification

Verified on 2026-10-29 with synthetic databases only.

- Step 1: `test/test_db_structure.py` pins every former non-module name of `db`
  (180 names; the `annotations` future flag excluded) and the 12-table ORM metadata;
  `git mv db.py db/__init__.py` kept the suite green (494 checks) and
  `manage_db.py --help` unchanged.
- Step 2: the AST-based split produced 16 topic modules (largest `duties.py` 456 lines,
  facade `db/__init__.py` 225 lines). An AST comparison of all 148 moved top-level
  definitions against the original file found no difference apart from
  `clock.duty_now` / `storage.SQLITE_TIMEOUT_SECONDS` attribute reads. pyflakes reports
  no undefined or unused names in the submodules; the one remaining message (unused
  `exc` in `release_slot`) already existed in the original.
- Tests: 17 `db.duty_now` patches, the `patch.object(db, "duty_now")` call and the global
  test clock in `helpers.py` now target `db.clock.duty_now`; the cake contention test
  patches `db.storage.SQLITE_TIMEOUT_SECONDS`. New guard tests: a patched
  `db.clock.duty_now` flips `db.duty_is_locked` at 09:00 and is visible through
  `db.duty_now()`; no db module imports `duty_now`/`SQLITE_TIMEOUT_SECONDS` by name; no
  test patches `db.duty_now`; submodules import only lower layers (annotation-only
  `TYPE_CHECKING` imports excepted).
- Runtime: `manage_db.py init` on an empty file, `migrate-schema --confirm-stopped`
  (reports schema head), `list-teams` and `alembic heads` (single head
  `0009_schedule_visibility`) work. A before/after comparison of 44 rendered pages and
  API responses on identical synthetic data was byte-identical except the audit row's
  own timestamp.
- `test/run_tests.sh`: exit 0, 497 passing checks (494 + 3 new guard tests), no failures.
- `openspec validate split-domain-module --strict`: valid.
