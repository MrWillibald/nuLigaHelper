# Design

## Context

`db.py` is imported by `webapp.py`, `main.py`, `manage_db.py`, `notifier.py`, `auth_abuse.py`, `operations.py`, `privacy.py`, `migrations/env.py` and almost every test. Tests patch `db.duty_now` (17×), `db.claim_slot`/`db.release_slot` (concurrency tests), `db.get_all_teams` and `db.logging.info`. Alembic autogenerate comparison and `schema_migrations.verify_head_schema()` rely on `db.Base.metadata` being complete once `db` is imported.

## Goals / Non-Goals

**Goals:** each domain topic in its own file of manageable size; unchanged public API through `db`; unchanged behavior, SQL and schema; explicit, testable patch points for time.

**Non-Goals:** no service/repository layer, no renaming of public functions, no moving `db` into the `nuligahelper/` package (that is `restructure-project-layout`), no rule changes.

## Decisions

### 1. Package facade with explicit re-exports

`db/__init__.py` imports the submodules in dependency order and re-exports public names with an explicit `__all__`. Private helpers (`_lock_game`, `_audit_assignment`, …) stay private to their submodule; if a test uses one, it is re-exported deliberately and listed in a comment. A test compares `dir(db)` public names before/after the split to prove nothing was dropped.

### 2. Dependency direction

`roles` → `engine` → `models` → (`visibility`, `audit`, `clock`) → (`people`, `sync`) → `staffing` → `duties` → `day_blocks`. Lower modules never import higher ones. Where today's code relies on late binding across this order (e.g. `offered_positions()` annotated with `Game` before the class exists), use `from __future__ import annotations` or `TYPE_CHECKING` imports, never runtime cycles.

### 3. Time is patched in one place

`clock.py` owns `duty_now()`, `duty_cutoff_at()`, `duty_is_locked()`. All submodules call `clock.duty_now()` via the module object. `db.duty_now` is still re-exported for callers, but patching it would no longer affect domain code, so the 17 test patches move to `db.clock.duty_now` in the same commit as the move. `common.effective_today()` is likewise always called as `common.effective_today()`.

### 4. Mutation entry points stay where callers expect them

Functions that tests replace on the facade to simulate races (`db.claim_slot`, `db.release_slot`) are only replaced for callers outside `db` (webapp/CLI). That remains true after the split because those callers go through the facade. Inside `db`, calls between submodules use the submodule attribute (`duties.claim_slot`) so that a facade patch never changes internal behavior unexpectedly.

### 5. Extraction order

From the edges inward, one commit per step with the full suite green: `roles` (+ `task_labels`) → `engine` → `models` → `visibility` → `sync` → `people` → `clock`/`audit` → `day_blocks` → `duties` → `staffing`.

## Risks / Trade-offs

- **Silent patch misses** (a test still patching `db.duty_now` passes vacuously) → add a guard test that patches `db.clock.duty_now` to a sentinel time and asserts a lock decision changes; grep that no `patch("db.duty_now"` remains.
- **Metadata completeness** → existing schema-drift and fresh-initialization tests run unchanged; add an assertion that `db.Base.metadata.tables` equals the pre-split table set.
- **SQLAlchemy relationship strings** resolve through the shared registry, so models must all live in `models.py` (or be imported before mapper configuration).
