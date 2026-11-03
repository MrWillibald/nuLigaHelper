# Proposal

## Why

`db.py` has 2,434 lines and holds every domain concern at once: role and task constants, SQLite engine/PRAGMA setup and startup verification, ORM models, scraped-game synchronization, persons/teams/memberships and registration, the binding-duty state machine (claim, release, replacement request/cancel/takeover, day-before lock), day blocks and cake configuration, staffing and age-eligibility evaluation, schedule visibility and audit snapshots. The rule "business rules belong in `db.py`" is right, but finding a rule in one 2,400-line file is slow, and every change touches the same file. This is phase 2 of `docs/REFACTORING-VORSCHLAG.md`.

## What Changes

- Turn `db.py` into the package `db/` with the same import name. `db/__init__.py` re-exports the complete current public API, so `import db`, `db.claim_slot(...)`, `migrations/env.py` (`db.Base.metadata`) and all callers stay unchanged.
- Split into topic modules: `roles.py`, `engine.py`, `models.py`, `visibility.py`, `sync.py`, `people.py`, `clock.py` (cutoff/lock time), `audit.py`, `duties.py` (game duty state machine), `day_blocks.py`, `staffing.py`. `age_eligibility.py`, `contact_validation.py` and `task_labels.py` remain separate top-level modules in this change; `task_labels.py` exists so that `age_eligibility.py` can use role labels without importing `db`, and folding it into `db/roles.py` would create an import cycle.
- Establish and document the patching rule: modules call `clock.duty_now()` and `common.effective_today()` through their module objects; tests patch `db.clock.duty_now` instead of `db.duty_now`.
- No rule, error message, constraint, transaction boundary or retry behavior changes.

## Capabilities

### New Capabilities

None.

### Modified Capabilities

None. Pure refactoring; observable behavior and the database schema are unchanged.

## Impact

- `db.py` → `db/` package (≈12 files).
- Tests: 17 `patch("db.duty_now", …)` call sites change their target to `db.clock.duty_now`; `patch.object(db, …)` on functions used only by external callers keeps working through the facade.
- `migrations/env.py` and Alembic metadata comparison must see the identical `Base.metadata` (verified by the existing drift tests).
- `AGENTS.md` and `CLAUDE.md` describe the new module map; the rule "business rules belong in `db`" is kept.
