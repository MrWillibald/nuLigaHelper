# Proposal

## Why

The developer-facing documentation has drifted from the code. `CLAUDE.md` still states that there is no migration layer and that schema changes require deleting the database, although Alembic revisions and guarded `migrate-schema` exist. It lists outdated routes and only three pages. The `AGENTS.md` module map names 9 of 25 modules, describes the web interface as LAN-only although the Caddy/Gunicorn/systemd deployment exists, and `run_webapp.sh` points to the private, gitignored `deploy/PRODUCTION.md`. `README.MD` has grown to 758 lines that mix German user guidance with English operator procedures, so newcomers cannot find the quick start. Contributors and coding agents follow these files literally, so wrong statements cause wrong changes.

This is phase 0 of `docs/REFACTORING-VORSCHLAG.md` and the prerequisite for the structural refactorings, which will update the same files.

## What Changes

- Correct `CLAUDE.md`: Alembic-based guarded migrations instead of "delete the `.db` file", current route list (`/api/assignment/claim|release|request-replacement|cancel-replacement|takeover`, `/api/block-assignment/*`, `/api/games/<id>/team`, `/api/games/<id>/candidates`, `/api/blocks/<id>/*`, `/dienst`, `/audit`, login/registration, legal pages).
- Complete the `AGENTS.md` module map with every top-level module and its single responsibility; replace the "local network only" statement with a pointer to the production deployment.
- Fix the `run_webapp.sh` hint to reference `DEPLOYMENT.md` / `release-assets/README.md`.
- Update the outdated header comment of `manage_db.py` ("until the web interface is available").
- Remove duplication between `CLAUDE.md` and `AGENTS.md`: `CLAUDE.md` currently repeats commands and invariants from the canonical `AGENTS.md` (and is the copy that went stale). Reduce it to a short orientation that imports the rulebook with `@AGENTS.md`, keeping only Claude-specific notes.
- Remove the per-file table from `test/README.md`; it lists 24 of 58 test files and cannot stay current. Keep run instructions and describe test groups instead.
- Replace the English behavior chapters in `README.MD` (web interface ~120 lines, home-game-day visibility, task descriptions, game duty positions, binding duties, age eligibility, cake delivery; ~360 lines in total) with a short feature list. Their content already exists in German in the user manuals for users and in `openspec/specs/` as the authoritative behavior contract; operator-relevant details move to `docs/betrieb/`.
- Shorten `README.MD` to overview, features, quick start, testing and links. Move the English operator sections (SQLite production profile, snapshot restore, schema initialization/migration, notification-wording migration, private deployment) unchanged into `docs/betrieb/` files and link them. Keep one language per file.

## Capabilities

### New Capabilities

None.

### Modified Capabilities

None. Documentation-only change; behavior and specifications stay unchanged.

## Impact

- `CLAUDE.md`, `AGENTS.md`, `README.MD`, `test/README.md`, `run_webapp.sh` (comment/echo text only), `manage_db.py` (header comment only), new `docs/betrieb/*.md`, `docs/README.md` index.
- Links in `docs/`, `release-assets/README.md`, `DEPLOYMENT.md` and `DATA-INVENTORY.md` that target moved README sections are updated.
- No Markdown file is orphaned today; `openspec/changes/archive/` (37 archived changes) is kept as the OpenSpec change history.
- No code path, CLI behavior, schema or deployment asset changes.
