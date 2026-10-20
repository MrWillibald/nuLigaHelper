# Proposal

## Why

A saved helper assignment is a commitment, but the current release flow can leave a required duty empty without a replacement. Helpers need to understand that commitment before saving and have a visible way to seek an eligible replacement while retaining responsibility.

## What Changes

- **BREAKING:** Replace ordinary withdrawal with `Ersatz gesucht` on the occupied assignment. The original helper stays assigned and responsible until an authorized, eligible replacement takes over atomically.
- Require a German pre-save confirmation with `Bestätigen` and `Abbrechen` for every UI flow creating an assignment. Selecting a helper or opening the dialog performs no write; cancel/dismissal performs no write. Confirmation still uses current authorization, eligibility and compare-and-swap checks.
- Lock occupied duties at 09:00 Europe/Berlin on the day before the duty, close outstanding replacement requests and retain their helpers, even if the reminder job is missed or delivery fails. Weekly reminders do not lock. After locking, the web service offers no occupied-duty release, replacement or change, including for administrators.
- Keep genuinely empty offered positions claimable after the boundary; successful late claims are immediately locked. Apply the lifecycle to every offered game duty, preparation, cleanup and configured cake position.
- Preserve existing staffing counts, age deficiencies, privacy, scope of tier permissions, and append-only history. Explicitly reconcile legacy release-only duties and operational cleanup paths instead of allowing accidental lock overrides.
- There is no 24-hour replacement window, timer, automatic return or periodic rollback. `Ersatz gesucht` is visible status; this change adds no replacement broadcast or external integration.

## Capabilities

### New Capabilities

None; extend existing assignment capabilities.

### Modified Capabilities

- `task-self-service`: binding confirmation, replacement requests, atomic takeover, cutoff, per-container permissions and compatibility for retained duties.
- `age-eligibility`: validate the resulting takeover group and remove unconditional withdrawal guarantees.
- `game-task-staffing`: preserve offeredness and reporting while qualifying release-only legacy maintenance.
- `game-day-task-blocks`: apply the same lifecycle to every phase and reconcile reminder, removal and cake configuration rules.
- `schedule-overview`: persistent replacement/locked state, truthful occupancy, accessible pre-save confirmation and private guest presentation.
- `access-control`: current candidate/write permissions for replacement and locked duties without expanded tier rights.
- `assignment-audit`: record request transitions, lock transitions and linked responsibility transfers atomically.
- `action-feedback`: pre-save cancellation and truthful atomic takeover outcomes instead of partial release-then-claim feedback.
- `user-accounts`: refuse deactivation or erroneous-record deletion when it would cancel a binding commitment, preserving assignment/account state atomically.
- `sqlite-runtime-safety`: serialize requests, takeovers and lock enforcement with existing claims/configuration.
- `schema-migrations`: explicit guarded migration preserving saved assignments and audit history.

## Impact

Later implementation affects `db.py`, Alembic revisions and schema verification, `main.py` reminder orchestration, `notifier.py`, Flask assignment/candidate routes, Jinja/JavaScript/CSS, synthetic offline tests, `README.MD` and `AGENTS.md`. Any later notification wording remains in `messages.py`. Existing SQLite, Flask and notification providers remain sufficient; no new service is proposed.

The user explicitly confirmed a fixed cutoff at 09:00 Europe/Berlin on the day before the duty, independent of missed jobs, missing contacts and failed delivery. The user also confirmed refusal of destructive maintenance when it would cancel binding commitments, including account deactivation/deletion and removal of an occupied last-game date or block. Authorized release-only retirement of obsolete retained roles remains possible before cutoff and is locked afterward. These confirmed decisions supersede the earlier actual-batch/default-maintenance design; no general administrator release or replacement override is introduced. The planning artifacts are ready for the separately authorized implementation workflow.
