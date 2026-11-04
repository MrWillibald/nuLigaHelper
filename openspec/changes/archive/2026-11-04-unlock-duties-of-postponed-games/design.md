# Design

## Context

- A game duty is locked if its assignment has `locked_at` set, a `DutyCutoff` marker
  exists for the container's `(season_year, date)`, or the clock is past
  `duty_cutoff_at(date)` (`db/clock.py: duty_is_locked`). Only the first condition is
  tied to the assignment; the other two follow the container's current date.
- `lock_day_before_batch()` (`db/locks.py`) sets `locked_at` on every occupied
  assignment of a dated cohort and audits `lock`; late claims after cutoff get
  `locked_at` immediately (`db/duties.py`).
- `sync_games()` (`db/sync.py`) serializes against claims, then for every existing game
  calls `enforce_due_duty_locks(session, game)` *before* overwriting `date`/`time`, so a
  due lock on the old date is persisted first. Nothing ever clears `locked_at`, which is
  why a postponed game stays frozen (see proposal.md).
- An open/unparseable date has no cutoff: `duty_cutoff_at()` returns `None`,
  `lock_day_before_batch()` refuses such a date and `lock_due_duties()` skips it, so no
  `DutyCutoff` marker can exist for it. Only a saved `locked_at` keeps such a game locked.
  Age rules already treat an invalid game date explicitly (timing/Ordner/non-youth
  Verkauf claims are refused), which also governs takeovers after an unlock.
- Every saved `locked_at` equals the cutoff of the date it was locked on: the batch
  copies `DutyCutoff.locked_at` (the 09:00 day-before cutoff, naive Berlin time) and a
  late claim uses `_claim_lock_time()` (saved marker or due cutoff). It is therefore
  never later than the moment of locking. A lock whose time is earlier than the current
  date's cutoff, while that cutoff is not yet due, can only come from an earlier date:
  the game was moved later. Production refuses date overrides, so a rolled-back clock
  cannot produce such locks there.
- Games postponed before this release already carry such stale locks; the database
  holds no record of the original date other than `locked_at` and audit snapshots.
- Requirements: `specs/task-self-service/spec.md` and `specs/assignment-audit/spec.md`
  of this change.

## Goals / Non-Goals

**Goals:**
- Release `locked_at` whenever the game's current stored date does not justify it
  (postponed to a later not-yet-due date, or open/unparseable), inside the sync
  transaction, with one `unlock` audit per released row and no entries on idempotent
  repeats.
- Repair games postponed before this release through the same rule, without a separate
  migration or operator command.

**Non-Goals:**
- No schema change, no new column, no change to `DutyCutoff` markers (they are per date
  and the old date keeps its marker for any other games on it).
- No unlock path in the web UI or CLI; no change to day blocks or notification texts.
- No Alembic revision: the rule depends on the duty clock, and schema revisions must
  stay clock-independent and run only with stopped writers.

## Decisions

1. **Unlock is a state-based check in `sync_games`, after the scraped fields are applied
   and before reconciliation, for every stored game of the synchronized season**
   (including stored games missing from the scrape). It runs under the sync's existing
   write serialization and in its transaction. Because it compares the saved lock with
   the game's *current* date instead of detecting a date change in this run, the same
   code repairs games moved before this release and games stored with an open date on
   the first daily sync after deployment. Alternative: compare old and new scraped date
   only. Rejected: it cannot see moves that happened before deployment. Alternative: a
   one-off CLI command or data migration. Rejected: a second code path with the same
   rule, an extra operator step, and migrations must stay clock-independent.
   Alternative: derive "unlocked" at read time by
   comparing `locked_at` with the current date's cutoff. Rejected: it would silently
   change the meaning of every saved lock, leave no audit evidence and spread the rule
   across every reader.

2. **Keep the existing `enforce_due_duty_locks` call before the date update.** If the
   old cutoff was due but not yet persisted, the lock (and closure of an open request)
   is recorded first and then released. The audit trail honestly shows that the
   original cutoff passed; closing the request at that cutoff is the existing rule.

3. **Unlock condition** (new helper `locks.release_unbound_game_locks(session, game)`),
   evaluated on the game's current stored date, per assignment with `locked_at` set
   (refreshed, like `lock_day_before_batch`):
   - `parse_game_date(game.date) is None` (open/unparseable) → release; or
   - the date is valid, `duty_is_locked(session, game)` without an assignment is
     `False` (no saved `DutyCutoff` for that date, clock before its 09:00 day-before
     cutoff), and `assignment.locked_at < duty_cutoff_at(game)` (naive Berlin
     comparison) → release;
   - otherwise (due or saved cutoff, earlier move, same date) → keep.
   The `locked_at` comparison is redundant in production but keeps the rule explicit
   and refuses to drop a lock that does not demonstrably belong to an earlier date.
   On release: clear it and append `_audit_assignment(..., "unlock", None,
   "system")`. Replacement tokens are not touched (already `None` after a lock).
   Reusing `duty_is_locked` keeps one definition of "due" and honours clock patches in
   tests (`db.clock.duty_now`). Only rows with `locked_at` set are touched, which makes
   the per-sync check idempotent. Alternative: compare only `locked_at` with the new
   cutoff. Rejected: it would ignore a saved marker for the new date.

4. **Audit action name `unlock`**, system actor, mirroring `lock`. The audit page
   renders action keys literally, so no web change is needed beyond documentation.

5. **Retained removed duties are treated like any other assignment**: after unlock
   their existing pre-cutoff release-only retirement becomes available again, which is
   consistent with the "before the actual cutoff" rule.

## Risks / Trade-offs

- [A wrong scrape shows a date jump forward and back] → Forward jump unlocks, the move
  back to the original (already due) date re-locks on the next authoritative check via
  `enforce_due_duty_locks`/`lock_due_duties`; both transitions are audited.
- [Helpers might not notice they could seek a replacement again] → Existing shift
  notification already informs assigned helpers of the new date; wording unchanged by
  decision (proposal non-goal).
- [nuLiga briefly shows a placeholder date and then the original, already due date
  again] → The placeholder unlocks, the original date re-locks on the next authoritative
  check; both transitions are audited, the helper never loses the assignment.
- [Ordinary claims/takeovers on an open-date game are limited] → Intended: existing
  invalid-date age rules decide; this change adds no exemption.
- [Day blocks of the old date keep their locks; if the moved game was the only game of
  that date, the existing rule still refuses removal of occupied blocks] → Unchanged
  behaviour, explicitly out of scope.

## Migration Plan

Code-only change; deploy with the ordinary release procedure, no schema migration.

Release transition (documented once in `docs/betrieb/deployment.md`):
1. Before activation, optionally note the current season's games with saved locks whose
   date is open or whose cutoff is still ahead (read-only SQL on a snapshot, IDs only),
   to compare afterwards.
2. The first daily run after the timers resume performs the repair in its sync: every
   falsely locked duty of an already moved or open-date game is released with one
   system `unlock` audit entry; nothing is sent because of the unlock itself.
3. Check the new `unlock` entries on the audit page: only postponed/open-date games
   were affected and their helpers remained assigned.
Games of earlier seasons are not touched (their dates are past and due anyway).

Rollback: redeploy the previous release. Released locks stay cleared; the old release
re-locks them at the new date's cutoff through the existing mechanism, which matches the
intended behaviour. Restoring the pre-activation snapshot is only needed if the release
must be undone completely.
