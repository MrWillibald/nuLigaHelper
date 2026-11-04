# Design

## Context

`release_slot()` exists only to retire obsolete duties and `_validate_retirement()` refuses
offered and locked ones; `release_block_slot()` is a compatibility refusal. The per-slot
compare-and-swap, serialization and retry pattern in `db/duties.py` and
`db/block_duties.py` is the model to follow. Web requests already build a `Notifier`
(`web/challenges.py:_account_notifier`). See proposal.md for motivation.

## Goals / Non-Goals

**Goals:** one narrow audited override, with the lock and binding rules untouched for all
other callers.

**Non-Goals:** any change to `release_slot()`, `_validate_retirement()` or the ordinary
release endpoints.

## Decisions

**Dedicated writers instead of an admin flag on release.** Add `db.reopen_slot()` and
`db.reopen_block_slot()`. A flag on `release_slot()` would put the lock bypass on the path
every tier uses, where one wrong argument cancels the guarantee.

**Authorization lives in `db`.** The writers take the actor, reload it after serialization
and require an active person with `is_admin` and `actor_tier == "admin"`. The web check is
only an early refusal. System callers (`actor=None`) are refused, so the CLI and jobs
cannot reach it by accident.

**Writer order.** Serialize with `_lock_game` / `_lock_day_block`, run
`enforce_due_duty_locks`, compare the expected occupant, refuse past dates, write audits,
delete the row, flush. Enforcing the lock first keeps due-lock evidence in the right order.
If a request is open, `request_close` is audited before `reopen`. Delete through the
pattern used in `release_slot()` and expire the collections, per the SQLAlchemy pitfall.

**Past check uses `common.effective_today()`.** Date strictly before today is refused;
`parse_game_date()` returning `None` means open date and is permitted. A duty on today's
date stays reopenable, which is the match-day case this exists for.

**No Verkauf check on reopening.** Emptying a slot cannot produce an invalid assigned
group, because collective adult coverage is enforced when the last slot is claimed.
Refusing here would block exactly the correction the administrator needs.

**Notification after commit, in the web layer.** The endpoint commits, then calls a new
`Notifier.notify_duty_reopened()` inside a guarded `try`. Sending inside the transaction
could notify about a change that rolls back. The JSON response adds `notified: false` when
`_dispatch()` returns 0 or sending raises; `app.js` and `/dienst` show it as a warning.
Move `_account_notifier()` to a shared web location rather than importing across topic
modules if the structure test requires it.

**Messages.** Add `duty.game.reopened`, `duty.spielfest.reopened` and `duty.block.reopened`
to `jobs/messages.py` with existing field sets where possible. An open date is shown as
stored.

**UI.** `duty_state()` gains `can_reopen` (admin, occupied, not past). Unlike the other
controls it is true on locked duties. `app.js` reuses the `assignment-confirmation`
dialog pattern; `/dienst` accepts `action=reopen` and requires `decision=confirm`, unlike
`release` and `request-replacement`, which skip the decision check.

## Risks / Trade-offs

- [An administrator empties a locked duty by mistake] → Confirmation names the helper, the
  audit records the actor, and the helper can be re-entered through an ordinary claim,
  which is immediately locked.
- [The helper misses the message on match day] → The administrator is warned on failure;
  a delivered message cannot be guaranteed, as with all notifications.
- [A reopened duty stays empty after cutoff] → Accepted: it appears as a vacancy, and the
  existing rule that empty slots stay claimable applies.
- [Weakens the promise that locked duties are stable] → Limited to current administrators
  and made visible through audit and notification.

## Migration Plan

No schema migration; deploy normally. Rollback to the previous release is safe: stored
`reopen` audit rows remain readable because the audit template renders unknown actions
as stored.
