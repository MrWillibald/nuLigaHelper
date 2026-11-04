# Proposal

## Why

Occupied duties can currently only change hands through a replacement takeover before the
day-before cutoff; nobody can empty a slot. When a helper drops out at short notice, or an
assignment was simply wrong, the administrator has no way to correct the schedule: the
name stays on the duty, the slot is not shown as open and nobody else can claim it.

## What Changes

- **BREAKING (domain rule):** administrators can switch an occupied game or day-block duty
  back to open, before and after the 09:00 day-before cutoff. This is a deliberate,
  admin-only exception to "no tier may release a locked duty" and "assignments never
  create a vacancy". All other tiers keep the existing rules unchanged.
- The operation is called `Auf offen setzen`. It requires a German `Bestätigen` /
  `Abbrechen` confirmation, also without JavaScript.
- Reopening empties the slot, closes any open replacement request and compares the
  expected occupant so that a stale page cannot remove a newer helper.
- Duties dated before today are refused, so reopening cannot rewrite completed work and
  statistics. Duties with an open or unparseable date are permitted.
- The reopened slot behaves like any empty slot: it is claimable, and a claim after the
  cutoff is immediately locked.
- The removed helper is notified through the ordinary channel preference. If the helper
  cannot be reached, the slot stays open and the administrator is warned.
- A new distinguishable `reopen` audit action names the administrator and removed helper.

Non-goals: a CLI command, reopening for MVs or members, bulk reopening, changing past
duties, changing ordinary release, or changing the cutoff itself.

## Capabilities

### New Capabilities

_None._

### Modified Capabilities

- `task-self-service`: adds administrator reopening and helper notification; narrows the
  admin and lock rules accordingly.
- `assignment-audit`: adds the `reopen` transition.
- `action-feedback`: adds the reopening confirmation and the control on locked duties.
- `game-day-task-blocks`: block duties follow the same exception; a reopened cake position
  can be removed by a later quantity reduction.

## Impact

- `nuligahelper/db/duties.py`, `db/block_duties.py`, `db/slots.py`: `reopen_slot()` and
  `reopen_block_slot()`.
- `nuligahelper/web/assignments.py`, `web/views.py`, `web/static/app.js`,
  `web/templates/schedule.html`, `assignment_form.html`: new JSON endpoints
  `/api/assignment/reopen` and `/api/block-assignment/reopen`, control, confirmation.
- `nuligahelper/jobs/messages.py`, `jobs/notifier.py`: new notification templates.
- No schema migration: `reopen` is a new value of the existing audit action column.
- Docs: `AGENTS.md` binding and lock rules, `README.md`, `docs/benutzer/` handbooks.
