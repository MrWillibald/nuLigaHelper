# Proposal

## Why

Occupied game duties lock at 09:00 Europe/Berlin on the day before the game, and a saved
lock is currently kept forever, even when nuLiga later postpones the game to a later date
or takes its date away ("offen", empty or otherwise unparseable). A postponed game (often
only discovered after its original date) therefore keeps helpers frozen for weeks: they
can no longer search for a replacement or be taken over, although the game is far from
any cutoff and the duty could be handled through the ordinary binding replacement rules.

## What Changes

- A saved duty lock is only kept while the game's **current** date justifies it. On
  every synchronization, each stored game of the synchronized season is checked: if its
  valid current date has no saved cutoff marker, its 09:00 day-before cutoff is not yet
  due and lies after the saved lock time (i.e. the game was moved to a later date), the
  saved locks of its occupied duties are released.
- Because the rule is evaluated on the stored state rather than on a detected date
  change, it also repairs **games that were already moved before this release** and
  still carry falsely saved locks: the first daily synchronization after deployment
  releases them, with audit evidence, without a separate migration or operator command.
- Saved locks are also released for every game whose current date is **open or
  unparseable** (empty or a placeholder such as "offen" when nuLiga postpones a game
  without a new date): such a game has no computable cutoff. This uses the same
  per-sync check, so games that already carry an open date are released on the first
  sync after deployment as well.
- Released duties return to the normal pre-cutoff binding state: the helper stays
  assigned and responsible, may open a replacement request, and eligible replacements may
  take over atomically. Existing age rules for unparseable game dates still apply to
  claims and takeovers.
- Each released lock appends one system `unlock` lifecycle audit entry for the unchanged
  responsible helper.
- Locks are **kept** when the game moves earlier between two valid dates, only its start
  time changes on the same date, or the new valid date's cutoff is already due or saved.
  As soon as a valid date is known, its cutoff locks the duties again through the
  existing mechanism.
- Replacement requests that were closed by the original lock stay closed; no request,
  assignment or vacancy is recreated, and nothing is released or reassigned automatically.
- Supersedes the current rule "retain any saved lock after game date/category changes"
  for postponements and open dates; category changes and all other cases keep their lock.

Non-goals: day-block duties (blocks belong to a valid date and never move with a game),
changing rescheduling notification wording, any admin/manual unlock, or any change to
the 09:00 day-before cutoff itself.

## Capabilities

### New Capabilities

_None._

### Modified Capabilities

- `task-self-service`: adds requirements that postponing a game to a later, not-yet-due
  date, or a game having an open/unparseable date, releases the saved locks of its
  occupied duties back to ordinary binding rules.
- `assignment-audit`: the lifecycle transitions recorded for assignments gain a
  distinguishable system `unlock` transition for postponement and open-date releases.

## Impact

- `nuligahelper/db/sync.py` (`sync_games`) and a new helper in `nuligahelper/db/locks.py`
  that releases saved game-duty locks whose game date no longer justifies them.
- `nuligahelper/db/clock.py` is reused unchanged (`duty_is_locked`, `duty_cutoff_at`).
- Audit page shows the new `unlock` action; no schema migration (only `locked_at` is
  cleared on existing rows).
- Release transition: no schema migration; the first daily run after deployment
  releases falsely locked duties of already moved games. `docs/betrieb/deployment.md`
  gets a one-time transition note on reviewing the resulting `unlock` audit entries.
- Tests in `tests/db/` (sync + binding assignments); docs: `AGENTS.md` domain rule on
  retained locks, `docs/benutzer/BENUTZERHANDBUCH.md` audit action list.
