# Proposal

## Why

Synchronization creates preparation, cake-delivery and cleanup blocks for every non-empty
date value, so a postponed game with the placeholder "offen" gets day tasks for a day that
does not exist. They have no calculable lock or meaningful time, they are offered to
helpers, and they appear as missing duties.

Separately, when the last game leaves a calendar date that has occupied blocks, the whole
daily synchronization fails and keeps failing until someone intervenes, so no other game
is updated and helpers keep duties for a day without games.

## What Changes

- Blocks are created only for calendar dates. Empty, placeholder and unparseable date
  values produce no blocks; the games keep their game duties.
- Synchronization removes existing empty open-date blocks.
- Occupied open-date blocks are retained so no commitment is cancelled. They no longer
  fail synchronization, stay visible even without a game on that value, and are removed by
  the first synchronization after they are empty.
- Retained blocks offer no new work: free positions cannot be claimed, cake settings
  cannot be changed, and free positions are not reported as missing. Replacement requests
  and takeovers of occupied positions keep working.
- When a game receives a calendar date, that date gets its blocks under the ordinary rules.
  Helpers of a retained block do not move to it.

- **BREAKING (domain rule):** when a scrape contains no game for a current or future
  calendar date, synchronization removes that date's blocks and assignments instead of
  failing, including locked duties and open replacement requests. This is a deliberate
  exception to "destructive maintenance may not cancel binding commitments".
- Each removed helper is notified after saving through the ordinary channel preference,
  and each removal gets a system `date_remove` audit entry.
- Dates before today keep occupied blocks, so completed work and statistics stay intact;
  they no longer fail synchronization either.
- An empty scrape for a season with stored games fails before any change, because a broken
  scrape would otherwise cancel every block duty of the season.

Non-goals: moving block assignments to a new date, restoring them when a date returns,
removing game duties, resending failed notifications, or automatically removing occupied
open-date blocks.

## Capabilities

### New Capabilities

_None._

### Modified Capabilities

- `game-day-task-blocks`: blocks require a calendar date; adds the handling of retained
  open-date blocks and of cancelled home-game dates.
- `assignment-audit`: adds the `date_remove` transition and replaces the block-removal
  refusal.
- `task-self-service`, `age-eligibility`, `game-task-staffing`: the destructive-maintenance
  refusal gains the cancelled-date exception.

## Impact

- `nuligahelper/db/sync.py` (`SyncEvents`, cancellation), `db/day_blocks.py`
  (`reconcile_day_blocks`), `db/block_duties.py` (claim and cake guards).
- `nuligahelper/jobs/daily.py`, `jobs/notifier.py`, `jobs/messages.py`: cancellation
  notification.
- `nuligahelper/web/views.py` (`build_schedule`, `block_view`),
  `web/statistics.py` (`build_statistics`).
- No schema migration: the first synchronization after deployment cleans up empty blocks
  and cancels occupied blocks of dates that have no games at that point.
- Docs: `AGENTS.md` day-block and destructive-maintenance rules,
  `docs/benutzer/BENUTZERHANDBUCH.md`,
  `docs/betrieb/deployment.md` transition note.
- Administrators empty a retained block through `admin-reopen-duties`. Until that change
  is deployed, occupied retained blocks simply stay, which is safe.
