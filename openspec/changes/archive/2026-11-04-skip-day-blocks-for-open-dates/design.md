# Design

## Context

`sync_games()` computes `active_dates` twice from any non-empty scraped date: for the
preflight that refuses removing occupied blocks, and for `reconcile_day_blocks()`, which
repeats the same refusal. If "offen" simply left the active set, an occupied "offen" block
would abort the whole daily sync at both places. `build_schedule()` derives its day groups
from game dates only, so a retained block without a game would disappear from the page.
`fetch_home_games()` does not check the HTTP status and can return an empty list. Stored
games missing from a scrape are kept and only reported. `day_blocks` sits below `audit` in
the `db` layer order, so `reconcile_day_blocks()` cannot write audits. See proposal.md for
motivation.

## Goals / Non-Goals

**Goals:** no new open-date blocks, no failed sync because of open or cancelled dates, and
no commitment cancelled without audit and notification.

**Non-Goals:** any change to game duties, or to refusals for deactivation and record
deletion.

## Decisions

**One definition of "open".** A date is a block date only if
`age_eligibility.parse_game_date()` accepts it, the same test the lock logic uses. A small
`block_date_valid()` helper in `db/day_blocks.py` serves sync, writers, views and
statistics. Matching the word "offen" was rejected because nuLiga may use other
placeholders, and `31.02.2026` must count as open too. `age_eligibility` is a lower layer,
so the import is allowed; `tests/db/test_db_structure.py` confirms it.

**Single active-date computation.** Compute the calendar dates once in `sync_games()` and
remove the preflight refusal. `sync_games()` classifies each season block that has no
scraped game: empty blocks are removed as today; occupied blocks with an open or past date
go into a `retained_dates` set passed to `reconcile_day_blocks()`; occupied blocks of today
or later are cancelled. Assignment relationships are expired first so a stale empty
collection cannot decide the outcome. `reconcile_day_blocks()` keeps its refusal for any
other occupied block as a last guard, which also keeps direct callers safe.

**Cancellation lives in `sync.py`.** A private `_cancel_date_block_duties()` runs before
reconciliation, after block rows have been serialized. Per assignment it audits
`request_close` when a request is open, then `date_remove` with system origin, records an
event and removes the row through collection removal and flush, per the SQLAlchemy pitfall.
Reconciliation then sees empty blocks and deletes them. Everything shares the existing
single commit, so a later failure leaves no partial cancellation. Placing it in
`day_blocks.py` was rejected because of the layer order.

**"No games" means no scraped game.** This matches the existing `active_dates` meaning and
covers both moved and vanished games. A vanished game stays stored on a date without
blocks. "Today" comes from the duty clock (`clock.duty_now().date()`), so a duty on today's
date is still cancelled.

**Empty-scrape guard.** `sync_games()` raises before any write when `scraped` is empty and
the season has stored games. The daily job already reports a failed sync. A first run on an
empty season still passes. A partial bad scrape is not detected; see Risks.

**Events carry snapshots.** `SyncEvents.cancelled_block_duties` holds `person_id`,
`block_date` and the numbered task label, because the rows no longer exist at notification
time. `Notifier.notify_cancelled_block_duties()` sends the new `block.date_cancelled`
template through `_dispatch_message()`, called from `_send_notifications()` in
`jobs/daily.py`. A person missing by then is skipped. Like shift notifications, a failed
stage is reported and not retried.

**Guards in the writers.** `claim_block_slot()` and `configure_cake_block()` refuse
open-date blocks after serialization. Takeover and replacement requests stay unchanged,
since they do not add occupied positions.

**Visibility.** `build_schedule()` appends the dates of retained blocks to its date list
and shows such a group even without games, unless a team filter is active. `block_view()`
reports no vacancies and no claimable free positions. `build_statistics()` skips missing
and setup reporting for these blocks; personal counts are untouched.

**No lock logic change.** `duty_cutoff_at()` already returns `None` for an open date and
`lock_due_duties()` skips it, so retained blocks stay unlocked.

## Risks / Trade-offs

- [A partial bad scrape cancels real duties] → Irreversible for those duties, and helpers
  are told wrongly. The empty-scrape guard covers the total failure; `date_remove` audits
  name everyone affected for manual re-entry, and the database backup holds the prior state.
  Accepted because vanished games were chosen to count as cancelled.
- [A game briefly moves away and back] → Assignments stay removed and helpers must claim
  again, as the notification tells them. Accepted.
- [A notification stage failure leaves helpers uninformed] → The daily job reports the
  failure, and the `date_remove` entries list who must be told by hand.
- [A retained block stays forever without an administrator] → It is visible on the
  schedule, and `admin-reopen-duties` provides the way to empty it.
- [A helper stays responsible for a block on a date that does not exist] → Accepted so that
  no commitment is cancelled silently. Reminders are date-driven and do not fire for it.
- [An unknown real date format would be treated as open] → The scraper stores
  `dd.mm.yyyy`, which all date logic already assumes.

## Migration Plan

No schema migration. The first synchronization after deployment removes empty open-date
blocks and logs each removal, and it cancels occupied blocks of current or future dates
without a scraped game. Since such a date fails synchronization today, a production
database that syncs successfully has none, so no burst of messages is expected. Rollback
restores the previous behaviour: blocks are recreated empty for the placeholder value and
retained blocks are unaffected, but cancelled assignments are only recoverable from a
database backup.
