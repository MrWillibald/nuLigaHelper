# Design

## Context

See `proposal.md` for motivation. Current game and block assignment rows represent occupancy without replacement-request or durable cutoff state. `db.claim_slot()` serializes on the game before fresh occupant, active-person, uniqueness and age checks; block claims and cake configuration serialize through `_lock_day_block()`. Release currently deletes the assignment. The JavaScript replacement flow releases first and then claims, which can leave a vacancy when the claim fails.

`main._send_notifications()` opens a non-writing session and dispatches game/block day-before reminders for `effective_today() + 1`; weekly reminders are separate at `+7`. Reminder delivery currently has no durable assignment transition. The user has explicitly confirmed a fixed cutoff of 09:00 Europe/Berlin on the day before the duty, including when the daily job or provider fails. This supersedes the earlier actual-batch-event design.

Existing specifications also promise unconditional release after reminders, release of the sole adult seller, retained-duty release-only maintenance, administrative past corrections, and assignment cleanup during person deactivation/deletion and disappearing day blocks. The delta specs replace or qualify these conflicting requirements. The existing offered/required sets, private birth dates, per-container uniqueness, German UI and centralized notification templates remain constraints.

## Goals / Non-Goals

**Goals:**

- Keep the responsible helper in place while searching, transfer only with validated atomic takeover, and make the saved responsibility understandable before any assignment write.
- Enforce the fixed local cutoff using a shared aware clock under serialization, with durable lock/request transitions coordinated with claims, takeovers, cake settings and reminder batches. Check authoritative state again after contention retries.
- Preserve original slot numbers, saved assignment identities where practical, physical counts, age deficiencies and readable audit history.
- Deliver the confirmation together with its supporting responsibility protections so the active notice is truthful; do not activate a notice that promises server guarantees before those guarantees exist.

**Non-Goals:**

- A search-duration timer, 24-hour grace period, automatic return, periodic rollback, replacement broadcast or new provider integration.
- A new administrator override or in-service arrangements after an occupied duty locks.
- A claim that an MV/admin acting on behalf of a person proves that person's personal acknowledgment.
- Changes to required-duty counts, age minimums, authentication or public roster privacy.

## Decisions

### 1. Explicit confirmation before the first write

For any assignment creation or takeover, select the target helper/duty locally and show a German commitment notice with `Bestätigen` / `Abbrechen`. Selecting or opening must not call the mutation endpoint. Explicit confirmation submits the same empty-occupant expectation (or current takeover generation) with CSRF and current permission checks. Cancel, Escape, dialog dismissal and any non-JavaScript cancel action perform no assignment mutation and never invoke a compensating release. If another client wins during the dialog, cancellation preserves that assignment and confirmation receives the existing conflict.

Use an accessible dialog where JavaScript is available, with associated title/body, focus placement/return and keyboard dismissal; provide a server-rendered `/dienst` picker and explicit confirmation step for non-JavaScript flows that create assignments. Cancellation returns to the schedule without calling an assignment writer. Identify duty/date/helper so the user knows what is being confirmed. Every creation path must use this boundary, including administrator/MV assignments and cake positions. The notice for a late empty claim explains that it becomes fixed immediately. Rejected alternative: save first and undo on Cancel; that creates responsibility, audit and race effects before consent.

### 2. Persist lifecycle state and an immutable request generation

Persist `replacement_token` and `locked_at` on game/block assignments. A fresh replacement token is the request generation used by takeover compare-and-swap. A `DutyCutoff` marker keyed by season/home-game date records application of the fixed day-before cutoff even when a slot was empty. A missing marker does not postpone the fixed deadline. Existing assignment rows keep their original person and occupancy while requests are open. An assignment lock marker or equivalent immutable association preserves a lock already applied to a particular commitment; ordinary metadata/date corrections must not silently reopen it.

Normal states are empty; occupied; occupied plus replacement wanted; and occupied plus locked. A valid takeover closes the current request and changes the responsible person in one transaction. There is no time-driven return state. A duty whose fixed deadline has been reached makes a later empty-slot claim immediately locked, even if no reminder batch ran. New block age rules are not introduced.

The current active helper may explicitly end their own open search before cutoff while keeping the same assignment and responsibility. Existing admin and responsible-team MV management scope is retained; a former helper or unrelated member gains no authority. `cancel_replacement()` / `cancel_block_replacement()` and matching `POST /api/assignment/cancel-replacement` / `POST /api/block-assignment/cancel-replacement` compare both expected occupant and request token after writer serialization and retries. Success only clears the token and records one `request_close` event with the actual actor. A stale, closed or reopened generation refuses without another successful event; reopening creates a fresh token. The German `Ersatzsuche beenden` action uses the same server boundary in JavaScript and non-JavaScript forms. This requires no additional schema revision.

Use assignment columns plus the dated cutoff table rather than overloading audit history as mutable state. Existing audit rows must stay unchanged and migration must not fabricate historical acknowledgment.

### 3. Atomic takeover uses the existing serialized domain boundary

Acquire the game's/block's writer serialization before refreshing the target, current person/request generation, actor rights, active account, offeredness/capacity and cutoff marker. Validate one task per person per container and incoming eligibility using saved data. For Verkauf, evaluate the resulting group with the outgoing seller excluded and the replacement included; an adult seeking replacement continues to supply coverage until the transaction actually replaces them.

The game and block request endpoints are `POST /api/assignment/request-replacement` and `POST /api/block-assignment/request-replacement`; atomic transfer uses the matching `/takeover` endpoints. A takeover submits `expected_person_id`, `expected_request_token` and the incoming `person_id`. Domain operations `request_replacement()` / `request_block_replacement()` and `takeover_slot()` / `takeover_block_slot()` enforce this state; the existing `/claim` routes stay empty-slot-only. Store responsibility transfer, request closure and linked before/after audit snapshots together. Use linked outgoing/incoming audit events or an equivalent explicit before/after snapshot, not an uncorrelated public release and claim. After a retry, re-read everything that could affect authorization or eligibility. No network I/O belongs inside this transaction. Rejected alternative: browser release-then-claim; it cannot preserve responsibility on a failed replacement.

### 4. Fixed day-before cutoff independent of job or delivery

The authoritative boundary is 09:00 Europe/Berlin on the calendar day before the scheduled duty date. Use shared timezone-aware `db.duty_now()` and calendar arithmetic in that zone, including daylight-saving transitions; the effective date continues to use `common.effective_today()`. Explicit `db.lock_due_duties()` preflight commits due locks before schedule/API mutations and notification work, without secretly committing unrelated pending work. Schedule and candidate reads derive the same due lock state without writing, so selecting, reviewing or cancelling an assignment remains mutation-free. The same scheduled-date rule covers every offered game position, retained occupied role and preparation/cake/cleanup position. Weekly reminders never lock assignments.

Every request, takeover, release and claim must recheck the clock after acquiring writer serialization and after retries. At or after the deadline an occupied duty is fixed even when its marker has not yet been written. Persist the dated cutoff and any first lock/request-closure transitions atomically and idempotently; stale pages, missed jobs, restarts, empty-at-cutoff positions and absent contacts do not defer enforcement. A saved assignment lock survives subsequent date/category changes. Empty offered duties remain claimable within existing date/eligibility rights and successful late claims become locked in that transaction.

Before dispatching day-before reminders at or after the deadline, commit the dated cohort transition for all game/block duties. Close open requests and retain their current helpers; include contactless helpers and empty-slot date state. Provider delivery remains outside writer serialization. Failure or retry cannot undo locks or duplicate lifecycle audit transitions. A reminder call before 09:00 does not advance the cutoff merely because it is a day-before call. No replacement broadcast is added; later notification wording changes use `messages.py`.

A missed batch remains an operational failure, but the 09:00 cutoff still applies. Authoritative reads and writes must not display or accept a post-deadline replacement merely because cron did not run. Persisting cutoff state is not evidence that a message was delivered. A takeover that obtains write access before 09:00 may win and its incoming helper later locks; a takeover that obtains access at or after 09:00 is refused and the original helper remains responsible. There is no automatic release, search-duration timer or periodic rollback.

### 5. Keep current tier scopes and public data boundaries

Members request replacement only for themselves and take over for themselves. An active request authorizes a prospective replacement to take that duty; it does not grant them authority to manage the outgoing helper's other duties. MVs retain the responsible-team and selected-person membership checks for game management, plus their ordinary self-service rights. Block MVs retain member scope. Administrators retain roster-wide staffing scope but no cutoff bypass. Confirmation belongs to the acting user; it does not invent communication or consent from an absent helper.

Candidates use current incoming-person authority, post-takeover eligibility, offeredness and lock state. Return no replacement candidates for a locked duty. Keep current occupants visible even if inactive or age-deficient; refresh only affected cards/sibling candidates after successful transfers. Guests may see helper names and read-only `Ersatz gesucht`/fixed status, never roster payloads, person IDs, contacts, full birth dates or exact ages. Requests do not change physical progress, missing counts or personal duty ownership. A takeover changes whose personal statistics count the duty without creating a required vacancy.

### 6. Refuse destructive maintenance that would cancel commitments

The user explicitly confirmed refusal of destructive maintenance when it would cancel a binding assignment. This supersedes the earlier compatibility default preserving lifecycle cancellation. Existing broad cleanup guarantees must be qualified consistently across domain helpers, web, CLI and synchronization; there is no administrator exception.

| Operation | Before fixed cutoff | At/after fixed cutoff | Responsibility/history |
| --- | --- | --- | --- |
| Ordinary withdrawal of an offered duty | Open `Ersatz gesucht`; retain occupant | Refuse | Only eligible atomic takeover transfers responsibility before lock |
| Account deactivation that would remove a current/future commitment | Refuse atomically | Refuse atomically | Retain account, MV appointments, assignments and request state; past history remains intact |
| Erroneous-record deletion that would remove any assigned duty, including past duties | Refuse atomically | Refuse atomically | Retain person and assignments; no successful removal audit |
| Game/day/block removal that would remove a binding commitment | Refuse atomically | Refuse atomically | Retain scheduled containers, occupants and request/history state; report the failed synchronization or maintenance |
| Retained removed-role retirement | Existing authorized release-only maintenance | Refuse | Audit original position; never claim/take over or change required-staffing counts |
| Cake quantity reduction removing an occupied position | Refuse | Refuse | No force release, renumbering or compaction |

Maintenance without affected binding commitments retains existing authority and history rules. Deactivation preserves past assignments, clears qualifying MV appointments only when deactivation succeeds, and reactivation never restores historically released assignments. Deletion of an erroneous uncommitted record remains admin-only and preserves readable historical audit entries. A blocked synchronization must not partially apply destructive game/day/block changes or emit successful-change notifications.

Retained Kasse/Reinigung assignments remain visible, counted personally, included in ordinary reminders and excluded from required progress/MV gaps. They never gain replacement candidates because the duty is no longer offered. An authorized explicit release before cutoff retires that obsolete duty; after cutoff the occupied lock applies. Original audit snapshots remain unchanged.

Moving an already locked game or changing its category does not reopen the commitment. Preserve eligibility reevaluation and the game's identity; the stored assignment lock survives metadata changes. Cancellation is not a lock bypass: if removal would destroy a binding duty, retain the saved commitment and report refusal. Cake reductions remain impossible while a removed numbered position is occupied; replacing the helper keeps it occupied.

## Risks / Trade-offs

- [Reminder transport fails after cutoff] → Persist first; failure cannot reopen duties. Preserve established skip/error logging and avoid claiming delivery succeeded.
- [Cutoff/claim/request races] → Share container serialization, compare request generation as well as occupant, and recheck after retries.
- [Full Verkauf loses adult coverage on replacement] → Validate the final group before changing the outgoing assignment.
- [Legacy UI/CLI still releases directly] → Inventory every mutation path and replace or guard it; qualify conflicting specs and documentation together.
- [Missed cron] → Enforce 09:00 Europe/Berlin in authoritative reads/writes without waiting for cron; expose delivery failure separately.
- [Migration invents consent/history] → Preserve assignments and audits; new state defaults cannot claim old acknowledgments.
- [Destructive cleanup breaks a commitment] → Refuse before mutation, preserve account/container/assignment state atomically and test admin/CLI/sync refusal.

## Migration Plan

1. Implement and verify solely with synthetic offline data; no real config, providers or database are required.
2. Add a reviewed Alembic revision joining the single head, update exact schema fingerprints/postflight checks, and preserve people, assignments, memberships, MV references and old audit snapshots.
3. Align domain helpers, CLI, sync, Flask routes, confirmation UX, tests and documentation with the confirmed fixed-clock and maintenance-refusal decisions before shipping the complete lifecycle.
4. For a later deployment, stop all writers, run the existing guarded backup-first migration, inspect verification and restart matching application code. Normal startup never migrates implicitly.
5. If verification or rollout fails, stop writers and restore the printed validated snapshot with the prior application. No in-place history rewrite or downgrade is proposed.

## Open Questions

No material product-policy question blocks implementation. The user explicitly approved ending an own replacement search while retaining the same helper; this is included in the scoped follow-up and is not a duty-release operation. The fixed 09:00 Europe/Berlin boundary applies even when cron or delivery is missed; delivery catch-up remains operational and cannot reopen commitments.
