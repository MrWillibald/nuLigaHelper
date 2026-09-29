## Context

See [proposal.md](proposal.md) for motivation. Today `build_schedule()` builds the active roster and sorts candidate dictionaries for each game slot; `schedule.html` renders those options into every collapsed card. Admin views multiply the full roster by six game slots and three slots per day-task block. The schedule also calls a team helper that loads team members although this page needs only team metadata. Assignment writes already use protected compare-and-swap endpoints; candidate loading must follow their tier and date rules.

## Goals / Non-Goals

**Goals:**

- Make initial schedule response size and DOM node count depend on displayed cards and occupied slots, rather than the number of unassigned roster members.
- Bound candidate work to the card opened by the viewer and keep one person representation per card response.
- Recheck permissions when candidates are requested and preserve existing assignment semantics.

**Non-Goals:**

- Change the roster page, assignment schema, date storage, or notification behavior.
- Change the current requirement for JavaScript to submit inline assignments.

## Decisions

### Fetch candidates per opened card

Add authenticated read endpoints for one game and one day-task block. The initial HTML contains the current occupant (including an inactive historical occupant) and a placeholder per editable slot, but no unassigned roster. On opening a card, JavaScript requests its candidates once and populates only that card. The controls stay disabled while loading and on failure; a subsequent opening or explicit retry can request again. Use `Cache-Control: private, no-store` because team membership and permissions can change. A GET needs no CSRF token, and the existing default-deny route guard protects it.

An alternative was embedding one shared roster in the initial page and creating options on expansion. That would reduce duplicate DOM nodes but still make the initial response and roster query grow with registrations. Another alternative was fetching on every individual select focus; this could issue six similar requests per game and make selecting a person feel slow.

### Return one roster per card with slot eligibility

Each response contains compact, presentation-safe person data (`id`, display name, complete team label, and category/hint metadata), the relevant slot keys with their permitted candidate IDs, and currently occupied person IDs. Only the union of people authorized for at least one editable slot is returned. The server determines eligibility from current tier, managed teams, game responsibility, slot occupant, date, and active account status. The browser uses the slot eligibility map to populate selects and excludes occupants of sibling slots. The current occupant is preserved from the rendered card even when absent from the active roster.

An alternative was a separate full candidate array for every slot. The per-card roster plus slot IDs avoids repeating names and team labels six times in the response. The write endpoints remain authoritative if state changes after loading.

### Reuse ordering and permission rules

Extract the existing slot editability and candidate category calculations so page rendering and candidate endpoints use the same rules. The server supplies stable category and name ordering metadata; the client applies that order when it creates options and reinserts a released person. Keep the playing-team precedence and advisory labels. A day-task block has no team category but retains alphabetical order and its own one-task limit. The schedule's team selector should use a lightweight team query; the candidate endpoints should eager-load active persons' memberships in a bounded number of queries.

### Verify growth and access with synthetic data

Test the initial response with a fixed schedule and progressively larger synthetic roster, asserting that unassigned names and repeated `<option>` nodes are absent. Test direct candidate reads for guest, pending, member, MV, and admin, including changed permissions and past dates. Test expansion, loading failure, selected historical occupants, duplicate names, warning metadata, sibling exclusion, and claim/release updates. Use response-size and query-count assertions where deterministic; treat wall-clock benchmarks as diagnostic rather than a flaky pass/fail gate.

## Risks / Trade-offs

- [First expansion adds a network round trip] → Load immediately on card open, show a short loading state, and request only one card at a time.
- [A stale page may show old candidates or occupants] → Recheck access on reads and writes, retain compare-and-swap conflicts, and reload the affected card or page after a conflict.
- [Accidental roster or contact disclosure through a new endpoint] → Keep the endpoint behind default-deny auth, build an explicit safe response shape, and test every tier directly.
- [An inactive occupant is missing from the active roster] → Preserve the selected occupant from the initial card and do not add that person to another slot when released.

## Migration Plan

This is an application-only change with no schema migration. Deploy the updated server and static assets together, then verify an anonymous overview, an admin card expansion, and one member self-service card. Roll back the application release if the new candidate request fails; existing database contents require no rollback.
