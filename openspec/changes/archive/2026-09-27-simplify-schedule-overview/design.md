## Context

See `proposal.md` for the motivation and `specs/schedule-overview/spec.md` for the behavior contract. The schedule is server-rendered from `build_schedule()` in `webapp.py`. It already groups games by date, applies GET filters before rendering day groups, puts day-task blocks around each visible date, and nests past dates inside a native `<details>` element. Assignment controls use the existing JSON claim/release endpoints and update their selects without a full page load. `db.REQUIRED_ROLE_SLOT_COUNT` defines five required game slots; `Unterstützung` is optional.

## Goals / Non-Goals

**Goals:**

- Keep filtering and card expansion available without JavaScript, including to guests and keyboard users.
- Calculate initial progress from the same assignment records used for the expanded fields, then keep it in sync with successful inline edits.
- Preserve current authorization, slot mutation, block timing, privacy, and game ordering behavior.

**Non-Goals:**

- Change assignment rules, database schema, notification completeness, or the meaning of optional `Unterstützung`.
- Persist a visitor's expansion state across navigation or page reloads.

## Decisions

### Filter dates on the server with the existing GET form

Build the dropdown from distinct current-season game dates in `db.game_sort_key()` order, before applying any other filter. The empty value means all dates; a nonempty value must match one of those dates. Apply the date condition before building visible day groups. This produces shareable URLs, works without JavaScript, and avoids duplicate client-side filtering logic. Client-side hiding was considered but would duplicate the grouping and block-bookending rules.

For each retained date, first find games satisfying the selected team filters; a date without one still fails the existing team filter rule. Apply the person search separately to each of those games and to each day-task block. Suppress all blocks when the responsible-team filter is selected. A day group survives if any game or eligible block remains, so a person search can produce a date with only one block. Take the day label and past/upcoming classification from the full date rather than the first visible game, which may not exist. Calculate visible block times from all games on the date even when no game card remains. With no person or responsible-team filter, both blocks continue to bookend any filtered game list.

Pass the selected date to the template. When it identifies a past date, set `open` on the existing outer past-games `<details>`; otherwise preserve its default closed state. This affects only the outer section. Each card starts closed on every page load.

### Use native disclosure cards

Render each game and block card as a native `<details>` with a `<summary>` containing identity, metadata, a labeled progress bar, and the expand affordance. A game summary also shows the current responsible-team name, or `– offen –` when unset, as read-only text for every viewer. Keep the responsible-team field and assignment fields in the disclosure body; the team's editable picker remains admin-only. Blocks have no responsible-team line. Preserve the current card IDs and data attributes on the outer card so `static/app.js` can still locate each card and its selects. Adapt the card CSS for collapsed and expanded layouts, focus visibility, and narrow screens. Native disclosure provides keyboard operation and useful behavior without adding a JavaScript state machine. A custom toggle was considered but would require extra keyboard and ARIA handling.

### Derive progress from required slots and confirmed edits

For a game, compute filled and total from `db.REQUIRED_ROLE_SLOT_COUNT`, iterating each role once and checking its defined slots; the total is five. The responsible team and optional `Unterstützung` are outside the fraction. For a block, count occupied slots out of `db.BLOCK_SLOT_COUNT` (three). Return filled, total, and a rounded whole-number percentage in each server-rendered card view. Show the numeric count and percentage beside the bar, and expose an accessible progress label.

After an inline assignment succeeds, `static/app.js` updates only the affected card's progress using the confirmed previous and new occupancy. A replacement leaves the count unchanged; a claim increments it; a release decrements it. Optional `Unterstützung` changes leave game progress unchanged. The existing two-step replacement flow may release successfully and then fail to claim; in that case the existing reload restores authoritative state. Failed single-step operations do not change the progress. This keeps the current JSON response contract and avoids a new progress endpoint. Returning counts from every mutation endpoint was considered, but adds API changes for a value already derivable from the confirmed UI operation.

## Risks / Trade-offs

- [A selected date has only past games, hidden by the outer disclosure] → Open the outer past section when that specific past date is selected; leave the individual cards closed.
- [A person-name search matches an assignment hidden in a collapsed card] → Keep each individually matched card visible with a clear expand control; expanding reveals the public occupant name.
- [A block-only result has no game view to supply its day or past/upcoming status] → Derive those values from the date's full game set and retain the date group while the block is visible.
- [A stale client may display progress from an earlier page load while another user edits a different slot] → Render server-authoritative counts on each page load and retain conflict-triggered reloads for stale writes. Do not update progress for rejected operations.
- [Nested disclosure styling could make cards hard to scan on phones] → Test the collapsed and expanded layouts at narrow widths and keep text counts visible beside each bar.

## Migration Plan

Deploy the route, template, CSS, and JavaScript changes together. No database migration or backfill is needed. Rolling back these files restores the current full-card presentation; existing filter URLs continue to work.
