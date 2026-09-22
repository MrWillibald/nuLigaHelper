## Context

See `proposal.md` for motivation. The server currently builds a globally name-ordered roster and gives each editable task slot a candidate list based on the viewer's tier. The template removes people taken by another slot and derives the playing/outside warnings. After an inline release, `static/app.js` clones the released option back into the other selects and currently sorts only by rendered label.

The category depends on the individual game, so it cannot be added once to the shared roster. Initial rendering and client-side reinsertion must use the same per-game category metadata without changing the compare-and-swap assignment APIs or tier restrictions.

## Goals / Non-Goals

**Goals:**

- Derive one deterministic, mutually exclusive category for each candidate in the context of a game.
- Keep initial server rendering and post-release client updates in the same order.
- Preserve candidate eligibility, duplicate-assignment exclusion, selected occupants, and warning semantics.

**Non-Goals:**

- Adding visible headings or HTML `optgroup` elements to the dropdowns.
- Changing who may assign whom, assignment validation, or persistence.
- Changing the responsible-team selector or schedule filters.

## Decisions

### Compute category metadata while building each game's slot options

Introduce a small server-side ordering helper that accepts the already authorized candidate list plus the responsible, Supporter, and playing team IDs. It will return per-game option dictionaries carrying a numeric sort group and normalized name key, ordered by `(group, normalized name, person id)`.

Category classification will check playing-team membership first so playing members always appear in the final warning group, then responsible-team membership, then Supporter membership, with all remaining people—including people without a team—in the third group. In normal data the categories do not overlap; the explicit precedence also gives deterministic behavior if an administrator makes the playing team responsible or the Supporter team is responsible.

This approach keeps authorization filtering where it already occurs and sorts only the resulting candidates. An alternative was to group in Jinja, but that would duplicate ordering logic and make the JavaScript update path harder to keep consistent.

### Encode sort metadata on each option

Render the numeric category, normalized display-name key, and person ID as `data-*` attributes on each person option. Warning classes, titles, and visible German hints remain as they are.

The client-side insertion function will compare this metadata rather than the full rendered label. Cloning an option therefore preserves both its warning state and its position rules. An alternative was reloading the page after every release, but that would discard the existing responsive inline update behavior.

### Filter taken people independently of sorting

Keep the existing per-game set of assigned person IDs and omit a taken person from other slots while allowing the current slot's selected occupant. The ordering helper does not alter this uniqueness rule, and the server-side claim validation remains authoritative.

## Risks / Trade-offs

- [Server and browser ordering diverge] → Use server-generated category/name/id metadata for both paths and cover initial rendering plus dynamic reinsertion in tests.
- [Overlapping team roles produce ambiguous categories] → Apply the documented precedence: playing last, then responsible, then Supporter, then other.
- [Sorting accidentally broadens MV or member choices] → Filter by the existing tier rules before applying the ordering helper and test restricted views.
- [Duplicate display names make tests or ordering unstable] → Use the immutable person ID as the final tie-breaker without exposing additional contact data.

## Migration Plan

No data migration or deployment sequencing is required. Deploy the server, template, and static JavaScript changes together. Rollback consists of reverting those files; stored assignments are unaffected.
