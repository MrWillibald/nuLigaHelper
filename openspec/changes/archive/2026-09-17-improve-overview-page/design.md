## Context

See `proposal.md` for motivation and the two delta specs for observable behavior. `build_schedule()` already sorts season games in Python, builds game cards, and splits day groups using `common.effective_today()`. `schedule.html` already uses a native `<details>` control for past days. The shared footer in `base.html` already contains both legal links, while `style.css` gives footer links no specific color. Guests may see assigned helper names on cards but no roster or person IDs.

## Goals / Non-Goals

**Goals:**

- Keep filtering usable without JavaScript, including on narrow screens and with a keyboard.
- Preserve the existing identity, privacy, date, sorting, and assignment rules while adding read-only filters.
- Reuse current page patterns and visual tokens for filter controls, past section, and footer.

**Non-Goals:**

- Change assignment eligibility, past-game edit rights, season selection, legal page content, or the database schema.
- Build a public person picker or publish a separate roster through the schedule.

## Decisions

### Filter with a GET form on the existing schedule route

Add playing-team and responsible-team selects plus a free-text assigned-person field. `build_schedule()` will evaluate the criteria against games before building day and month groups, then return the filtered groups and filter state to the template. Existing team IDs identify teams; the person search matches normalized, case-insensitive text against current assignment occupants' display names. Empty values mean no filter. Treat a malformed or unknown selected team ID as matching no games rather than broadening the result. Use a link to the bare schedule route to clear filters. This makes filtered URLs shareable and works without JavaScript. Client-side filtering was considered, but it would require exposing searchable assignment metadata and maintaining duplicate DOM visibility logic.

### Keep guest filtering within the public card data

The person field is a name search, not a selector backed by `person_options()`. The template must not render `persons` or assignment person IDs to guests; the server uses assignment records only to decide which cards to include. Two people with the same name can both match; this is appropriate for read-only discovery and does not affect mutation identity. Person filtering by ID was considered, but putting a public person list or IDs into controls would violate the guest schedule boundary.

### Group only filtered games, then split past and upcoming days

Reuse `db.game_sort_key()` and the existing effective-day comparison. The past `<details>` stays closed initially and uses the number of filtered past day groups in its summary. Show the no-games message only when the season has no games; show a separate no-results message when filters remove every game. Refreshes naturally restore the default collapsed state. This keeps month headers and date blocks from appearing without cards.

### Refine styles through existing components

Use the established filter card, form field, button, and spacing patterns for the overview. Give the past section a clear neutral boundary and muted day/card colors while keeping text legible instead of lowering the entire card's opacity. Explicitly style `.footer a` for light text, subtle hover treatment, and a high-visibility `:focus-visible` outline; retain link underlines or another clear affordance. Apply responsive rules already used by other filter cards. The shared footer template can remain unchanged unless a small semantic adjustment improves the result.

## Risks / Trade-offs

- [Name search is ambiguous for duplicate display names] → Label it as a name search and match all assigned occupants with that text; keep every mutation based on `Person.id`.
- [A guest could infer a name has an assignment by searching] → Results only contain game cards and assigned names already public on the unfiltered schedule; return no roster, IDs, or contacts.
- [Many filters could produce an apparently empty page while past matches are collapsed] → Keep the filtered past-day count visible and provide a distinct all-results-empty message only when both sections have no matches.
- [Muted styling could impair readability] → Use explicit neutral colors with readable contrast and verify on desktop and narrow layouts rather than applying low opacity to all card contents.

## Migration Plan

No data migration is needed. Deploy the template, route, and CSS changes together; rollback restores the previous overview rendering. Existing unfiltered schedule URLs remain valid.
