## Context

See `proposal.md` for the motivation and scope. The current person page renders the
administrative form and team/MV controls in a left column beside the member cards. The
server already derives `g.tier` and `g.mv_team_ids`, and registration approval already
authorizes MVs for their managed teams. The current add-person route, however, is
admin-only. The audit route already provides newest-first entries and game/person filters,
but its template has no comparable page-specific visual treatment.

## Goals / Non-Goals

**Goals:**

- Make the roster the primary, full-width content on the management page.
- Make routine MV user creation available without weakening server-side team boundaries.
- Make the tier-dependent administration cards obvious and consistent.
- Add safe, shareable roster filtering while preserving contact-data privacy.
- Give audit entries a coherent desktop presentation and a non-scrolling mobile
  presentation.

**Non-Goals:**

- Do not change the database schema or add a migration layer.
- Do not change contact storage, authentication routes, notification fallback, or the
  separate e-mail/phone fields.
- Do not change team derivation, MV appointment rules, assignment semantics, or audit
  persistence.
- Do not expose the audit protocol to non-admins.

## Decisions

### 1. Stack the management page by importance

Render the member area as the first full-width section. Render a lower administrative grid
with three independently permissioned cards: `Neue Benutzer`, `Offene Registrierungen`, and
MV assignment. Keep the member cards editable according to the existing rules, rather than
introducing a separate member-management page.

The lower grid collapses to one column on narrow screens. Cards that are not authorized are
not rendered, leaving no empty placeholders.

### 2. Restrict MV creation on the server and in the form

Allow the add-person endpoint for admins and MVs. For an MV, validate the submitted team
against the complete set of teams represented by `g.mv_team_ids`; do not rely on the select
options as authorization. Render only those managed teams for an MV. Admins retain all
existing team choices.

An MV-created person follows the existing direct-create behavior: active roster status,
normal contact validation, and no registration challenge. The support team is offered only
if it is actually among the MV's managed teams; otherwise it is unavailable to that MV.

### 3. Use query parameters for roster filtering

Represent name, team, and admin-only status filters as GET parameters on `/personen`. This
keeps filtering server-authorized, works without JavaScript, and makes filtered views
bookmarkable. Apply filters only after the existing tier-based visible-person selection.
Preserve the selected filter values when rendering the form.

Client-side JavaScript is not required for correctness. It may be used only for a small
interaction enhancement if it does not replace server filtering or reveal hidden data.

### 4. Preserve the existing registration approval contract

Keep the existing MV approval and rejection authorization and use the lower registration
card to expose it to authorized MVs. Display only registrations whose desired team is in
the viewer's managed-team set for MVs; admins continue to see all pending registrations.
Registrations for support or teams without an MV continue to require admin action.

### 5. Adapt audit markup instead of changing audit data

Keep the current route, filters, ordering, and snapshot fields. On desktop, use a styled
filter card and table with consistent typography, borders, spacing, and action/tier badges.
On mobile, use responsive entry cards or equivalent stacked markup so every field remains
visible without horizontal scrolling. Use semantic labels for stacked fields and retain
read-only markup.

### 6. Keep responsive behavior CSS-first

Add narrowly scoped management and audit classes to `style.css`, reusing existing design
tokens, card radii, buttons, and breakpoints. Avoid changing global navigation or schedule
styles. Use the existing `app.js` only for interactions that require asynchronous behavior,
such as MV selection; roster filtering and audit readability must work without it.

## Risks / Trade-offs

- [Risk] An MV for several teams could submit a forged team ID. -> Validate membership in
  `g.mv_team_ids` in the POST handler and add direct unauthorized-request tests.
- [Risk] A filter implementation could accidentally reveal hidden or inactive people. ->
  Build the filtered list only from the already tier-filtered query/result set and test
  member, MV, and admin responses separately.
- [Risk] Three cards may become cramped on tablets. -> Use a two-column lower grid at
  medium widths and one column at the existing mobile breakpoint.
- [Risk] A wide audit table may remain difficult to read on intermediate widths. -> Switch
  to stacked audit entries before the viewport reaches the mobile breakpoint, or provide a
  deliberate table-to-card layout rather than relying on overflow.
- [Risk] MV-created active users bypass verification because they are direct roster
  entries. -> Preserve the existing admin creation contract and make this behavior explicit
  in tests and the UI copy.

## Migration Plan

No data migration is required. Deploy the template, CSS, route, and test changes together;
existing people, contacts, registrations, teams, assignments, and audit entries remain
valid. Rollback consists of reverting the application files; no persisted data needs to be
transformed.
