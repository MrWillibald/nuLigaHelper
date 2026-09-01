## 1. Authorization And View Data

- [x] 1.1 Extend person creation authorization so admins can choose any existing team and MVs can choose only teams in their managed-team set; verify with forged team-ID tests for member, MV, and admin requests.
- [x] 1.2 Expose the complete managed-team set to the MV creation form and preserve the existing admin team options; verify an MV managing multiple teams sees and can select every one.
- [x] 1.3 Add server-side roster filters for name and team for signed-in users and account status for admins, applied after tier-based visibility filtering; verify filtered responses do not reveal unauthorized people or contact data.
- [x] 1.4 Preserve and surface the existing MV registration approval/rejection behavior in the management view, including managed-team scope and admin-only handling for support or unassigned teams; verify authorized and unauthorized decisions.

## 2. Person Management Layout

- [x] 2.1 Restructure `persons.html` so the full-width member-card area is first and the lower area contains the three independently permissioned cards "Neue Benutzer", "Offene Registrierungen", and MV assignment; verify the cards render in the required tier combinations.
- [x] 2.2 Add the roster search and filter controls with preserved query values and accessible labels; verify filtering remains usable without JavaScript.
- [x] 2.3 Keep separate e-mail and phone fields, existing contact privacy, self-editing, status actions, and duplicate-name team context unchanged; verify the existing webapp CRUD and privacy tests pass.
- [x] 2.4 Add responsive management styles for the full-width member area and lower card grid using existing design tokens; verify the lower cards collapse cleanly on narrow viewports through rendered markup and CSS review.

## 3. Activity Protocol Presentation

- [x] 3.1 Restyle `audit.html` with the established section header, filter card, controls, table, and action/tier treatments without changing route behavior; verify newest-first ordering and game/person filters remain intact.
- [x] 3.2 Add mobile stacked audit-entry markup or equivalent responsive presentation with semantic field labels and no horizontal-scrolling dependency; verify all audit fields remain visible in the narrow-screen structure.
- [x] 3.3 Preserve read-only audit behavior and admin-only access; verify non-admin refusal and absence of edit/delete controls.
- [x] 3.4 Use one shared desktop and responsive table treatment for statistics and the activity protocol; verify both pages render the shared table classes and mobile field labels.

## 4. Verification

- [x] 4.1 Extend the offline webapp tests for MV multi-team creation, member/admin filter visibility, tier-specific card rendering, and responsive audit markup; verify with `test/run_tests.sh`.
- [x] 4.2 Validate the complete OpenSpec change and inspect the final rendered templates for contact-data privacy and unchanged contact behavior; verify with `openspec validate "redesign-management-activity-ui" --type change --strict` and the full test suite.
