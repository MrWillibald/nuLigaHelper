## Why

The person-management page gives administrative controls more visual prominence than the
member roster and places related controls in an unintuitive side column. The activity
protocol also does not use the application's established visual language and is difficult
to read on small screens. MVs additionally need a practical way to maintain their teams'
people without depending on admins for routine work.

## What Changes

- Restructure the person-management page with a full-width member-card area at the top.
- Add member search and filtering by name, team, and appropriate account status.
- Move administration below the member area into three cards: "Neue Benutzer", "Offene
  Registrierungen", and MV assignment.
- Show the new-user card to MVs and admins; restrict MV team choices to all teams managed
  by the current MV while keeping admins unrestricted.
- Allow MVs to create active users for teams they manage.
- Show open registrations to admins and MVs, with MVs limited to registrations for teams
  they manage; keep MV assignment admin-only.
- Leave the existing separate e-mail and phone fields and all contact, authentication,
  notification, and privacy behavior unchanged.
- Restyle the activity protocol to match the existing page, card, filter, button, and
  typography patterns.
- Present activity entries as readable stacked cards on mobile while preserving the
  existing filters, newest-first ordering, read-only behavior, and audit data.

## Capabilities

### New Capabilities

None.

### Modified Capabilities

- `user-accounts`: MVs may create active users for teams they manage, and the management
  interface groups user creation and registration approval according to tier.
- `access-control`: MV authorization includes creating users for managed teams and acting
  on registrations for managed teams, while MV appointment remains admin-only.
- `assignment-audit`: The admin review interface gains the established responsive visual
  treatment without changing its append-only data or access rules.

## Impact

- `templates/persons.html` and `templates/audit.html` require structural and presentation
  changes.
- `static/style.css` and possibly `static/app.js` require responsive layout, filtering, and
  interaction updates.
- `webapp.py` requires server-side authorization and team-option changes for MV user
  creation, plus member filtering data handling.
- Existing registration approval authorization and audit persistence should be preserved.
- Tests for tier permissions, filtering, responsive markup, and existing CRUD behavior need
  updating or extension; no database migration is expected.
