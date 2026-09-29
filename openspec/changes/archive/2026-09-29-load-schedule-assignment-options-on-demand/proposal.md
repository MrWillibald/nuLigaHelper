## Why

The schedule currently renders every eligible person into every editable task dropdown, including cards that remain collapsed. As registrations grow, the response, server rendering work, and browser DOM grow with the product of roster size and task slots, slowing the production overview.

## What Changes

- Keep the initial schedule compact: show current assignments and controls, but load candidate lists only when an editable game or day-task card opens.
- Return only candidates the current viewer may assign, with names, team labels, ordering and advisory information needed by the dropdowns. Do not return contact data.
- Preserve the existing one-task-per-container rule, selected occupants, category order, warnings, and compare-and-swap assignment behavior.
- Make candidate-loading failures visible and retryable without changing stored assignments.
- Add offline checks for response size, access scope, and expanded-card behavior; document the loading behavior.

## Capabilities

### New Capabilities

None.

### Modified Capabilities

- `schedule-overview`: Editable card expansion loads assignment candidates on demand while the initial overview remains compact.
- `task-self-service`: Candidate ordering, exclusion, selection, and warning rules also hold for on-demand lists.
- `access-control`: Candidate requests enforce current tier and assignment scope without exposing contacts or a guest roster.

## Impact

The schedule builder in `webapp.py`, `templates/schedule.html`, and `static/app.js` will change. Authenticated candidate-read API endpoints and focused tests will be added. No database migration or provider integration is expected.
