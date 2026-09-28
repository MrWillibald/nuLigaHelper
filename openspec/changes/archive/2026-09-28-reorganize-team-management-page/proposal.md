## Why

On "Helfer verwalten", the roster appears before the controls for creating users, deciding registrations, and appointing MVs. Administrators and MVs must scroll past the roster to reach the actions they use to organize their teams.

## What Changes

- Show the available team-management controls before the roster on the person-management page.
- Keep the roster, including its filters and person entries, after those controls for administrators and MVs. Members continue to see the roster without an unavailable management section.
- Preserve the existing visibility and authorization rules for each action and all roster data.

## Capabilities

### New Capabilities

None.

### Modified Capabilities

- `access-control`: Specify the order of the tier-appropriate management controls and roster on the person-management page without changing permissions.

## Impact

- The `templates/persons.html` page structure and focused management-page rendering tests will change.
- No database schema, API, authorization rule, or dependency changes are expected.
