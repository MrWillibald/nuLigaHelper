## Why

Verified registrations currently notify only one administrator, so other administrators may not learn that approval is needed. After approval, the registrant receives no confirmation and no invitation to begin participating in home-game duties.

## What Changes

- Notify every active administrator after a registrant successfully verifies their contact, using each administrator's automatic notification channel preference.
- Give administrators a personal greeting, identify the registrant and selected teams, and direct them to review the registration under "Helfer verwalten".
- After an administrator successfully approves a registration, send the newly activated user a welcome notification confirming approval and inviting them to sign in and take open duties in the Heimspielplan.
- Treat these informational deliveries as best effort: log missing contacts or delivery failures without undoing verification or approval.
- Prevent stale or repeated approval requests from producing duplicate welcome notifications.

## Capabilities

### New Capabilities

None.

### Modified Capabilities

- `user-accounts`: Define administrator notifications for verified registrations and user welcome notifications after approval.

## Impact

- `webapp.py`: registration-verification and approval notification orchestration.
- `notifier.py`: existing automatic e-mail-first, SMS-fallback account-message delivery.
- `test/test_auth.py` and `test/test_notifier.py`: coverage for recipients, message content, channel preference, delivery failure, and duplicate prevention.
- `README.MD`: registration and approval notification behavior.
- No database schema, API, provider, or external dependency changes are expected.
