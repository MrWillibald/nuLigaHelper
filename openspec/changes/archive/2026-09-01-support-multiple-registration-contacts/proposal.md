## Why

Self-registration currently permits only one contact route, forcing users to add
their second route later and limiting both notification resilience and login
options. Registration should collect the user's complete contact information up
front while still requiring a verified route before the account becomes usable.

## What Changes

- Allow registration to accept an e-mail address, a phone number, or both.
- Validate every supplied contact server-side and block submission while any
  supplied contact is invalid; an invalid field must be corrected or cleared.
- Reorder the registration form so the e-mail and SMS fields appear before the
  contact-route selection.
- Enable route selection only for supplied contacts that pass validation.
- Send the registration verification code through the explicitly selected valid
  route and persist both supplied canonical contacts.
- Allow passwordless login through either stored route when both are available.
- Preserve e-mail as the preferred channel for all automatic notifications, with
  SMS as fallback when no usable e-mail address exists.
- Keep existing privacy, uniqueness, rate-limiting, verification, and approval
  behavior for registrations.

## Capabilities

### New Capabilities

None.

### Modified Capabilities

- `user-accounts`: Change registration to support multiple valid contacts and
  define route selection, automatic-notification preference, and dual-route login.

## Impact

- Registration validation and persistence in `webapp.py` and `db.py`.
- Registration contact fields and route-selection behavior in
  `templates/register.html`, `templates/_auth_fields.html`, and
  `static/app.js`.
- Authentication and notification behavior covered by `webapp.py` and
  `notifier.py`.
- Authentication tests and user-account OpenSpec requirements.
