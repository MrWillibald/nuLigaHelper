## Why

The recently added passwordless authentication pages do not use the established form-card design consistently and split requesting and entering a code across separate pages. Registration and login should instead present the same clear two-step flow, validate contact data reliably, and make the primary completion action visually distinct from requesting a code.

## What Changes

- Replace the separate authentication and generic code-entry pages with matching, responsive registration and login cards that keep code request and code confirmation on one page.
- Add labelled fields with short German descriptions, an explicit E-Mail/SMS contact selector, and a country-calling-code selector for SMS numbers.
- Add server-side e-mail and international phone-number validation and canonicalization, shared by authentication and person contact-data writes; browser validation remains an early usability aid rather than the source of truth.
- Send a six-digit, single-use code through either selected channel and reveal the confirmation step after a privacy-safe request response.
- Preserve registration consent, approval gates, CSRF protection, rate limits, token expiry, single-use consumption, and account-enumeration resistance.
- Give the code-request action a subdued secondary treatment and the final registration/login action the existing saturated primary treatment.
- Redirect active accounts to the Heimspielplan after successful login; verified registrations awaiting approval continue to see only their status.
- **BREAKING** Replace the e-mail magic-link-first login and verification experience with code entry so e-mail and SMS follow the same visible interaction.

## Capabilities

### New Capabilities

None.

### Modified Capabilities

- `user-accounts`: Change registration and passwordless login to a unified two-step code flow for both e-mail and SMS, with validated and canonical contact addresses and defined post-login navigation.

## Impact

- Authentication routes and contact lookup/token issuance in `webapp.py` will change, including compatibility handling for the existing `/login/code`, `/registrieren/code`, and signed-link routes.
- `templates/login.html`, `templates/register.html`, the generic code template, `static/style.css`, and a small progressive-enhancement portion of `static/app.js` are affected.
- Contact validation should be centralized and reused by registration plus person create/edit endpoints so stored addresses remain comparable and unique.
- A maintained phone-number parsing/validation dependency is expected for E.164 normalization; e-mail validation may use a focused dependency or an equivalently robust centralized validator.
- Authentication and web-interface tests must cover both channels, validation failures, privacy-safe responses, responsive form states, code replay/expiry, and redirects.
- `README.MD`, `config_template.json` only if a new setting is introduced, and the repository instructions must remain consistent with the resulting behavior.
