## 1. Contact Validation and Canonicalization

- [x] 1.1 Add pinned `email-validator` and `phonenumbers` dependencies, install from `requirements.txt`, and verify both import successfully in the project virtual environment.
- [x] 1.2 Add centralized e-mail normalization and country-code-aware phone parsing helpers with German validation errors, and verify focused offline unit tests cover valid values, malformed values, formatting variations, E.164 output and mismatched international prefixes.
- [x] 1.3 Apply the helpers to registration, login lookup, admin person creation/editing and member self-editing without mutating records on failure, and verify route tests cover invalid writes plus canonically equivalent uniqueness and lookup.
- [x] 1.4 Add a read-only contact preflight report for invalid, changed and colliding existing values, and verify it reports synthetic fixtures without changing the database.

## 2. Unified Authentication Challenges

- [x] 2.1 Issue a nonce and six-digit code for both e-mail and SMS challenges, sign an opaque purpose-bound challenge context, and verify tests cover valid consumption, wrong purpose, wrong code, expiry, replay and two-consumer races.
- [x] 2.2 Add explicit e-mail or SMS delivery for authentication codes without changing ordinary notification fallback behavior, and verify channel-selection tests cover people with one channel, both channels and no channels.
- [x] 2.3 Generate same-shaped dummy challenge state for syntactically valid unknown or ineligible contacts, and verify visible responses contain no account-specific difference while no message or session is created.
- [x] 2.4 Rework `/login` into request-code and confirm-code actions on one route, preserve CSRF and rate limits, redirect active accounts to the Heimspielplan and verified registrants to their status page, and verify route tests cover both channels and both redirects.
- [x] 2.5 Rework `/registrieren` into request-code and confirm-code actions on one route, preserving team selection, publication consent and the approval gate, and verify tests cover missing consent, duplicate contacts, replacement codes for an unverified registration, verification and approver notification.
- [x] 2.6 Redirect the legacy standalone code-entry URLs to their unified forms and retain consumption of already-issued e-mail links for the 15-minute transition window, and verify compatibility tests cover both behaviors.

## 3. Authentication Form Design

- [x] 3.1 Introduce shared Jinja structure for labelled auth fields, descriptions, contact-route radios, the country-calling-code selector and validation messages, and verify rendered controls have associated labels, fieldsets and stable error references.
- [x] 3.2 Rebuild the registration template as a centered two-step card with name, team, consent, contact controls, code request and final registration actions, and verify HTML tests assert field order, German help text and preserved challenge state.
- [x] 3.3 Rebuild the login template with the same contact and challenge components, and verify HTML tests assert parity with registration and removal of the standalone SMS-code link.
- [x] 3.4 Add narrowly scoped responsive auth styles for inputs, selects, radio choices, help/error text, steps and focus states; give code request the subdued secondary style and completion the orange primary style, and verify both actions remain distinguishable without color alone at desktop and narrow widths.
- [x] 3.5 Add progressive enhancement that switches contact inputs, custom country prefix and focus state without owning submission, and verify the forms can request and confirm codes with JavaScript disabled.
- [x] 3.6 Replace or remove the now-unused generic code template and audit auth selectors for leakage into person-management forms, and verify no route or template reference is stale.

## 4. Security and Regression Coverage

- [x] 4.1 Extend authentication tests for invalid e-mail and phone input, canonical matching, code masking, unknown contacts, per-person/per-client limits and stricter SMS limits, and verify `test/test_auth.py` passes standalone.
- [x] 4.2 Add tests that expired, replayed, mismatched and tampered challenges leave account state and sessions unchanged, and verify concurrent consumption still produces exactly one winner.
- [x] 4.3 Update the ordered webapp scenario and authorization/refusal tests for the new form posts without weakening the default-deny route guard or CSRF checks, and verify the affected test files pass standalone and together.

## 5. Documentation and Completion

- [x] 5.1 Update `README.MD` with the two-step code flow, supported contact formats, approval outcome and new dependencies, and verify the documented routes and commands match the implementation.
- [x] 5.2 Update `AGENTS.md` only where authentication or contact invariants changed, keeping the 15-minute token, one-hour session and no-migration rules explicit, and verify no documentation still describes new e-mail logins as magic-link-only.
- [x] 5.3 Run `test/run_tests.sh` with synthetic secrets and databases, verify every test passes offline, and confirm both debug switches remain `False`.
- [x] 5.4 Manually exercise registration and login through E-Mail and SMS at desktop and mobile viewport widths, verifying validation messages, keyboard/focus behavior, visual hierarchy, approval status and the successful Heimspielplan redirect.
