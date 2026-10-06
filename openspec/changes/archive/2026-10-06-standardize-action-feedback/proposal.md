# Proposal

## Why

User action feedback currently varies between a bottom bubble, page-top banners, inline messages and silent redirects. Some bubbles disappear immediately during reloads, and successful saves with advisory warnings use the same red appearance as failed actions.

## What Changes

- Use the shared bubble at the bottom of the current viewport for action results throughout the website: game and day-duty assignments/releases, responsible-team and MV changes, cake settings, person and membership maintenance, registration decisions, and authentication actions.
- Route both JavaScript results and server-rendered form feedback through the same component, including confirmations displayed after redirects or reloads.
- Give warnings an orange appearance; reserve red feedback for errors. A saved action with an advisory must explicitly say that it succeeded.
- Add explicit confirmation for registration approval/rejection and successful login/logout, and retain generic wording for authentication-code requests.
- Keep field validation, code-entry instructions, candidate loading/retry controls and ongoing account/staffing status in their relevant page context.
- Make bubbles readable and accessible on desktop and narrow screens, with no close button and automatic expiry for every severity after five seconds of unpaused visible reading time.
- Report refused, stale, partially completed and unconfirmed operations accurately without changing assignment or authorization rules.

## Capabilities

### New Capabilities

- `action-feedback`: Defines shared viewport-bottom action notifications, severity semantics, complete action coverage, navigation persistence, accurate outcomes and accessible presentation.

### Modified Capabilities

None. Existing authentication, access-control, assignment, schedule and outbound-message requirements continue to apply; this capability adds a common presentation contract for action feedback.

## Impact

The shared layout and feedback styles/JavaScript, schedule and management/authentication templates, Flask form/redirect feedback, and offline web/JavaScript regression tests are affected. README documentation will describe the resulting feedback behavior. No database migration, new runtime dependency or change to outbound e-mail/SMS delivery is needed.
