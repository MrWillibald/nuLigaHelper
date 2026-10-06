# Tasks

## 1. Shared feedback component and navigation lifecycle

- [x] 1.1 Replace binary toast feedback with explicit success/info/warning/error descriptors and a shared bounded scrollable stack; add offline JavaScript checks for every severity's five-second visible timers, hover/focus/visibility pause, no close buttons and a new success remaining visible after multiple unread errors.
- [x] 1.2 Render escaped server feedback through `templates/base.html` and enhance it without duplicate entries or announcements; verify shared rendering on schedule, management and auth pages, plain-text handling of markup-like names, and a readable document-flow no-JavaScript fallback without inert controls in synthetic web tests.
- [x] 1.3 Add orange warning/red error styling, text severity cues, keyboard-accessible recovery links, narrow-screen wrapping, visible-viewport/safe-area handling and reduced-motion behavior; verify markup/placement behavior offline, retain task-help focus/scroll regressions, and inspect a synthetic narrow-screen view with a result arriving while task help has keyboard focus.
- [x] 1.4 Implement one-use 60-second same-tab transfer for action-required reloads/sign-in navigation, destination matching and storage-unavailable recovery; verify full destination lifetimes, no replay, malformed/expired/mismatched data rejection, authentication-boundary cleanup and absence of form/API secrets in transferred data.
- [x] 1.5 Document feedback colors, uniform automatic lifetimes, absence of close buttons and preserved contextual messages in README.MD; verify the text agrees with the action-feedback spec and implemented defaults.

## 2. Schedule action results and truthful reconciliation

- [x] 2.1 Route game/day-duty assignments and releases through the shared manager with distinct German action wording and server-provided orange advisories; extend the existing assignment/progress JavaScript regressions to verify feedback severity together with saved occupancy, progress and sibling-candidate consistency.
- [x] 2.2 Route cake-settings results through shared feedback while keeping field validation and ongoing configuration/staffing/candidate status contextual; verify success, invalid settings, occupied reductions and stale settings in the cake web/JavaScript regressions, with no duplicate transient confirmation.
- [x] 2.3 Carry responsible-team/MV confirmations and assignment conflict errors through required reloads, and handle expiry/sign-in recovery for every AJAX mutation; verify destination feedback once, rollback of refused selections, current-state reconciliation and consistent expiry wording with mocked responses.
- [x] 2.4 Distinguish release-success/claim-refusal from unconfirmed replacement outcomes, and keep confirmed save feedback accurate if candidate refresh fails; extend replacement and candidate-loading regressions to assert both explanatory messages and the retained/reconciled saved state.
- [x] 2.5 Document how successful saves with advisories, interrupted replacements and unconfirmed outcomes are reported in README.MD; verify examples match the corresponding saved-state regression scenarios.

## 3. Management and authentication feedback

- [x] 3.1 Route person creation/editing, membership changes, deletion, deactivation and reactivation through base feedback; remove the management-only flash presenter, classify successful contact removal as warning, and render management contact/birth-date validation beside authorized form fields; extend synthetic management tests to verify one correctly classified action result after redirects and local field errors without exposing private data through transfer.
- [x] 3.2 Add registration approval/rejection confirmation and readable HTML outcomes for ordinary form refusals while retaining API JSON/status and tier/CSRF guards; verify confirmed decisions, forbidden membership changes and malformed form requests through web/auth/refusal regressions without granting additional roster access.
- [x] 3.3 Add successful login/logout and contact-verification confirmation, and move generic code-request information into shared feedback; extend auth tests to verify intentionally identical request feedback across known/unknown/ineligible/throttled contacts, logout cleanup, local field errors, code instructions and persistent awaiting-approval status.
- [x] 3.4 Verify server-submitted feedback and both auth steps remain readable/usable without JavaScript, and the pre-deletion warning remains intact; add focused template/web checks for these progressive-enhancement and confirmation contracts.
- [x] 3.5 Update README.MD with management/auth action coverage and the distinction between contact verification and administrative approval; verify documented behavior against the synthetic auth/management scenarios.

## 4. Integration verification

- [x] 4.1 Inspect synthetic browser flows spanning an expanded schedule card, cake settings, helper maintenance, team MV change, registration decision and login/logout; verify bottom placement when scrolled, orange warnings/red errors, keyboard/zoom/narrow-screen readability and feedback survival across navigation.
- [x] 4.2 Run `test/run_tests.sh` and resolve regressions; verify the full offline suite is green with synthetic databases/secrets and no production configuration or provider dispatch.
- [x] 4.3 Run `openspec validate standardize-action-feedback --strict` and inspect the final diff; verify proposal/spec/design/tasks and README match the implementation, required checks are recorded, and committed debug switches remain disabled.

## Verification notes

- Synthetic web regressions cover escaped one-use server feedback, authorized local
  contact/date errors, HTML refusals, management decisions, generic code requests,
  login/logout cleanup and verification with persistent approval status.
- Offline JavaScript regressions cover visible-time lifecycle, unread stacks,
  navigation transfer/recovery/privacy, partial/unconfirmed replacements, saved
  progress, authoritative advisories and failed post-save candidate refreshes.
- Browser inspection used disposable synthetic accounts and stubbed providers on
  localhost: expanded game/cake cards, occupied cake reduction, helper membership,
  responsible team, team MV, registration rejection and login/logout. At 390×740
  and 320×420 CSS-pixel visible viewports (including the reduced space equivalent
  to zoom), text wrapped and guidance scrolled. The earlier inspection also
  verified keyboard dismissal of the then-present feedback controls. A
  delayed warning arrived with focus inside guidance; focus stayed there and the
  rectangles did not overlap. Offset/resize behavior is also checked offline.
- README describes colors, uniform automatic durations, absence of close buttons, navigation,
  truthful saved/partial/unconfirmed outcomes and contextual status/validation.
- Before the latest lifecycle refinement, the full `test/run_tests.sh` run passed on 2026-10-06 with exit status 0.
  Tests used synthetic databases/secrets and stubbed notification providers.
- Every severity now follows the flat management-badge treatment, with tinted
  backgrounds, bold dark text and no border or shadow. Confirmation reuses the
  active badge colors, information reuses the team badge colors, and errors reuse
  the delete-button colors. Warnings keep their matching amber treatment. Browser
  previews compared the feedback with the actual management CSS classes.
- A subsequent visual refinement uses 24px bubble corners and extra horizontal
  padding for the requested pill style. Its browser preview verified single-line
  and wrapped messages; the round close controls shown then are removed by the
  latest lifecycle refinement.
- Before the latest lifecycle refinement, strict OpenSpec validation and `git diff --check` passed. The final diff was
  reviewed against the planning artifacts and README; `DEBUG_FLAG` and `CHANGE_DAY`
  remain `False`. Existing unrelated message and database-sidecar changes were
  preserved.
- The latest explicit user preference removes feedback close buttons and makes
  every severity expire after five seconds of unpaused visible reading time.
  Hover/focus/hidden-document/out-of-stack pauses, navigation privacy, contextual
  status and the no-JavaScript fallback are retained. Offline regressions verify
  the five-second boundary and hover/focus pauses for every severity, unread-stack
  preservation, no close buttons and a fresh error lifetime after navigation.
  A synthetic browser preview displayed all four severities without close buttons
  and confirmed that all entries disappeared automatically without moving focus.
  The revised full `test/run_tests.sh` suite passed with exit status 0, as did
  strict OpenSpec validation and `git diff --check`.
- The user-approved theme palette matches the site's navy, mint and orange
  accents. A synthetic browser preview checked all four pills against the site's
  navigation colors; all foreground/background pairs meet WCAG AA for normal
  text (the revised amber warning has 6.92:1 contrast). The full offline suite,
  strict OpenSpec validation and `git diff --check`
  passed again after this palette refinement.
- The warning's light amber fill, bold text and absence of border/shadow were
  compared with actual management badge and delete-button classes in a synthetic
  browser preview. The full offline suite, strict OpenSpec validation and
  `git diff --check` passed after this warning refinement.
- The corresponding flat treatment was then applied to confirmation, information
  and error pills. A synthetic browser preview compared all four with the actual
  active/team badges and delete-button styles. All text/background pairs meet
  WCAG AA (confirmation 5.90:1, information 14.48:1, warning 6.92:1, error 7.52:1).
  The full offline suite passed with exit status 0; strict OpenSpec validation and
  `git diff --check` also passed after this final style refinement.
