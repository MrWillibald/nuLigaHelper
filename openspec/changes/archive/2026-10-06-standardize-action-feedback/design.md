# Design

## Context

See `proposal.md` for the motivation and `specs/action-feedback/spec.md` for the behavior contract. This is a cross-cutting presentation change, so a design document is required.

`templates/base.html` already includes one `#toast` on every page. `static/app.js` replaces its text and hides it after 2.6 seconds; `static/style.css` fixes it at the bottom center. Schedule assignment handlers use it, cake settings write inline text, and management forms use Flask flashes rendered only by `templates/persons.html`. Responsible-team changes reload after 500 milliseconds and MV changes reload immediately. Registration decisions and login/logout redirect without explicit confirmation.

Assignments retain their current per-slot compare-and-swap behavior. Replacing an occupant is a confirmed release followed by a separate claim; a later failure does not undo the release. Saved occupancy/progress, candidate recovery, authentication privacy and no-JavaScript authentication remain governed by the existing capabilities. The `message-templates` capability concerns dispatched notifications; this design does not change that catalog or its triggers.

## Goals / Non-Goals

**Goals:**

- Consolidate rendering, severity, lifecycle and navigation transfer behind one feedback interface, without replacing the current Flask/vanilla-JavaScript structure.
- Keep confirmed results separate from advisory information, refused writes and unavailable follow-up reads.
- Preserve server-rendered form operation and provide a readable shared fallback when JavaScript is unavailable.

**Non-Goals:**

- Notification history, cross-tab broadcasting, undo actions, or changes to e-mail/SMS delivery.
- Atomic helper replacement, new assignment rules, permission changes or database schema work.
- A management-form redesign, draft recovery, or changes to filtering/navigation beyond carrying the action result.
- Replacing contextual field errors, permanent status messages or the existing pre-deletion confirmation.

## Decisions

### 1. Extend the existing shared feedback layer with explicit severities

Replace the boolean success/error parameter with a descriptor containing plain `message`, `severity` (`success`, `info`, `warning`, `error`) and an optional allowlisted local action such as sign-in or refresh. Keep a single manager in the shared JavaScript and one feedback region in the base template. Normalize existing Flask `ok` categories to success during migration and update warning producers to use warning explicitly.

All severities follow the management badges and delete-button treatment with bold text, a flat tinted background and no border or shadow. Success reuses the active badge's `--ok-bg`/`--ok-fg`; information reuses the team badge's `#eef3f8` and navy text; errors reuse the delete button's `--err-bg`/`--err-fg`. Warnings use light amber (`#fff0cf`, text `#774709`). Pills use 24px corners. Include a German severity label or suitable icon/text cue so color is not the only distinction. The scope of this color rule is feedback; destructive action buttons are not redesigned.

Prefer extending the current component to introducing a toast library: all required behavior fits the existing shared layout and JavaScript, and a new dependency would add operational cost without changing the result. Treating every non-success message as an error is rejected because successful saves can carry advisories.

### 2. Use one rendering pipeline for client results and server forms

`base.html` consumes Flask flashes and renders escaped feedback entries into the shared region. Initially render that region within the document flow; JavaScript enhances it into the floating bubble region with timers and announcements. Neither presentation includes a close button. Without JavaScript it remains readable, does not time out or obstruct controls. Remove the competing flash loop from `persons.html` and transient confirmation elements from auth/cake templates when those outcomes use the shared region. Keep field errors and persistent instructions in place.

Route all normal form results through this mechanism. Add missing success flashes for registration decisions and successful login/logout. For logout, clear the old session and then create only the intentional logout confirmation. Generic code-request messages become informational feedback on the returned auth page; the code-entry guidance remains local. Contact verification confirms verification separately from the persistent awaiting-approval state.

Distinguish HTML form failures from API failures at their entry points. Use a shared HTML result/error helper where forms currently return raw API JSON, retaining required refusal statuses and authorization checks. Continue returning the existing JSON shape and status to API callers. Rendering feedback must never introduce additional access to roster data.

Management contact/birth-date validation currently uses ordinary flashes. Classify these as field validation before converting other flashes to action results: re-render the authorized person form with keyed local explanations, rather than making the bubble the only error location. Do not transfer submitted birth dates/contact values through browser feedback storage or broaden access to another person's fields. This is local error rendering, not draft recovery or a management-form redesign.

Keeping page-specific banners alongside bubbles was considered and rejected: it would duplicate outcomes and preserve the location inconsistency. Moving all page state into bubbles was rejected because errors beside a field and retry/status controls must remain available in context.

### 3. Carry JavaScript outcomes across required navigation once

Normal POST/redirect feedback uses Flask's existing one-use flash delivery. For JavaScript actions requiring a reload or sign-in navigation, store a minimal feedback envelope in `sessionStorage` immediately before navigation. It contains only validated feedback descriptors, the intended same-origin destination path, entry identifiers and an expiry time of 60 seconds. On the destination page, remove the envelope before parsing/displaying it, reject malformed, expired or mismatched envelopes, and start each visible timer afresh.

Transfer only the results associated with the action causing navigation; do not serialize unrelated active notifications, API payloads or form values. Preserve a result once, rather than briefly showing it on the old page and then announcing it again on the destination. Clear obsolete transfer state at authentication boundaries, while allowing the explicitly transferred expiry notice and the current login/logout outcome. Do not create durable browser notification history.

Catch unavailable browser storage. If a safe transfer cannot be made, keep the result readable on the current page, reconcile affected state where possible, and offer explicit refresh/sign-in instead of automatically destroying the result. This uses the existing action links and avoids making feedback dependent on storage availability.

An authenticated server bridge for every JavaScript outcome was considered but would add another endpoint/request and a new failure path. Arbitrary message query parameters and durable local storage were rejected because of replay, untrusted copy and retention concerns. Replacing all required reloads with broader page-state updates would expand this change beyond feedback.

### 4. Give every feedback severity the same automatic reading period

Use a bounded scrollable stack with the newest outcome visible at its bottom edge and older unread messages retained above it. Reveal newly added outcomes without deleting earlier entries or moving keyboard focus. Expire success, information, warning and error entries after five seconds of unpaused visible time; pause on hover, focus within the notification, while the document is hidden, or when the entry is outside the stack's visible region. Entries have no close button. Older unread errors cannot hide a newer result behind a queue. This uniform automatic lifetime follows the user's explicit preference and replaces the initial persistent warning/error design.

Bound the region's width and height to the visible viewport, wrap long text, and allow contained scrolling if several long notifications need space. The region itself must not intercept the surrounding page; any explicit refresh/sign-in links remain clickable and pause expiry while focused.

The current overwrite-and-reset behavior was rejected because it can hide an unread warning. A sequential queue or a fixed visible-entry limit with queued overflow was rejected because earlier unread results could delay later confirmations. A bounded scrollable stack makes new results visible while retaining earlier feedback until its reading period has elapsed.

### 5. Report the actual operation outcome, including partial replacements

Update action handlers to produce action-specific feedback after inspecting the saved response. Use the existing server `warning` on successful game claims, rather than inferring advisory status from stale option classes. A warning message begins with the confirmation that the duty was saved, followed by the advisory. A release uses release wording; successful registration rejection uses success severity.

Preserve the current reconciliation paths for 409 conflicts, saved cake configuration, progress and candidate refresh. Carry an error through any required reload. Apply session-expiry feedback and sign-in recovery to responsible-team and MV changes as well as assignment/cake writes.

Track the confirmed release step during replacements. A subsequent rejected claim reports that the prior assignment was released and the replacement was refused; an unavailable response reports that the replacement outcome is unconfirmed. Network and uninterpretable-response failures must advise checking current state without asserting that the server did not save. These are message/reconciliation changes, not a new transaction boundary.

A failed candidate refresh after a confirmed save remains a local read/recovery problem. Keep the success notification and saved occupancy/progress; retain disabled incomplete pickers and their retry/sign-in controls. Do not generate a second bubble purporting to reverse the confirmed save.

### 6. Support accessible viewport placement and existing overlays

Announce success/information/warnings through polite status semantics and errors through alert semantics, without automatically moving focus. Use one announcement path per logical notification; avoid replaying server-rendered messages through a second live region. Provide visible keyboard focus and honor reduced-motion preferences.

Keep bottom-center placement with narrow-screen margins, a constrained width, safe-area padding and visible-viewport updates when the on-screen keyboard or zoom changes the usable area. Reuse the project's existing visible-viewport approach from task-help positioning where helpful.

Native dialogs and task popovers occupy the browser top layer, so raising a CSS `z-index` alone is insufficient. Use a non-modal manual popover for the enhanced feedback region where supported, with the fixed-position fallback, and do not autofocus it. Account for the occupied bottom feedback area when positioning task-help popovers so focused or pinned guidance and new feedback can remain readable together. Retain the description's existing focus and scroll position; do not invoke its focus-returning dismissal merely to show an asynchronous result. Management dialog submissions already navigate, and their results are delivered after that dialog is gone. Do not treat unsubmitted team-picker drafts as saved actions or close a blocking management dialog merely to show unrelated feedback.

### 7. Keep feedback transfer and rendering minimal

Render server entries with ordinary Jinja escaping and client entries with `textContent`. Validate severity and allowlisted local actions; cap transferred entry count and text length, treating invalid transport data as unusable rather than injecting it. Do not carry contact values, full birth dates, exact ages, authentication codes/challenges or CSRF/session tokens. Existing generic authentication request wording remains the same for all account/limit outcomes.

No extension of `messages.py` is required: it remains the authoritative outbound notification catalog. Browser feedback is UI copy, with small shared German messages in the feedback layer and domain explanations supplied by existing responses. Never change dispatched notification wording or recipients as a side effect of normalizing UI feedback.

## Risks / Trade-offs

- Several unread outcomes can cover page content -> Bound and scroll the region, reveal the newest result, and expire every entry after its five-second visible reading period.
- Browser storage can be blocked or stale -> Use a one-use short-lived same-tab envelope, discard mismatches, and preserve current-page feedback with explicit recovery when transfer is unavailable.
- A write can complete without a readable response -> Use unconfirmed-outcome wording and existing saved-state reconciliation rather than implying rollback.
- Extra live-region events can cause repeated announcements -> Hydrate server-rendered entries through one notification path and consume navigation transfer once.
- Moving feedback can weaken authentication guidance -> Keep field errors, code instructions and persistent approval status local, and verify the no-JavaScript flow and generic response shape.
- Refactoring shared JavaScript can affect synthetic DOM harnesses -> Update test fixtures to exercise the manager, while retaining meaningful saved-state, replacement and candidate-recovery regressions.

## Migration Plan

Implement the shared manager/markup first, then route client and server producers through it, followed by removal of obsolete transient presenters. Add offline regression coverage for lifecycle/navigation and the existing action paths, inspect narrow-screen and keyboard behavior against synthetic data, and update README documentation. Run `test/run_tests.sh` before completion.

Deploy as an ordinary application update with no schema operation or provider configuration change. Roll back by restoring the previous application version; no data transformation or notification replay is required.
