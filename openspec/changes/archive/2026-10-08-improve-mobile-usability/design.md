# Design

## Context

See `proposal.md` for motivation and scope. The Flask interface renders Jinja templates using a shared layout, stylesheet and JavaScript bundle. The shared layout already has `body#top`. Game/day-task cards use native disclosures and fetch candidates only on expansion. Statistics uses an expanded card grid and two expanded tables. Roster entries render authorized edit forms inline; `error_person_id` and field errors identify the form requiring correction. Mobile styles already use a 700px breakpoint; the 900px breakpoint also changes tables to stacked rows but does not define this change's initial disclosure defaults.

The change spans four templates and shared interaction, with particular attention to focus, viewport defaults and existing privacy boundaries; a design is therefore required.

## Goals / Non-Goals

**Goals:** Use one small, reusable disclosure enhancement across the affected pages. Keep server-rendered content and form submissions authoritative. Make failure fallbacks usable and keep deliberate user interaction stable during rotation or resizing.

**Non-Goals:** Fetch statistics or person-maintenance data on demand, persist disclosure preferences across visits, introduce a UI library, or change statistics/domain calculations. Existing game-card summaries and administration panels keep their structure.

## Decisions

### Native disclosure with conservative server defaults

Use `<details>/<summary>` for introductions, filters, authorized person-maintenance areas and statistics sections. Annotate only these new elements for a shared initializer in `static/app.js`; do not reuse game-card candidate-loading attributes. Native disclosure supplies manual keyboard/touch operation without a bespoke accordion implementation or new dependency.

Render introductions, filter forms and team statistics open as a desktop/no-JavaScript fallback. Render person-maintenance areas and the other statistics sections closed; only an affected person-maintenance area with returned validation errors is rendered open. Person maintenance keeps this closed default on desktop as well as mobile, while every action remains reachable through native disclosure without JavaScript. Initialize the other responsive defaults using `matchMedia('(max-width: 700px)')` when the page loads. With JavaScript unavailable, mobile may have longer expanded introductory and filter sections, but every form and text remains reachable and team statistics still opens on desktop. This is preferable to user-agent-based server rendering, duplicated markup, or CSS that visually hides an open disclosure while its accessibility state remains open. A short initial expanded layout before enhancement is an accepted trade-off for introductions and filters; initialize early in the existing bundle and avoid collapse animations.

Apply initial viewport defaults before binding user-toggle tracking. On later breakpoint changes, update only untouched disclosures without dirty fields, active focus inside, or validation errors. Programmatic toggle events must not be mistaken for user choices. Track dirty state via input/change and user choices via disclosure activation; use no local storage. Keep multiple sections open independently.

### Narrow explanation boundaries

Wrap the general top-of-page hints on the schedule, persons and statistics pages. The schedule's generic age-rule explanation belongs in this introduction. Do not broadly wrap `.hint` elements: missing-birth-date follow-up, cake settings guidance, validation errors, empty-result actions and live staffing/eligibility warnings must stay in their existing contexts. Task-help popovers, auth forms and legal pages are unaffected. Use concise German content captions that remain unchanged when the disclosure opens or closes: "Hinweise", "Filter", "Bearbeitung", "Details" for game/day cards and "Vergangene Spieltage". Statistics keeps its section titles. Native disclosure state remains accessible; alternate action labels are unnecessary. Schedule game/day cards show a triangle beside "Details": right when closed, down when open. Style that decorative indicator from the native `[open]` state so it also works without JavaScript.

### Keep compact roster identity outside maintenance

Render each person's name and complete team badges in an always-visible card header, with any status already authorized for that viewer. Put profile editing, membership editing and admin record actions in that entry's maintenance disclosure. Show "Bearbeitung" only if at least one action is permitted: own-profile/admin editing or MV membership maintenance qualifies. This lets an MV reach authorized team actions without exposing other profile fields. Maintenance starts closed on mobile, desktop and without JavaScript, with returned validation errors as the sole initial-open exception; it opens through the same native control on every viewport.

Keep existing membership dialogs and native form actions; avoid nested forms and place dialogs so closing the disclosure does not disrupt their lifecycle. Identify cards with internal person IDs, never display names. Use `error_person_id` to force the affected disclosure open on a validation return; the responsive initializer must honor this state and field errors. Verify whether submitted values are currently preserved and retain or explicitly pass them in that validation response if necessary. New-person creation remains in the current administration panel and is not collapsed by this change.

### Summarize filters from validated presentation context

Retain existing GET forms and query parameters, wrapping them in native disclosure. Render an outside summary and clear/reset link only when non-default filters apply. Resolve team names from existing permitted options and use the route's effective filter context, not arbitrary query strings, for administrative filters. Escape all entered name values normally. For unknown active values, use a safe German unknown-filter label without exposing internal IDs; preserve the route's existing empty-result behavior. Opening or closing the wrapper does not submit a form, and clearing filters uses the existing route URL. Do not add auto-submit or search requests.

### Statistics retain current order and calculations

Wrap each current section in a separate disclosure, keeping team statistics first. Use existing `team_stats`, `person_stats` and `gaps` lengths for header counts and label the last count as affected games/day blocks. It is not a sum of missing positions. Existing gap rows already separate physical vacancies, eligibility deficiencies and cake setup needs; preserve that content. Render all three headers even for empty datasets, with an expanded empty state inside each. No new backend statistics aggregation is required.

### Return-to-top progressively enhances a real link

Add a shared "Nach oben" anchor targeting the existing top location, with a small inline arrow icon rather than an external icon library. Without JavaScript it is a footer link. Enhancement changes it to a floating control, shown at `scrollY >= viewport height`, with passive scroll handling and a frame-scheduled visibility update. Use a minimum 44px target, safe-area bottom/right offsets and visible focus. Respect reduced motion when scrolling; after activation move focus to a meaningful top target without hiding the currently focused control and losing keyboard position.

Keep the arrow below feedback and task-help layers. Coordinate its bottom-right position with action feedback so both remain readable; temporarily suppress the arrow if feedback needs that area or an active dialog/popover would conflict. Avoid fixed padding that merely assumes short feedback text. Provide sufficient content clearance at the document bottom. A permanently floating control was considered but adds unnecessary obstruction on short pages.

## Risks / Trade-offs

- Expanded mobile fallback or brief initial layout change -> Native defaults favor availability when JavaScript fails; initialize responsive state early without animations.
- Generic disclosure listeners accidentally triggering candidate fetches -> Use distinct attributes and test that new disclosures never initiate game/day-block requests.
- Collapsing an error or dirty form on resize -> Honor server error state, user interaction, focused descendants and dirty fields.
- Treating hidden private fields as a permission boundary -> Preserve server-side omission and existing route authorization; verify actual response contents for each tier.
- Arrow covering feedback or mobile keyboard UI -> Test narrow portrait/landscape layouts and safe-area placement with live feedback, popovers and dialogs; allow temporary suppression.
- Roster still scales with total result size -> This change reduces visible height rather than adding pagination or shrinking authorized maintenance payloads.

## Migration Plan

Implement and verify the presentation changes with synthetic offline data, then deploy the normal application update. No schema revision, data migration or new dependency is required. Update `README.MD` with responsive defaults and no-JavaScript fallback behavior. Rollback restores the preceding templates and static assets; no stored data or audit history needs rollback.
