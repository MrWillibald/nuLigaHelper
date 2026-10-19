# Design

## Context

See `proposal.md` for motivation and the prerequisite mobile implementation. That implementation renders introductions, filters and team statistics as native `<details open>` fallbacks, then calls `initializeResponsiveDisclosures()` from `static/app.js` at the end of `templates/base.html`. Its default calculation uses `matchMedia('(max-width: 700px)')`; deliberate activation, dirty fields, focus and errors protect subsequent viewport changes.

Production permits only same-origin scripts through `script-src 'self'`. The initial state must therefore be applied through an external asset, without relaxing the policy. The existing script tests inspect fully constructed elements after importing the bundle; they cannot detect expanded content appearing earlier during HTML parsing.

This design is required because the change spans shared JavaScript and several templates and introduces a parser-time loading path with security and performance constraints.

## Goals / Non-Goals

**Goals:** Set native disclosure state before parsing the content body, keep the early path independent of feedback and assignment code, and give each disclosure one lifecycle across early and later application loading.

**Non-Goals:** Device detection on the server, disclosure preference persistence, animations, changes to existing captions or filter behavior, deferred fetching of page content, and changing game/day-card candidate loading. This follow-up does not implement or merge its mobile prerequisite.

## Decisions

### Run a small classic external asset between each summary and its contents

Extract the shared disclosure controller into a small asset such as `static/disclosures.js`. Render a classic same-origin script immediately after each responsive disclosure's `</summary>` and before its `.disclosure-content`. Do not use `async`, `defer` or `type="module"`: the content must remain unparsed until this invocation has applied its enclosing disclosure's state. Use a shared Jinja helper or include to keep the insertion consistent across introductions, filters, statistics and authorized maintenance areas.

Each invocation obtains its enclosing `<details>` from `document.currentScript.parentElement` and initializes only that disclosure. The first invocation creates the shared controller; later invocations reuse it. Add a preload for the same asset URL in the shared layout to start fetching before the first panel and reduce the parser pause. Repeated script elements reuse the small asset; verify both cold and warm asset loading and a representative larger roster rather than assuming caching removes all cost.

Keeping the original server `open` attributes retains the expanded fallback if JavaScript is disabled, the asset fails to load or the required viewport API is unavailable. A successful mobile invocation clears the actual `.open` property before any content body exists. Wider viewports retain the original expanded defaults. No CSS loading guard or visual concealment of an open panel is needed.

Removing server `open` would also solve mobile expansion but changes desktop/no-JavaScript defaults. Inline scripts violate the current policy. Moving the full application bundle or a deferred initializer earlier still couples initial state to unrelated work and does not establish the required parser-before-content ordering.

### Use one idempotent controller for early and later loading

Keep per-disclosure state in a shared registry, for example a `WeakMap`, recording deliberate activation and dirty fields. Bind summary clicks, bubbling input/change events and viewport-change handling during the early invocation. The existing focus and error guards remain part of the default calculation. Native activation produces the same click tracking for keyboard and touch; programmatic toggle events do not imply deliberate interaction.

Isolate the repeatedly executed asset so top-level lexical declarations cannot collide. Once initialized, a disclosure must not receive duplicate listeners or another initial-default reset. Remove the standalone default initializer from `app.js` and make the per-header invocations own initial defaults. Any defensive later binding delegates to the shared controller, skips registered elements and preserves the current native state of unregistered, fully parsed elements rather than applying an initial default late. This covers partial failure: one panel's asset invocation can fail while a later invocation successfully creates the controller. If the early controller is entirely unavailable, the main bundle also leaves native server fallback states intact instead of introducing a late collapse.

This preserves deliberate interaction and edits from successful early initialization onwards, including while `app.js` is delayed. Before the early script is available, the native fallback is the operative interface; this design does not promise to record activation events that occurred before its listeners existed.

### Make server error metadata authoritative during parsing

The early script runs before descendant fields exist. Use the existing `data-disclosure-error="true"` on an affected person-maintenance `<details>` to retain its server-rendered open state. Do not depend on querying descendant `aria-invalid` fields for the initial decision. Keep descendant error detection for subsequent viewport changes once the form is parsed. Confirm every responsive validation-error wrapper carries its server marker.

### Verify rendering before the main bundle can participate

Use synthetic offline pages and separate two forms of evidence:

- Rendered-markup checks establish that every targeted header invokes the external script before its content, with no asynchronous attributes and the existing error/default metadata intact.
- Browser checks hold `feedback.js` and `app.js` responses pending while serving the early asset normally. At 390px and exactly 700px, observe the initial native state and visible content before releasing the bundles. Exercise summary activation, edits and focus during this interval, then release them and resize. Check 701px and a desktop width, filter responses, statistics and error returns. Deliberately delay the early asset too and verify its unparsed contents never flash during a successful load; fail it separately to verify the expanded fallback. Include production CSP and JavaScript-disabled checks.

Extend the existing dependency-free JavaScript checks to exercise repeated early invocations, controller reuse, error markers before descendant construction and viewport guards. Keep candidate-fetch spies separate so initializing general panels cannot initiate roster requests.

## Risks / Trade-offs

- A synchronous asset can delay parsing -> Keep it independent and small, preload one shared URL and measure repeated invocations on a larger roster. This change removes panel-content expansion; it does not promise zero layout movement from unrelated page assets.
- The disclosure marker may initially reflect the native expanded fallback while the early asset downloads -> The panel content has not been parsed, so it cannot expand visibly. Preloading reduces this header-only interval.
- Later bundle loading could reset early choices -> One shared registry owns initialization and interaction state; test before and after releasing the delayed bundles.
- Missing descendant fields could hide an error form -> Require the existing server error marker on the wrapper before initialization and verify a real synthetic validation response.
- Asset failure could leave panels unusable -> Keep native HTML defaults and controls, avoid CSS concealment, and ensure the main bundle cannot apply a late reset to any panel that missed early initialization, even when another panel initialized successfully.

## Migration Plan

Start from a checkout containing the mobile prerequisite and its synchronized main specifications. Implement this change as an ordinary template/static-asset update, update `README.MD`, and run the offline suite plus the rendering checks. Deploy the asset and all referencing templates together. No database migration or provider changes are involved. Rollback restores the preceding templates and JavaScript, bringing back the previous late initialization and expanded fallback without changing stored data.
