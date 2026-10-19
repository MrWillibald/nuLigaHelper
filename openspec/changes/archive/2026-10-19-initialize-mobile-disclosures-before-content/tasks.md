# Tasks

## 1. Shared early disclosure controller

- [x] 1.1 Confirm the implementation baseline contains `improve-mobile-usability` and its synchronized specifications; verify the existing responsive markup and default/error metadata before changing the initializer.
- [x] 1.2 Extract the responsive initializer into an isolated classic same-origin external asset that initializes its enclosing disclosure before its contents exist; verify mobile/desktop defaults and the server error marker using synthetic elements with only their summaries constructed.
- [x] 1.3 Preserve deliberate activation, dirty fields, active focus and error guards in one idempotent controller, and remove the standalone `app.js` default initializer; any defensive later binding must preserve unregistered panels' current native state. Extend the dependency-free JavaScript checks to verify repeated invocations, interaction before later bundle loading, resize behavior and fallback under total or partial early-initialization failure.

## 2. Parser-time template integration

- [x] 2.1 Add one shared template helper for the external script immediately after each responsive summary and before its content, including introductions, schedule/roster filters, statistics and authorized maintenance; extend synthetic rendered-page checks to verify ordering, classic synchronous loading, original server defaults and error markers across the relevant viewer tiers.
- [x] 2.2 Preload the same early asset URL from the shared layout and preserve production CSP; verify the served asset and generated script elements work under `script-src 'self'` without inline scripts or policy changes, and inspect cold/warm loads on a larger synthetic roster for repeated-fetch or execution overhead.
- [x] 2.3 Verify rendering in a browser with synthetic offline data while `feedback.js` and `app.js` responses are held pending: at 390px and 700px, introductions/filters/team statistics must have no expanded-content flash; at 701px and a desktop width, the existing desktop defaults must hold. Record reproducible steps and observations, including a separately delayed early asset, reloads and schedule/roster filter responses.
- [x] 2.4 Update `README.MD` to explain early mobile initialization and the retained JavaScript-disabled fallback; verify the wording matches the 700px breakpoint and actual browser observations.

## 3. Integration verification

- [x] 3.1 Exercise native keyboard/touch activation, unsaved edits, active focus, validation errors and device rotation before and after releasing the delayed main bundle; verify no later reset, duplicate handlers, hidden actionable errors, unintended candidate requests or data mutations occur.
- [x] 3.2 Verify JavaScript-disabled pages and failed early-asset responses retain manually usable expanded introduction/filter/team-statistics fallbacks even when the main bundle later loads, including one failed panel invocation followed by a successful invocation on the same page; confirm filter submissions and authorized error correction still work.
- [x] 3.3 Run `test/run_tests.sh`, `openspec validate initialize-mobile-disclosures-before-content --strict` and `git diff --check`; record passing results and ensure changes remain limited to the planned presentation, tests and documentation.
