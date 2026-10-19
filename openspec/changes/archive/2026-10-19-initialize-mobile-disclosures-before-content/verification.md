# Implementation verification

Verified on 2026-10-19 with Chrome 155.0.8059.84 and the already installed
Playwright runtime. The opt-in harness uses a temporary SQLite database built
from `test/helpers.py` and the existing synthetic roster fixture, extended to
127 people. Its Flask server listens only on loopback. Production configuration,
contacts, database files and notification providers were not used. No dependency
was installed or added to the application.

## Reproduce the browser checks

From the repository root, point `PLAYWRIGHT_MODULE` at an installed Playwright
ES module and run:

```bash
PLAYWRIGHT_MODULE=/home/manu/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright/index.mjs \
BROWSER_REPORT_PATH=/tmp/disclosure-browser-report.json \
BROWSER_SCREENSHOT_DIR=/tmp/disclosure-browser-screenshots \
node test/browser_disclosures.mjs
```

`BROWSER_EXECUTABLE` and `PYTHON_EXECUTABLE` can override the defaults of
`/usr/bin/google-chrome` and `./venv/bin/python`. Report/screenshot paths are
optional. This is an explicit browser verification command, independent of the
dependency-free regular test suite. In the restricted workspace it needed the
normal sandbox escalation to open its temporary local test socket.

The final command passed all 36 check groups. To reproduce an individual case
manually with the same synthetic app, hold `feedback.js` and `app.js` responses
pending, serve `disclosures.js` normally, and inspect the page before releasing
the held responses. The harness waits until the footer script is actually
parsed; a speculative loader request alone does not establish that parsing has
reached the end of a larger roster.

## Initial presentation and policy

Schedule, helper management and statistics passed at 390, 700, 701 and 1280 CSS
pixels. While the main bundles remained pending, mutation observations recorded
the first constructed disclosure bodies and animation-frame observations
recorded their native open/visible state. At 390 and exactly 700 pixels none of
the mobile introduction, filter or statistics bodies appeared expanded. At 701
and 1280 pixels introductions, filters and team statistics were expanded, while
other statistics sections and ordinary maintenance stayed closed. Releasing the
bundles preserved these states. Mobile and 701-pixel roster screenshots were
also inspected.

With the early asset separately held pending, the first summary was present
but its body remained unparsed, including throughout an additional 120 ms
observation interval. Once that asset was released, mobile content was already
collapsed when constructed. Thus the delay exposed only the native header,
without an expanded content frame.

Every browser response used the exact current
`webapp.PRODUCTION_SECURITY_HEADERS` policy, including `script-src 'self'`.
The same-origin asset returned HTTP 200 under that policy; normal checks had
no JavaScript errors. The ordinary production-runtime regression separately
verifies that actual production middleware serves the asset with JavaScript
MIME type and the unchanged security headers. Rendered-response checks cover
the synchronous external script placement, shared preload URL, original
fallback/default attributes and authoritative error markers across viewer
tiers; no inline script or policy relaxation was introduced.

## Interaction, filters and fallback

- Native Enter and touch activation before the main bundles loaded survived
  their eventual release and portrait/landscape breakpoint changes. Independent
  untouched panels continued to follow the responsive defaults.
- Unsaved filter edits and focused fields were preserved. Separate initially
  expanded desktop cases verified dirty and focus protection without deliberate
  summary activation. Reinitializing the same real browser element ten times
  left its event-listener count unchanged.
- Native schedule and roster GET submissions retained selected values, displayed
  their active summary/reset outside the collapsed filter, and retained the
  correct state after reload. Clearing the filter removed the active summary.
- Failed early responses retained expanded introduction/filter/team-statistics
  fallbacks after the main bundle loaded. Native closing/opening, resizing and
  filter submission remained usable. A separate partial-failure test faults only
  the first invocation by adding a test-only query parameter to that script URL;
  the next normal invocation successfully initializes its own panel. The first
  panel keeps its native/manual state and receives no late collapse.
- Removing `matchMedia` retained the same manually usable native fallbacks on all
  three pages, including filter submissions. JavaScript-disabled browser contexts
  independently passed native expansion/touch and offered GET filters.
- Real authorized person-edit POST responses with an invalid future birth date
  retained the submitted date, `aria-invalid` fields and expanded error wrapper.
  With scripts enabled or the early asset failed, errors were checked while the
  main bundles were held, then after release and rotation. JavaScript-disabled
  errors were checked separately. Native correction/submission succeeded in all
  three modes; the correction restored the original synthetic saved date.
- Initialization and general-panel interactions started no `/api/` candidate
  request or mutation request. A hash of every persisted application table was
  unchanged across each initial browser case, before and after releasing scripts.

## Repeated loading and execution cost

A separate browser context with request routing disabled retained normal browser
caching. Its larger roster rendered 128 synchronous invocations of the one asset
URL. Chrome reported exactly one resource entry per navigation:

| Observation | Cold navigation | Warm navigation |
| --- | ---: | ---: |
| Script elements | 128 | 128 |
| Asset resource entries | 1 | 1 |
| Transferred bytes, including headers | 2,249 | 300 |
| Encoded response body bytes | 1,949 | 0 |
| Asset resource duration | 16.0 ms | 12.9 ms |
| Sampled CPU attributed to early asset | 3.80 ms | 1.68 ms |
| DOMContentLoaded | 282.9 ms | 299.0 ms |
| Load event | 456.0 ms | 470.6 ms |

The warm navigation revalidated the existing asset without retransmitting its
body. Repeated script tags did not create a request per person. CPU values came
from Chrome's sampling profiler at a 100 μs interval and are approximate local
observations; navigation timing includes the full synthetic roster and is not a
Raspberry Pi performance benchmark.

## Final validation

- `test/run_tests.sh`: passed, exit status 0, including the extended production
  asset/header regression and dependency-free disclosure checks.
- `openspec validate initialize-mobile-disclosures-before-content --strict`:
  passed.
- `git diff --check`: passed.
- `test/browser_disclosures.mjs`: all 36 browser check groups passed with the
  command above.

The implementation remains scoped to presentation assets/templates, regression
checks and documentation. It introduces no schema, domain, assignment API,
provider or authorization change.
