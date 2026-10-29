/**
 * Opt-in browser integration checks for parser-time responsive disclosures.
 *
 * Uses a preinstalled Playwright runtime and Chromium; installs nothing. Run:
 * PLAYWRIGHT_MODULE=/absolute/path/to/playwright/index.mjs \
 *   node test/browser_disclosures.mjs
 * Optional: BROWSER_EXECUTABLE, PYTHON_EXECUTABLE, BROWSER_REPORT_PATH.
 * Only a synthetic helper-backed app on loopback is served; all other browser
 * requests are refused in fault-injection cases. No production config is read.
 */
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { createInterface } from "node:readline";
import { mkdir, writeFile } from "node:fs/promises";
import { resolve } from "node:path";
import { pathToFileURL } from "node:url";

const { chromium } = await import(process.env.PLAYWRIGHT_MODULE
  ? pathToFileURL(resolve(process.env.PLAYWRIGHT_MODULE)).href : "playwright");

const fixtureServer = String.raw`
import hashlib, json, logging, os, sys
from datetime import date
sys.path.insert(0, os.getcwd())
from tests import helpers as h
from nuligahelper import db
from nuligahelper import web as webapp
from tests.web.test_mobile_persons import _roster
from flask import jsonify
from sqlalchemy import select
from werkzeug.serving import make_server

app, engine, ids, names = _roster()
with h.Session(engine) as session:
    support = db.get_support_team(session)
    session.add_all([db.Person(name=f'Browser Fixture {index:03}',
        birth_date=date(1990, 1, 1), teams=[support]) for index in range(120)])
    session.commit()

@app.after_request
def synthetic_csp(response):
    for name, value in webapp.PRODUCTION_SECURITY_HEADERS.items():
        response.headers[name] = value
    return response

@app.get('/__verification_snapshot')
def verification_snapshot():
    # Auth/session transport is synthetic; hash all persisted application data.
    with engine.connect() as connection:
        data = {table.name: [list(row) for row in connection.execute(select(table))]
                for table in db.Base.metadata.sorted_tables}
    return jsonify(hash=hashlib.sha256(json.dumps(data, default=str,
        sort_keys=True).encode()).hexdigest())

clients = {}
for role in ('admin', 'member', 'mv'):
    client = app.test_client()
    h.sign_in(client, ids[role])
    clients[role] = client.get_cookie(app.config['SESSION_COOKIE_NAME']).value
logging.getLogger('werkzeug').setLevel(logging.ERROR)
server = make_server('127.0.0.1', 0, app, threaded=True)
print('BROWSER_FIXTURE ' + json.dumps({'port': server.server_port, 'ids': ids,
    'cookies': clients, 'cookieName': app.config['SESSION_COOKIE_NAME'],
    'csp': webapp.PRODUCTION_SECURITY_HEADERS['Content-Security-Policy']}), flush=True)
server.serve_forever()
`;

const server = spawn(process.env.PYTHON_EXECUTABLE || "./venv/bin/python", ["-u", "-c", fixtureServer],
  { cwd: process.cwd(), stdio: ["ignore", "pipe", "pipe"] });
let serverErrors = "";
server.stderr.on("data", (value) => { serverErrors += value; });
const meta = await new Promise((accept, reject) => {
  const lines = createInterface({ input: server.stdout });
  lines.on("line", (line) => {
    if (line.startsWith("BROWSER_FIXTURE ")) accept(JSON.parse(line.slice(16)));
  });
  server.once("exit", (code) => reject(new Error(`Synthetic server exited ${code}: ${serverErrors}`)));
});
const origin = `http://127.0.0.1:${meta.port}`;
const browser = await chromium.launch({ executablePath: process.env.BROWSER_EXECUTABLE || "/usr/bin/google-chrome",
  headless: true, args: ["--no-sandbox", "--disable-background-networking"] });
const report = { browser: browser.version(), syntheticPeople: 127, checks: [], performance: [] };
const contexts = [];
const check = (label) => { report.checks.push(label); console.log(`PASS ${label}`); };
const panels = "details[data-responsive-disclosure]";
const contentSelector = `${panels} > .disclosure-content`;

async function contextFor(width = 390, options = {}) {
  const context = await browser.newContext({ viewport: { width, height: 900 }, hasTouch: true,
    javaScriptEnabled: options.javaScriptEnabled ?? true });
  contexts.push(context);
  if (options.role !== "guest") await context.addCookies([{ name: meta.cookieName,
    value: meta.cookies[options.role || "admin"], url: origin }]);
  const page = await context.newPage();
  page.requests = [];
  page.errors = [];
  page.on("request", (request) => page.requests.push({ url: request.url(), method: request.method() }));
  page.on("pageerror", (error) => page.errors.push(String(error)));
  if (options.noMatchMedia) await page.addInitScript(() => { window.matchMedia = undefined; });
  await page.addInitScript(() => {
    window.__disclosureSamples = [];
    window.__disclosureFrames = [];
    const capture = () => {
      document.querySelectorAll("details[data-responsive-disclosure]").forEach((details, index) => {
        const content = details.querySelector(":scope > .disclosure-content");
        if (content && !details.__observedContent) {
          details.__observedContent = true;
          window.__disclosureSamples.push({ index, open: details.open,
            desktopOpen: details.dataset.desktopOpen === "true", error: details.dataset.disclosureError === "true" });
        }
      });
    };
    new MutationObserver(capture).observe(document, { childList: true, subtree: true });
    const frame = () => {
      document.querySelectorAll("details[data-responsive-disclosure]").forEach((details, index) => {
        const content = details.querySelector(":scope > .disclosure-content");
        if (content && window.__disclosureFrames.length < 10000) {
          window.__disclosureFrames.push({ index, open: details.open, visible: content.checkVisibility(),
            desktopOpen: details.dataset.desktopOpen === "true", error: details.dataset.disclosureError === "true" });
        }
      });
      requestAnimationFrame(frame);
    };
    requestAnimationFrame(frame);
  });
  return { context, page };
}

async function delayedPage(path, width = 390, options = {}) {
  const pair = await contextFor(width, options);
  const { page } = pair;
  let releaseMain;
  const pendingMain = new Promise((accept) => { releaseMain = accept; });
  let releaseEarly;
  const pendingEarly = new Promise((accept) => { releaseEarly = accept; });
  let earlySeen = false;
  let mainSeen = false;
  await page.route("**/*", async (route) => {
    const url = new URL(route.request().url());
    if (url.origin !== origin) return route.abort();
    if (url.pathname === "/static/feedback.js" || url.pathname === "/static/app.js") {
      mainSeen = true;
      await pendingMain;
    }
    if (url.pathname === "/static/disclosures.js") {
      earlySeen = true;
      if (options.earlyDelay) await pendingEarly;
      if (options.earlyFail || url.searchParams.has("fail-one")) {
        return route.fulfill({ status: 503, contentType: "application/javascript", body: "" });
      }
    }
    if (options.partialFailure && route.request().isNavigationRequest()) {
      const response = await route.fetch();
      const body = (await response.text()).replace(/(<script\s+src=")[^"]*disclosures\.js("[^>]*>)/,
        "$1/static/disclosures.js?fail-one$2");
      return route.fulfill({ response, body });
    }
    return route.continue();
  });
  const navigation = page.goto(origin + path, { waitUntil: "load" });
  // A failed assertion still closes pages with deliberately pending navigation.
  navigation.catch(() => {});
  await page.waitForFunction(() => document.querySelector("details[data-responsive-disclosure] > summary"));
  if (!options.earlyDelay) {
    while (!mainSeen) await page.waitForTimeout(10);
    // The speculative loader can request the footer script before HTML parsing
    // reaches it. Wait for its actual element while its response remains held.
    await page.waitForFunction(() => document.querySelector('script[src*="/static/feedback.js"]'));
  } else {
    while (!earlySeen) await page.waitForTimeout(10);
  }
  return { ...pair, navigation, releaseMain, releaseEarly,
    finish: async () => { releaseEarly(); releaseMain(); await navigation; } };
}

async function state(page) {
  return page.locator(panels).evaluateAll((nodes) => nodes.map((details) => ({
    open: details.open, desktopOpen: details.dataset.desktopOpen === "true",
    error: details.dataset.disclosureError === "true", title: details.querySelector("summary").textContent.trim(),
  })));
}
async function expectedDefaults(page, width, fallback = false) {
  const values = await state(page);
  assert.ok(values.length > 0);
  for (const value of values) assert.equal(value.open,
    value.error || ((fallback || width > 700) && value.desktopOpen), JSON.stringify(value));
  return values;
}
async function noUnexpectedRequests(page) {
  assert.deepEqual(page.requests.filter(({ url, method }) => new URL(url).pathname.startsWith("/api/") || method !== "GET"), []);
  assert.deepEqual(page.errors, []);
}
async function snapshot(context) {
  return (await (await context.request.get(origin + "/__verification_snapshot")).json()).hash;
}

try {
  for (const width of [390, 700, 701, 1280]) {
    for (const path of ["/", "/personen", "/statistik"]) {
      const run = await delayedPage(path, width);
      const before = await snapshot(run.context);
      await expectedDefaults(run.page, width);
      const samples = await run.page.evaluate(() => window.__disclosureSamples);
      assert.equal(samples.length, (await state(run.page)).length);
      for (const sample of samples) assert.equal(sample.open, sample.error || (width > 700 && sample.desktopOpen),
        "The first observed parsed content must already have its responsive state");
      await run.page.waitForTimeout(40);
      const frames = await run.page.evaluate(() => window.__disclosureFrames);
      assert.ok(frames.length > 0, "Capture rendered frames while main bundles are pending");
      for (const frame of frames) assert.equal(frame.visible, frame.error || (width > 700 && frame.desktopOpen),
        "Rendered mobile frames must contain no expanded body");
      if (process.env.BROWSER_SCREENSHOT_DIR && path === "/personen" && [390, 701].includes(width)) {
        await mkdir(process.env.BROWSER_SCREENSHOT_DIR, { recursive: true });
        await run.page.screenshot({ path: resolve(process.env.BROWSER_SCREENSHOT_DIR, `roster-${width}.png`) });
      }
      const response = await run.context.request.get(origin + "/static/disclosures.js");
      assert.equal(response.status(), 200);
      assert.equal(response.headers()["content-security-policy"], meta.csp);
      await run.finish();
      await expectedDefaults(run.page, width);
      await noUnexpectedRequests(run.page);
      assert.equal(await snapshot(run.context), before, "Initialization must not write application data");
      check(`${path} at ${width}px: defaults before/after delayed bundles, parser samples, CSP, no data/API effects`);
      await run.context.close();
    }
  }

  // Delay the early parser-blocking resource separately from the main bundles.
  const early = await delayedPage("/personen", 390, { earlyDelay: true });
  assert.equal(await early.page.locator(contentSelector).count(), 0,
    "While early loading is blocked the first disclosure body must be unparsed");
  await early.page.waitForTimeout(120);
  assert.equal(await early.page.locator(contentSelector).count(), 0);
  early.releaseEarly();
  await early.page.waitForFunction(() => document.querySelectorAll("details[data-responsive-disclosure]").length > 100);
  await expectedDefaults(early.page, 390);
  await early.finish();
  check("Separately delayed early asset blocks content parsing; successful load never exposes expanded mobile content");
  await early.context.close();

  // Deliberate keyboard/touch decisions, untouched resize, edits and focus.
  const interactive = await delayedPage("/personen", 390);
  const panel = interactive.page.locator(panels);
  await panel.nth(0).locator("summary").focus();
  await interactive.page.keyboard.press("Enter");
  assert.equal(await panel.nth(0).evaluate((details) => details.open), true);
  await panel.nth(1).locator("summary").tap();
  await interactive.page.locator("#person-name-filter").fill("Unsubmitted draft");
  await interactive.page.setViewportSize({ width: 844, height: 390 });
  await interactive.page.setViewportSize({ width: 390, height: 844 });
  await interactive.finish();
  assert.equal(await panel.nth(0).evaluate((details) => details.open), true);
  assert.equal(await panel.nth(1).evaluate((details) => details.open), true);
  assert.equal(await interactive.page.locator("#person-name-filter").inputValue(), "Unsubmitted draft");
  assert.equal(await interactive.page.locator("#person-name-filter").evaluate((input) => document.activeElement === input), true);
  await interactive.page.setViewportSize({ width: 1280, height: 900 });
  await interactive.page.setViewportSize({ width: 390, height: 900 });
  assert.equal(await panel.nth(0).evaluate((details) => details.open), true);
  await noUnexpectedRequests(interactive.page);
  const inspector = await interactive.context.newCDPSession(interactive.page);
  const { result: inspected } = await inspector.send("Runtime.evaluate", {
    expression: 'document.querySelector("[data-responsive-disclosure]")', returnByValue: false });
  const beforeListeners = await inspector.send("DOMDebugger.getEventListeners", { objectId: inspected.objectId });
  await interactive.page.evaluate(() => {
    const details = document.querySelector("[data-responsive-disclosure]");
    for (let index = 0; index < 10; index++) window.nuLigaDisclosures.initialize(details);
  });
  const afterListeners = await inspector.send("DOMDebugger.getEventListeners", { objectId: inspected.objectId });
  assert.equal(afterListeners.listeners.length, beforeListeners.listeners.length,
    "Repeated initialization must not bind duplicate browser handlers");
  await inspector.detach();
  check("Native Enter/touch, draft and focused field survive bundle release and portrait/landscape breakpoint changes");
  await interactive.context.close();

  for (const guard of ["dirty", "focus", "untouched"]) {
    const run = await delayedPage("/personen", 1280);
    const filter = run.page.locator(".roster-filter-disclosure");
    if (guard === "dirty") {
      await run.page.locator("#person-name-filter").fill("Dirty without summary activation");
      await run.page.locator("#page-navigation").focus();
    }
    if (guard === "focus") await run.page.locator("#person-name-filter").focus();
    await run.page.setViewportSize({ width: 390, height: 844 });
    assert.equal(await filter.evaluate((details) => details.open), guard !== "untouched");
    await run.finish();
    assert.equal(await filter.evaluate((details) => details.open), guard !== "untouched");
    check(`Independent ${guard} viewport guard survives delayed bundle release`);
    await run.context.close();
  }

  // Actual GET form submissions and reloads retain collapsed filters and outside summaries.
  for (const path of ["/", "/personen"]) {
    const run = await delayedPage(path, 390);
    await run.finish();
    await run.page.locator(".filter-disclosure > summary").tap();
    const input = run.page.locator(path === "/" ? "#person-filter" : "#person-name-filter");
    await input.fill(path === "/" ? "Nobody Fixture" : "Browser Fixture 00");
    await Promise.all([run.page.waitForNavigation(), run.page.locator(".filter-disclosure button[type=submit]").click()]);
    await expectedDefaults(run.page, 390);
    assert.ok(await run.page.locator(".active-filter-summary").isVisible());
    assert.equal(await run.page.locator(".active-filter-summary").locator("xpath=ancestor::details").count(), 0);
    assert.equal(await run.page.locator(path === "/" ? "#person-filter" : "#person-name-filter").inputValue(),
      path === "/" ? "Nobody Fixture" : "Browser Fixture 00");
    await run.page.reload();
    await expectedDefaults(run.page, 390);
    await run.page.locator(".active-filter-summary a").click();
    assert.equal(await run.page.locator(".active-filter-summary").count(), 0);
    check(`${path}: native GET filter response, selected draft, visible summary/reset and reload`);
    await run.context.close();
  }

  // Total/partial failure must not let the later application collapse native fallback.
  for (const options of [{ earlyFail: true }, { partialFailure: true }, { noMatchMedia: true }]) {
    for (const path of ["/", "/personen", "/statistik"]) {
      const run = await delayedPage(path, 390, options);
      if (!options.partialFailure) await expectedDefaults(run.page, 390, true);
      else {
        const values = await state(run.page);
        assert.equal(values[0].open, true);
        assert.ok(values.slice(1).every((value) => !value.open));
      }
      const first = run.page.locator(panels).first();
      await first.locator("summary").tap();
      assert.equal(await first.evaluate((details) => details.open), false);
      await run.finish();
      assert.equal(await first.evaluate((details) => details.open), false);
      await first.locator("summary").tap();
      assert.equal(await first.evaluate((details) => details.open), true);
      await run.page.setViewportSize({ width: 1280, height: 900 });
      await run.page.setViewportSize({ width: 390, height: 900 });
      assert.equal(await first.evaluate((details) => details.open), true);
      if (path !== "/statistik") {
        const filter = run.page.locator(".filter-disclosure");
        if (!await filter.evaluate((details) => details.open)) await filter.locator("summary").tap();
        await run.page.locator(path === "/" ? "#person-filter" : "#person-name-filter").fill("Browser Fixture");
        await Promise.all([run.page.waitForNavigation(), filter.locator("button[type=submit]").click()]);
        assert.ok(await run.page.locator(".active-filter-summary").isVisible());
        if (!options.partialFailure) await expectedDefaults(run.page, 390, true);
      }
      await noUnexpectedRequests(run.page);
      check(`${path}: ${options.earlyFail ? "total early failure" : options.noMatchMedia ? "unavailable viewport API" : "first invocation failure followed by successful initialization"} keeps usable fallback after later bundle`);
      await run.context.close();
    }
  }

  // Native error responses and correction use the real authorized POST route.
  for (const mode of ["enabled", "disabled", "failed"]) {
    const pair = mode === "failed" ? await delayedPage("/personen", 390, { role: "member", earlyFail: true })
      : await contextFor(390, { role: "member", javaScriptEnabled: mode !== "disabled" });
    if (pair.finish) await pair.finish();
    else await pair.page.goto(origin + "/personen");
    const { page, context } = pair;
    if (mode === "disabled") await expectedDefaults(page, 390, true);
    const maintenance = page.locator(`#person-maintenance-${meta.ids.member}`);
    await maintenance.locator("summary").click();
    const form = maintenance.locator("form.person-edit-form");
    let releaseErrorMain;
    if (mode !== "disabled") {
      const pending = new Promise((accept) => { releaseErrorMain = accept; });
      await page.route(/\/static\/(?:feedback|app)\.js$/, async (route) => {
        await pending;
        return route.continue();
      });
    }
    // bypass client validation to obtain the same authoritative server error response.
    await form.locator('[name="birth_date"]').fill("9999-01-01");
    await Promise.all([page.waitForNavigation({ waitUntil: mode === "disabled" ? "load" : "commit" }),
      form.evaluate((node) => node.submit())]);
    if (mode !== "disabled") {
      await page.waitForFunction(() => document.querySelector('script[src*="/static/feedback.js"]'));
    }
    assert.equal(await maintenance.getAttribute("data-disclosure-error"), "true");
    assert.equal(await maintenance.evaluate((details) => details.open), true);
    assert.ok(await maintenance.locator('[aria-invalid="true"]').count());
    assert.equal(await maintenance.locator('[name="birth_date"]').inputValue(), "9999-01-01");
    if (process.env.BROWSER_SCREENSHOT_DIR && mode === "enabled") {
      await page.screenshot({ path: resolve(process.env.BROWSER_SCREENSHOT_DIR, "validation-mobile.png"), fullPage: true });
    }
    if (releaseErrorMain) {
      releaseErrorMain();
      await page.waitForLoadState("load");
      assert.equal(await maintenance.evaluate((details) => details.open), true);
    }
    await page.setViewportSize({ width: 1280, height: 900 });
    await page.setViewportSize({ width: 390, height: 900 });
    assert.equal(await maintenance.evaluate((details) => details.open), true);
    await maintenance.locator('[name="birth_date"]').fill("2000-04-05");
    await Promise.all([page.waitForNavigation(), maintenance.locator('form.person-edit-form button[type="submit"]').click()]);
    assert.equal(await page.locator('[data-disclosure-error="true"]').count(), 0);
    check(`${mode}: authorized validation return stays open with errors/values; native correction succeeds`);
    await context.close();
  }

  for (const path of ["/", "/personen", "/statistik"]) {
    const { context, page } = await contextFor(390, { javaScriptEnabled: false });
    await page.goto(origin + path);
    await expectedDefaults(page, 390, true);
    const first = page.locator(panels).first();
    await first.locator("summary").tap();
    assert.equal(await first.evaluate((details) => details.open), false);
    await first.locator("summary").tap();
    assert.equal(await first.evaluate((details) => details.open), true);
    if (path !== "/statistik") {
      await page.locator(path === "/" ? "#person-filter" : "#person-name-filter").fill("Browser Fixture");
      await Promise.all([page.waitForNavigation(), page.locator(".filter-disclosure button[type=submit]").click()]);
      await expectedDefaults(page, 390, true);
      assert.ok(await page.locator(".active-filter-summary").isVisible());
    }
    check(`${path}: JavaScript-disabled expanded native fallback, touch and filter submission where offered`);
    await context.close();
  }

  // No routing in this context: browser caching remains enabled for measurement.
  const measured = await contextFor(390);
  const profiler = await measured.context.newCDPSession(measured.page);
  await profiler.send("Profiler.enable");
  await profiler.send("Profiler.setSamplingInterval", { interval: 100 });
  for (const temperature of ["cold", "warm"]) {
    await profiler.send("Profiler.start");
    await measured.page.goto(origin + "/personen");
    const { profile } = await profiler.send("Profiler.stop");
    const earlyNodes = new Set(profile.nodes.filter((node) => node.callFrame.url.endsWith("/static/disclosures.js"))
      .map((node) => node.id));
    const sampledEarlyCpuMs = Number((profile.samples.reduce((sum, node, index) => sum
      + (earlyNodes.has(node) ? profile.timeDeltas[index] : 0), 0) / 1000).toFixed(2));
    const metric = await measured.page.evaluate(() => {
      const entries = performance.getEntriesByType("resource").filter((entry) => entry.name.includes("/static/disclosures.js"));
      const nav = performance.getEntriesByType("navigation")[0];
      return { scriptElements: document.querySelectorAll('script[src*="/static/disclosures.js"]').length,
        resourceEntries: entries.length, transferBytes: entries.reduce((sum, entry) => sum + entry.transferSize, 0),
        encodedBytes: entries.reduce((sum, entry) => sum + entry.encodedBodySize, 0),
        assetDurationsMs: entries.map((entry) => Number(entry.duration.toFixed(2))),
        domContentLoadedMs: Number(nav.domContentLoadedEventEnd.toFixed(2)),
        loadMs: Number(nav.loadEventEnd.toFixed(2)) };
    });
    assert.ok(metric.scriptElements > 100, "A representative larger roster must exercise repeated invocations");
    assert.equal(metric.resourceEntries, 1, "Repeated invocation should fetch one shared asset per navigation");
    report.performance.push({ temperature, sampledEarlyCpuMs, ...metric });
    check(`${temperature} roster: ${metric.scriptElements} invocations, ${metric.resourceEntries} resource entry, ${metric.transferBytes} transferred bytes`);
  }
  await measured.context.close();
  console.log(JSON.stringify(report, null, 2));
  if (process.env.BROWSER_REPORT_PATH) await writeFile(process.env.BROWSER_REPORT_PATH, JSON.stringify(report, null, 2) + "\n");
} finally {
  await Promise.allSettled(contexts.map((context) => context.close()));
  await browser.close();
  server.kill();
}
