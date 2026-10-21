/** Opt-in synthetic layout check; requires preinstalled Playwright and Chromium.
 * PLAYWRIGHT_MODULE=/absolute/path/to/playwright/index.mjs node test/browser_schedule_gauges.mjs
 * Optional GAUGE_BEFORE_DIR contains pre-change schedule.html/style.css snapshots.
 * Optional BROWSER_REPORT_PATH and BROWSER_SCREENSHOT_DIR retain review evidence.
 */
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { createInterface } from 'node:readline';
import { mkdir, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE
  ? pathToFileURL(resolve(process.env.PLAYWRIGHT_MODULE)).href : 'playwright');
const server = spawn(process.env.PYTHON_EXECUTABLE || './venv/bin/python', ['-u', 'test/schedule_gauge_fixture.py'],
  { stdio: ['ignore', 'pipe', 'pipe'] });
let errors = '';
server.stderr.on('data', chunk => { errors += chunk; });
const meta = await new Promise((accept, reject) => {
  createInterface({ input: server.stdout }).on('line', line => {
    if (line.startsWith('GAUGE_FIXTURE ')) accept(JSON.parse(line.slice(14)));
  });
  server.on('exit', code => reject(new Error(`Fixture failed ${code}: ${errors}`)));
});
const origin = `http://127.0.0.1:${meta.port}`;
let browser;
const report = { widths: [], checks: [] };
try {
  browser = await chromium.launch({ executablePath: process.env.BROWSER_EXECUTABLE || '/usr/bin/google-chrome',
    headless: true, args: ['--no-sandbox', '--disable-background-networking'] });
  browser.on("disconnected", () => console.error("Browser disconnected"));
  report.browser = browser.version();
  for (const width of [320, 360, 390, 700, 900, 1280]) {
    for (const scale of [1, 2]) {
      console.error(`Checking ${width}px / ${scale * 100}%`);
      const context = await browser.newContext({ viewport: { width, height: 900 }, hasTouch: true, javaScriptEnabled: false });
      await context.addCookies([{ ...meta.cookie, url: origin }]);
      const page = await context.newPage();
      page.setDefaultTimeout(15000);
      const metrics = async () => page.locator('.game-summary').evaluateAll(nodes => nodes.map(node => {
        const info = node.querySelector('.game-meta').textContent.trim();
        const badges = node.querySelector('[data-vacancies]');
        return { info, height: node.getBoundingClientRect().height,
          badgeRows: badges && !badges.hidden ? new Set([...badges.children].map(badge => badge.offsetTop)).size : 0 };
      }));
      let before;
      if (process.env.GAUGE_BEFORE_DIR) {
        await page.goto(origin + '/__before');
        if (scale === 2) await page.evaluate(() => { document.documentElement.style.fontSize = '200%'; });
        before = await metrics();
      }
      await page.goto(origin + '/');
      if (scale === 2) await page.evaluate(() => { document.documentElement.style.fontSize = '200%'; });
      const after = await metrics();
      if (before && scale === 1 && width <= 390) {
        for (const index of [0, 1, 2, 3]) assert.ok(after[index].height < before[index].height,
          `Normal mobile card ${index} did not get shorter at ${width}px`);
      }
      assert.equal(after.length, 16, 'Every game and dated phase must remain present');
      const overflow = await page.evaluate(() => {
        const summaries = [...document.querySelectorAll('.game-summary')];
        return summaries.flatMap(node => [node, ...node.querySelectorAll('*')]).filter(node => {
          const bounds = node.getBoundingClientRect();
          return bounds.width && bounds.height && (bounds.right > innerWidth + 1 || bounds.left < -1
            || node.scrollWidth > node.clientWidth + 1 && !['svg', 'circle'].includes(node.tagName));
        }).map(node => ({ tag: node.tagName, text: node.textContent.trim().slice(0, 80) }));
      });
      assert.deepEqual(overflow, [], `Summary overflow at ${width}px / ${scale * 100}%`);
      const overlap = await page.locator('.coverage:not([hidden])').evaluateAll(nodes => nodes.some(node => {
        const list = node.querySelector('[data-vacancies]');
        if (list.hidden) return false;
        const gauge = node.querySelector('.coverage-gauge').getBoundingClientRect();
        const badges = [...list.children].map(badge => badge.getBoundingClientRect());
        return badges.some((badge, index) => badge.left < gauge.right || badges.slice(index + 1).some(other =>
          badge.left < other.right && other.left < badge.right && badge.top < other.bottom && other.top < badge.bottom));
      }));
      assert.equal(overlap, false, `Ring/badge overlap at ${width}px / ${scale * 100}%`);
      if (!await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)) {
        console.error(await page.evaluate(() => [...document.querySelectorAll('body *')].filter(node => {
          const r = node.getBoundingClientRect();
          return r.width && r.height && (r.right > innerWidth + 1 || r.left < -1 || node.clientWidth && node.scrollWidth > node.clientWidth + 1 && getComputedStyle(node).overflowX === "visible");
        }).slice(0, 15).map(node => ({ tag: node.tagName, cls: node.className, text: node.textContent.slice(0, 70),
          scrollWidth: node.scrollWidth, clientWidth: node.clientWidth, width: node.getBoundingClientRect().width, right: node.getBoundingClientRect().right }))));
      }
      assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth), true,
        `Page overflow at ${width}px / ${scale * 100}%`);
      // Native fallback: full staffing/identity are visible before any expansion.
      assert.equal(await page.locator('.game-card[open]').count(), 0);
      assert.equal(await page.locator('.coverage:not([hidden]) [role=progressbar]').count(), 13);
      assert.equal(await page.locator('.coverage[hidden] [role=progressbar]').count(), 0);
      assert.equal(await page.locator('[data-progress-percent]').count(), 0);
      assert.ok(await page.locator('[data-eligibility-status]:not([hidden])').count());
      const first = page.locator('.game-card').first();
      const firstSummary = first.locator('summary');
      await firstSummary.focus();
      assert.equal(await firstSummary.evaluate(node => getComputedStyle(node).outlineStyle), 'solid');
      await page.keyboard.press('Enter');
      assert.equal(await first.evaluate(node => node.open), true);
      assert.equal(await page.locator('.game-card[open]').count(), 1);
      await firstSummary.tap();
      assert.equal(await first.evaluate(node => node.open), false);
      report.widths.push({ width, textPercent: scale * 100, overflow: false,
        cards: after.map((row, index) => ({ ...row, ...(before ? { beforeHeight: before[index].height } : {}) })) });
      // Reload after the touch checks so a transient native tap highlight is
      // absent from the saved layout evidence.
      await page.goto(origin + '/');
      if (scale === 2) await page.evaluate(() => { document.documentElement.style.fontSize = '200%'; });
      // Isolate cards in screenshots; a fixed navigation overlay is unrelated
      // to a full element capture that scrolls through a very tall warning card.
      await page.evaluate(() => {
        document.getElementById('page-navigation').style.visibility = 'hidden';
        document.activeElement.blur();
        document.querySelectorAll('.game-summary').forEach(node => { node.style.webkitTapHighlightColor = 'transparent'; });
      });
      await page.mouse.move(0, 0);
      if (process.env.BROWSER_SCREENSHOT_DIR && [320, 360, 390, 1280].includes(width)) {
        await mkdir(process.env.BROWSER_SCREENSHOT_DIR, { recursive: true });
        await page.locator('.day-block').first().screenshot({ path: resolve(process.env.BROWSER_SCREENSHOT_DIR, `schedule-${width}-${scale * 100}.png`) });
        if (width === 320) {
          await page.locator('.day-block').nth(1).screenshot({ path: resolve(process.env.BROWSER_SCREENSHOT_DIR, `all-vacant-zero-${width}-${scale * 100}.png`) });
          await page.locator('.day-block').nth(2).screenshot({ path: resolve(process.env.BROWSER_SCREENSHOT_DIR, `full-unconfigured-${width}-${scale * 100}.png`) });
        }
        if (width === 320) await page.locator('.day-block').last().screenshot({ path: resolve(process.env.BROWSER_SCREENSHOT_DIR, `long-warning-${width}-${scale * 100}.png`) });
      }
      await context.close();
    }
  }
  // Delayed application code must preserve initial summary and native expansion.
  for (const width of [360, 1280]) {
    const context = await browser.newContext({ viewport: { width, height: 900 }, reducedMotion: 'reduce' });
    await context.addCookies([{ ...meta.cookie, url: origin }]);
    const page = await context.newPage();
    const pageErrors = [];
    page.on('pageerror', error => pageErrors.push(error.message));
    const writes = [];
    page.on('request', request => { if (request.method() === 'POST') writes.push(request.url()); });
    let release;
    const pending = new Promise(accept => { release = accept; });
    await page.route(/\/static\/(app|feedback)\.js$/, async route => { await pending; await route.continue(); });
    await page.goto(origin + '/', { waitUntil: 'commit' });
    await page.waitForSelector('.game-summary');
    const snapshot = await (await page.request.get(origin + '/__snapshot')).json();
    const summaries = async () => page.locator('.coverage').evaluateAll(nodes => nodes.map(node => ({
      filled: node.dataset.progressFilled, total: node.dataset.progressTotal, hidden: node.hidden,
      count: node.querySelector('[data-progress-count]').textContent,
      arc: Number(node.querySelector('.coverage-fill').getAttribute('stroke-dashoffset')),
      badges: [...node.querySelectorAll('[data-vacancies] .chip')].map(badge => badge.textContent),
    })));
    const initial = await summaries();
    await page.locator('.game-card').first().locator('summary').click();
    release();
    await page.waitForLoadState('load');
    await page.waitForFunction(() => globalThis.nuLigaProgressTools);
    // Read-only candidate refreshes must not change progress or write commitments.
    await page.waitForFunction(() => !document.querySelector('.game-card')._candidateLoading);
    assert.deepEqual(await summaries(), initial, 'Delayed initialization changed the saved summaries');
    assert.deepEqual((await (await page.request.get(origin + '/__snapshot')).json()), snapshot);
    assert.deepEqual(writes, []);
    assert.deepEqual(pageErrors, []);
    assert.equal(await page.locator('.coverage-fill').first().evaluate(node => getComputedStyle(node).transitionDuration), '0s');
    assert.equal(await page.locator('.filter-disclosure').evaluate(node => node.open), width > 700);
    await page.evaluate(() => {
      const card = document.querySelector('.task-block-cake_delivery');
      globalThis.nuLigaProgressTools.updateBlockSummary(card, {
        progress: { filled: 9999, total: 10000 }, vacancies: { Kuchenlieferung: 1 },
      });
      const gauge = card.querySelector('.coverage-gauge').getBoundingClientRect();
      const count = card.querySelector('[data-progress-count]').getBoundingClientRect();
      if (count.left < gauge.left + 5 || count.right > gauge.right - 5
          || document.documentElement.scrollWidth > innerWidth) throw new Error('Large cake fraction clips or overflows');
    });
    report.checks.push(`${width}px: delayed script, native disclosure, reduced motion, unchanged database and filters, large cake fraction`);
    if (process.env.GAUGE_BEFORE_DIR) {
      await page.goto(origin + '/statistik');
      const badgeLayout = async () => page.locator('.gap-list .chip').evaluateAll(nodes => nodes.map(node => {
        const style = getComputedStyle(node), rect = node.getBoundingClientRect();
        return { text: node.textContent, width: rect.width, height: rect.height,
          padding: style.padding, fontSize: style.fontSize, background: style.backgroundColor, color: style.color };
      }));
      const currentBadges = await badgeLayout();
      assert.ok(currentBadges.length > 0);
      const oldCss = await (await page.request.get(origin + '/__before.css')).text();
      await page.evaluate(css => {
        document.querySelector('link[rel=stylesheet]').remove();
        const style = document.createElement('style'); style.textContent = css; document.head.append(style);
      }, oldCss);
      assert.deepEqual(await badgeLayout(), currentBadges, 'Statistics badge layout changed');
      report.checks.push(`${width}px: statistics badge sizes, colors and text match original CSS`);
    }
    await context.close();
  }
  console.log(JSON.stringify(report, null, 2));
  if (process.env.BROWSER_REPORT_PATH) await writeFile(process.env.BROWSER_REPORT_PATH, JSON.stringify(report, null, 2) + '\n');
} finally {
  await browser?.close();
  server.kill();
}
