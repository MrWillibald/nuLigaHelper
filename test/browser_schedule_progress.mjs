/** Opt-in synthetic layout verification using preinstalled Playwright/Chromium.
 * See test/README.md for environment variables and retained evidence commands.
 */
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { createInterface } from 'node:readline';
import { mkdir, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE
  ? pathToFileURL(resolve(process.env.PLAYWRIGHT_MODULE)).href : 'playwright');
const server = spawn(process.env.PYTHON_EXECUTABLE || './venv/bin/python', ['-u', 'test/schedule_progress_fixture.py'],
  { stdio: ['ignore', 'pipe', 'pipe'] });
let errors = '';
server.stderr.on('data', chunk => { errors += chunk; });
const meta = await new Promise((accept, reject) => {
  createInterface({ input: server.stdout }).on('line', line => {
    if (line.startsWith('PROGRESS_FIXTURE ')) accept(JSON.parse(line.slice(17)));
  });
  server.on('exit', code => reject(new Error(`Fixture failed ${code}: ${errors}`)));
});
const origin = `http://127.0.0.1:${meta.port}`;
let browser;
const report = { widths: [], checks: [] };
const prepare = async (page, scale) => page.evaluate(scale => {
  document.documentElement.style.fontSize = `${scale * 100}%`;
  document.querySelector('.past-details').open = true;
}, scale);
const summaries = page => page.locator('.coverage').evaluateAll(nodes => nodes.map(node => ({
  filled: node.dataset.progressFilled, total: node.dataset.progressTotal, hidden: node.hidden,
  count: node.querySelector('[data-progress-count]').textContent,
  percent: node.querySelector('[data-progress-percent]').textContent,
  width: node.querySelector('.coverage-fill').style.width,
  accessible: [...node.querySelector('.coverage-track').attributes].map(attr => [attr.name, attr.value]),
})));
const checkLayout = async (page, width) => {
  const problems = await page.locator('.game-summary').evaluateAll((nodes, width) => {
    const issues = [], rect = node => node.getBoundingClientRect();
    for (const summary of nodes) {
      for (const node of [summary, ...summary.querySelectorAll('*')]) {
        const r = rect(node);
        if (r.width && r.height && (r.left < -1 || r.right > innerWidth + 1
            || node.scrollWidth > node.clientWidth + 1)) issues.push(`overflow: ${node.className}`);
      }
      const info = summary.querySelector('.game-info'), side = summary.querySelector('.game-summary-side');
      const rows = [...info.children].slice(0, info.classList.contains('task-block-info') ? 3 : 4);
      rows.forEach((row, i) => {
        if (i && rect(row).top < rect(rows[i - 1]).bottom - 1) issues.push('information rows overlap');
      });
      if (width <= 900 ? rect(side).top < rect(info).bottom : rect(side).left < rect(info).right)
        issues.push('responsive information/progress order');
      const label = side.querySelector('.disclosure-label');
      if (!label || getComputedStyle(label).position === 'absolute') issues.push('Details is outside normal flow');
      if (Math.abs(rect(label).right - rect(side).right) > 1) issues.push('Details is not right aligned');
      for (const status of [...side.children].filter(node => node !== label && !node.hidden)) {
        if (rect(status).bottom > rect(label).top) issues.push('Details overlaps status');
      }
      const coverage = summary.querySelector('.coverage');
      if (coverage.hidden) continue;
      const count = coverage.querySelector('[data-progress-count]'), percent = coverage.querySelector('[data-progress-percent]');
      const track = coverage.querySelector('.coverage-track'), fill = coverage.querySelector('.coverage-fill');
      const filled = Number(coverage.dataset.progressFilled), total = Number(coverage.dataset.progressTotal);
      const expected = Math.floor((filled * 100 + Math.floor(total / 2)) / total);
      if (percent.textContent !== `${expected} %` || fill.style.width !== `${expected}%`) issues.push('percentage/fill mismatch');
      if (track.getAttribute('aria-valuetext') !== count.textContent
          || track.getAttribute('aria-valuenow') !== String(filled)
          || track.getAttribute('aria-valuemax') !== String(total)) issues.push('accessible values mismatch');
      if (rect(track).top < Math.max(rect(count).bottom, rect(percent).bottom)) issues.push('bar overlaps text');
      if (Math.abs(rect(track).width - rect(side).width) > 1) issues.push('bar does not span column');
      const a = rect(count), b = rect(percent);
      if (a.left < b.right && b.left < a.right && a.top < b.bottom && b.top < a.bottom) issues.push('count overlaps percentage');
    }
    if (document.documentElement.scrollWidth > innerWidth) issues.push('page overflow');
    return issues;
  }, width);
  assert.deepEqual(problems, [], `Layout at ${width}px`);
};
try {
  browser = await chromium.launch({ executablePath: process.env.BROWSER_EXECUTABLE || '/usr/bin/google-chrome',
    headless: true, args: ['--no-sandbox', '--disable-background-networking'] });
  report.browser = browser.version();
  for (const width of [320, 360, 390, 700, 900, 901, 1280]) {
    for (const scale of [1, 2]) {
      console.error(`Checking ${width}px / ${scale * 100}%`);
      const context = await browser.newContext({ viewport: { width, height: 900 }, hasTouch: true, javaScriptEnabled: false });
      await context.addCookies([{ ...meta.cookie, url: origin }]);
      const page = await context.newPage();
      await page.goto(origin + '/');
      assert.equal(await page.locator('.past-details').evaluate(node => node.open), false);
      assert.equal(await page.locator('.filter-disclosure').evaluate(node => node.open), true, 'No-JS filters remain available');
      await prepare(page, scale);
      assert.equal(await page.locator('.game-card').count(), 20);
      assert.equal(await page.locator('.game-card[open]').count(), 0);
      assert.equal(await page.locator('.coverage:not([hidden]) [role=progressbar]').count(), 16);
      assert.equal(await page.locator('.coverage[hidden] [role=progressbar]').count(), 0);
      assert.equal(await page.locator('.coverage-gauge, [data-vacancies], .game-summary svg, .game-summary .chip').count(), 0);
      assert.ok(await page.locator('[data-eligibility-status]:not([hidden])').count());
      assert.ok(await page.locator('.classification-status').count());
      await checkLayout(page, width);
      // Check every card type and past card independently, using native keyboard and touch.
      for (const index of [0, 1, 2, 3, 16, 17, 18, 19]) {
        const card = page.locator('.game-card').nth(index), summary = card.locator('summary');
        await summary.focus();
        assert.equal(await summary.evaluate(node => getComputedStyle(node).outlineStyle), 'solid');
        if (index === 0 && width === 360 && scale === 1 && process.env.BROWSER_SCREENSHOT_DIR) {
          await mkdir(process.env.BROWSER_SCREENSHOT_DIR, { recursive: true });
          await summary.screenshot({ path: resolve(process.env.BROWSER_SCREENSHOT_DIR, 'keyboard-focus-360.png') });
        }
        await page.keyboard.press('Enter');
        assert.equal(await card.evaluate(node => node.open), true);
        assert.equal(await page.locator('.game-card[open]').count(), 1);
        await checkLayout(page, width);
        await summary.tap();
        assert.equal(await card.evaluate(node => node.open), false);
      }
      const cards = await page.locator('.game-summary').evaluateAll(nodes => nodes.map(node => ({
        info: node.querySelector('.game-meta').textContent.trim(), height: node.getBoundingClientRect().height,
        padding: getComputedStyle(node).padding, gap: getComputedStyle(node).gap,
        columns: getComputedStyle(node).gridTemplateColumns,
      })));
      let reference;
      if (process.env.PROGRESS_REFERENCE_DIR) {
        await page.goto(origin + '/__before'); await prepare(page, scale);
        reference = await page.locator('.game-summary').evaluateAll(nodes => nodes.map(node => ({
          height: node.getBoundingClientRect().height, padding: getComputedStyle(node).padding,
          gap: getComputedStyle(node).gap, columns: getComputedStyle(node).gridTemplateColumns,
        })));
        assert.equal(cards[0].padding, reference[0].padding);
        assert.equal(cards[0].gap, reference[0].gap);
        assert.equal(cards[0].columns.split(" ").length, reference[0].columns.split(" ").length);
      }
      report.widths.push({ width, textPercent: scale * 100, overflow: false, overlap: false, cards, reference });
      if (process.env.BROWSER_SCREENSHOT_DIR && [320, 360, 390, 901, 1280].includes(width)) {
        await mkdir(process.env.BROWSER_SCREENSHOT_DIR, { recursive: true });
        for (const [url, prefix] of [[origin + '/', 'schedule'], ...(reference ? [[origin + '/__before', 'reference']] : [])]) {
          await page.goto(url); await prepare(page, scale);
          await page.evaluate(() => {
            document.getElementById('page-navigation').style.visibility = 'hidden';
            document.activeElement.blur();
          });
          await page.mouse.move(0, 0);
          await page.locator('.day-block:not(.past)').first().screenshot({ path: resolve(process.env.BROWSER_SCREENSHOT_DIR, `${prefix}-${width}-${scale * 100}.png`) });
          if (prefix === 'schedule' && width === 320) {
            for (const [index, name] of [[1, 'all-vacant-zero'], [2, 'full-unconfigured'], [3, 'long-warning']]) {
              await page.locator('.day-block:not(.past)').nth(index).screenshot({ path: resolve(process.env.BROWSER_SCREENSHOT_DIR, `${name}-${width}-${scale * 100}.png`) });
            }
            await page.locator('.day-block.past').screenshot({ path: resolve(process.env.BROWSER_SCREENSHOT_DIR, `past-${width}-${scale * 100}.png`) });
          }
        }
      }
      await context.close();
    }
  }
  // Public layout uses the same counts and bar semantics without private controls.
  for (const width of [320, 901, 1280]) {
    const context = await browser.newContext({ viewport: { width, height: 900 }, javaScriptEnabled: false });
    const page = await context.newPage();
    await page.goto(origin + '/'); await prepare(page, 2); await checkLayout(page, width);
    assert.equal(await page.locator('[data-occupant-id], [data-candidate-url], select[data-role], [data-block-assignment]').count(), 0);
    assert.equal(await page.locator('.game-summary [data-progress-count]').count(), 20);
    report.checks.push(`${width}px: guest layout at 200% text, no private identifiers or assignment controls`);
    await context.close();
  }
  for (const width of [360, 1280]) {
    const context = await browser.newContext({ viewport: { width, height: 900 }, reducedMotion: 'reduce' });
    await context.addCookies([{ ...meta.cookie, url: origin }]);
    const page = await context.newPage(), pageErrors = [], writes = [];
    page.on('pageerror', error => pageErrors.push(error.message));
    page.on('request', request => { if (request.method() === 'POST') writes.push(request.url()); });
    let release;
    const pending = new Promise(accept => { release = accept; });
    await page.route(/\/static\/(app|feedback)\.js$/, async route => { await pending; await route.continue(); });
    await page.goto(origin + '/', { waitUntil: 'commit' });
    await page.waitForSelector('.game-summary');
    const snapshot = await (await page.request.get(origin + '/__snapshot')).json();
    const initial = await summaries(page);
    await page.locator('.game-card').first().locator('summary').click();
    release();
    await page.waitForLoadState('load');
    await page.waitForFunction(() => globalThis.nuLigaProgressTools);
    await page.waitForFunction(() => !document.querySelector('.game-card')._candidateLoading);
    assert.deepEqual(await summaries(page), initial, 'Delayed initialization changed saved summaries');
    assert.deepEqual((await (await page.request.get(origin + '/__snapshot')).json()), snapshot);
    assert.deepEqual(writes, []); assert.deepEqual(pageErrors, []);
    assert.equal(await page.locator('.coverage-fill').first().evaluate(node => getComputedStyle(node).transitionDuration), '0s');
    for (const selector of ['.filter-disclosure', '.page-introduction']) {
      assert.equal(await page.locator(selector).evaluate(node => node.open), width > 700);
    }
    await page.evaluate(() => {
      globalThis.nuLigaProgressTools.updateBlockSummary(document.querySelector('.task-block-cake_delivery'), {
        phase: 'cake_delivery', configured: true, cake_quantity: 10000, progress: { filled: 9999, total: 10000 },
      });
    });
    await prepare(page, 2); await checkLayout(page, width);
    report.checks.push(`${width}px: delayed script, read-only candidate refresh, unchanged database, disclosure defaults, reduced motion, large cake counts`);
    await page.goto(origin + '/statistik');
    const badgeLayout = () => page.locator('.gap-list .chip').evaluateAll(nodes => nodes.map(node => {
      const style = getComputedStyle(node), rect = node.getBoundingClientRect();
      return { text: node.textContent, width: rect.width, height: rect.height,
        padding: style.padding, fontSize: style.fontSize, background: style.backgroundColor, color: style.color };
    }));
    const badges = await badgeLayout(); assert.ok(badges.length);
    if (process.env.PROGRESS_REFERENCE_DIR) {
      const css = await (await page.request.get(origin + '/__before.css')).text();
      await page.evaluate(css => {
        document.querySelector('link[rel=stylesheet]').remove();
        const style = document.createElement('style'); style.textContent = css; document.head.append(style);
      }, css);
      assert.deepEqual(await badgeLayout(), badges, 'Statistics badges changed');
    }
    report.checks.push(`${width}px: statistics badges retained`);
    await context.close();
  }
  console.log(JSON.stringify(report, null, 2));
  if (process.env.BROWSER_REPORT_PATH) await writeFile(process.env.BROWSER_REPORT_PATH, JSON.stringify(report, null, 2) + '\n');
} finally {
  await browser?.close(); server.kill();
}
