/** Opt-in Playwright regression using only a synthetic helper-backed app.
 * PLAYWRIGHT_MODULE=/path/to/playwright/index.mjs node test/browser_schedule_visibility.mjs
 * Optional: BROWSER_EXECUTABLE, PYTHON_EXECUTABLE, BROWSER_REPORT_DIR.
 */
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { createInterface } from 'node:readline';
import { mkdir, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

const { chromium } = await import(process.env.PLAYWRIGHT_MODULE
  ? pathToFileURL(resolve(process.env.PLAYWRIGHT_MODULE)).href : 'playwright');
const fixture = String.raw`
import json, logging, os, sys
sys.path.insert(0, os.path.join(os.getcwd(), 'test'))
from test_schedule_visibility import site, set_boundary, TODAY
from nuligahelper import common
from helpers import sign_in
from werkzeug.serving import make_server
app, engine, ids = site()
set_boundary(engine, ids)
common.effective_today = lambda: TODAY
browser = app.test_client()
sign_in(browser, ids['admin'])
cookie = browser.get_cookie(app.config['SESSION_COOKIE_NAME'])
logging.getLogger('werkzeug').setLevel(logging.ERROR)
server = make_server('127.0.0.1', 0, app, threaded=True)
print('FIXTURE ' + json.dumps({'port': server.server_port, 'cookieName': cookie.key, 'cookie': cookie.value}), flush=True)
server.serve_forever()
`;
const server = spawn(process.env.PYTHON_EXECUTABLE || './venv/bin/python', ['-u', '-c', fixture],
  { cwd: process.cwd(), stdio: ['ignore', 'pipe', 'pipe'] });
let errors = '';
server.stderr.on('data', data => { errors += data; });
let browser;
const deadline = setTimeout(() => { console.error('Browser regression exceeded its 90-second deadline'); server.kill('SIGTERM'); process.exit(1); }, 90000);
const output = process.env.BROWSER_REPORT_DIR || '/tmp/nuliga-visibility-browser';
const checks = [];
try {
  const meta = await new Promise((accept, reject) => {
    const lines = createInterface({ input: server.stdout });
    lines.on('line', line => { if (line.startsWith('FIXTURE ')) accept(JSON.parse(line.slice(8))); });
    server.once('exit', code => reject(new Error(`Fixture failed (${code}): ${errors}`)));
  });
  const origin = `http://127.0.0.1:${meta.port}`;
  browser = await chromium.launch({ executablePath: process.env.BROWSER_EXECUTABLE || '/usr/bin/google-chrome',
    headless: true, args: ['--no-sandbox', '--disable-background-networking'] });
  await mkdir(output, { recursive: true });
  for (const js of [false, true]) {
    const context = await browser.newContext({ javaScriptEnabled: js, locale: 'de-DE' });
    context.setDefaultTimeout(10000);
    await context.addCookies([{ name: meta.cookieName, value: meta.cookie, url: origin }]);
    await context.route('**/*', route => route.request().url().startsWith(origin)
      ? route.continue() : route.abort());
    const page = await context.newPage();
    const form = page.getByRole('form', { name: 'Sichtbarkeit der Heimspieltage' });
    for (const width of [320, 360, 640]) {
      console.log(`Checking ${width}px, js=${js}`);
      await page.setViewportSize({ width, height: 900 });
      await page.goto(origin);
      if (width === 640) await page.evaluate(() => { document.documentElement.style.fontSize = '200%'; });
      assert.equal(await form.getByLabel('Heimspieltage ausblenden ab').inputValue(), '2026-12-01');
      await form.getByLabel('Heimspieltage ausblenden ab').focus();
      for (let step = 0; step < 6 && await page.evaluate(() => document.activeElement.tagName) !== 'BUTTON'; step++) {
        await page.keyboard.press('Tab');
      }
      assert.equal(await page.evaluate(() => document.activeElement.textContent.trim()), 'Sichtbarkeit speichern');
      const fits = await page.evaluate(() => {
        const width = document.documentElement.clientWidth;
        return { page: document.documentElement.scrollWidth <= width + 1,
          controls: [...document.querySelectorAll('.visibility-settings input[type=date], .visibility-settings button, .visibility-scope a')]
            .every(element => { const box = element.getBoundingClientRect(); return box.left >= 0 && box.right <= width + 1; }) };
      });
      assert.deepEqual(fits, { page: true, controls: true }, `schedule at ${width}px, js=${js}`);
      await page.screenshot({ path: `${output}/schedule-${width}-js-${js}.png`, fullPage: false });
      await page.getByRole('link', { name: 'Ausgeblendete Spieltage anzeigen', exact: true }).focus();
      await Promise.all([page.waitForURL('**/?show_hidden=1'), page.keyboard.press('Enter')]);
      assert.match(await page.locator('main').textContent(), /Opponent 2/);
      const navigation = page.getByRole('navigation', { name: 'Seitennavigation' });
      await navigation.getByRole('link', { name: 'Statistik', exact: true }).click();
      assert.match(page.url(), /show_hidden=1/);
      if (width === 640) await page.evaluate(() => { document.documentElement.style.fontSize = '200%'; });
      const outstanding = page.locator('details.statistics-disclosure').last();
      await outstanding.locator('summary').focus();
      if (!(await outstanding.getAttribute('open') !== null)) await page.keyboard.press('Enter');
      assert.match(await outstanding.textContent(), /01.12.2026/);
      assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= document.documentElement.clientWidth + 1), true);
      await page.screenshot({ path: `${output}/statistics-${width}-js-${js}.png`, fullPage: false });
      checks.push(`native keyboard links/disclosures and wrapping at ${width}px; js=${js}`);
    }
    // Save and clear use a native form and PRG with no script dependency.
    await page.goto(origin);
    await form.getByLabel('Heimspieltage ausblenden ab').fill('2027-01-01');
    await Promise.all([page.waitForURL(origin + '/'), form.getByRole('button', { name: 'Sichtbarkeit speichern' }).click()]);
    await page.waitForLoadState('load');
    assert.equal(await form.getByLabel('Heimspieltage ausblenden ab').inputValue(), '2027-01-01');
    await form.getByLabel('Heimspieltage ausblenden ab').fill('');
    await form.getByRole('button', { name: 'Sichtbarkeit speichern' }).click();
    await page.waitForLoadState('load');
    assert.equal(await page.locator('.visibility-scope').count(), 0);
    assert.match(await page.locator('main').textContent(), /Opponent 2/);
    await form.getByLabel('Heimspieltage ausblenden ab').fill('2026-12-01');
    await form.getByRole('button', { name: 'Sichtbarkeit speichern' }).click();
    await page.waitForLoadState('load');
    assert.equal(await form.getByLabel('Heimspieltage ausblenden ab').inputValue(), '2026-12-01');
    checks.push(`save/clear/restore via native form; js=${js}`);
    await context.close();
  }
  await writeFile(`${output}/report.json`, JSON.stringify({ browser: browser.version(), checks }, null, 2));
  console.log(checks.map(check => `PASS ${check}`).join('\n'));
} finally {
  clearTimeout(deadline);
  await browser?.close();
  server.kill('SIGTERM');
}
