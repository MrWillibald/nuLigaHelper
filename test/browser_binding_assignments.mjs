/** Optional real-browser verification; uses installed Playwright/Chromium only.
 * PLAYWRIGHT_MODULE=/path/to/playwright/index.mjs node test/browser_binding_assignments.mjs
 * Serves only synthetic fixtures on loopback; no real config, DB or providers.
 */
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { createInterface } from 'node:readline';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE
  ? pathToFileURL(resolve(process.env.PLAYWRIGHT_MODULE)).href : 'playwright');
const fixture = String.raw`
import json, logging, os, sys
from datetime import date, datetime
from zoneinfo import ZoneInfo
sys.path.insert(0, os.path.join(os.getcwd(), 'test'))
import helpers as h
from nuligahelper import db
from nuligahelper import common
from nuligahelper import web as webapp
from test_sticky_web import fixture
from flask import jsonify, request
from werkzeug.serving import make_server
cleanup, engine, app, ids = fixture()
common.effective_today = lambda: date(2026,10,30)
clock = [datetime(2026,10,30,8,tzinfo=ZoneInfo('Europe/Berlin'))]
db.duty_now = lambda: clock[0]
@app.get('/__binding_snapshot')
def snapshot():
    with h.Session(engine) as session:
        return jsonify(game=[dict(slot=a.slot,role=a.role,person=a.person_id,locked=a.locked_at is not None,request=bool(a.replacement_token)) for a in session.query(db.Assignment).order_by(db.Assignment.id)],
            blocks=[dict(block=a.block_id,slot=a.slot,person=a.person_id,locked=a.locked_at is not None,request=bool(a.replacement_token)) for a in session.query(db.BlockAssignment).order_by(db.BlockAssignment.id)],
            audits=session.query(db.AssignmentAudit).count())
@app.post('/__binding_clock')
def set_clock():
    clock[0]=datetime(2026,10,31,9,tzinfo=ZoneInfo('Europe/Berlin'))
    return jsonify(ok=True)
@app.after_request
def csp(response):
    for key,value in webapp.PRODUCTION_SECURITY_HEADERS.items(): response.headers[key]=value
    return response
cookies={}
for name in ('Original','Replacement','Admin'):
    client=app.test_client(); h.sign_in(client,ids[name])
    cookies[name]=client.get_cookie(app.config['SESSION_COOKIE_NAME']).value
logging.getLogger('werkzeug').setLevel(logging.ERROR)
server=make_server('127.0.0.1',0,app,threaded=True)
print('BINDING_FIXTURE '+json.dumps(dict(port=server.server_port,ids=ids,cookies=cookies,cookieName=app.config['SESSION_COOKIE_NAME'])),flush=True)
server.serve_forever()
`;
const server = spawn(process.env.PYTHON_EXECUTABLE || './venv/bin/python', ['-u', '-c', fixture],
  { cwd: process.cwd(), stdio: ['ignore', 'pipe', 'pipe'] });
let errors = '';
server.stderr.on('data', chunk => { errors += chunk; });
let browser;
try {
  const meta = await new Promise((accept, reject) => {
    createInterface({ input: server.stdout }).on('line', line => {
      if (line.startsWith('BINDING_FIXTURE ')) accept(JSON.parse(line.slice(16)));
    });
    server.once('exit', code => reject(new Error(`Fixture exited ${code}: ${errors}`)));
  });
  const origin = `http://127.0.0.1:${meta.port}`;
  browser = await chromium.launch({ executablePath: process.env.BROWSER_EXECUTABLE || '/usr/bin/google-chrome',
    headless: true, args: ['--no-sandbox', '--disable-background-networking'] });
  const contexts = [];
  async function pageFor(name, javaScriptEnabled = true) {
    const context = await browser.newContext({ viewport: { width: 390, height: 900 }, javaScriptEnabled });
    contexts.push(context);
    await context.addCookies([{ name: meta.cookieName, value: meta.cookies[name], url: origin }]);
    const page = await context.newPage();
    await page.route('**/*', route => new URL(route.request().url()).origin === origin ? route.continue() : route.abort());
    await page.goto(origin);
    return page;
  }
  const original = await pageFor('Original');
  const replacement = await pageFor('Replacement');
  const admin = await pageFor('Admin');
  const csrf = { 'X-CSRF-Token': 'test-csrf-token' };
  const snapshot = async page => (await page.request.get(origin + '/__binding_snapshot')).json();
  async function choose(page, cardId, selector, helper) {
    const card = page.locator(cardId);
    if (!await card.evaluate(node => node.open)) await card.locator(':scope > summary').click();
    const select = card.locator(selector);
    await select.waitFor({ state: 'visible' });
    await page.waitForFunction(element => !element.disabled, await select.elementHandle());
    await select.selectOption(String(helper));
    await page.locator('#assignment-confirmation').waitFor({ state: 'visible' });
    if (await select.getAttribute('data-role') === 'Ordnungsdienst') {
      assert.equal(await select.getAttribute('data-duty-label'),'Ordner');
      const context = await page.locator('[data-confirmation-context]').textContent();
      assert.match(context,/übernimmt Ordner/);
      assert.equal(context.includes('Ordnungsdienst'),false,'raw role keys must not become confirmation text');
      assert.equal(await card.locator('[aria-label="Informationen zu Ordner"]').count(),1);
    }
    return select;
  }
  const targets = [
    { card: `#game-${meta.ids.game}`, selector: 'select[data-role="Ordnungsdienst"][data-slot="0"]', game: true },
    ...['preparation','cake_delivery','cleanup'].map(phase => ({ card: `#block-${meta.ids[phase]}`,
      selector: 'select[data-block-assignment][data-slot="0"]', block: meta.ids[phase] })),
  ];
  for (const [index,target] of targets.entries()) {
    const before = await snapshot(original);
    const select = await choose(original,target.card,target.selector,meta.ids.Original);
    assert.deepEqual(await snapshot(original),before,'opening native confirmation must not save');
    assert.equal(await original.locator('[data-confirmation-cancel]').evaluate(node => document.activeElement === node),true);
    if (index % 2) await original.keyboard.press('Escape');
    else await original.locator('[data-confirmation-cancel]').click();
    await original.locator('#assignment-confirmation').waitFor({ state: 'hidden' });
    assert.deepEqual(await snapshot(original),before,'Cancel/Escape must perform no mutation');
    assert.equal(await select.evaluate(node => document.activeElement === node),true,'Cancel must restore focus');
    await choose(original,target.card,target.selector,meta.ids.Original);
    const saved = original.waitForResponse(response => response.request().method()==='POST'
      && response.url().endsWith('/claim') && response.status()===200);
    await original.locator('[data-confirmation-confirm]').click();
    await saved;
    const assigned = await snapshot(original);
    assert.equal(assigned.game.length + assigned.blocks.length,before.game.length + before.blocks.length + 1);
    const button = target.game ? original.locator(target.card).locator('[data-request-replacement][data-role="Ordnungsdienst"][data-slot="0"]')
      : original.locator(target.card).locator('[data-request-replacement][data-slot="0"]');
    await button.click();
    await original.locator(target.card).locator('[data-duty-state]').filter({hasText:'Ersatz gesucht'}).waitFor();
    const wanted = await snapshot(original);
    const stored = target.game ? wanted.game.find(a => a.role === 'Ordnungsdienst' && a.slot===0)
      : wanted.blocks.find(a => a.block===target.block && a.slot===0);
    assert.equal(stored.person,meta.ids.Original); assert.equal(stored.request,true);
    assert.equal(await button.textContent(),'Ersatzsuche beenden');
    const closing = original.waitForResponse(response => response.url().endsWith('/cancel-replacement') && response.status()===200);
    await button.click(); await closing;
    await button.filter({hasText:'Ersatz suchen'}).waitFor();
    await original.waitForFunction(card => !card._mutationPending, await original.locator(target.card).elementHandle());
    assert.equal(await button.evaluate(node => document.activeElement===node),true,'search closure must preserve button focus, including rerendered cakes');
    const closed = await snapshot(original);
    const retained = target.game ? closed.game.find(a => a.role==='Ordnungsdienst' && a.slot===0)
      : closed.blocks.find(a => a.block===target.block && a.slot===0);
    assert.equal(retained.person,meta.ids.Original); assert.equal(retained.request,false);
    assert.equal(closed.game.length + closed.blocks.length,wanted.game.length + wanted.blocks.length);
    const reopening = original.waitForResponse(response => response.url().endsWith('/request-replacement') && response.status()===200);
    await button.click(); await reopening;
    await button.filter({hasText:'Ersatzsuche beenden'}).waitFor();
    if (target.block===meta.ids.cake_delivery) {
      const nojsSearch = await pageFor('Original',false);
      const searchUrl = origin+`/dienst?block_id=${target.block}&slot=0`;
      await nojsSearch.goto(searchUrl);
      await nojsSearch.getByRole('button',{name:'Ersatzsuche beenden',exact:true}).click();
      const nojsClosed = (await snapshot(nojsSearch)).blocks.find(a => a.block===target.block && a.slot===0);
      assert.equal(nojsClosed.person,meta.ids.Original); assert.equal(nojsClosed.request,false);
      await nojsSearch.goto(searchUrl);
      await nojsSearch.getByRole('button',{name:'Ersatz suchen',exact:true}).click();
      assert.equal((await snapshot(nojsSearch)).blocks.find(a => a.block===target.block && a.slot===0).request,true);
      console.log('PASS non-JavaScript own search closure and reopening keeps responsibility');
    }
    await replacement.reload();
    await choose(replacement,target.card,target.selector,meta.ids.Replacement);
    await replacement.locator('[data-confirmation-confirm]').click();
    await replacement.locator(target.card).locator('[data-duty-occupant]').filter({hasText:'Replacement'}).waitFor();
    const after = await snapshot(replacement);
    const transferred = target.game ? after.game.find(a => a.role==='Ordnungsdienst' && a.slot===0)
      : after.blocks.find(a => a.block===target.block && a.slot===0);
    assert.equal(transferred.person,meta.ids.Replacement); assert.equal(transferred.request,false);
    assert.equal(after.game.length + after.blocks.length,wanted.game.length + wanted.blocks.length);
    console.log(`PASS native confirmation/cancel/focus/search closure/reopening/atomic takeover ${target.game ? 'game' : target.block}`);
  }
  await original.reload();
  const gameCard = `#game-${meta.ids.game}`;
  await choose(original,gameCard,'select[data-role="Verkauf"][data-slot="1"]',meta.ids.Original);
  assert.equal((await admin.request.post(origin+'/api/assignment/claim',{headers:csrf,data:{game_id:meta.ids.game,
    role:'Verkauf',slot:1,person_id:meta.ids.Admin,expected_person_id:null}})).status(),200);
  const winner = await snapshot(original);
  await original.keyboard.press('Escape');
  assert.deepEqual(await snapshot(original),winner,'cancel must preserve a concurrent winner');
  await choose(original,gameCard,'select[data-role="Kasse"][data-slot="0"]',meta.ids.Original);
  assert.equal((await admin.request.post(origin+'/api/assignment/claim',{headers:csrf,data:{game_id:meta.ids.game,
    role:'Kasse',slot:0,person_id:meta.ids.MV,expected_person_id:null}})).status(),200);
  const secondWinner = await snapshot(original);
  const conflict = original.waitForResponse(response => response.url().endsWith('/api/assignment/claim') && response.status()===409);
  await original.locator('[data-confirmation-confirm]').click(); await conflict;
  assert.deepEqual(await snapshot(original),secondWinner,'confirmed stale selection must not overwrite the winner');
  console.log('PASS native concurrent Cancel and confirmed CAS conflict');
  await original.request.post(origin+'/__binding_clock',{headers:csrf});
  await original.goto(origin);
  await choose(original,`#block-${meta.ids.preparation}`,'select[data-block-assignment][data-slot="1"]',meta.ids.Original);
  assert.equal(await original.locator('[data-confirmation-late]').isVisible(),true);
  await original.locator('[data-confirmation-confirm]').click();
  await original.locator(`#block-${meta.ids.preparation} [data-block-slot="1"] [data-duty-state]`).filter({hasText:'Fest eingeteilt'}).waitFor();
  assert.equal((await snapshot(original)).blocks.find(a => a.block===meta.ids.preparation && a.slot===1).locked,true);
  console.log('PASS native late empty claim immediately locks');
  const nojs = await pageFor('Original',false);
  const url = origin+`/dienst?block_id=${meta.ids.cleanup}&slot=2`;
  await nojs.goto(url); await nojs.selectOption('select[name="person_id"]',String(meta.ids.Original));
  const beforeReview = await snapshot(nojs);
  await nojs.getByRole('button',{name:'Weiter zur Bestätigung'}).click();
  assert.deepEqual(await snapshot(nojs),beforeReview);
  await nojs.getByRole('button',{name:'Abbrechen',exact:true}).click();
  assert.deepEqual(await snapshot(nojs),beforeReview);
  await nojs.goto(url); await nojs.selectOption('select[name="person_id"]',String(meta.ids.Original));
  await nojs.getByRole('button',{name:'Weiter zur Bestätigung'}).click();
  await nojs.getByRole('button',{name:'Bestätigen',exact:true}).click();
  assert.equal((await snapshot(nojs)).blocks.find(a => a.block===meta.ids.cleanup && a.slot===2).locked,true);
  console.log('PASS non-JavaScript review/cancel/confirm with late lock');
  for (const context of contexts) await context.close();
} finally {
  if (browser) await browser.close();
  server.kill('SIGTERM');
}
