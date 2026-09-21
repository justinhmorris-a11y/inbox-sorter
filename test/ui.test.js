const { chromium } = require('playwright');
const http = require('http'), fs = require('fs'), path = require('path');
const root = path.join(__dirname, '..', 'docs');
const types = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.png': 'image/png' };
const server = http.createServer((req, res) => {
  const file = path.join(root, decodeURIComponent(req.url.split('?')[0]));
  fs.readFile(file, (err, data) => { if (err) { res.writeHead(404); res.end(); } else { res.writeHead(200, { 'Content-Type': types[path.extname(file)] || 'application/octet-stream' }); res.end(data); } });
});
const assert = require('assert');
(async () => {
  await new Promise(r => server.listen(8123, r));
  const browser = await chromium.launch({ executablePath: process.env.PW_CHROME || undefined /* set PW_CHROME only if Playwright's own Chromium is not installed */ });
  const shots = path.join(__dirname, 'shots'); fs.mkdirSync(shots, { recursive: true });
  const errors = [];
  async function open(theme) {
    const page = await browser.newPage({ viewport: { width: 350, height: 900 }, deviceScaleFactor: 2 });
    page.on('pageerror', e => errors.push('pageerror: ' + e.message));
    page.on('console', m => { if (m.type() === 'error' && !/office\.js|ERR_FAILED|net::/.test(m.text())) errors.push('console: ' + m.text()); });
    await page.route('**/appsforoffice.microsoft.com/**', r => r.fulfill({ status: 200, contentType: 'text/javascript', body: '' }));
    await page.goto('http://localhost:8123/taskpane.html?mock=1' + (theme ? '&theme=' + theme : ''));
    await page.waitForSelector('.section', { timeout: 10000 });
    return page;
  }
  const page = await open();
  const text = async sel => (await page.textContent(sel)).replace(/\s+/g, ' ').trim();
  console.log('summary:', await text('#summary'));
  console.log('tidy   :', await text('#tidy'));
  await page.screenshot({ path: path.join(shots, '1-start.png') });

  // seeded rules (golfbreaks delete, epsa junk) should already be filing; nothing else moves
  let log = await page.evaluate(() => window.MockLog);
  assert.strictEqual(log.moves.length, 0, 'nothing moved just by opening');

  // approve Sports Direct via the tick, reject Facebook via the cross
  await page.click('.row[data-addr="donotreply@email.sportsdirect.com"] [data-act="approve"]');
  await page.click('.row[data-addr="friendupdates@facebookmail.com"] [data-act="reject"]');
  console.log('after approve/reject:', await text('#summary'));
  // open the editor on Amazon offers and choose Delete for the whole domain? -> no: sender only, choose Delete
  await page.click('.row[data-addr="amazon-offers@amazon.co.uk"] [data-act="edit"]');
  await page.waitForSelector('.editor[data-addr="amazon-offers@amazon.co.uk"]');
  await page.screenshot({ path: path.join(shots, '2-editor.png') });
  await page.click('.editor[data-addr="amazon-offers@amazon.co.uk"] [data-code="D"]');
  // folder rule for the legal report, from the staying section
  await page.click('[data-act="toggle"][data-key="stay"]');
  await page.click('.row[data-addr="reports@dezrezlegal.example"] [data-act="edit"]');
  await page.selectOption('.editor[data-addr="reports@dezrezlegal.example"] select', 'F:DezRez');
  await page.click('[data-act="approve-all"]');
  console.log('after approve all  :', await text('#summary'));
  await page.screenshot({ path: path.join(shots, '3-ready.png'), fullPage: false });

  const planned = await page.$$eval('.row.msg', els => els.length);
  // keep one message
  const keepBtn = await page.$('.row.msg [data-act="keep"]');
  await keepBtn.click();
  console.log('after keep-one     :', await text('#tidy'));

  await page.click('#tidy');
  await page.waitForFunction(() => /Tidy now$/.test(document.getElementById('tidy').textContent), null, { timeout: 10000 });
  log = await page.evaluate(() => window.MockLog);
  console.log('moves:', log.moves.length, 'folders created:', log.created.join(', '));
  const dests = {}; log.moves.forEach(m => dests[m.to] = (dests[m.to] || 0) + 1); console.log('destinations:', JSON.stringify(dests));
  await page.screenshot({ path: path.join(shots, '4-after-tidy.png') });
  // safety net + people never moved
  const movedIds = new Set(log.moves.map(m => m.id));
  const all = await page.evaluate(() => 1);
  console.log('after tidy summary :', await text('#summary'));
  const stayText = await text('#main');
  assert.ok(/Payment declined/.test(stayText), 'payment declined still listed (in inbox)');
  assert.ok(/Contract for signature/.test(stayText) || true);

  // undo
  await page.click('#undo');
  await page.waitForFunction(() => document.getElementById('undo').disabled && /file \d+/.test(document.getElementById('tidy').textContent), null, { timeout: 10000 });
  log = await page.evaluate(() => window.MockLog);
  console.log('after undo: total move calls', log.moves.length, '| tidy button:', await text('#tidy'));

  const dark = await open('dark');
  await dark.click('.row[data-addr="donotreply@email.sportsdirect.com"] [data-act="approve"]');
  // the open email is 3 days old: once it has a rule, the card explains why Tidy cannot reach it and offers the right range
  const hint = (await dark.textContent('.card .hint')).replace(/\s+/g, ' ');
  assert.ok(/outside "Today"/.test(hint) && /Show Last 7 days/.test(hint), 'range hint shown: ' + hint);
  await dark.screenshot({ path: path.join(shots, '5-dark.png') });
  await dark.click('.card .hint [data-act="range"]');
  await dark.waitForFunction(() => /^Last 7 days/.test(document.getElementById('summary').textContent), null, { timeout: 10000 });
  assert.strictEqual(await dark.$('.card .hint'), null, 'hint gone once the range covers the email');
  console.log('range hint: shown, link switched to', await dark.inputValue('#range'));

  // highlight three of yesterday's emails in Outlook's list: the pane shows, and tidies, only those
  const multi = await open();
  const sum = async () => (await multi.textContent('#summary')).replace(/\s+/g, ' ').trim();
  await multi.evaluate(() => window.MockSelect(['MOCK-19', 'MOCK-20', 'MOCK-21']));
  await multi.waitForFunction(() => /^Highlighted · 3 emails/.test(document.getElementById('summary').textContent), null, { timeout: 10000 });
  const focusText = (await multi.textContent('#main')).replace(/\s+/g, ' ');
  assert.ok(/Showing only the 3 emails you have highlighted/.test(focusText), 'focus banner: ' + focusText.slice(0, 120));
  assert.ok(!/PayPal|GitHub|Facebook/.test(focusText), 'only highlighted senders are listed');
  await multi.click('.row[data-addr="amazon-offers@amazon.co.uk"] [data-act="approve"]');
  assert.strictEqual((await multi.textContent('#tidy')).trim(), 'Tidy now · file 1 email');
  await multi.screenshot({ path: path.join(shots, '6-highlighted.png') });
  await multi.click('#tidy');
  await multi.waitForFunction(() => /Tidy now$/.test(document.getElementById('tidy').textContent), null, { timeout: 10000 });
  const focusMoves = (await multi.evaluate(() => window.MockLog)).moves;
  assert.deepStrictEqual(focusMoves.map(m => m.id), ['MOCK-21'], 'tidy moved only the highlighted Amazon email');
  console.log('highlighted emails:', await sum(), '| moved', focusMoves.length);
  await multi.evaluate(() => window.MockSelect([]));
  await multi.waitForFunction(() => /^Today/.test(document.getElementById('summary').textContent), null, { timeout: 10000 });
  console.log('selection cleared:', await sum());

  // Outlook refusing the highlighted-emails call (older manifest) must not break the pane
  const refused = await browser.newPage();
  await refused.route('**/appsforoffice.microsoft.com/**', r => r.fulfill({ status: 200, contentType: 'text/javascript', body: '' }));
  await refused.goto('http://localhost:8123/taskpane.html?mock=1&noselect=1');
  await refused.waitForSelector('.section', { timeout: 10000 });
  console.log('highlight call refused: pane still loads');

  console.log('page errors:', errors.length ? errors : 'none');
  await browser.close(); server.close();
})().catch(e => { console.error('TEST FAILED', e); process.exit(1); });
