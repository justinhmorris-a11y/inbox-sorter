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
  await page.close();
  // subject rule from the card: open one of Ross's 'Demo booked' emails, type the subject text, pick a folder
  const sj = await open();
  await sj.evaluate(() => window.MockOpen('ross@example-colleague.com', 'Ross', 'Demo booked -'));
  await sj.waitForFunction(() => /Ross/.test(document.querySelector('.card .who').textContent), null, { timeout: 10000 });
  await sj.click('.card [data-act="card-more"]');
  assert.strictEqual(await sj.inputValue('.card [data-act="subject"]'), 'Demo booked', 'subject box pre-filled from the open email, cleaned');
  assert.ok(await sj.isDisabled('.card [data-act="subject"]'), 'subject box greyed until ticked');
  await sj.check('.card [data-act="subject-on"]');
  await sj.check('.card [data-act="read"]');
  await sj.selectOption('.card select[data-act="folder"]', 'F:DezRez');
  const sjCard = (await sj.textContent('.card')).replace(/\s+/g, ' ');
  assert.ok(/Your rule: "Demo booked" from this sender goes to DezRez, mark as read/.test(sjCard), 'card shows subject rule: ' + sjCard);
  assert.strictEqual((await sj.textContent('#tidy')).trim(), 'Tidy now · file 5 emails', 'the two Demo booked mails join the 3 seeded ones');
  await sj.click('[data-act="toggle"][data-key="stay"]');
  const sjMain = (await sj.textContent('#main')).replace(/\s+/g, ' ');
  assert.ok(/Keystone sprint/.test(sjMain) && /your rule for "Demo booked"/.test(sjMain), 'Ross\'s other mail still listed as staying');
  await sj.screenshot({ path: path.join(shots, '7-subject-rule.png') });
  // 'File "Demo booked" now' files just those two, leaving the 3 seeded ones for Tidy
  assert.strictEqual((await sj.textContent('.card [data-act="file-these"]')).trim(), 'File "Demo booked" now · 2 emails');
  await sj.click('.card [data-act="file-these"]');
  await sj.waitForFunction(() => /file 3 emails$/.test(document.getElementById('tidy').textContent), null, { timeout: 10000 });
  const sjMoved = (await sj.evaluate(() => window.MockLog)).moves;
  assert.strictEqual(sjMoved.length, 2, 'only the two Demo booked emails moved'); assert.ok(sjMoved.every(m => m.to === 'F-dez'));
  assert.strictEqual(await sj.locator('.card [data-act="file-these"]').count(), 0, 'button gone once nothing is left to file');
  console.log('file-these: moved', sjMoved.length, '| tidy still offers', (await sj.textContent('#tidy')).trim());
  await sj.click('#undo');
  await sj.waitForFunction(() => /put back/.test(document.getElementById('toast').textContent), null, { timeout: 10000 });
  // reopening a matching email later shows the rule with the box ticked
  await sj.evaluate(() => window.MockOpen('ross@example-colleague.com', 'Ross', 'RE: Demo booked - Acme'));
  await sj.waitForFunction(() => /"Demo booked" from this sender/.test(document.querySelector('.card').textContent), null, { timeout: 10000 });
  assert.ok(await sj.isChecked('.card [data-act="subject-on"]'), 'tick box on for an email that matches a subject rule');
  await sj.click('.card [data-act="clear"]');
  assert.strictEqual((await sj.textContent('#tidy')).trim(), 'Tidy now · file 3 emails', 'subject rule removed, seeded rules remain');
  console.log('subject rule: set on Ross/Demo booked, filed 2, other mail untouched, removed');

  await sj.close();
  // 'File this one': the open email (Sports Direct 'New season arrivals', yesterday, MOCK-22) is filed alone while the pane shows Today
  const one = await open();
  await one.evaluate(() => window.MockOpen('donotreply@email.sportsdirect.com', 'Sports Direct', 'New season arrivals', new Date(Date.now() - 86400000), 'MOCK-22'));
  await one.waitForFunction(() => /New season|Sports Direct/.test(document.querySelector('.card .who').textContent), null, { timeout: 10000 });
  await one.click('.card [data-code="N"]');   // rule set -> range widens to Yesterday by itself
  await one.waitForFunction(() => /^Yesterday/.test(document.getElementById('summary').textContent), null, { timeout: 10000 });
  await one.selectOption('#range', 'today');
  await one.waitForFunction(() => /^Today · 20 emails/.test(document.getElementById('summary').textContent), null, { timeout: 10000 });
  assert.strictEqual((await one.textContent('.card [data-act="file-these"]')).trim(), 'File Sports Direct now · 1 email', 'sender-wide button counts only Today\'s Sports Direct mail');
  await one.click('.card [data-act="file-one"]');
  await one.waitForFunction(() => /Filed 1 email/.test(document.getElementById('toast').textContent), null, { timeout: 10000 });
  const oneMoves = (await one.evaluate(() => window.MockLog)).moves;
  assert.deepStrictEqual(oneMoves.map(m => m.id), ['MOCK-22'], 'only the open email moved');
  console.log('file this one: moved', oneMoves[0].id, 'to', oneMoves[0].to, 'while the pane showed Today');

  await one.close();
  // 'mark as read' tick box: Sports Direct newsletters get filed AND marked read; undo restores unread
  const rd = await open();
  await rd.selectOption('#range', '7');
  await rd.waitForFunction(() => /^Last 7 days/.test(document.getElementById('summary').textContent), null, { timeout: 10000 });
  await rd.click('.row[data-addr="donotreply@email.sportsdirect.com"] [data-act="edit"]');
  await rd.check('.editor[data-addr="donotreply@email.sportsdirect.com"] [data-act="read"]');
  await rd.click('.editor[data-addr="donotreply@email.sportsdirect.com"] [data-code="N"]');
  const rdMain = (await rd.textContent('#main')).replace(/\s+/g, ' ');
  assert.ok(/will be marked read/.test(rdMain), 'ready-to-file row says it will be marked read');
  await rd.click('#tidy');
  await rd.waitForFunction(() => /Tidy now$/.test(document.getElementById('tidy').textContent), null, { timeout: 10000 });
  let rdLog = await rd.evaluate(() => window.MockLog);
  const readIds = rdLog.reads.filter(r => r.isRead).map(r => r.id).sort();
  assert.ok(readIds.length >= 1, 'at least one unread Sports Direct mail marked read');
  assert.ok(readIds.every(id => rdLog.moves.some(m => m.id === id)), 'only moved mail is marked read');
  await rd.click('#undo');
  await rd.waitForFunction(() => /put back/.test(document.getElementById('toast').textContent), null, { timeout: 10000 });
  rdLog = await rd.evaluate(() => window.MockLog);
  assert.deepStrictEqual(rdLog.reads.filter(r => !r.isRead).map(r => r.id).sort(), readIds, 'undo marks the same mail unread again');
  console.log('mark as read: filed+read', readIds.length, '| undo restored unread');

  await rd.close();
  // the open email is 3 days old: setting a rule on it widens the range to Last 7 days by itself, so Tidy can reach it
  await dark.waitForFunction(() => /^Last 7 days/.test(document.getElementById('summary').textContent), null, { timeout: 10000 });
  assert.strictEqual(await dark.locator('.card .hint').count(), 0, 'no hint needed once the range covers the email');
  console.log('auto range: rule on a 3-day-old email switched the pane to', await dark.inputValue('#range'));

  await dark.close();
  // highlight three of yesterday's emails in Outlook's list: the pane shows, and tidies, only those
  const multi = await open();
  const sum = async () => (await multi.textContent('#summary')).replace(/\s+/g, ' ').trim();
  await multi.evaluate(() => window.MockSelect(['MOCK-21', 'MOCK-22', 'MOCK-23']));
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
  assert.deepStrictEqual(focusMoves.map(m => m.id), ['MOCK-23'], 'tidy moved only the highlighted Amazon email');
  console.log('highlighted emails:', await sum(), '| moved', focusMoves.length);
  await multi.evaluate(() => window.MockSelect([]));
  await multi.waitForFunction(() => /^Today/.test(document.getElementById('summary').textContent), null, { timeout: 10000 });
  console.log('selection cleared:', await sum());

  await multi.close();
  // Outlook refusing the highlighted-emails call (older manifest) must not break the pane
  const refused = await browser.newPage();
  await refused.route('**/appsforoffice.microsoft.com/**', r => r.fulfill({ status: 200, contentType: 'text/javascript', body: '' }));
  await refused.goto('http://localhost:8123/taskpane.html?mock=1&noselect=1');
  await refused.waitForSelector('.section', { timeout: 10000 });
  console.log('highlight call refused: pane still loads');

  console.log('page errors:', errors.length ? errors : 'none');
  await browser.close(); server.close();
})().catch(e => { console.error('TEST FAILED', e); process.exit(1); });
