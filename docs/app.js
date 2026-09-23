/* Inbox Sorter - task pane UI. Vanilla JS, no build step. */
(function () {
  'use strict';
  var E = window.SorterEngine, G = window.SorterGraph;
  var $ = function (id) { return document.getElementById(id); };

  var PUBLIC_DOMAINS = /^(gmail|googlemail|outlook|hotmail|live|msn|yahoo|icloud|me|aol|btinternet|sky|virginmedia|talktalk|protonmail|proton)\.(com|co\.uk|me|net)$/i;
  var SENT_MAX = 3000, SENT_TTL_DAYS = 7, EVIDENCE_PER_REFRESH = 80;

  var S = {
    range: 'today', busy: false, me: null,
    messages: [], senders: {}, rows: [], assessments: {},
    rules: E.emptyRules(), ctx: { sentTo: {}, myDomains: {}, evidence: {} },
    folders: [], selected: null, focus: null,   // focus: the emails highlighted in Outlook's list; while set, the pane shows and tidies only those
    editing: null, domainFlag: {}, readFlag: {}, subjectText: {}, subjectOn: {}, cardMore: false,
    open: { suggest: true, file: true, stay: false }, openDest: {},
    ruleSnapshot: null, toastTimer: null
  };

  var ICON = {
    check: '<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M13.7 3.8a.75.75 0 0 1 0 1.06l-7 7a.75.75 0 0 1-1.06 0l-3.3-3.3a.75.75 0 1 1 1.06-1.06l2.77 2.77 6.47-6.47a.75.75 0 0 1 1.06 0Z"/></svg>',
    cross: '<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M3.7 3.7a.75.75 0 0 1 1.06 0L8 6.94l3.24-3.24a.75.75 0 1 1 1.06 1.06L9.06 8l3.24 3.24a.75.75 0 1 1-1.06 1.06L8 9.06 4.76 12.3a.75.75 0 0 1-1.06-1.06L6.94 8 3.7 4.76a.75.75 0 0 1 0-1.06Z"/></svg>',
    more: '<svg viewBox="0 0 16 16" aria-hidden="true"><circle cx="3.5" cy="8" r="1.25"/><circle cx="8" cy="8" r="1.25"/><circle cx="12.5" cy="8" r="1.25"/></svg>',
    chev: '<svg class="chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M4.2 2.2a.75.75 0 0 1 1.06 0l3.27 3.27a.75.75 0 0 1 0 1.06L5.26 9.8A.75.75 0 0 1 4.2 8.74L6.94 6 4.2 3.26a.75.75 0 0 1 0-1.06Z"/></svg>'
  };

  function esc(v) { return String(v == null ? '' : v).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function plural(n, word) { return n + ' ' + word + (n === 1 ? '' : 's'); }
  function pillClass(code) { return E.BUCKETS[code] ? code : 'F'; }
  function pill(code, ghost) { return '<span class="pill ' + pillClass(code) + (ghost ? ' ghost' : '') + '">' + esc(E.bucketName(code)) + '</span>'; }

  // ------------------------------------------------------------------ boot
  function applyTheme() {
    var dark = false;
    try {
      var t = Office.context.officeTheme;
      if (t && t.bodyBackgroundColor) {
        var hex = t.bodyBackgroundColor.replace('#', '');
        var lum = (parseInt(hex.substr(0, 2), 16) * 299 + parseInt(hex.substr(2, 2), 16) * 587 + parseInt(hex.substr(4, 2), 16) * 114) / 255000;
        dark = lum < 0.4;
      } else { dark = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches; }
    } catch (e) { dark = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches; }
    if (/[?&]theme=dark/.test(location.search)) dark = true;
    document.documentElement.setAttribute('data-theme', dark ? 'dark' : 'light');
  }

  function seedRules() {
    var seed = (window.SORTER_CONFIG && window.SORTER_CONFIG.seedRules) || {};
    return { v: 1, senders: seed.senders || {}, domains: seed.domains || {}, keep: {} };
  }

  function pruneKeep() {
    var cutoff = Date.now() - 14 * 86400000, keys = Object.keys(S.rules.keep);
    keys.forEach(function (k) { if (S.rules.keep[k] < cutoff) delete S.rules.keep[k]; });
    keys = Object.keys(S.rules.keep).sort(function (a, b) { return S.rules.keep[b] - S.rules.keep[a]; });
    keys.slice(40).forEach(function (k) { delete S.rules.keep[k]; });
  }

  async function boot() {
    applyTheme();
    wire();
    showLoading('Connecting to your mailbox...');
    await G.initAuth();
    S.rules = E.normaliseRules(G.loadRules() || seedRules());
    if (!G.rulesInMailbox()) G.saveRules(S.rules);   // only a local backup (or seeds) so far: store them in the mailbox
    pruneKeep();
    S.me = await G.me();
    S.ctx.myDomains = S.me.domains;
    S.ctx.evidence = G.local.get('is.evidence.v1') || {};
    await ensurePeople();
    G.listFolders().then(function (f) { S.folders = f; }).catch(function () { /* loaded again on demand */ });
    watchSelection();
    await refresh();
  }

  async function ensurePeople() {
    var cache = G.local.get('is.sent.v1');
    var fresh = cache && (Date.now() - cache.at) < SENT_TTL_DAYS * 86400000;
    if (cache) S.ctx.sentTo = cache.to;
    if (fresh) return;
    var load = G.sentRecipients(SENT_MAX, function (n) { if (!cache) showLoading('Learning who you write to... ' + n + ' sent emails'); })
      .then(function (r) { S.ctx.sentTo = r.to; G.local.set('is.sent.v1', { at: Date.now(), scanned: r.scanned, to: r.to }); });
    if (!cache) await load;                       // first run: wait, it decides who counts as a person
    else load.then(replan).catch(function () { /* stale cache is fine */ });
  }

  // ------------------------------------------------------------------ data
  function rangeDates(key) {
    var n = new Date(), y = n.getFullYear(), m = n.getMonth(), d = n.getDate();
    if (key === 'today') return [new Date(y, m, d), new Date(y, m, d + 1)];
    if (key === 'yesterday') return [new Date(y, m, d - 1), new Date(y, m, d)];
    var days = parseInt(key, 10) || 7;
    return [new Date(y, m, d - days + 1), new Date(y, m, d + 1)];
  }
  function rangeLabel() { return { today: 'Today', yesterday: 'Yesterday', '7': 'Last 7 days', '30': 'Last 30 days' }[S.range] || 'Today'; }
  function scopeLabel() { return S.focus ? 'Highlighted' : rangeLabel(); }

  // Each refresh gets a number; if the range changes while a read is in flight, the older read's result is thrown away.
  var refreshSeq = 0;
  async function refresh() {
    if (S.focus) { pickCache = {}; readHighlighted(); return; }
    var seq = ++refreshSeq;
    S.busy = true; $('refresh').classList.add('spin');
    try {
      var r = rangeDates(S.range);
      if (!S.messages.length) showLoading('Reading ' + rangeLabel().toLowerCase() + '...');
      var listed = await G.listInbox(r[0], r[1], function (n) { if (seq === refreshSeq && !S.focus) $('summary').textContent = 'Reading... ' + n + ' emails'; });
      if (seq !== refreshSeq || S.focus) return;
      S.messages = listed;
      S.senders = E.buildSenders(S.messages);
      await gatherEvidence();
      if (seq !== refreshSeq) return;
      S.busy = false;
      replan();
    } catch (err) { if (seq === refreshSeq) showError(err); }
    finally { if (seq === refreshSeq) { S.busy = false; $('refresh').classList.remove('spin'); } }
  }

  /** Look at the headers of one message per unfamiliar sender (unsubscribe link, bulk / automated markers). */
  async function gatherEvidence() {
    var todo = Object.keys(S.senders).map(function (a) { return S.senders[a]; }).filter(function (s) {
      if (S.ctx.evidence[s.address] || E.ruleFor(S.rules, s.address) || !E.addressParts(s.address)) return false;
      return E.assessSender(s, S.ctx).kind !== 'people';
    }).sort(function (a, b) { return b.total - a.total; }).slice(0, EVIDENCE_PER_REFRESH);
    if (!todo.length) return;
    var i = 0, done = 0;
    async function worker() {
      while (i < todo.length) {
        var s = todo[i++];
        try { S.ctx.evidence[s.address] = E.readHeaders(await G.messageHeaders(s.latest.id)); }
        catch (e) { /* message moved or deleted meanwhile - skip */ }
        $('summary').textContent = 'Checking senders... ' + (++done) + ' of ' + todo.length;
      }
    }
    await Promise.all([worker(), worker(), worker(), worker()]);
    var keys = Object.keys(S.ctx.evidence);
    if (keys.length > 2000) keys.slice(0, keys.length - 2000).forEach(function (k) { delete S.ctx.evidence[k]; });
    G.local.set('is.evidence.v1', S.ctx.evidence);
  }

  function replan() {
    var p = E.plan(S.messages, S.senders, S.rules, S.ctx, { now: new Date() });
    S.rows = p.rows; S.assessments = p.assessments;
    render();
  }

  // ------------------------------------------------------------------ rules
  function domainOf(address) { var p = E.addressParts(address); return p ? E.registrableDomain(p.domain) : null; }
  function canUseDomain(address) { var d = domainOf(address); return !!d && !PUBLIC_DOMAINS.test(d); }
  function senderName(address) { var s = S.senders[address]; if (s) return s.name; if (S.selected && S.selected.address === address) return S.selected.name || address; return address; }

  function usesDomain(address) {
    if (S.domainFlag[address] !== undefined) return S.domainFlag[address];
    var r = E.ruleFor(S.rules, address);
    return !!(r && r.scope === 'domain');
  }

  function snapshot() { S.ruleSnapshot = JSON.stringify(S.rules); }

  // The subject filter in the card's editor: what is typed, else the existing subject rule for the open email, else ''.
  // 'Re: Fwd: Demo booked -  ' -> 'Demo booked'
  function cleanSubject(s) { return String(s || '').replace(/^\s*((re|fw|fwd|aw|wg)\s*:\s*)+/i, '').replace(/[\s\-–—:|]+$/, '').trim(); }
  function existingSubjectRule(address) {
    var sel = S.selected && S.selected.address === address ? S.selected : null;
    return sel ? E.subjectRuleFor(S.rules, address, sel.subject) : null;
  }
  // Text in the subject box: what was typed, else the matching subject rule's text, else the open email's subject.
  function subjectText(address) {
    if (S.subjectText[address] !== undefined) return S.subjectText[address];
    var sr = existingSubjectRule(address);
    if (sr) return sr.has;
    return S.selected && S.selected.address === address ? cleanSubject(S.selected.subject) : '';
  }
  // Is the 'only when the subject has' box ticked? Defaults to ticked when the open email already matches a subject rule.
  function subjectOn(address) {
    if (S.subjectOn[address] !== undefined) return S.subjectOn[address];
    return !!existingSubjectRule(address);
  }
  function subjectFilter(address) { return subjectOn(address) ? subjectText(address) : ''; }
  function cardRule(address) {   // the rule the card is editing: a subject rule if a filter is set, else the sender/domain rule
    var has = subjectFilter(address).trim();
    if (has) { var m = (S.rules.subjects || []).filter(function (r) { return r.from === address && r.has.toLowerCase() === has.toLowerCase(); })[0]; var c = m ? E.splitCode(m.code) : null; return m ? { bucket: c.bucket, read: c.read, scope: 'subject', key: address, has: m.has } : null; }
    return E.ruleFor(S.rules, address);
  }

  function readFlag(address) {
    if (S.readFlag[address] !== undefined) return S.readFlag[address];
    var r = cardRule(address);
    return !!(r && r.read);
  }

  function setRule(address, code, quiet) {
    snapshot();
    code = E.joinCode(E.splitCode(code).bucket, readFlag(address));
    var has = subjectFilter(address).trim();
    if (has) {
      S.rules.subjects = (S.rules.subjects || []).filter(function (r) { return !(r.from === address && r.has.toLowerCase() === has.toLowerCase()); });
      S.rules.subjects.push({ from: address, has: has, code: code });
      var sb = E.splitCode(code).bucket;
      afterRuleChange(quiet ? null : '"' + has + '" from ' + senderName(address) + ': ' + (sb === 'I' ? 'keep in inbox' : sb === 'D' ? 'delete' : E.bucketName(sb)) + (E.splitCode(code).read ? ', mark as read' : ''));
      return;
    }
    var existing = E.ruleFor(S.rules, address);
    if (existing && existing.scope === 'domain') delete S.rules.domains[existing.key];
    delete S.rules.senders[address];
    var dom = domainOf(address);
    if (usesDomain(address) && canUseDomain(address)) S.rules.domains[dom] = code;
    else S.rules.senders[address] = code;
    var b = E.splitCode(code).bucket;
    var label = (b === 'I' ? 'keep in inbox' : b === 'D' ? 'delete' : E.bucketName(b)) + (E.splitCode(code).read ? ', mark as read' : '');
    afterRuleChange(quiet ? null : senderName(address) + ': ' + label);
  }

  function clearRule(address) {
    snapshot();
    var has = subjectFilter(address).trim();
    if (has) {
      S.rules.subjects = (S.rules.subjects || []).filter(function (r) { return !(r.from === address && r.has.toLowerCase() === has.toLowerCase()); });
      delete S.readFlag[address]; delete S.subjectText[address]; delete S.subjectOn[address];
      afterRuleChange('Subject rule removed for ' + senderName(address));
      return;
    }
    var existing = E.ruleFor(S.rules, address);
    if (existing && existing.scope === 'domain') delete S.rules.domains[existing.key];
    delete S.rules.senders[address];
    delete S.domainFlag[address]; delete S.readFlag[address];
    afterRuleChange('Rule removed for ' + senderName(address));
  }

  function afterRuleChange(message) {
    G.saveRules(S.rules).then(function (r) {
      if (r.size > 30000) toast('Your rule list is getting large for Outlook to store - tell Claude.', null);
    });
    // the open email got a rule but sits outside the dates shown: widen the range so Tidy can reach it straight away
    var sel = S.selected, when = sel && sel.received, widened = false;
    if (!S.focus && when && !isNaN(when)) {
      var r0 = rangeDates(S.range), cr = cardRule(sel.address);
      if (cr && cr.bucket !== 'I' && !(when >= r0[0] && when < r0[1])) {
        var want = E.rangeContaining(when);
        if (want) { var selEl = $('range'); selEl.value = want; S.range = want; S.messages = []; S.editing = null; refresh(); widened = true; }
      }
    }
    if (!widened) replan();
    if (message) toast(message, function () { S.rules = E.normaliseRules(JSON.parse(S.ruleSnapshot)); S.domainFlag = {}; S.readFlag = {}; S.subjectText = {}; S.subjectOn = {}; G.saveRules(S.rules); replan(); });
  }

  function suggestions() {
    return Object.keys(S.senders).filter(function (a) {
      return !E.ruleFor(S.rules, a) && S.assessments[a] && S.assessments[a].kind === 'suggest';
    }).map(function (a) { return S.senders[a]; }).sort(function (a, b) { return b.total - a.total || a.name.localeCompare(b.name); });
  }

  function approveAll() {
    var list = suggestions();
    if (!list.length) return;
    snapshot();
    list.forEach(function (s) { S.rules.senders[s.address] = S.assessments[s.address].bucket; });
    afterRuleChange('Approved ' + plural(list.length, 'rule'));
  }

  // ------------------------------------------------------------------ tidy + undo
  async function folderIdFor(code) {
    var b = E.BUCKETS[code];
    if (b && (b.folder === 'deleteditems' || b.folder === 'junkemail')) return b.folder;
    var name = E.bucketName(code);
    if (!S.folders.length) S.folders = await G.listFolders();
    var hit = S.folders.filter(function (f) { return f.name.toLowerCase() === name.toLowerCase(); })[0];
    if (hit) return hit.id;
    var made = await G.createFolder(name);
    S.folders.push(made);
    return made.id;
  }

  function setProgress(fraction) {
    $('progress').hidden = fraction == null;
    if (fraction != null) $('bar').style.width = Math.round(fraction * 100) + '%';
  }

  async function runMoves(items, label) {
    // items: [{ id, to }] ; returns ids that moved
    var moved = [], failed = 0, i = 0;
    setProgress(0);
    async function worker() {
      while (i < items.length) {
        var it = items[i++];
        try { await G.moveMessage(it.id, it.to); moved.push(it.id); } catch (e) { failed++; if (failed === 1) console.error(e); }
        setProgress((moved.length + failed) / items.length);
        $('summary').textContent = label + ' ' + (moved.length + failed) + ' of ' + items.length + '...';
      }
    }
    await Promise.all([worker(), worker(), worker()]);
    setProgress(null);
    return { moved: moved, failed: failed };
  }

  // items: [{ id, isRead }] ; returns ids updated
  async function runEach(items, label) {
    var done = [], i = 0;
    async function worker() {
      while (i < items.length) {
        var it = items[i++];
        try { await G.setRead(it.id, it.isRead); done.push(it.id); } catch (e) { /* leave as is */ }
        $('summary').textContent = label + ' ' + i + ' of ' + items.length + '...';
      }
    }
    await Promise.all([worker(), worker(), worker()]);
    return { done: done };
  }

  // Rows the open email's rule would file in the period shown (that sender, or that subject rule's matches).
  function rowsForSelected() {
    var sel = S.selected; if (!sel || S.focus) return [];
    var cr = cardRule(sel.address); if (!cr || cr.bucket === 'I') return [];
    return S.rows.filter(function (r) {
      if (r.dest === 'I' || r.msg.from !== sel.address) return false;
      return cr.scope === 'subject' ? (r.rule && r.rule.scope === 'subject' && r.rule.has === cr.has) : !(r.rule && r.rule.scope === 'subject');
    });
  }

  // 'File this one': just the open email, by its rule, whatever period the pane is showing.
  async function fileSelectedOne() {
    var sel = S.selected; if (!sel || !sel.itemId || S.busy) return;
    var cr = cardRule(sel.address); if (!cr || cr.bucket === 'I') return;
    S.busy = true; $('tidy').disabled = true; $('undo').disabled = true; $('summary').textContent = 'Reading that email...';
    try {
      var id = Office.context.mailbox.convertToRestId ? Office.context.mailbox.convertToRestId(sel.itemId, Office.MailboxEnums.RestVersion.v2_0) : sel.itemId;
      var m = await G.messageInfo(id);
      var row = E.plan([m], E.buildSenders([m]), S.rules, S.ctx, { now: new Date() }).rows[0];
      S.busy = false;
      if (!row || row.dest === 'I') { replan(); toast(row && row.why ? 'Kept in the inbox: ' + row.why : 'Nothing to file for this email', null); return; }
      await tidy([row]);
    } catch (err) { S.busy = false; setProgress(null); showError(err); }
    finally { S.busy = false; }
  }

  async function tidy(only) {
    var rows = Array.isArray(only) ? only : S.rows.filter(function (r) { return r.dest !== 'I'; });
    if (!rows.length || S.busy) return;
    S.busy = true; $('tidy').disabled = true; $('undo').disabled = true;
    try {
      var ids = {}, codes = rows.map(function (r) { return r.dest; }).filter(function (c, idx, a) { return a.indexOf(c) === idx; });
      for (var c = 0; c < codes.length; c++) ids[codes[c]] = await folderIdFor(codes[c]);
      var result = await runMoves(rows.map(function (r) { return { id: r.msg.id, to: ids[r.dest] }; }), 'Filing');
      var from = {}, wasUnread = [];
      rows.forEach(function (r) { if (r.msg.folder) from[r.msg.id] = r.msg.folder; });
      var movedSet = {}; result.moved.forEach(function (id) { movedSet[id] = true; });
      var toRead = rows.filter(function (r) { return r.markRead && movedSet[r.msg.id]; }).map(function (r) { return r.msg.id; });
      if (toRead.length) { var readRes = await runEach(toRead.map(function (id) { return { id: id, isRead: true }; }), 'Marking as read'); wasUnread = readRes.done; }
      G.local.set('is.undo.v1', { at: Date.now(), ids: result.moved, from: from, unread: wasUnread });
      var gone = {}; result.moved.forEach(function (id) { gone[id] = true; });
      S.messages = S.messages.filter(function (m) { return !gone[m.id]; });
      if (S.focus) S.focus = S.messages;
      S.senders = E.buildSenders(S.messages);
      S.busy = false;
      replan();
      toast('Filed ' + plural(result.moved.length, 'email') + (result.failed ? ' (' + result.failed + ' could not be moved)' : ''), undo, 'Undo');
    } catch (err) { S.busy = false; setProgress(null); showError(err); }
    finally { S.busy = false; }
  }

  async function undo() {
    var last = G.local.get('is.undo.v1');
    if (!last || !last.ids || !last.ids.length || S.busy) return;
    S.busy = true; $('tidy').disabled = true; $('undo').disabled = true;
    try {
      var result = await runMoves(last.ids.map(function (id) { return { id: id, to: (last.from && last.from[id]) || 'inbox' }; }), 'Putting back');
      if (last.unread && last.unread.length) await runEach(last.unread.map(function (id) { return { id: id, isRead: false }; }), 'Marking unread again');
      G.local.remove('is.undo.v1');
      S.busy = false;
      toast(plural(result.moved.length, 'email') + ' put back', null);
      await refresh();
    } catch (err) { S.busy = false; setProgress(null); showError(err); }
    finally { S.busy = false; }
  }

  // ------------------------------------------------------------------ selected email
  function readSelection() {
    try {
      S.cardMore = false; S.subjectText = {}; S.subjectOn = {}; S.readFlag = {};
      var item = Office.context.mailbox.item;
      if (item && item.from && item.from.emailAddress) {
        S.selected = { address: String(item.from.emailAddress).toLowerCase(), name: item.from.displayName || '', subject: item.subject || '', received: item.dateTimeCreated ? new Date(item.dateTimeCreated) : null, itemId: item.itemId || null };
      } else { S.selected = null; }
    } catch (e) { S.selected = null; }
  }
  // Several emails highlighted in Outlook's list: show their senders so rules can be set on just those.
  var pickCache = {}, pickRun = 0;
  function readHighlighted() {
    var mb = Office.context.mailbox, run = ++pickRun;
    if (!mb.getSelectedItemsAsync) return;
    try { mb.getSelectedItemsAsync(onItems); } catch (e) { console.warn('highlighted emails unavailable', e); }   // older manifest / Outlook: single selection still works
    function onItems(res) {
      if (run !== pickRun) return;
      var items = res && res.status === 'succeeded' && res.value ? res.value.filter(function (i) { return !i.itemType || String(i.itemType).toLowerCase() === 'message'; }) : [];
      if (items.length < 2) { if (S.focus) { S.focus = null; S.messages = []; S.editing = null; refresh(); } return; }
      loadHighlighted(items, run);
    }
  }
  async function loadHighlighted(items, run) {
    if (!S.focus) { S.focus = []; S.editing = null; showLoading('Looking at ' + items.length + ' highlighted emails...'); }
    var i = 0, found = [], seen = {};
    async function worker() {
      while (i < items.length) {
        var it = items[i++];
        try {
          if (!pickCache[it.itemId]) {
            var id = Office.context.mailbox.convertToRestId ? Office.context.mailbox.convertToRestId(it.itemId, Office.MailboxEnums.RestVersion.v2_0) : it.itemId;
            pickCache[it.itemId] = await G.messageInfo(id);
          }
          var m = pickCache[it.itemId];
          if (!seen[m.id]) { seen[m.id] = true; found.push(m); }
        } catch (e) { /* not a message we can read (other mailbox, draft...) - skip */ }
      }
    }
    await Promise.all([worker(), worker(), worker(), worker()]);
    if (run !== pickRun) return;
    var evidenceChanged = false;
    found.forEach(function (m) { if (E.addressParts(m.from) && !S.ctx.evidence[m.from]) { S.ctx.evidence[m.from] = E.readHeaders(m.headers); evidenceChanged = true; } });
    if (evidenceChanged) G.local.set('is.evidence.v1', S.ctx.evidence);
    found.sort(function (a, b) { return b.received - a.received; });
    S.focus = found; S.messages = found; S.senders = E.buildSenders(found);
    replan();
  }

  function watchSelection() {
    readSelection(); readHighlighted();
    try { Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, function () { readSelection(); render(); }); } catch (e) { /* not pinned-capable */ }
    try { if (Office.EventType.SelectedItemsChanged) Office.context.mailbox.addHandlerAsync(Office.EventType.SelectedItemsChanged, function () { readHighlighted(); }); } catch (e) { /* single selection only */ }
  }

  // ------------------------------------------------------------------ rendering
  function editorHtml(address, inCard) {
    var rule = inCard ? cardRule(address) : E.ruleFor(S.rules, address);
    var current = rule ? rule.bucket : null;
    var has = inCard ? subjectFilter(address) : '';
    var subjOn = inCard && subjectOn(address), subjTxt = inCard ? subjectText(address) : '';
    var chips = ['I', 'N', 'R', 'T', 'D', 'J'].map(function (code) {
      return '<button class="chip ' + code + '" data-act="set" data-code="' + code + '" aria-pressed="' + (current === code) + '">' + esc(E.BUCKETS[code].label) + '</button>';
    }).join('');
    var skip = { 'inbox': 1, 'deleted items': 1, 'junk email': 1, 'drafts': 1, 'sent items': 1, 'outbox': 1, 'conversation history': 1, 'newsletters': 1, 'receipts': 1, 'notifications': 1 };
    var folderCode = current && !E.BUCKETS[current] ? current : '';
    var options = S.folders.filter(function (f) { return !skip[f.name.toLowerCase()]; })
      .sort(function (a, b) { return a.name.localeCompare(b.name); })
      .map(function (f) { var v = 'F:' + f.name; return '<option value="' + esc(v) + '"' + (v === folderCode ? ' selected' : '') + '>' + esc(f.name) + '</option>'; }).join('');
    var dom = domainOf(address);
    var domainLine = canUseDomain(address) && !has.trim()
      ? '<label class="line"><input type="checkbox" data-act="domain"' + (usesDomain(address) ? ' checked' : '') + '> Everything from ' + esc(dom) + '</label>' : '';
    var subjectLine = inCard && S.selected ? '<label class="line subj"><input type="checkbox" data-act="subject-on"' + (subjOn ? ' checked' : '') + '> Only when the subject has</label>'
      + '<div class="line subj-text"><input type="text" data-act="subject" value="' + esc(subjTxt) + '"' + (subjOn ? '' : ' disabled') + ' title="Rule applies only to this sender\'s emails whose subject contains this text"></div>' : '';
    var readLine = '<label class="line"><input type="checkbox" data-act="read"' + (readFlag(address) ? ' checked' : '') + '> Mark as read when filed</label>';
    var extras = (options ? '<div class="line"><span>or folder</span><select data-act="folder"><option value="">Choose...</option>' + options + '</select></div>' : '') + subjectLine + domainLine + readLine;
    var showExtras = !inCard || S.cardMore || !!folderCode || usesDomain(address) || readFlag(address) || subjOn;
    return '<div class="editor" data-addr="' + esc(address) + '">'
      + (inCard ? (has.trim() ? '<p class="label">Mail from this sender with "' + esc(has.trim()) + '" in the subject goes to:</p>' : '') : '<p class="label">Mail from ' + esc(address) + ' always goes to:</p>')
      + '<div class="chips">' + chips + '</div>'
      + (showExtras ? extras : '')
      + '<div class="foot">' + (rule ? '<button class="link" data-act="clear">Remove rule</button>' : '<span></span>')
      + (inCard ? (showExtras ? '' : '<button class="link" data-act="card-more">Folder, domain, subject, mark as read...</button>') : '<button class="link" data-act="close">Done</button>') + '</div></div>';
  }

  function focusCard() {
    return '<div class="card"><p class="eyebrow">Highlighted emails</p><p class="state" style="margin:0">Showing only the ' + plural(S.focus.length, 'email') + ' you have highlighted. Tidy files just these. Click a single email to go back to ' + esc(rangeLabel().toLowerCase()) + '.</p></div>';
  }

  function selectedCard() {
    if (S.focus) return focusCard();
    if (!S.selected) return '';
    var a = S.selected.address, rule = cardRule(a), state;
    if (rule) state = (rule.scope === 'subject' ? 'Your rule: "' + rule.has + '" from this sender ' : 'Your rule: always ') + (rule.bucket === 'I' ? (rule.scope === 'subject' ? 'stays in the inbox' : 'keep in the inbox') : (rule.bucket === 'D' ? 'delete' : (rule.scope === 'subject' ? 'goes to ' : 'file under ') + E.bucketName(rule.bucket))) + (rule.scope === 'domain' ? ' (all of ' + rule.key + ')' : '') + (rule.read ? ', mark as read' : '');
    else if (subjectFilter(a).trim()) state = 'No rule yet for "' + subjectFilter(a).trim() + '" from this sender - pick where it goes';
    else {
      var as = S.assessments[a] || (E.addressParts(a) ? E.assessSender({ address: a, total: 1, transN: 0 }, S.ctx) : { kind: 'unknown' });
      if (as.kind === 'people') state = 'No rule - stays in the inbox (' + as.why + ')';
      else if (as.kind === 'suggest') state = 'No rule yet. Suggested: ' + E.bucketName(as.bucket) + ' (' + as.why + ')';
      else state = 'No rule yet - stays in the inbox';
    }
    return '<div class="card"><p class="eyebrow">Selected email</p><div class="who">' + esc(S.selected.name || a) + '</div><div class="addr">' + esc(a) + '</div>'
      + '<p class="state">' + esc(state) + '</p>' + rangeHint(rule, S.selected.received) + editorHtml(a, true) + fileTheseButton() + '</div>';
  }

  // 'File Ocado now · 4 emails': apply just the open email's rule, leaving everything else for Tidy.
  function fileTheseButton() {
    var sel = S.selected, cr = cardRule(sel.address);
    if (!cr || cr.bucket === 'I') return '';
    var rows = rowsForSelected(), dis = S.busy ? ' disabled' : '';
    var one = sel.itemId ? '<button class="btn file-these" data-act="file-one"' + dis + '>File this one</button>' : '';
    var who = cr.scope === 'subject' ? '"' + cr.has + '"' : (sel.name || sel.address);
    var all = rows.length ? '<button class="btn file-these" data-act="file-these"' + dis + '>File ' + esc(who) + ' now · ' + plural(rows.length, 'email') + '</button>' : '';
    return one || all ? '<div class="file-row">' + one + all + '</div>' : '';
  }

  // A rule was set on the open email but it is outside the dates shown, so Tidy cannot reach it: say so.
  function rangeHint(rule, when) {
    if (!rule || rule.bucket === 'I' || !when || isNaN(when)) return '';
    var r = rangeDates(S.range);
    if (when >= r[0] && when < r[1]) return '';
    var want = E.rangeContaining(when), labels = { today: 'Today', yesterday: 'Yesterday', '7': 'Last 7 days', '30': 'Last 30 days' };
    var day = (Date.now() - when) < 6 * 86400000 ? when.toLocaleDateString('en-GB', { weekday: 'long' }) : when.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });
    if (!want) return '<p class="hint">This email is from ' + esc(day) + '. The date list only reaches back 30 days; highlight the emails in Outlook to file older ones.</p>';
    return '<p class="hint">This email is from ' + esc(day) + ', outside "' + esc(rangeLabel()) + '", so Tidy will not pick it up here. '
      + '<button class="link" data-act="range" data-range="' + want + '">Show ' + esc(labels[want]) + '</button></p>';
  }

  function senderRow(s, opts) {
    opts = opts || {};
    var a = s.address, rule = E.ruleFor(S.rules, a), as = S.assessments[a] || {};
    var acts = '';
    if (opts.suggest) {
      acts += '<button class="round yes" data-act="approve" title="Approve: always ' + esc(E.bucketName(as.bucket)) + '" aria-label="Approve">' + ICON.check + '</button>'
        + '<button class="round no" data-act="reject" title="No - keep this sender in the inbox" aria-label="Reject">' + ICON.cross + '</button>';
    }
    acts += '<button class="round more" data-act="edit" title="Choose something else" aria-label="Choose a rule" aria-expanded="' + (S.editing === a) + '">' + ICON.more + '</button>';
    // suggestions: name, then folder + latest subject (the reason is the tooltip). Everything else: one line.
    var subject = esc(s.latest ? s.latest.subject : ''), tip = esc(a + (as.why ? ' - ' + as.why : ''));
    var who = '<div class="who">' + esc(s.name) + '<span class="n">' + s.total + '</span></div>';
    var tag = opts.suggest ? pill(as.bucket) : (rule && (rule.bucket !== 'I' || rule.scope === 'domain') ? pill(rule.bucket) : '');
    var body = opts.suggest
      ? who + '<div class="what">' + tag + ' ' + subject + '</div>'
      : '<div class="one">' + who + tag + '<div class="what">' + subject + '</div></div>';
    return '<div class="row' + (opts.suggest ? '' : ' slim') + '" data-addr="' + esc(a) + '" title="' + tip + '"><div class="text">' + body + '</div>'
      + '<div class="acts">' + acts + '</div></div>'
      + (S.editing === a ? editorHtml(a, false) : '');
  }

  function messageRow(r, action) {
    var m = r.msg, a = m.from;
    var link = action === 'keep' ? '<button class="link" data-act="keep" title="Leave this one email in the inbox">Keep</button>'
      : action === 'unkeep' ? '<button class="link" data-act="unkeep">File it</button>' : '';
    if (r.rule && r.rule.scope === 'subject' && r.why === 'your rule') r = Object.assign({}, r, { why: 'your rule for "' + r.rule.has + '"' });
    var note = r.group !== 'file' || r.why !== 'your rule' ? '<div class="why">' + esc(r.why) + (r.wouldBe ? ' · otherwise ' + esc(E.bucketName(r.wouldBe)) : '') + (r.markRead ? ' · will be marked read' : '') + '</div>' : (r.markRead ? '<div class="why">will be marked read</div>' : '');
    return '<div class="row msg slim" data-addr="' + esc(a) + '" data-id="' + esc(m.id) + '" title="' + esc(a) + '"><div class="text">'
      + '<div class="one"><div class="who">' + (m.isRead ? '' : '<span class="unread-dot"></span>') + esc(m.fromName || a) + '</div><div class="what" title="' + esc(m.subject) + '">' + esc(m.subject) + '</div></div>' + note + '</div>'
      + '<div class="acts">' + link + '<button class="round more" data-act="edit" aria-label="Change the rule" title="Change the rule for this sender" aria-expanded="' + (S.editing === a) + '">' + ICON.more + '</button></div></div>';
  }

  function section(key, title, count, extra, body) {
    var open = S.open[key];
    return '<section class="section"><button class="section-head" data-act="toggle" data-key="' + key + '" aria-expanded="' + open + '">' + ICON.chev
      + '<span>' + esc(title) + '</span><span class="count">' + count + '</span><span class="spacer"></span></button>'
      + (open ? (extra || '') + body : '') + '</section>';
  }

  function uniqueSenders(rows) {
    var seen = {}, out = [];
    rows.forEach(function (r) { if (!seen[r.msg.from]) { seen[r.msg.from] = true; out.push(S.senders[r.msg.from]); } });
    return out.sort(function (a, b) { return b.total - a.total || a.name.localeCompare(b.name); });
  }

  function render() {
    var main = $('main'), scroll = main.scrollTop, html = selectedCard();
    var sug = suggestions();
    var filing = S.rows.filter(function (r) { return r.dest !== 'I'; });
    var staying = S.rows.filter(function (r) { return r.dest === 'I'; });

    if (!S.messages.length) {
      html += '<div class="state-box"><h2>Nothing here</h2><p>' + (S.focus ? 'None of the highlighted emails could be read.' : 'No email in the inbox for ' + esc(rangeLabel().toLowerCase()) + '.') + '</p></div>';
    } else {
      // 1. suggestions
      if (sug.length) {
        html += section('suggest', 'Suggested rules', sug.length,
          '<p class="hint">Tick to approve, cross to keep that sender in the inbox. <button class="link" data-act="approve-all">Approve all</button></p>',
          sug.map(function (s) { return senderRow(s, { suggest: true }); }).join(''));
      }
      // 2. ready to file
      var byDest = {}, order = [];
      filing.forEach(function (r) { if (!byDest[r.dest]) { byDest[r.dest] = []; order.push(r.dest); } byDest[r.dest].push(r); });
      var rank = { R: 1, T: 2, N: 3, J: 8, D: 9 };
      order.sort(function (a, b) { return (rank[a] || 5) - (rank[b] || 5); });
      var fileBody = filing.length ? order.map(function (code) {
        var open = S.openDest[code] !== undefined ? S.openDest[code] : filing.length <= 15;
        return '<button class="section-head" style="padding-left:30px;font-weight:400" data-act="toggle-dest" data-code="' + esc(code) + '" aria-expanded="' + open + '">' + ICON.chev
          + pill(code) + '<span class="count">' + byDest[code].length + '</span></button>'
          + (open ? byDest[code].map(function (r) { return messageRow(r, 'keep') + (S.editing === r.msg.from && byDest[code].filter(function (x) { return x.msg.from === r.msg.from; })[0] === r ? editorHtml(r.msg.from, false) : ''); }).join('') : '');
      }).join('') : '<p class="hint">Nothing to file yet. Approve a suggestion above, or pick a rule for any sender.</p>';
      html += section('file', 'Ready to file', filing.length, '', fileBody);

      // 3. staying
      var groups = [
        ['kept', 'Looks like it needs you', 'msg'], ['keep1', 'You kept these', 'unkeep'],
        ['people', 'People', 'sender'], ['rule-inbox', 'Your rule: keep in inbox', 'sender'], ['unknown', 'No rule yet', 'sender']
      ];
      var stayBody = groups.map(function (g) {
        var rows = staying.filter(function (r) { return r.group === g[0]; });
        if (!rows.length) return '';
        var inner = g[2] === 'sender' ? uniqueSenders(rows).map(function (s) { return senderRow(s); }).join('')
          : rows.map(function (r) { return messageRow(r, g[2] === 'unkeep' ? 'unkeep' : '') + (S.editing === r.msg.from && rows.filter(function (x) { return x.msg.from === r.msg.from; })[0] === r ? editorHtml(r.msg.from, false) : ''); }).join('');
        return '<div class="sub">' + esc(g[1]) + ' · ' + rows.length + '</div>' + inner;
      }).join('');
      var stayCount = staying.filter(function (r) { return r.group !== 'suggest'; }).length;
      if (stayCount) html += section('stay', 'Staying in the inbox', stayCount, '', stayBody);
    }

    main.innerHTML = html;
    main.scrollTop = scroll;

    $('summary').textContent = scopeLabel() + ' · ' + plural(S.messages.length, 'email') + ' · ' + filing.length + ' to file' + (sug.length ? ' · ' + plural(sug.length, 'suggestion') : '');
    $('tidy').disabled = !filing.length || S.busy;
    $('tidy').textContent = filing.length ? 'Tidy now · file ' + plural(filing.length, 'email') : 'Tidy now';
    var last = G.local.get('is.undo.v1');
    $('undo').disabled = !(last && last.ids && last.ids.length) || S.busy;
  }

  function showLoading(text) { $('main').innerHTML = '<div class="state-box"><div class="spinner"></div><p>' + esc(text) + '</p></div>'; $('summary').textContent = ''; }

  function showError(err) {
    console.error(err);
    var setup = err && err.name === 'SetupError';
    $('main').innerHTML = '<div class="state-box"><h2>' + (setup ? 'One more setup step' : 'Something went wrong') + '</h2><p>' + esc(setup ? err.message : 'Nothing was changed. You can try again.') + '</p>'
      + (setup ? '' : '<button class="btn" data-act="retry">Try again</button><code>' + esc((err && err.message) || err) + '</code>') + '</div>';
    $('summary').textContent = '';
  }

  function toast(message, onUndo, label) {
    var t = $('toast');
    clearTimeout(S.toastTimer);
    t.innerHTML = '<span>' + esc(message) + '</span>' + (onUndo ? '<button id="toast-undo">' + esc(label || 'Undo') + '</button>' : '');
    t.hidden = false;
    if (onUndo) $('toast-undo').onclick = function () { t.hidden = true; onUndo(); };
    S.toastTimer = setTimeout(function () { t.hidden = true; }, onUndo ? 7000 : 3500);
  }

  // ------------------------------------------------------------------ events
  var wired = false;
  function wire() {
    if (wired) return;
    wired = true;
    $('range').addEventListener('change', function (e) { S.range = e.target.value; S.focus = null; S.messages = []; S.editing = null; refresh(); });
    $('refresh').addEventListener('click', function () { refresh(); });
    $('tidy').addEventListener('click', function () { tidy(); });
    $('undo').addEventListener('click', undo);

    $('main').addEventListener('click', function (e) {
      var el = e.target.closest('[data-act]');
      if (!el) return;
      var act = el.getAttribute('data-act');
      if (act === 'domain' || act === 'folder' || act === 'read' || act === 'subject' || act === 'subject-on') return;   // handled by 'change' / 'input'
      var holder = el.closest('[data-addr]'), address = holder && holder.getAttribute('data-addr');
      var idHolder = el.closest('[data-id]'), id = idHolder && idHolder.getAttribute('data-id');
      if (act === 'toggle') { var k = el.getAttribute('data-key'); S.open[k] = !S.open[k]; render(); }
      else if (act === 'toggle-dest') { var c = el.getAttribute('data-code'); var cur = S.openDest[c] !== undefined ? S.openDest[c] : S.rows.filter(function (r) { return r.dest !== 'I'; }).length <= 15; S.openDest[c] = !cur; render(); }
      else if (act === 'approve') { setRule(address, S.assessments[address].bucket); }
      else if (act === 'reject') { setRule(address, 'I'); }
      else if (act === 'approve-all') { approveAll(); }
      else if (act === 'edit') { S.editing = S.editing === address ? null : address; render(); }
      else if (act === 'close') { S.editing = null; render(); }
      else if (act === 'card-more') { S.cardMore = true; render(); }
      else if (act === 'set') { setRule(address, el.getAttribute('data-code')); }
      else if (act === 'clear') { clearRule(address); }
      else if (act === 'keep') { snapshot(); S.rules.keep[id] = Date.now(); afterRuleChange('Keeping that one in the inbox'); }
      else if (act === 'unkeep') { snapshot(); delete S.rules.keep[id]; afterRuleChange('It will be filed on the next tidy'); }
      else if (act === 'range') { var sel = $('range'); sel.value = el.getAttribute('data-range'); sel.dispatchEvent(new Event('change')); }
      else if (act === 'file-these') { tidy(rowsForSelected()); }
      else if (act === 'file-one') { fileSelectedOne(); }
      else if (act === 'retry') { boot().catch(showError); }
    });

    $('main').addEventListener('change', function (e) {
      var el = e.target, act = el.getAttribute && el.getAttribute('data-act');
      var holder = el.closest('[data-addr]'), address = holder && holder.getAttribute('data-addr');
      if (!address) return;
      if (act === 'folder' && el.value) { setRule(address, el.value); }
      else if (act === 'subject-on') {
        S.subjectOn[address] = el.checked; delete S.readFlag[address]; delete S.domainFlag[address];
        render();
        if (el.checked) { var box = document.querySelector('.card .editor [data-act="subject"]'); if (box) { box.focus(); box.select(); } }
      }
      else if (act === 'subject') {
        S.subjectText[address] = el.value; delete S.readFlag[address];
        render();
      }
      else if (act === 'read') {
        S.readFlag[address] = el.checked;
        var rr = cardRule(address);
        if (rr) setRule(address, rr.bucket); else render();
      }
      else if (act === 'domain') {
        S.domainFlag[address] = el.checked;
        var rule = E.ruleFor(S.rules, address);
        if (rule) setRule(address, rule.bucket); else render();
      }
    });
    // live typing in the subject box only updates state; the rule is written when a destination is chosen
    $('main').addEventListener('input', function (e) {
      var el = e.target; if (!el.getAttribute || el.getAttribute('data-act') !== 'subject') return;
      var holder = el.closest('[data-addr]'); if (holder) S.subjectText[holder.getAttribute('data-addr')] = el.value;
    });
  }

  function start() { boot().catch(showError); }
  if (window.Office && Office.onReady) Office.onReady(start); else document.addEventListener('DOMContentLoaded', start);
})();
