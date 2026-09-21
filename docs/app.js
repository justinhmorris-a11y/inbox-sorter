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
    folders: [], selected: null, picked: null, pickedNote: '', pickAssess: {}, editing: null, domainFlag: {}, cardMore: false,
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

  async function refresh() {
    if (S.busy) return;
    S.busy = true; $('refresh').classList.add('spin');
    try {
      var r = rangeDates(S.range);
      if (!S.messages.length) showLoading('Reading ' + rangeLabel().toLowerCase() + '...');
      S.messages = await G.listInbox(r[0], r[1], function (n) { $('summary').textContent = 'Reading... ' + n + ' emails'; });
      S.senders = E.buildSenders(S.messages);
      await gatherEvidence();
      S.busy = false;
      replan();
    } catch (err) { showError(err); }
    finally { S.busy = false; $('refresh').classList.remove('spin'); }
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
  function senderName(address) { var s = S.senders[address]; if (s) return s.name; if (S.selected && S.selected.address === address) return S.selected.name || address; var pk = (S.picked || []).filter(function (x) { return x.address === address; })[0]; if (pk) return pk.name; return address; }

  function usesDomain(address) {
    if (S.domainFlag[address] !== undefined) return S.domainFlag[address];
    var r = E.ruleFor(S.rules, address);
    return !!(r && r.scope === 'domain');
  }

  function snapshot() { S.ruleSnapshot = JSON.stringify(S.rules); }

  function setRule(address, code, quiet) {
    snapshot();
    var existing = E.ruleFor(S.rules, address);
    if (existing && existing.scope === 'domain') delete S.rules.domains[existing.key];
    delete S.rules.senders[address];
    var dom = domainOf(address);
    if (usesDomain(address) && canUseDomain(address)) S.rules.domains[dom] = code;
    else S.rules.senders[address] = code;
    var label = code === 'I' ? 'keep in inbox' : code === 'D' ? 'delete' : E.bucketName(code);
    afterRuleChange(quiet ? null : senderName(address) + ': ' + label);
  }

  function clearRule(address) {
    snapshot();
    var existing = E.ruleFor(S.rules, address);
    if (existing && existing.scope === 'domain') delete S.rules.domains[existing.key];
    delete S.rules.senders[address];
    delete S.domainFlag[address];
    afterRuleChange('Rule removed for ' + senderName(address));
  }

  function afterRuleChange(message) {
    G.saveRules(S.rules).then(function (r) {
      if (r.size > 30000) toast('Your rule list is getting large for Outlook to store - tell Claude.', null);
    });
    replan();
    if (message) toast(message, function () { S.rules = E.normaliseRules(JSON.parse(S.ruleSnapshot)); S.domainFlag = {}; G.saveRules(S.rules); replan(); });
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

  async function tidy() {
    var rows = S.rows.filter(function (r) { return r.dest !== 'I'; });
    if (!rows.length || S.busy) return;
    S.busy = true; $('tidy').disabled = true; $('undo').disabled = true;
    try {
      var ids = {}, codes = rows.map(function (r) { return r.dest; }).filter(function (c, idx, a) { return a.indexOf(c) === idx; });
      for (var c = 0; c < codes.length; c++) ids[codes[c]] = await folderIdFor(codes[c]);
      var result = await runMoves(rows.map(function (r) { return { id: r.msg.id, to: ids[r.dest] }; }), 'Filing');
      G.local.set('is.undo.v1', { at: Date.now(), ids: result.moved });
      var gone = {}; result.moved.forEach(function (id) { gone[id] = true; });
      S.messages = S.messages.filter(function (m) { return !gone[m.id]; });
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
      var result = await runMoves(last.ids.map(function (id) { return { id: id, to: 'inbox' }; }), 'Putting back');
      G.local.remove('is.undo.v1');
      S.busy = false;
      toast(plural(result.moved.length, 'email') + ' back in the inbox', null);
      await refresh();
    } catch (err) { S.busy = false; setProgress(null); showError(err); }
    finally { S.busy = false; }
  }

  // ------------------------------------------------------------------ selected email
  function readSelection() {
    try {
      S.cardMore = false;
      var item = Office.context.mailbox.item;
      if (item && item.from && item.from.emailAddress) {
        S.selected = { address: String(item.from.emailAddress).toLowerCase(), name: item.from.displayName || '', subject: item.subject || '', received: item.dateTimeCreated ? new Date(item.dateTimeCreated) : null };
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
      var items = res && res.status === 'succeeded' && res.value ? res.value.filter(function (i) { return !i.itemType || i.itemType === 'message'; }) : [];
      if (items.length < 2) { if (S.picked || S.pickedNote) { S.picked = null; S.pickedNote = ''; render(); } return; }
      loadHighlighted(items, run);
    }
  }
  async function loadHighlighted(items, run) {
    S.pickedNote = 'Looking at ' + items.length + ' highlighted emails...'; S.picked = null; render();
    var i = 0, found = [];
    async function worker() {
      while (i < items.length) {
        var it = items[i++];
        try {
          if (!pickCache[it.itemId]) {
            var id = Office.context.mailbox.convertToRestId ? Office.context.mailbox.convertToRestId(it.itemId, Office.MailboxEnums.RestVersion.v2_0) : it.itemId;
            pickCache[it.itemId] = await G.messageInfo(id);
          }
          found.push(pickCache[it.itemId]);
        } catch (e) { /* not a message we can read (other mailbox, draft...) - skip */ }
      }
    }
    await Promise.all([worker(), worker(), worker(), worker()]);
    if (run !== pickRun) return;
    var by = {}, evidenceChanged = false;
    found.forEach(function (m) {
      if (!E.addressParts(m.from)) return;
      var s = by[m.from] || (by[m.from] = { address: m.from, name: m.fromName || m.from, total: 0, transN: 0, latest: m, oldest: m.received });
      s.total++; if (m.received > s.latest.received) s.latest = m; if (m.received < s.oldest) s.oldest = m.received;
      if (!S.ctx.evidence[m.from]) { S.ctx.evidence[m.from] = E.readHeaders(m.headers); evidenceChanged = true; }
    });
    if (evidenceChanged) G.local.set('is.evidence.v1', S.ctx.evidence);
    S.picked = Object.keys(by).map(function (k) { return by[k]; }).sort(function (x, y) { return y.total - x.total || x.name.localeCompare(y.name); });
    S.pickedNote = found.length + ' highlighted email' + (found.length === 1 ? '' : 's') + ' · ' + S.picked.length + ' sender' + (S.picked.length === 1 ? '' : 's');
    render();
  }

  function watchSelection() {
    readSelection(); readHighlighted();
    try { Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, function () { readSelection(); render(); }); } catch (e) { /* not pinned-capable */ }
    try { if (Office.EventType.SelectedItemsChanged) Office.context.mailbox.addHandlerAsync(Office.EventType.SelectedItemsChanged, function () { readHighlighted(); }); } catch (e) { /* single selection only */ }
  }

  // ------------------------------------------------------------------ rendering
  function editorHtml(address, inCard) {
    var rule = E.ruleFor(S.rules, address);
    var current = rule ? rule.bucket : null;
    var chips = ['I', 'N', 'R', 'T', 'D', 'J'].map(function (code) {
      return '<button class="chip ' + code + '" data-act="set" data-code="' + code + '" aria-pressed="' + (current === code) + '">' + esc(E.BUCKETS[code].label) + '</button>';
    }).join('');
    var skip = { 'inbox': 1, 'deleted items': 1, 'junk email': 1, 'drafts': 1, 'sent items': 1, 'outbox': 1, 'conversation history': 1, 'newsletters': 1, 'receipts': 1, 'notifications': 1 };
    var folderCode = current && !E.BUCKETS[current] ? current : '';
    var options = S.folders.filter(function (f) { return !skip[f.name.toLowerCase()]; })
      .sort(function (a, b) { return a.name.localeCompare(b.name); })
      .map(function (f) { var v = 'F:' + f.name; return '<option value="' + esc(v) + '"' + (v === folderCode ? ' selected' : '') + '>' + esc(f.name) + '</option>'; }).join('');
    var dom = domainOf(address);
    var domainLine = canUseDomain(address)
      ? '<label class="line"><input type="checkbox" data-act="domain"' + (usesDomain(address) ? ' checked' : '') + '> Everything from ' + esc(dom) + '</label>' : '';
    var extras = (options ? '<div class="line"><span>or folder</span><select data-act="folder"><option value="">Choose...</option>' + options + '</select></div>' : '') + domainLine;
    var showExtras = !inCard || S.cardMore || !!folderCode || usesDomain(address);
    return '<div class="editor" data-addr="' + esc(address) + '">'
      + (inCard ? '' : '<p class="label">Mail from ' + esc(address) + ' always goes to:</p>')
      + '<div class="chips">' + chips + '</div>'
      + (showExtras ? extras : '')
      + '<div class="foot">' + (rule ? '<button class="link" data-act="clear">Remove rule</button>' : '<span></span>')
      + (inCard ? (showExtras ? '' : '<button class="link" data-act="card-more">Folder or whole domain...</button>') : '<button class="link" data-act="close">Done</button>') + '</div></div>';
  }

  function pickedCard() {
    S.pickAssess = {};
    if (!S.picked) return '<div class="card"><p class="eyebrow">Highlighted emails</p><p class="state">' + esc(S.pickedNote) + '</p></div>';
    var oldest = null, oldestRule = null;
    var rows = S.picked.map(function (s) {
      var rule = E.ruleFor(S.rules, s.address), as = S.assessments[s.address] || E.assessSender(s, S.ctx);
      S.pickAssess[s.address] = as;
      if (rule && rule.bucket !== 'I' && (!oldest || s.oldest < oldest)) { oldest = s.oldest; oldestRule = rule; }
      return senderRow(s, { suggest: !rule && as.kind === 'suggest', as: as });
    }).join('');
    return '<div class="card picked"><p class="eyebrow">' + esc(S.pickedNote) + '</p>' + rows + rangeHint(oldestRule, oldest, true) + '</div>';
  }

  function selectedCard() {
    if (S.picked || S.pickedNote) return pickedCard();
    if (!S.selected) return '';
    var a = S.selected.address, rule = E.ruleFor(S.rules, a), state;
    if (rule) state = (rule.bucket === 'I' ? 'Your rule: always keep in the inbox' : 'Your rule: always ' + (rule.bucket === 'D' ? 'delete' : 'file under ' + E.bucketName(rule.bucket))) + (rule.scope === 'domain' ? ' (all of ' + rule.key + ')' : '');
    else {
      var as = S.assessments[a] || (E.addressParts(a) ? E.assessSender({ address: a, total: 1, transN: 0 }, S.ctx) : { kind: 'unknown' });
      if (as.kind === 'people') state = 'No rule - stays in the inbox (' + as.why + ')';
      else if (as.kind === 'suggest') state = 'No rule yet. Suggested: ' + E.bucketName(as.bucket) + ' (' + as.why + ')';
      else state = 'No rule yet - stays in the inbox';
    }
    return '<div class="card"><p class="eyebrow">Selected email</p><div class="who">' + esc(S.selected.name || a) + '</div><div class="addr">' + esc(a) + '</div>'
      + '<p class="state">' + esc(state) + '</p>' + rangeHint(rule, S.selected.received) + editorHtml(a, true) + '</div>';
  }

  // A rule was set on the open email but it is outside the dates shown, so Tidy cannot reach it: say so.
  function rangeHint(rule, when, many) {
    if (!rule || rule.bucket === 'I' || !when || isNaN(when)) return '';
    var r = rangeDates(S.range);
    if (when >= r[0] && when < r[1]) return '';
    var want = E.rangeContaining(when), labels = { today: 'Today', yesterday: 'Yesterday', '7': 'Last 7 days', '30': 'Last 30 days' };
    var day = (Date.now() - when) < 6 * 86400000 ? when.toLocaleDateString('en-GB', { weekday: 'long' }) : when.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });
    if (!want) return '<p class="hint">' + (many ? 'The oldest of these is from ' : 'This email is from ') + esc(day) + '. Tidy only reaches back 30 days for now, so it stays where it is.</p>';
    return '<p class="hint">' + (many ? 'The oldest of these is from ' : 'This email is from ') + esc(day) + ', outside "' + esc(rangeLabel()) + '", so Tidy will not pick it up here. '
      + '<button class="link" data-act="range" data-range="' + want + '">Show ' + esc(labels[want]) + '</button></p>';
  }

  function senderRow(s, opts) {
    opts = opts || {};
    var a = s.address, rule = E.ruleFor(S.rules, a), as = opts.as || S.assessments[a] || {};
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
    var note = r.group !== 'file' || r.why !== 'your rule' ? '<div class="why">' + esc(r.why) + (r.wouldBe ? ' · otherwise ' + esc(E.bucketName(r.wouldBe)) : '') + '</div>' : '';
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
      html += '<div class="state-box"><h2>Nothing here</h2><p>No email in the inbox for ' + esc(rangeLabel().toLowerCase()) + '.</p></div>';
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

    $('summary').textContent = rangeLabel() + ' · ' + plural(S.messages.length, 'email') + ' · ' + filing.length + ' to file' + (sug.length ? ' · ' + plural(sug.length, 'suggestion') : '');
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
    $('range').addEventListener('change', function (e) { S.range = e.target.value; S.messages = []; S.editing = null; refresh(); });
    $('refresh').addEventListener('click', function () { refresh(); });
    $('tidy').addEventListener('click', tidy);
    $('undo').addEventListener('click', undo);

    $('main').addEventListener('click', function (e) {
      var el = e.target.closest('[data-act]');
      if (!el) return;
      var act = el.getAttribute('data-act');
      if (act === 'domain' || act === 'folder') return;              // handled by 'change'
      var holder = el.closest('[data-addr]'), address = holder && holder.getAttribute('data-addr');
      var idHolder = el.closest('[data-id]'), id = idHolder && idHolder.getAttribute('data-id');
      if (act === 'toggle') { var k = el.getAttribute('data-key'); S.open[k] = !S.open[k]; render(); }
      else if (act === 'toggle-dest') { var c = el.getAttribute('data-code'); var cur = S.openDest[c] !== undefined ? S.openDest[c] : S.rows.filter(function (r) { return r.dest !== 'I'; }).length <= 15; S.openDest[c] = !cur; render(); }
      else if (act === 'approve') { setRule(address, (S.assessments[address] || S.pickAssess[address]).bucket); }
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
      else if (act === 'retry') { boot().catch(showError); }
    });

    $('main').addEventListener('change', function (e) {
      var el = e.target, act = el.getAttribute && el.getAttribute('data-act');
      var holder = el.closest('[data-addr]'), address = holder && holder.getAttribute('data-addr');
      if (!address) return;
      if (act === 'folder' && el.value) { setRule(address, el.value); }
      else if (act === 'domain') {
        S.domainFlag[address] = el.checked;
        var rule = E.ruleFor(S.rules, address);
        if (rule) setRule(address, rule.bucket); else render();
      }
    });
  }

  function start() { boot().catch(showError); }
  if (window.Office && Office.onReady) Office.onReady(start); else document.addEventListener('DOMContentLoaded', start);
})();
