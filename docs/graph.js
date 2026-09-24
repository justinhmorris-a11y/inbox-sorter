/* Inbox Sorter - sign-in, Microsoft Graph calls and storage.
 * Sign-in uses nested app authentication: Outlook itself brokers the token, so there is no
 * separate login window in normal use.
 */
(function (global) {
  'use strict';

  var GRAPH = 'https://graph.microsoft.com/v1.0';
  var SCOPES = ['Mail.ReadWrite', 'MailboxSettings.ReadWrite', 'User.Read'];
  var pca = null;
  var cached = { token: null, expires: 0 };

  function SetupError(message) { this.name = 'SetupError'; this.message = message; }
  SetupError.prototype = Object.create(Error.prototype);

  function GraphError(status, body) {
    this.name = 'GraphError'; this.status = status;
    var text = body;
    try { var j = JSON.parse(body); if (j && j.error) text = j.error.code + ': ' + j.error.message; } catch (e) { /* keep raw */ }
    this.message = 'Microsoft Graph said ' + status + ' - ' + String(text).slice(0, 300);
  }
  GraphError.prototype = Object.create(Error.prototype);

  function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

  async function initAuth() {
    if (global.SORTER_MOCK) return;
    var cfg = global.SORTER_CONFIG || {};
    if (!cfg.clientId || /REPLACE/i.test(cfg.clientId)) {
      throw new SetupError('Setup is not finished: config.js does not have the app registration (client ID) yet.');
    }
    var ok = false;
    try { ok = Office.context.requirements.isSetSupported('NestedAppAuth', '1.1'); } catch (e) { ok = false; }
    if (!ok) throw new SetupError('This Outlook build does not support the sign-in method the add-in uses (nested app authentication). Update Microsoft 365 Apps and reopen Outlook.');
    pca = await msal.createNestablePublicClientApplication({
      auth: { clientId: cfg.clientId, authority: 'https://login.microsoftonline.com/' + (cfg.tenantId || 'common') },
      cache: { cacheLocation: 'localStorage' }
    });
  }

  async function getToken() {
    if (global.SORTER_MOCK) return 'mock';
    if (cached.token && Date.now() < cached.expires - 120000) return cached.token;
    var request = { scopes: SCOPES };
    var result;
    try { result = await pca.acquireTokenSilent(request); }
    catch (silentError) { result = await pca.acquireTokenPopup(request); }
    cached.token = result.accessToken;
    cached.expires = result.expiresOn ? new Date(result.expiresOn).getTime() : Date.now() + 1800000;
    return cached.token;
  }

  async function call(path, options) {
    options = options || {};
    var url = path.indexOf('http') === 0 ? path : GRAPH + path;
    var method = options.method || 'GET';
    if (global.SORTER_MOCK) return global.MockGraph.handle(method, url, options.body);
    for (var attempt = 1; ; attempt++) {
      var token = await getToken();
      var headers = { Authorization: 'Bearer ' + token };
      if (options.immutable !== false) headers.Prefer = 'IdType="ImmutableId"' + (options.bodyText ? ', outlook.body-content-type="text"' : '');
      if (options.body) headers['Content-Type'] = 'application/json';
      if (options.consistency) headers.ConsistencyLevel = 'eventual';
      var res;
      try { res = await fetch(url, { method: method, headers: headers, body: options.body ? JSON.stringify(options.body) : undefined }); }
      catch (networkError) {
        if (attempt >= 5) throw new Error('Could not reach Microsoft 365 - check your connection.');
        await sleep(1000 * Math.pow(2, attempt)); continue;
      }
      if (res.status === 401 && attempt < 3) { cached.token = null; continue; }
      if (res.status === 429 || res.status >= 500) {
        if (attempt >= 6) throw new GraphError(res.status, await res.text());
        var retryAfter = parseInt(res.headers.get('Retry-After') || '0', 10);
        await sleep(Math.max(retryAfter * 1000, 1000 * Math.pow(2, attempt)));
        continue;
      }
      if (!res.ok) throw new GraphError(res.status, await res.text());
      if (res.status === 204) return null;
      return res.json();
    }
  }

  async function me() {
    var u = await call('/me?$select=displayName,mail,userPrincipalName,proxyAddresses');
    var addresses = {};
    [u.mail, u.userPrincipalName].forEach(function (a) { if (a) addresses[String(a).toLowerCase()] = true; });
    (u.proxyAddresses || []).forEach(function (p) { var m = /^smtp:(.+)$/i.exec(p || ''); if (m) addresses[m[1].toLowerCase()] = true; });
    var domains = {};
    Object.keys(addresses).forEach(function (a) { var d = a.slice(a.indexOf('@') + 1); if (!/onmicrosoft\.com$/.test(d)) domains[d] = true; });
    return { name: u.displayName, addresses: addresses, domains: domains };
  }

  function mapMessage(m) {
    var ea = (m.from && m.from.emailAddress) || (m.sender && m.sender.emailAddress) || {};
    return {
      id: m.id,
      from: String(ea.address || '(no sender)').trim().toLowerCase(),
      fromName: ea.name || '',
      subject: m.subject || '(no subject)',
      received: new Date(m.receivedDateTime),
      isRead: !!m.isRead,
      flagged: !!(m.flag && m.flag.flagStatus === 'flagged'),
      importance: m.importance || 'normal',
      folder: m.parentFolderId || null
    };
  }

  /** Inbox messages received in [from, to), newest first. */
  async function listInbox(from, to, onProgress) {
    var filter = 'receivedDateTime ge ' + from.toISOString().slice(0, 19) + 'Z and receivedDateTime lt ' + to.toISOString().slice(0, 19) + 'Z';
    var url = '/me/mailFolders/inbox/messages?$select=id,from,sender,subject,receivedDateTime,isRead,flag,importance'
      + '&$filter=' + encodeURIComponent(filter) + '&$orderby=' + encodeURIComponent('receivedDateTime desc') + '&$top=200';
    var out = [], seen = {};
    while (url) {
      var page = await call(url);
      (page.value || []).forEach(function (m) { if (!seen[m.id]) { seen[m.id] = true; out.push(mapMessage(m)); } });
      if (onProgress) onProgress(out.length);
      url = page['@odata.nextLink'] || null;
    }
    return out;
  }

  var INBOX_SELECT = '$select=id,from,sender,subject,receivedDateTime,isRead,flag,importance,parentFolderId';
  /** How many inbox emails come from this address. */
  async function inboxCountFrom(address) {
    var url = '/me/mailFolders/inbox/messages?$filter=' + encodeURIComponent("from/emailAddress/address eq '" + String(address).replace(/'/g, "''") + "'") + '&$count=true&$top=1&$select=id';
    var page = await call(url, { consistency: true });
    return typeof page['@odata.count'] === 'number' ? page['@odata.count'] : (page.value || []).length;
  }
  /** Every inbox email from this address. */
  async function listInboxFrom(address, onProgress) {
    var url = '/me/mailFolders/inbox/messages?$filter=' + encodeURIComponent("from/emailAddress/address eq '" + String(address).replace(/'/g, "''") + "'") + '&' + INBOX_SELECT + '&$top=500';
    return pageAll(url, onProgress);
  }
  /** The whole inbox, newest first. */
  async function listInboxAll(onProgress) {
    return pageAll('/me/mailFolders/inbox/messages?' + INBOX_SELECT + '&$orderby=' + encodeURIComponent('receivedDateTime desc') + '&$top=1000', onProgress);
  }
  async function pageAll(url, onProgress) {
    var out = [], seen = {};
    while (url) {
      var page = await call(url);
      (page.value || []).forEach(function (m) { if (!seen[m.id]) { seen[m.id] = true; out.push(mapMessage(m)); } });
      if (onProgress) onProgress(out.length);
      url = page['@odata.nextLink'] || null;
    }
    return out;
  }

  /** Who the user writes to: { address: count } from the most recent sent mail. */
  async function sentRecipients(max, onProgress) {
    var url = '/me/mailFolders/sentitems/messages?$select=toRecipients,ccRecipients&$orderby=' + encodeURIComponent('receivedDateTime desc') + '&$top=500';
    var to = {}, scanned = 0;
    while (url && scanned < max) {
      var page = await call(url);
      (page.value || []).forEach(function (m) {
        scanned++;
        (m.toRecipients || []).concat(m.ccRecipients || []).forEach(function (r) {
          var a = r && r.emailAddress && r.emailAddress.address;
          if (a) { a = String(a).trim().toLowerCase(); to[a] = (to[a] || 0) + 1; }
        });
      });
      if (onProgress) onProgress(scanned);
      url = page['@odata.nextLink'] || null;
    }
    return { scanned: scanned, to: to };
  }

  /** Every inbox message counted by sender: { address: { address, name, total, unread, last90, transN, latest: { id, subject, received } } }.
   *  isTrans(subject) says whether a subject looks like an order / receipt. Reads the whole inbox, 1,000 at a time. */
  async function senderCounts(isTrans, onProgress) {
    var url = '/me/mailFolders/inbox/messages?$select=id,from,receivedDateTime,isRead,subject&$top=1000';
    var by = {}, n = 0, d90 = Date.now() - 90 * 86400000;
    while (url) {
      var page = await call(url);
      (page.value || []).forEach(function (raw) {
        var m = mapMessage(raw);
        var s = by[m.from] || (by[m.from] = { address: m.from, name: m.fromName || m.from, total: 0, unread: 0, last90: 0, transN: 0, latest: null });
        s.total++; if (!m.isRead) s.unread++; if (m.received > d90) s.last90++; if (isTrans && isTrans(m.subject)) s.transN++;
        if (!s.latest || m.received > s.latest.received) s.latest = { id: m.id, subject: m.subject, received: m.received };
        if (!s.name && m.fromName) s.name = m.fromName;
      });
      n += (page.value || []).length;
      if (onProgress) onProgress(n);
      url = page['@odata.nextLink'] || null;
    }
    return { scanned: n, senders: by };
  }

  async function messageHeaders(id) {
    var m = await call('/me/messages/' + encodeURIComponent(id) + '?$select=internetMessageHeaders');
    return m.internetMessageHeaders || [];
  }

  /** One message by id (any id format): who sent it, when, and its header evidence. */
  async function messageInfo(id) {
    var m = await call('/me/messages/' + encodeURIComponent(id) + '?$select=id,from,sender,subject,receivedDateTime,isRead,flag,importance,parentFolderId,internetMessageHeaders');
    var out = mapMessage(m); out.headers = m.internetMessageHeaders || [];
    return out;
  }

  async function setRead(id, isRead) {
    return call('/me/messages/' + encodeURIComponent(id), { method: 'PATCH', body: { isRead: !!isRead } });
  }

  async function moveMessage(id, destinationId, immutable) {
    return call('/me/messages/' + encodeURIComponent(id) + '/move', { method: 'POST', body: { destinationId: destinationId }, immutable: immutable });
  }

  /** Top-level folders plus the Inbox's own subfolders: [{ id, name }] */
  async function listFolders() {
    var out = [];
    var top = await call('/me/mailFolders?$top=200&$select=id,displayName');
    (top.value || []).forEach(function (f) { out.push({ id: f.id, name: f.displayName }); });
    try {
      var sub = await call('/me/mailFolders/inbox/childFolders?$top=200&$select=id,displayName');
      (sub.value || []).forEach(function (f) { out.push({ id: f.id, name: f.displayName }); });
    } catch (e) { /* not fatal */ }
    return out;
  }

  /** The real id of a well-known folder ('junkemail', 'deleteditems'): server rules need ids, not names. */
  async function wellKnownFolderId(name) {
    var f = await call('/me/mailFolders/' + name + '?$select=id');
    return f.id;
  }

  // ---- Outlook server rules ("apply at arrival") -----------------------------------------------
  async function listRules() {
    var out = [], url = '/me/mailFolders/inbox/messageRules?$top=100';
    while (url) { var page = await call(url); out = out.concat(page.value || []); url = page['@odata.nextLink'] || null; }
    return out;
  }
  function createRule(rule) { return call('/me/mailFolders/inbox/messageRules', { method: 'POST', body: rule }); }
  function deleteRule(id) { return call('/me/mailFolders/inbox/messageRules/' + encodeURIComponent(id), { method: 'DELETE' }); }

  async function createFolder(name) {
    var f = await call('/me/mailFolders', { method: 'POST', body: { displayName: name } });
    return { id: f.id, name: f.displayName };
  }

  // ---- storage -------------------------------------------------------------------------
  // Rules live in the mailbox (Outlook roaming settings) so they survive a cleared cache.
  // Everything else is a convenience cache in localStorage and is rebuilt if missing.
  var local = {
    get: function (key) { try { var v = global.localStorage.getItem(key); return v ? JSON.parse(v) : null; } catch (e) { return null; } },
    set: function (key, value) { try { global.localStorage.setItem(key, JSON.stringify(value)); } catch (e) { /* cache only */ } },
    remove: function (key) { try { global.localStorage.removeItem(key); } catch (e) { /* cache only */ } }
  };

  // ---- the rule store ---------------------------------------------------------------------
  // Rules live in one draft message inside a hidden folder of the mailbox: room for tens of thousands
  // of rules, roams with the mailbox, never shown in Outlook. Outlook's 32 KB roaming setting (the old
  // store) and a localStorage backup remain as fallbacks and are migrated on first load.
  var STORE_FOLDER = 'Inbox Sorter (data)', STORE_SUBJECT = 'Inbox Sorter rules - do not delete';
  var store = { folderId: null, messageId: null };

  async function storeFolderId() {
    if (store.folderId) return store.folderId;
    var q = await call('/me/mailFolders?includeHiddenFolders=true&$filter=' + encodeURIComponent("displayName eq '" + STORE_FOLDER + "'") + '&$select=id');
    var f = (q.value || [])[0];
    if (!f) f = await call('/me/mailFolders', { method: 'POST', body: { displayName: STORE_FOLDER, isHidden: true } });
    store.folderId = f.id;
    return f.id;
  }

  async function storeMessage() {
    var fid = await storeFolderId();
    var q = await call('/me/mailFolders/' + encodeURIComponent(fid) + '/messages?$filter=' + encodeURIComponent("subject eq '" + STORE_SUBJECT + "'") + '&$select=id,body&$top=2', { bodyText: true });
    var m = (q.value || [])[0];
    if (m) store.messageId = m.id;
    return m || null;
  }

  async function loadRulesRemote() {
    var m = await storeMessage();
    if (!m || !m.body || !m.body.content) return null;
    var text = String(m.body.content).replace(/<[^>]+>/g, '').trim();   // text body; strip tags if Outlook wrapped it
    var start = text.indexOf('{');
    return start < 0 ? null : JSON.parse(text.slice(start, text.lastIndexOf('}') + 1));
  }

  async function saveRulesRemote(text) {
    if (!store.messageId) await storeMessage();
    var body = { contentType: 'text', content: text };
    if (store.messageId) { await call('/me/messages/' + encodeURIComponent(store.messageId), { method: 'PATCH', body: { body: body } }); return; }
    var fid = await storeFolderId();
    var m = await call('/me/mailFolders/' + encodeURIComponent(fid) + '/messages', { method: 'POST', body: { subject: STORE_SUBJECT, body: body, isRead: true } });
    store.messageId = m.id;
  }

  function loadRulesLegacy() {
    var raw = null;
    try { raw = Office.context.roamingSettings.get('rules'); } catch (e) { raw = null; }
    if (typeof raw === 'string') { try { raw = JSON.parse(raw); } catch (e2) { raw = null; } }
    return raw;
  }

  /** The rules: from the mailbox store, else the old roaming setting, else the local backup. { rules, where } */
  async function loadRules() {
    try { var r = await loadRulesRemote(); if (r) return { rules: r, where: 'store' }; } catch (e) { console.warn('rule store unreadable, using fallbacks', e); }
    var legacy = loadRulesLegacy();
    if (legacy) return { rules: legacy, where: 'roaming' };
    var backup = local.get('is.rules.backup');
    return { rules: backup, where: backup ? 'backup' : 'none' };
  }

  var saving = null, pending = null;
  /** Save to the mailbox store (one write at a time; a burst of edits collapses into the last). Resolves { ok, size }. */
  function saveRules(rules) {
    var text = JSON.stringify(rules);
    local.set('is.rules.backup', rules);
    if (global.SORTER_MOCK) { try { Office.context.roamingSettings.set('rules', text); } catch (e) { /* mock */ } }
    pending = text;
    if (!saving) saving = (async function loop() {
      while (pending !== null) { var t = pending; pending = null; try { await saveRulesRemote(t); } catch (e) { console.error('rule store save failed', e); saving = null; return { ok: false, size: t.length }; } }
      saving = null; return { ok: true, size: text.length };
    })();
    return saving.then(function (r) { return { ok: r.ok, size: text.length }; });
  }

  global.SorterGraph = {
    SetupError: SetupError, GraphError: GraphError,
    initAuth: initAuth, me: me, listInbox: listInbox, sentRecipients: sentRecipients, messageHeaders: messageHeaders, messageInfo: messageInfo, senderCounts: senderCounts,
    moveMessage: moveMessage, setRead: setRead, inboxCountFrom: inboxCountFrom, listInboxFrom: listInboxFrom, listInboxAll: listInboxAll, wellKnownFolderId: wellKnownFolderId, listRules: listRules, createRule: createRule, deleteRule: deleteRule, listFolders: listFolders, createFolder: createFolder,
    local: local, loadRules: loadRules, saveRules: saveRules
  };
})(typeof self !== 'undefined' ? self : this);
