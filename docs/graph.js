/* Inbox Sorter - sign-in, Microsoft Graph calls and storage.
 * Sign-in uses nested app authentication: Outlook itself brokers the token, so there is no
 * separate login window in normal use.
 */
(function (global) {
  'use strict';

  var GRAPH = 'https://graph.microsoft.com/v1.0';
  var SCOPES = ['Mail.ReadWrite', 'User.Read'];
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
      if (options.immutable !== false) headers.Prefer = 'IdType="ImmutableId"';
      if (options.body) headers['Content-Type'] = 'application/json';
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
      importance: m.importance || 'normal'
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

  async function messageHeaders(id) {
    var m = await call('/me/messages/' + encodeURIComponent(id) + '?$select=internetMessageHeaders');
    return m.internetMessageHeaders || [];
  }

  /** One message by id (any id format): who sent it, when, and its header evidence. */
  async function messageInfo(id) {
    var m = await call('/me/messages/' + encodeURIComponent(id) + '?$select=id,from,sender,subject,receivedDateTime,isRead,flag,importance,internetMessageHeaders');
    var out = mapMessage(m); out.headers = m.internetMessageHeaders || [];
    return out;
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

  function loadRules() {
    var raw = null;
    try { raw = Office.context.roamingSettings.get('rules'); } catch (e) { raw = null; }
    if (typeof raw === 'string') { try { raw = JSON.parse(raw); } catch (e2) { raw = null; } }
    if (!raw) raw = local.get('is.rules.backup');
    return raw;
  }

  function saveRules(rules) {
    var text = JSON.stringify(rules);
    local.set('is.rules.backup', rules);
    return new Promise(function (resolve) {
      try {
        Office.context.roamingSettings.set('rules', text);
        Office.context.roamingSettings.saveAsync(function (r) { resolve({ ok: r && r.status === 'succeeded', size: text.length }); });
      } catch (e) { resolve({ ok: false, size: text.length }); }
    });
  }

  global.SorterGraph = {
    SetupError: SetupError, GraphError: GraphError,
    initAuth: initAuth, me: me, listInbox: listInbox, sentRecipients: sentRecipients, messageHeaders: messageHeaders, messageInfo: messageInfo,
    moveMessage: moveMessage, listFolders: listFolders, createFolder: createFolder,
    local: local, loadRules: loadRules, saveRules: saveRules
  };
})(typeof self !== 'undefined' ? self : this);
