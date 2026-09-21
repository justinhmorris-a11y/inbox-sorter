/* Inbox Sorter - preview mode. Add ?mock=1 to the URL to try the panel in an ordinary browser with a
   made-up mailbox. Does nothing at all inside Outlook. */
(function (global) {
  'use strict';
  if (!/[?&]mock=1/.test(global.location.search)) return;
  global.SORTER_MOCK = true;

  var settings = {}, handlers = {}, highlighted = [];
  global.MockSelect = function (ids) { highlighted = ids; if (handlers.selectedItemsChanged) handlers.selectedItemsChanged(); };
  var item = { from: { emailAddress: 'donotreply@email.sportsdirect.com', displayName: 'Sports Direct' }, subject: 'Outlet savings under £50', dateTimeCreated: new Date(Date.now() - 3 * 86400000) };
  global.Office = {
    onReady: function (cb) { setTimeout(cb, 0); return Promise.resolve(); },
    EventType: { ItemChanged: 'itemChanged', SelectedItemsChanged: 'selectedItemsChanged' },
    MailboxEnums: { RestVersion: { v2_0: 'v2.0' } },
    context: {
      requirements: { isSetSupported: function () { return true; } },
      roamingSettings: { get: function (k) { return settings[k]; }, set: function (k, v) { settings[k] = v; }, saveAsync: function (cb) { cb({ status: 'succeeded' }); } },
      mailbox: { item: item, addHandlerAsync: function (type, fn) { handlers[type] = fn; }, convertToRestId: function (id) { return id; },
        getSelectedItemsAsync: function (cb) { if (/[?&]noselect=1/.test(global.location.search)) throw new Error("Elevated permission is required to call the method: 'getSelectedItemsAsync'."); cb({ status: 'succeeded', value: highlighted.map(function (id) { return { itemId: id, itemType: 'message' }; }) }); } }
    }
  };

  var now = new Date(), n = 0;
  function at(daysAgo, h, m) { return new Date(now.getFullYear(), now.getMonth(), now.getDate() - daysAgo, h, m).toISOString(); }
  function mail(addr, name, subject, daysAgo, h, m, o) {
    o = o || {};
    return { id: 'MOCK-' + (++n), from: { emailAddress: { address: addr, name: name } }, subject: subject, receivedDateTime: at(daysAgo, h, m),
      isRead: !!o.read, flag: { flagStatus: o.flagged ? 'flagged' : 'notFlagged' }, importance: o.high ? 'high' : 'normal' };
  }
  var inbox = [
    mail('donotreply@email.sportsdirect.com', 'Sports Direct', 'Outlet savings under £50', 0, 8, 31, { read: true }),
    mail('reports@dezrezlegal.example', 'Legal Reporting', 'Extraction report - September', 0, 8, 30, { read: true }),
    mail('friendupdates@facebookmail.com', 'Facebook', 'Daisy posted something new', 0, 8, 26),
    mail('amazon-offers@amazon.co.uk', 'Amazon.co.uk', "Prime member, don't forget your 2 free audiobooks", 0, 8, 5),
    mail('info-uk@epsa.com', 'EPSA', 'Download the 2026 Innovation Loans guide', 0, 8, 3),
    mail('ross@example-colleague.com', 'Ross', 'RE: Keystone sprint planning', 0, 7, 50),
    mail('auto-confirm@amazon.co.uk', 'Amazon.co.uk', "Ordered: 'Notebook A5, 3 pack'", 0, 7, 40, { read: true }),
    mail('shipment-tracking@amazon.co.uk', 'Amazon.co.uk', "Dispatched: 'Notebook A5, 3 pack'", 0, 7, 20),
    mail('payments-messages@amazon.co.uk', 'Amazon Payments', 'Payment declined: please update your payment method', 0, 7, 10),
    mail('info@emails.golfbreaks.com', 'Golfbreaks', 'Justin, make the Algarve your next stop', 0, 6, 30),
    mail('info@emails.golfbreaks.com', 'Golfbreaks', 'Your booking confirmation GB998877', 0, 6, 10),
    mail('service@paypal.co.uk', 'PayPal', 'Receipt for your payment to Steam', 0, 6, 5),
    mail('service@paypal.co.uk', 'PayPal', 'Receipt for your payment to Xero Shoes', 0, 5, 45),
    mail('noreply@steampowered.com', 'Steam Support', 'Your Steam account: Access from new computer', 0, 5, 30),
    mail('newsletter@progressiveguitar.example', 'Progressive Guitar', "I'm Yours - Jason Mraz", 0, 5, 0),
    mail('notifications@github.com', 'GitHub', '[dezrez/app] Run failed: CI - main', 0, 4, 0),
    mail('andy@example-friend.com', 'Andy Moore', 'The Community Cup | Saturday', 0, 3, 0, { read: true }),
    mail('ceo@partner.example', 'Partner CEO', 'Contract for signature', 0, 2, 0, { flagged: true }),
    mail('donotreply@email.sportsdirect.com', 'Sports Direct', 'Your order SD12345 has been dispatched', 1, 15, 0),
    mail('donotreply@email.sportsdirect.com', 'Sports Direct', 'New season arrivals', 1, 9, 0),
    mail('amazon-offers@amazon.co.uk', 'Amazon.co.uk', 'Discover our tablets for every budget', 1, 9, 30),
    mail('info@emails.golfbreaks.com', 'Golfbreaks', 'Golfer, make the Algarve your next stop', 1, 8, 0)
  ];
  var headers = { unsub: [{ name: 'List-Unsubscribe', value: '<x>' }] };
  var unsubSenders = { 'donotreply@email.sportsdirect.com': 1, 'friendupdates@facebookmail.com': 1, 'amazon-offers@amazon.co.uk': 1, 'info-uk@epsa.com': 1, 'info@emails.golfbreaks.com': 1, 'newsletter@progressiveguitar.example': 1 };
  var folders = [{ id: 'F-inbox', displayName: 'Inbox' }, { id: 'F-audit', displayName: 'Audit' }, { id: 'F-dez', displayName: 'DezRez' }, { id: 'F-hipz', displayName: 'Hipz' }, { id: 'F-ski', displayName: 'Ski erg' }, { id: 'F-trips', displayName: 'Trips' }];
  var moved = {};
  global.MockLog = { moves: [], created: [] };

  global.MockGraph = {
    handle: function (method, url, body) {
      return new Promise(function (resolve, reject) {
        setTimeout(function () {
          try { resolve(route(method, decodeURIComponent(url), body)); } catch (e) { reject(e); }
        }, 25);
      });
    }
  };

  function route(method, url, body) {
    var m;
    if (method === 'POST' && /\/me\/mailFolders$/.test(url)) { var f = { id: 'F-' + body.displayName, displayName: body.displayName }; folders.push(f); global.MockLog.created.push(body.displayName); return f; }
    if (method === 'POST' && (m = /\/me\/messages\/([^/]+)\/move$/.exec(url))) {
      if (body.destinationId === 'inbox') delete moved[m[1]]; else moved[m[1]] = body.destinationId;
      global.MockLog.moves.push({ id: m[1], to: body.destinationId });
      return { id: m[1] };
    }
    if (/\/me\?/.test(url)) return { displayName: 'Justin', mail: 'justin@example-colleague.com', userPrincipalName: 'justin@example-colleague.com', proxyAddresses: [] };
    if (/sentitems\/messages/.test(url)) return { value: [{ toRecipients: [{ emailAddress: { address: 'andy@example-friend.com' } }], ccRecipients: [] }] };
    if (/inbox\/childFolders/.test(url)) return { value: [] };
    if (/\/me\/mailFolders\?/.test(url)) return { value: folders };
    if ((m = /\/me\/messages\/([^?]+)\?\$select=id,from/.exec(url))) {
      var one = inbox.filter(function (x) { return x.id === m[1]; })[0];
      if (!one) throw new Error('mock: no such message');
      var copy = JSON.parse(JSON.stringify(one)); copy.internetMessageHeaders = unsubSenders[one.from.emailAddress.address] ? headers.unsub : [];
      return copy;
    }
    if ((m = /\/me\/messages\/([^?]+)\?\$select=internetMessageHeaders/.exec(url))) {
      var msg = inbox.filter(function (x) { return x.id === m[1]; })[0];
      return { internetMessageHeaders: msg && unsubSenders[msg.from.emailAddress.address] ? headers.unsub : [] };
    }
    if (/inbox\/messages/.test(url)) {
      var ge = /ge (\S+Z)/.exec(url), lt = /lt (\S+Z)/.exec(url);
      return { value: inbox.filter(function (x) { return !moved[x.id] && x.receivedDateTime >= ge[1].replace('Z', '.000Z') && x.receivedDateTime < lt[1].replace('Z', '.000Z'); }) };
    }
    throw new Error('mock: unexpected ' + method + ' ' + url);
  }
})(window);
