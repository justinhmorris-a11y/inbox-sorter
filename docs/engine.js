/* Inbox Sorter - sorting engine (no Outlook / network code in here, so it can be unit tested).
 *
 * Principles
 *   - Mail only ever moves because of a rule the user has APPROVED. The engine merely SUGGESTS rules.
 *   - People the user writes to, colleagues and anything unclear always stay in the Inbox.
 *   - Safety net: recent mail that is flagged, high importance, or whose subject looks like it needs
 *     action stays in the Inbox even when its sender has a rule.
 *   - Order rescue: a genuine order / dispatch / refund mail from a sender whose mail is normally
 *     treated as newsletters, deleted or junked is filed under Receipts instead.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.SorterEngine = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // Bucket codes are short because the rule set lives in Outlook roaming settings (32 KB cap).
  var BUCKETS = {
    I: { name: 'Inbox', label: 'Keep in inbox', folder: null },
    N: { name: 'Newsletters', label: 'Newsletters', folder: 'Newsletters' },
    R: { name: 'Receipts', label: 'Receipts', folder: 'Receipts' },
    T: { name: 'Notifications', label: 'Notifications', folder: 'Notifications' },
    D: { name: 'Delete', label: 'Delete', folder: 'deleteditems' },
    J: { name: 'Junk', label: 'Junk', folder: 'junkemail' }
  };

  function rx(parts) { return new RegExp(parts.join('|'), 'i'); }

  var RX_ACTION = rx([
    'payment.{0,25}(fail|declin|unsuccess|problem|issue|overdue|\\bdue\\b|requir|could ?n.t|unable|reject|retry|update)',
    '(fail|declin|unsuccess|unable|could ?n.t|problem|issue).{0,25}(payment|charge|card|direct debit|renew|deliver|process|bill|order)',
    'card.{0,20}(declin|expir|fail|updat)',
    '(action|response|attention|reply|signature|approval|verification|confirmation) (is )?(required|needed|requested)',
    'requires? your (attention|action|approval|response)',
    '\\burgent\\b', '\\boverdue\\b', 'past due', 'final (notice|reminder|demand|warning)', 'last chance to (pay|renew)',
    '(will|about to|set to|going to|due to) (expire|be (suspended|cancell?ed|closed|deleted|deactivated|removed))',
    '(has|have) (expired|been (suspended|cancell?ed|declined|locked|disabled|deactivated|put on hold))',
    'expir(es|ing) (soon|today|tomorrow|in \\d+|on )',
    '(account|subscription|service|domain|licen[cs]e|certificate|policy|membership|insurance|\\bmot\\b|\\btax\\b).{0,25}(suspend|expir|cancel|locked|on hold|at risk|renewal|\\bdue\\b)',
    'security alert', 'unusual (sign|activity|login)', 'access from (a )?new', 'new (device|computer|login|location)', 'unrecogni[sz]ed', 'password (was )?(reset|changed)', 'suspicious', 'new sign-?in', 'verify your', 'confirm your (email|account|identity|details)',
    '(missed|failed|unable to) deliver', 'delivery (attempt|failed|failure|problem|exception)', 'we missed you', 'sorry we missed',
    '(invoice|bill|balance|amount|payment).{0,20}(overdue|outstanding|unpaid|is due|now due)',
    'refund.{0,20}(fail|declin|problem|issue)',
    '(please|kindly) (sign|review and sign|approve|respond|confirm|complete|pay|update)',
    'docusign', 'adobe sign', 'signature request', 'waiting for your',
    '\\bdeadline\\b', '\\breminder\\b'
  ]);

  var RX_TRANS = rx([
    '\\border(ed|s)?\\b', 'cancell?ed', 'receipt', 'invoice', 'payment', '\\bpaid\\b', 'dispatch', 'shipped', 'shipment', 'deliver', 'parcel', 'tracking',
    'booking', 'reservation', 'itinerary', 'e-?ticket', 'statement', 'subscription', 'renewal', 'refund', 'purchase', 'transaction',
    'confirmation', 'confirmed', 'your account', 'direct debit', '\\bbill\\b', 'returns?\\b', 'warranty', 'policy', 'premium', 'top-?up'
  ]);

  var RX_TRANS_STRONG = rx([
    'your order', 'order (confirm|number|no\\.?|#|ref|update|received|dispatched|cancell?ed|summary)', '\\bordered\\b', 'dispatched', 'has (been )?shipped',
    'out for delivery', '\\bdelivered\\b', 'cancell?ed', 'refund', 'receipt', 'invoice', 'return (label|received|request)', 'ready (for|to) collect',
    'booking (confirm|reference|ref)', 'your (booking|reservation|itinerary|e-?tickets?)', 'payment (received|confirm)'
  ]);

  var RX_NOREPLY = /no[-_.]?reply|do[-_.]?not[-_.]?reply|mailer-daemon|postmaster|bounce/i;
  var RX_NOTIF_LOCAL = /notif|alert|monitor|status|build|deploy|pipeline|incident|error|exception|system|automat|robot|daemon|jira|jenkins|teamcity|octopus|devops|sentry|nagios|zabbix|cron|backup/i;
  var RX_NEWS_LOCAL = /news|letter|marketing|promo|offer|deals?|digest|updates?|campaign|mailing|subscri|community|events?|webinar|^(hello|hi|team|info|email|mail|contact)$/i;
  var RX_NEWS_DOMAIN = /^(e|em|email|emails|mail|mailer|mailing|news|newsletter|marketing|info|mg|send|t|click|go|reply)\d*\./i;
  var RX_DEV_DOMAIN = /(^|\.)(github\.com|gitlab\.com|bitbucket\.org|atlassian\.(net|com)|visualstudio\.com|azure\.com|sentry\.io|pagerduty\.com|datadoghq\.(com|eu)|newrelic\.com|statuspage\.io|uptimerobot\.com|pingdom\.com|opsgenie\.(net|com)|octopus\.(com|app)|circleci\.com|travis-ci\.(com|org)|appveyor\.com|sonarcloud\.io|snyk\.io|vercel\.com|netlify\.com|heroku\.com|amazonaws\.com|supabase\.(io|com)|npmjs\.com|nuget\.org|docker\.com|slack\.com|trello\.com|asana\.com|linear\.app|raygun\.(io|com)|bugsnag\.com|rollbar\.com|elmah\.io|papertrailapp\.com|loggly\.com)$/i;
  var RX_ESP_HEADER = /^(x-mailgun-|x-sg-eid|x-mc-user|x-mandrill-user|x-ses-outgoing|feedback-id|x-csa-complaints|x-campaign|x-mailchimp|x-hubspot|x-sfmc|x-klaviyo)/i;

  function addressParts(address) {
    var a = String(address || '').toLowerCase();
    var at = a.indexOf('@');
    if (at < 1) return null;
    return { local: a.slice(0, at), domain: a.slice(at + 1) };
  }

  // mail.shop.co.uk -> shop.co.uk ; email.sportsdirect.com -> sportsdirect.com
  function registrableDomain(domain) {
    var labels = String(domain || '').toLowerCase().split('.');
    if (labels.length <= 2) return labels.join('.');
    var last = labels[labels.length - 1], second = labels[labels.length - 2];
    var twoPart = last.length === 2 && /^(co|com|org|net|gov|ac|ltd|plc|me|sch|nhs)$/.test(second);
    return labels.slice(twoPart ? -3 : -2).join('.');
  }

  function emptyRules() { return { v: 1, senders: {}, domains: {}, keep: {} }; }

  function normaliseRules(r) {
    var out = emptyRules();
    if (r && typeof r === 'object') {
      ['senders', 'domains', 'keep'].forEach(function (k) { if (r[k] && typeof r[k] === 'object') out[k] = r[k]; });
    }
    return out;
  }

  /** The rule that applies to an address, or null. Sender rules beat domain rules; the most specific domain wins. */
  function ruleFor(rules, address) {
    var a = String(address || '').toLowerCase();
    if (rules.senders[a]) return { bucket: rules.senders[a], scope: 'sender', key: a };
    var p = addressParts(a);
    if (!p) return null;
    var d = p.domain;
    while (d) {
      if (rules.domains[d]) return { bucket: rules.domains[d], scope: 'domain', key: d };
      var dot = d.indexOf('.');
      if (dot < 0) break;
      d = d.slice(dot + 1);
    }
    return null;
  }

  function bucketName(code) {
    if (BUCKETS[code]) return BUCKETS[code].name;
    if (typeof code === 'string' && code.indexOf('F:') === 0) return code.slice(2);
    return String(code);
  }

  function readHeaders(headers) {
    var ev = { unsub: false, bulk: false, auto: false, esp: false };
    (headers || []).forEach(function (h) {
      if (!h || !h.name) return;
      var n = String(h.name).toLowerCase(), v = String(h.value || '');
      if (n === 'list-unsubscribe' || n === 'list-unsubscribe-post') ev.unsub = true;
      else if (n === 'list-id') ev.bulk = true;
      else if (n === 'precedence' && /bulk|list|junk/i.test(v)) ev.bulk = true;
      else if (n === 'auto-submitted' && !/^\s*no\s*$/i.test(v)) ev.auto = true;
      else if (RX_ESP_HEADER.test(n)) ev.esp = true;
    });
    return ev;
  }

  /** Per-sender totals for a slice of messages. */
  function buildSenders(messages) {
    var map = {};
    messages.forEach(function (m) {
      var s = map[m.from];
      if (!s) { s = map[m.from] = { address: m.from, name: m.fromName || m.from, total: 0, unread: 0, transN: 0, actionN: 0, latest: null }; }
      s.total++;
      if (!m.isRead) s.unread++;
      if (RX_TRANS.test(m.subject || '')) s.transN++;
      if (RX_ACTION.test(m.subject || '')) s.actionN++;
      if (!s.latest || m.received > s.latest.received) { s.latest = m; if (m.fromName) s.name = m.fromName; }
    });
    return map;
  }

  /**
   * What the engine thinks of a sender that has NO rule yet.
   * ctx: { sentTo: {address: count}, myDomains: {domain: true}, evidence: {address: {unsub,bulk,auto,esp}} }
   * returns { kind: 'people' | 'suggest' | 'unknown', bucket?: code, why: string }
   */
  function assessSender(s, ctx) {
    var p = addressParts(s.address);
    if (!p) {
      if (s.address === '(no sender)') return { kind: 'unknown', why: 'no sender address' };
      return { kind: 'people', why: 'internal (Exchange address)' };
    }
    var sent = (ctx.sentTo && ctx.sentTo[s.address]) || 0;
    if (sent >= 1) return { kind: 'people', why: 'you have written to them' + (sent > 1 ? ' x' + sent : '') };
    if (ctx.myDomains && ctx.myDomains[p.domain] && !RX_NOREPLY.test(p.local) && !RX_NOTIF_LOCAL.test(p.local)) {
      return { kind: 'people', why: 'colleague' };
    }
    var ev = (ctx.evidence && ctx.evidence[s.address]) || {};
    var transPct = s.total ? Math.round(100 * s.transN / s.total) : 0;
    var automated = RX_NOREPLY.test(p.local) || ev.auto || ev.esp || ev.unsub || ev.bulk;

    if (RX_DEV_DOMAIN.test(p.domain)) return { kind: 'suggest', bucket: 'T', why: 'dev / ops tool' };
    if (RX_NOTIF_LOCAL.test(p.local)) return { kind: 'suggest', bucket: 'T', why: 'alert-style address' };
    if (ev.unsub) return { kind: 'suggest', bucket: 'N', why: 'has an unsubscribe link' };
    if (ev.bulk) return { kind: 'suggest', bucket: 'N', why: 'bulk-mail header' };
    if (s.total >= 2 && transPct >= 50) return { kind: 'suggest', bucket: 'R', why: transPct + '% order / account subjects' };
    if (s.total < 2 && s.transN >= 1 && automated) return { kind: 'suggest', bucket: 'R', why: 'order / account subject' };
    if (RX_NOREPLY.test(p.local)) return { kind: 'suggest', bucket: 'T', why: 'no-reply address' };
    if (ev.auto) return { kind: 'suggest', bucket: 'T', why: 'automated sender' };
    if (RX_NEWS_LOCAL.test(p.local) || RX_NEWS_DOMAIN.test(p.domain)) return { kind: 'suggest', bucket: 'N', why: 'looks like a mailing address' };
    return { kind: 'unknown', why: 'no clear signal' };
  }

  /**
   * Decide every message.
   * opts: { now: Date, recentDays: number }
   * returns rows: { msg, dest: code ('I' = stays), group, why, rule }
   *   group: 'file' | 'kept' (safety net) | 'keep1' (user kept this one) | 'rule-inbox' | 'people' | 'suggest' | 'unknown'
   */
  function plan(messages, senders, rules, ctx, opts) {
    var now = (opts && opts.now) || new Date();
    var recentMs = ((opts && opts.recentDays) || 90) * 86400000;
    var assessments = {};
    Object.keys(senders).forEach(function (a) { assessments[a] = assessSender(senders[a], ctx); });

    var rows = messages.map(function (m) {
      var rule = ruleFor(rules, m.from);
      var subject = m.subject || '';
      if (!rule) {
        var as = assessments[m.from] || { kind: 'unknown', why: 'no clear signal' };
        return { msg: m, dest: 'I', group: as.kind, why: as.why, rule: null, suggestion: as.kind === 'suggest' ? as.bucket : null };
      }
      if (rule.bucket === 'I') return { msg: m, dest: 'I', group: 'rule-inbox', why: 'your rule: keep in inbox', rule: rule };

      var bucket = rule.bucket, why = 'your rule';
      if ((bucket === 'N' || bucket === 'D' || bucket === 'J') && RX_TRANS_STRONG.test(subject)) {
        why = 'looks like a real order, so Receipts instead of ' + bucketName(bucket); bucket = 'R';
      }
      if (rules.keep[m.id]) return { msg: m, dest: 'I', group: 'keep1', why: 'you chose to keep this one', rule: rule, wouldBe: bucket };

      var recent = (now - m.received) <= recentMs;
      var protect = '';
      if (recent) {
        if (m.flagged) protect = 'flagged';
        else if (m.importance === 'high') protect = 'high importance';
        else if (RX_ACTION.test(subject)) protect = 'subject looks like it needs action';
      }
      if (protect) return { msg: m, dest: 'I', group: 'kept', why: protect, rule: rule, wouldBe: bucket };
      return { msg: m, dest: bucket, group: 'file', why: why, rule: rule };
    });
    return { rows: rows, assessments: assessments };
  }

  return {
    BUCKETS: BUCKETS, bucketName: bucketName, addressParts: addressParts, registrableDomain: registrableDomain,
    emptyRules: emptyRules, normaliseRules: normaliseRules, ruleFor: ruleFor, readHeaders: readHeaders,
    buildSenders: buildSenders, assessSender: assessSender, plan: plan,
    _rx: { action: RX_ACTION, trans: RX_TRANS, strong: RX_TRANS_STRONG }
  };
});
