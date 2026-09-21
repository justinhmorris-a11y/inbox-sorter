const E = require('../docs/engine.js');
const assert = require('assert');
const now = new Date('2026-09-21T09:00:00Z');
let n = 0;
const msg = (from, fromName, subject, daysAgo, o = {}) => ({ id: 'm' + (++n), from, fromName, subject, received: new Date(now - daysAgo * 86400000), isRead: !!o.read, flagged: !!o.flagged, importance: o.importance || 'normal' });
const messages = [
  msg('colleague@dezrez.com', 'A Colleague', 'RE: Project sprint', 0),
  msg('new.person@dezrez.com', 'New Person', 'Intro', 0),
  msg('donotreply@email.sportsdirect.com', 'Sports Direct', 'Outlet savings under £50', 0),
  msg('donotreply@email.sportsdirect.com', 'Sports Direct', 'Your order SD123 has been dispatched', 0),
  msg('donotreply@email.sportsdirect.com', 'Sports Direct', 'Free delivery on orders over £50', 0),
  msg('info@emails.golfbreaks.com', 'Golfbreaks', 'Justin, make the Algarve your next stop', 0),
  msg('info@emails.golfbreaks.com', 'Golfbreaks', 'Action required: balance payment is due', 1),
  msg('info@emails.golfbreaks.com', 'Golfbreaks', 'Your booking confirmation GB99', 1),
  msg('payments-messages@amazon.co.uk', 'Amazon', 'Payment declined: update your payment method', 2),
  msg('payments-messages@amazon.co.uk', 'Amazon', 'Payment declined: update your payment method', 800),
  msg('auto-confirm@amazon.co.uk', 'Amazon.co.uk', "Ordered: 'Notebook A5'", 0),
  msg('auto-confirm@amazon.co.uk', 'Amazon.co.uk', "Delivered: 'Notebook A5'", 0),
  msg('amazon-offers@amazon.co.uk', 'Amazon.co.uk', "Prime member, don't forget", 0),
  msg('notifications@github.com', 'GitHub', '[dezrez/app] Run failed: CI', 0),
  msg('reports@dezrezlegal.co.uk', 'Dezrez Legal Reporting', 'Extraction report - September 2026', 0),
  msg('info-uk@epsa.com', 'EPSA', 'Download the 2026 Innovation Loans guide', 0),
  msg('ceo@partner.com', 'Partner', 'Contract', 0, { flagged: true }),
];
const ctx = {
  sentTo: { 'colleague@dezrez.com': 30 }, myDomains: { 'dezrez.com': true },
  evidence: { 'donotreply@email.sportsdirect.com': { unsub: true }, 'amazon-offers@amazon.co.uk': { unsub: true }, 'info-uk@epsa.com': { unsub: true }, 'info@emails.golfbreaks.com': { unsub: true } }
};
const senders = E.buildSenders(messages);
const by = (rows, subj) => rows.find(r => r.msg.subject.startsWith(subj));

// 1. No rules at all: NOTHING moves, only suggestions
let p = E.plan(messages, senders, E.emptyRules(), ctx, { now });
assert.strictEqual(p.rows.filter(r => r.dest !== 'I').length, 0, 'nothing moves without approved rules');
assert.strictEqual(p.assessments['colleague@dezrez.com'].kind, 'people');
assert.strictEqual(p.assessments['new.person@dezrez.com'].kind, 'people');
assert.deepStrictEqual([p.assessments['donotreply@email.sportsdirect.com'].kind, p.assessments['donotreply@email.sportsdirect.com'].bucket], ['suggest', 'N']);
assert.strictEqual(p.assessments['notifications@github.com'].bucket, 'T');
assert.strictEqual(p.assessments['auto-confirm@amazon.co.uk'].bucket, 'R');
assert.strictEqual(p.assessments['payments-messages@amazon.co.uk'].bucket, 'R');
assert.strictEqual(p.assessments['reports@dezrezlegal.co.uk'].kind, 'unknown');
assert.strictEqual(p.assessments['ceo@partner.com'].kind, 'unknown');

// 2. Approved rules
const rules = E.normaliseRules({ senders: { 'donotreply@email.sportsdirect.com': 'N', 'payments-messages@amazon.co.uk': 'R', 'auto-confirm@amazon.co.uk': 'R', 'reports@dezrezlegal.co.uk': 'F:DezRez' }, domains: { 'golfbreaks.com': 'D', 'epsa.com': 'J', 'github.com': 'T' } });
p = E.plan(messages, senders, rules, ctx, { now });
assert.strictEqual(by(p.rows, 'Outlet savings').dest, 'N');
assert.strictEqual(by(p.rows, 'Free delivery').dest, 'N', 'marketing mention of delivery is still a newsletter');
assert.strictEqual(by(p.rows, 'Your order SD123').dest, 'R', 'real order rescued to Receipts');
assert.strictEqual(by(p.rows, 'Justin, make the Algarve').dest, 'D');
assert.strictEqual(by(p.rows, 'Action required: balance').dest, 'I', 'safety net beats delete rule');
assert.strictEqual(by(p.rows, 'Action required: balance').group, 'kept');
assert.strictEqual(by(p.rows, 'Your booking confirmation').dest, 'R', 'booking rescued from delete');
const declined = p.rows.filter(r => r.msg.subject.startsWith('Payment declined'));
assert.deepStrictEqual(declined.map(r => r.dest).sort(), ['I', 'R'], 'recent payment problem stays, 800-day-old one is history');
assert.strictEqual(by(p.rows, 'Ordered:').dest, 'R');
assert.strictEqual(by(p.rows, '[dezrez/app]').dest, 'T');
assert.strictEqual(by(p.rows, 'Extraction report').dest, 'F:DezRez');
assert.strictEqual(E.bucketName('F:DezRez'), 'DezRez');
assert.strictEqual(by(p.rows, 'Download the 2026').dest, 'J');
assert.strictEqual(by(p.rows, 'Contract').dest, 'I');
assert.strictEqual(by(p.rows, 'RE: Project').dest, 'I');
assert.strictEqual(by(p.rows, 'Prime member').dest, 'I', 'no rule yet -> stays, as a suggestion');
assert.strictEqual(by(p.rows, 'Prime member').suggestion, 'N');

// 3. keep-this-one, reject (= Inbox rule), sender beats domain
rules.keep[by(p.rows, 'Outlet savings').msg.id] = 1;
rules.senders['info@emails.golfbreaks.com'] = 'I';
p = E.plan(messages, senders, rules, ctx, { now });
assert.strictEqual(by(p.rows, 'Outlet savings').dest, 'I');
assert.strictEqual(by(p.rows, 'Outlet savings').group, 'keep1');
assert.strictEqual(by(p.rows, 'Justin, make the Algarve').group, 'rule-inbox');

assert.strictEqual(E.registrableDomain('email.sportsdirect.com'), 'sportsdirect.com');
assert.strictEqual(E.registrableDomain('mail.shop.co.uk'), 'shop.co.uk');
assert.strictEqual(E.registrableDomain('amazon.co.uk'), 'amazon.co.uk');
assert.deepStrictEqual(E.readHeaders([{ name: 'List-Unsubscribe', value: '<x>' }, { name: 'Auto-Submitted', value: 'no' }]), { unsub: true, bulk: false, auto: false, esp: false });
console.log('engine tests passed:', p.rows.length, 'messages planned');
