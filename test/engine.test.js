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
// subject rules: sender + subject text, beat sender/domain rules, never touch the sender's other mail
const subjRules = E.normaliseRules({ senders: { 'ross@dezrez.com': 'I' }, subjects: [{ from: 'ross@dezrez.com', has: 'demo booked', code: 'F:Demos*' }] });
assert.deepStrictEqual(E.ruleFor(subjRules, 'ross@dezrez.com', 'Demo booked - Acme'), { bucket: 'F:Demos', read: true, arrival: false, scope: 'subject', key: 'ross@dezrez.com', has: 'demo booked' });
assert.strictEqual(E.ruleFor(subjRules, 'ross@dezrez.com', 'RE: Project sprint').scope, 'sender', 'other subjects fall through to the sender rule');
assert.strictEqual(E.ruleFor(subjRules, 'ross@dezrez.com').scope, 'sender', 'no subject given: sender rule');
assert.strictEqual(E.subjectRuleFor(subjRules, 'other@dezrez.com', 'Demo booked'), null, 'subject rule is per sender');
const subjMsgs = messages.concat([msg('colleague@dezrez.com', 'A Colleague', 'Demo booked - Acme Ltd', 0), msg('colleague@dezrez.com', 'A Colleague', 'Demo booked - URGENT reply needed', 0)]);
const subjRules2 = E.normaliseRules({ subjects: [{ from: 'colleague@dezrez.com', has: 'Demo booked', code: 'F:Demos' }] });
p = E.plan(subjMsgs, E.buildSenders(subjMsgs), subjRules2, ctx, { now });
assert.strictEqual(by(p.rows, 'Demo booked - Acme').dest, 'F:Demos');
assert.strictEqual(by(p.rows, 'RE: Project').dest, 'I', 'colleague\'s normal mail stays');
assert.strictEqual(by(p.rows, 'Demo booked - URGENT').dest, 'I', 'safety net still applies to subject rules');
assert.strictEqual(E.normaliseRules({ subjects: [{ from: 'x' }, null, { from: 'a@b.c', has: 'q', code: 'N' }] }).subjects.length, 1, 'broken subject rules dropped');

// arrival flag: '^' on a code; defaults for junk / delete / read newsletters; one-off migration of old rules
assert.deepStrictEqual(E.splitCode('N*^'), { bucket: 'N', read: true, arrival: true });
assert.deepStrictEqual(E.splitCode('J^'), { bucket: 'J', read: false, arrival: true });
assert.strictEqual(E.joinCode('N', true, true), 'N*^'); assert.strictEqual(E.joinCode('R', false, false), 'R'); assert.strictEqual(E.joinCode('I', true, true), 'I');
assert.strictEqual(E.arrivalDefault('J', false), true); assert.strictEqual(E.arrivalDefault('N', false), false); assert.strictEqual(E.arrivalDefault('N', true), true); assert.strictEqual(E.arrivalDefault('R', true), false);
const old = E.normaliseRules({ senders: { 'a@x.com': 'J', 'b@x.com': 'N*', 'c@x.com': 'N', 'd@x.com': 'R*' }, domains: { 'spam.com': 'D' }, subjects: [{ from: 'e@x.com', has: 'digest', code: 'N*' }] });
assert.strictEqual(E.applyArrivalDefaults(old), true);
assert.deepStrictEqual(old.senders, { 'a@x.com': 'J^', 'b@x.com': 'N*^', 'c@x.com': 'N', 'd@x.com': 'R*' }); assert.strictEqual(old.domains['spam.com'], 'D^'); assert.strictEqual(old.subjects[0].code, 'N*^'); assert.strictEqual(old.v, 2);
assert.strictEqual(E.applyArrivalDefaults(old), false, 'runs once');
assert.strictEqual(E.ruleFor(old, 'a@x.com').arrival, true); assert.strictEqual(E.ruleFor(old, 'c@x.com').arrival, false);

// 'mark as read' flag: stored as a '*' suffix, stripped by ruleFor, surfaced as markRead only for unread mail that files
assert.deepStrictEqual(E.splitCode('N*'), { bucket: 'N', read: true, arrival: false });
assert.deepStrictEqual(E.splitCode('F:DezRez'), { bucket: 'F:DezRez', read: false, arrival: false });
assert.strictEqual(E.joinCode('N', true), 'N*'); assert.strictEqual(E.joinCode('I', true), 'I');
const readRules = E.normaliseRules({ senders: { 'donotreply@email.sportsdirect.com': 'N*' }, domains: { 'golfbreaks.com': 'D' } });
assert.deepStrictEqual(E.ruleFor(readRules, 'donotreply@email.sportsdirect.com'), { bucket: 'N', read: true, arrival: false, scope: 'sender', key: 'donotreply@email.sportsdirect.com' });
assert.strictEqual(E.ruleFor(readRules, 'info@emails.golfbreaks.com').read, false);
p = E.plan(messages, senders, readRules, ctx, { now });
assert.strictEqual(by(p.rows, 'Outlet savings').dest, 'N');
assert.strictEqual(by(p.rows, 'Outlet savings').markRead, !by(p.rows, 'Outlet savings').msg.isRead, 'unread newsletter gets marked read');
assert.strictEqual(by(p.rows, 'Your order SD123').markRead, !by(p.rows, 'Your order SD123').msg.isRead, 'rescued order still honours the flag');
assert.strictEqual(by(p.rows, 'Justin, make the Algarve').markRead, false, 'no flag on golfbreaks');

// which pane range reaches a given email (drives the 'outside the dates shown' hint)
const noon = new Date(2026, 8, 21, 12, 0);
assert.strictEqual(E.rangeContaining(new Date(2026, 8, 21, 0, 5), noon), 'today');
assert.strictEqual(E.rangeContaining(new Date(2026, 8, 20, 23, 59), noon), 'yesterday');
assert.strictEqual(E.rangeContaining(new Date(2026, 8, 15, 9, 0), noon), '7');
assert.strictEqual(E.rangeContaining(new Date(2026, 8, 14, 9, 0), noon), '30');
assert.strictEqual(E.rangeContaining(new Date(2026, 7, 23, 9, 0), noon), '30');
assert.strictEqual(E.rangeContaining(new Date(2026, 7, 22, 9, 0), noon), null);
assert.strictEqual(E.rangeContaining(new Date(2026, 8, 22, 9, 0), noon), null);
assert.strictEqual(E.rangeContaining('nonsense', noon), null);

console.log('engine tests passed:', p.rows.length, 'messages planned');
