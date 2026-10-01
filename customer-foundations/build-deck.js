const pptxgen = require('pptxgenjs');
const React = require('react');
const RDS = require('react-dom/server');
const sharp = require('sharp');
const Fi = require('react-icons/fi');

const OUT = process.argv[2] || 'deck.pptx';
const C = {
  bg: 'F6F7F9', ink: '1D1D1F', soft: '6E6E73', border: 'DCE1E7', white: 'FFFFFF',
  teal: '0B7A6F', tealSoft: 'E3F4F1', gold: '9A6410', goldSoft: 'FFF3D6',
  peach: 'FBEFE4', lav: 'EDEBFA', lavInk: '5B4BB7', red: 'B42318', redSoft: 'FDECEA',
};
const HEAD = 'SF Pro Display', BODY = 'SF Pro Text';
const FOOT = 'Homeaglow  ·  New Hire Care Training';

const iconCache = {};
async function icon(name, color) {
  const k = name + color;
  if (!iconCache[k]) {
    const svg = RDS.renderToStaticMarkup(React.createElement(Fi[name], { color: '#' + color, size: 256 }));
    iconCache[k] = 'image/png;base64,' + (await sharp(Buffer.from(svg)).png().toBuffer()).toString('base64');
  }
  return iconCache[k];
}

const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9'; // 10 x 5.625
pres.title = 'C CRM & Customer Foundations';

pres.defineSlideMaster({
  title: 'CONTENT',
  background: { color: C.bg },
  objects: [
    { text: { text: FOOT, options: { x: 0.48, y: 5.25, w: 4, h: 0.2, fontFace: BODY, fontSize: 7, color: C.soft, margin: 0 } } },
  ],
  slideNumber: { x: 9.0, y: 5.25, w: 0.5, h: 0.2, fontFace: BODY, fontSize: 7, color: C.soft, align: 'right' },
});
pres.defineSlideMaster({ title: 'COVER', background: { color: C.bg }, objects: [] });

const T = (s, text, o) => s.addText(text, Object.assign({ margin: 0, isTextBox: true, fontFace: BODY, color: C.ink, valign: 'top' }, o));
const box = (s, x, y, w, h, fill, line) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: line || fill, width: 0.75 }, rectRadius: 0.08 });
const badge = (s, x, y, n, fill, ink) => {
  s.addShape(pres.shapes.OVAL, { x, y, w: 0.3, h: 0.3, fill: { color: fill }, line: { color: fill } });
  T(s, String(n), { x, y, w: 0.3, h: 0.3, fontFace: HEAD, bold: true, fontSize: 9, color: ink, align: 'center', valign: 'middle' });
};
async function iconDot(s, x, y, name, d = 0.36, fill = C.tealSoft, ink = C.teal) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill } });
  const i = d * 0.52;
  s.addImage({ data: await icon(name, ink), x: x + (d - i) / 2, y: y + (d - i) / 2, w: i, h: i });
}

function content(eyebrow, title, sub, notes) {
  const s = pres.addSlide({ masterName: 'CONTENT' });
  T(s, eyebrow.toUpperCase(), { x: 0.48, y: 0.3, w: 9, h: 0.2, fontFace: HEAD, bold: true, fontSize: 8, color: C.teal, charSpacing: 0.5 });
  T(s, title, { x: 0.48, y: 0.5, w: 9, h: 0.5, fontFace: HEAD, bold: true, fontSize: 24 });
  if (sub) T(s, sub, { x: 0.48, y: 1.05, w: 9, h: 0.3, fontSize: 12, color: C.soft });
  if (notes) s.addNotes(notes);
  return s;
}

// Callout strip (yellow), as in the inspiration deck
async function tip(s, y, label, text, ico = 'FiInfo') {
  box(s, 0.45, y, 9.1, 0.42, C.goldSoft);
  s.addImage({ data: await icon(ico, C.gold), x: 0.62, y: y + 0.12, w: 0.18, h: 0.18 });
  T(s, [{ text: label + '  ', options: { bold: true, color: C.gold } }, { text }], { x: 0.9, y: y + 0.11, w: 8.5, h: 0.22, fontSize: 10 });
}

// Card with numbered badge or icon, title and one-line body
async function card(s, x, y, w, h, { n, ico, title, body, tint = C.tealSoft, ink = C.teal }) {
  box(s, x, y, w, h, C.white, C.border);
  if (ico) await iconDot(s, x + 0.18, y + 0.18, ico, 0.36, tint, ink); else badge(s, x + 0.18, y + 0.2, n, tint, ink);
  T(s, title, { x: x + 0.18, y: y + 0.66, w: w - 0.36, h: 0.3, fontFace: HEAD, bold: true, fontSize: 12 });
  if (body) T(s, body, { x: x + 0.18, y: y + 0.98, w: w - 0.36, h: h - 1.08, fontSize: 9.5, color: C.soft });
}
async function cards(s, y, h, items, opts = {}) {
  const gap = 0.15, n = items.length, w = (9.1 - gap * (n - 1)) / n;
  for (let i = 0; i < n; i++) await card(s, 0.45 + i * (w + gap), y, w, h, Object.assign({ n: i + 1 }, opts, items[i]));
}

// Topic break slide
let topicNo = 0;
const TOPICS = 6;
async function topic(title, sub, learn, ico, tint, ink, notes) {
  topicNo++;
  const s = pres.addSlide({ masterName: 'CONTENT' });
  T(s, `TOPIC ${String(topicNo).padStart(2, '0')} OF ${String(TOPICS).padStart(2, '0')}`, { x: 0.62, y: 1.32, w: 4.5, h: 0.2, fontFace: HEAD, bold: true, fontSize: 8, color: ink, charSpacing: 0.5 });
  T(s, title, { x: 0.62, y: 1.6, w: 4.8, h: 1.0, fontFace: HEAD, bold: true, fontSize: 28, valign: 'bottom' });
  T(s, sub, { x: 0.62, y: 2.7, w: 4.6, h: 0.5, fontSize: 11.5, color: C.soft });
  T(s, "YOU'LL LEARN", { x: 0.62, y: 3.28, w: 3, h: 0.2, fontFace: HEAD, bold: true, fontSize: 8, color: C.soft, charSpacing: 0.5 });
  for (let i = 0; i < learn.length; i++) {
    const y = 3.55 + i * 0.3;
    s.addShape(pres.shapes.OVAL, { x: 0.62, y: y + 0.02, w: 0.17, h: 0.17, fill: { color: tint }, line: { color: tint } });
    s.addImage({ data: await icon('FiCheck', ink), x: 0.655, y: y + 0.055, w: 0.1, h: 0.1 });
    T(s, learn[i], { x: 0.9, y, w: 4.3, h: 0.22, fontSize: 10.5 });
  }
  // Illustration panel: tinted card, white circle, icon. Swap the icon for an illustration later if wanted.
  box(s, 5.81, 0.68, 3.7, 4.26, tint);
  s.addShape(pres.shapes.OVAL, { x: 6.26, y: 1.2, w: 2.8, h: 2.8, fill: { color: C.white }, line: { color: C.white } });
  s.addImage({ data: await icon(ico, ink), x: 7.06, y: 2.0, w: 1.2, h: 1.2 });
  for (const [x, y] of [[6.04, 0.9], [8.96, 0.9], [8.96, 4.42]]) {
    s.addShape(pres.shapes.OVAL, { x, y, w: 0.36, h: 0.36, fill: { color: C.white }, line: { color: C.white } });
    s.addShape(pres.shapes.OVAL, { x: x + 0.135, y: y + 0.135, w: 0.09, h: 0.09, fill: { color: ink }, line: { color: ink } });
  }
  if (notes) s.addNotes(notes);
  return s;
}

// Trainer walkthrough slide: hand-off to the live system
async function walkthrough(eyebrow, system, lookFor, notes) {
  const s = content(eyebrow, 'Live walkthrough', null, notes);
  box(s, 0.45, 1.25, 9.1, 2.1, C.tealSoft);
  await iconDot(s, 0.8, 1.6, 'FiMonitor', 0.9, C.white, C.teal);
  T(s, 'TRAINER WALKTHROUGH', { x: 2.0, y: 1.6, w: 7, h: 0.2, fontFace: HEAD, bold: true, fontSize: 8, color: C.teal, charSpacing: 0.5 });
  T(s, `The trainer will open the live ${system} and walk you through the actual system.`, { x: 2.0, y: 1.85, w: 7.2, h: 0.75, fontFace: HEAD, bold: true, fontSize: 17 });
  T(s, 'Follow along and note your questions.', { x: 2.0, y: 2.7, w: 7, h: 0.25, fontSize: 11, color: C.soft });
  T(s, 'AS YOU WATCH, LOOK FOR', { x: 0.48, y: 3.6, w: 5, h: 0.2, fontFace: HEAD, bold: true, fontSize: 8, color: C.soft, charSpacing: 0.5 });
  const gap = 0.15, w = (9.1 - gap * (lookFor.length - 1)) / lookFor.length;
  for (let i = 0; i < lookFor.length; i++) {
    const x = 0.45 + i * (w + gap);
    box(s, x, 3.88, w, 0.62, C.white, C.border);
    await iconDot(s, x + 0.14, 4.02, lookFor[i][0], 0.34);
    T(s, lookFor[i][1], { x: x + 0.58, y: 3.88, w: w - 0.68, h: 0.62, fontSize: 10, bold: true, valign: 'middle' });
  }
  return s;
}

(async () => {
  // ---------- Cover ----------
  {
    const s = pres.addSlide({ masterName: 'COVER' });
    T(s, 'HOMEAGLOW  ·  NEW HIRE CARE TRAINING', { x: 0.62, y: 0.6, w: 5, h: 0.2, fontFace: HEAD, bold: true, fontSize: 8, color: C.teal, charSpacing: 0.5 });
    T(s, 'Customer Foundations\n& the C CRM', { x: 0.62, y: 1.0, w: 4.9, h: 1.6, fontFace: HEAD, bold: true, fontSize: 32 });
    T(s, 'The rules we never bend, the path every customer takes, and the tools we use to help them.', { x: 0.62, y: 2.7, w: 4.6, h: 0.7, fontSize: 12.5, color: C.soft });
    T(s, 'Internal training material', { x: 0.62, y: 5.0, w: 3, h: 0.2, fontSize: 7, color: C.soft });
    box(s, 5.81, 0.4, 3.7, 4.82, C.peach);
    s.addShape(pres.shapes.OVAL, { x: 6.21, y: 1.36, w: 2.9, h: 2.9, fill: { color: C.white }, line: { color: C.white } });
    s.addImage({ data: await icon('FiHeadphones', C.gold), x: 7.06, y: 2.21, w: 1.2, h: 1.2 });
    s.addNotes('Welcome trainees. This module covers six topics: Zero Tolerance Policy, the Customer Journey, the Legacy C CRM, the C Dashboard, the New C CRM, and a hands-on CRM activity.');
  }

  // ---------- Agenda ----------
  {
    const s = content('Overview', 'Agenda', 'Six topics. The last four happen mostly in the live systems.',
      'Walk through the agenda. Point out that topics 3 to 5 are live demos: the slides only set up what trainees will see.');
    const rows = [
      ['01', 'Zero Tolerance Policy', 'The non-negotiables and their penalties'],
      ['02', 'Customer Journey', 'The five steps every customer moves through'],
      ['03', 'Legacy C CRM', 'Our central hub, and a live walkthrough'],
      ['04', 'C Dashboard', 'What the customer sees and manages'],
      ['05', 'New C CRM', 'Where we respond to customers'],
      ['06', 'C CRM Activity', 'Hands-on practice, set by your trainer'],
    ];
    const hdr = ['#', 'Topic', 'You will learn'].map(t => ({ text: t, options: { bold: true, color: C.teal, fill: { color: C.tealSoft }, fontFace: HEAD } }));
    const body = rows.map(r => r.map((t, i) => ({ text: t, options: { color: i === 0 ? C.teal : i === 1 ? C.ink : C.soft, bold: i < 2, fill: { color: C.white } } })));
    s.addTable([hdr, ...body], { x: 0.45, y: 1.5, w: 9.1, colW: [0.7, 3.0, 5.4], rowH: 0.42, fontFace: BODY, fontSize: 10.5, valign: 'middle', border: { type: 'solid', pt: 0.75, color: C.border }, margin: [0, 0.14, 0, 0.14] });
  }

  // ================= 1. ZERO TOLERANCE POLICY =================
  await topic('Zero Tolerance Policy', 'The non-negotiables: behaviors that are never acceptable with customers or cleaners.',
    ['What ZTP is and why it exists', 'Violations by level', 'Penalties and other conditions', 'The rules, area by area'],
    'FiShield', C.redSoft, C.red,
    'ZTP is different from the rest of the Knowledge Library. Most articles are about judgment; this one is about hard lines that create legal, security, compliance or trust risk.');

  {
    const s = content('Zero Tolerance Policy', 'Not judgment calls. Hard lines.', 'Violations carry consequences regardless of intent.',
      'Most KB articles teach judgment and best practice. ZTP is the handful of things that, done wrong, are not just a bad customer experience: they create legal, security or compliance exposure for Homeaglow. Knowing it protects the agent\'s QA score and standing, and keeps every interaction safe and trustworthy. Rule of thumb: if you are ever unsure whether an action crosses a line, stop and ask a Team Lead first.');
    await cards(s, 1.6, 1.75, [
      { ico: 'FiFileText', title: 'Legal', body: 'Wording, auto-renewal and state laws we must follow.' },
      { ico: 'FiLock', title: 'Security', body: 'Personal data, logins and payment details stay protected.' },
      { ico: 'FiHeart', title: 'Trust', body: 'Customers and cleaners can rely on how we treat them.' },
    ], { tint: C.redSoft, ink: C.red });
    await tip(s, 3.6, 'Unsure?', 'Stop and ask a Team Lead before you act. Asking is always better than a violation.', 'FiHelpCircle');
  }

  {
    const s = content('Zero Tolerance Policy', 'Violations at a glance', 'Every violation below is an auto-fail on QA, except capturing work screens.',
      'Full list from the KB, by level. L3: Security violation; Capturing or sharing work screens (no QA score). L2: Not removing card info sent to CS; Ticket avoidance; Incorrect handling of OCC-ARL accounts; Non-compliance with legal wording; Sending internal notes as email/SMS; Disclosing internal information; Improper use of AI tools; Not unsubscribing customers who opt out; Mishandling Washington State cancellations. L1: Rude/argumentative comms; Not following the bug reporting process; Editing Django inappropriately; Handling a Thumbtack account.');
    const cols = [
      ['Level 3', 'Most severe', C.redSoft, C.red, ['Security violation', 'Capturing or sharing work screens']],
      ['Level 2', '', C.goldSoft, C.gold, ['Card info not removed', 'Ticket avoidance', 'OCC / ARL accounts', 'Legal wording', 'Internal notes sent out', 'Disclosing internal info', 'Improper AI tool use', 'Not unsubscribing', 'WA cancellations']],
      ['Level 1', '', C.tealSoft, C.teal, ['Rude / argumentative comms', 'Bug reporting process', 'Editing Django', 'Thumbtack accounts']],
    ];
    const gap = 0.15, w = (9.1 - 2 * gap) / 3;
    cols.forEach(([lvl, , tint, ink, items], i) => {
      const x = 0.45 + i * (w + gap);
      box(s, x, 1.5, w, 3.3, C.white, C.border);
      box(s, x + 0.15, 1.65, 0.75, 0.28, tint);
      T(s, lvl, { x: x + 0.15, y: 1.65, w: 0.75, h: 0.28, fontFace: HEAD, bold: true, fontSize: 9, color: ink, align: 'center', valign: 'middle' });
      const isL2 = items.length > 5;
      T(s, items.map((t, j) => ({ text: t, options: { bullet: { indent: 10 }, breakLine: j < items.length - 1 } })),
        { x: x + 0.15, y: 2.08, w: w - 0.3, h: 2.8, fontSize: isL2 ? 9.5 : 10.5, color: C.ink, paraSpaceAfter: isL2 ? 3 : 6 });
    });
  }

  {
    const s = content('Zero Tolerance Policy', 'Penalties', 'Penalties grow with the level and with each repeat.',
      'Other conditions: 3 ZTPs of any level means termination. An L1 followed by an L2 means a 1-week suspension. Previously suspended for an L1 and then commits an L2 means termination. Multiple ZTPs of different levels in the same week: the penalty follows the higher level. L1 clearing period is 6 months. For Incorrect Handling of OCC accounts, the QA score is set at a minimum of -$25, not stacked on the LTV computation; if the standard markdown is higher, that is used.');
    const H = (t, c) => ({ text: t, options: { bold: true, color: c || C.ink, fill: { color: C.white }, fontFace: HEAD, align: 'center' } });
    const cell = (t, tint, ink) => ({ text: t, options: { color: t === '—' ? C.soft : ink, fill: { color: t === '—' ? C.white : tint }, align: 'center', bold: t.startsWith('Subject') } });
    const rows = [
      [H(''), H('Level 1', C.teal), H('Level 2', C.gold), H('Level 3', C.red)],
      [H('1st instance'), cell('Final warning (NTE)', C.tealSoft, C.ink), cell('Suspension, up to 5 days', C.goldSoft, C.ink), cell('Subject to termination', C.redSoft, C.red)],
      [H('2nd instance'), cell('Suspension, up to 3 days', C.tealSoft, C.ink), cell('Subject to termination', C.redSoft, C.red), cell('—')],
      [H('3rd instance'), cell('Subject to termination', C.redSoft, C.red), cell('—'), cell('—')],
    ];
    s.addTable(rows, { x: 0.45, y: 1.5, w: 9.1, colW: [1.6, 2.5, 2.5, 2.5], rowH: 0.55, fontFace: BODY, fontSize: 10.5, valign: 'middle', border: { type: 'solid', pt: 0.75, color: C.border } });
    await tip(s, 3.95, 'Remember:', '3 ZTPs of any level means termination. L1s clear after 6 months.', 'FiAlertTriangle');
  }

  {
    const s = content('Security & compliance', 'Protect what is private', null,
      'Security violations: sharing sensitive info or system access incorrectly, e.g. an auto-login link to the wrong person, sharing credentials, mishandling documents with personal data (payment info, IDs, tax forms). Always double-check recipients. Card info: if a C or CP sends card details, a Team Lead must remove it from the CRM immediately (PCI). Never ask for card info or photos, never put card details in internal notes; direct the customer to the CC update page on their dashboard. Internal notes are for risk assessments, offers/action plans and follow-ups; never send them as email or SMS. Never disclose: CP DTC flags, CP Swap details, CP Tiering levels, known fraud/safety flags, or features not yet released on NCW. Test: if it is not in the Help Center or public comms, it is internal.');
    await cards(s, 1.4, 1.95, [
      { ico: 'FiKey', title: 'Security', body: 'Double-check every recipient before sending a login link or personal data.' },
      { ico: 'FiCreditCard', title: 'Card details', body: 'Never accept them. A Team Lead removes any that arrive.' },
      { ico: 'FiEyeOff', title: 'Internal notes', body: 'Never sent as an email or SMS.' },
      { ico: 'FiSlash', title: 'Internal info', body: 'Not in the Help Center? Then it stays internal.' },
    ], { tint: C.redSoft, ink: C.red });
    await tip(s, 3.6, 'Quick test:', '"If it\'s not in the Help Center or public-facing communication, it\'s internal."', 'FiHelpCircle');
  }

  {
    const s = content('Security & compliance', 'Say it the legal way', 'Cleaners are independent contractors, not Homeaglow employees. Our words must show it.',
      'Legal Wording Compliance. CPs are Independent Contractors. Referring to Homeaglow itself is fine: "your Homeaglow cleaning" and "your cleaning on Homeaglow" are both OK. Ask trainees to rewrite a few sentences out loud.');
    const pairs = [
      ['Your cleaner on Homeaglow', 'Your Homeaglow cleaner'],
      ['Find a cleaner on Homeaglow', 'Find a Homeaglow cleaner'],
      ['Your cleaner / Your cleaner Alice', 'Our cleaner'],
      ['Your cleaner will see you at <address>', 'We will see you at <address>'],
      ['Your cleaner was unable to', 'We were unable to'],
      ['Your requested cleaner is still reviewing', 'We are still waiting for your requested cleaner'],
    ];
    const head = [
      { text: 'Do say', options: { bold: true, color: C.teal, fill: { color: C.tealSoft }, fontFace: HEAD } },
      { text: "Don't say", options: { bold: true, color: C.red, fill: { color: C.redSoft }, fontFace: HEAD } },
    ];
    const body = pairs.map(([a, b]) => [{ text: a, options: { color: C.ink, fill: { color: C.white } } }, { text: b, options: { color: C.soft, fill: { color: C.white } } }]);
    s.addTable([head, ...body], { x: 0.45, y: 1.55, w: 9.1, colW: [4.55, 4.55], rowH: 0.42, fontFace: BODY, fontSize: 10.5, valign: 'middle', border: { type: 'solid', pt: 0.75, color: C.border }, margin: [0, 0.16, 0, 0.16] });
  }

  {
    const s = content('Security & compliance', 'Work screens stay on work screens', 'No exceptions, even if it looks blurred or empty.',
      'Anything on a work screen can include names, addresses, payment details, internal notes, flags, other tabs or pop-up notifications. Once a capture exists outside our systems we cannot delete it, track it or undo it. "I blurred it" or "nothing was on screen" does not count: the risk is the capture itself. It is a ZTP when an agent takes a photo/video of a work screen with any device, takes a screenshot or screen recording, or shares any of these anywhere, including personal chats, social media or coworker group chats.');
    await cards(s, 1.6, 1.75, [
      { ico: 'FiCamera', title: 'No photos or videos', body: 'With a phone or any other device.' },
      { ico: 'FiMaximize', title: 'No screenshots', body: 'And no screen recordings.' },
      { ico: 'FiShare2', title: 'No sharing', body: 'Not in personal chats, social media or coworker group chats.' },
    ], { tint: C.redSoft, ink: C.red });
    await tip(s, 3.6, 'Why:', 'Once a capture leaves our systems, we can\'t delete it, track it or undo it.', 'FiAlertTriangle');
  }

  {
    const s = content('Customer handling', 'Handle these customers by the rules', null,
      'OCC-ARL: some states have Auto-Renewal Laws requiring immediate cancellation on request with no extra steps or delays. Failing to cancel an FC when required is the ZTP; offering retention when the correct call was not to cancel is a markdown, not a ZTP. When in doubt, check the OCC Non-Retention Process or a Team Lead. Unsubscribe: if a customer says they do not want contact or replies STOP/UNSUBSCRIBE, unsubscribe them, confirm it, and handle other pending items as BAU. Washington State: WA customers are exempt from MCT and the ETF (a standing legal exemption). Process the cancellation immediately: no ETF, no citing MCT, no retention offers first, no presenting ETF/MCT as if they apply. Tone: avoid condescending ("You didn\'t understand...") and passive-aggressive ("As I\'ve told you before...") language. Educate, personalize, focus on resolution; don\'t argue, use generic statements, or over-empathize.');
    await cards(s, 1.4, 2.0, [
      { ico: 'FiRefreshCcw', title: 'OCC / ARL accounts', body: 'Auto-renewal law states: cancel right away when asked.' },
      { ico: 'FiBellOff', title: 'Unsubscribe requests', body: '"STOP" means unsubscribe them and confirm it.' },
      { ico: 'FiMapPin', title: 'Washington State', body: 'Cancel right away. No ETF, no MCT, no retention offers.' },
      { ico: 'FiMessageCircle', title: 'Tone', body: 'Educate, never argue. No condescending or passive-aggressive replies.' },
    ], { tint: C.redSoft, ink: C.red });
    await tip(s, 3.65, 'In doubt?', 'Check the OCC Non-Retention Process or ask a Team Lead before replying.', 'FiHelpCircle');
  }

  {
    const s = content('Operational compliance', 'Follow the process, every time', null,
      'Ticket avoidance: deliberately skipping a ticket that needs action, e.g. using CRM/NavBar filters to sidestep hard tickets, clearing without resolving, leaving reassigned tickets unattended, skipping charge disputes. Clearing a "thanks / got it" ticket is NOT avoidance. Unsure? Ask in the designated Slack channels. Bug reporting: file a report via the Bug Validation Form when an issue can\'t be fixed by a dashboard/CRM update; check the Confirmed Bugs Tracker first; don\'t add sample or unrelated comments on escalated Trello cards. Django: bypasses system codes, so only specific updates are CS-editable (see KB list); escalate review edits in NCW and job history edits for LMCs/LOs to SMEs/Managers. When in doubt, don\'t edit, escalate. Thumbtack: identified by a system flag in New CRM, an internal note, or the customer saying so. Transfer to the Thumbtack queue, notify an available Aftersales agent, and stop working the ticket. Transferring without reaching anyone is still compliant.');
    await cards(s, 1.4, 2.0, [
      { ico: 'FiInbox', title: 'Ticket avoidance', body: 'Work every ticket assigned to you, including the hard ones.' },
      { ico: 'FiAlertOctagon', title: 'Bug reporting', body: 'Use the Bug Validation Form. Check the tracker first.' },
      { ico: 'FiDatabase', title: 'Editing Django', body: 'Only approved edits. When in doubt, escalate.' },
      { ico: 'FiRepeat', title: 'Thumbtack accounts', body: 'Transfer to the Thumbtack queue, then stop.' },
    ], { tint: C.redSoft, ink: C.red });
    await tip(s, 3.65, 'Remember:', 'Taking a few extra minutes to check is always better than creating a risk.', 'FiShield');
  }

  // ================= 2. CUSTOMER JOURNEY =================
  await topic('Customer Journey', 'The map of a customer\'s whole experience, from first hearing about us to their cleaning.',
    ['The five steps', 'How customers find us', 'Three ways to sign up', 'Where tickets come from'],
    'FiMap', C.goldSoft, C.gold,
    'Every ticket sits somewhere on this journey. Knowing where an issue sits lets you anticipate needs and pick the right playbook.');

  {
    const s = content('Customer Journey', 'Five steps, one journey', 'Every ticket you handle sits somewhere on this path.',
      'Understanding the full shape lets you anticipate what a customer needs before they fully explain it, catch friction before it becomes a complaint, and know where an issue sits. A booking problem needs a very different response from a post-cleaning service issue, even when the tone sounds the same. This page is the map; playbooks like PCQ, Overcharged Hours and Retention are what you use once you know where you are.');
    const steps = [['FiSearch', 'Learns about us'], ['FiUserPlus', 'Signs up'], ['FiCalendar', 'Books'], ['FiSliders', 'Manages the appointment'], ['FiStar', 'Receives and pays']];
    const w = 1.62, gap = (9.1 - 5 * w) / 4;
    for (let i = 0; i < 5; i++) {
      const x = 0.45 + i * (w + gap);
      box(s, x, 1.7, w, 2.0, C.white, C.border);
      await iconDot(s, x + (w - 0.6) / 2, 1.95, steps[i][0], 0.6, C.goldSoft, C.gold);
      T(s, `STEP ${i + 1}`, { x, y: 2.72, w, h: 0.2, fontFace: HEAD, bold: true, fontSize: 8, color: C.gold, align: 'center', charSpacing: 0.5 });
      T(s, steps[i][1], { x: x + 0.1, y: 2.95, w: w - 0.2, h: 0.55, fontFace: HEAD, bold: true, fontSize: 11.5, align: 'center' });
      if (i < 4) s.addImage({ data: await icon('FiChevronRight', C.soft), x: x + w + gap / 2 - 0.09, y: 2.6, w: 0.18, h: 0.18 });
    }
    await tip(s, 4.1, 'Why it matters:', 'Know where the customer is, and you\'ll know which playbook to use.', 'FiCompass');
  }

  {
    const s = content('Step 1  ·  Learns about us', 'How customers find us', null,
      'Customers find us through our website (homeaglow.com) or a sister brand (Dazzling Cleaning, Cozy Maid, Bubbly Cleaning, Dapper Maids); search engines, the most common source; TV commercials; friends and family referrals; and social media (Facebook, YouTube, TikTok, etc).');
    const items = [['FiSearch', 'Search engines', 'The most common source'], ['FiGlobe', 'Our websites', 'Homeaglow and sister brands'], ['FiTv', 'TV commercials', ''], ['FiUsers', 'Friends & family', 'Referrals'], ['FiThumbsUp', 'Social media', 'Facebook, YouTube, TikTok']];
    const w = 1.7, gap = (9.1 - 5 * w) / 4;
    for (let i = 0; i < 5; i++) {
      const x = 0.45 + i * (w + gap);
      box(s, x, 1.4, w, 2.3, C.white, C.border);
      await iconDot(s, x + (w - 0.7) / 2, 1.7, items[i][0], 0.7, C.goldSoft, C.gold);
      T(s, items[i][1], { x: x + 0.1, y: 2.6, w: w - 0.2, h: 0.3, fontFace: HEAD, bold: true, fontSize: 11.5, align: 'center' });
      T(s, items[i][2], { x: x + 0.1, y: 2.92, w: w - 0.2, h: 0.5, fontSize: 9, color: C.soft, align: 'center' });
    }
    await tip(s, 3.95, 'Most common:', 'Search engines.', 'FiTrendingUp');
  }

  {
    const s = content('Step 2  ·  Signs up', 'Three ways to sign up', 'The path a customer took shapes their billing and what options they have.',
      'Legacy DHJ: the customer buys a discounted voucher from <brand>.com/deal. First Cleaning Free (FCF): signs up for a membership directly, pays $19 upfront (more for a longer first cleaning), membership starts immediately. From August 3, 2026 all new sign-ups are FCF only, so DHJ is a closed, shrinking group. Non-member: third-party merchant vouchers (Groupon, Living Social, Gilt City) never converted; Homeaglow Gift Cards (never trigger FC); one-time cleanings sold by Inside Sales (paid upfront, full price); trial cleanings sold by Inside Sales (regular price, 30 days to cancel the membership without obligation). Non-members are billed differently and have nothing to "cancel"; the question is usually just whether a specific charge should be refunded.');
    await cards(s, 1.6, 1.75, [
      { n: 1, title: 'Legacy DHJ', body: 'Bought a discounted voucher from <brand>.com/deal.' },
      { n: 2, title: 'First Cleaning Free', body: 'Joins the membership directly. Pays $19 upfront.' },
      { n: 3, title: 'Non-member', body: 'Groupon, gift card, one-time or trial cleaning. No membership.' },
    ], { tint: C.goldSoft, ink: C.gold });
    await tip(s, 3.6, 'Since Aug 3, 2026:', 'all new membership sign-ups are First Cleaning Free.', 'FiCalendar');
  }

  {
    const s = content('Steps 3 & 4  ·  Books and manages', 'Booking and managing', null,
      'Booking: customers with a voucher or membership are prompted to opt into a recurring schedule and add alternate start times; both increase the odds of a successful match. Customers without one are not required to. Managing: through their dashboard, up to 6 hours before the start time, a customer can reschedule, cancel, or match with a different cleaner.');
    // Left: booking
    box(s, 0.45, 1.3, 4.45, 3.3, C.white, C.border);
    await iconDot(s, 0.65, 1.5, 'FiCalendar', 0.4, C.goldSoft, C.gold);
    T(s, 'Books', { x: 0.65, y: 2.02, w: 4, h: 0.3, fontFace: HEAD, bold: true, fontSize: 14 });
    T(s, 'Members are prompted to add:', { x: 0.65, y: 2.38, w: 4, h: 0.25, fontSize: 10, color: C.soft });
    for (const [i, t] of ['A recurring schedule', 'Alternate start times'].entries()) {
      box(s, 0.65, 2.75 + i * 0.5, 4.05, 0.4, C.goldSoft);
      T(s, t, { x: 0.82, y: 2.75 + i * 0.5, w: 3.8, h: 0.4, fontSize: 10.5, bold: true, valign: 'middle' });
    }
    T(s, 'Both raise the odds of a match.', { x: 0.65, y: 3.85, w: 4, h: 0.25, fontSize: 9.5, color: C.soft, italic: true });
    // Right: manage
    box(s, 5.1, 1.3, 4.45, 3.3, C.white, C.border);
    T(s, '6', { x: 5.3, y: 1.4, w: 0.7, h: 0.9, fontFace: HEAD, bold: true, fontSize: 48, color: C.gold });
    T(s, 'hours before start', { x: 5.95, y: 1.75, w: 3.4, h: 0.3, fontFace: HEAD, bold: true, fontSize: 12, color: C.gold });
    T(s, 'Up until then, from their dashboard, customers can:', { x: 5.3, y: 2.38, w: 4.1, h: 0.25, fontSize: 10, color: C.soft });
    const acts = [['FiClock', 'Reschedule'], ['FiXCircle', 'Cancel'], ['FiUserCheck', 'Match a new cleaner']];
    for (let i = 0; i < 3; i++) {
      const x = 5.3 + i * 1.37;
      box(s, x, 2.75, 1.27, 1.1, C.goldSoft);
      s.addImage({ data: await icon(acts[i][0], C.gold), x: x + 0.49, y: 2.88, w: 0.28, h: 0.28 });
      T(s, acts[i][1], { x: x + 0.05, y: 3.25, w: 1.17, h: 0.5, fontSize: 9.5, bold: true, align: 'center' });
    }
  }

  {
    const s = content('Step 5  ·  Receives and pays', 'Where service issues start', 'The booking finalizes here, and most Service Issues tickets begin here too.',
      'Common issues after the cleaning: Poor Cleaning Quality (didn\'t meet expectations), False Invoice (charged when no cleaning happened), Overcharged Hours (charged for more time than the CP worked), Theft/Damage (missing or damaged item). Each has its own playbook under Service Issues; this page only shows where they surface.');
    await cards(s, 1.6, 1.75, [
      { ico: 'FiFrown', title: 'Poor cleaning quality', body: 'The cleaning didn\'t meet expectations.' },
      { ico: 'FiFileMinus', title: 'False invoice', body: 'Charged when no cleaning happened.' },
      { ico: 'FiClock', title: 'Overcharged hours', body: 'Charged for more time than was worked.' },
      { ico: 'FiPackage', title: 'Theft / damage', body: 'An item missing or damaged after the cleaning.' },
    ], { tint: C.goldSoft, ink: C.gold });
    await tip(s, 3.6, 'Next:', 'each issue has its own playbook in the Knowledge Library.', 'FiBookOpen');
  }

  // ================= 3. LEGACY C CRM =================
  await topic('Introduction to C CRM (Legacy)', 'Our central hub for customers, cleaners, jobs and payments.',
    ['What the CRM is for', 'Why we have two CRMs', 'A live walkthrough'],
    'FiDatabase', C.tealSoft, C.teal,
    'We currently have two CRM systems. We start with the Legacy C CRM because many actions still live there.');

  {
    const s = content('Legacy C CRM', 'One hub for everything', 'Where we manage accounts, talk to users and find what we need to help them.',
      'The CRM is the central hub for managing customer and cleaner accounts, jobs and payments. It is also where we communicate with users, manage tickets, see our daily targets, and access the information that helps us handle users\' concerns. Note: we currently have two CRM systems. We introduce the Legacy CRM first.');
    const tiles = [['FiUsers', 'Customer & cleaner accounts'], ['FiBriefcase', 'Jobs'], ['FiCreditCard', 'Payments'], ['FiMessageSquare', 'Messages with users'], ['FiInbox', 'Tickets'], ['FiTarget', 'Daily targets']];
    const w = (9.1 - 0.3) / 3, h = 0.95;
    for (let i = 0; i < 6; i++) {
      const x = 0.45 + (i % 3) * (w + 0.15), y = 1.5 + Math.floor(i / 3) * (h + 0.15);
      box(s, x, y, w, h, C.white, C.border);
      await iconDot(s, x + 0.2, y + (h - 0.5) / 2, tiles[i][0], 0.5);
      T(s, tiles[i][1], { x: x + 0.85, y, w: w - 1.0, h, fontFace: HEAD, bold: true, fontSize: 12, valign: 'middle' });
    }
    await tip(s, 3.85, 'Two CRMs:', 'we have a Legacy CRM and a New CRM. We start with Legacy.', 'FiLayers');
  }

  await walkthrough('Legacy C CRM', 'Legacy C CRM', [['FiUser', 'Account details'], ['FiBriefcase', 'Jobs and charges'], ['FiMessageSquare', 'Messages and tickets'], ['FiTarget', 'Daily targets']],
    'Switch to the live Legacy C CRM now. Do not demo from the slides. Suggested flow: open a customer account, show account details and flags, job history and charges, the message/ticket thread, and where daily targets show. Adjust to the cohort.');

  // ================= 4. C DASHBOARD =================
  await topic('Introduction to C Dashboard', 'Where customers manage their account and bookings.',
    ['What the customer sees', 'What they can do themselves', 'A live walkthrough'],
    'FiLayout', C.lav, C.lavInk,
    'The C Dashboard is the customer-facing side. Seeing it helps agents guide customers and understand what they already tried.');

  {
    const s = content('C Dashboard', 'The customer\'s side', 'Where customers manage their account and bookings.',
      'The Customer dashboard is where customers manage their account and bookings: book a cleaning, see upcoming and past cleanings and alerts (e.g. a cancelled cleaning), and, up to 6 hours before start, reschedule, cancel or match with a different cleaner. Knowing what they see lets agents point customers to self-serve steps and understand what they have already tried.');
    await cards(s, 1.6, 1.75, [
      { ico: 'FiPlusCircle', title: 'Book a cleaning', body: 'Start a new booking.' },
      { ico: 'FiCalendar', title: 'Upcoming cleanings', body: 'Reschedule, cancel or match a new cleaner.' },
      { ico: 'FiClock', title: 'Past cleanings', body: 'Cleanings and the cleaners who did them.' },
      { ico: 'FiUser', title: 'Account', body: 'Their profile and account settings.' },
    ], { tint: C.lav, ink: C.lavInk });
    await tip(s, 3.6, 'Why it helps:', 'when you know what the customer sees, you can guide them step by step.', 'FiEye');
  }

  await walkthrough('C Dashboard', 'C Dashboard', [['FiHome', 'The home view'], ['FiCalendar', 'Upcoming cleanings'], ['FiClock', 'Past cleanings'], ['FiSettings', 'Account settings']],
    'Switch to the live C Dashboard now. Suggested flow: home view and alerts, booking a cleaning, managing an upcoming cleaning (reschedule, cancel, match), past cleanings, and account settings.');

  // ================= 5. NEW C CRM =================
  await topic('Introduction to New C CRM', 'Where we respond to customers today.',
    ['What the New CRM is for', 'When you still need Legacy', 'A live walkthrough'],
    'FiMessageSquare', C.tealSoft, C.teal,
    'The New C CRM is where agents respond to customers. Some actions are still only available in the Legacy CRM.');

  {
    const s = content('New C CRM', 'Two CRMs, one workflow', null,
      'Important: we respond to customers through the New CRM. At this time, there are still some actions that can only be done in the Legacy CRM, so agents use both. The trainer will point out which actions still require Legacy during the walkthrough.');
    const col = async (x, tint, ink, ico, eyebrow, title, body) => {
      box(s, x, 1.3, 4.1, 2.55, C.white, C.border);
      await iconDot(s, x + 0.25, 1.55, ico, 0.6, tint, ink);
      T(s, eyebrow, { x: x + 0.25, y: 2.35, w: 3.6, h: 0.2, fontFace: HEAD, bold: true, fontSize: 8, color: ink, charSpacing: 0.5 });
      T(s, title, { x: x + 0.25, y: 2.6, w: 3.6, h: 0.4, fontFace: HEAD, bold: true, fontSize: 16 });
      T(s, body, { x: x + 0.25, y: 3.05, w: 3.6, h: 0.6, fontSize: 10.5, color: C.soft });
    };
    await col(0.45, C.tealSoft, C.teal, 'FiMessageSquare', 'NEW C CRM', 'Respond to customers', 'Every reply to a customer goes out from here.');
    s.addImage({ data: await icon('FiRepeat', C.soft), x: 4.8, y: 2.4, w: 0.4, h: 0.4 });
    await col(5.45, C.goldSoft, C.gold, 'FiDatabase', 'LEGACY C CRM', 'Some actions still live here', 'Use it for actions the New CRM can\'t do yet.');
    await tip(s, 4.1, 'Important:', 'reply in the New CRM. Switch to Legacy only for actions the New CRM can\'t do yet.', 'FiAlertCircle');
  }

  await walkthrough('New C CRM', 'New C CRM', [['FiInbox', 'Ticket and status'], ['FiUser', 'Customer info'], ['FiMessageSquare', 'Replying to customers'], ['FiLayers', 'What still needs Legacy']],
    'Switch to the live New C CRM now. Suggested flow: ticket header (status, assignee, category, SLA), customer information panel, related jobs, the message thread and internal notes, composing a reply with macros. Call out which actions still require the Legacy CRM.');

  // ================= 6. C CRM ACTIVITY =================
  await topic('C CRM Activity', 'Put what you saw into practice.',
    ['Work in the live CRM', 'Practice real scenarios', 'Ask questions as you go'],
    'FiTarget', C.peach, C.gold,
    'Trainer decides the activity format, scenarios, and level of difficulty based on trainee readiness.');

  {
    const s = content('C CRM Activity', 'Activity setup', 'Your trainer will set the activity for this cohort.',
      'TRAINER NOTE: Trainer decides the activity format, scenarios, and level of difficulty based on trainee readiness. Fill in the three boxes on this slide before the session (or live with the class).');
    const fields = [['FiGrid', 'Format', 'Solo, pairs or groups'], ['FiFileText', 'Scenarios', 'Which cases to work'], ['FiBarChart2', 'Difficulty', 'Based on readiness']];
    const gap = 0.15, w = (9.1 - 2 * gap) / 3;
    for (let i = 0; i < 3; i++) {
      const x = 0.45 + i * (w + gap);
      box(s, x, 1.5, w, 2.75, C.white, C.border);
      await iconDot(s, x + 0.2, 1.7, fields[i][0], 0.42, C.peach, C.gold);
      T(s, fields[i][1], { x: x + 0.2, y: 2.25, w: w - 0.4, h: 0.3, fontFace: HEAD, bold: true, fontSize: 13 });
      T(s, fields[i][2], { x: x + 0.2, y: 2.57, w: w - 0.4, h: 0.25, fontSize: 9.5, color: C.soft });
      // Editable blank for the trainer to fill in
      box(s, x + 0.2, 2.95, w - 0.4, 1.1, C.bg, C.border);
      T(s, 'Click to add', { x: x + 0.32, y: 2.95, w: w - 0.64, h: 1.1, fontSize: 9.5, color: C.soft, italic: true, valign: 'middle' });
    }
    await tip(s, 4.45, 'Trainer note:', 'trainer decides the format, scenarios and difficulty based on trainee readiness.', 'FiEdit3');
  }

  // ---------- Close ----------
  {
    const s = pres.addSlide({ masterName: 'CONTENT' });
    T(s, 'WRAP-UP', { x: 0.62, y: 1.9, w: 5, h: 0.2, fontFace: HEAD, bold: true, fontSize: 8, color: C.teal, charSpacing: 0.5 });
    T(s, 'Questions?', { x: 0.62, y: 2.15, w: 5, h: 0.75, fontFace: HEAD, bold: true, fontSize: 34 });
    T(s, 'When in doubt, check the Knowledge Library or ask your Trainer or Team Lead.', { x: 0.62, y: 2.95, w: 4.6, h: 0.5, fontSize: 12, color: C.soft });
    box(s, 5.81, 0.68, 3.7, 4.26, C.tealSoft);
    s.addShape(pres.shapes.OVAL, { x: 6.26, y: 1.2, w: 2.8, h: 2.8, fill: { color: C.white }, line: { color: C.white } });
    s.addImage({ data: await icon('FiHelpCircle', C.teal), x: 7.06, y: 2.0, w: 1.2, h: 1.2 });
    s.addNotes('Open the floor for questions. Recap: ZTP hard lines, the five-step customer journey, and the three systems: Legacy C CRM, C Dashboard and New C CRM.');
  }

  await pres.writeFile({ fileName: OUT });
  console.log('wrote', OUT);
})();
