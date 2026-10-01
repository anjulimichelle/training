// Builds the "Customer Foundations & C CRM" training deck.
// Usage: node deck.js out.pptx [imageDir]
const path = require('path');
const pptxgen = require('pptxgenjs');
const React = require('react');
const RDS = require('react-dom/server');
const sharp = require('sharp');
const Fi = require('react-icons/fi');

const OUT = process.argv[2] || 'deck.pptx';
const IMG = process.argv[3] || path.join(__dirname, 'img');

// One palette, taken from the Care training deck: teal accents on a light grey page,
// white cards, and a soft yellow strip for tips.
const C = {
  bg: 'F6F7F9', ink: '1D1D1F', soft: '6E6E73', border: 'DCE1E7', white: 'FFFFFF',
  teal: '0B7A6F', tealSoft: 'E3F4F1', gold: '9A6410', goldSoft: 'FFF3D6',
};
const HEAD = 'SF Pro Display', BODY = 'SF Pro Text';
const FOOT = 'Homeaglow  ·  New Hire Care Training';

const iconCache = {};
async function icon(name, color = C.teal) {
  const k = name + color;
  if (!iconCache[k]) {
    const svg = RDS.renderToStaticMarkup(React.createElement(Fi[name], { color: '#' + color, size: 256 }));
    iconCache[k] = 'image/png;base64,' + (await sharp(Buffer.from(svg)).png().toBuffer()).toString('base64');
  }
  return iconCache[k];
}

// Place an image file inside a box, keeping its aspect ratio, centered.
async function fitImage(s, file, x, y, w, h) {
  const p = path.join(IMG, file);
  const { width, height } = await sharp(p).metadata();
  const r = Math.min(w / width, h / height), iw = width * r, ih = height * r;
  s.addImage({ path: p, x: x + (w - iw) / 2, y: y + (h - ih) / 2, w: iw, h: ih });
  return { x: x + (w - iw) / 2, y: y + (h - ih) / 2, w: iw, h: ih };
}

const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9'; // 10 x 5.625 in
pres.title = 'Customer Foundations & the C CRM';

pres.defineSlideMaster({
  title: 'CONTENT',
  background: { color: C.bg },
  objects: [{ text: { text: FOOT, options: { x: 0.48, y: 5.25, w: 4, h: 0.2, fontFace: BODY, fontSize: 7, color: C.soft, margin: 0 } } }],
  slideNumber: { x: 9.0, y: 5.25, w: 0.5, h: 0.2, fontFace: BODY, fontSize: 7, color: C.soft, align: 'right' },
});
pres.defineSlideMaster({ title: 'COVER', background: { color: C.bg }, objects: [] });

const T = (s, text, o) => s.addText(text, Object.assign({ margin: 0, isTextBox: true, fontFace: BODY, color: C.ink, valign: 'top' }, o));
const box = (s, x, y, w, h, fill, line) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: line || fill, width: 0.75 }, rectRadius: 0.08 });
const eyebrow = (s, text, x, y, w = 5, color = C.teal) => T(s, text.toUpperCase(), { x, y, w, h: 0.2, fontFace: HEAD, bold: true, fontSize: 8, color, charSpacing: 0.5 });
const badge = (s, x, y, n) => {
  s.addShape(pres.shapes.OVAL, { x, y, w: 0.3, h: 0.3, fill: { color: C.tealSoft }, line: { color: C.tealSoft } });
  T(s, String(n), { x, y, w: 0.3, h: 0.3, fontFace: HEAD, bold: true, fontSize: 9, color: C.teal, align: 'center', valign: 'middle' });
};
async function iconDot(s, x, y, name, d = 0.36, fill = C.tealSoft) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill } });
  const i = d * 0.52;
  s.addImage({ data: await icon(name), x: x + (d - i) / 2, y: y + (d - i) / 2, w: i, h: i });
}
const bullets = (items) => items.map((t, j) => ({ text: t, options: { bullet: { indent: 10 }, breakLine: j < items.length - 1 } }));

// ZTP level chip. Level 3 is solid teal, Level 2 tinted, Level 1 outlined.
function levelChip(s, x, y, lvl) {
  const fill = lvl === 3 ? C.teal : lvl === 2 ? C.tealSoft : C.white;
  const ink = lvl === 3 ? C.white : C.teal;
  box(s, x, y, 0.62, 0.24, fill, lvl === 1 ? C.teal : fill);
  T(s, 'Level ' + lvl, { x, y, w: 0.62, h: 0.24, fontFace: HEAD, bold: true, fontSize: 8, color: ink, align: 'center', valign: 'middle' });
}

function content(eb, title, sub, notes) {
  const s = pres.addSlide({ masterName: 'CONTENT' });
  eyebrow(s, eb, 0.48, 0.3, 9);
  T(s, title, { x: 0.48, y: 0.5, w: 9, h: 0.5, fontFace: HEAD, bold: true, fontSize: 24 });
  if (sub) T(s, sub, { x: 0.48, y: 1.05, w: 9, h: 0.3, fontSize: 12, color: C.soft });
  if (notes) s.addNotes(notes);
  return s;
}

async function tip(s, y, label, text, ico = 'FiInfo') {
  box(s, 0.45, y, 9.1, 0.42, C.goldSoft);
  s.addImage({ data: await icon(ico, C.gold), x: 0.62, y: y + 0.12, w: 0.18, h: 0.18 });
  T(s, [{ text: label + '  ', options: { bold: true, color: C.gold } }, { text }], { x: 0.9, y: y + 0.11, w: 8.5, h: 0.22, fontSize: 10 });
}

async function card(s, x, y, w, h, { n, ico, title, body }) {
  box(s, x, y, w, h, C.white, C.border);
  if (ico) await iconDot(s, x + 0.18, y + 0.18, ico, 0.36); else badge(s, x + 0.18, y + 0.2, n);
  T(s, title, { x: x + 0.18, y: y + 0.66, w: w - 0.36, h: 0.3, fontFace: HEAD, bold: true, fontSize: 12 });
  if (body) T(s, body, { x: x + 0.18, y: y + 0.98, w: w - 0.36, h: h - 1.08, fontSize: 9.5, color: C.soft });
}
async function cards(s, y, h, items) {
  const gap = 0.15, n = items.length, w = (9.1 - gap * (n - 1)) / n;
  for (let i = 0; i < n; i++) await card(s, 0.45 + i * (w + gap), y, w, h, Object.assign({ n: i + 1 }, items[i]));
}

// ZTP rule card: level chip, title, then exactly what the agent must do.
async function ztpCard(s, x, y, w, h, { lvl, ico, title, do: steps }) {
  box(s, x, y, w, h, C.white, C.border);
  await iconDot(s, x + 0.16, y + 0.15, ico, 0.36);
  T(s, title, { x: x + 0.62, y: y + 0.15, w: w - 1.45, h: 0.36, fontFace: HEAD, bold: true, fontSize: 11.5, valign: 'middle' });
  levelChip(s, x + w - 0.78, y + 0.21, lvl);
  T(s, bullets(steps), { x: x + 0.18, y: y + 0.62, w: w - 0.36, h: h - 0.7, fontSize: 9, color: C.ink, paraSpaceAfter: 3 });
}
async function ztpGrid(s, items, y = 1.3, h = 1.68) {
  const gap = 0.15, w = (9.1 - gap) / 2;
  for (let i = 0; i < items.length; i++) await ztpCard(s, 0.45 + (i % 2) * (w + gap), y + Math.floor(i / 2) * (h + gap), w, h, items[i]);
}

// Right-hand illustration panel used on cover, topic and closing slides.
async function panel(s, ill, y = 0.68, h = 4.26) {
  box(s, 5.81, y, 3.7, h, C.tealSoft);
  const cy = y + h / 2;
  s.addShape(pres.shapes.OVAL, { x: 6.26, y: cy - 1.4, w: 2.8, h: 2.8, fill: { color: C.white }, line: { color: C.white } });
  await fitImage(s, ill, 6.06, cy - 1.15, 3.2, 2.3);
  for (const [x, yy] of [[6.04, y + 0.22], [8.96, y + 0.22], [8.96, y + h - 0.58]]) {
    s.addShape(pres.shapes.OVAL, { x, y: yy, w: 0.36, h: 0.36, fill: { color: C.white }, line: { color: C.white } });
    s.addShape(pres.shapes.OVAL, { x: x + 0.135, y: yy + 0.135, w: 0.09, h: 0.09, fill: { color: C.teal }, line: { color: C.teal } });
  }
}

let topicNo = 0;
const TOPICS = 6;
async function topic(title, sub, learn, ill, notes) {
  topicNo++;
  const s = pres.addSlide({ masterName: 'CONTENT' });
  eyebrow(s, `Topic ${String(topicNo).padStart(2, '0')} of ${String(TOPICS).padStart(2, '0')}`, 0.62, 1.32, 4.5);
  T(s, title, { x: 0.62, y: 1.6, w: 4.8, h: 1.0, fontFace: HEAD, bold: true, fontSize: 28, valign: 'bottom' });
  T(s, sub, { x: 0.62, y: 2.7, w: 4.6, h: 0.5, fontSize: 11.5, color: C.soft });
  eyebrow(s, "You'll learn", 0.62, 3.28, 3, C.soft);
  for (let i = 0; i < learn.length; i++) {
    const y = 3.55 + i * 0.3;
    s.addShape(pres.shapes.OVAL, { x: 0.62, y: y + 0.02, w: 0.17, h: 0.17, fill: { color: C.tealSoft }, line: { color: C.tealSoft } });
    s.addImage({ data: await icon('FiCheck'), x: 0.655, y: y + 0.055, w: 0.1, h: 0.1 });
    T(s, learn[i], { x: 0.9, y, w: 4.3, h: 0.22, fontSize: 10.5 });
  }
  await panel(s, ill);
  if (notes) s.addNotes(notes);
  return s;
}

// Screen preview: a framed screenshot (customer details blurred).
async function screen(s, file, x, y, w, h, caption) {
  box(s, x, y, w, h, C.white, C.border);
  const r = await fitImage(s, file, x + 0.1, y + 0.1, w - 0.2, h - (caption ? 0.42 : 0.2));
  s.addShape(pres.shapes.RECTANGLE, { x: r.x, y: r.y, w: r.w, h: r.h, fill: { type: 'none' }, line: { color: C.border, width: 0.5 } });
  if (caption) T(s, caption, { x: x + 0.1, y: y + h - 0.3, w: w - 0.2, h: 0.2, fontSize: 8, color: C.soft, italic: true, align: 'center' });
}

// Icon list (left column) next to a screenshot.
async function iconList(s, x, y, w, items, rowH = 0.5) {
  for (let i = 0; i < items.length; i++) {
    const yy = y + i * (rowH + 0.08);
    box(s, x, yy, w, rowH, C.white, C.border);
    await iconDot(s, x + 0.12, yy + (rowH - 0.32) / 2, items[i][0], 0.32);
    T(s, items[i][1], { x: x + 0.55, y: yy, w: w - 0.65, h: rowH, fontFace: HEAD, bold: true, fontSize: 10.5, valign: 'middle' });
  }
}

async function walkthrough(eb, system, lookFor, notes) {
  const s = content(eb, 'Live walkthrough', null, notes);
  box(s, 0.45, 1.25, 9.1, 2.1, C.tealSoft);
  await iconDot(s, 0.8, 1.6, 'FiMonitor', 0.9, C.white);
  eyebrow(s, 'Trainer walkthrough', 2.0, 1.6, 7);
  T(s, `The trainer will open the live ${system} and walk you through the actual system.`, { x: 2.0, y: 1.85, w: 7.2, h: 0.75, fontFace: HEAD, bold: true, fontSize: 17 });
  T(s, 'Follow along and note your questions.', { x: 2.0, y: 2.7, w: 7, h: 0.25, fontSize: 11, color: C.soft });
  eyebrow(s, 'As you watch, look for', 0.48, 3.6, 5, C.soft);
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
    eyebrow(s, 'Homeaglow  ·  New Hire Care Training', 0.62, 0.6);
    T(s, 'Customer Foundations\n& the C CRM', { x: 0.62, y: 1.0, w: 4.9, h: 1.6, fontFace: HEAD, bold: true, fontSize: 32 });
    T(s, 'The rules we never bend, the path every customer takes, and the tools we use to help them.', { x: 0.62, y: 2.7, w: 4.6, h: 0.7, fontSize: 12.5, color: C.soft });
    T(s, 'Internal training material', { x: 0.62, y: 5.0, w: 3, h: 0.2, fontSize: 7, color: C.soft });
    await panel(s, 'ill6.png', 0.4, 4.82);
    s.addNotes('Welcome trainees. This module covers six topics: Zero Tolerance Policy, the Customer Journey, the Legacy C CRM, the C Dashboard, the New C CRM, and a hands-on CRM activity.');
  }

  // ---------- Agenda ----------
  {
    const s = content('Overview', 'Agenda', 'Six topics. Topics 3 to 6 happen mostly in the live systems.',
      'Walk through the agenda. Point out that topics 3 to 5 are live demos: the slides only set up what trainees will see.');
    const rows = [
      ['01', 'Zero Tolerance Policy', 'The non-negotiables, and exactly how to avoid them'],
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
    ['What ZTP is and why it exists', 'Violations by level', 'Penalties and other conditions', 'Exactly what to do to avoid each one'],
    'ill14.png',
    'ZTP is different from the rest of the Knowledge Library. Most articles are about judgment; this one is about hard lines that create legal, security, compliance or trust risk. Spend time here: every rule on the next slides says exactly what to do.');

  {
    const s = content('Zero Tolerance Policy', 'Not judgment calls. Hard lines.', 'Violations carry consequences regardless of intent.',
      'Most KB articles teach judgment and best practice. ZTP is the handful of things that, done wrong, are not just a bad customer experience: they create legal, security or compliance exposure for Homeaglow. Knowing it protects the agent\'s QA score and standing, and keeps every interaction safe and trustworthy. Rule of thumb: if you are ever unsure whether an action crosses a line, stop and ask a Team Lead first.');
    await cards(s, 1.6, 1.75, [
      { ico: 'FiFileText', title: 'Legal', body: 'Wording, auto-renewal and state laws we must follow.' },
      { ico: 'FiLock', title: 'Security', body: 'Personal data, logins and payment details stay protected.' },
      { ico: 'FiHeart', title: 'Trust', body: 'Customers and cleaners can rely on how we treat them.' },
    ]);
    await tip(s, 3.6, 'Unsure?', 'Stop and ask a Team Lead before you act. Asking is always better than a violation.', 'FiHelpCircle');
  }

  {
    const s = content('Zero Tolerance Policy', 'Violations at a glance', 'Every violation is an auto-fail on QA, except capturing or sharing work screens.',
      'Full list from the KB, by level. L3: Security violation; Capturing or sharing work screens (no QA score). L2: Not removing a C or CP\'s card info sent to CS; Ticket avoidance; Incorrect handling of OCC-ARL accounts; Non-compliance with legal wording changes; Sending internal notes as email/SMS; Disclosing internal information; Improper usage of AI tools; Not unsubscribing customers who opt out; Mishandling Washington State cancellations. L1: Rude/argumentative comms; Not following the bug reporting process; Editing Django inappropriately; Handling a Thumbtack account.');
    const cols = [
      [3, ['Security violation', 'Capturing or sharing work screens']],
      [2, ['Card info not removed', 'Ticket avoidance', 'OCC / ARL accounts', 'Legal wording', 'Internal notes sent out', 'Disclosing internal info', 'Improper AI tool use', 'Not unsubscribing', 'WA cancellations']],
      [1, ['Rude / argumentative comms', 'Bug reporting process', 'Editing Django', 'Thumbtack accounts']],
    ];
    const gap = 0.15, w = (9.1 - 2 * gap) / 3;
    cols.forEach(([lvl, items], i) => {
      const x = 0.45 + i * (w + gap);
      box(s, x, 1.5, w, 3.3, C.white, C.border);
      levelChip(s, x + 0.15, 1.66, lvl);
      T(s, lvl === 3 ? 'Most severe' : lvl === 2 ? 'Serious' : 'Still a ZTP', { x: x + 0.85, y: 1.66, w: w - 1.0, h: 0.24, fontSize: 9, color: C.soft, valign: 'middle' });
      const big = items.length > 5;
      T(s, bullets(items), { x: x + 0.15, y: 2.08, w: w - 0.3, h: 2.6, fontSize: big ? 9.5 : 10.5, paraSpaceAfter: big ? 3 : 6 });
    });
  }

  {
    const s = content('Zero Tolerance Policy', 'Penalties', 'Penalties grow with the level and with each repeat.',
      'Other conditions: 3 ZTPs of any level means termination. An L1 followed by an L2 means a 1-week suspension. Previously suspended for an L1 and then commits an L2 means termination. Multiple ZTPs of different levels in the same week: the penalty follows the higher level. L1 clearing period is 6 months. For Incorrect Handling of OCC accounts, the QA score is set at a minimum of -$25, not stacked on the LTV computation; if the standard markdown is higher, that is used.');
    const H = (t) => ({ text: t, options: { bold: true, color: C.teal, fill: { color: C.tealSoft }, fontFace: HEAD, align: 'center' } });
    const cell = (t) => ({ text: t, options: { color: t === '—' ? C.soft : C.ink, fill: { color: C.white }, align: 'center', bold: t.startsWith('Subject') } });
    const rows = [
      [H(''), H('Level 1'), H('Level 2'), H('Level 3')],
      [H('1st instance'), cell('Final warning (NTE)'), cell('Suspension, up to 5 days'), cell('Subject to termination')],
      [H('2nd instance'), cell('Suspension, up to 3 days'), cell('Subject to termination'), cell('—')],
      [H('3rd instance'), cell('Subject to termination'), cell('—'), cell('—')],
    ];
    s.addTable(rows, { x: 0.45, y: 1.5, w: 9.1, colW: [1.6, 2.5, 2.5, 2.5], rowH: 0.55, fontFace: BODY, fontSize: 10.5, valign: 'middle', border: { type: 'solid', pt: 0.75, color: C.border } });
    await tip(s, 3.95, 'Remember:', '3 ZTPs of any level means termination. L1s clear after 6 months.', 'FiAlertTriangle');
  }

  {
    const s = content('Security & compliance', 'Protect sensitive information', null,
      'Security violations: sending an auto-login link to the wrong person, sharing login credentials, mishandling documents with personal data (payment info, IDs, tax forms). A wrong auto-login link can expose private details and create legal/privacy risk. Card info: required for PCI compliance. If you receive card info, notify a Team Lead to remove it from the CRM and direct the customer to the CC update page on their dashboard. Internal notes are for risk assessments, offers/action plans and internal follow-ups. Internal info never to discuss: CP DTC flags, CP Swap details, CP Tiering levels, known fraud/safety flags, features not yet released on NCW.');
    await ztpGrid(s, [
      { lvl: 3, ico: 'FiKey', title: 'Security violation', do: ['Double-check the recipient before sending any auto-login link or secure info.', 'Never share login credentials.', 'Keep documents with personal data (payment info, IDs, tax forms) inside our systems.'] },
      { lvl: 2, ico: 'FiCreditCard', title: 'Card details sent to CS', do: ['Notify a Team Lead right away so the card info is removed from the CRM.', 'Direct the customer to the CC update page on their dashboard.', 'Never ask for card info or card photos, and never put card details in internal notes.'] },
      { lvl: 2, ico: 'FiEyeOff', title: 'Internal notes sent out', do: ['Use internal notes only for risk assessments, offers/action plans and internal follow-ups.', 'Check the channel before you send: an internal note never goes out as an email or SMS.'] },
      { lvl: 2, ico: 'FiSlash', title: 'Disclosing internal info', do: ['Never discuss CP DTC flags, CP Swap details, CP Tiering levels, fraud/safety flags or unreleased NCW features.', 'Test: not in the Help Center or public comms? It\'s internal.'] },
    ]);
  }

  {
    const s = content('Security & compliance  ·  Level 2', 'Say it the legal way', 'Cleaners are independent contractors, not Homeaglow employees. Always word it that way.',
      'Legal Wording Compliance (Level 2, auto-fail). CPs are Independent Contractors. Referring to Homeaglow itself is fine: "your Homeaglow cleaning" and "your cleaning on Homeaglow" are both OK. Ask trainees to rewrite a few sentences out loud.');
    const pairs = [
      ['Your cleaner on Homeaglow', 'Your Homeaglow cleaner'],
      ['Find a cleaner on Homeaglow', 'Find a Homeaglow cleaner'],
      ['Your cleaner / Your cleaner Alice', 'Our cleaner'],
      ['Your cleaner will see you at <address>', 'We will see you at <address>'],
      ['Your cleaner was unable to', 'We were unable to'],
      ['Your requested cleaner is still reviewing', 'We are still waiting for your requested cleaner'],
    ];
    const head = [
      { text: 'Do say', options: { bold: true, color: C.white, fill: { color: C.teal }, fontFace: HEAD } },
      { text: "Don't say", options: { bold: true, color: C.teal, fill: { color: C.tealSoft }, fontFace: HEAD } },
    ];
    const body = pairs.map(([a, b]) => [{ text: a, options: { color: C.ink, bold: true, fill: { color: C.white } } }, { text: b, options: { color: C.soft, fill: { color: C.white } } }]);
    s.addTable([head, ...body], { x: 0.45, y: 1.55, w: 9.1, colW: [4.55, 4.55], rowH: 0.42, fontFace: BODY, fontSize: 10.5, valign: 'middle', border: { type: 'solid', pt: 0.75, color: C.border }, margin: [0, 0.16, 0, 0.16] });
  }

  {
    const s = content('Security & compliance  ·  Level 3', 'Work screens stay on work screens', 'No exceptions, even if the screen looks blurred, covered or empty.',
      'Anything on a work screen can include names, addresses, payment details, internal notes, flags, other tabs or pop-up notifications. Once a capture exists outside our systems we cannot delete it, track it or undo it. "I blurred it" or "nothing was on screen" does not count: the risk is the capture itself. Level 3: subject to termination on the first instance.');
    await cards(s, 1.6, 1.75, [
      { ico: 'FiCamera', title: 'Never photograph it', body: 'No photos or videos of a work screen, with a phone or any other device.' },
      { ico: 'FiMaximize', title: 'Never capture it', body: 'No screenshots or screen recordings of a work screen.' },
      { ico: 'FiShare2', title: 'Never share it', body: 'Nowhere: not personal chats, social media or coworker group chats.' },
    ]);
    await tip(s, 3.6, 'Why:', 'once a capture leaves our systems, we can\'t delete it, track it or undo it.', 'FiAlertTriangle');
  }

  {
    const s = content('Customer handling', 'Handle these customers by the rules', null,
      'OCC-ARL: some states have Auto-Renewal Laws requiring immediate cancellation on request, with no extra steps or delays. Failing to cancel an FC when required is the ZTP; if the correct call was not to cancel but a retention offer was given anyway, that is a markdown, not a ZTP. Unsubscribe: if a customer says they don\'t want to be contacted, or replies STOP/UNSUBSCRIBE to an email. Washington State: WA customers are exempt from both MCT and the ETF; this is a standing legal exemption, not a case-by-case waiver. Tone: condescending ("You didn\'t understand...", "As I\'ve already said...") and passive-aggressive ("As I\'ve told you before...", "You should already know this.") replies. Do: educate, personalize, focus on resolution. Don\'t: argue, use generic statements, over-empathize.');
    await ztpGrid(s, [
      { lvl: 2, ico: 'FiRefreshCcw', title: 'OCC / ARL accounts', do: ['In Auto-Renewal Law states, cancel the FC immediately when asked: no extra steps, no delays.', 'Unsure? Check the OCC Non-Retention Process or ask a Team Lead before replying.'] },
      { lvl: 2, ico: 'FiBellOff', title: 'Unsubscribe requests', do: ['Customer says "don\'t contact me" or replies STOP / UNSUBSCRIBE: unsubscribe them.', 'Confirm the unsubscribe to them, then handle any other pending items as usual.'] },
      { lvl: 2, ico: 'FiMapPin', title: 'Washington State cancellations', do: ['Process the cancellation immediately.', 'No ETF, never cite MCT, and no retention offers before cancelling.', 'Never present ETF or MCT terms as if they apply.'] },
      { lvl: 1, ico: 'FiMessageCircle', title: 'Rude / argumentative comms', do: ['No condescending ("You didn\'t understand...") or passive-aggressive ("As I\'ve told you before...") replies.', 'Educate, personalize and focus on resolution. Never argue.'] },
    ]);
  }

  {
    const s = content('Operational compliance', 'Follow the process, every time', null,
      'Ticket avoidance: deliberately skipping a ticket that needs a response or action. Clearing a ticket where the customer just says "thanks", "got it" or similar is NOT avoidance. Bug reporting: file a report when an issue can\'t be fixed by a dashboard/CRM update, or when something might be a bug or product idea. Bug Validation Form and Confirmed Bugs Tracker are linked in the KB. Django bypasses system codes, so one wrong edit can create bugs or data issues. CS-editable: duplicate account placeholders, a CP\'s SSN, invalidating flags (except Holdpay), reassigning/adjusting voucher duration, editing a DHJ voucher\'s MF/MCT (if no FC table yet), marking a CP review hidden (approved cases only), editing/restricting Requested CPs (if unavailable on OCW), manually adding CPs to recreated jobs. Thumbtack: transferring without reaching an Aftersales agent is still compliant: the transfer is the compliance action.');
    await ztpGrid(s, [
      { lvl: 2, ico: 'FiInbox', title: 'Ticket avoidance', do: ['Work every ticket assigned to your department, including hard ones and charge disputes.', 'Never use CRM/NavBar filters to dodge tickets, or clear a ticket without resolving it.', 'Unsure if it needs action? Ask in the designated Slack channel.'] },
      { lvl: 1, ico: 'FiAlertOctagon', title: 'Bug reporting process', do: ['Check the Confirmed Bugs Tracker first, then submit the Bug Validation Form.', 'Once a card is escalated, don\'t add sample or unrelated comments.'] },
      { lvl: 1, ico: 'FiDatabase', title: 'Editing Django', do: ['Only make the CS-approved edits listed in the KB.', 'Escalate review edits in NCW and job history edits for LMCs/LOs to SMEs or Managers.', 'When in doubt, don\'t edit it. Escalate it.'] },
      { lvl: 1, ico: 'FiRepeat', title: 'Thumbtack accounts', do: ['Spot it: system flag in New CRM, an internal note, or the customer says so.', 'Transfer to the Thumbtack queue and notify an available Aftersales agent.', 'Then stop: don\'t reply, resolve or take any other action.'] },
    ]);
  }

  // ================= 2. CUSTOMER JOURNEY =================
  await topic('Customer Journey', 'The map of a customer\'s whole experience, from first hearing about us to their cleaning.',
    ['The five steps', 'How customers find us', 'Three ways to sign up', 'Where tickets come from'],
    'ill3.png',
    'Every ticket sits somewhere on this journey. Knowing where an issue sits lets you anticipate needs and pick the right playbook.');

  {
    const s = content('Customer Journey', 'Five steps, one journey', 'Every ticket you handle sits somewhere on this path.',
      'Understanding the full shape lets you anticipate what a customer needs before they fully explain it, catch friction before it becomes a complaint, and know where an issue sits. A booking problem needs a very different response from a post-cleaning service issue, even when the tone sounds the same. This page is the map; playbooks like PCQ, Overcharged Hours and Retention are what you use once you know where you are.');
    const steps = [['FiSearch', 'Learns about us'], ['FiUserPlus', 'Signs up'], ['FiCalendar', 'Books'], ['FiSliders', 'Manages the appointment'], ['FiStar', 'Receives and pays']];
    const w = 1.62, gap = (9.1 - 5 * w) / 4;
    for (let i = 0; i < 5; i++) {
      const x = 0.45 + i * (w + gap);
      box(s, x, 1.7, w, 2.0, C.white, C.border);
      await iconDot(s, x + (w - 0.6) / 2, 1.95, steps[i][0], 0.6);
      T(s, `STEP ${i + 1}`, { x, y: 2.72, w, h: 0.2, fontFace: HEAD, bold: true, fontSize: 8, color: C.teal, align: 'center', charSpacing: 0.5 });
      T(s, steps[i][1], { x: x + 0.1, y: 2.95, w: w - 0.2, h: 0.55, fontFace: HEAD, bold: true, fontSize: 11.5, align: 'center' });
      if (i < 4) s.addImage({ data: await icon('FiChevronRight', C.soft), x: x + w + gap / 2 - 0.09, y: 2.6, w: 0.18, h: 0.18 });
    }
    await tip(s, 4.1, 'Why it matters:', 'know where the customer is, and you\'ll know which playbook to use.', 'FiCompass');
  }

  {
    const s = content('Step 1  ·  Learns about us', 'How customers find us', null,
      'Customers find us through our website (homeaglow.com) or a sister brand (Dazzling Cleaning, Cozy Maid, Bubbly Cleaning, Dapper Maids); search engines, the most common source; TV commercials; friends and family referrals; and social media (Facebook, YouTube, TikTok, etc).');
    const items = [['FiSearch', 'Search engines', 'The most common source'], ['FiGlobe', 'Our websites', 'Homeaglow and sister brands'], ['FiTv', 'TV commercials', ''], ['FiUsers', 'Friends & family', 'Referrals'], ['FiThumbsUp', 'Social media', 'Facebook, YouTube, TikTok']];
    const w = 1.7, gap = (9.1 - 5 * w) / 4;
    for (let i = 0; i < 5; i++) {
      const x = 0.45 + i * (w + gap);
      box(s, x, 1.4, w, 2.3, C.white, C.border);
      await iconDot(s, x + (w - 0.7) / 2, 1.7, items[i][0], 0.7);
      T(s, items[i][1], { x: x + 0.1, y: 2.6, w: w - 0.2, h: 0.3, fontFace: HEAD, bold: true, fontSize: 11.5, align: 'center' });
      T(s, items[i][2], { x: x + 0.1, y: 2.92, w: w - 0.2, h: 0.5, fontSize: 9, color: C.soft, align: 'center' });
    }
    await tip(s, 3.95, 'Most common:', 'search engines.', 'FiTrendingUp');
  }

  {
    const s = content('Step 2  ·  Signs up', 'Three ways to sign up', 'The path a customer took shapes their billing and what options they have.',
      'Legacy DHJ: the customer buys a discounted voucher from <brand>.com/deal. First Cleaning Free (FCF): signs up for a membership directly, pays $19 upfront (more for a longer first cleaning), membership starts immediately. From August 3, 2026 all new sign-ups are FCF only, so DHJ is a closed, shrinking group. Non-member: third-party merchant vouchers (Groupon, Living Social, Gilt City) never converted; Homeaglow Gift Cards (never trigger FC); one-time cleanings sold by Inside Sales (paid upfront, full price); trial cleanings sold by Inside Sales (regular price, 30 days to cancel the membership without obligation). Non-members are billed differently and have nothing to "cancel"; the question is usually just whether a specific charge should be refunded.');
    await cards(s, 1.6, 1.75, [
      { n: 1, title: 'Legacy DHJ', body: 'Bought a discounted voucher from <brand>.com/deal.' },
      { n: 2, title: 'First Cleaning Free', body: 'Joins the membership directly. Pays $19 upfront.' },
      { n: 3, title: 'Non-member', body: 'Groupon, gift card, one-time or trial cleaning. No membership.' },
    ]);
    await tip(s, 3.6, 'Since Aug 3, 2026:', 'all new membership sign-ups are First Cleaning Free.', 'FiCalendar');
  }

  {
    const s = content('Steps 3 & 4  ·  Books and manages', 'Booking and managing', null,
      'Booking: customers with a voucher or membership are prompted to opt into a recurring schedule and add alternate start times; both increase the odds of a successful match. Customers without one are not required to. Managing: through their dashboard, up to 6 hours before the start time, a customer can reschedule, cancel, or match with a different cleaner.');
    box(s, 0.45, 1.3, 4.45, 3.3, C.white, C.border);
    await iconDot(s, 0.65, 1.5, 'FiCalendar', 0.4);
    T(s, 'Books', { x: 0.65, y: 2.02, w: 4, h: 0.3, fontFace: HEAD, bold: true, fontSize: 14 });
    T(s, 'Members are prompted to add:', { x: 0.65, y: 2.38, w: 4, h: 0.25, fontSize: 10, color: C.soft });
    for (const [i, t] of ['A recurring schedule', 'Alternate start times'].entries()) {
      box(s, 0.65, 2.75 + i * 0.5, 4.05, 0.4, C.tealSoft);
      T(s, t, { x: 0.82, y: 2.75 + i * 0.5, w: 3.8, h: 0.4, fontSize: 10.5, bold: true, color: C.teal, valign: 'middle' });
    }
    T(s, 'Both raise the odds of a match.', { x: 0.65, y: 3.85, w: 4, h: 0.25, fontSize: 9.5, color: C.soft, italic: true });
    box(s, 5.1, 1.3, 4.45, 3.3, C.white, C.border);
    T(s, '6', { x: 5.3, y: 1.4, w: 0.7, h: 0.9, fontFace: HEAD, bold: true, fontSize: 48, color: C.teal });
    T(s, 'hours before start', { x: 5.95, y: 1.75, w: 3.4, h: 0.3, fontFace: HEAD, bold: true, fontSize: 12, color: C.teal });
    T(s, 'Up until then, from their dashboard, customers can:', { x: 5.3, y: 2.38, w: 4.1, h: 0.25, fontSize: 10, color: C.soft });
    const acts = [['FiClock', 'Reschedule'], ['FiXCircle', 'Cancel'], ['FiUserCheck', 'Match a new cleaner']];
    for (let i = 0; i < 3; i++) {
      const x = 5.3 + i * 1.37;
      box(s, x, 2.75, 1.27, 1.1, C.tealSoft);
      s.addImage({ data: await icon(acts[i][0]), x: x + 0.49, y: 2.88, w: 0.28, h: 0.28 });
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
    ]);
    await tip(s, 3.6, 'Next:', 'each issue has its own playbook in the Knowledge Library.', 'FiBookOpen');
  }

  // ================= 3. LEGACY C CRM =================
  await topic('Introduction to C CRM (Legacy)', 'Our central hub for customers, cleaners, jobs and payments.',
    ['What the CRM is for', 'Why we have two CRMs', 'A live walkthrough'],
    'ill11.png',
    'We currently have two CRM systems. We start with the Legacy C CRM because many actions still live there.');

  {
    const s = content('Legacy C CRM', 'One hub for everything', 'Where we manage accounts, talk to users and find what we need to help them.',
      'The CRM is the central hub for managing customer and cleaner accounts, jobs and payments. It is also where we communicate with users, manage tickets, see our daily targets, and access the information that helps us handle users\' concerns. We currently have two CRM systems; we introduce the Legacy CRM first. The screenshot is only a preview; the trainer will walk through the live system next. Customer details in the screenshot are blurred on purpose (see ZTP: work screens).');
    await iconList(s, 0.45, 1.5, 3.0, [['FiUsers', 'Customer & cleaner accounts'], ['FiBriefcase', 'Jobs'], ['FiCreditCard', 'Payments'], ['FiMessageSquare', 'Messages with users'], ['FiInbox', 'Tickets'], ['FiTarget', 'Daily targets']], 0.42);
    await screen(s, 'legacy.png', 3.6, 1.5, 5.95, 3.15, 'Legacy C CRM, preview only. Customer details blurred.');
    await tip(s, 4.75, 'Two CRMs:', 'we have a Legacy CRM and a New CRM. We start with Legacy.', 'FiLayers');
  }

  await walkthrough('Legacy C CRM', 'Legacy C CRM', [['FiUser', 'Account details'], ['FiBriefcase', 'Jobs and charges'], ['FiMessageSquare', 'Messages and tickets'], ['FiTarget', 'Daily targets']],
    'Switch to the live Legacy C CRM now. Do not demo from the slides. Suggested flow: open a customer account, show account details and flags, job history and charges, the message/ticket thread, and where daily targets show. Adjust to the cohort.');

  // ================= 4. C DASHBOARD =================
  await topic('Introduction to C Dashboard', 'Where customers manage their account and bookings.',
    ['What the customer sees', 'What they can do themselves', 'A live walkthrough'],
    'ill22.png',
    'The C Dashboard is the customer-facing side. Seeing it helps agents guide customers and understand what they already tried.');

  {
    const s = content('C Dashboard', 'The customer\'s side', 'Where customers manage their account and bookings.',
      'The Customer dashboard is where customers manage their account and bookings: book a cleaning, see upcoming and past cleanings and alerts (e.g. a cancelled cleaning), and, up to 6 hours before start, reschedule, cancel or match with a different cleaner. Knowing what they see lets agents point customers to self-serve steps and understand what they have already tried. Customer name and cleaner photos in the screenshot are blurred.');
    await iconList(s, 0.45, 1.5, 3.0, [['FiPlusCircle', 'Book a cleaning'], ['FiBell', 'Alerts about their bookings'], ['FiCalendar', 'Upcoming cleanings'], ['FiClock', 'Past cleanings'], ['FiUser', 'Account settings']], 0.5);
    await screen(s, 'dashboard.png', 3.6, 1.5, 5.95, 3.15, 'C Dashboard, preview only. Customer details blurred.');
    await tip(s, 4.75, 'Why it helps:', 'when you know what the customer sees, you can guide them step by step.', 'FiEye');
  }

  await walkthrough('C Dashboard', 'C Dashboard', [['FiHome', 'The home view'], ['FiCalendar', 'Upcoming cleanings'], ['FiClock', 'Past cleanings'], ['FiSettings', 'Account settings']],
    'Switch to the live C Dashboard now. Suggested flow: home view and alerts, booking a cleaning, managing an upcoming cleaning (reschedule, cancel, match), past cleanings, and account settings.');

  // ================= 5. NEW C CRM =================
  await topic('Introduction to New C CRM', 'Where we respond to customers today.',
    ['What the New CRM is for', 'When you still need Legacy', 'A live walkthrough'],
    'ill30.png',
    'The New C CRM is where agents respond to customers. Some actions are still only available in the Legacy CRM.');

  {
    const s = content('New C CRM', 'Where we respond to customers', 'Every reply to a customer goes out from the New C CRM.',
      'The New C CRM is where we respond to customers. The screenshot is a preview of a ticket: ticket header, customer information, related jobs, and the message thread with the reply box. Customer details are blurred. The trainer walks through the live system next.');
    await iconList(s, 0.45, 1.5, 3.0, [['FiInbox', 'Ticket status and SLA'], ['FiUser', 'Customer information'], ['FiBriefcase', 'Related jobs'], ['FiMessageSquare', 'Messages and internal notes'], ['FiSend', 'Reply with macros']], 0.5);
    await screen(s, 'newcrm.png', 3.6, 1.5, 5.95, 3.15, 'New C CRM, preview only. Customer details blurred.');
    await tip(s, 4.75, 'Important:', 'some actions can still only be done in the Legacy CRM.', 'FiAlertCircle');
  }

  {
    const s = content('New C CRM', 'Two CRMs, one workflow', null,
      'Important: we respond to customers through the New CRM. At this time, there are still some actions that can only be done in the Legacy CRM, so agents use both. The trainer will point out which actions still require Legacy during the walkthrough.');
    const col = async (x, ico, eb, title, body) => {
      box(s, x, 1.3, 4.1, 2.55, C.white, C.border);
      await iconDot(s, x + 0.25, 1.55, ico, 0.6);
      eyebrow(s, eb, x + 0.25, 2.35, 3.6);
      T(s, title, { x: x + 0.25, y: 2.6, w: 3.6, h: 0.4, fontFace: HEAD, bold: true, fontSize: 16 });
      T(s, body, { x: x + 0.25, y: 3.05, w: 3.6, h: 0.6, fontSize: 10.5, color: C.soft });
    };
    await col(0.45, 'FiMessageSquare', 'New C CRM', 'Respond to customers', 'Every reply to a customer goes out from here.');
    s.addImage({ data: await icon('FiRepeat', C.soft), x: 4.8, y: 2.4, w: 0.4, h: 0.4 });
    await col(5.45, 'FiDatabase', 'Legacy C CRM', 'Some actions still live here', 'Use it for actions the New CRM can\'t do yet.');
    await tip(s, 4.1, 'Important:', 'reply in the New CRM. Switch to Legacy only for actions the New CRM can\'t do yet.', 'FiAlertCircle');
  }

  await walkthrough('New C CRM', 'New C CRM', [['FiInbox', 'Ticket and status'], ['FiUser', 'Customer info'], ['FiMessageSquare', 'Replying to customers'], ['FiLayers', 'What still needs Legacy']],
    'Switch to the live New C CRM now. Suggested flow: ticket header (status, assignee, category, SLA), customer information panel, related jobs, the message thread and internal notes, composing a reply with macros. Call out which actions still require the Legacy CRM.');

  // ================= 6. C CRM ACTIVITY =================
  await topic('C CRM Activity', 'Put what you saw into practice.',
    ['Work in the live CRM', 'Practice real scenarios', 'Ask questions as you go'],
    'ill12.png',
    'Trainer decides the activity format, scenarios, and level of difficulty based on trainee readiness.');

  {
    const s = content('C CRM Activity', 'Activity setup', 'Your trainer will set the activity for this cohort.',
      'TRAINER NOTE: Trainer decides the activity format, scenarios, and level of difficulty based on trainee readiness. Fill in the three boxes on this slide before the session (or live with the class).');
    const fields = [['FiGrid', 'Format', 'Solo, pairs or groups'], ['FiFileText', 'Scenarios', 'Which cases to work'], ['FiBarChart2', 'Difficulty', 'Based on readiness']];
    const gap = 0.15, w = (9.1 - 2 * gap) / 3;
    for (let i = 0; i < 3; i++) {
      const x = 0.45 + i * (w + gap);
      box(s, x, 1.5, w, 2.75, C.white, C.border);
      await iconDot(s, x + 0.2, 1.7, fields[i][0], 0.42);
      T(s, fields[i][1], { x: x + 0.2, y: 2.25, w: w - 0.4, h: 0.3, fontFace: HEAD, bold: true, fontSize: 13 });
      T(s, fields[i][2], { x: x + 0.2, y: 2.57, w: w - 0.4, h: 0.25, fontSize: 9.5, color: C.soft });
      box(s, x + 0.2, 2.95, w - 0.4, 1.1, C.bg, C.border);
      T(s, 'Click to add', { x: x + 0.32, y: 2.95, w: w - 0.64, h: 1.1, fontSize: 9.5, color: C.soft, italic: true, valign: 'middle' });
    }
    await tip(s, 4.45, 'Trainer note:', 'trainer decides the format, scenarios and difficulty based on trainee readiness.', 'FiEdit3');
  }

  // ---------- Close ----------
  {
    const s = pres.addSlide({ masterName: 'CONTENT' });
    eyebrow(s, 'Wrap-up', 0.62, 1.9);
    T(s, 'Questions?', { x: 0.62, y: 2.15, w: 5, h: 0.75, fontFace: HEAD, bold: true, fontSize: 34 });
    T(s, 'When in doubt, check the Knowledge Library or ask your Trainer or Team Lead.', { x: 0.62, y: 2.95, w: 4.6, h: 0.5, fontSize: 12, color: C.soft });
    await panel(s, 'ill9.png');
    s.addNotes('Open the floor for questions. Recap: ZTP hard lines, the five-step customer journey, and the three systems: Legacy C CRM, C Dashboard and New C CRM.');
  }

  await pres.writeFile({ fileName: OUT });
  console.log('wrote', OUT);
})();
