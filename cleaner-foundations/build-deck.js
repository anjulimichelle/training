// Builds the "Cleaner Foundations & CP CRM" training deck.
// Usage: node deck.js out.pptx [imageDir]
const path = require('path');
const pptxgen = require('pptxgenjs');
const React = require('react');
const RDS = require('react-dom/server');
const sharp = require('sharp');
const Fi = require('react-icons/fi');

const OUT = process.argv[2] || 'deck.pptx';
const IMG = process.argv[3] || path.join(__dirname, '../img2');

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
pres.title = 'Cleaner Foundations & the CP CRM';

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
const TOPICS = 11;
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

// Screen preview: a framed screenshot (sample training details).
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

// Simple bordered table in the deck style. head: [labels], rows: [[cells]], colW: inches.
function table(s, head, rows, { y = 1.45, colW, fontSize = 9, rowH = 0.3, boldFirst = true } = {}) {
  const H = head.map(t => ({ text: t, options: { bold: true, color: C.teal, fill: { color: C.tealSoft }, fontFace: HEAD } }));
  const R = rows.map(r => r.map((t, i) => ({ text: t, options: { color: i === 0 ? C.ink : C.soft, bold: boldFirst && i === 0, fill: { color: C.white } } })));
  s.addTable([H, ...R], { x: 0.45, y, w: 9.1, colW, rowH, fontFace: BODY, fontSize, valign: 'middle', border: { type: 'solid', pt: 0.75, color: C.border }, margin: [0.03, 0.1, 0.03, 0.1] });
}

// Two side-by-side panels, each a title plus bullets.
async function twoCol(s, left, right, y = 1.4, h = 2.6) {
  const w = (9.1 - 0.15) / 2;
  for (const [i, p] of [left, right].entries()) {
    const x = 0.45 + i * (w + 0.15);
    box(s, x, y, w, h, C.white, C.border);
    await iconDot(s, x + 0.2, y + 0.2, p.ico, 0.42);
    T(s, p.title, { x: x + 0.75, y: y + 0.2, w: w - 0.95, h: 0.42, fontFace: HEAD, bold: true, fontSize: 13, valign: 'middle' });
    T(s, bullets(p.items), { x: x + 0.2, y: y + 0.78, w: w - 0.4, h: h - 0.9, fontSize: 10, paraSpaceAfter: 5 });
  }
}

// Full-width infographic with a caption.
async function infographic(eb, title, file, notes) {
  const s = content(eb, title, null, notes);
  box(s, 0.45, 1.15, 9.1, 3.95, C.white, C.border);
  await fitImage(s, file, 0.6, 1.25, 8.8, 3.75);
  return s;
}


// Screenshot with numbered highlight boxes. marks: [[n, x0, y0, x1, y1]] in image pixels.
async function shot(s, file, x, y, w, h, marks = []) {
  box(s, x, y, w, h, C.white, C.border);
  const r = await fitImage(s, file, x + 0.08, y + 0.08, w - 0.16, h - 0.16);
  const { width } = await sharp(path.join(IMG, file)).metadata();
  const k = r.w / width;
  for (const [n, x0, y0, x1, y1] of marks) {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: r.x + x0 * k, y: r.y + y0 * k, w: (x1 - x0) * k, h: (y1 - y0) * k, fill: { type: 'none' }, line: { color: C.teal, width: 2 }, rectRadius: 0.03 });
    const bx = r.x + x0 * k - 0.11, by = r.y + y0 * k - 0.11;
    s.addShape(pres.shapes.OVAL, { x: bx, y: by, w: 0.22, h: 0.22, fill: { color: C.teal }, line: { color: C.white, width: 1 } });
    T(s, String(n), { x: bx, y: by, w: 0.22, h: 0.22, fontFace: HEAD, bold: true, fontSize: 8, color: C.white, align: 'center', valign: 'middle' });
  }
}

// Numbered legend rows matching the highlight numbers.
// Item marker: a number matching a highlight, or '—' for a row with no highlight.
function legend(s, x, y, w, items, rowH = 0.55, start = 1) {
  items.forEach(([title, body, mark], i) => {
    const yy = y + i * (rowH + 0.1);
    const label = mark || String(start + i), plain = label === '—';
    box(s, x, yy, w, rowH, C.white, C.border);
    s.addShape(pres.shapes.OVAL, { x: x + 0.14, y: yy + (rowH - 0.28) / 2, w: 0.28, h: 0.28, fill: { color: plain ? C.tealSoft : C.teal }, line: { color: plain ? C.tealSoft : C.teal } });
    T(s, label, { x: x + 0.14, y: yy + (rowH - 0.28) / 2, w: 0.28, h: 0.28, fontFace: HEAD, bold: true, fontSize: 9, color: plain ? C.teal : C.white, align: 'center', valign: 'middle' });
    T(s, [{ text: title, options: { bold: true, fontFace: HEAD, breakLine: true } }, { text: body, options: { color: C.soft, fontSize: 9 } }], { x: x + 0.55, y: yy + 0.05, w: w - 0.65, h: rowH - 0.1, fontSize: 10.5, valign: 'middle' });
  });
}

(async () => {
  // ---------- Cover ----------
  {
    const s = pres.addSlide({ masterName: 'COVER' });
    eyebrow(s, 'Homeaglow  ·  New Hire Care Training', 0.62, 0.6);
    T(s, 'Cleaner Foundations\n& the CP CRM', { x: 0.62, y: 1.0, w: 4.9, h: 1.6, fontFace: HEAD, bold: true, fontSize: 32 });
    T(s, 'How cleaners join and work on Homeaglow, and the tools and rules we use to support them.', { x: 0.62, y: 2.7, w: 4.6, h: 0.7, fontSize: 12.5, color: C.soft });
    T(s, 'Internal training material', { x: 0.62, y: 5.0, w: 3, h: 0.2, fontSize: 7, color: C.soft });
    await panel(s, 'ill7.png', 0.4, 4.82);
    s.addNotes('Welcome. This module covers the cleaner (CP) side of Homeaglow: the Cleaner Journey, the Legacy CP CRM and CP Dashboard, CP profile statuses, deactivation reasons, CP penalties on the C side, operating models, tiering, job statuses, and the Pending Invoice action guide. It ends with a knowledge check.');
  }

  // ---------- Agenda ----------
  {
    const s = content('Overview', 'Agenda', 'Eleven topics. Topics 2 and 3 happen mostly in the live systems.',
      'Walk through the agenda. Topics 2 and 3 are live walkthroughs: the slides only set up what trainees will see.');
    const rows = [
      ['01', 'Cleaner Journey', 'How a cleaner goes from applying to getting paid'],
      ['02', 'Legacy CP CRM', 'Where we manage cleaner accounts'],
      ['03', 'CP Dashboard', 'What cleaners see and manage'],
      ['04', 'CP Profile Status', 'What each status lets a cleaner do'],
      ['05', 'CP-Facing Deactivation', 'The reason codes cleaners see'],
      ['06', 'CP Penalty: C-Side Handling', 'What to do on the cleaner\'s side when a customer reports them'],
      ['07', 'Operating Models', 'Why location changes the rules'],
      ['08', 'CP Tiering', 'Who sees and claims jobs first'],
      ['09', 'Job Statuses', 'Submitted, Claimed, Invoiced, Cancelled, Pending Invoice'],
      ['10', 'AG: Pending Invoice', 'The action guide and cancellation reason codes'],
      ['11', 'Knowledge Check', 'Set by your trainer'],
    ];
    const hdr = ['#', 'Topic', 'You will learn'].map(t => ({ text: t, options: { bold: true, color: C.teal, fill: { color: C.tealSoft }, fontFace: HEAD } }));
    const body = rows.map(r => r.map((t, i) => ({ text: t, options: { color: i === 0 ? C.teal : i === 1 ? C.ink : C.soft, bold: i < 2, fill: { color: C.white } } })));
    s.addTable([hdr, ...body], { x: 0.45, y: 1.45, w: 9.1, colW: [0.7, 3.0, 5.4], rowH: 0.29, fontFace: BODY, fontSize: 9.5, valign: 'middle', border: { type: 'solid', pt: 0.75, color: C.border }, margin: [0, 0.14, 0, 0.14] });
  }

  // ================= 1. CLEANER JOURNEY =================
  await topic('Cleaner Journey', 'The path a cleaner takes, from first hearing about us to getting paid for their first job.',
    ['How cleaners find us and sign up', 'What blocks a cleaner from claiming', 'When the background check happens', 'What happens after the first job'],
    'ill7.png',
    'This is the mirror image of the Customer Journey, from the CP\'s side. It explains things you\'ll see constantly in CP tickets: why a CP\'s status shows cpprofile_setpay or cpsignica instead of active, why a background check hasn\'t cleared even though the CP claimed a job days ago, or why a CP can\'t claim jobs despite finishing registration.');

  await infographic('Cleaner Journey', 'The ideal cleaner journey', 'info10.png',
    'Step 1 Explore information: cleaners learn about us through homeaglow.com/apply (job descriptions, benefits, how-to videos, dashboard walkthroughs), search engines (most common), testimonials, friends and family, or by contacting Support. Step 2 Submit job application (registration). Step 3 Claim and manage jobs: the CP moves from registered to actively working. Step 4 Background check: starts only after the CP claims their first job, so they are verified before their first appointment. Step 5 Complete the job and get paid through their chosen payment method.');

  await infographic('Cleaner Journey  ·  Onboarding', 'The ideal onboarding process', 'info11.png',
    'Step 1: go to homeaglow.com/apply and enter a zip code. Step 2: fill out details (name, email, address etc.). Step 3: take a short quiz about cleaning homes. Step 4: set up their profile: pay rate, coverage area, schedule, profile photo, SSN and pay details. Step 5: watch the welcome video, after which their profile is created and they can begin claiming jobs. Everything before the first claim is the registration phase; the key uncertainty is whether the CP will follow through and complete a job.');

  {
    const s = content('Cleaner Journey', 'Before the first job', 'Two things are required to claim jobs. The background check comes after the first claim.',
      'Registration is normally quick, but two things commonly block a CP from claiming jobs: no pay details (debit cards and bank accounts are accepted) and no SSN (needed to run the background check, which protects customer safety). Once a CP claims their first job, we assume they are committed, intend to complete it and get paid, and could keep claiming. The background check (BGC) starts only after that first claim, not at registration, so the CP is verified before the first scheduled appointment. BGCs are a free benefit for CPs and a core reason customers trust the platform.');
    // Left: what's required before claiming
    box(s, 0.45, 1.45, 5.75, 2.4, C.tealSoft);
    eyebrow(s, 'Required to claim jobs', 0.65, 1.6, 5);
    const w = (5.75 - 0.4 - 0.15) / 2;
    await card(s, 0.65, 1.9, w, 1.75, { ico: 'FiCreditCard', title: 'Pay details', body: 'So the cleaner can get paid. Debit cards and bank accounts are accepted.' });
    await card(s, 0.65 + w + 0.15, 1.9, w, 1.75, { ico: 'FiFileText', title: 'SSN', body: 'Needed to run the background check that keeps customers safe.' });
    s.addImage({ data: await icon('FiArrowRight', C.soft), x: 6.3, y: 2.5, w: 0.35, h: 0.35 });
    // Right: what happens after the first claim
    box(s, 6.75, 1.45, 2.8, 2.4, C.white, C.border);
    eyebrow(s, 'After the first claim', 6.95, 1.6, 2.5);
    await iconDot(s, 6.95, 1.95, 'FiShield', 0.36);
    T(s, 'Background check', { x: 6.95, y: 2.43, w: 2.45, h: 0.3, fontFace: HEAD, bold: true, fontSize: 12 });
    T(s, 'Starts once the cleaner claims their first job, so they\'re verified before that job happens.', { x: 6.95, y: 2.75, w: 2.45, h: 0.9, fontSize: 9.5, color: C.soft });
    await tip(s, 4.1, 'Why it matters:', 'a cleaner who "finished registering" but can\'t claim is usually missing pay details or an SSN.', 'FiInfo');
  }

  {
    const s = content('Cleaner Journey  ·  Step 5', 'After the job is complete', null,
      'Once a claimed job is completed: the CP processes payment based on their selected payment method; the customer gets email/SMS confirmation and is encouraged to leave a review and tip (optional); if there is a service concern such as quality, the customer may request a refund; the CP can issue a refund themselves, either in response to a request or proactively if they know they didn\'t fully meet expectations (some CPs do this to protect their review score).');
    await cards(s, 1.4, 1.95, [
      { ico: 'FiDollarSign', title: 'Cleaner gets paid', body: 'Through the payment method they chose.' },
      { ico: 'FiStar', title: 'Review and tip', body: 'The customer gets a confirmation. Leaving a review and tip is optional.' },
      { ico: 'FiRotateCcw', title: 'Refund requests', body: 'If there\'s a service concern, the customer may ask for a refund.' },
      { ico: 'FiHeart', title: 'Cleaner refunds', body: 'Cleaners can refund on request, or proactively.' },
    ]);
  }

  // ================= 2. LEGACY CP CRM =================
  await topic('Introduction to CP CRM (Legacy)', 'Where we manage cleaners\' (CPs\') accounts.',
    ['What the CP CRM is for', 'What you\'ll find on a CP\'s page', 'A live walkthrough'],
    'ill11.png',
    'The Legacy CP CRM is where we manage CP (cleaner) accounts.');

  {
    const s = content('Legacy CP CRM', 'Where we manage cleaner accounts', 'Everything about a cleaner, in one place.',
      'The Legacy CP CRM is where we manage CP accounts. The screenshot is a preview with sample training details (Truffle Wuffle); the trainer will walk through the live system next.');
    await iconList(s, 0.45, 1.5, 3.0, [['FiUser', 'Profile, status and rate'], ['FiFlag', 'Flags and status history'], ['FiAward', 'Tier and keep rate'], ['FiBriefcase', 'Claimed and invoiced jobs'], ['FiMessageSquare', 'Messages with C and CP']], 0.5);
    await screen(s, 'cpcrm.png', 3.6, 1.5, 5.95, 3.15, 'Legacy CP CRM, preview only. Sample training details.');
    await tip(s, 4.75, 'Remember:', 'tier levels and flags are internal. Never share them with a cleaner.', 'FiLock');
  }

  await walkthrough('Legacy CP CRM', 'Legacy CP CRM', [['FiUser', 'Profile and status'], ['FiFlag', 'Flags table'], ['FiAward', 'Tier and keep rate'], ['FiMessageSquare', 'Messages']],
    'Switch to the live Legacy CP CRM now. Do not demo from the slides. Suggested flow: open a CP account, show profile status and rate, the Flags table and status history, tier and keep rate, claimed and invoiced jobs, and the message thread. Adjust to the cohort.');

  // ================= 3. CP DASHBOARD =================
  await topic('Introduction to CP Dashboard', 'Where cleaners claim and manage their jobs, and manage their account.',
    ['What the cleaner sees', 'Claiming and managing jobs', 'A live walkthrough'],
    'ill22.png',
    'The CP Dashboard is the cleaner-facing side. Seeing it helps agents guide CPs step by step.');

  {
    const s = content('CP Dashboard', 'The cleaner\'s side', 'Where cleaners claim and manage their jobs, and manage their account.',
      'The CP Dashboard is where cleaners claim and manage their jobs and manage their account. The screenshot shows claimed jobs on a calendar, Steps to Success guides, and links to update rate and availability. Customer addresses are sample data.');
    await iconList(s, 0.45, 1.5, 3.0, [['FiPlusCircle', 'Claim new jobs'], ['FiCalendar', 'Manage claimed jobs'], ['FiSliders', 'Update rate & availability'], ['FiBookOpen', 'Steps to Success guides'], ['FiUser', 'Account & Help Center']], 0.5);
    await screen(s, 'cpdash.png', 3.6, 1.5, 5.95, 3.15, 'CP Dashboard, preview only. Sample training details.');
    await tip(s, 4.75, 'Why it helps:', 'when you know what the cleaner sees, you can guide them step by step.', 'FiEye');
  }

  await walkthrough('CP Dashboard', 'CP Dashboard', [['FiPlusCircle', 'Claiming jobs'], ['FiCalendar', 'Job calendar'], ['FiSliders', 'Rate & availability'], ['FiUser', 'Account']],
    'Switch to the live CP Dashboard now. Suggested flow: claimable jobs, the claimed-jobs calendar, a job\'s details, Charge Client, updating rate and availability, and account settings.');

  // ================= 4. CP PROFILE STATUS =================
  await topic('CP Profile Status', 'The label on a cleaner\'s account that controls what they can do on the platform.',
    ['The three statuses you\'ll use most', 'Onboarding statuses', 'What the system changes automatically', 'The test before you change one'],
    'ill21.png',
    'Every time you coach or penalize a CP in a PCQ, Overcharged Hours or Reliability ticket, you are changing (or deciding not to change) their status. Understanding what each status does is what separates "I flagged the CP" from "I understood what that flag does to their ability to work." Status history is in the CP\'s Flags table: when it changed, to what, and the previous status.');

  {
    const s = content('CP Profile Status', 'The three you\'ll use most', 'What each status lets a cleaner do.',
      'active: full standing; can claim new jobs, keep upcoming claimed jobs, full dashboard. norequests: can\'t claim new jobs, keeps upcoming jobs and full dashboard. suspended: can\'t claim, removed from upcoming jobs, limited dashboard (can still see Completed Jobs to review past pay). norequests and suspended CPs can reactivate with a fee ($15-$25 by cohort) unless they have a do_not_reactivate flag or live in a CA/CAWA Operating Model state. Probation applies for the first jobs after joining and after reactivation from suspension. cp-opt-out: the CP chose not to be available (illness, moved, no longer cleaning); they can reactivate free via the dashboard or CS.');
    table(s, ['Status', 'Claim new jobs', 'Keep claimed jobs', 'Dashboard', 'Reactivate'], [
      ['active', 'Yes', 'Yes', 'Full access', '—'],
      ['norequests', 'No', 'Yes', 'Full access', '$15–25 fee*'],
      ['suspended', 'No', 'No, removed', 'Limited (Completed Jobs only)', '$15–25 fee*'],
      ['cp-opt-out', 'No', '—', '—', 'Free, via dashboard or CS'],
    ], { colW: [1.5, 1.5, 1.6, 2.4, 2.1], fontSize: 10, rowH: 0.45 });
    T(s, '* Not available if the cleaner has a do_not_reactivate flag, or is in a CA/CA-WA Operating Model state.', { x: 0.48, y: 3.85, w: 9, h: 0.25, fontSize: 9, color: C.soft, italic: true });
    await tip(s, 4.25, 'Check history:', 'the Flags table shows when a status changed, to what, and what it was before.', 'FiFlag');
  }

  {
    const s = content('CP Profile Status', 'Where to see and change it', 'In the Legacy CP CRM.',
      'Status history: the Flags table on the CP\'s CRM page shows when a status changed, what it changed to, and the previous status (e.g. cp-step-active > suspended). Changing a status: use the Change Step menu and pick the new status; norequests, suspended and cp-opt-out are grouped under one option. The trainer will show the exact clicks in the live CRM. Names and photo in the screenshots are sample training details.');
    await shot(s, 'cp_header.png', 0.45, 1.4, 4.9, 2.4, [[1, 28, 186, 792, 345]]);
    await shot(s, 'cp_changestep.png', 5.5, 1.4, 4.05, 2.4, [[2, 17, 40, 167, 353]]);
    legend(s, 0.45, 3.95, 4.9, [['Flags table', 'When the status changed, to what, and what it was before']]);
    legend(s, 5.5, 3.95, 4.05, [['Change Step', 'Pick the new status. Check the history first']], 0.55, 2);
  }

  {
    const s = content('CP Profile Status', 'Onboarding statuses', 'Where a new cleaner is stuck in sign-up.',
      'Imported: was a Homejoy CP; info imported. cpintrovideo: hasn\'t finished the intro video. cpprofile_setpay: hasn\'t set hourly rate (CPs set their own rate; a $5/hr match fee may apply to first jobs with each client in General Operating Model states; CPs can update rate anytime). cpprofilebio: hasn\'t uploaded photo/bio. cpprofilesetuppayment: hasn\'t entered legal info for the BGC or payment info (they can temporarily skip payment details to see claimable jobs). cpsignica: hasn\'t accepted the Platform Access Agreement. cphowitworks: hasn\'t done the quiz, OR was active but hasn\'t accessed the dashboard in 30+ days (auto_update_inactive_cp_step flag), OR was active but has no photo (auto_update_bad_photo_cp_step flag).');
    const items = [
      ['cpintrovideo', 'Hasn\'t finished the intro video'],
      ['cpprofile_setpay', 'Hasn\'t set their hourly rate'],
      ['cpprofilebio', 'Hasn\'t added a photo or bio'],
      ['cpprofilesetuppayment', 'Hasn\'t added legal or payment info'],
      ['cpsignica', 'Hasn\'t accepted the Platform Access Agreement'],
      ['cphowitworks', 'Quiz not done, inactive 30+ days, or no photo'],
      ['Imported', 'Former Homejoy cleaner, info imported'],
    ];
    const gap = 0.15, w = (9.1 - 3 * gap) / 4, h = 1.3;
    for (let i = 0; i < items.length; i++) {
      const x = 0.45 + (i % 4) * (w + gap), y = 1.5 + Math.floor(i / 4) * (h + gap);
      box(s, x, y, w, h, C.white, C.border);
      box(s, x + 0.15, y + 0.18, w - 0.3, 0.32, C.tealSoft);
      T(s, items[i][0], { x: x + 0.15, y: y + 0.18, w: w - 0.3, h: 0.32, fontFace: HEAD, bold: true, fontSize: 9.5, color: C.teal, align: 'center', valign: 'middle' });
      T(s, items[i][1], { x: x + 0.15, y: y + 0.62, w: w - 0.3, h: 0.6, fontSize: 9.5, color: C.ink });
    }
  }

  {
    const s = content('CP Profile Status', 'What the system changes automatically', null,
      'System-generated status changes. Keep rate: <=80% (more than 20% cancellations) norequests; <70% suspended; in probation <100% suspended. No-shows: 2 in the last 5 jobs suspended; in probation any no-show suspended. Reviews (only OCW reviews): average <4.0 across various counts, last 4 averaged 2.8, or last 6 averaged 3.0 norequests; in probation any 1- or 2-star review suspended. Reschedules: various calculations norequests. Serious risks (false invoices, theft, damage): 2 refund requests in last 10 jobs norequests; new CP 2 suspended. Cleaning supplies photo: rejected norequests, rejected 3+ suspended. Profile photo: same. BG check: info matches another CP suspended or norequests; same info as a fraud-flagged CP suspended; not clear suspended. Misc: CP rescheduled a job to before their account was created suspended. New CPs (0-2 invoiced jobs since their latest activation) are on probation with stricter thresholds.');
    table(s, ['What', 'Moves to norequests', 'Moves to suspended'], [
      ['Keep rate', '80% or lower', 'Below 70%  ·  probation: below 100%'],
      ['No-shows', '—', '2 in last 5 jobs  ·  probation: any'],
      ['Reviews (OCW)', 'Average below 4.0', 'Probation: any 1- or 2-star review'],
      ['Reschedules', 'Based on various calculations', '—'],
      ['Refund requests', '2 in last 10 jobs', 'New cleaner: 2'],
      ['Supplies / profile photo', 'Rejected', 'Rejected 3 or more times'],
      ['Background check', 'Info matches another CP', 'Matches a fraud-flagged CP, or not clear'],
      ['Other', '—', 'Rescheduled a job to before the account existed'],
    ], { y: 1.25, colW: [2.1, 3.0, 4.0], fontSize: 9.5, rowH: 0.36 });
    await tip(s, 4.6, 'Probation:', 'new cleaners (0–2 invoiced jobs since activation) face stricter thresholds.', 'FiAlertTriangle');
  }

  {
    const s = content('CP Profile Status', 'Check the numbers', 'The CP\'s CRM page shows their status and how they compare to the limits.',
      'The panel shows the CP\'s current status and rate, their review score, lifetime issues (damage, theft, false invoice, overcharged hours), and a Limit vs Actual table for the last 60 days (no-shows, reschedules, cancels, lockouts, last-minute cancels). Actual is issues divided by invoiced jobs. Use it with the flag history before deciding on a status change.');
    await shot(s, 'cp_reliability.png', 0.45, 1.4, 5.6, 2.8, [[1, 26, 95, 66, 112], [2, 26, 135, 372, 220], [3, 397, 113, 741, 277]]);
    legend(s, 6.2, 1.4, 3.35, [['Status and rate', 'The current status, e.g. active'], ['Lifetime issues', 'Damage, theft, false invoices, overcharged hours'], ['Limit vs Actual', 'Last 60 days: issues out of invoiced jobs']], 0.75);
  }

  {
    const s = content('CP Profile Status', 'Before you change a status', null,
      'Ask: "If I were the customer, would I want this CP for my cleaning?" If no, norequests or suspended is likely appropriate after proper review. A new customer with a bad first job is 50% less likely to stay on the platform, so this protects retention as well as quality. Never change a CP\'s status without checking the reason or flag history first: a status change without context is how legitimate CPs get penalized for something that wasn\'t their fault.');
    box(s, 0.45, 1.3, 9.1, 1.35, C.tealSoft);
    eyebrow(s, 'Ask yourself', 0.75, 1.5, 6);
    T(s, '"If I were the customer, would I want this cleaner for my cleaning?"', { x: 0.75, y: 1.75, w: 8.5, h: 0.7, fontFace: HEAD, bold: true, fontSize: 17 });
    const w = (9.1 - 0.15) / 2;
    await card(s, 0.45, 2.85, w, 1.6, { ico: 'FiSearch', title: 'Check the history first', body: 'Never change a status without checking the reason or flag history.' });
    await card(s, 0.45 + w + 0.15, 2.85, w, 1.6, { ico: 'FiTrendingDown', title: 'First jobs matter', body: 'A new customer with a bad first job is 50% less likely to stay.' });
  }

  // ================= 5. CP-FACING DEACTIVATION =================
  await topic('CP-Facing Deactivation Reasons', 'The exact message a cleaner sees when they are set to norequests or suspended.',
    ['Who reads the reason', 'Where you enter it', 'The reason codes and their format'],
    'ill14.png',
    'A CP-facing deactivation reason is a specific reason code plus a templated sentence. The same code is stored as the value of the deactivation flag.');

  {
    const s = content('CP-Facing Deactivation', 'One reason, two audiences', 'The reason you enter isn\'t just an internal note.',
      'When you change a CP\'s status to norequests or suspended, the reason shows on the CP\'s own dashboard and is stored in our database as the status flag\'s value, which is how we track deactivation patterns. Using the wrong code, or skipping the format, breaks both. Rule: always use the exact reason code format. Never write a freeform reason.');
    const w = (9.1 - 0.15) / 2;
    await card(s, 0.45, 1.55, w, 1.8, { ico: 'FiMonitor', title: 'The cleaner', body: 'Sees it on their dashboard as the reason they were deactivated.' });
    await card(s, 0.45 + w + 0.15, 1.55, w, 1.8, { ico: 'FiDatabase', title: 'Our data', body: 'Stored as the flag\'s value, so we can track deactivation patterns.' });
    await tip(s, 3.6, 'Rule:', 'always use the exact reason code format. Never write a freeform reason.', 'FiAlertCircle');
  }

  {
    const s = content('CP-Facing Deactivation', 'Where you enter the reason', 'The Change CP step window in the Legacy CP CRM.',
      'Top: the CP-facing reason box ("Explain to CP in 1 sentence: Why was I deactivated? & Call to Action"); the sample shows cpq_false_invoice. Middle: reference tables with templated wording by situation and codes to use (HGTA, 199, ssn_validation_failed, bgcheck_not_clear). Bottom: the three status buttons: norequests & notify CP (no new jobs, keeps existing jobs and dashboard); suspended & remove CP from claimed jobs (no new jobs, existing jobs or dashboard); cp-opt-out & remove CP from claimed jobs (also stops system messages), with notify and do_not_reactivate checkboxes. The window also shows whether the CP is eligible for timed reactivation.');
    await shot(s, 'cp_stepmodal.png', 0.45, 1.3, 3.1, 3.85, [[1, 8, 58, 616, 110], [2, 8, 120, 616, 350], [3, 8, 478, 610, 555]]);
    legend(s, 3.75, 1.3, 5.8, [['Reason box', 'Type the code and sentence the cleaner will see'], ['Reference tables', 'Templated wording and codes by situation'], ['Status buttons', 'norequests, suspended or cp-opt-out, plus notify options']], 0.8);
    box(s, 3.75, 4.05, 5.8, 0.5, C.goldSoft);
    s.addImage({ data: await icon('FiAlertCircle', C.gold), x: 3.92, y: 4.21, w: 0.18, h: 0.18 });
    T(s, [{ text: 'Rule:  ', options: { bold: true, color: C.gold } }, { text: 'use the exact code format. Never a freeform reason.' }], { x: 4.2, y: 4.15, w: 5.2, h: 0.3, fontSize: 10, valign: 'middle' });
  }

  {
    const s = content('CP-Facing Deactivation', 'Reason codes', 'Fill in the <placeholders>. Keep the rest word for word.',
      'Examples: cpq_cash Cash Payment J 1234567 / cpq_cancel Your customer Joseph reported you cancelled their Monday Oct 9 job. Submit reactivation appeal through dashboard. / cpq_false_invoice Customer Mary reported incorrect hours charged. Submit reactivation appeal. / cpq_noshow Your customer Janis reported you didn\'t show up to their Monday Oct 9 job. Submit reactivation appeal through dashboard. / cpq_overcharge Customer Kyle reported incorrect hours charged. Submit reactivation appeal. / cpq_invalid_cancel Upon review, your customer cancellation report has been deemed invalid. Submit reactivation appeal through dashboard. / cpq_reschedule Your customer George reported you rescheduled their Monday Oct 9 job without permission. Submit reactivation appeal through dashboard. / HGTA 1234567. Rule: always use the exact reason code format; never write a freeform reason.');
    table(s, ['Reason for deactivation', 'CP-facing reason format'], [
      ['Accepted cash payment', 'cpq_cash Cash Payment J <JOB ID>'],
      ['Cancellation', 'cpq_cancel Your customer <C\'S NAME> reported you cancelled their <DAY, DATE> job. Submit reactivation appeal through dashboard.'],
      ['False invoice', 'cpq_false_invoice Customer <C\'S NAME> reported incorrect hours charged. Submit reactivation appeal.'],
      ['No-show', 'cpq_noshow Your customer <C\'S NAME> reported you didn\'t show up to their <DAY, DATE> job. Submit reactivation appeal through dashboard.'],
      ['Overcharged hours', 'cpq_overcharge Customer <C\'S NAME> reported incorrect hours charged. Submit reactivation appeal.'],
      ['Invalid use of "Report Issue" to cancel', 'cpq_invalid_cancel Upon review, your <emergency/customer cancellation> report has been deemed invalid. Submit reactivation appeal through dashboard.'],
      ['Invalid use of "Report Issue" to reschedule on C\'s behalf', 'cpq_invalid_reschedule Upon review, your customer reschedule report has been deemed invalid. Submit reactivation appeal through dashboard.'],
      ['Reschedule unauthorized by the customer', 'cpq_reschedule Your customer <C\'S NAME> reported you rescheduled their <DAY, DATE> job without permission. Submit reactivation appeal through dashboard.'],
      ['Rude/unprofessional behavior', 'cpq_quality Unprofessional behavior'],
      ['Theft', 'HGTA <JOB ID>'],
    ], { y: 1.4, colW: [2.75, 6.35], fontSize: 8.5, rowH: 0.33 });
  }

  // ================= CP PENALTY: C-SIDE HANDLING =================
  await topic('CP Penalty: C-Side Handling', 'What to do on the cleaner\'s side when a customer reports a problem with their cleaner.',
    ['Your two jobs, in order', 'Why we pause instead of suspend', 'The five-step process', 'The action for each offense'],
    'ill12.png',
    'This covers what you do on the CP side when a customer reports a no-show, cancellation, false invoice, overcharged hours, missing supplies, rude behavior and similar issues. Category: Service Issue.');

  {
    const s = content('CP Penalty: C-Side Handling', 'Two jobs, in this order', null,
      'First resolve the customer\'s issue using the related article for that issue. Then take the CP-side action: pause, coach, flag and document. You don\'t make the permanent decision about the CP, and you don\'t suspend CPs yourself: that belongs to the CP-side team, Trust & Safety, or the system. You are seeing one incident, from one side; the CP-side team sees the CP\'s full history and hears the CP\'s side before deciding anything permanent.');
    const w = (9.1 - 0.15) / 2;
    await card(s, 0.45, 1.4, w, 1.85, { n: 1, title: 'Resolve the customer\'s issue', body: 'Use the related article for that issue first.' });
    await card(s, 0.45 + w + 0.15, 1.4, w, 1.85, { n: 2, title: 'Take the CP-side action', body: 'Pause, coach, flag and document.' });
    await tip(s, 3.5, 'Not your call:', 'permanent decisions and suspensions belong to the CP-side team, Trust & Safety, or the system.', 'FiInfo');
  }

  {
    const s = content('CP Penalty: C-Side Handling', 'Pause, don\'t suspend', 'You\'re seeing one incident from one side. The two statuses affect other customers very differently.',
      'norequests: the CP can\'t claim new jobs but keeps their upcoming claimed jobs. suspended: the CP is removed from every upcoming claimed job, and those customers are notified. Suspending over one complaint may cancel several other customers\' cleanings. Example: a customer reports the CP arrived without supplies. The CP has four claimed jobs this week with other customers who had no problems. norequests stops new claims and lets those four customers keep their cleanings; suspending would cancel all four.');
    await twoCol(s,
      { ico: 'FiPauseCircle', title: 'norequests (pause)', items: ['Can\'t claim new jobs', 'Keeps upcoming claimed jobs, so other customers keep their cleanings'] },
      { ico: 'FiXOctagon', title: 'suspended', items: ['Removed from every upcoming claimed job', 'Those customers are notified their cleaning is affected'] }, 1.55, 1.85);
    box(s, 0.45, 3.6, 9.1, 0.95, C.tealSoft);
    eyebrow(s, 'Example', 0.7, 3.72, 3);
    T(s, 'A customer reports their cleaner came without supplies. The cleaner has 4 other jobs this week. Pausing lets those 4 customers keep their cleanings. Suspending would cancel all 4.', { x: 0.7, y: 3.95, w: 8.6, h: 0.5, fontSize: 10.5 });
  }

  {
    const s = content('CP Penalty: C-Side Handling', 'The default, and what you don\'t do', null,
      'Default CP-side action: norequests + coaching + the relevant flag + an internal note. C-side agents don\'t suspend a CP (except the false-invoice exception), add a DNR flag, permanently deactivate a CP, or reactivate a CP.');
    box(s, 0.45, 1.3, 9.1, 1.0, C.teal);
    T(s, 'DEFAULT CP-SIDE ACTION', { x: 0.75, y: 1.43, w: 6, h: 0.2, fontFace: HEAD, bold: true, fontSize: 8, color: C.white, charSpacing: 0.5 });
    T(s, 'norequests  +  coaching  +  the relevant flag  +  an internal note', { x: 0.75, y: 1.68, w: 8.6, h: 0.45, fontFace: HEAD, bold: true, fontSize: 16, color: C.white });
    eyebrow(s, 'C-side agents don\'t', 0.48, 2.55, 5, C.soft);
    const items = [['FiXOctagon', 'Suspend a CP', 'Except the false-invoice exception'], ['FiSlash', 'Add a DNR flag', 'Never'], ['FiUserX', 'Deactivate a CP', 'Never permanently'], ['FiRotateCcw', 'Reactivate a CP', 'The CP-side team decides']];
    const gap = 0.15, w = (9.1 - 3 * gap) / 4;
    for (let i = 0; i < 4; i++) {
      const x = 0.45 + i * (w + gap);
      box(s, x, 2.85, w, 1.55, C.white, C.border);
      await iconDot(s, x + 0.18, 3.0, items[i][0], 0.4);
      T(s, items[i][1], { x: x + 0.18, y: 3.5, w: w - 0.36, h: 0.3, fontFace: HEAD, bold: true, fontSize: 11.5 });
      T(s, items[i][2], { x: x + 0.18, y: 3.82, w: w - 0.36, h: 0.45, fontSize: 9.5, color: C.soft });
    }
  }

  {
    const s = content('CP Penalty: C-Side Handling', 'Exceptions to the default', null,
      'Safety concern: theft, damage, threats, harassment, someone else doing jobs under the CP\'s account, or anything that makes the customer feel unsafe: route to Trust & Safety and follow the T&S process. System auto-suspended after your flag/report (e.g. no-show report, second cp_false_invoice on a new CP): leave it, don\'t override or downgrade; coach the CP, tell them about the penalty and how to appeal from their dashboard, add your internal note. First instance of a warning-level offense: coach only, no status change (false invoices are not warning-level). False invoice: norequests and ask the CP what happened, even when the customer says the CP never came (GPS alone can\'t confirm; CPs aren\'t required to turn it on). Multiple false invoices with an abusive pattern: you may set suspended; the only case a C-side agent sets suspended. Never add DNR.');
    table(s, ['Situation', 'What to do'], [
      ['Safety concern (theft, damage, threats, harassment, someone else working the account)', 'Route to Trust & Safety and follow their process'],
      ['The system already suspended the CP after your flag or report', 'Leave it. Coach, explain how to appeal from the dashboard, add your note'],
      ['First warning-level offense', 'Coach only. No status change'],
      ['False invoice (even "the cleaner never came")', 'norequests, and ask the CP what happened. GPS alone can\'t confirm'],
      ['Multiple false invoices that look abusive', 'You may set suspended. The only time you do. Never add DNR'],
    ], { y: 1.25, colW: [4.7, 4.4], fontSize: 9, rowH: 0.46 });
    await tip(s, 4.65, 'Safety first:', 'anything that makes the customer feel unsafe goes to Trust & Safety.', 'FiShield');
  }

  {
    const s = content('CP Penalty: C-Side Handling', 'The process', 'Five steps, in order.',
      'Step 1: Is this a safety issue? Yes: route to Trust & Safety; you\'re done on the CP side. Step 2: Did the system already suspend the CP? Check the CP\'s Issues table after adding your flag or report. Yes: leave the status, send the coaching macro, add your note; you\'re done. Step 3: find the offense in the table and take the listed action. Step 4: coach the CP. Step 5: leave an internal note.');
    const steps = [
      ['FiShield', 'Safety issue?', 'Yes: route to Trust & Safety. Done'],
      ['FiCheckSquare', 'Already suspended?', 'Check the Issues table. Yes: leave it, coach, note. Done'],
      ['FiList', 'Find the offense', 'Take the action in the table'],
      ['FiMessageSquare', 'Coach the CP', 'What was reported and how to appeal'],
      ['FiEdit3', 'Leave a note', 'So the CP-side team doesn\'t re-investigate'],
    ];
    const w = 1.66, gap = (9.1 - 5 * w) / 4;
    for (let i = 0; i < 5; i++) {
      const x = 0.45 + i * (w + gap);
      box(s, x, 1.5, w, 2.6, C.white, C.border);
      await iconDot(s, x + (w - 0.55) / 2, 1.7, steps[i][0], 0.55);
      T(s, `STEP ${i + 1}`, { x, y: 2.38, w, h: 0.2, fontFace: HEAD, bold: true, fontSize: 8, color: C.teal, align: 'center', charSpacing: 0.5 });
      T(s, steps[i][1], { x: x + 0.1, y: 2.6, w: w - 0.2, h: 0.45, fontFace: HEAD, bold: true, fontSize: 11, align: 'center' });
      T(s, steps[i][2], { x: x + 0.12, y: 3.1, w: w - 0.24, h: 1.05, fontSize: 9, color: C.soft, align: 'center' });
      if (i < 4) s.addImage({ data: await icon('FiChevronRight', C.soft), x: x + w + gap / 2 - 0.08, y: 2.7, w: 0.16, h: 0.16 });
    }
  }

  {
    const s = content('CP Penalty: C-Side Handling  ·  Step 3', 'The action for each offense', null,
      'No-show or CP cancellation: report via Job Admin > No Show, or cancel the job from the CP Dashboard when the CP should be penalized; either usually makes the system set norequests or suspended. Lean coaching; set norequests yourself only if the CP already has a couple of no-shows or cancellations. Overcharged hours: cp_overcharged_hours flag if the refund is 30+ minutes. Accepted cash: the macro adds accepting_cash_not_allowed_warning. Supplies (Gen Op): require_new_cleaning_supplies_photo sets norequests automatically. Easy to miss: unauthorized reschedule and accepted cash are coaching-only on the first instance; check message history and flags for a previous warning before pausing. Macros: No Show/Cancellation (No Request); Unauthorized Reschedule (Warning); False Invoice (With Penalty); Overcharged Hours series; Accepted Cash (Coaching) / (Warning + Account Adjustment); No/Incomplete Cleaning Supplies; Rude to Customer (Coaching).');
    table(s, ['Customer reported', 'Flag / report', 'CP status', 'Coaching'], [
      ['No-show or CP cancellation', 'Job Admin > No Show, or cancel from the CP Dashboard', 'Lean coaching. norequests only if repeat', 'No-show / cancellation macro'],
      ['Unauthorized reschedule', '—', '1st: no change · 2nd+: norequests', 'Warning macro'],
      ['False invoice (incl. "never came")', 'cp_false_invoice (may auto-suspend new CPs)', 'norequests if not auto-suspended', 'False invoice macro'],
      ['Overcharged hours', 'cp_overcharged_hours (refund 30+ min)', 'norequests', 'Overcharged Hours series'],
      ['Accepted cash', 'accepting_cash_not_allowed_warning (by macro)', '1st: no change · 2nd+: norequests', 'Accepted Cash macro'],
      ['Missing supplies (Gen Op)', 'require_new_cleaning_supplies_photo', 'Set by the flag', 'No/Incomplete Supplies macro'],
      ['Missing supplies (non-Gen Op)', '—', 'No change', 'No/Incomplete Supplies macro'],
      ['Rude / unprofessional (not safety)', '—', 'norequests', 'Coaching macro'],
    ], { y: 1.15, colW: [2.15, 2.85, 2.3, 1.8], fontSize: 8, rowH: 0.36 });
    await tip(s, 4.72, 'Easy to miss:', 'reschedules and accepted cash are coaching-only the first time. Check for a past warning.', 'FiAlertTriangle');
  }

  {
    const s = content('CP Penalty: C-Side Handling  ·  Steps 4 & 5', 'Coach, then document', null,
      'Coaching message: explain specifically what was reported; tell the CP their account is paused (if you paused it) and how to appeal through their dashboard; do not hint at an outcome ("you\'ll probably be reactivated", "this could be permanent"). The CP-side team decides. Internal note: the CP-side team uses it to decide on reactivation; a clear note saves them from investigating again. Template: CP No Request — J <JOB ID> / Reported by C: <what the customer said, one line> / Evidence: <what you verified: comms, GPS if available, CTJ, photos> / Action: <flag added> / <status set> / <macro sent> / C-side resolution: <refund, credits, rebook, etc.>');
    const w = (9.1 - 0.15) / 2;
    box(s, 0.45, 1.3, w, 3.3, C.white, C.border);
    await iconDot(s, 0.65, 1.5, 'FiMessageSquare', 0.42);
    T(s, 'Step 4: Coach the CP', { x: 1.2, y: 1.5, w: w - 0.9, h: 0.42, fontFace: HEAD, bold: true, fontSize: 13, valign: 'middle' });
    T(s, bullets(['Explain exactly what was reported', 'If you paused them, say so and explain how to appeal from their dashboard', 'Never hint at an outcome ("you\'ll probably be reactivated"). The CP-side team decides']), { x: 0.65, y: 2.1, w: w - 0.4, h: 2.3, fontSize: 10.5, paraSpaceAfter: 6 });
    const x2 = 0.45 + w + 0.15;
    box(s, x2, 1.3, w, 3.3, C.white, C.border);
    await iconDot(s, x2 + 0.2, 1.5, 'FiEdit3', 0.42);
    T(s, 'Step 5: Internal note', { x: x2 + 0.75, y: 1.5, w: w - 0.9, h: 0.42, fontFace: HEAD, bold: true, fontSize: 13, valign: 'middle' });
    box(s, x2 + 0.2, 2.1, w - 0.4, 2.3, C.bg, C.border);
    T(s, [
      { text: 'CP No Request — J <JOB ID>', options: { bold: true, breakLine: true } },
      { text: 'Reported by C: <what the customer said, one line>', options: { breakLine: true } },
      { text: 'Evidence: <comms, GPS if available, CTJ, photos>', options: { breakLine: true } },
      { text: 'Action: <flag> / <status> / <macro sent>', options: { breakLine: true } },
      { text: 'C-side resolution: <refund, credits, rebook>' },
    ], { x: x2 + 0.35, y: 2.2, w: w - 0.7, h: 2.1, fontFace: 'Courier New', fontSize: 9.5, paraSpaceAfter: 6 });
  }

  // ================= 6. OPERATING MODELS =================
  await topic('Operating Models', 'The rules that apply to a customer or cleaner, based on where they are.',
    ['The two models', 'How to tell which one applies', 'What changes between them'],
    'ill28.png',
    'Apply the wrong model and you are applying the wrong policy entirely: reactivation fees, supply requirements, hourly rates and cancellation penalties all differ.');

  {
    const s = content('Operating Models', 'Two models', 'Location decides which rules apply.',
      'CA / CA-WA Operating Model: California, New Jersey, Illinois, Massachusetts, Washington. Users in this model have a banner next to their name on the CRM. General Operating Model: all other states; no banner. If no operating model banner is present, the user is in the General Operating Model.');
    await twoCol(s,
      { ico: 'FiMapPin', title: 'CA / CA-WA Operating Model', items: ['California, New Jersey, Illinois, Massachusetts, Washington', 'Shows as a banner next to the user\'s name on the CRM'] },
      { ico: 'FiGlobe', title: 'General Operating Model', items: ['All other states', 'No banner on the CRM'] }, 1.4, 2.0);
    await tip(s, 3.65, 'Quick check:', 'no operating model banner on the CRM means General Operating Model.', 'FiEye');
  }

  {
    const s = content('Operating Models', 'How to spot it', 'Look next to the user\'s name on the CRM.',
      'Users in the CA or CA-WA Operating Model have a banner next to their name on the CRM. No banner means the General Operating Model. The sample shows a CA Operating Model cleaner.');
    await shot(s, 'cp_header_top.png', 0.45, 1.35, 9.1, 1.7, [[1, 410, 8, 538, 29]]);
    legend(s, 0.45, 3.2, 4.475, [['Operating model banner', 'CA or CA-WA Operating Model']], 0.65);
    legend(s, 5.075, 3.2, 4.475, [['No banner?', 'General Operating Model', '—']], 0.65);
    await tip(s, 4.05, 'Check first:', 'look for the banner before applying fees, penalties or pay rules.', 'FiEye');
  }

  {
    const s = content('Operating Models', 'What changes between them', null,
      'Day to day: if a CA/CAWA CP didn\'t bring cleaning supplies, coach but don\'t penalize, and don\'t apply the require_new_cleaning_supplies_photo flag. When calculating first-time-client pay, check the model: General CPs lose $5/hr on that job; CA/CAWA CPs don\'t. Whether a CP cancellation is penalized depends on the model and how much notice was given. CA/CAWA CPs must certify their business is registered and give the business name, except WA CPs, who register with the Department of Revenue and give a Universal Business Identifier.');
    table(s, ['', 'CA / CA-WA', 'General'], [
      ['Pay to reactivate', 'Disabled: no reactivation fee', 'Enabled: $15–25 fee'],
      ['Cleaning supplies photo', 'Not required. Don\'t penalize', 'Required. Missing means norequests'],
      ['CP hourly rate', 'Separate first-time and repeat-client rates', 'One rate. $5/hr match fee on first-time client jobs'],
      ['CP job cancellation', 'Can cancel within a window without penalty, depending on notice', 'Penalized. Affects keep rate'],
      ['Own cleaning business', 'Must certify it\'s registered (WA: Universal Business Identifier)', 'Not applicable'],
      ['CP Swap', 'Enabled', 'Enabled'],
    ], { y: 1.25, colW: [2.1, 3.6, 3.4], fontSize: 9.5, rowH: 0.45 });
    await tip(s, 4.5, 'Check first:', 'look for the banner before applying fees, penalties or pay rules.', 'FiAlertCircle');
  }

  // ================= 7. CP TIERING =================
  await topic('CP Tiering', 'How cleaners are ranked, and why it decides who sees and claims a job first.',
    ['Why tiering exists', 'What each tier gets', 'Requested CPs and swapping'],
    'ill29.png',
    'Important: CP Tiering is internal information only. Disclosing tier levels, tiering logic or how matching works to a CP or customer is a Zero Tolerance Policy violation.');

  {
    const s = content('CP Tiering', 'Stronger cleaners see jobs first', 'Without tiers, the first cleaner to click would get the job, whatever their record.',
      'Claimable jobs are first-come, first-served, so standing alone has little effect on who claims. Without tiering, customers could be matched with CPs with low reviews or poor reliability just because they claimed first. Tiering controls when CPs even see a job. It explains why a requested CP couldn\'t claim in time, or why a higher-tier CP swapped into a lower-tier CP\'s job. To check a CP\'s tier: click the Tier banner on CRM, or CP View > CleanerTier for full Tier History.');
    await cards(s, 1.55, 1.85, [
      { ico: 'FiBarChart2', title: 'Tier 1 to 10', body: 'A ranking based on how the cleaner has performed.' },
      { ico: 'FiEye', title: 'Earlier access', body: 'Higher tiers see jobs earlier.' },
      { ico: 'FiLayers', title: 'More jobs', body: 'Higher tiers can claim more jobs at once.' },
    ]);
    await tip(s, 3.65, 'ZTP:', 'tiering is internal. Never share tier levels or tiering logic with a cleaner or customer.', 'FiLock');
  }

  {
    const s = content('CP Tiering', 'Where to see the tier', 'On the cleaner\'s CRM page.',
      'Click the CP\'s Tier banner on the CRM, or go to CP View > CleanerTier, to see their full Tier History. Tiering is internal: never share it with a CP or customer (ZTP).');
    await shot(s, 'cp_header_top.png', 0.45, 1.35, 9.1, 1.7, [[1, 109, 30, 165, 51]]);
    legend(s, 0.45, 3.2, 4.475, [['Tier banner', 'Click it for full tier history']], 0.65);
    legend(s, 5.075, 3.2, 4.475, [['Or CP View > CleanerTier', 'Same tier history', '—']], 0.65);
    await tip(s, 4.05, 'ZTP:', 'tier levels are internal. Never share them with a cleaner or customer.', 'FiLock');
  }

  {
    const s = content('CP Tiering', 'What each tier gets', '"With requested CPs" means the customer asked for specific cleaners on that job.',
      'Requested CP: sees FC jobs once submitted; off-brand and claim limit follow their actual tier; protected from swaps at all times. Tier 10: FC jobs 48 hrs before start if requested CPs exist, otherwise once submitted; can\'t see off-brand; always protected; no limit. Tier 9: 48 hrs (requested) / 96 hrs (none); no off-brand; can\'t be swapped <48 hrs; no limit. Tier 8: 36 / 72 hrs; off-brand 24 hrs before start; protected <36 hrs; no limit. Tier 7: 24 / 48 hrs; off-brand once submitted; protected <24 hrs; max 4. Tier 6: 24 / 36 hrs; max 3. Tier 5: 24 hrs if requested CPs exist; max 2. Tiers 4 and 3: same as Tier 5; max 2. Tiers 2 and 1: same; max 1.');
    table(s, ['Tier', 'Sees FC jobs (with / without requested CPs)', 'Sees off-brand jobs', 'Swap protection', 'Claim limit'], [
      ['Requested CP', 'Once submitted', 'Per their tier', 'Always', 'Per their tier'],
      ['Tier 10', '48 hrs before / once submitted', 'Never', 'Always', 'No limit'],
      ['Tier 9', '48 hrs / 96 hrs before', 'Never', 'Under 48 hrs', 'No limit'],
      ['Tier 8', '36 hrs / 72 hrs before', '24 hrs before', 'Under 36 hrs', 'No limit'],
      ['Tier 7', '24 hrs / 48 hrs before', 'Once submitted', 'Under 24 hrs', '4'],
      ['Tier 6', '24 hrs / 36 hrs before', 'Once submitted', 'Under 24 hrs', '3'],
      ['Tiers 5 – 3', '24 hrs before (with requested CPs)', 'Once submitted', 'Under 24 hrs', '2'],
      ['Tiers 2 – 1', '24 hrs before (with requested CPs)', 'Once submitted', 'Under 24 hrs', '1'],
    ], { y: 1.4, colW: [1.35, 3.05, 1.6, 1.55, 1.55], fontSize: 9, rowH: 0.38 });
  }

  {
    const s = content('CP Tiering', 'Requested cleaners and swapping', null,
      'Requested CPs are set in Job Admin under job.customer.requested_cp. Swapping: higher-tier CPs can swap into a job claimed by a lower-tier CP if the job is more than 24-48 hours away (depending on tier). The customer sees they were matched with a different, usually better-suited cleaner; the swapped-out CP is told the customer cancelled. This framing is deliberate. Exception: a CP is protected if they have sent and received at least one non-system message with the customer in the past month. Swapping is internal: disclosing it is a ZTP violation.');
    await twoCol(s,
      { ico: 'FiUserCheck', title: 'Requested cleaners', items: ['See the job as soon as the customer submits it', 'Can claim it and replace a non-requested cleaner', 'Still follow their own tier\'s claim limit'] },
      { ico: 'FiRepeat', title: 'Swapping', items: ['A higher-tier cleaner can swap in if the job is 24–48+ hrs away', 'Customer sees a new match. The swapped-out cleaner is told the customer cancelled', 'Protected if they messaged with the customer in the past month'] }, 1.3, 2.75);
    await tip(s, 4.25, 'ZTP:', 'swapping is internal. Never tell a cleaner or customer that a swap happened.', 'FiLock');
  }

  // ================= 8. JOB STATUSES =================
  await topic('Job Statuses', 'The stage a booking is in, and what that means for the ticket in front of you.',
    ['The job lifecycle', 'Submitted, Claimed, Invoiced', 'Cancelled and Pending Invoice', 'What the customer and cleaner can do'],
    'ill27.png',
    'The same complaint ("my cleaner never showed") is handled completely differently depending on whether the job is Submitted, was Claimed and abandoned, or is in Pending Invoice. Knowing the statuses means reading the account instead of guessing.');

  {
    const s = content('Job Statuses', 'The job lifecycle', 'A job can be cancelled at almost any stage. It doesn\'t have to pass through every status.',
      'Normal path: Pending (right after the customer clicks Book; only visible via direct link) > Submitted > Claimed > Pending Invoice > Invoiced. Submitted, Claimed, Pending Invoice and Invoiced can each become Cancelled. Pending can also become Cancelled (e.g. an expired Handshake booking). Invoiced only becomes Cancelled through a full refund, not a direct status change.');
    const steps = ['Pending', 'Submitted', 'Claimed', 'Pending Invoice', 'Invoiced'];
    const w = 1.55, gap = (9.1 - 5 * w) / 4, y = 1.65;
    for (let i = 0; i < 5; i++) {
      const x = 0.45 + i * (w + gap);
      box(s, x, y, w, 0.7, i === 4 ? C.teal : C.white, C.teal);
      T(s, steps[i], { x, y, w, h: 0.7, fontFace: HEAD, bold: true, fontSize: 12, color: i === 4 ? C.white : C.teal, align: 'center', valign: 'middle' });
      if (i < 4) s.addShape(pres.shapes.LINE, { x: x + w + 0.04, y: y + 0.35, w: gap - 0.08, h: 0, line: { color: C.teal, width: 1.5, endArrowType: 'triangle' } });
      // dashed branch down to Cancelled
      const cx = x + w / 2, top = y + 0.7, cy = 3.35, tx = 4.0 + i * 0.5;
      s.addShape(pres.shapes.LINE, { x: Math.min(cx, tx), y: top, w: Math.abs(tx - cx) || 0.001, h: cy - top, flipH: tx < cx, line: { color: C.soft, width: 1, dashType: 'dash', endArrowType: 'triangle' } });
    }
    box(s, 3.75, 3.35, 2.5, 0.6, C.goldSoft, C.gold);
    T(s, 'Cancelled', { x: 3.75, y: 3.35, w: 2.5, h: 0.6, fontFace: HEAD, bold: true, fontSize: 12, color: C.gold, align: 'center', valign: 'middle' });
    await tip(s, 4.25, 'Note:', 'an Invoiced job only becomes Cancelled through a full refund.', 'FiInfo');
  }

  {
    const s = content('Job Statuses', 'Where to look in the CRM', 'On the customer\'s page in the New C CRM, jobs are grouped by status.',
      'The customer page groups jobs into panels: Pending Invoice Jobs (past their scheduled time but not yet invoiced; a count badge shows how many), Claimed & Submitted (upcoming jobs, either still waiting for a cleaner or already claimed), Completed Jobs (invoiced jobs, with type, hours requested vs invoiced, cleaner, rating and review; a Failed Charge tag shows a charge that didn\'t go through), and Cancelled Jobs. Each panel links to View all in admin and View job history in admin.');
    await shot(s, 'jobs_panels.png', 0.45, 1.35, 5.75, 3.4, [[1, 6, 5, 906, 83], [2, 6, 87, 906, 198], [3, 6, 202, 906, 397], [4, 6, 401, 906, 512]]);
    legend(s, 6.35, 1.35, 3.2, [['Pending Invoice Jobs', 'Time has passed, not invoiced yet'], ['Claimed & Submitted', 'Upcoming: waiting for a cleaner, or claimed'], ['Completed Jobs', 'Invoiced. Watch for a Failed Charge tag'], ['Cancelled Jobs', 'Cancelled at any stage']], 0.72);
  }

  {
    const s = content('Job Statuses  ·  Submitted', 'Submitted', 'The job is officially requested and sent to available cleaners nearby.',
      'Who can book: customers (dashboard), CPs for repeat clients (dashboard or the CP App Handshake feature, if the C has <2 failed charges and no known_customer_fraud flag, the CP is active, under their claim limit, rate <= $30 and no more than $5 above the C\'s last max requested rate, no deactivation flag, and they aren\'t banned), and CS via CRM. Bookings can be made up to 12 hours before start; within 48 hours a system priority fee applies (not for CRM bookings). If still unclaimed 24 hours before start, the system adds alternate start times (10 AM-2 PM that day and the next 2 days); add do_not_auto_add_alternate_start_times (C CRM > Do) if the customer objects. can_substitute may flip to True if unclaimed <=48 hrs before start and no requested CP can claim. A submitted job becomes a Priority Request within 24-48 hours of start: more notifications, and CPs with rates up to 14% above the C\'s previous max can claim. Common tickets: match anxiety, last-minute requests, gender preference (rematch >6 hrs before), specific CP requests (only change Requested CPs if the customer explicitly asks).');
    await cards(s, 1.55, 1.9, [
      { ico: 'FiUsers', title: 'Who can book', body: 'Customers, cleaners for repeat clients, and CS via the CRM.' },
      { ico: 'FiClock', title: '12 and 48 hours', body: 'Book at least 12 hrs ahead. Under 48 hrs adds a priority fee.' },
      { ico: 'FiCalendar', title: 'Alternate times', body: 'Unclaimed 24 hrs before start? The system adds alternate times.' },
      { ico: 'FiZap', title: 'Priority Request', body: 'Within 24–48 hrs of start, the job is boosted to more cleaners.' },
    ]);
    await tip(s, 3.7, 'Customer can\'t do alternate times?', 'add the do_not_auto_add_alternate_start_times flag (C CRM > Do).', 'FiFlag');
  }

  {
    const s = content('Job Statuses  ·  Claimed', 'Claimed', 'A cleaner claimed the job, and the customer gets the cleaner\'s info.',
      'Customers can reschedule via dashboard 6+ hours before start (new bookings need 12+ hours; don\'t apply the 12-hour rule to a reschedule). A priority fee may apply if the new slot is <48 hours away. Rematch: C Dashboard > Manage Appointment > Change My Cleaner, available until 6 hours before start. CP swapping can happen (internal). The CP can file a lockout (OCW: up to 24 hrs after start; CP App: 5 min before to 1 hr after start), reschedule (with the C\'s agreement, even after the job time passed), or cancel from their dashboard if no agreement. Reschedule may be unavailable if the C had 2+ undesired changes (CP reschedules, no-shows, system cancels) in the last 7 days.');
    await twoCol(s,
      { ico: 'FiUser', title: 'The customer can', items: ['Reschedule 6+ hrs before start (not the 12-hour new-booking rule)', 'Rematch: Manage Appointment > Change My Cleaner, until 6 hrs before', 'Cancel'] },
      { ico: 'FiTool', title: 'The cleaner can', items: ['File a lockout claim', 'Reschedule, if the customer agrees', 'Cancel from their dashboard if they can\'t agree'] }, 1.4, 2.55);
    await tip(s, 4.15, 'ZTP:', 'higher-tier cleaners may swap in. Swapping is internal: never mention it.', 'FiLock');
  }

  {
    const s = content('Job Statuses  ·  Invoiced', 'Invoiced', 'The cleaner clicked Charge Client. The customer is charged and the cleaner gets paid.',
      'This status occurs after the cleaner completes the job and invoices by clicking Charge Client on their dashboard. Once invoiced, the customer can leave a review and tip, or request a refund if there is an issue (covered in a separate session). The cleaner can issue a refund and leave a review for the C via the CP App. An invoiced job becomes Cancelled once a full refund is issued.');
    await twoCol(s,
      { ico: 'FiUser', title: 'The customer can', items: ['Leave a review and tip', 'Request a refund if there\'s an issue'] },
      { ico: 'FiTool', title: 'The cleaner can', items: ['Issue a refund', 'Review the customer in the CP App'] }, 1.4, 2.0);
    await tip(s, 3.65, 'Note:', 'a full refund changes an Invoiced job to Cancelled.', 'FiInfo');
  }

  {
    const s = content('Job Statuses  ·  Cancelled', 'Cancelled', 'Can happen at almost any stage. Why depends on where the job was.',
      'Customers can cancel via dashboard unless invoiced (a Last-Minute Cancellation fee may apply <6 hrs before start). CPs can cancel via dashboard or CP App unless invoiced, but it may affect their Keep Rate; no penalty if they report an emergency, a safety concern, or that the client wants to cancel via the CP App. CS can cancel via C CRM > Do > Job > Cancel: lets us waive LMC fees (job admin page) and doesn\'t penalize the CP. System cancels if unclaimed at 0, 12 or 24 hours before start, or if the C has a known_customer_fraud flag, no card on file, a declined pre-auth, requested CPs ineligible with can_substitute = False, or an unapproved Handshake booking.');
    table(s, ['Status before', 'Why it was cancelled'], [
      ['Pending', 'Booked by the cleaner with Handshake, but the customer didn\'t confirm before it expired'],
      ['Submitted', 'Cancelled by the customer, by CS in the CRM, or automatically by the system'],
      ['Claimed', 'Cleaner cancelled too late for another to claim, or the customer/CS cancelled or reported a no-show'],
      ['Pending Invoice', 'Cancelled by the customer or cleaner, or by CS in the CRM'],
      ['Invoiced', 'Fully refunded by the cleaner or CS'],
    ], { y: 1.45, colW: [1.9, 7.2], fontSize: 10, rowH: 0.42 });
    await tip(s, 4.2, 'Cancel via CRM:', 'you can waive last-minute fees, and the cleaner isn\'t penalized.', 'FiInfo');
  }

  {
    const s = content('Job Statuses  ·  Pending Invoice', 'Pending Invoice', 'The booked time has passed and the job isn\'t invoiced yet. The status changes automatically.',
      'Once the scheduled duration has elapsed from the start time and the CP hasn\'t invoiced, the job moves from Claimed to Pending Invoice. It can be invoiced by the CP (Charge Client) or the customer (reports "My cleaning was completed"). Changes are still possible: reschedule (CP only), cancel, edit details. The C can cancel (LMC fee may apply), edit duration and extras, update address, edit notes, and via Report an Issue: confirm completion, ask for ETA, cancel. The CP can edit duration before Charge Client but needs the C\'s permission to add hours; can review the C; can file a lockout (OCPW up to 24 hrs after start; CP App 5 min before to 1 hr after); can reschedule with the C\'s agreement or cancel.');
    await twoCol(s,
      { ico: 'FiUser', title: 'The customer can', items: ['Report "My cleaning was completed" (this invoices it)', 'Cancel. A last-minute fee may apply', 'Edit duration, extras, address and notes'] },
      { ico: 'FiTool', title: 'The cleaner can', items: ['Click Charge Client to invoice', 'Edit hours. Adding hours needs the customer\'s OK', 'Reschedule (cleaner only), cancel, or file a lockout'] }, 1.4, 2.55);
    await tip(s, 4.15, 'Next:', 'Pending Invoice doesn\'t mean the job was done. See the action guide.', 'FiArrowRight');
  }

  {
    const s = content('Job Statuses  ·  Pending Invoice', 'Spotting a Pending Invoice job', 'It shows in its own panel at the top of the customer\'s jobs.',
      'When a customer has a Pending Invoice job, the Pending Invoice Jobs panel shows a count badge and lists the job with a PENDING_INVOICE status, date, type, hours and cleaner. CS must act on it even when it isn\'t the customer\'s stated concern; see the AG: Pending Invoice action guide.');
    await shot(s, 'jobs_pending.png', 0.45, 1.35, 9.1, 1.45, [[1, 141, 4, 170, 26], [2, 161, 49, 257, 73]]);
    legend(s, 0.45, 3.15, 4.475, [['Count badge', 'How many Pending Invoice jobs the customer has']], 0.65);
    legend(s, 5.075, 3.15, 4.475, [['PENDING_INVOICE status', 'The job, date, hours and cleaner']], 0.65, 2);
    await tip(s, 4.0, 'Act on it:', 'a Pending Invoice job needs action, even if the customer asked about something else.', 'FiAlertCircle');
  }

  // ================= 9. AG: PENDING INVOICE =================
  await topic('AG: Pending Invoice', 'What to do with a Pending Invoice job, and how to pick the right cancellation reason.',
    ['Why Pending Invoice needs action', 'The three-step action guide', 'Agreed reschedules', 'Cancellation reason codes'],
    'ill24.png',
    'Action Guide: Pending Invoice. CS is required to take action on Pending Invoice jobs even when it isn\'t the customer\'s or cleaner\'s stated concern.');

  {
    const s = content('AG: Pending Invoice', 'Pending Invoice doesn\'t mean done', 'The status changes on a timer, not because the cleaning happened.',
      'Example: a job on January 2 at 10 AM for 3 hours changes to Pending Invoice at 1 PM automatically. It may also mean the CP couldn\'t complete the cleaning and is coordinating a reschedule, the customer requested a cancellation that hasn\'t been processed, or the job was completed but not invoiced. That is why CS must act on Pending Invoice jobs even when it isn\'t the stated concern.');
    box(s, 0.45, 1.5, 3.0, 2.3, C.tealSoft);
    eyebrow(s, 'Example', 0.7, 1.7, 2.5);
    T(s, 'Jan 2, 10 AM\n3-hour job', { x: 0.7, y: 1.95, w: 2.6, h: 0.7, fontFace: HEAD, bold: true, fontSize: 15 });
    T(s, 'Pending Invoice at 1 PM', { x: 0.7, y: 2.8, w: 2.6, h: 0.35, fontFace: HEAD, bold: true, fontSize: 12, color: C.teal });
    T(s, 'automatically', { x: 0.7, y: 3.12, w: 2.6, h: 0.3, fontSize: 10, color: C.soft });
    T(s, 'IT COULD MEAN', { x: 3.7, y: 1.5, w: 5, h: 0.2, fontFace: HEAD, bold: true, fontSize: 8, color: C.soft, charSpacing: 0.5 });
    await iconList(s, 3.6, 1.8, 5.95, [['FiCalendar', 'The cleaner couldn\'t finish and is rescheduling'], ['FiXCircle', 'The customer asked to cancel, not processed yet'], ['FiCheckCircle', 'The job was done but not invoiced']], 0.55);
    await tip(s, 4.1, 'Rule:', 'act on every Pending Invoice job, even when it isn\'t what the user asked about.', 'FiAlertCircle');
  }

  {
    const s = content('AG: Pending Invoice', 'The action guide', 'Three steps, in order.',
      'Step 1: review the C/CP comms and any C or CP messages to CS to identify whether the job was completed. Step 2: act on the evidence (next slide). Step 3: if evidence is insufficient, check whether 72 hours have passed since the job\'s end time. Yes: waive the cancellation fee so the C isn\'t charged, cancel with the appropriate reason, send C and/or CP comms. No: message the C to confirm whether the job was completed, and consider calling the C/CP. Even before 72 hours, you can cancel if comms clearly show the cleaning didn\'t happen.');
    await cards(s, 1.5, 2.2, [
      { n: 1, title: 'Review the comms', body: 'Read C and CP messages, and anything they sent to CS. Was the job completed?' },
      { n: 2, title: 'Act on the evidence', body: 'Completed, not completed, or not sure: each has its own path.' },
      { n: 3, title: 'Not sure? Check 72 hrs', body: 'Past 72 hrs since end: waive the fee, cancel, send comms. Not yet: ask the customer, consider calling.' },
    ]);
    await tip(s, 3.9, 'Clear evidence?', 'you can cancel before 72 hrs if the comms show the cleaning didn\'t happen.', 'FiInfo');
  }

  {
    const s = content('AG: Pending Invoice  ·  Step 2', 'Act on the evidence', null,
      'Completed: check whether the C messaged us confirming completion. If yes and the C confirmed a different duration: invoice in the CP Dashboard with the updated duration, tell the C we marked it completed, tell the CP we invoiced on their behalf with the duration, and coach the CP about leaving jobs pending. If yes and same duration: mark completed in the C dashboard (Select job > Report an Issue > My Cleaning was Completed). If the C hasn\'t confirmed: message the C to confirm the duration and the CP to invoice. Not completed: if the C initiated the cancellation, cancel from the C CRM. If the CP initiated, cancel from the CP dashboard to penalize the CP, or from the C CRM if no penalty should apply. Not sure: go to Step 3.');
    const cols = [
      ['FiCheckCircle', 'Completed', ['Customer confirmed a different duration: invoice in the CP Dashboard with the new hours, message both, coach the cleaner', 'Customer confirmed, same duration: Report an Issue > My Cleaning was Completed', 'Not confirmed yet: ask the customer for the duration, ask the cleaner to invoice']],
      ['FiXCircle', 'Not completed', ['Customer cancelled: cancel from the C CRM', 'Cleaner cancelled: cancel from the CP Dashboard to penalize, or the C CRM for no penalty']],
      ['FiHelpCircle', 'Not sure', ['Go to Step 3: check whether 72 hrs have passed since the job ended']],
    ];
    const gap = 0.15, w = (9.1 - 2 * gap) / 3;
    for (let i = 0; i < 3; i++) {
      const x = 0.45 + i * (w + gap);
      box(s, x, 1.2, w, 3.85, C.white, C.border);
      await iconDot(s, x + 0.18, 1.38, cols[i][0], 0.4);
      T(s, cols[i][1], { x: x + 0.7, y: 1.38, w: w - 0.85, h: 0.4, fontFace: HEAD, bold: true, fontSize: 13, valign: 'middle' });
      T(s, bullets(cols[i][2]), { x: x + 0.18, y: 1.95, w: w - 0.36, h: 3.0, fontSize: 9.5, paraSpaceAfter: 6 });
    }
  }

  {
    const s = content('AG: Pending Invoice  ·  Scenario', 'They agreed to reschedule, but the job wasn\'t changed', 'Check the comms: did they agree on a specific date and time?',
      'Yes, specific date/time agreed: proactively reschedule the job to that date/time via the CP dashboard. No specific date/time: check whether 72 hours have passed since the job end time. Not yet: send C and CP comms about the pending job with reschedule instructions. Yes: waive the cancellation fee, cancel via C CRM > Do with the reason that matches who initiated the reschedule, and tell the C the job was cancelled with rebooking instructions.');
    await twoCol(s,
      { ico: 'FiCheck', title: 'Yes, a specific date and time', items: ['Reschedule the job to that date and time in the CP Dashboard'] },
      { ico: 'FiClock', title: 'No specific date or time', items: ['Under 72 hrs since end: message both with reschedule instructions', '72+ hrs: waive the fee, cancel via C CRM > Do with the reason for whoever asked to reschedule, and send the customer rebooking steps'] }, 1.5, 2.6);
  }

  {
    const s = content('Cancellation reason codes', 'Customer reasons', 'Pick the reason that matches what your investigation found.',
      'Customer credit card failed: proactive, if we likely can\'t charge an invalid card. Failed pre-auth: explain the pre-auth first; cancel only if the C still wants to. Customer disputed credit card charge: proactive, if the C previously disputed Homeaglow charges. Customer price: disagrees with the cost. Customer wants one time: doesn\'t want a recurring schedule or only wanted a one-time. Customer doesn\'t want service: no reason given, membership-only reason, or invalid address with an unresponsive C after multiple attempts. Customer moved. Customer schedule: schedule conflict. Customer change frequency: changing RC frequency. Customer last job was lockout/LMC: proactive, inactive C with a prior LO/LMC. Customer double-booked: two cleanings on the same date or under 7 days apart. Customer wants requested CP: only wants their requested cleaner, who couldn\'t take it.');
    table(s, ['Reason code', 'Use when', 'Reason code', 'Use when'], [
      ['Customer credit card failed', 'Card is invalid, can\'t be charged', 'Customer moved', 'Moved or moving'],
      ['Failed pre-auth', 'Still wants to cancel after the pre-auth is explained', 'Customer schedule', 'A schedule conflict'],
      ['Customer disputed charge', 'Disputed Homeaglow charges before', 'Customer change frequency', 'Changing their recurring frequency'],
      ['Customer price', 'Disagrees with the cost', 'Last job lockout / LMC', 'Inactive, last job was a lockout or LMC'],
      ['Customer wants one time', 'Doesn\'t want recurring cleanings', 'Customer double-booked', 'Two cleanings under 7 days apart'],
      ['Customer doesn\'t want service', 'No reason, or unreachable with an invalid address', 'Customer wants requested CP', 'Only wants their requested cleaner'],
    ], { y: 1.45, colW: [2.1, 2.45, 2.1, 2.45], fontSize: 9, rowH: 0.48 });
  }

  {
    const s = content('Cancellation reason codes', 'Cleaner and other reasons', null,
      'CP quality: the C wants to cancel because the previous CP didn\'t meet expectations, or due to PCQ. CP no-show/reschedule: the CP needs to cancel but tried to reschedule; no C/CP comms and the C hasn\'t said why; or the CP rescheduled without consent. CP cancellation: CP-initiated with no reschedule attempts. No CP claimed: the C cancels because no cleaner claimed. Escalated safety issue. Homeaglow UI: system or dashboard issue. Voucher invalid move-out: job was a move-out clean and the C refused required extra hours. Natural disaster.');
    table(s, ['Reason code', 'Use when', 'Reason code', 'Use when'], [
      ['CP Quality', 'Previous cleaner didn\'t meet expectations (PCQ)', 'Escalated safety issue', 'Cancelled for a safety concern'],
      ['CP no-show / reschedule', 'Cleaner tried to reschedule, rescheduled without consent, or no comms', 'Homeaglow UI', 'A system or dashboard issue'],
      ['CP cancellation', 'Cleaner cancelled, no reschedule attempt', 'Voucher invalid move-out', 'Move-out clean, extra hours refused'],
      ['No CP Claimed', 'No cleaner claimed the job', 'Natural Disaster', 'Cancelled due to a natural disaster'],
    ], { y: 1.3, colW: [2.1, 2.45, 2.1, 2.45], fontSize: 9, rowH: 0.5 });
    await tip(s, 4.6, 'Remember:', 'the reason should match who initiated the cancellation and why.', 'FiInfo');
  }

  // ---------- Close ----------
  {
    const s = pres.addSlide({ masterName: 'CONTENT' });
    eyebrow(s, 'Wrap-up', 0.62, 1.9);
    T(s, 'Questions?', { x: 0.62, y: 2.15, w: 5, h: 0.75, fontFace: HEAD, bold: true, fontSize: 34 });
    T(s, 'When in doubt, check the Knowledge Library or ask your Trainer or Team Lead.', { x: 0.62, y: 2.95, w: 4.6, h: 0.5, fontSize: 12, color: C.soft });
    await panel(s, 'ill9.png');
    s.addNotes('Open the floor for questions. Recap: the cleaner journey, CP CRM and Dashboard, profile statuses, deactivation codes, operating models, tiering, job statuses and the Pending Invoice action guide.');
  }

  // ================= 10. KNOWLEDGE CHECK =================
  {
    topicNo++;
    const s = pres.addSlide({ masterName: 'CONTENT' });
    eyebrow(s, `Topic ${String(topicNo).padStart(2, '0')} of ${String(TOPICS).padStart(2, '0')}`, 0.62, 1.55, 4.5);
    T(s, 'Knowledge Check', { x: 0.62, y: 1.8, w: 4.8, h: 0.75, fontFace: HEAD, bold: true, fontSize: 30 });
    T(s, 'Time to check what you\'ve learned. Your trainer will explain the format.', { x: 0.62, y: 2.6, w: 4.6, h: 0.6, fontSize: 12, color: C.soft });
    box(s, 0.62, 3.4, 4.6, 0.55, C.goldSoft);
    s.addImage({ data: await icon('FiEdit3', C.gold), x: 0.8, y: 3.58, w: 0.2, h: 0.2 });
    T(s, [{ text: 'Trainer note:  ', options: { bold: true, color: C.gold } }, { text: 'format and questions set by the trainer.' }], { x: 1.1, y: 3.5, w: 4.0, h: 0.35, fontSize: 10, valign: 'middle' });
    await panel(s, 'ill10.png');
    s.addNotes('TRAINER NOTE: The trainer decides the knowledge check format, questions and difficulty based on the cohort.');
  }

  await pres.writeFile({ fileName: OUT });
  console.log('wrote', OUT);
})();
