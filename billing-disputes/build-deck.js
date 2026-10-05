// Builds the "Billing Disputes & Cash Payments" training deck.
// Usage: node deck.js out.pptx [imageDir]
const path = require('path');
const pptxgen = require('pptxgenjs');
const React = require('react');
const RDS = require('react-dom/server');
const sharp = require('sharp');
const Fi = require('react-icons/fi');

const OUT = process.argv[2] || 'deck.pptx';
const IMG = process.argv[3] || path.join(__dirname, '../img4');

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
pres.title = 'Billing Disputes & Cash Payments';

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
const TOPICS = 4;
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

// Common-ticket rows: number, ticket (+ optional quote), how to handle it, optional gold chip.
async function ticketRows(s, rows) {
  const h = Math.min(0.68, (3.7 - 0.07 * (rows.length - 1)) / rows.length);
  for (let i = 0; i < rows.length; i++) {
    const [title, quote, answer, chip] = rows[i], y = 1.42 + i * (h + 0.07);
    box(s, 0.45, y, 9.1, h, C.white, C.border);
    badge(s, 0.6, y + 0.2, i + 1);
    T(s, [{ text: title, options: { bold: true, fontFace: HEAD, fontSize: 10.5, breakLine: !!quote } }].concat(quote ? [{ text: quote, options: { color: C.soft, fontSize: 8.5, italic: true } }] : []),
      { x: 1.05, y, w: 2.75, h, valign: 'middle' });
    T(s, answer, { x: 3.9, y, w: chip ? 4.35 : 5.55, h, fontSize: 9.5, valign: 'middle' });
    if (chip) {
      box(s, 8.3, y + 0.2, 1.1, 0.28, C.goldSoft);
      T(s, chip, { x: 8.3, y: y + 0.2, w: 1.1, h: 0.28, fontSize: 8, bold: true, color: C.gold, align: 'center', valign: 'middle' });
    }
  }
}

// Knowledge check: question on the left, answer on the right revealed one per click (step1..step4).
async function knowledgeCheck(eb, qa, notes) {
  const s = content(eb + '  ·  Knowledge check', 'Knowledge check', 'Ask each question, then click to reveal the answer.', notes);
  const n = qa.length, h = Math.min(0.86, (3.75 - 0.08 * (n - 1)) / n);
  for (let i = 0; i < n; i++) {
    const y = 1.4 + i * (h + 0.08);
    box(s, 0.45, y, 5.05, h, C.white, C.border);
    badge(s, 0.6, y + (h - 0.3) / 2, i + 1);
    T(s, qa[i][0], { x: 1.05, y, w: 4.35, h, fontSize: 10, valign: 'middle' });
    box(s, 5.6, y, 3.95, h, C.tealSoft);
    T(s, qa[i][1], { x: 5.75, y, w: 3.7, h, fontSize: 9, valign: 'middle', objectName: `step${i + 1}Ans` });
  }
  return s;
}

// Wrap-up: three key takeaways plus the open-floor banner.
async function wrapUp(eb, takeaways, notes) {
  const s = content(eb + '  ·  Wrap-up', 'Wrap-up', 'Key takeaways before we move on.', notes);
  await cards(s, 1.45, 1.75, takeaways.map(([title, body]) => ({ title, body })));
  box(s, 0.45, 3.4, 9.1, 1.0, C.teal);
  s.addImage({ data: await icon('FiMessageCircle', C.white), x: 0.8, y: 3.7, w: 0.4, h: 0.4 });
  T(s, [{ text: 'Questions?', options: { bold: true, fontFace: HEAD, fontSize: 20, breakLine: true } }, { text: 'Open the floor before we move to the next topic.', options: { fontSize: 11 } }],
    { x: 1.45, y: 3.4, w: 7.9, h: 1.0, color: C.white, valign: 'middle' });
  return s;
}

// Horizontal numbered steps.
async function steps(s, y, h, items) {
  await cards(s, y, h, items.map(([title, body]) => ({ title, body })));
}

(async () => {
  // ---------- Cover ----------
  {
    const s = pres.addSlide({ masterName: 'COVER' });
    eyebrow(s, 'Homeaglow  ·  New Hire Care Training', 0.62, 0.6);
    T(s, 'Billing Disputes\n& Cash Payments', { x: 0.62, y: 1.0, w: 4.9, h: 1.6, fontFace: HEAD, bold: true, fontSize: 32 });
    T(s, 'When a customer is charged for time they didn\'t get, and when they pay the cleaner off the platform.', { x: 0.62, y: 2.7, w: 4.6, h: 0.7, fontSize: 12.5, color: C.soft });
    T(s, 'Internal training material', { x: 0.62, y: 5.0, w: 3, h: 0.2, fontSize: 7, color: C.soft });
    await panel(s, 'ill6.png', 0.4, 4.82);
    s.addNotes('Welcome. This module covers four topics: Overcharged Hours, Unauthorized Addition of Hours, False Invoice, and Disintermediation (Cash Payment). Each topic ends with a knowledge check and a wrap-up with time for questions.');
  }

  // ---------- Agenda ----------
  {
    const s = content('Overview', 'Agenda', 'Four topics. Each ends with a knowledge check and time for questions.',
      'The first three topics are the "unworked hours" family: they all look the same to the customer ("I was charged for time I didn\'t get"), but each has a different question, different evidence and a different resolution. The fourth topic covers customers paying cleaners off the platform.');
    const rows = [
      ['01', 'Overcharged Hours (OCH)', 'The cleaner billed more time than they worked'],
      ['02', 'Unauthorized Addition of Hours', 'The cleaner added time without the customer\'s OK'],
      ['03', 'False Invoice', 'The cleaner billed a job they never did'],
      ['04', 'Disintermediation: Cash Payment', 'The customer paid the cleaner directly'],
    ];
    const hdr = ['#', 'Topic', 'In one line'].map(t => ({ text: t, options: { bold: true, color: C.teal, fill: { color: C.tealSoft }, fontFace: HEAD } }));
    const body = rows.map(r => r.map((t, i) => ({ text: t, options: { color: i === 0 ? C.teal : i === 1 ? C.ink : C.soft, bold: i < 2, fill: { color: C.white } } })));
    s.addTable([hdr, ...body], { x: 0.45, y: 1.5, w: 9.1, colW: [0.7, 3.4, 5.0], rowH: 0.5, fontFace: BODY, fontSize: 11, valign: 'middle', border: { type: 'solid', pt: 0.75, color: C.border }, margin: [0, 0.14, 0, 0.14] });
  }

  // ================= 1. OVERCHARGED HOURS =================
  const OCH = 'Overcharged Hours';
  await topic('Overcharged Hours (OCH)', 'The cleaner charged the customer for more time than they actually worked.',
    ['How OCH differs from similar issues', 'Policies and what counts as evidence', 'Branch A and Branch B', 'Penalties and watch-outs'],
    'ill27.png',
    'Example: the customer booked 3 hours, the cleaner actually cleaned for 2, but invoiced the full 3. Category: Service Issues > Billing.');

  {
    const s = content(OCH, 'The "unworked hours" family', 'To the customer they all look the same: "I got charged for time I didn\'t get."',
      'The test isn\'t just "was there a heads-up". It\'s whether the cleaner actually worked the hours they billed. Did the cleaner show up at all? If not: False Invoice. Did they work the hours, but never get the customer\'s OK to add the extra time? Unauthorized Hours. Did they bill for more time than they worked? Overcharged Hours. Getting this right matters because each needs different evidence (GPS/CTJ data vs proof of a conversation) and leads to different resolutions and penalties.');
    await cards(s, 1.45, 1.9, [
      { ico: 'FiUserX', title: 'Didn\'t show up at all', body: 'False Invoice' },
      { ico: 'FiMessageSquare', title: 'Worked it, but never asked', body: 'Unauthorized Addition of Hours' },
      { ico: 'FiClock', title: 'Billed more than worked', body: 'Overcharged Hours: this topic' },
    ]);
    await tip(s, 3.6, 'Why it matters:', 'each one needs different evidence and leads to a different resolution and penalty.', 'FiInfo');
  }

  {
    const s = content(OCH, 'Policies', null,
      'Claim window: only jobs completed in the last 120 days. Exception: the window does not apply when the overcharge is already clearly proven when the customer reports it (CTJ, message history or job timestamps make it obvious; "likely" or "probably" is not enough). Then refund the difference via the CP Dashboard regardless of job age or the cleaner\'s status, and follow the rest of the process including the penalty ladder. Example: CTJ shows a 45-minute visit but 2 hours invoiced, reported 150 days later: refund and apply the ladder. If not clearly proven and beyond 120 days: don\'t request documentation or refund; close the ticket as unresolved instead of escalating (this includes cases where the only reason to refund is the CP\'s Suspended + DNR history). Evidence-based refund: refund only the hours the evidence supports, not automatically the full disputed amount. Suspended + DNR exemption: if the cleaner is already suspended + DNR for repeated false invoice/overcharge offenses, refund the excess hours even without clear evidence; a documented pattern removes the benefit of the doubt.');
    await cards(s, 1.4, 2.15, [
      { ico: 'FiCalendar', title: '120-day claim window', body: 'Unless the overcharge is clearly proven when reported. Then the job\'s age doesn\'t matter.' },
      { ico: 'FiSliders', title: 'Refund what\'s proven', body: 'Only the hours the evidence supports, not the full disputed amount.' },
      { ico: 'FiAlertOctagon', title: 'Suspended + DNR', body: 'Already DNR for repeat offenses? Refund the excess even without clear evidence.' },
    ]);
  }

  {
    const s = content(OCH, 'What counts as valid evidence', 'Documentation from the cleaner must meet all three.',
      'Not altered or edited: for example, a GPS screenshot must clearly show the route. If GPS was on while travelling, the screenshot shows the actual route taken; a straight line usually means the location was added manually. It clearly shows arrival and departure times. It clearly shows the cleaner was at the address for the whole job. If the documentation wasn\'t just weak but fabricated, stop and go to AG - Fraud (CP Document Fraud): invalid evidence means the claim fails; faked evidence is a separate offence with a separate outcome.');
    await cards(s, 1.45, 1.85, [
      { ico: 'FiShield', title: 'Not altered', body: 'A GPS route, not a straight line added by hand.' },
      { ico: 'FiClock', title: 'Arrival and departure', body: 'Both times are clearly shown.' },
      { ico: 'FiMapPin', title: 'On site the whole time', body: 'At the address for the entire job.' },
    ]);
    await tip(s, 3.5, 'Faked, not just weak?', 'stop and go to AG – Fraud (CP Document Fraud). It\'s a separate offence.', 'FiAlertTriangle');
  }

  {
    const s = content(OCH, 'Always start here: investigate', 'The burden of proof is on the cleaner, unless valid evidence is already in front of you.',
      'If the CTJ, the customer, or the message history already gives you something valid, decide on it; don\'t go to the cleaner for proof you already have. Evidence from the customer counts: a Ring camera clip or timestamped photo is treated the same as a CTJ record, as long as it\'s legitimate. Example: a Ring clip shows the cleaner on site from 2:13pm to 3:24pm against 2 hours invoiced: refund the difference without asking the cleaner first. Check all four sources even if one looks conclusive: a CTJ mismatch might have an innocent explanation in the messages.');
    eyebrow(s, 'Check all four sources', 0.48, 1.5, 4, C.soft);
    await iconList(s, 0.45, 1.8, 4.4, [
      ['FiNavigation', 'Cleaner\'s Travel Journey (CTJ)'],
      ['FiMessageSquare', 'C/CP messages and account activity'],
      ['FiList', 'Job History timestamps'],
      ['FiStar', 'C/CP reviews of each other'],
    ], 0.55);
    await card(s, 5.1, 1.5, 4.45, 2.35, { ico: 'FiVideo', title: 'Customer evidence counts', body: 'A Ring camera clip or a timestamped photo counts the same as a CTJ record, as long as it\'s legitimate. You don\'t need the cleaner\'s proof first.' });
  }

  {
    const s = content(OCH, 'What you find decides the branch', null,
      'Customer claims a shorter duration than what you found: if CTJ confirms an overcharge, go to Branch A and refund only what CTJ supports, even if it doesn\'t match what the customer reported. If only job history or other sources (e.g. C/CP message timestamps) suggest it, go to Branch B. Known bug (Trello): don\'t treat an abnormally short CTJ interval (30 minutes or less between ARRIVED AT and COMPLETED AT) on its own as clear evidence; it\'s inconclusive unless other evidence supports it.');
    table(s, ['What you find', 'Go to'], [
      ['CTJ shorter than the hours invoiced', 'Branch A'],
      ['Customer sends valid proof of the actual duration (Ring clip, timestamped photo)', 'Branch A: refund only what it supports'],
      ['Cleaner admits the overcharge in messages', 'Branch A'],
      ['Cleaner already refunded after the complaint', 'Branch A'],
      ['CTJ matches the hours invoiced', 'Stop: no refund. Address other concerns'],
      ['Customer\'s claim backed only by job history or message timestamps', 'Branch B'],
      ['None of the above', 'Branch B'],
    ], { y: 1.25, colW: [6.0, 3.1], fontSize: 9.5, rowH: 0.42 });
    await tip(s, 4.55, 'Known bug:', 'a very short CTJ visit (30 min or less) alone is inconclusive, not clear evidence.', 'FiAlertTriangle');
  }

  {
    const s = content(OCH, 'Branch A: clear evidence', 'A straight line: refund, check Premium, then the 30-minute rule.',
      'Step 1: refund the overcharged hour(s) via the CP Dashboard; the evidence already settles it. Step 2: if the customer was also charged a Premium rate, refund it: a confirmed overcharge is itself the service issue, no separate request needed. Step 3: check how much was overcharged. 30 minutes or more: apply the cp_overcharged_hours flag, ban the C/CP pairing, send comms to both, then go to the Penalty section. Under 30 minutes: ban the C/CP pairing, send comms to both, done (no flag, no penalty).');
    await steps(s, 1.4, 1.5, [
      ['Refund the hours', 'Via the CP Dashboard.'],
      ['Refund Premium', 'If charged. No separate request needed.'],
      ['How much?', 'The 30-minute rule decides what\'s next.'],
    ]);
    const w = (9.1 - 0.15) / 2;
    await card(s, 0.45, 3.05, w, 1.55, { ico: 'FiFlag', title: '30 min or more', body: 'cp_overcharged_hours flag, ban the pairing, comms to both, then the Penalty section.' });
    await card(s, 0.45 + w + 0.15, 3.05, w, 1.55, { ico: 'FiCheck', title: 'Under 30 min', body: 'Ban the pairing, comms to both. Done: no flag, no penalty.' });
  }

  {
    const s = content(OCH, 'Branch B: no clear evidence', 'May span more than one contact with the customer.',
      'B1: do we know how much was allegedly overcharged? If not, ask the customer for detail and address other pain points; done for this contact. B2: is the cleaner suspended + DNR for multiple overcharged hours? If yes, skip documentation and treat it as valid: go to Branch A from step 1. If not, B3: request documentation. 1) Refund the Premium Upsell fee if charged. 2) Create a Timed Reminder on the C side: 72 hours (Action On: 72 hours from handling; Who Should Act: Any CS; Action: C <C_ID> reported that CP overcharged J <JOB_ID> for <# OF OVERCHARGED HOURS>. Follow When the Timed Reminder Triggers). 3) Tell the customer we\'ll follow up in 4-5 business days and the cleaner is blocked from their future requests (not penalized yet). 4) Block the pairing and tell the cleaner documentation is due within 3 days, or the alleged hours will be refunded on their behalf. Done for this contact.');
    await steps(s, 1.4, 1.45, [
      ['Know the amount?', 'No: ask the customer, then wait.'],
      ['Suspended + DNR?', 'Yes: treat as valid, go to Branch A.'],
      ['Request proof', 'Cleaner has 3 days to send documentation.'],
    ]);
    eyebrow(s, 'When you request proof', 0.48, 3.05, 5, C.soft);
    await iconList(s, 0.45, 3.32, 4.5, [['FiDollarSign', 'Refund Premium Upsell, if charged'], ['FiBell', 'Timed Reminder on the C side: 72 hrs']], 0.5);
    await iconList(s, 5.05, 3.32, 4.5, [['FiMail', 'Tell the customer: 4–5 business days'], ['FiSlash', 'Block the pairing, ask the cleaner (3 days)']], 0.5);
  }

  {
    const s = content(OCH, 'Branch B: when the reminder triggers', 'A new contact. Two questions.',
      'Was the cleaner already advised (by a previous CS) to submit documentation? If not: send comms now, give 48 hours, set a new Timed Reminder; done for this contact. If yes: did the cleaner respond with documentation proving the hours? Yes: tell the customer no refund will be issued since the cleaner proved the hours; confirm the cleaner is blocked from future bookings with this customer. No or insufficient: refund the alleged hours only via the CP Dashboard; check/refund Premium if charged; log to the Ticket Tracker. 30 minutes or more: go to the Penalty section. Under 30: done.');
    await steps(s, 1.4, 1.5, [
      ['Already asked for proof?', 'No: send comms, give 48 hrs, set a new reminder.'],
      ['Proof received?', 'Check it meets all three evidence rules.'],
    ]);
    const w = (9.1 - 0.15) / 2;
    await card(s, 0.45, 3.05, w, 1.55, { ico: 'FiCheckCircle', title: 'Valid proof', body: 'No refund. Confirm the cleaner stays blocked from this customer.' });
    await card(s, 0.45 + w + 0.15, 3.05, w, 1.55, { ico: 'FiXCircle', title: 'No reply or not enough', body: 'Refund the alleged hours and Premium. 30 min or more: Penalty section.' });
  }

  {
    const s = content(OCH, 'Penalty section', 'Shared ending for Branch A and Branch B.',
      'Check who invoiced the job before any penalty. If the customer invoiced it, there is no overcharge penalty for the cleaner: refund the difference and close the CP side with coaching if something else warrants it. Example: the customer self-invoiced 2 hours for a job the cleaner worked for 1h11m: refund, no penalty. 30 minutes or more overcharged: cp_overcharged_hours + norequests + the Overcharged Hours macro series. Check the cleaner\'s past OCH and false invoice history and put it in your internal note; a pattern is the CP-side team\'s call. C-side agents never suspend or add DNR. CP-facing reason code (use exactly): cpq_overcharge Customer <C\'S NAME> reported incorrect hours charged. Submit reactivation appeal. Overcharged hours can also show in the cleaner\'s lifetime issue history alongside False Invoice: check both.');
    await cards(s, 1.4, 1.75, [
      { ico: 'FiSearch', title: 'Who invoiced?', body: 'Customer invoiced it? Refund, no cleaner penalty.' },
      { ico: 'FiFlag', title: '30 min or more', body: 'cp_overcharged_hours + norequests + macro series.' },
      { ico: 'FiFileText', title: 'Note the history', body: 'Past OCH and false invoices go in your internal note.' },
    ]);
    box(s, 0.45, 3.3, 9.1, 0.62, C.white, C.border);
    T(s, [{ text: 'CP-facing reason  ', options: { bold: true, color: C.teal, fontFace: HEAD } }, { text: 'cpq_overcharge Customer <C\'S NAME> reported incorrect hours charged. Submit reactivation appeal.' }], { x: 0.65, y: 3.3, w: 8.8, h: 0.62, fontSize: 10, valign: 'middle' });
    await tip(s, 4.1, 'Never:', 'C-side agents never suspend or add DNR. A pattern is the CP-side team\'s call.', 'FiLock');
  }

  {
    const s = content(OCH, 'Every refund: the checklist', 'Not a sequence. Check each one off.',
      'Applies whenever a refund is issued, in any branch. Apply the cp_overcharged_hours flag if 30 minutes or more. Check and refund the Premium fee if charged; log to the Ticket Tracker form. Block the C/CP pairing (always, unless the customer specifically wants the same cleaner). If the cleaner has a Guardian Angel/Tier 10 banner with pending action items: acknowledge the inquiry, do not clear pending items, reassign to TL Allen (allentolentino).');
    await cards(s, 1.4, 2.0, [
      { ico: 'FiFlag', title: 'Flag', body: 'cp_overcharged_hours if 30 min or more.' },
      { ico: 'FiDollarSign', title: 'Premium', body: 'Refund it if charged. Log to the Ticket Tracker.' },
      { ico: 'FiSlash', title: 'Block the pairing', body: 'Always, unless the customer wants the same cleaner.' },
      { ico: 'FiAward', title: 'GA / Tier 10 banner', body: 'Don\'t clear pending items. Reassign to TL Allen.' },
    ]);
  }

  {
    const s = content(OCH, 'Related scenarios', null,
      'Cleaner brought a helper, worked half the duration, charged in full: if the customer authorized it, tell them records show they did, address other pain points, don\'t penalize. If not, go to Branch B: request documentation and coach that helpers need permission and can\'t justify full-rate billing. Cleaner refunds after being notified: treat the refund as a sign the overcharge happened unless evidence proves otherwise; if 30 minutes or more, still apply the flag (a voluntary refund confirms the violation, it doesn\'t erase it); if it happens before the first-response Timed Reminder triggers, invalidate that reminder; tell the customer the cleaner refunded. Spotted a likely overcharge before the customer reported it? Reach out proactively. Sample: "Just checking in to see how your recent cleaning with <<Cleaner Name>> on <<Date>> went. We also wanted to quickly confirm the number of hours worked during the visit, just to make sure everything looks right on your end. Let us know if you have any questions or if anything seems off - we\'re happy to help!"');
    await cards(s, 1.4, 2.4, [
      { ico: 'FiUsers', title: 'Helper, full charge', body: 'Customer authorized it? No penalty. If not: Branch B, and coach the cleaner.' },
      { ico: 'FiRotateCcw', title: 'Cleaner refunds first', body: 'Treat it as confirmation. 30 min or more still gets the flag.' },
      { ico: 'FiEye', title: 'You spot it first', body: 'Reach out proactively to confirm the hours worked.' },
    ]);
  }

  {
    const s = content(OCH, 'Watch-outs', 'System quirks that can block a refund.',
      'Admin courtesy voucher already on the job: the CP Dashboard full refund option may not appear, because the system treats the voucher as compensation. If the voucher was already used on another job: issue an admin refund for this job and create a CP Payment Holdback. If unused: invalidate the voucher first (this reopens the CP Dashboard refund), then issue credits so the account\'s total credits equal $20. Cleaner is cp-opt-out but you need to refund from their dashboard: log in to their dashboard as admin, do NOT click Reactivate My Account, change the URL to https://homeaglow.com/cp/job/<JOB_ID>, and process the refund as usual.');
    table(s, ['Admin courtesy voucher on the job', 'What to do'], [
      ['Already used on another job', 'Admin refund this job, and create a CP Payment Holdback'],
      ['Unused', 'Invalidate it first (reopens the CP Dashboard refund), then top credits up to $20'],
    ], { y: 1.4, colW: [3.2, 5.9], fontSize: 10, rowH: 0.5 });
    await card(s, 0.45, 3.0, 9.1, 1.3, { ico: 'FiLogIn', title: 'Cleaner is cp-opt-out?', body: 'Log in to their dashboard as admin. Don\'t click Reactivate. Go to homeaglow.com/cp/job/<JOB_ID> and refund as usual.' });
  }

  await knowledgeCheck(OCH, [
    ['Booked and invoiced 3 hrs. CTJ shows the cleaner on site for 1 hr 40 min. Which branch, and what happens?', 'Branch A. Refund 1 hr 20 min via the CP Dashboard (and Premium if charged). It\'s 30 min or more: flag, ban the pairing, comms to both, Penalty section.'],
    ['No CTJ or other evidence. The customer says the cleaner left an hour early. What now?', 'Branch B. Refund Premium if charged, set a 72-hr Timed Reminder, tell the customer 4–5 business days, block the pairing, give the cleaner 3 days for proof.'],
    ['The job is 150 days old, but CTJ clearly shows 45 minutes against 2 hours invoiced. Refund?', 'Yes. Clearly proven, so the 120-day window doesn\'t apply. Refund the difference and follow the penalty ladder.'],
    ['The customer self-invoiced 2 hrs for a job the cleaner worked 1 hr 11 min. Penalize the cleaner?', 'No. Refund the difference. The customer entered the hours, so there\'s no penalty for the cleaner.'],
  ], 'Ask each question and let trainees answer before revealing. Each click reveals the next answer.');

  await wrapUp(OCH, [
    ['Find the right issue', 'Billed more than worked is OCH. Didn\'t show or didn\'t ask are different issues.'],
    ['Evidence picks the branch', 'Clear evidence: Branch A. None: Branch B and ask the cleaner for proof.'],
    ['Refund only what\'s proven', '30 min or more adds the flag and norequests. Never suspend or DNR.'],
  ], 'TRAINER: recap the three takeaways, then open the floor for questions on Overcharged Hours before moving on.');

  // ================= 2. UNAUTHORIZED ADDITION OF HOURS =================
  const UAH = 'Unauthorized Hours';
  await topic('Unauthorized Addition of Hours', 'The cleaner extended the job before invoicing, without the customer\'s OK.',
    ['How it differs from OCH', 'The consent rule', 'The four-step process', 'The helper scenario'],
    'ill24.png',
    'Cleaners are allowed to update a job\'s duration before invoicing: that part is normal platform behavior. What\'s not allowed is doing it silently. Cleaners are onboarded to always set expectations and get approval before adding time. The work itself is usually genuine, but if they never got the customer\'s OK, the extra hours are unauthorized however well the job was done.');

  {
    const s = content(UAH, 'Not the same as OCH', 'Both feel like "I was billed more than I expected." They\'re resolved differently.',
      'Overcharged Hours: the cleaner claims to have worked X hours, but the evidence shows less; a dispute over what happened. Example: booked 6 hours, finished in 4, billed the full 6. Unauthorized Hours: the cleaner genuinely worked the extra time but never checked with the customer before adding it; a dispute over whether permission was given. Example: booked 3 hours, the cleaner extended and finished after 5, then invoiced 5 hours without asking.');
    await twoCol(s,
      { ico: 'FiClock', title: 'Overcharged Hours', items: ['A dispute over what happened', 'Booked 6 hrs, worked 4, billed 6'] },
      { ico: 'FiMessageSquare', title: 'Unauthorized Hours', items: ['A dispute over whether permission was given', 'Booked 3 hrs, worked 5, billed 5 without asking'] }, 1.45, 2.0);
    await tip(s, 3.7, 'Cleaners can update the duration before invoicing.', 'Doing it without the customer\'s OK is the problem.', 'FiInfo');
  }

  {
    const s = content(UAH, 'Policies', null,
      'Consent required before extending: any additional hours the client did not authorize are treated as Unauthorized Hours and are refundable. Documentation exception: if the cleaner is already permanently suspended, skip requesting documentation; there\'s no reactivation path to justify it, so resolve directly with the customer.');
    const w = (9.1 - 0.15) / 2;
    await card(s, 0.45, 1.45, w, 2.0, { ico: 'FiCheckSquare', title: 'Consent first', body: 'Any hours the customer didn\'t authorize are refundable.' });
    await card(s, 0.45 + w + 0.15, 1.45, w, 2.0, { ico: 'FiUserX', title: 'Permanently suspended?', body: 'Skip asking the cleaner for proof. Resolve directly with the customer.' });
  }

  {
    const s = content(UAH, 'The process', 'Refund first, verify after.',
      'Step 1: refund only the alleged unauthorized hours via the cleaner\'s dashboard, not the full job; the base appointment was authorized and worked. Step 2: coach the cleaner on asking permission before extending; send comms about the action taken; give them 3 days to provide documentation proving the customer authorized the extra hours. Refunding first protects the customer\'s money without making them wait on an investigation. Step 3: if a Premium fee was charged, refund it; unauthorized hours is itself the service issue. Step 4: tell the customer a refund was processed on their cleaner\'s behalf, remind them to leave clear instructions for future cleaners, and block the C/CP pairing. Macros: Unauthorized Addition of Hours (customer refund); Unauthorized Addition of Hours > CP Coaching.');
    await steps(s, 1.4, 2.2, [
      ['Refund added hours', 'Via the CP Dashboard. Only the added time, not the full job.'],
      ['Coach the cleaner', 'Ask first, every time. They have 3 days to prove consent.'],
      ['Refund Premium', 'If charged. No separate request needed.'],
      ['Close the loop', 'Tell the customer, remind them to leave clear notes, block the pairing.'],
    ]);
    await tip(s, 3.8, 'Why refund first?', 'it protects the customer\'s money without making them wait on an investigation.', 'FiShield');
  }

  {
    const s = content(UAH, 'Scenario: the helper', 'The cleaner brought a helper, worked half the duration, and charged in full.',
      'Is there evidence the customer authorized the helper and the reduced hours? Yes: tell the customer records show they authorized it, address other pain points, don\'t penalize the cleaner. No: follow the main flow: request documentation from the cleaner and coach that helpers require permission and can\'t justify full-rate billing. If the cleaner\'s document wasn\'t just weak but faked, stop and go to AG - Fraud > CP Document Fraud.');
    const w = (9.1 - 0.15) / 2;
    await card(s, 0.45, 1.45, w, 1.9, { ico: 'FiCheckCircle', title: 'Customer authorized it', body: 'Tell them records show they did. Address other pain points. No penalty.' });
    await card(s, 0.45 + w + 0.15, 1.45, w, 1.9, { ico: 'FiXCircle', title: 'No authorization', body: 'Main flow: request documentation. Coach: helpers need permission.' });
    await tip(s, 3.55, 'Faked documents?', 'stop and go to AG – Fraud (CP Document Fraud).', 'FiAlertTriangle');
  }

  await knowledgeCheck(UAH, [
    ['Booked 3 hrs. The cleaner worked 5 hrs and invoiced 5 without asking. OCH or Unauthorized Hours?', 'Unauthorized Hours. The work was real, but the customer never gave permission.'],
    ['What do you refund?', 'Only the 2 added hours, via the CP Dashboard. Plus Premium if charged. Not the full job.'],
    ['Do you wait for the cleaner\'s proof before refunding?', 'No. Refund first. The cleaner then has 3 days to prove the customer authorized it.'],
    ['The cleaner is permanently suspended. Do you request documentation?', 'No. There\'s no reactivation path, so resolve directly with the customer.'],
  ], 'Ask each question and let trainees answer before revealing. Each click reveals the next answer.');

  await wrapUp(UAH, [
    ['Permission is the question', 'The work was real. What\'s missing is the customer\'s OK.'],
    ['Refund first, verify after', 'Only the added hours, plus Premium if charged.'],
    ['Coach and close the loop', '3 days for proof. Block the pairing.'],
  ], 'TRAINER: recap the three takeaways, then open the floor for questions on Unauthorized Addition of Hours before moving on.');

  // ================= 3. FALSE INVOICE =================
  const FI = 'False Invoice';
  await topic('False Invoice', 'The cleaner clearly did no cleaning at all, but invoiced the job and the customer was charged.',
    ['The first check: who invoiced?', 'Policies', 'Investigation and the 8 steps', 'The penalty ladder'],
    'ill14.png',
    'Distinct from Overcharged Hours (the cleaner showed up but charged more than worked): here no work happened at all. Not every falsely invoiced job is the cleaner\'s fault, so always check who invoiced first.');

  {
    const s = content(FI, 'First check: who invoiced?', 'This fork starts every False Invoice ticket.',
      'Always check the Job History for who actually invoiced the job before doing anything else. If the customer invoiced it themselves (a system flow allows this, usually by accident), issue the refund from the CP Dashboard: we still recover the pay already issued to the cleaner since no cleaning happened, but do not penalize the cleaner. Skipping this check risks penalizing a cleaner for something they didn\'t do.');
    const w = (9.1 - 0.15) / 2;
    await card(s, 0.45, 1.45, w, 2.0, { ico: 'FiUser', title: 'The customer invoiced it', body: 'Refund via the CP Dashboard. We recover the cleaner\'s pay. No penalty.' });
    await card(s, 0.45 + w + 0.15, 1.45, w, 2.0, { ico: 'FiTool', title: 'The cleaner invoiced it', body: 'A valid False Invoice. Follow the full process and penalty ladder.' });
    await tip(s, 3.65, 'Where to look:', 'the Job History shows who invoiced the job.', 'FiList');
  }

  {
    const s = content(FI, 'Policies', null,
      'Claim window: only jobs completed in the last 120 days. Exception: the limit doesn\'t apply when C/CP comms or CTJ data clearly prove the cleaner never did the cleaning but invoiced anyway ("probably" isn\'t enough). Then issue the full refund via the CP Dashboard no matter how old the job or the cleaner\'s status, and follow the penalty ladder; a late discovery doesn\'t reduce the penalty. Example: the cleaner confirmed the appointment, the customer said they weren\'t proceeding, and four days later the cleaner invoiced anyway; reported after 120 days: refund in full and penalize. First-and-only-job exception (Legacy DHJ only): once refunded, the DHJ voucher is reinstated and any MFs within 120 days are auto-refunded, since the membership shouldn\'t have started. Doesn\'t apply to FCF: FCF membership starts at signup, so only the job charge is refunded and the MF stays. Consecutive false invoices: two in a row may make the customer eligible for an ETF waiver if they have cancellation intent.');
    await cards(s, 1.4, 2.25, [
      { ico: 'FiCalendar', title: '120-day window', body: 'Unless comms or CTJ clearly prove no cleaning happened.' },
      { ico: 'FiGift', title: 'First and only job', body: 'Legacy DHJ only: voucher reinstated, MFs auto-refunded. Not FCF.' },
      { ico: 'FiRepeat', title: 'Two in a row', body: 'Customer may be eligible for an ETF waiver if they want to cancel.' },
    ]);
  }

  {
    const s = content(FI, 'Investigate', 'Gather everything before you act.',
      'Gather: customer and cleaner message histories and account activity; Job History timestamps; the cleaner\'s CTJ data; customer and cleaner reviews of each other. First and most important: who invoiced? Don\'t wait for the customer to report it: if something looks off (duration, charges, or details that don\'t match the booking), reach out proactively to check if the cleaning happened and the invoice is accurate.');
    await iconList(s, 0.45, 1.45, 4.4, [
      ['FiMessageSquare', 'C/CP messages and account activity'],
      ['FiList', 'Job History timestamps'],
      ['FiNavigation', 'Cleaner\'s Travel Journey (CTJ)'],
      ['FiStar', 'C/CP reviews of each other'],
    ], 0.55);
    await card(s, 5.1, 1.45, 4.45, 2.2, { ico: 'FiEye', title: 'Spot it first', body: 'Duration or charges don\'t match the booking? Reach out to the customer before they complain.' });
  }

  {
    const s = content(FI, 'Confirmed (cleaner-caused): 8 steps', null,
      '1) Issue a full refund via the CP Dashboard. 2) Check the cleaner\'s Flags table for an existing cp_false_invoice flag on this job. 3) If the system already flagged it, validate that the Job ID matches; on the flag\'s Django page set Value to {"admin": "<YOUR CRM NAME>", "admin_flag_comments": "<THE ID TRIPLET>", "cp_false_invoice_job_id": <THE JOB ID>}. If no flag exists, add one: CP CRM > Do > Flag > cp_false_invoice. Always include the ID triplet (Customer ID | CP ID | Job ID) and a short summary. 4) Check the Issues Table for prior false invoice / overcharge history and penalize accordingly. Reason: cpq_false_invoice Customer <C\'S NAME> reported incorrect hours charged. Submit reactivation appeal. 5) Coach the cleaner: don\'t claim jobs you can\'t complete or invoice without working; ask the customer to reschedule, or cancel from your dashboard if they disagree. Guardian Angel or Tier 10 banner with pending items: acknowledge, don\'t clear, reassign to allentolentino. 6) Refund the Premium fee if charged (AG - Refund: Premium Upsell). 7) Block the C/CP pairing unless the customer wants the same cleaner. 8) Tell the customer the refund was issued and arrives in 5-10 business days; address other pain points. Macros: False Invoice (C-facing); False Invoice (With Penalty) (CP-facing).');
    const items = [
      ['Full refund', 'Via the CP Dashboard'], ['Check the flag', 'cp_false_invoice on this job?'], ['Validate or add it', 'Include the ID triplet'], ['Penalize', 'Check the Issues Table first'],
      ['Coach the cleaner', 'Reschedule or cancel, never invoice'], ['Refund Premium', 'If charged'], ['Block the pairing', 'Unless the customer wants them'], ['Tell the customer', 'Refund in 5–10 business days'],
    ];
    const w = (9.1 - 0.45) / 4;
    for (let i = 0; i < 8; i++) {
      await card(s, 0.45 + (i % 4) * (w + 0.15), 1.25 + Math.floor(i / 4) * 1.9, w, 1.75, { n: i + 1, title: items[i][0], body: items[i][1] });
    }
  }

  {
    const s = content(FI, 'The penalty ladder', 'C-side agents never add DNR.',
      'Any false invoice, first or repeat, including "the cleaner never came": set norequests and ask the cleaner what happened (the False Invoice (With Penalty) macro handles this). Already multiple false invoices and the pattern looks abusive: you may set the cleaner to suspended; never add DNR. The system auto-suspended the cleaner (new cleaners are auto-suspended at 2+ cp_false_invoice flags): leave it, coach the cleaner and tell them about the penalty. Reactivation: C-side agents don\'t reactivate cleaners. If the cleaner\'s explanation matches one of these, put it in your internal note for the CP-side team: they accidentally invoiced the wrong customer; they invoiced a Lockout job by mistake; it\'s not their first false invoice, but the last one was more than a year ago. If the cleaner is already permanently suspended, skip requesting documentation.');
    await cards(s, 1.4, 1.95, [
      { ico: 'FiPauseCircle', title: 'Any false invoice', body: 'norequests, and ask what happened.' },
      { ico: 'FiAlertOctagon', title: 'Repeated and abusive', body: 'You may suspend. Never add DNR.' },
      { ico: 'FiCpu', title: 'System auto-suspended', body: 'Leave it. Coach and explain the penalty.' },
    ]);
    await tip(s, 3.55, 'Reactivation is the CP-side team\'s call.', 'Note any wrong-customer or Lockout mix-up in your internal note.', 'FiFileText');
  }

  await knowledgeCheck(FI, [
    ['What\'s the first thing you check on a False Invoice ticket?', 'Who invoiced the job: check the Job History.'],
    ['The customer invoiced it by accident. What do you do?', 'Refund via the CP Dashboard. We recover the cleaner\'s pay, but don\'t penalize the cleaner.'],
    ['It was a Legacy DHJ customer\'s first and only job. What happens after the refund?', 'The DHJ voucher is reinstated and MFs within 120 days are auto-refunded. On FCF, only the job is refunded.'],
    ['It\'s the cleaner\'s first false invoice. What status do you set?', 'norequests, using the False Invoice (With Penalty) macro. Never add DNR.'],
  ], 'Ask each question and let trainees answer before revealing. Each click reveals the next answer.');

  await wrapUp(FI, [
    ['Who invoiced, first', 'Customer invoiced it? Refund, no penalty.'],
    ['Full refund, then flag', 'cp_false_invoice with the ID triplet. Block the pairing.'],
    ['norequests, never DNR', 'Repeated abuse may be suspended. Reactivation is CP-side.'],
  ], 'TRAINER: recap the three takeaways, then open the floor for questions on False Invoice before moving on.');

  // ================= 4. CASH PAYMENT =================
  const CASH = 'Cash Payment';
  await topic('Disintermediation: Cash Payment', 'The customer paid the cleaner directly instead of through the platform.',
    ['What it is and how it\'s reported', 'Key principles', 'The penalty ladder', 'CP-side and C-side paths'],
    'ill_pay.png',
    'Cash is the most common form, but Venmo, Zelle, Cash App and PayPal count too.');

  {
    const s = content(CASH, 'What it looks like', 'Cash, Venmo, Zelle, Cash App or PayPal: any payment outside the platform.',
      'Cleaners can no longer tag a job as a cash job. If a cleaner receives cash, they use Report Issue to tell us. That creates a ticket for a CP-side agent; it does not leave any flag on the job. The job stays at Pending Invoice, which keeps the customer from being auto-charged. Two starting points: a CP-side agent gets the ticket directly (the expected path), or a C-side agent hears it from the customer first, either because the cleaner\'s report hasn\'t been worked yet, or the cleaner never reported it (job still Pending Invoice, or already invoiced). C-side: always check the CP CRM before assuming nothing was reported.');
    await steps(s, 1.4, 1.6, [
      ['Cleaner uses Report Issue', 'They tell us they received cash.'],
      ['A CP-side ticket is created', 'No flag is left on the job itself.'],
      ['Job stays Pending Invoice', 'So the customer isn\'t auto-charged.'],
    ]);
    await tip(s, 3.25, 'C-side agent?', 'always check the CP CRM first. The cleaner may already have reported it.', 'FiSearch');
  }

  {
    const s = content(CASH, 'The job status decides the action', 'Nobody should pay twice, and nobody should be paid twice.',
      'The customer shouldn\'t pay twice (cash and the platform) and the cleaner shouldn\'t be paid twice (cash and their payout). Pending Invoice: nothing has been charged; cancel via the C CRM with the "does not want the service" reason code so it never gets charged. Invoiced: refund from the CP Dashboard, which reverses both the customer\'s charge and the cleaner\'s payout in one action. Cancelled (the customer says the cleaning happened and they paid cash): nothing was ever charged, so there is nothing to refund. Do not recreate the job: that\'s how the customer ends up paying twice. Exception: the cleaner tells us separately they were never paid; then work it as a normal cash case from the job\'s actual status. Otherwise educate on the cash policy, cancel the RC and upcoming jobs if the customer no longer wants service, and close.');
    await cards(s, 1.4, 2.0, [
      { ico: 'FiClock', title: 'Pending Invoice', body: 'Cancel via the C CRM: "does not want the service".' },
      { ico: 'FiDollarSign', title: 'Invoiced', body: 'Refund from the CP Dashboard. Reverses both sides.' },
      { ico: 'FiXCircle', title: 'Cancelled', body: 'Nothing to refund. Don\'t recreate the job.' },
    ]);
    await tip(s, 3.6, 'Partial cash payment?', 'invoice (or refund) only the difference, based on CP Pay, not the C Price.', 'FiPercent');
  }

  {
    const s = content(CASH, 'Key principles', null,
      'Don\'t penalize customers for paying off-platform: they\'re not breaking the rules, the cleaner is. Don\'t act as an investigator when there\'s nothing to go on: with no cleaner report and no proof from the customer, focus on the customer\'s experience. How the cash payment was confirmed changes the tone of the coaching, not the cleaner\'s status. Remind the customer that payment runs through the platform, not cash, unless it\'s a tip: many don\'t realise the job still has to be invoiced, and a cash payment today can still turn into a charge later.');
    await cards(s, 1.4, 2.25, [
      { ico: 'FiHeart', title: 'Not the customer', body: 'The cleaner broke the rule, not them.' },
      { ico: 'FiSearch', title: 'Nothing to go on?', body: 'Nothing to go on? Focus on the customer\'s experience.' },
      { ico: 'FiMessageSquare', title: 'Tone, not status', body: 'How it was confirmed changes the coaching tone.' },
      { ico: 'FiCreditCard', title: 'Remind the customer', body: 'Pay through the platform. Cash is only for tips.' },
    ]);
  }

  {
    const s = content(CASH, 'The penalty ladder', 'C-side agents never suspend or add DNR for cash.',
      'Quick reference from the KB. Cleaner admits it (soft ladder): 1st instance coach, no status change; 2nd instance norequests. Confirmed only by customer proof (hard ladder): 1st instance coach and warn, no status change; 2nd instance norequests. Repeated instances are the CP-side team\'s call: put the history in your internal note. If the cleaner denies it and the customer never produces proof, nothing is confirmed: no refund, no penalty; assess for retention instead. NOTE FOR TRAINER: the KB process steps further down still say a 3rd instance (soft) and 2nd instance (hard) lead to Suspended + DNR, which conflicts with this quick reference and the key principle. Confirm the current rule before teaching.');
    table(s, ['Instance', 'Cleaner admits it (soft ladder)', 'Only customer proof (hard ladder)'], [
      ['1st', 'Coach. No status change', 'Coach and warn. No status change'],
      ['2nd', 'norequests', 'norequests'],
    ], { y: 1.45, colW: [1.4, 3.85, 3.85], fontSize: 11, rowH: 0.55 });
    await tip(s, 3.35, 'Denied and no proof?', 'nothing is confirmed: no refund, no penalty. Assess for retention.', 'FiInfo');
  }

  {
    const s = content(CASH, 'Branch A: the cleaner reported it', 'CP-side agent. This is the path most cases take.',
      '1) The cleaner reports cash via Report Issue and the ticket lands in your queue. 2) Did they say how much they received? If not, pause: don\'t act on the job yet; contact the cleaner and get a number, or a clear "I was paid in full". 3) Check the job status: Pending Invoice: cancel via the C CRM ("does not want the service"). Invoiced: refund from the CP Dashboard. Cancelled: nothing to refund; don\'t recreate the job. Partial cash payment with the job still Pending Invoice: invoice on the cleaner\'s behalf for the remaining balance. Full vs partial is based on CP Pay, not the full C Price. 4) Apply the soft ladder (a cleaner admission). Done: no further customer action unless they reach out.');
    await steps(s, 1.4, 2.0, [
      ['Ticket arrives', 'The cleaner used Report Issue.'],
      ['Get the amount', 'No amount? Pause and ask the cleaner first.'],
      ['Act on the status', 'Cancel, refund, or nothing if already cancelled.'],
      ['Soft ladder', 'It\'s an admission: coach first.'],
    ]);
    await tip(s, 3.6, 'Partial payment, still Pending Invoice?', 'invoice the remaining balance, based on CP Pay.', 'FiPercent');
  }

  {
    const s = content(CASH, 'Branch B: the customer told us first', 'C-side agent. Check the CP CRM before anything else.',
      'Report found, not yet actioned: resolve it yourself, no need to wait for the CP-side queue. Same amount gate as Branch A. Pending Invoice: cancel via the C CRM ("does not want the service"); if partial, invoice the remaining balance. Invoiced: refund via the CP Dashboard; if partial, admin refund the customer for the cash-paid portion (based on CP Pay) and apply a CP Holdback for the same amount. Soft ladder. Nothing on file: ask the customer for proof (encourage, don\'t require). Call the cleaner within business hours (8AM-8PM their local time); no answer, send an SMS; outside business hours, email. Set a 48-hour Timed Reminder. Customer provides proof: resolve per the job status, hard ladder. No proof but the cleaner admits it: resolve per the job status, soft ladder. Cleaner denies and no proof: no refund, no penalty; assess for retention: high-value or retention-likely (multiple completed jobs, high LTNR): voucher covering the full job hours; otherwise $20-$50 in platform credits; Internal Reason: cash_payment. Macros: Accepted Cash (Coaching); Accepted Cash (Warning + Account Adjustment); Refund confirmation: credit card refund processed.');
    const w = (9.1 - 0.15) / 2;
    await card(s, 0.45, 1.4, w, 1.45, { ico: 'FiFileText', title: 'Report on file', body: 'Resolve it yourself. Get the amount, act on the status, soft ladder.' });
    await card(s, 0.45 + w + 0.15, 1.4, w, 1.45, { ico: 'FiPhoneCall', title: 'Nothing on file', body: 'Ask the customer for proof. Call or text the cleaner. 48-hr Timed Reminder.' });
    table(s, ['Outcome', 'What to do'], [
      ['Customer provides proof', 'Resolve per job status. Hard ladder'],
      ['Cleaner admits it', 'Resolve per job status. Soft ladder'],
      ['Cleaner denies, no proof', 'No refund, no penalty. Retention: voucher (high value) or $20–$50 credits'],
    ], { y: 3.0, colW: [2.8, 6.3], fontSize: 9.5, rowH: 0.4 });
  }

  await knowledgeCheck(CASH, [
    ['The cleaner reported full cash payment. The job is still Pending Invoice. What do you do?', 'Cancel via the C CRM with "does not want the service". Nothing was charged, so nothing to refund.'],
    ['Same case, but the job was already invoiced.', 'Refund from the CP Dashboard. It reverses the customer\'s charge and the cleaner\'s payout.'],
    ['The job shows Cancelled, but the customer says they paid cash. Recreate the job?', 'No. Nothing was charged. Recreating it would make the customer pay twice.'],
    ['The cleaner denies it and the customer has no proof. What now?', 'No refund, no penalty. Assess retention: voucher (high value) or $20–$50 credits, reason cash_payment.'],
  ], 'Ask each question and let trainees answer before revealing. Each click reveals the next answer.');

  await wrapUp(CASH, [
    ['Status decides the action', 'Pending Invoice: cancel. Invoiced: refund. Cancelled: don\'t recreate.'],
    ['Get the amount first', 'Full vs partial is based on CP Pay, not the C Price.'],
    ['Never penalize the customer', 'Coach the cleaner. C-side never suspends or adds DNR.'],
  ], 'TRAINER: recap the three takeaways, then open the floor for questions on Cash Payment.');

  // ---------- Close ----------
  {
    const s = pres.addSlide({ masterName: 'CONTENT' });
    eyebrow(s, 'Wrap-up', 0.62, 1.9);
    T(s, 'Thank you!', { x: 0.62, y: 2.15, w: 5, h: 0.75, fontFace: HEAD, bold: true, fontSize: 34 });
    T(s, 'When in doubt, check the Knowledge Library or ask your Trainer or Team Lead.', { x: 0.62, y: 2.95, w: 4.6, h: 0.5, fontSize: 12, color: C.soft });
    await panel(s, 'ill9.png');
    s.addNotes('Close the module. Recap: OCH (billed more than worked), Unauthorized Hours (worked but never asked), False Invoice (never showed, check who invoiced first), and Cash Payment (job status decides the action).');
  }

  await pres.writeFile({ fileName: OUT });
  console.log('wrote', OUT);
})();
