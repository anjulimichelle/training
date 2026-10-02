// Builds the "Platform & Operating Mechanics" training deck.
// Usage: node deck.js out.pptx [imageDir]
const path = require('path');
const pptxgen = require('pptxgenjs');
const React = require('react');
const RDS = require('react-dom/server');
const sharp = require('sharp');
const Fi = require('react-icons/fi');

const OUT = process.argv[2] || 'deck.pptx';
const IMG = process.argv[3] || path.join(__dirname, '../img3');

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
pres.title = 'Platform & Operating Mechanics';

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
  eyebrow(s, `Topic ${String(topicNo).padStart(2, "0")}`, 0.62, 1.32, 4.5);
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
    T(s, 'Platform &\nOperating Mechanics', { x: 0.62, y: 1.0, w: 4.9, h: 1.6, fontFace: HEAD, bold: true, fontSize: 32 });
    T(s, 'How jobs, callbacks, pricing, vouchers, unverified tickets and rate caps work behind the scenes.', { x: 0.62, y: 2.7, w: 4.6, h: 0.7, fontSize: 12.5, color: C.soft });
    T(s, 'Internal training material', { x: 0.62, y: 5.0, w: 3, h: 0.2, fontSize: 7, color: C.soft });
    await panel(s, 'ill6.png', 0.4, 4.82);
    s.addNotes('Welcome. This module covers Understanding Job History, Callback Handling, C Price vs CP Pay, Vouchers in CRM, Non-Logged-In Visitor Tickets, and Max CP Rate. More topics will be added to this deck.');
  }

  // ================= UNDERSTANDING JOB HISTORY =================
  await topic('Understanding Job History', 'The record of every change made to a job, and what caused it.',
    ['Why you always check it', 'Where to find it', 'How to read the columns', 'Practice reading one'],
    'ill21.png',
    'There are times we can\'t tell the full story of what happened to a job from the C/CP comms alone. The job history records every change made to the job, so we can understand how its status changed and what caused it.');

  {
    const s = content('Understanding Job History', 'Why you always check it', null,
      'There are instances where we cannot tell the full story of what happened to the job if we solely refer to the C/CP comms. The job history records any changes done within the job, so we can understand how the job status changed and what caused it. Check the job history whenever you\'re investigating a particular job, so you give the correct resolution.');
    await cards(s, 1.4, 1.95, [
      { ico: 'FiMessageSquare', title: 'Comms aren\'t enough', body: 'C and CP messages don\'t always tell the full story of a job.' },
      { ico: 'FiList', title: 'Every change is logged', body: 'How the status changed, and what caused it.' },
      { ico: 'FiSearch', title: 'Check it every time', body: 'Whenever you investigate a job, before you decide on a resolution.' },
    ]);
    await tip(s, 3.6, 'Rule:', 'investigating a job? Open its job history before you resolve the ticket.', 'FiAlertCircle');
  }

  {
    const s = content('Understanding Job History', 'Where to find it', 'Click the ⓘ icon next to a job ID to open its job history.',
      'On the customer\'s page in the New C CRM, every job in the jobs panels (Claimed & Submitted, Completed Jobs and so on) has a small ⓘ icon next to the job ID. Clicking it opens that job\'s history. Names in the screenshot are sample training details.');
    await shot(s, 'jh_open.png', 0.45, 1.4, 6.1, 2.75);
    legend(s, 6.7, 1.4, 2.85, [['The ⓘ icon', 'Next to each job ID, in every jobs panel', 'i'], ['Click it', 'The job\'s history opens', '→']], 0.85);
    await tip(s, 4.35, 'Tip:', 'investigating a job? Open its history before you decide on a resolution.', 'FiInfo');
  }

  {
    const s = content('Understanding Job History', 'The job history', 'Click a number to reveal what that column tells you. Read the rows from the bottom up.',
      'TRAINER: each click (on a number or anywhere on the slide) reveals the next column description, 1 to 7, while the job history stays on screen. 1 Job Start: the job start time and date at the time the action was made. 2 Job Start After Action: the job start time after the action was made. 3 Action On: the time the action was taken. 4 Job Status Before Action: the status of the job before the action was taken. 5 Actor: the one who took the action (it can be the C, CP, CS, or system). 6 Action: the action taken on the job; actions are written in a very self-explanatory manner. 7 Comment: logs any essential detail to the action. The first column (ID) is the action ID, mostly used by product or support to edit the action in Django; CS doesn\'t use it. Names are sample training details.');
    const file = 'jh_table_hdr.png';
    box(s, 0.45, 1.3, 5.95, 3.85, C.white, C.border);
    const r = await fitImage(s, file, 0.53, 1.38, 5.79, 3.69);
    const { width } = await sharp(path.join(IMG, file)).metadata();
    const k = r.w / width;
    const cols = [[92, 212], [219, 333], [342, 400], [462, 588], [594, 634], [641, 700], [1141, 1196]];
    const hy0 = 62, hy1 = 81, d = 0.2;
    cols.forEach(([x0, x1], i) => {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: r.x + x0 * k, y: r.y + hy0 * k, w: (x1 - x0) * k, h: (hy1 - hy0) * k, fill: { type: 'none' }, line: { color: C.teal, width: 1.5 }, rectRadius: 0.02 });
      const cx = r.x + ((x0 + x1) / 2) * k;
      s.addShape(pres.shapes.OVAL, { x: cx - d / 2, y: r.y + hy0 * k - d - 0.03, w: d, h: d, fill: { color: C.teal }, line: { color: C.white, width: 1 } });
      T(s, String(i + 1), { x: cx - d / 2, y: r.y + hy0 * k - d - 0.03, w: d, h: d, fontFace: HEAD, bold: true, fontSize: 7.5, color: C.white, align: 'center', valign: 'middle' });
    });
    // Descriptions panel: empty until clicked; each click reveals the next line
    box(s, 6.55, 1.3, 3.0, 3.85, C.tealSoft);
    const items = [
      ['Job Start', 'the job start time and date at the time the action was made'],
      ['Job Start After Action', 'the job start time after the action was made'],
      ['Action On', 'the time the action was taken'],
      ['Job Status Before Action', 'the status of the job before the action was taken'],
      ['Actor', 'the one who took the action: the C, CP, CS, or system'],
      ['Action', 'the action taken on the job. Written to be self-explanatory'],
      ['Comment', 'logs any essential detail to the action'],
    ];
    const runs = [];
    items.forEach(([name, desc], i) => {
      runs.push({ text: `${i + 1}  ${name}: `, options: { bold: true, color: C.teal } });
      runs.push({ text: desc, options: { breakLine: i < items.length - 1 } });
    });
    T(s, runs, { x: 6.68, y: 1.42, w: 2.75, h: 3.6, fontSize: 8.5, paraSpaceAfter: 4, objectName: 'revealColumns' });
  }

  {
    const s = content('Understanding Job History', 'Two things to remember', null,
      'Read the job history from bottom to top. Dates and times are in CT (Central Time): our system time is CT, and customers may be in a different time zone. The time formula in the upper-left corner of the CRM helps you convert (e.g. PDT -2, MDT -1, CDT 0, EDT +1).');
    const w = (9.1 - 0.15) / 2;
    await card(s, 0.45, 1.35, w, 1.75, { ico: 'FiArrowUp', title: 'Read bottom to top', body: 'The oldest action is at the bottom. The newest is at the top.' });
    await card(s, 0.45 + w + 0.15, 1.35, w, 1.75, { ico: 'FiClock', title: 'Times are in CT', body: 'Our system time is Central Time. The customer may be in another time zone.' });
    T(s, 'CONVERT WITH THE TIME FORMULA, UPPER LEFT OF THE CRM', { x: 0.48, y: 3.3, w: 9, h: 0.2, fontFace: HEAD, bold: true, fontSize: 8, color: C.soft, charSpacing: 0.5 });
    await shot(s, 'legacy_time.png', 0.45, 3.55, 4.0, 1.05, [[1, 30, 1, 210, 34]]);
    T(s, [{ text: 'Example: ', options: { bold: true } }, { text: '8:23 AM CDT is 6:23 AM in Pacific time (PDT −2) and 9:23 AM in Eastern time (EDT +1).' }], { x: 4.7, y: 3.65, w: 4.85, h: 0.85, fontSize: 10.5, valign: 'middle' });
  }

  {
    const s = content('Understanding Job History', 'Reading the example', 'Read the job history from the bottom up. Then click The Story.',
      'TRAINER: let trainees read the job history first (bottom to top). Then click The Story (or press the right arrow) to reveal the story one line at a time, staying on this slide. Story: On Sept 7, 2026, the system booked a cleaning for C Anjuli Kintanar scheduled for Sept 21 10AM CT through the recurring cleaning plan. On Sept 14, CP Truffle Wuffle claimed the job as scheduled. On Sept 21 9:19AM CT, CP Truffle Wuffle rescheduled the appointment from Sept 21, 10AM CT to Sept 22, 11AM. On Sept 21 9:50 AM CT, CP Truffle Wuffle rescheduled the appointment again from Sept 22, 11AM to Sept 24, 12:30 PM CT. On Sept 24, 4:09 PM CT, CP Truffle Wuffle invoiced the appointment and the system logged it in snapshot.');
    await shot(s, 'jh_table.png', 0.45, 1.3, 5.75, 2.9);
    s.addImage({ data: await icon('FiArrowUp'), x: 0.55, y: 4.3, w: 0.2, h: 0.2 });
    T(s, 'Read bottom to top', { x: 0.8, y: 4.27, w: 3, h: 0.26, fontSize: 9.5, bold: true, color: C.teal, valign: 'middle' });
    // "The Story" button
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.35, y: 1.3, w: 3.2, h: 0.5, fill: { color: C.teal }, line: { color: C.teal }, rectRadius: 0.1 });
    s.addImage({ data: await icon('FiBookOpen', C.white), x: 6.55, y: 1.43, w: 0.24, h: 0.24 });
    T(s, 'The Story', { x: 6.9, y: 1.3, w: 2.5, h: 0.5, fontFace: HEAD, bold: true, fontSize: 13, color: C.white, valign: 'middle' });
    // Story area: empty panel; the text appears line by line on click
    box(s, 6.35, 1.9, 3.2, 3.25, C.tealSoft);
    const lines = [
      ['Sept 7, 2026: ', 'the system booked a cleaning for C Anjuli Kintanar for Sept 21, 10 AM CT, through the recurring cleaning plan.'],
      ['Sept 14: ', 'CP Truffle Wuffle claimed the job as scheduled.'],
      ['Sept 21, 9:19 AM CT: ', 'CP Truffle Wuffle rescheduled from Sept 21, 10 AM CT to Sept 22, 11 AM.'],
      ['Sept 21, 9:50 AM CT: ', 'CP Truffle Wuffle rescheduled again, from Sept 22, 11 AM to Sept 24, 12:30 PM CT.'],
      ['Sept 24, 4:09 PM CT: ', 'CP Truffle Wuffle invoiced the appointment and the system logged it in snapshot.'],
    ];
    const runs = [];
    lines.forEach(([d, t], i) => {
      runs.push({ text: d, options: { bold: true, color: C.teal } });
      runs.push({ text: t, options: { breakLine: i < lines.length - 1 } });
    });
    T(s, runs, { x: 6.5, y: 2.02, w: 2.9, h: 3.05, fontSize: 9, paraSpaceAfter: 5, objectName: 'revealStory' });
  }

  {
    const s = content('Understanding Job History', 'Practice: read a job history', 'Open the practice ticket and answer these four questions.',
      'TRAINER NOTE: Open the practice job history during the slideshow using the button (https://www.homeaglow.com/nocm/homeaglow/jobhistory/?job=12685237). Give trainees time to answer, then go through the answers together, row by row from the bottom up.');
    const url = 'https://www.homeaglow.com/nocm/homeaglow/jobhistory/?job=12685237';
    // Button drawn as an image so the whole thing is one clickable link in Slides
    const btnSvg = `<svg xmlns="http://www.w3.org/2000/svg" width="930" height="180"><rect width="930" height="180" rx="30" fill="#${C.teal}"/><text x="465" y="114" font-family="Inter, Arial, sans-serif" font-size="64" font-weight="700" fill="#ffffff" text-anchor="middle">Open Practice Ticket  →</text></svg>`;
    const btn = 'image/png;base64,' + (await sharp(Buffer.from(btnSvg)).png().toBuffer()).toString('base64');
    s.addImage({ data: btn, x: 0.45, y: 1.45, w: 3.1, h: 0.6, hyperlink: { url, tooltip: 'Open the practice job history' } });
    T(s, [{ text: 'Job 12685237', options: { hyperlink: { url }, color: C.teal } }], { x: 3.75, y: 1.6, w: 3, h: 0.3, fontSize: 10, color: C.teal, valign: 'middle' });
    const qs = [
      'Who booked the appointment?',
      'When was the appointment initially scheduled for?',
      'The customer writes: "I initially booked for Jan 12 at 1 PM, but I received a notification that I am booked for 2 PM. Why is that?"',
      'What happened on Jan 8 at 5:58 PM?',
    ];
    const w = (9.1 - 0.15) / 2, h = 1.2;
    for (let i = 0; i < 4; i++) {
      const x = 0.45 + (i % 2) * (w + 0.15), y = 2.3 + Math.floor(i / 2) * (h + 0.15);
      box(s, x, y, w, h, C.white, C.border);
      badge(s, x + 0.18, y + 0.2, i + 1);
      T(s, qs[i], { x: x + 0.65, y: y + 0.15, w: w - 0.85, h: h - 0.3, fontSize: 10.5, valign: 'middle' });
    }
  }

  // ================= CALLBACK HANDLING =================
  await topic('Callback Handling', 'When to call a customer back, how many times to try, and what to do next.',
    ['When to call', 'The call steps', 'Answered, missed or voicemail', 'Keeping callbacks from slipping'],
    'ill30.png',
    'We don\'t have inbound phone support, but outbound calls are a tool we use whenever a customer asks for one, or when something needs an immediate answer (like confirming a double-booking). Category: Platform & Operating Mechanics.');

  {
    const s = content('Callback Handling', 'Asked for a call? Call them.', 'We don\'t take inbound calls, but we call out when a customer asks, or when something needs an answer now.',
      'Agent Do: if a customer requests a callback within business hours (8 AM - 8 PM, customer local time), call them right away. No LTV risk or escalation is needed to justify it; the request is enough. Why: the customer\'s preferred communication method is respected regardless of how serious the ticket looks, and consistent steps (timed reminders, endorsement) keep a promised callback from falling through between shifts. Outbound calls are also used when something needs an immediate answer, like confirming a double-booking.');
    box(s, 0.45, 1.5, 9.1, 1.1, C.teal);
    eyebrow(s, 'Agent do', 0.75, 1.63, 5, C.white);
    T(s, 'Callback request between 8 AM and 8 PM (customer\'s local time)? Call right away. The request is enough.', { x: 0.75, y: 1.88, w: 8.6, h: 0.6, fontFace: HEAD, bold: true, fontSize: 14, color: C.white });
    const w = (9.1 - 0.15) / 2;
    await card(s, 0.45, 2.8, w, 1.65, { ico: 'FiPhone', title: 'Respect their choice', body: 'They asked for a call, so they get one, however minor the ticket looks.' });
    await card(s, 0.45 + w + 0.15, 2.8, w, 1.65, { ico: 'FiRepeat', title: 'Never lose a callback', body: 'Reminders and the tracker keep promises from slipping between shifts.' });
  }

  {
    const s = content('Callback Handling', 'The call steps', null,
      'Step 1: did the customer request a callback? If not, this article doesn\'t apply; handle the ticket through normal channels. Step 2: within business hours (8 AM-8 PM customer local time)? Yes: attempt to call. No: address the concern via email and ask for their availability for a callback. Step 3: attempt to call up to 3 times, in this order: Primary CRM, then Google Voice, then Secondary CRM.');
    const steps = [
      ['FiHelpCircle', 'Callback requested?', 'No: handle the ticket as usual'],
      ['FiClock', '8 AM – 8 PM their time?', 'No: email, and ask for their availability'],
      ['FiPhoneCall', 'Call up to 3 times', 'Primary CRM, then Google Voice, then Secondary CRM'],
    ];
    const w = (9.1 - 2 * 0.5) / 3;
    for (let i = 0; i < 3; i++) {
      const x = 0.45 + i * (w + 0.5);
      box(s, x, 1.4, w, 2.4, C.white, C.border);
      await iconDot(s, x + (w - 0.6) / 2, 1.6, steps[i][0], 0.6);
      T(s, `STEP ${i + 1}`, { x, y: 2.32, w, h: 0.2, fontFace: HEAD, bold: true, fontSize: 8, color: C.teal, align: 'center', charSpacing: 0.5 });
      T(s, steps[i][1], { x: x + 0.15, y: 2.55, w: w - 0.3, h: 0.35, fontFace: HEAD, bold: true, fontSize: 12.5, align: 'center' });
      T(s, steps[i][2], { x: x + 0.15, y: 2.95, w: w - 0.3, h: 0.7, fontSize: 10, color: C.soft, align: 'center' });
      if (i < 2) s.addImage({ data: await icon('FiChevronRight', C.soft), x: x + w + 0.13, y: 2.45, w: 0.24, h: 0.24 });
    }
    await tip(s, 4.05, 'Call order:', 'Primary CRM, then Google Voice, then Secondary CRM.', 'FiPhone');
  }

  {
    const s = content('Callback Handling', 'Where to call from', 'On the customer\'s page in the New C CRM.',
      'Call options at the top right of Customer Information: 1 Callbox with the primary number; 2 GVoice (Google Voice); 3 Callbox with the secondary number (same Callbox button, choose the Secondary Phone in the From list). Call order: Primary CRM, then Google Voice, then Secondary CRM. TRAINER: click Show screens (or press the right arrow) to reveal each screen in order. Important: check the customer\'s Local Time before making a call (8 AM - 8 PM their time). Customer details are sample training data.');
    // Left: CRM (top half) + options + reminder
    await shot(s, 'callback_crm_half.png', 0.45, 1.35, 4.35, 1.25, [[1, 540, 5, 602, 30], [2, 476, 5, 539, 30], ['!', 319, 96, 436, 114]]);
    legend(s, 0.45, 2.75, 4.35, [['Callbox: primary number', 'Call 1st, from the primary number'], ['GVoice', 'Call 2nd, with Google Voice'], ['Callbox: secondary number', 'Call 3rd, same Callbox button, Secondary Phone']], 0.5);
    box(s, 0.45, 4.55, 4.35, 0.6, C.goldSoft);
    s.addImage({ data: await icon('FiClock', C.gold), x: 0.6, y: 4.74, w: 0.22, h: 0.22 });
    T(s, [{ text: 'Important:  ', options: { bold: true, color: C.gold } }, { text: 'check Local Time (!) before making a call. 8 AM – 8 PM their time.' }], { x: 0.92, y: 4.55, w: 3.8, h: 0.6, fontSize: 9.5, valign: 'middle' });
    // Right: Show screens button + panel; each click reveals the next screen
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 4.95, y: 1.35, w: 4.6, h: 0.45, fill: { color: C.teal }, line: { color: C.teal }, rectRadius: 0.1 });
    s.addImage({ data: await icon('FiEye', C.white), x: 5.12, y: 1.465, w: 0.22, h: 0.22 });
    T(s, 'Show screens', { x: 5.45, y: 1.35, w: 3.8, h: 0.45, fontFace: HEAD, bold: true, fontSize: 12.5, color: C.white, valign: 'middle' });
    box(s, 4.95, 1.9, 4.6, 3.25, C.tealSoft);
    const screens = [['call_primary_crop.png', 'Callbox: primary number', 'From list set to the primary number'], ['call_gvoice_crop.png', 'GVoice', 'Sign in with the shared Homeaglow account'], ['call_secondary_crop.png', 'Callbox: secondary number', 'From list set to the Secondary Phone']];
    for (let i = 0; i < 3; i++) {
      const y = 1.98 + i * 1.06;
      const nm = 'step' + (i + 1);
      s.addShape(pres.shapes.OVAL, { x: 5.07, y: y + 0.01, w: 0.2, h: 0.2, fill: { color: C.teal }, line: { color: C.teal }, objectName: nm + 'a' });
      T(s, [{ text: String(i + 1), options: { color: C.white, bold: true } }], { x: 5.07, y: y + 0.01, w: 0.2, h: 0.2, fontSize: 7.5, align: 'center', valign: 'middle', objectName: nm + 'b' });
      T(s, [{ text: screens[i][1] + '  ', options: { bold: true, color: C.teal } }, { text: screens[i][2], options: { color: C.soft } }], { x: 5.35, y, w: 4.1, h: 0.22, fontSize: 9, valign: 'middle', objectName: nm + 'c' });
      s.addShape(pres.shapes.RECTANGLE, { x: 5.07, y: y + 0.26, w: 4.36, h: 0.74, fill: { color: C.white }, line: { color: C.border, width: 0.5 }, objectName: nm + 'd' });
      const p = path.join(IMG, screens[i][0]);
      const { width, height } = await sharp(p).metadata();
      const r = Math.min(4.28 / width, 0.68 / height), iw = width * r, ih = height * r;
      s.addImage({ path: p, x: 5.07 + (4.36 - iw) / 2, y: y + 0.26 + (0.74 - ih) / 2, w: iw, h: ih, objectName: nm + 'e' });
    }
  }

  {
    const s = content('Callback Handling', 'What happens next', 'Every call ends one of three ways.',
      'Answered: attempt to resolve the issue on the call, send an email summary of what was discussed/resolved, and log an internal note about the conversation (calls are not recorded, so the next agents need to know what happened). Not answered: address the concern via email and ask for their availability for a callback. Voicemail: leave a voicemail if applicable, send email + SMS, then apply the 10-minute buffer rule: after sending the SMS, wait 10 minutes and check the ticket again. If the customer replies asking for an immediate callback, call right away, buffer or not. The buffer exists so we don\'t call back-to-back with no sign the customer has seen our message.');
    const cols = [
      ['FiPhoneIncoming', 'Answered', ['Try to resolve it on the call', 'Email a summary of what was discussed', 'Log an internal note. Calls aren\'t recorded']],
      ['FiPhoneMissed', 'Not answered', ['Address the concern by email', 'Ask for their availability for a callback']],
      ['FiVoicemail', 'Voicemail', ['Leave a voicemail, if applicable', 'Send an email and an SMS', 'Wait 10 minutes, then check the ticket again']],
    ];
    const gap = 0.15, w = (9.1 - 2 * gap) / 3;
    for (let i = 0; i < 3; i++) {
      const x = 0.45 + i * (w + gap);
      box(s, x, 1.45, w, 2.55, C.white, C.border);
      await iconDot(s, x + 0.18, 1.62, cols[i][0], 0.42);
      T(s, cols[i][1], { x: x + 0.72, y: 1.62, w: w - 0.85, h: 0.42, fontFace: HEAD, bold: true, fontSize: 13, valign: 'middle' });
      T(s, bullets(cols[i][2]), { x: x + 0.18, y: 2.2, w: w - 0.36, h: 1.7, fontSize: 10, paraSpaceAfter: 6 });
    }
    await tip(s, 4.2, '"Call me now"?', 'if they reply asking for an immediate call, call right away, even inside the 10 minutes.', 'FiZap');
  }

  {
    const s = content('Callback Handling', 'Once they share their availability', 'Create a timed reminder assigned to "Any CS", then pick your path.',
      'After the customer provides availability, create a timed reminder assigned to "Any CS". If you\'re on shift at the scheduled time: set a calendar reminder and include the ticket link in the invite so you can jump straight back in. If you\'re not on shift: log the ticket in the Callback Requests Tracker (Google Sheet). This two-path structure keeps a promised callback from getting lost across a shift change.');
    box(s, 0.45, 1.5, 9.1, 0.75, C.teal);
    s.addImage({ data: await icon('FiBell', C.white), x: 0.75, y: 1.71, w: 0.32, h: 0.32 });
    T(s, 'Create a timed reminder, assigned to "Any CS"', { x: 1.25, y: 1.5, w: 8, h: 0.75, fontFace: HEAD, bold: true, fontSize: 14, color: C.white, valign: 'middle' });
    await twoCol(s,
      { ico: 'FiCalendar', title: 'On shift at that time', items: ['Set a calendar reminder', 'Put the ticket link in the invite'] },
      { ico: 'FiFileText', title: 'Not on shift', items: ['Log the ticket in the Callback Requests Tracker'] }, 2.45, 1.75);
    await tip(s, 4.4, 'Why two paths:', 'a promised callback never gets lost across a shift change.', 'FiRepeat');
  }

  {
    const s = content('Callback Handling', 'What to send', 'Adapt these samples to the customer.',
      'Sample email if outside business hours: "Hi <<Name>>, I\'d be happy to hop on a call to discuss this with you. Please let me know what days and times generally work best for this week. In the meantime, I pulled up your account and <<address surface-level issues/resolutions>>. <<If applicable>> Let me know if that fully resolves your concerns! Otherwise, just reply with your availability and I\'ll get you scheduled." Sample text if a callback attempt was missed: "This is <<Name>> from Homeaglow. I tried calling but it went to voicemail — no worries, I\'ve already addressed your concern over email. If you\'d still like to hop on a call, just text me back with your availability and we\'ll set it up."');
    const w = (9.1 - 0.15) / 2;
    const msg = async (x, ico, title, text) => {
      box(s, x, 1.4, w, 3.3, C.white, C.border);
      await iconDot(s, x + 0.2, 1.58, ico, 0.42);
      T(s, title, { x: x + 0.75, y: 1.58, w: w - 0.9, h: 0.42, fontFace: HEAD, bold: true, fontSize: 12.5, valign: 'middle' });
      box(s, x + 0.2, 2.15, w - 0.4, 2.35, C.tealSoft);
      T(s, text, { x: x + 0.38, y: 2.27, w: w - 0.76, h: 2.15, fontSize: 9.5, italic: true });
    };
    await msg(0.45, 'FiMail', 'Email: outside business hours', '"Hi <<Name>>, I\'d be happy to hop on a call to discuss this with you. Please let me know what days and times generally work best for this week. In the meantime, I pulled up your account and <<address surface-level issues/resolutions>>. Let me know if that fully resolves your concerns! Otherwise, just reply with your availability and I\'ll get you scheduled."');
    await msg(0.45 + w + 0.15, 'FiMessageCircle', 'Text: missed callback', '"This is <<Name>> from Homeaglow. I tried calling but it went to voicemail — no worries, I\'ve already addressed your concern over email. If you\'d still like to hop on a call, just text me back with your availability and we\'ll set it up."');
  }

  // ================= C PRICE vs CP PAY =================
  await topic('C Price vs CP Pay', 'What a customer is charged, and what the cleaner is paid. Related, but calculated separately.',
    ['The three pricing cohorts', 'How the price is built', 'How the cleaner is paid', 'When a "platform fee" shows up'],
    'ill_pay.png',
    'The customer\'s price includes platform fees and taxes that don\'t go to the CP. The CP\'s pay depends on their own hourly rate and their operating model. Billing disputes are some of the most common tickets, and this is the formula reference behind Overcharged Hours, False Invoice and ETF calculations.');

  {
    const s = content('C Price vs CP Pay', 'Two separate calculations', 'You can\'t judge whether a charge is fair until you know how it\'s built.',
      'Customer price includes platform fees and taxes that don\'t go to the CP. CP pay depends on the CP\'s own hourly rate and operating model. Billing disputes are among the most common tickets; this is the formula reference that Overcharged Hours, False Invoice and ETF calculations draw from.');
    const w = (9.1 - 0.15) / 2;
    await card(s, 0.45, 1.5, w, 1.85, { ico: 'FiUser', title: 'Customer price', body: 'The cleaner\'s rate plus platform fees and taxes that don\'t go to the cleaner.' });
    await card(s, 0.45 + w + 0.15, 1.5, w, 1.85, { ico: 'FiDollarSign', title: 'Cleaner pay', body: 'Based on the cleaner\'s own hourly rate and their operating model.' });
    await tip(s, 3.6, 'Used everywhere:', 'Overcharged Hours, False Invoice and ETF calculations all start from these formulas.', 'FiInfo');
  }

  {
    const s = content('C Price vs CP Pay', 'Three pricing cohorts', 'What a customer pays on top of the cleaner\'s rate depends on their membership.',
      'Non-FC customer: $15.00/hr platform fee, 5% processing fee, sales tax varies by state. FC customer: no platform fee, 15% processing fee, sales tax varies. Deactivated FC customer: $30.00/hr platform fee, 5% processing fee, sales tax varies.');
    table(s, ['Cohort', 'Platform fee / hr', 'Processing fee', 'Sales tax'], [
      ['Non-FC customer', '$15.00', '5%', 'Varies by state'],
      ['FC customer', 'None', '15%', 'Varies by state'],
      ['Deactivated FC customer', '$30.00', '5%', 'Varies by state'],
    ], { y: 1.5, colW: [2.8, 2.1, 2.1, 2.1], fontSize: 11, rowH: 0.55 });
  }

  {
    const s = content('C Price vs CP Pay', 'How the cleaning price is built', null,
      'Cleaning Price = (CP Hourly Rate + Platform Hourly Rate) x Job Duration + Processing Fee + Sales Tax + Premium Fee. If the customer has an active voucher, voucher hours are deducted from the job duration before this calculation runs. If they have active credits, those are subtracted from the final total.');
    box(s, 0.45, 1.35, 9.1, 1.5, C.teal);
    eyebrow(s, 'Cleaning price', 0.75, 1.5, 5, C.white);
    T(s, '(CP hourly rate + Platform hourly rate) × Job duration\n+ Processing fee + Sales tax + Premium fee', { x: 0.75, y: 1.78, w: 8.6, h: 0.95, fontFace: HEAD, bold: true, fontSize: 16, color: C.white });
    const w = (9.1 - 0.15) / 2;
    await card(s, 0.45, 3.05, w, 1.5, { ico: 'FiTag', title: 'Active voucher', body: 'Voucher hours come off the job duration before the price is calculated.' });
    await card(s, 0.45 + w + 0.15, 3.05, w, 1.5, { ico: 'FiCreditCard', title: 'Active credits', body: 'Credits come off the final total.' });
  }

  {
    const s = content('C Price vs CP Pay', 'Let\'s practice!', 'Calculate the job cost for all 3 customer pricing cohorts using the details below.',
      'TRAINER: give trainees time to calculate, then click Show Answers (or press the right arrow). Each click reveals one cohort. Working: billable hours = 4 - 1.5 voucher = 2.5. FC: $25 x 2.5 = $62.50; + 15% processing = $71.88; + 8% tax = $77.63; - $12 credits = $65.63. Non-FC: ($25 + $15) x 2.5 = $100.00; + 5% = $105.00; + 8% = $113.40; - $12 = $101.40. Deactivated (cancelled) FC: ($25 + $30) x 2.5 = $137.50; + 5% = $144.38; + 8% = $155.93; - $12 = $143.93.');
    // Scenario card
    box(s, 0.45, 1.4, 3.3, 3.7, C.white, C.border);
    await iconDot(s, 0.65, 1.58, 'FiFileText', 0.42);
    T(s, 'Scenario', { x: 1.2, y: 1.58, w: 2.4, h: 0.42, fontFace: HEAD, bold: true, fontSize: 14, valign: 'middle' });
    const rows = [['CP rate', '$25.00/hr'], ['Duration', '4 hours'], ['Voucher', '1.5 hours'], ['Sales tax', '8%'], ['Credits', '$12.00']];
    rows.forEach(([k, v], i) => {
      const y = 2.2 + i * 0.52;
      box(s, 0.65, y, 2.9, 0.42, i % 2 ? C.white : C.bg, C.border);
      T(s, k, { x: 0.8, y, w: 1.3, h: 0.42, fontSize: 10.5, color: C.soft, valign: 'middle' });
      T(s, v, { x: 2.0, y, w: 1.45, h: 0.42, fontFace: HEAD, bold: true, fontSize: 12, align: 'right', valign: 'middle' });
    });
    // Show Answers button
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 3.95, y: 1.4, w: 5.6, h: 0.5, fill: { color: C.teal }, line: { color: C.teal }, rectRadius: 0.1 });
    s.addImage({ data: await icon('FiEye', C.white), x: 4.15, y: 1.53, w: 0.24, h: 0.24 });
    T(s, 'Show Answers', { x: 4.5, y: 1.4, w: 4.5, h: 0.5, fontFace: HEAD, bold: true, fontSize: 13, color: C.white, valign: 'middle' });
    // Answers panel: empty until clicked; one cohort per click
    box(s, 3.95, 2.0, 5.6, 3.1, C.tealSoft);
    const ans = [
      ['FC customer:  $65.63', '$25/hr CP rate × 2.5 hrs + 15% processing fee + 8% tax − $12 credits'],
      ['Non-FC customer:  $101.40', '($25/hr CP rate + $15/hr platform fee) × 2.5 hrs + 5% processing fee + 8% tax − $12 credits'],
      ['Deactivated (cancelled) FC customer:  $143.93', '($25/hr CP rate + $30/hr platform fee) × 2.5 hrs + 5% processing fee + 8% tax − $12 credits'],
    ];
    const runs = [];
    ans.forEach(([head, how], i) => {
      runs.push({ text: head, options: { bold: true, color: C.teal, fontSize: 12.5, fontFace: HEAD } });
      runs.push({ text: how, options: { softBreakBefore: true, breakLine: i < ans.length - 1, color: C.ink, fontSize: 9.5 } });
    });
    T(s, runs, { x: 4.15, y: 2.15, w: 5.2, h: 2.85, paraSpaceAfter: 10, objectName: 'revealAnswers' });
    T(s, 'Hint: take the voucher hours off the duration first.', { x: 0.65, y: 4.8, w: 3.0, h: 0.22, fontSize: 8.5, italic: true, color: C.soft });
  }

  {
    const s = content('C Price vs CP Pay', 'How the cleaner is paid', null,
      'CA/CAWA: CP\'s hourly rate x duration + tips + surge bonus, for first-time and repeat clients. General: first-time client (CP\'s hourly rate - $5.00 match fee) x duration + tips + surge bonus; repeat client CP\'s hourly rate x duration + tips + surge bonus. Key point: Homeaglow pays the CP directly via direct deposit regardless of whether the customer\'s own charge succeeds; a failed card, voucher or credits never affects CP pay. Tips: 100% go to the CP. Surge bonus: an extra $1-$50 Homeaglow may pay on certain jobs upon invoicing, positioned to CPs as "Homeaglow is especially busy."');
    table(s, ['Operating model', 'First-time client', 'Repeat client'], [
      ['CA / CA-WA', 'Rate × duration + tips + surge bonus', 'Same'],
      ['General', '(Rate − $5 match fee) × duration + tips + surge bonus', 'Rate × duration + tips + surge bonus'],
    ], { y: 1.3, colW: [1.9, 4.0, 3.2], fontSize: 10, rowH: 0.5 });
    await cards(s, 2.95, 1.55, [
      { ico: 'FiShield', title: 'Always paid', body: 'Even if the customer\'s card, voucher or credits fail.' },
      { ico: 'FiHeart', title: 'Tips', body: '100% of every tip goes to the cleaner.' },
      { ico: 'FiZap', title: 'Surge bonus', body: '$1–$50 extra on some jobs, at invoicing.' },
    ]);
  }

  {
    const s = content('C Price vs CP Pay', 'Let\'s practice!', 'Determine the CP pay in this scenario.',
      'Scenario: C is a repeat client and books a cleaning for 5 hours. C has a 3-hour cleaning voucher that will be applied to the job. Calculate the total job pay of CP Sarah and CP Gab if the job also has a $12.50 surge bonus and a $10.00 tip. TRAINER: click Show Answers (or press the right arrow); each click reveals one answer, then the key takeaway. Working: the voucher only changes what the customer pays. The CP is paid for all 5 hours (Homeaglow pays the CP regardless of vouchers, credits or failed charges). Sarah (CA/CAWA, repeat rate $28): $28 x 5 = $140 + $12.50 surge + $10 tip = $162.50. Gab (General, repeat client so no $5 match fee): $24 x 5 = $120 + $12.50 + $10 = $142.50.');
    // Job details
    box(s, 0.45, 1.35, 4.4, 1.55, C.white, C.border);
    await iconDot(s, 0.62, 1.48, 'FiBriefcase', 0.36);
    T(s, 'The job', { x: 1.08, y: 1.48, w: 3, h: 0.36, fontFace: HEAD, bold: true, fontSize: 12.5, valign: 'middle' });
    const job = [['Client', 'Repeat'], ['Booked', '5 hours'], ['Voucher', '3 hours'], ['Surge bonus', '$12.50'], ['Tip', '$10.00']];
    job.forEach(([k, v], i) => {
      const x = 0.62 + (i % 3) * 1.38, y = 1.95 + Math.floor(i / 3) * 0.45;
      T(s, [{ text: k + '  ', options: { color: C.soft, fontSize: 9 } }, { text: v, options: { bold: true, fontFace: HEAD, fontSize: 11 } }], { x, y, w: 1.35, h: 0.4, valign: 'middle' });
    });
    // CP cards
    const cps = [['CP Sarah', 'CA/CA-WA Operating Model', ['$30.00/hr  first-time C/CP pair', '$28.00/hr  repeat C/CP pair']], ['CP Gab', 'General Operating Model', ['$24.00/hr  one rate']]];
    for (let i = 0; i < 2; i++) {
      const x = 0.45 + i * 2.25;
      box(s, x, 3.05, 2.15, 2.1, C.white, C.border);
      await iconDot(s, x + 0.15, 3.18, 'FiUser', 0.45);
      T(s, cps[i][0], { x: x + 0.7, y: 3.18, w: 1.4, h: 0.45, fontFace: HEAD, bold: true, fontSize: 12.5, valign: 'middle' });
      T(s, cps[i][1], { x: x + 0.15, y: 3.72, w: 1.9, h: 0.3, fontSize: 9, bold: true, color: C.teal });
      T(s, cps[i][2].map((t, j) => ({ text: t, options: { breakLine: j < cps[i][2].length - 1 } })), { x: x + 0.15, y: 4.05, w: 1.9, h: 1.0, fontSize: 9, paraSpaceAfter: 4 });
    }
    // Show Answers
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.0, y: 1.35, w: 4.55, h: 0.5, fill: { color: C.teal }, line: { color: C.teal }, rectRadius: 0.1 });
    s.addImage({ data: await icon('FiEye', C.white), x: 5.2, y: 1.48, w: 0.24, h: 0.24 });
    T(s, 'Show Answers', { x: 5.55, y: 1.35, w: 3.8, h: 0.5, fontFace: HEAD, bold: true, fontSize: 13, color: C.white, valign: 'middle' });
    box(s, 5.0, 1.95, 4.55, 3.2, C.tealSoft);
    const ans = [
      ['CP Sarah:  $162.50', '$28/hr repeat rate × 5 hrs + $12.50 surge bonus + $10.00 tip'],
      ['CP Gab:  $142.50', '$24/hr × 5 hrs + $12.50 surge bonus + $10.00 tip (repeat client, so no $5 match fee)'],
      ['Key point', 'The 3-hour voucher only lowers what the customer pays. The cleaner is still paid for all 5 hours.'],
    ];
    const runs = [];
    ans.forEach(([head, how], i) => {
      runs.push({ text: head, options: { bold: true, color: C.teal, fontSize: 12.5, fontFace: HEAD } });
      runs.push({ text: how, options: { softBreakBefore: true, breakLine: i < ans.length - 1, color: C.ink, fontSize: 9.5 } });
    });
    T(s, runs, { x: 5.2, y: 2.1, w: 4.15, h: 2.95, paraSpaceAfter: 12, objectName: 'revealCpPay' });
  }

  {
    const s = content('C Price vs CP Pay', 'When a "platform fee" shows up', 'FC customers have no platform fee. If one shows on an FC invoice, it\'s one of these cases. Not an error.',
      'FC, General, first C/CP pairing: customer pays the CP\'s rate; the platform fee line is the $5/hr marketing fee deducted from the cleaner\'s pay, not an extra cost to the customer. FC, General, repeat pairing: nothing. FC, CA/CAWA, CP sets different first and repeat rates: customer pays the higher of the two whichever pairing it is; the line is the difference; the CP is paid their applicable rate and the difference goes to Homeaglow. FC, CA/CAWA, one rate: nothing. Non-FC or deactivated FC: customer pays the max CP rate, not the CP\'s actual rate; the line is the cohort platform fee ($15/hr non-FC, $30/hr deactivated FC) plus the gap between max rate and the CP\'s actual rate. Why the models differ: no marketing fee in CA/CAWA (state regulation); in exchange CPs may set separate first-pairing and repeat rates and we charge the higher.');
    table(s, ['Situation', 'Customer pays', 'The "platform fee" line is'], [
      ['FC · General · first pairing', 'The cleaner\'s rate', 'The $5/hr marketing fee, taken from the cleaner\'s pay'],
      ['FC · General · repeat pairing', 'The cleaner\'s rate', 'Nothing'],
      ['FC · CA/CA-WA · two rates', 'The higher of the two rates', 'The difference between the rates'],
      ['FC · CA/CA-WA · one rate', 'The cleaner\'s rate', 'Nothing'],
      ['Non-FC or deactivated FC', 'The max CP rate', 'Cohort fee ($15 or $30/hr) + gap to the cleaner\'s rate'],
    ], { y: 1.45, colW: [2.7, 2.4, 4.0], fontSize: 9.5, rowH: 0.44 });
    await tip(s, 4.4, 'The principle:', 'if the customer pays more than the cleaner\'s rate, the cleaner still gets their rate. The rest goes to Homeaglow.', 'FiInfo');
  }

  {
    const s = content('C Price vs CP Pay', 'Worked example', 'A deactivated FC customer sees a $40/hr platform fee. Is it right?',
      'Deactivated-FC customer, max rate $25, claiming CP\'s rate $15, platform fee shows $40/hr. That is the $30/hr deactivated-FC platform fee plus the $10/hr gap. The charge is correct: do not reprice it and do not refund the difference. Max rate behaves differently by membership: FC customers have a max rate too, but are charged the CP\'s actual rate; non-FC and deactivated-FC customers are charged the max rate while the CP is paid their own rate.');
    const parts = [['$30/hr', 'Deactivated FC\nplatform fee'], ['+', ''], ['$10/hr', 'Max rate $25 −\ncleaner\'s rate $15'], ['=', ''], ['$40/hr', 'Platform fee\non the invoice']];
    const ws = [2.4, 0.5, 2.4, 0.5, 2.4];
    let x = 0.65;
    for (let i = 0; i < 5; i++) {
      const [big, small] = parts[i];
      if (small) {
        box(s, x, 1.5, ws[i], 1.75, i === 4 ? C.teal : C.white, i === 4 ? C.teal : C.border);
        T(s, big, { x, y: 1.65, w: ws[i], h: 0.7, fontFace: HEAD, bold: true, fontSize: 28, color: i === 4 ? C.white : C.teal, align: 'center', valign: 'middle' });
        T(s, small, { x, y: 2.4, w: ws[i], h: 0.7, fontSize: 10, color: i === 4 ? C.white : C.soft, align: 'center' });
      } else {
        T(s, big, { x, y: 1.5, w: ws[i], h: 1.75, fontFace: HEAD, bold: true, fontSize: 26, color: C.soft, align: 'center', valign: 'middle' });
      }
      x += ws[i] + 0.1;
    }
    await tip(s, 3.55, 'Correct charge:', 'don\'t reprice it and don\'t refund the difference. The cleaner is still paid $15/hr.', 'FiCheckCircle');
  }

  {
    const s = content('C Price vs CP Pay', 'Max requested rate', 'Non-FC and deactivated FC customers are priced at the highest rate among the cleaners they requested.',
      'For non-FC customers, pricing follows the highest hourly rate among all the CPs they requested, not whichever CP claims. Example: the customer requests Mona ($25/hr), Cassandra ($30/hr) and Yolanda ($23/hr); the max requested rate is $30/hr. If a requested CP claims (say Mona), the customer is charged Mona\'s actual rate ($25/hr). If a non-requested CP claims (say Jeff at $18/hr), the customer is still charged $30/hr; Jeff is paid $18/hr and the difference goes to Homeaglow. Why: requesting Cassandra signals willingness to pay her rate, so it becomes their ceiling.');
    eyebrow(s, 'Requested cleaners', 0.48, 1.45, 5, C.soft);
    const req = [['Mona', '$25/hr', false], ['Cassandra', '$30/hr', true], ['Yolanda', '$23/hr', false]];
    for (let i = 0; i < 3; i++) {
      const x = 0.45 + i * 1.5;
      box(s, x, 1.72, 1.38, 1.0, req[i][2] ? C.teal : C.white, req[i][2] ? C.teal : C.border);
      T(s, req[i][0], { x, y: 1.8, w: 1.38, h: 0.35, fontFace: HEAD, bold: true, fontSize: 12, color: req[i][2] ? C.white : C.ink, align: 'center' });
      T(s, req[i][1], { x, y: 2.15, w: 1.38, h: 0.4, fontFace: HEAD, bold: true, fontSize: 15, color: req[i][2] ? C.white : C.teal, align: 'center' });
    }
    T(s, 'Max requested rate: $30/hr', { x: 0.48, y: 2.85, w: 4.4, h: 0.3, fontFace: HEAD, bold: true, fontSize: 12, color: C.teal });
    const w = 4.45;
    await card(s, 5.1, 1.45, w, 1.45, { ico: 'FiUserCheck', title: 'A requested cleaner claims', body: 'Mona claims: the customer pays Mona\'s rate, $25/hr.' });
    await card(s, 5.1, 3.0, w, 1.45, { ico: 'FiUserPlus', title: 'Someone else claims', body: 'Jeff ($18/hr) claims: the customer pays $30/hr. Jeff gets $18/hr.' });
    T(s, 'Requesting Cassandra shows the customer is willing to pay her rate, so it becomes their ceiling.', { x: 0.48, y: 3.3, w: 4.4, h: 0.9, fontSize: 10.5, color: C.soft });
  }

  {
    const s = content('C Price vs CP Pay', 'Before you explain or refund a fee', 'Three answers tell you which case you\'re in.',
      'Before explaining or refunding any platform-fee line, establish the operating model, the membership status, and whether it is that CP\'s first job with this customer. Those three answers decide which row of the platform fee table you are in.');
    await cards(s, 1.5, 1.85, [
      { n: 1, title: 'Operating model', body: 'CA/CA-WA or General? Look for the banner.' },
      { n: 2, title: 'Membership status', body: 'FC, non-FC, or deactivated FC?' },
      { n: 3, title: 'First job together?', body: 'Is this the cleaner\'s first job with this customer?' },
    ]);
    await tip(s, 3.6, 'Then:', 'find the matching row in the platform fee table before you explain or refund anything.', 'FiArrowRight');
  }

  // ================= VOUCHERS IN CRM =================
  await topic('Vouchers in CRM', 'What a voucher is, how to read it in the CRM, and how each type behaves.',
    ['What a voucher is', 'How vouchers work', 'Reading a voucher in the CRM', 'Voucher types at a glance'],
    'ill24.png',
    'Vouchers show up constantly in refund, cancellation and retention tickets. Getting the type wrong leads to real mistakes, like invalidating a voucher that should have been refundable, or telling a customer their gift card triggers a membership when it does not.');

  {
    const s = content('Vouchers in CRM', 'What is a voucher?', 'A discount on a customer\'s cleaning, redeemed with a code and tied to their account.',
      'Depending on how a customer signed up, they may have paid for it directly (Legacy DHJ), received it automatically when signing up for a membership (FCF), got it through a referral or gift, or received one from Care as a courtesy. Vouchers are the entry point for Legacy DHJ memberships: the customer\'s first interaction is "I bought a voucher." For FCF it is the reverse: the customer signs up for a membership directly, and the system issues a free DHJ voucher behind the scenes. All new customers move to FCF from August 3, 2026, so most new sign-ups won\'t have "bought a voucher" from their own perspective, even though a voucher exists on their account.');
    await cards(s, 1.45, 1.6, [
      { ico: 'FiShoppingCart', title: 'Bought it', body: 'Legacy DHJ: the customer paid for the voucher.' },
      { ico: 'FiStar', title: 'With membership', body: 'FCF: issued free when they signed up.' },
      { ico: 'FiGift', title: 'Gift or referral', body: 'Sent by a friend, or bought as a gift card.' },
      { ico: 'FiHeart', title: 'From Care', body: 'An admin courtesy voucher issued by us.' },
    ]);
    await tip(s, 3.3, 'Watch your words:', 'an FCF customer who says "I never bought a voucher" is right. Don\'t contradict them.', 'FiMessageCircle');
  }

  {
    const s = content('Vouchers in CRM', 'How vouchers work', null,
      'Voucher hours are deducted from a cleaning\'s total hours regardless of the CP\'s hourly rate. If the customer books longer than the voucher, the extra hours are charged at Homeaglow\'s standard rate. Customers redeem a voucher at <<Brand Name>>.com/redeem and must enter Zip code, Email and Voucher Code. One voucher per household: we generally allow one paid voucher per household; the system detects a paid voucher on the registered address and blocks another purchase.');
    await cards(s, 1.4, 1.75, [
      { ico: 'FiClock', title: 'Hours, not dollars', body: 'Voucher hours come off the job, whatever the cleaner\'s rate.' },
      { ico: 'FiPlusCircle', title: 'Longer job?', body: 'Extra hours are charged at Homeaglow\'s standard rate.' },
      { ico: 'FiKey', title: 'Redeem online', body: 'brand.com/redeem with zip code, email and voucher code.' },
      { ico: 'FiHome', title: 'One per household', body: 'Only one paid voucher per address.' },
    ]);
  }

  {
    const s = content('Vouchers in CRM', 'Reading a voucher in the new CRM', 'Once redeemed, the voucher shows on the customer\'s account.',
      'Walk through each column. Code: the 16-letter redemption code (e.g. ABCD-EFGH-IJKL-MNOP). Merchant: the voucher type, e.g. dhj, dhj_gift, groupon. Price and Hours: what was paid and the voucher duration; free FCF DHJ vouchers have no price. Purchased and Expires. FC and Refund. Use "+ Issue voucher" to issue one, and "View all in legacy CRM" for the full voucher history. In the legacy CRM the voucher sits below the card logo in an orange box; an applied voucher shows a green box with a check mark, job ID and appointment date; canceled, invalidated or expired vouchers are grayed out with the reason next to the code.');
    await shot(s, 'vouchers_newcrm.png', 0.45, 1.45, 9.1, 2.15, [
      [1, 22, 76, 320, 172], [2, 320, 76, 464, 172], [3, 464, 76, 676, 172], [4, 676, 76, 980, 186], [5, 980, 76, 1190, 172], [6, 18, 182, 400, 220],
    ]);
    legend(s, 0.45, 3.75, 2.95, [['Code', '16-letter redemption code'], ['Merchant', 'The voucher type (dhj, groupon…)']], 0.5, 1);
    legend(s, 3.53, 3.75, 2.95, [['Price & hours', 'Paid amount and duration'], ['Purchased & expires', 'When it was bought, when it ends']], 0.5, 3);
    legend(s, 6.6, 3.75, 2.95, [['FC & refund', 'Tied to FC? Refunded?'], ['Actions', 'Issue a voucher, or open legacy CRM']], 0.5, 5);
  }

  {
    const s = content('Vouchers in CRM', 'Voucher status', 'In the legacy CRM, the color tells you where the voucher stands.',
      'Unused (issued or redeemed): assigned to the account and will apply to the next invoiced job, shown in an orange box with duration, type, price and the 16-letter code. Applied (used): a green box with a check mark, the job ID and the appointment date and time. Grayed out: canceled/invalidated (removed through a refund, dispute or admin invalidation), converted (turned into credits equal to the purchase price), or expired. The reason usually appears beside the code. Refunded: refunded to the customer\'s card on file. Examples on the slide: an unused 3hr DHJ ($19.00) sitting below the card on file; a 3hr DHJ applied to Job 12242911 on 10/31/25 (green tag); a 2hr DHJ ($9.00) grayed out with the reason admin_refund on 5/1/26.');
    const rows = [
      ['FiTag', 'Unused', 'Orange tag. Applies to the next invoiced job.', 'v_unused.png'],
      ['FiCheckCircle', 'Applied', 'Green tag with a check mark, job ID and date.', 'v_applied.png'],
      ['FiSlash', 'Grayed out', 'Canceled, invalidated, converted or expired. Reason next to the code.', 'v_grayed.png'],
    ];
    for (let i = 0; i < 3; i++) {
      const y = 1.45 + i * 1.02, [ico, title, body, file] = rows[i];
      box(s, 0.45, y, 2.85, 0.92, C.white, C.border);
      await iconDot(s, 0.6, y + 0.25, ico, 0.42);
      T(s, [{ text: title, options: { bold: true, fontFace: HEAD, fontSize: 12, breakLine: true } }, { text: body, options: { color: C.soft, fontSize: 9 } }], { x: 1.12, y: y + 0.06, w: 2.1, h: 0.8, valign: 'middle' });
      box(s, 3.42, y, 6.13, 0.92, C.white, C.border);
      const { width, height } = await sharp(path.join(IMG, file)).metadata();
      const k = Math.min(5.9 / width, 0.76 / height);
      s.addImage({ path: path.join(IMG, file), x: 3.54, y: y + (0.92 - height * k) / 2, w: width * k, h: height * k });
    }
    await tip(s, 4.6, 'Converted?', 'the voucher was turned into credits equal to what the customer paid for it.', 'FiRefreshCw');
  }

  {
    const s = content('Vouchers in CRM', 'Voucher types at a glance', null,
      'Use this table before you invalidate, refund or explain any voucher. DHJ refunds can be done from the C Dashboard if purchased one year ago or less, or from the CRM. A DHJ converts to credits if an unqualified customer tries to redeem a second DHJ. Third-party merchant vouchers convert to credits on expiry. Legacy DHJ is deprecated: all new customers are on the new membership model from August 3, 2026, except those who bought a DHJ before the flip and customers who came from an AD.');
    table(s, ['Type', 'Comes from', 'Invalidate?', 'Refund?', 'Expires', 'To credits?'], [
      ['DHJ', 'Customer purchase', 'Yes (CRM)', 'Yes (Dashboard ≤1 yr, CRM)', 'Never', 'Yes, 2nd DHJ'],
      ['DHJ Gift', 'Referral', 'Yes (CRM)', 'No', '90 days', 'No'],
      ['First-Time', 'System', 'Yes (CRM)', 'No', 'Never', 'N/A'],
      ['Third-party (Groupon)', 'Customer purchase', 'Yes', 'Yes', '90 days', 'Yes, on expiry'],
      ['CP Offer', 'Cleaner', 'Yes (CRM)', 'No', '30 days', 'No'],
      ['After Nth Job', 'System', 'No', 'No', '3 months', 'No'],
      ['HG 2nd', 'Customer purchase', 'No', 'Yes', '1 year', 'Yes'],
      ['Admin Courtesy', 'CS / System', 'Yes (CRM)', 'No', '60 days', 'No'],
      ['Homeaglow Gift Card', 'Customer purchase', 'Yes (CRM)', 'Yes', '2 years', 'No'],
    ], { y: 1.2, colW: [1.85, 1.55, 1.2, 2.1, 1.0, 1.4], fontSize: 8.5, rowH: 0.36 });
  }

  {
    const s = content('Vouchers in CRM', 'Types you\'ll see most', null,
      'Free DHJ (new membership model): instead of buying a voucher, the customer signs up for a membership, pays a $19 MF upfront, and the system issues a free 3-hour DHJ voucher with no price indicator in the CRM; the membership is charged when the DHJ is applied to the account. DHJ Gift: shown as dhj_gift on the guest\'s account; single use (only the first guest can redeem); triggers FC once used; expires in 90 days. Third-party merchant (Groupon, Living Social, Gilt City): expires 90 days after redemption; cannot be split across appointments, and no admin courtesy voucher is issued for unused hours; converts to credits on expiry. CP Offer: up to 30 minutes off from a CP the customer has worked with; claim within 2 days, use within 30 days. Admin Courtesy: issued by CS or the system for refund requests, unworked hours, reviews or retention. Homeaglow Gift Card (since December 2025): a one-time cleaning bought as a gift; does NOT trigger FC and is not tied to the buyer\'s account.');
    const items = [
      ['FiStar', 'Free DHJ', '$19 MF at sign-up, free 3-hr voucher, no price shown.'],
      ['FiGift', 'DHJ Gift', 'Single use. Triggers FC. Expires in 90 days.'],
      ['FiShoppingBag', 'Third-party (Groupon)', 'Can\'t be split. Converts to credits on expiry.'],
      ['FiUser', 'CP Offer', 'Up to 30 min off. Claim in 2 days, use in 30.'],
      ['FiHeart', 'Admin Courtesy', 'Issued by CS or the system. Lasts 60 days.'],
      ['FiCreditCard', 'Homeaglow Gift Card', 'Does not trigger FC. Valid for 2 years.'],
    ];
    const w = (9.1 - 0.3) / 3;
    for (let i = 0; i < 6; i++) {
      await card(s, 0.45 + (i % 3) * (w + 0.15), 1.25 + Math.floor(i / 3) * 1.95, w, 1.8, { ico: items[i][0], title: items[i][1], body: items[i][2] });
    }
  }

  {
    const s = content('Vouchers in CRM', 'When a customer wants a refund', 'Always find and fix the reason first.',
      'Groupon: determine the reason (e.g. no show, cleaning quality) and address it; use the "Refund voucher (groupon)" macro in the comms kit for self-refund instructions; invalidate the voucher, otherwise Groupon may be unable to refund. Second <<BRAND>> voucher: first response, address the reason, encourage them to keep it, and give the self-refund link and instructions; second response, refund via C CRM > Do > Voucher > Refund > 2nd Voucher and let the customer know. Homeaglow Gift Card: first response, address the reason, remind them the card is valid for 2 years and can be gifted later, point them to the purchase confirmation email for the code, and invite them to reach out again; second response, escalate to your department\'s CSQ sheet so the charge is refunded via Stripe.');
    table(s, ['Voucher', 'First response', 'If they ask again'], [
      ['Groupon', 'Fix the reason, send self-refund steps (comms kit macro)', 'Invalidate the voucher so Groupon can refund'],
      ['Second brand voucher', 'Fix the reason, encourage them to keep it, send the self-refund link', 'Refund it: C CRM › Do › Voucher › Refund › 2nd Voucher'],
      ['Homeaglow Gift Card', 'Fix the reason. Valid for 2 years, code is in the confirmation email', 'Escalate to your CSQ sheet for a Stripe refund'],
    ], { y: 1.5, colW: [2.1, 3.6, 3.4], fontSize: 9.5, rowH: 0.62 });
    await tip(s, 4.15, 'Groupon:', 'invalidate the voucher, or Groupon may not be able to refund it.', 'FiAlertCircle');
  }

  {
    const s = content('Vouchers in CRM', 'How the system handles combos', null,
      'Scenario 1: two Groupon vouchers, the second converts to credits. Scenario 2: Groupon or second brand voucher applied to a shorter job, no admin courtesy voucher for unused hours (unlike DHJ). Scenario 3: DHJ voucher or active FC, then a Groupon, the Groupon converts to credits. Scenario 4: past invoiced job, then a Groupon, converts to credits because third-party vouchers are for new customers only. Scenario 5: Groupon first then DHJ: allowed, so we can convert them to FC; the DHJ is applied first; both can be used on separate appointments. Scenario 6: active FC customers can buy a 2nd DHJ if only one is true: booking address previously serviced by a DHJ, a card tied to a previous DHJ, or a used DHJ on the account. Scenario 7: deactivated FC customers can buy a new DHJ via the /deals page if they have no unused vouchers, no active FC, and none of the flags disputed_stripe_charge_id, customer_disputed_voucher, known_customer_fraud.');
    const items = [
      ['FiCopy', 'Two Groupons', 'The second one converts to credits.'],
      ['FiMinusCircle', 'Short job on a Groupon', 'No courtesy voucher for the unused hours.'],
      ['FiLayers', 'DHJ or FC, then Groupon', 'The Groupon converts to credits.'],
      ['FiRotateCcw', 'Past job, then Groupon', 'Converts to credits. New customers only.'],
      ['FiArrowUpCircle', 'Groupon, then DHJ', 'Allowed. DHJ is applied first.'],
      ['FiUserCheck', '2nd DHJ', 'Possible for some active or deactivated FC customers.'],
    ];
    const w = (9.1 - 0.15) / 2;
    for (let i = 0; i < 6; i++) {
      const x = 0.45 + (i % 2) * (w + 0.15), y = 1.25 + Math.floor(i / 2) * 1.2;
      box(s, x, y, w, 1.05, C.white, C.border);
      await iconDot(s, x + 0.18, y + 0.3, items[i][0], 0.42);
      T(s, [{ text: items[i][1], options: { bold: true, fontFace: HEAD, fontSize: 12, breakLine: true } }, { text: items[i][2], options: { color: C.soft, fontSize: 10 } }], { x: x + 0.78, y: y + 0.1, w: w - 0.95, h: 0.85, valign: 'middle' });
    }
  }

  // ================= NON-LOGGED-IN VISITOR =================
  await topic('Non-Logged-In Visitor Tickets', 'Someone used an existing customer\'s email, but wasn\'t logged in.',
    ['What the ticket looks like', 'Why it\'s a security checkpoint', 'How to respond'],
    'ill14.png',
    'Someone submits a ticket, usually via a public contact form or the Help Center, using an email address that already belongs to a registered customer. They haven\'t logged in, and may not actually be that customer.');

  {
    const s = content('Non-Logged-In Visitor Tickets', 'What it looks like', 'The new CRM flags these tickets for you.',
      'Point out: 1) the Triggered ticket badge and the source (Help Center); 2) the system note: "A non-logged in visitor submitted the below ticket using this user\'s email address. To avoid leaking sensitive information, ask the user to log in and re-submit a ticket as appropriate."; 3) the visitor\'s actual message. A matching email does not prove the person is the account owner. Anyone can type an email address into a form.');
    await shot(s, 'ticket_nonlogged.png', 0.45, 1.45, 9.1, 1.65, [
      [1, 402, 8, 786, 60], [2, 22, 70, 1400, 128], [3, 22, 142, 650, 180],
    ]);
    legend(s, 0.45, 3.3, 2.95, [['Triggered ticket', 'Came in from the Help Center']], 0.62, 1);
    legend(s, 3.53, 3.3, 2.95, [['System note', 'Sender wasn\'t logged in']], 0.62, 2);
    legend(s, 6.6, 3.3, 2.95, [['Their message', 'What the visitor is asking']], 0.62, 3);
    await tip(s, 4.15, 'Remember:', 'a matching email isn\'t proof. Anyone can type an email into a form.', 'FiAlertCircle');
  }

  {
    const s = content('Non-Logged-In Visitor Tickets', 'Why it matters', 'This is a security and privacy checkpoint, not a routine ticket.',
      'Two different customers, or a customer and someone impersonating them, could share visibility into the same account if you respond to the wrong party as if they were the account owner. Treat verifying identity as the first step, not an afterthought.');
    await cards(s, 1.5, 1.85, [
      { ico: 'FiMail', title: 'Anyone can type an email', body: 'A match doesn\'t mean it\'s the account owner.' },
      { ico: 'FiEyeOff', title: 'Wrong person, real data', body: 'Replying to the wrong person exposes the account.' },
      { ico: 'FiShield', title: 'Verify first', body: 'Confirm who you\'re talking to before you share anything.' },
    ]);
  }

  {
    const s = content('Non-Logged-In Visitor Tickets', 'How to respond', 'Check the account first. What you find decides your reply.',
      'Step 1: don\'t assume the submitter is the account owner just because the email matches. Step 2: check the account for identifying details, such as recent job details, the address, or an issue description that matches what\'s on file. Confirmed (their description matches account history, they reference specifics only the owner would know): proceed normally, no extra hoops for a genuinely verified customer. Can\'t confirm: respond generically; don\'t confirm account details, upcoming appointments or charges; ask them to log in to their dashboard or give a couple of identifying details first. Clearly not the owner (mismatched details, or asking something the real owner wouldn\'t need to ask): treat it as an unverified inquiry and share nothing account-specific.');
    box(s, 0.45, 1.45, 9.1, 0.62, C.teal);
    s.addImage({ data: await icon('FiSearch', C.white), x: 0.68, y: 1.63, w: 0.26, h: 0.26 });
    T(s, [{ text: 'First, check the account:  ', options: { bold: true, fontFace: HEAD } }, { text: 'recent jobs, address, and whether their issue matches what\'s on file.' }], { x: 1.08, y: 1.45, w: 8.3, h: 0.62, fontSize: 11.5, color: C.white, valign: 'middle' });
    await cards(s, 2.25, 2.0, [
      { ico: 'FiCheckCircle', title: 'Details match', body: 'Proceed as normal. No extra hoops for a verified customer.' },
      { ico: 'FiHelpCircle', title: 'Can\'t tell', body: 'Reply generically. Ask them to log in, or share a couple of identifying details.' },
      { ico: 'FiXCircle', title: 'Clearly not them', body: 'Treat it as unverified. Share nothing about the account.' },
    ]);
    await tip(s, 4.45, 'Until verified:', 'never confirm account details, upcoming appointments or charges.', 'FiLock');
  }

  // ================= MAX CP RATE =================
  await topic('Max CP Rate', 'The C-facing rate cap: the highest hourly rate a cleaner can claim this customer\'s jobs at.',
    ['What the cap does', 'What to flag before setting it', 'How the cap behaves', 'What to tell the customer'],
    'ill29.png',
    'An agent can set a maximum hourly rate on a customer\'s behalf. It\'s a pricing lever with real trade-offs: a lower cap protects the customer\'s wallet, but shrinks the pool of cleaners who can claim their jobs, possibly including their current cleaner.');

  {
    const s = content('Max CP Rate', 'What the cap does', 'Set in the legacy CRM: C CRM › Do › Update Max Rate.',
      'An agent can set a maximum hourly rate on a customer\'s behalf via the Update Max Rate action in the legacy CRM (C CRM > Do). It caps what CPs can claim the customer\'s jobs at, going forward; CPs above that rate simply can\'t claim. Trade-off: a lower cap protects the customer\'s wallet, but shrinks the pool of CPs who can claim, possibly including their current CP if that CP\'s rate is above the new cap. Setting it without understanding how it interacts with booked jobs, rescheduling and admin-locked jobs is how a well-intentioned cap turns into an unclaimed-job ticket a week later.');
    await twoCol(s,
      { ico: 'FiDollarSign', title: 'Protects the customer\'s wallet', items: ['Cleaners above the cap can\'t claim', 'New jobs price at or below the cap'] },
      { ico: 'FiUsers', title: 'Shrinks the cleaner pool', items: ['Their current cleaner may be above it', 'Too low, and jobs can go unclaimed'] },
      1.5, 2.2);
    await tip(s, 3.9, 'Not a pure win:', 'set it carefully, or it becomes an unclaimed-job ticket a week later.', 'FiAlertTriangle');
  }

  {
    const s = content('Max CP Rate', 'Before you set a cap', 'Flag these to the customer first.',
      'CPs set their own rates, so a lower cap can mean losing access to some of the platform\'s best cleaners, including the customer\'s current CP if that CP\'s rate is above the new cap. Setting it too low can mean jobs go unclaimed, especially in areas with fewer CPs at that price point. Most customers ask for $19-23/hr. Before setting the cap, check the customer\'s dashboard to see what CPs in their area actually charge. Don\'t set a number blind.');
    await cards(s, 1.45, 1.85, [
      { ico: 'FiUserX', title: 'They may lose their cleaner', body: 'Cleaners set their own rates. Theirs may be above the cap.' },
      { ico: 'FiCalendar', title: 'Jobs may go unclaimed', body: 'Especially where few cleaners charge that little.' },
      { ico: 'FiMap', title: 'Check the area first', body: 'Look at what cleaners near them charge on their dashboard.' },
    ]);
    await tip(s, 3.5, 'Typical ask:', 'most customers ask for $19–23/hr. Don\'t set a number blind.', 'FiInfo');
  }

  {
    const s = content('Max CP Rate', 'How the cap behaves', 'Several of these aren\'t what you\'d expect.',
      'The two easiest mistakes: assuming raising a cap will correct an existing job\'s rate (it won\'t; a cap can never raise a rate), and forgetting that admin-locked jobs (rate set via the Django job page) bypass the cap and the nightly sweep entirely. So a customer saying "I set a cap and this job still charged more" isn\'t necessarily wrong. Check whether that job was set via Django first.');
    table(s, ['Scenario', 'What happens'], [
      ['Lowering the cap', 'Immediately lowers the rate on booked future jobs (submitted or claimed)'],
      ['Raising the cap', 'Does nothing to existing jobs. A cap can never raise a rate'],
      ['Customer reschedules', 'The job keeps its rate. The nightly recalculation can\'t push it above the cap'],
      ['New job booked after the cap', 'Prices at or below the cap automatically'],
      ['Cleaner tries to claim above it', 'Can\'t. The cap blocks the claim'],
      ['Customer taps Standard / Premium', 'The cap stays in effect either way'],
      ['Rate set on the Django job page', 'Admin-locked: skips the cap and the nightly sweep entirely'],
    ], { y: 1.35, colW: [3.0, 6.1], fontSize: 9, rowH: 0.36 });
    await tip(s, 4.6, 'Charged above the cap?', 'check whether that job\'s rate was set via Django first.', 'FiSearch');
  }

  {
    const s = content('Max CP Rate', 'What to tell the customer', 'After the cap is set.',
      'Jobs are matched to CPs at or below their selected rate whenever possible; avoid "only" or "guaranteed" language. An invoice estimate shows on their dashboard right away once a CP claims the job. They can request a rematch up to 6 hours before the appointment. If they are lowering their cap, tell them it may affect upcoming jobs they\'ve already booked, not just future ones; this is the part that most contradicts intuition. Don\'t guarantee: this is a newly shipped feature still being monitored for edge cases, and admin-locked jobs bypass the cap. Terminology: don\'t say "max CP rate" or "max rate" to the customer; use plain language, e.g. "the highest rate a cleaner can charge for your jobs".');
    await cards(s, 1.4, 1.75, [
      { ico: 'FiUsers', title: 'Best match', body: 'To cleaners at or below their rate. Never "only" or "guaranteed".' },
      { ico: 'FiFileText', title: 'Estimate on claim', body: 'The invoice estimate shows on their dashboard right away.' },
      { ico: 'FiRepeat', title: 'Rematch', body: 'They can request one up to 6 hours before the appointment.' },
      { ico: 'FiTrendingDown', title: 'Lowering it?', body: 'It can change jobs they\'ve already booked too.' },
    ]);
    await tip(s, 3.35, 'Say it plainly:', '"the highest rate a cleaner can charge for your jobs", not "max CP rate".', 'FiMessageCircle');
  }

  // ---------- Close ----------
  {
    const s = pres.addSlide({ masterName: 'CONTENT' });
    eyebrow(s, 'Wrap-up', 0.62, 1.9);
    T(s, 'Questions?', { x: 0.62, y: 2.15, w: 5, h: 0.75, fontFace: HEAD, bold: true, fontSize: 34 });
    T(s, 'When in doubt, check the Knowledge Library or ask your Trainer or Team Lead.', { x: 0.62, y: 2.95, w: 4.6, h: 0.5, fontSize: 12, color: C.soft });
    await panel(s, 'ill9.png');
    s.addNotes('Open the floor for questions.');
  }

  await pres.writeFile({ fileName: OUT });
  console.log('wrote', OUT);
})();
