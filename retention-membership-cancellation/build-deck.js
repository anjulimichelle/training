// Builds the "Customer Foundations & C CRM" training deck.
// Usage: node deck.js out.pptx [imageDir]
const path = require('path');
const pptxgen = require('pptxgenjs');
const React = require('react');
const RDS = require('react-dom/server');
const sharp = require('sharp');
const Fi = require('react-icons/fi');

const OUT = process.argv[2] || 'deck.pptx';
const IMG = process.argv[3] || path.join(__dirname, '../img7');

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
pres.title = 'Retention & Membership Cancellation';

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
// An item may be a string, or an array of runs (mixed formatting within one bullet).
const bullets = (items) => items.flatMap((t, j) => {
  const runs = Array.isArray(t) ? t : [{ text: t }], last = j === items.length - 1;
  return runs.map((r, k) => ({ text: r.text, options: Object.assign({}, r.options, k === 0 ? { bullet: { indent: 10 } } : {}, k === runs.length - 1 && !last ? { breakLine: true } : {}) }));
});

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
const TOPICS = 7;
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

// Flowchart pieces in the deck palette: action (teal), decision (gold), penalty (dark), end (outline).
async function flowBox(s, x, y, w, h, text, kind) {
  const st = { action: [C.teal, C.teal, C.white], decision: [C.goldSoft, C.gold, C.gold], penalty: [C.ink, C.ink, C.white], end: [C.white, C.border, C.ink] }[kind];
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: st[0] }, line: { color: st[1], width: 1 }, rectRadius: 0.08 });
  T(s, text, { x, y, w, h, fontFace: HEAD, bold: true, fontSize: 10, color: st[2], align: 'center', valign: 'middle' });
}
function flowLine(s, pts, arrow) {
  for (let i = 0; i < pts.length - 1; i++) {
    const [x1, y1] = pts[i], [x2, y2] = pts[i + 1], last = i === pts.length - 2;
    s.addShape(pres.shapes.LINE, { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1), line: { color: C.soft, width: 1, endArrowType: last && arrow ? 'triangle' : undefined } });
  }
}
// Decision split: down from (cx, y0) to a bar at yBar, then down to two children at lx and rx.
function flowSplit(s, cx, y0, yBar, lx, rx, yChild, lLabel, rLabel) {
  flowLine(s, [[cx, y0], [cx, yBar]]);
  flowLine(s, [[lx, yBar], [rx, yBar]]);
  flowLine(s, [[lx, yBar], [lx, yChild]], true);
  flowLine(s, [[rx, yBar], [rx, yChild]], true);
  T(s, lLabel, { x: lx + 0.08, y: yBar + 0.02, w: 0.6, h: 0.22, fontSize: 8.5, bold: true, color: C.soft });
  T(s, rLabel, { x: rx + 0.08, y: yBar + 0.02, w: 0.6, h: 0.22, fontSize: 8.5, bold: true, color: C.soft });
}

// Button that cues the click-to-reveal descriptions on a flowchart slide.
async function detailsButton(s, x, y) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 1.95, h: 0.4, fill: { color: C.teal }, line: { color: C.teal }, rectRadius: 0.08 });
  s.addImage({ data: await icon('FiPlayCircle', C.white), x: x + 0.15, y: y + 0.1, w: 0.2, h: 0.2 });
  T(s, 'Show details', { x: x + 0.42, y, w: 1.45, h: 0.4, fontFace: HEAD, bold: true, fontSize: 11, color: C.white, valign: 'middle' });
}

// Small gold flag naming the recommended macro.
async function macroFlag(s, x, y, label, macro, w = 3.2) {
  box(s, x, y, w, 0.32, C.goldSoft, C.gold);
  s.addImage({ data: await icon('FiMessageSquare', C.gold), x: x + 0.12, y: y + 0.08, w: 0.16, h: 0.16 });
  T(s, [{ text: label + '  ', options: { bold: true, color: C.gold } }, { text: macro, options: { bold: true, color: C.ink } }],
    { x: x + 0.35, y, w: w - 0.4, h: 0.32, fontFace: HEAD, fontSize: 9.5, valign: 'middle' });
}

// Knowledge check: question on the left, answer on the right revealed one per click (step1..step4).
async function knowledgeCheck(eb, qa, notes, title = 'Knowledge check', sub = 'Ask each question, then click to reveal the answer.') {
  const s = content(eb + '  ·  Knowledge check', title, sub, notes);
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

// Practice: a live ticket on the left (trainer provides it), three numbered prompts on the right, then the Questions strip.
async function practice(eb, prompts, notes) {
  const s = content(eb + '  ·  Practice', 'Let\'s investigate together!', 'Review the live ticket and investigate.', notes);
  const top = 1.42, bottom = 4.3, gap = 0.15, lw = 3.7;
  box(s, 0.45, top, lw, bottom - top, C.white, C.border);
  await iconDot(s, 0.65, top + 0.2, 'FiInbox', 0.42);
  T(s, 'Live ticket', { x: 0.65, y: top + 0.78, w: lw - 0.4, h: 0.32, fontFace: HEAD, bold: true, fontSize: 13 });
  T(s, 'Your trainer will provide a live ticket. Review it and investigate before you decide anything.', { x: 0.65, y: top + 1.15, w: lw - 0.4, h: 0.8, fontSize: 10.5, color: C.soft });
  s.addText('Trainer provides', { shape: pres.shapes.ROUNDED_RECTANGLE, x: 0.65, y: bottom - 0.58, w: 1.5, h: 0.36, rectRadius: 0.08,
    fill: { color: C.goldSoft }, line: { color: C.goldSoft }, fontFace: BODY, bold: true, fontSize: 9.5, color: C.gold, align: 'center', valign: 'middle', margin: 0 });
  const rx = 0.45 + lw + gap, rw = 9.55 - rx, rh = (bottom - top - 2 * gap) / 3;
  for (let n = 0; n < prompts.length; n++) {
    const y = top + n * (rh + gap);
    box(s, rx, y, rw, rh, C.white, C.border);
    badge(s, rx + 0.2, y + (rh - 0.3) / 2, n + 1);
    T(s, prompts[n][0], { x: rx + 0.7, y: y + 0.14, w: rw - 0.9, h: 0.32, fontFace: HEAD, bold: true, fontSize: 12.5, valign: 'middle' });
    T(s, prompts[n][1], { x: rx + 0.7, y: y + 0.48, w: rw - 0.9, h: 0.28, fontSize: 10, color: C.soft, valign: 'middle' });
  }
  box(s, 0.45, 4.45, 9.1, 0.55, C.teal);
  s.addImage({ data: await icon('FiMessageCircle', C.white), x: 0.7, y: 4.585, w: 0.28, h: 0.28 });
  T(s, [{ text: 'Questions?  ', options: { bold: true, fontFace: HEAD, fontSize: 14 } }, { text: 'Open the floor before we move to the next topic.', options: { fontSize: 11 } }],
    { x: 1.15, y: 4.45, w: 8.2, h: 0.55, color: C.white, valign: 'middle' });
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

// Stat tile: a big number with a caption.
function stat(s, x, y, w, h, big, caption, color = C.teal) {
  box(s, x, y, w, h, C.white, C.border);
  T(s, big, { x, y: y + 0.18, w, h: 0.7, fontFace: HEAD, bold: true, fontSize: 30, color, align: 'center', valign: 'middle' });
  T(s, caption, { x: x + 0.2, y: y + 0.9, w: w - 0.4, h: h - 1.0, fontSize: 10, color: C.soft, align: 'center', valign: 'top' });
}

// Question on the left, what to check on the right: one row per validation question.
async function checkRows(s, y, rows, h = 0.62, gap = 0.08) {
  for (let i = 0; i < rows.length; i++) {
    const yy = y + i * (h + gap);
    box(s, 0.45, yy, 3.3, h, C.tealSoft, C.tealSoft);
    badge(s, 0.58, yy + (h - 0.3) / 2, i + 1);
    T(s, rows[i][0], { x: 1.0, y: yy, w: 2.7, h, fontFace: HEAD, bold: true, fontSize: 10.5, valign: 'middle' });
    box(s, 3.85, yy, 5.7, h, C.white, C.border);
    T(s, rows[i][1], { x: 4.0, y: yy, w: 5.45, h, fontSize: 9, color: C.ink, valign: 'middle' });
  }
}

// ---------- Deck 7: Retention & Membership Cancellation ----------
const RED = 'B3261E', PINK = 'FCEBEA';
const KC_TITLE = 'Let’s See What You’ve Got!';

// A sample line the agent can say, in a soft teal box with a quote mark.
function quote(s, x, y, w, h, text, label) {
  box(s, x, y, w, h, C.tealSoft);
  s.addShape(pres.shapes.RECTANGLE, { x, y: y + 0.08, w: 0.05, h: h - 0.16, fill: { color: C.teal }, line: { color: C.teal } });
  const runs = [];
  if (label) runs.push({ text: label, options: { bold: true, fontFace: HEAD, fontSize: 8, color: C.teal, breakLine: true } });
  runs.push({ text: '“' + text + '”', options: { italic: true, color: C.ink } });
  T(s, runs, { x: x + 0.2, y, w: w - 0.35, h, fontSize: 10, valign: 'middle' });
}

// Red "never" banner.
async function warn(s, y, label, text, h = 0.42) {
  box(s, 0.45, y, 9.1, h, PINK, RED);
  s.addImage({ data: await icon('FiAlertTriangle', RED), x: 0.62, y: y + (h - 0.18) / 2, w: 0.18, h: 0.18 });
  T(s, [{ text: label + '  ', options: { bold: true, color: RED } }, { text }], { x: 0.9, y, w: 8.5, h, fontSize: 10, valign: 'middle' });
}

// Table at any x/width (table() is always full width).
function tableAt(s, x, w, head, rows, { y = 1.45, colW, fontSize = 9, rowH = 0.3 } = {}) {
  const H = head.map(t => ({ text: t, options: { bold: true, color: C.teal, fill: { color: C.tealSoft }, fontFace: HEAD } }));
  const R = rows.map(r => r.map((t, i) => ({ text: t, options: { color: i === 0 ? C.ink : C.teal, bold: i !== 0, fill: { color: C.white } } })));
  s.addTable([H, ...R], { x, y, w, colW, rowH, fontFace: BODY, fontSize, valign: 'middle', border: { type: 'solid', pt: 0.75, color: C.border }, margin: [0.03, 0.1, 0.03, 0.1] });
}

// Coloured pill chip.
function chip(s, x, y, w, text, fill = C.tealSoft, color = C.teal, h = 0.3) {
  s.addText(text, { shape: pres.shapes.ROUNDED_RECTANGLE, x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { color: fill },
    fontFace: BODY, bold: true, fontSize: 9, color, align: 'center', valign: 'middle', margin: 0 });
}

// Panel with a coloured header strip and bullets underneath.
function headPanel(s, x, y, w, h, title, items, col = C.teal, soft = C.tealSoft, fontSize = 9.5) {
  box(s, x, y, w, h, soft, col);
  s.addText(title, { shape: pres.shapes.ROUNDED_RECTANGLE, x, y, w, h: 0.36, rectRadius: 0.06, fill: { color: col }, line: { color: col, width: 1 },
    fontFace: HEAD, bold: true, fontSize: 10.5, color: C.white, align: 'center', valign: 'middle', margin: 0 });
  T(s, bullets(items), { x: x + 0.18, y: y + 0.48, w: w - 0.36, h: h - 0.56, fontSize, paraSpaceAfter: 4 });
}

(async () => {
  // ---------- Cover ----------
  {
    const s = pres.addSlide({ masterName: 'COVER' });
    eyebrow(s, 'Homeaglow  ·  New Hire Care Training', 0.62, 0.6);
    T(s, 'Retention &\nMembership\nCancellation', { x: 0.62, y: 0.95, w: 4.9, h: 1.75, fontFace: HEAD, bold: true, fontSize: 30 });
    T(s, 'How we decide when to retain, when to make an offer, and how to let a customer go fairly.', { x: 0.62, y: 2.7, w: 4.6, h: 0.7, fontSize: 12.5, color: C.soft });
    T(s, 'Internal training material', { x: 0.62, y: 5.0, w: 3, h: 0.2, fontSize: 7, color: C.soft });
    await panel(s, 'ill6.png', 0.4, 4.82);
    s.addNotes('Welcome. This module covers four Knowledge Library articles from the Retention and Membership Cancellation category: Washington State: Temporary FC Cancellation Handling; Understanding Offer Logic: When and How Should I Make an Offer?; Cancellation Intent Surfaced in C/CP Comms or Reported by CP; the General Retention Playbook; Unused DHJ Voucher; FCF Pre-Invoice Cancellation and Refund Requests; and Trial Cleaning / One-Time Cleaning Conversion. The thread through all four: understand why the customer wants to leave, match what you do to that reason, and know the cases where retention is off the table.');
  }

  // ---------- Agenda ----------
  const agendaLinks = [], topicSlides = [];
  let agendaSlide;
  const addTopic = topic;
  topic = async (...args) => { const t = await addTopic(...args); topicSlides.push(t._slideNum); return t; };
  {
    const s = agendaSlide = content('Overview', 'Agenda', 'Seven topics. Click a topic to jump to it.',
      'Topic 1 is a state-specific override: Washington customers cancel with no ETF, no MCT and no retention. Topic 2 explains what each offer is built to do, so you can pick the right one. Topic 3 covers opening the conversation when you spot cancellation intent in C/CP messages. Topic 4 is the foundation for every retention conversation: the five key factors, matching the offer to the root cause, and when retention is not the right call. Topic 5 covers customers with an unused DHJ voucher. Topic 6 covers FCF customers cancelling before their first completed job, and which refund applies. Topic 7 covers Trial Cleaning and One-Time Cleaning as alternatives to a full membership. In slideshow mode, click any topic to jump straight to it.');
    const rows = [
      ['Washington State: FC Cancellation', 'ETF $0, MCT 0. Cancel right away, no retention.'],
      ['Understanding Offer Logic', 'What each offer is built to do, and when to use it.'],
      ['Cancellation Intent in C/CP Comms', 'The customer told the CP, not us. Reach out first.'],
      ['General Retention Playbook', 'Fix the real problem. Weigh the five key factors.'],
      ['Unused DHJ Voucher', 'Bought, not used. Sort the request, then act.'],
      ['FCF Pre-Invoice Cancellation', 'No completed job yet. Which refund applies?'],
      ['Trial & One-Time Cleaning', 'Alternatives to a full membership.'],
    ];
    const cw = (9.1 - 0.15) / 2, h = 0.8, gy = 0.1;
    for (let i = 0; i < rows.length; i++) {
      const [title, line] = rows[i], x = 0.45 + (i < 4 ? 0 : cw + 0.15), y = 1.42 + (i % 4) * (h + gy);
      const link = { slide: 1, tooltip: 'Go to ' + title };
      agendaLinks.push(link);
      s.addText([
        { text: title, options: { color: C.ink, bold: true, fontFace: HEAD, fontSize: 13, breakLine: true, hyperlink: link } },
        { text: line, options: { color: C.soft, fontSize: 10, hyperlink: link } },
      ], { shape: pres.shapes.ROUNDED_RECTANGLE, x, y, w: cw, h, rectRadius: 0.06, fill: { color: C.white }, line: { color: C.border, width: 0.75 },
        fontFace: BODY, valign: 'middle', margin: [56, 30, 0, 0], hyperlink: link });
      const num = { slide: 1, tooltip: 'Go to ' + title };
      agendaLinks.push(num);
      s.addText(String(i + 1).padStart(2, '0'), { x: x + 0.1, y, w: 0.55, h, fontFace: HEAD, bold: true, fontSize: 16, color: C.teal, align: 'center', valign: 'middle', hyperlink: num });
      const arrow = { slide: 1, tooltip: 'Go to ' + title };
      agendaLinks.push(arrow);
      s.addText('›', { x: x + cw - 0.42, y, w: 0.35, h, fontFace: HEAD, bold: true, fontSize: 20, color: C.teal, align: 'center', valign: 'middle', hyperlink: arrow });
    }
  }

  // ================= 1. WASHINGTON STATE =================
  const WA = 'Washington State';
  await topic('Washington State: FC Cancellation', 'A temporary, state-specific override of standard ETF and MCT handling.',
    ['Who it applies to', 'Why there\'s no ETF talk and no retention', 'The 5-step process', 'Edge cases'],
    'ill14.png',
    'Washington State: Temporary FC Cancellation Handling. Category: Retention and Membership Cancellation. This article covers a temporary, state-specific override of standard ETF/MCT handling for customers based in Washington State. For these customers specifically, the usual retention, ETF and MCT playbooks don\'t apply the way they normally would.');

  {
    const s = content(WA, 'What is this?', 'For WA customers, the system already set the fee and the term to zero.',
      'Effective May 11, 2026, Engineering set ETF to $0 and MCT to 0 months for all Washington customers: existing and new, any membership type. IMPORTANT: confirm the customer actually resides in Washington before applying any of this. This guidance does not apply just because a customer mentions the Homeaglow vs. Washington lawsuit; only their actual residence matters. A non-WA customer who references the lawsuit, with no material legal threat, is handled as BAU.');
    const w = (9.1 - 0.3) / 3;
    stat(s, 0.45, 1.45, w, 1.45, '$0', 'ETF for every WA customer');
    stat(s, 0.45 + w + 0.15, 1.45, w, 1.45, '0', 'months of MCT');
    stat(s, 0.45 + 2 * (w + 0.15), 1.45, w, 1.45, 'May 11', 'Effective 2026. Existing and new customers, any membership type', C.gold);
    await tip(s, 3.15, 'Confirm residency first:', 'mentioning the WA lawsuit isn\'t enough. Only where they live counts.', 'FiMapPin');
    await warn(s, 3.72, 'Zero Tolerance Policy:', 'mishandling a WA cancellation request is a ZTP violation.');
  }

  {
    const s = content(WA, 'Why it\'s handled differently', 'A legal and regulatory situation, not a more sympathetic customer.',
      'Washington customers are in this bucket because of a legal/regulatory situation (the Consent Decree and WA Attorney General involvement), not because their circumstances are more sympathetic than anyone else\'s. 1) The system already fixed the fee: ETF is $0 and MCT is 0 months at the account level. You\'re not deciding whether they deserve a waiver; it already happened automatically. Your job is to confirm it and act on it. 2) That\'s why you don\'t mention the ETF at all: there\'s nothing to explain or waive, so bringing it up only invites a conversation about something that no longer applies. Silence here isn\'t withholding information. 3) That\'s also why retention is off the table: the customer\'s path to cancel needs to be as friction-free as possible.');
    await cards(s, 1.45, 2.0, [
      { ico: 'FiCheckSquare', title: 'Already waived', body: 'ETF $0 and MCT 0 are set on the account. You confirm it and act on it.' },
      { ico: 'FiVolumeX', title: 'Never mention the ETF', body: 'Not to explain it, not to say it\'s waived. Nothing to discuss.' },
      { ico: 'FiSlash', title: 'No retention', body: 'The path to cancel must be as friction-free as possible.' },
    ]);
    eyebrow(s, 'ZTP violations', 0.45, 3.7, 4, RED);
    const v = ['Applying an ETF', 'Citing MCT to delay', 'Retention offers before cancelling', 'Disclosing ETF/MCT as if they apply'];
    const cw = (9.1 - 0.3) / 4;
    v.forEach((t, i) => chip(s, 0.45 + i * (cw + 0.1), 3.95, cw, '✕  ' + t, PINK, RED, 0.4));
  }

  {
    const s = content(WA, 'The process - Steps 1-2', 'Check the account, then cancel immediately.',
      'Step 1, confirm WA residency and account state. Agent Do: verify in the account that MCT = 0 months and ETF = $0. If both are set correctly, go to Step 2. If either isn\'t set to $0 / 0 months, don\'t try to fix it yourself: escalate to TL/Support first so they can raise it to get fixed. Step 2, cancellation request handling: 1) Do not attempt retention. 2) Acknowledge the request and confirm the customer\'s intent to cancel. 3) Once the account state is confirmed, cancel the FC membership immediately. 4) In your response: confirm the membership is cancelled; do not reference ETF in any way; include a soft re-engagement note in general terms only. Only get into specific alternative plans or pricing if the customer follows up later and shows interest; don\'t lead with options. If an ETF was accidentally charged anyway (it shouldn\'t be, given the $0 setting): refund it immediately, no need to wait for the customer to ask.');
    const w = 2.1, gap = 0.233;
    const flow = [['Confirm they live in WA', 'end'], ['Check: MCT 0, ETF $0', 'decision'], ['Confirm intent. No retention', 'end'], ['', 'action']];
    for (let i = 0; i < flow.length; i++) {
      const x = 0.45 + i * (w + gap);
      await flowBox(s, x, 1.45, w, 0.6, flow[i][0], flow[i][1]);
      if (i < flow.length - 1) flowLine(s, [[x + w, 1.75], [x + w + gap, 1.75]], true);
    }
    T(s, 'Cancel FC immediately', { x: 0.45 + 3 * (w + gap), y: 1.47, w, h: 0.33, fontFace: HEAD, bold: true, fontSize: 10, color: C.white, align: 'center', valign: 'bottom' });
    // Video link inside the last box (its own shape-level link so Google keeps the colour; keepU keeps the underline).
    s.addText('▷ Watch & Learn', { x: 0.45 + 3 * (w + gap) + (w - 1.2) / 2, y: 1.83, w: 1.2, h: 0.18, margin: 0, fontFace: BODY, bold: true, fontSize: 9, color: 'B6D7A8', underline: { style: 'sng' }, align: 'center', valign: 'middle',
      objectName: 'keepU_watch', hyperlink: { url: 'https://drive.google.com/file/d/1FNrb0bVkfw4z9XO_nxwU_teVxbLA2VtL/view?usp=sharing', tooltip: 'Play: cancel FC for a WA customer' } });
    const nx = 0.45 + w + gap + w / 2;
    flowLine(s, [[nx, 2.05], [nx, 2.3]], true);
    s.addText('Not set? Escalate to TL/Support. Don\'t fix it yourself.', { shape: pres.shapes.ROUNDED_RECTANGLE, x: nx - 1.5, y: 2.3, w: 3.0, h: 0.4, rectRadius: 0.08,
      fill: { color: PINK }, line: { color: RED, width: 0.75 }, fontFace: HEAD, bold: true, fontSize: 9, color: RED, align: 'center', valign: 'middle', margin: 0 });
    eyebrow(s, 'In your reply', 0.45, 2.92, 4, C.soft);
    for (const [i, [t, b]] of [['Confirm it\'s cancelled', 'Say the membership is cancelled.'], ['No ETF, at all', 'Not to explain it, not to say it\'s waived.']].entries()) {
      const y = 3.15 + i * 0.58;
      box(s, 0.45, y, 4.55, 0.5, C.white, C.border);
      badge(s, 0.58, y + 0.1, i + 1);
      T(s, [{ text: t + '  ', options: { bold: true, fontFace: HEAD, fontSize: 10.5 } }, { text: b, options: { fontSize: 9, color: C.soft } }], { x: 1.0, y, w: 3.9, h: 0.5, valign: 'middle' });
    }
    quote(s, 5.2, 3.15, 4.35, 1.08, 'If you change your mind, reach out and we can explore alternative pricing options that might fit better.', 'SOFT RE-ENGAGEMENT NOTE');
    await tip(s, 4.45, 'ETF charged anyway?', 'refund it right away. Don\'t wait for the customer to ask.', 'FiRotateCcw');
  }

  {
    const s = content(WA, 'Step 3: MF refund requests', 'The one place WA customers are treated more conservatively, not less.',
      'Step 3, Monthly Fee refund requests. This cuts against the instinct to be generous just because you\'re already waiving other things. Default: do not refund MF requests from WA customers. 1) Cancel FC proactively. 2) Tell the customer their membership is cancelled, but the MF isn\'t refundable, unless another policy would separately make them eligible (e.g. a standard refund window that would have applied regardless of WA status). Exception, WA AG notification refund requests: some WA customers cite a notice from the WA Attorney General and ask for MF refunds on that basis. Do not refund under the Consent Decree alone; handle as BAU, unless one of these is true: we failed to cancel them before their most recent MF charge (they asked to cancel and we were too slow), or we made cancellation harder than it should have been (pushed retention offers, slow responses). Suggested comms are on the slide.');
    headPanel(s, 0.45, 1.45, 4.45, 1.75, 'Default: no MF refund', ['Cancel FC proactively', 'Explain the MF isn\'t refundable', 'Unless another policy makes them eligible anyway'], C.ink, C.white);
    headPanel(s, 5.1, 1.45, 4.45, 1.75, 'Cites the WA AG notice?', ['Consent Decree alone: handle as BAU', [{ text: 'Refund only if ', options: {} }, { text: 'we were too slow', options: { bold: true } }, { text: ' to cancel before the latest MF' }], [{ text: 'Or ', options: {} }, { text: 'we made cancelling harder', options: { bold: true } }, { text: ' (retention push, slow replies)' }]], C.teal, C.tealSoft);
    quote(s, 0.45, 3.4, 9.1, 0.95, 'Thank you for reaching out. I\'ve cancelled your FC membership for you, so you\'re all set and won\'t be charged anymore. If anything about your experience wasn\'t quite right, I\'m here if you want to share. And if you ever decide to come back, we can definitely help you find something that fits better.', 'SUGGESTED COMMS');
  }

  {
    const s = content(WA, 'Steps 4–5: Vouchers & CP reports', 'No retention attempt in either case.',
      'Step 4, customers with unused DHJ vouchers (Legacy DHJ: the membership hasn\'t started yet) asking for the voucher to be refunded: proactively refund and invalidate the unused DHJ voucher and deactivate the FC table. No retention attempt needed. Step 5, the customer told the CP, not us: if a customer mentioned wanting to cancel to their CP but FC is still active in our system, reach out proactively. Let them know their CP flagged that they may want to cancel, and send the self-cancellation link, without mentioning ETF. Offer either option: they self-cancel, or you cancel it for them.');
    await cards(s, 1.45, 2.3, [
      { ico: 'FiGift', title: 'Step 4: Unused DHJ voucher', body: 'Legacy DHJ, membership not started, wants a refund: refund and invalidate the voucher, deactivate the FC table.' },
      { ico: 'FiMessageCircle', title: 'Step 5: They told the CP they want to cancel, not HG', body: 'FC still active? Reach out first. Send the self-cancel link, no ETF mention. They self-cancel, or you do it.' },
    ]);
    await tip(s, 4.0, 'Both cases:', 'no retention attempt needed.', 'FiSlash');
  }

  {
    const s = content(WA, 'Edge cases', 'When was the ETF charged, and does the customer actually live in WA?',
      'Common edge cases. Customer self-cancels and is charged an ETF: refund it; this shouldn\'t happen since ETF is set to $0 for self-cancels. Customer was charged ETF before the May 11, 2026 go-live: do not refund; it predates the change. Customer charged ETF after go-live but didn\'t explicitly ask for a refund: refund proactively anyway, don\'t wait to be asked. Non-WA customer references the WA lawsuit, with no material legal threat: handle as BAU; this guidance is residency-specific, not topic-specific. Related article: Zero Tolerance Policy > Mishandling Cancellation Requests for Washington State Customers.');
    table(s, ['Situation', 'What you do'], [
      ['Self-cancelled and was charged an ETF', 'Refund it'],
      ['ETF charged before May 11, 2026', 'Don\'t refund: it predates the change'],
      ['ETF charged after go-live, no refund asked', 'Refund it proactively'],
      ['Non-WA customer mentions the WA lawsuit', 'Handle as BAU (no material legal threat)'],
    ], { y: 1.45, colW: [4.6, 4.5], fontSize: 10.5, rowH: [0.4, 0.5, 0.5, 0.5, 0.5] });
  }

  await knowledgeCheck(WA, [
    ['A WA customer asks to cancel. Do you mention the ETF is waived?', 'No. Don\'t reference the ETF at all.'],
    ['The account shows ETF $140 for a WA customer. Fix it yourself?', 'No. Escalate to TL/Support, who raise it to get it fixed.'],
    ['A WA customer cites the AG notice and wants their MF back. We cancelled on time.', 'Handle as BAU. No refund on the Consent Decree alone.'],
    ['A Texas customer mentions the WA lawsuit. Apply this process?', 'No. It\'s residency-based. Handle as BAU.'],
  ], 'Ask each question and let trainees answer before revealing. Each click reveals the next answer.', KC_TITLE, '');

  await practice(WA, [
    ['Does the WA process apply? Why?', 'Confirm residency and the account state.'],
    ['What actions are you going to take?', 'Cancel, refund, escalate, or BAU?'],
    ['5-min writing challenge', 'Create your comms for the customer.'],
  ], 'TRAINER: share a live ticket from a Washington customer. Give trainees a few minutes to review it on their own, then work through it together. 1) Does the WA process apply? Confirm the customer resides in WA (not just that they mention the lawsuit) and check MCT = 0 and ETF = $0; if not set, escalate to TL/Support. 2) What actions are you going to take? Cancel FC immediately with no retention; never reference the ETF; refund any ETF charged after go-live; MF refunds: default no, unless we were too slow to cancel or made cancelling harder. 3) 5-minute writing challenge: each trainee writes the comms, including the soft re-engagement note and no ETF mention. Read a few aloud and compare. Then open the floor for questions.');

  // ================= 2. OFFER LOGIC =================
  const OL = 'Offer Logic';
  await topic('Understanding Offer Logic', 'When and how should I make an offer?',
    ['You don\'t need cancellation intent', 'Retention, Exit and Make-Right offers', 'What each offer is built to do', 'Caps, existing offers and guardrails'],
    'ill12.png',
    'Understanding Offer Logic: When and How Should I Make an Offer? Category: Retention and Membership Cancellation. One of the biggest misconceptions about offers is that you need cancellation intent before you can make one. You don\'t. This guide isn\'t about when you\'re allowed to offer; it\'s about what each offer is built to accomplish. Once you know that, deciding whether to use it takes care of itself.');

  {
    const s = content(OL, 'No cancellation intent needed', 'But offers aren\'t rewards, apologies, or something we issue just because we can.',
      'If the situation calls for it (a service issue, a cost concern, anything an offer is meant to fix), make it in your first response. Waiting for "I want to cancel" doesn\'t protect the company; it just delays solving a problem you could have fixed right away. The easiest customers to retain are usually the ones who haven\'t decided to leave yet. But don\'t swing too far the other way: "no cancellation intent needed" does not mean "offer something whenever a customer seems unhappy". Every offer exists to do a specific job. If an offer isn\'t doing one of these jobs, it probably doesn\'t need to happen. If you\'re ever unsure, don\'t ask "Am I allowed to use this?" Ask "Is this the problem this offer was designed to solve?"');
    box(s, 0.45, 1.45, 3.4, 2.55, C.teal);
    T(s, [{ text: 'Offer in your first response', options: { bold: true, fontFace: HEAD, fontSize: 15, breakLine: true } },
      { text: 'if the situation calls for it. The easiest customers to retain haven\'t decided to leave yet.', options: { fontSize: 10.5 } }],
      { x: 0.7, y: 1.45, w: 2.95, h: 2.55, color: C.white, valign: 'middle' });
    eyebrow(s, 'Every offer has a job', 4.05, 1.45, 5, C.soft);
    const jobs = [['FiHeart', 'Rebuild trust after a bad experience'], ['FiDollarSign', 'Remove a financial barrier'], ['FiRepeat', 'Make it easier to keep using Homeaglow'], ['FiLogOut', 'Close out a membership fairly']];
    for (const [i, [ico, t]] of jobs.entries()) {
      const y = 1.72 + i * 0.58;
      box(s, 4.05, y, 5.5, 0.5, C.white, C.border);
      await iconDot(s, 4.17, y + 0.07, ico, 0.36);
      T(s, t, { x: 4.65, y, w: 4.8, h: 0.5, fontFace: HEAD, bold: true, fontSize: 11, valign: 'middle' });
    }
    await tip(s, 4.2, 'Ask:', '"Is this the problem this offer was designed to solve?" Not "Am I allowed to use this?"', 'FiHelpCircle');
  }

  {
    const s = content(OL, 'Three kinds of offers', 'Every offer does one of three jobs.',
      'Retention Offers: used when the goal is to keep the customer as an active member; they remove whatever is in the way of staying. Exit Offers: used when the customer is leaving and isn\'t (or shouldn\'t be) retained; they make the exit fair and low-cost, not keep them enrolled. Make-Right Offers: used when Homeaglow or the system caused a charge or error that shouldn\'t have happened, whether the customer stays or leaves; they correct our own mistakes. Vouchers and credits stay Retention Offers even in a cancellation-intent conversation (the PCQ Retention Framework pairs a voucher with a customer who already said they want to cancel). The line isn\'t "did the customer ask to cancel", it\'s "have I assessed this customer as unretainable". Make-Right example: a customer cancels FC, but their RC stays active due to a system behaviour, a job still runs and charges a Lockout fee. Refunding that LO isn\'t retention or exit; it\'s making up for a mistake our system caused. Make-Right doesn\'t introduce new tools: see the individual charge/refund pages (LMC, Lockout, Priority Fee, etc.).');
    const cols = [
      ['Retention', C.teal, C.tealSoft, 'Keep the membership active', 'What\'s standing in the way of them staying?', 'Vouchers & credits, MF reduction, MCT reduction (Job 1), free months'],
      ['Exit', C.ink, C.white, 'Close the membership out fairly', 'What\'s the lowest-cost way for them to leave?', 'MCT reduction (Job 2), ETF reduction'],
      ['Make-Right', C.gold, C.goldSoft, 'Correct a Homeaglow or system error', 'What did we cause, and what makes it right?', 'The individual charge and refund pages (LMC, Lockout, Priority Fee…)'],
    ];
    const w = (9.1 - 0.3) / 3;
    cols.forEach(([t, col, soft, goal, q, tools], i) => {
      const x = 0.45 + i * (w + 0.15);
      box(s, x, 1.45, w, 3.0, soft, col);
      s.addText(t, { shape: pres.shapes.ROUNDED_RECTANGLE, x, y: 1.45, w, h: 0.42, rectRadius: 0.06, fill: { color: col }, line: { color: col },
        fontFace: HEAD, bold: true, fontSize: 13, color: C.white, align: 'center', valign: 'middle', margin: 0 });
      T(s, [
        { text: 'GOAL', options: { bold: true, fontFace: HEAD, fontSize: 7.5, color: C.soft, breakLine: true } },
        { text: goal, options: { bold: true, fontFace: HEAD, fontSize: 11, breakLine: true } },
        { text: ' ', options: { fontSize: 5, breakLine: true } },
        { text: 'ASK', options: { bold: true, fontFace: HEAD, fontSize: 7.5, color: C.soft, breakLine: true } },
        { text: q, options: { italic: true, fontSize: 10, breakLine: true } },
        { text: ' ', options: { fontSize: 5, breakLine: true } },
        { text: 'TOOLS', options: { bold: true, fontFace: HEAD, fontSize: 7.5, color: C.soft, breakLine: true } },
        { text: tools, options: { fontSize: 9.5 } },
      ], { x: x + 0.18, y: 2.0, w: w - 0.36, h: 2.35, valign: 'top' });
    });
    await tip(s, 4.6, 'The line:', 'not "did they ask to cancel?" but "have I assessed them as unretainable?"', 'FiGitBranch');
  }

  {
    const s = content(OL, 'Before you choose an offer', 'Run the quick test, then three questions.',
      'Quick test: 1) Is there a genuine account error or system-caused charge here? Then it\'s Make-Right, regardless of retention or exit status. 2) If not: have I already determined this customer isn\'t retainable? No: Retention. Yes: Exit. Then three questions. 1) What problem is the customer actually trying to solve? This is where most offer decisions go wrong: focus on why they\'re asking, not what they\'re asking for. A poor cleaning? The membership too expensive? Stuck in the commitment? About to be charged at a bad time? Simply done? 2) Does an offer solve that problem? Not every conversation needs one: sometimes the customer just wants information, an explanation, or the issue is already resolved. 3) Has this problem already been addressed? An existing offer doesn\'t automatically stop you issuing another. If it already addresses the current concern, use it; if not, upgrade it or pick a different offer. The number of offers isn\'t what matters.');
    // Quick test flow, left.
    eyebrow(s, 'Quick test', 0.45, 1.42, 3, C.soft);
    await flowBox(s, 0.45, 1.65, 3.6, 0.5, 'Account error or system-caused charge?', 'decision');
    flowLine(s, [[4.05, 1.9], [4.3, 1.9]], true);
    await flowBox(s, 4.3, 1.68, 1.05, 0.44, 'Make-Right', 'end');
    T(s, 'Yes', { x: 4.08, y: 1.7, w: 0.3, h: 0.18, fontSize: 7.5, bold: true, color: C.soft });
    flowLine(s, [[2.25, 2.15], [2.25, 2.4]], true);
    T(s, 'No', { x: 2.32, y: 2.17, w: 0.4, h: 0.2, fontSize: 7.5, bold: true, color: C.soft });
    await flowBox(s, 0.45, 2.4, 3.6, 0.5, 'Already assessed as unretainable?', 'decision');
    flowSplit(s, 2.25, 2.9, 3.05, 1.15, 3.35, 3.2, 'No', 'Yes');
    await flowBox(s, 0.45, 3.2, 1.4, 0.44, 'Retention', 'action');
    await flowBox(s, 2.65, 3.2, 1.4, 0.44, 'Exit', 'penalty');
    // Three questions, right.
    const qs = [['What problem are they really solving?', 'Why they ask, not what they ask for.'], ['Does an offer solve it?', 'Sometimes they just want an explanation.'], ['Already addressed?', 'Use the existing offer, or upgrade it.']];
    for (const [i, [t, b]] of qs.entries()) {
      const y = 1.45 + i * 0.75;
      box(s, 5.6, y, 3.95, 0.66, C.white, C.border);
      badge(s, 5.72, y + 0.18, i + 1);
      T(s, [{ text: t, options: { bold: true, fontFace: HEAD, fontSize: 10.5, breakLine: true } }, { text: b, options: { fontSize: 9, color: C.soft } }], { x: 6.12, y, w: 3.35, h: 0.66, valign: 'middle' });
    }
    await tip(s, 3.95, 'Not the count:', 'what matters is whether their current concern is already addressed.', 'FiLayers');
  }

  {
    const s = content(OL, 'Vouchers & credits', 'Retention offer. Two separate jobs.',
      'What they\'re for: getting the customer to book again, and/or making up for a service issue that\'s still sitting with them emotionally. When it doesn\'t make sense: if the customer is already willing to book again and seems genuinely fine with how the issue was handled, both jobs are done. But don\'t confuse "willing to book" with "the issue is resolved": a customer can be ready to book and still be annoyed about the last one; then still offer it, because their trust is what needs repair. If the customer\'s history suggests they may be abusing the compensation process, don\'t issue another voucher or credit right away; review the current complaint and confirm it qualifies first. Existing offer? Check whether they already have an unused one. Less than what\'s appropriate: top it up (a $5 credit when the situation qualifies for a 1-hour voucher: upgrade it, don\'t add another). Already meets or exceeds: use that instead. Already used: you may issue new compensation if the situation qualifies (unless there are abuse concerns).');
    await cards(s, 1.45, 1.4, [
      { ico: 'FiCalendar', title: 'Job 1: book again', body: 'Get the customer booking their next cleaning.' },
      { ico: 'FiHeart', title: 'Job 2: repair trust', body: 'Make up for an issue still sitting with them. Ready to book ≠ resolved.' },
    ]);
    eyebrow(s, 'Check for an unused one first', 0.45, 3.02, 6, C.soft);
    const ex = [['Less than appropriate', 'Top it up'], ['Meets or exceeds', 'Use that one'], ['Already used', 'Issue new, if it qualifies']];
    const w = (9.1 - 0.3) / 3;
    ex.forEach(([a, b], i) => {
      const x = 0.45 + i * (w + 0.15);
      box(s, x, 3.27, w, 0.62, C.tealSoft);
      T(s, [{ text: a, options: { fontSize: 9, color: C.soft, breakLine: true } }, { text: '→ ' + b, options: { bold: true, fontFace: HEAD, fontSize: 11, color: C.teal } }], { x: x + 0.15, y: 3.27, w: w - 0.3, h: 0.62, valign: 'middle' });
    });
    await warn(s, 4.08, 'Possible abuse?', 'don\'t issue another right away. Confirm the current complaint qualifies first.');
  }

  {
    const s = content(OL, 'MF reduction', 'Retention offer for when the membership fee itself feels too costly.',
      'What it\'s for: the membership fee feeling too costly for the customer\'s situation, not the cost of any individual cleaning. "The membership isn\'t worth it" / "I can\'t afford this right now": MF reduction is your default. "My cleaner\'s rate is too high": that\'s a per-cleaning cost problem; see Setting Max CP Rate instead. MF reduction only comes back as a small concession if the customer wants to keep their current, higher-rate CP. Why the gap matters: a $40/hr CP vs a wanted $25/hr is $45 on a 3-hour job; a $15 MF reduction doesn\'t come close. Don\'t shorten the cleaning duration to cut the cost: that trades their cost complaint for a quality complaint, unless the duration was genuinely excessive. Amounts: 1 to 5 months paid, $10–$15 is common; rare cases up to 50% (usually needs Support approval at 1–4 months). Existing offer? Always calculate from the original $59 fee, never the current discounted fee. If the new reduction is greater, replace the old one (on $49 with a $10 reduction, now qualifies for $15: $59 − $15 = $44, not $34). If the existing one is equal or greater, leave it.');
    twoColQuick(s);
    function twoColQuick(s) {
      headPanel(s, 0.45, 1.45, 4.45, 1.3, 'Membership cost', ['"It isn\'t worth it" / "I can\'t afford it"', [{ text: 'MF reduction is your default', options: { bold: true } }]]);
      headPanel(s, 5.1, 1.45, 4.45, 1.3, 'Cleaner\'s rate too high', ['A per-cleaning cost problem', [{ text: 'See Setting Max CP Rate', options: { bold: true } }]], C.ink, C.white);
    }
    stat(s, 0.45, 2.95, 2.9, 1.3, '$10–$15', 'common, 1 to 5 months paid');
    stat(s, 3.5, 2.95, 2.9, 1.3, 'Up to 50%', 'rare cases, usually with Support approval', C.gold);
    box(s, 6.55, 2.95, 3.0, 1.3, C.white, C.border);
    T(s, [{ text: 'Always from $59', options: { bold: true, fontFace: HEAD, fontSize: 12, color: C.teal, breakLine: true } },
      { text: 'On $49, now qualifies for $15:', options: { fontSize: 9, color: C.soft, breakLine: true } },
      { text: '$59 − $15 = $44', options: { bold: true, fontSize: 11, breakLine: true } },
      { text: 'not $34', options: { fontSize: 9, color: RED } }], { x: 6.7, y: 2.95, w: 2.75, h: 1.3, valign: 'middle' });
    await tip(s, 4.4, 'Don\'t shorten the cleaning:', 'it trades a cost complaint for a quality complaint.', 'FiScissors');
  }

  {
    const s = content(OL, 'MCT reduction', 'Two jobs. Cap: 2 months off the original 6-month MCT.',
      'Job 1, retention: use it when the concern is the length of the commitment itself ("six months is too long"), not the monthly fee. Size the offer to the objection, not the tone: general or vague concern, start with 1 month; a specific, concrete reason 1 month clearly won\'t resolve, go straight to 2 months. Decide upfront from what they said, not by offering 1 and upgrading if they push back. Job 2, exit: once the customer isn\'t retainable, 2 months is on the table if they\'re incredibly escalated or it\'s just the more sensible, lower-risk, amicable way to close; use judgment and flag to your TL if unsure. By months paid: 1–3 months, 1–2 months off; 4 months, 1 month off by default (ETF still applies), full 2 months only if incredibly escalated; 5 months, 1 month off, effectively waiving the ETF. Heads up: the reduction can\'t take anyone past 6 total months. 4 paid months + 1 off = 5, still 1 short; 5 paid + 1 off completes MCT and waives the ETF. Existing offer? Always calculate from the original 6-month MCT; if an existing reduction is equal or greater, leave it. Applies to both jobs.');
    headPanel(s, 0.45, 1.45, 4.45, 1.4, 'Job 1: Retention', ['"Six months is too long"', 'Vague concern: start with 1 month', 'Concrete reason: go straight to 2'], C.teal, C.tealSoft);
    headPanel(s, 5.1, 1.45, 4.45, 1.4, 'Job 2: Exit', ['Customer isn\'t retainable', '2 months if incredibly escalated', 'Or it\'s the most amicable close'], C.ink, C.white);
    table(s, ['Months paid', 'MCT reduction'], [
      ['1–3 months', '1–2 months off the original MCT'],
      ['4 months', '1 month off (ETF still applies). 2 only if incredibly escalated'],
      ['5 months', '1 month off: effectively waives the ETF'],
    ], { y: 3.0, colW: [2.2, 6.9], fontSize: 10, rowH: [0.32, 0.36, 0.36, 0.36] });
    await tip(s, 4.55, 'Decide upfront:', 'don\'t offer 1 month and upgrade when they push back.', 'FiTarget');
  }

  {
    const s = content(OL, 'Free months: two jobs', 'Retention offer. Give back paid time, or block a badly timed charge.',
      'Job 1, giving back time they already paid for: inactivity, forgetting about FC, a dormant membership, or a customer who moved and needs time before booking. Eligibility: any unused MF still within the 120-day disputable window qualifies; it doesn\'t need to be recent or in a row. Unused MFs older than 120 days: don\'t issue directly; send to Support for a second opinion and approval. Check by month, not by customer: Jan–Mar unused, 3 free months given, next MF pushed to Jul; if Jul–Sep then go unused, you can give up to 3 more for those. There\'s no lifetime cap unless you see a pattern that looks like abuse; flag it to Support. Job 2, protecting the customer from a charge about to hit at a bad time: Legacy DHJ, wants to cancel and 3 weeks or less into the current MF cycle, give a free month. FCF, wants to cancel because of a service issue AND the next MF is within 2 weeks: both must be true. Either model, wants to cancel and the next MF is ~5 days away: give a free month (or 1–3 weeks) right away. Either model, wants to cancel and the MF is overdue: bring it current. Doesn\'t want to cancel but has 2+ overdue MFs (or will within 7 days): bring the MF current. Give it now, don\'t offer and wait, and explain it clearly in writing.');
    headPanel(s, 0.45, 1.45, 4.0, 2.75, 'Job 1: Give back paid time', ['Unused MF within the 120-day window', 'Doesn\'t need to be recent or in a row', 'Check by month, not by customer', [{ text: 'Older than 120 days: ', options: { bold: true } }, { text: 'Support approves first' }]]);
    tableAt(s, 4.6, 4.95, ['Job 2: charge coming at a bad time', 'Give'], [
      ['Legacy DHJ: cancelling, ≤3 weeks into the MF cycle', 'Free month'],
      ['FCF: cancelling over a service issue AND MF due within 2 weeks', 'Free month'],
      ['Cancelling, next MF in ~5 days', 'Free time now'],
      ['Cancelling, MF already overdue', 'Bring MF current'],
      ['Not cancelling, 2+ overdue MFs (or within 7 days)', 'Bring MF current'],
    ], { y: 1.45, colW: [3.55, 1.4], fontSize: 8.5, rowH: [0.34, 0.48, 0.48, 0.42, 0.42, 0.48] });
    await tip(s, 4.35, 'Job 2:', 'give it now. Don\'t offer and wait, or they may be charged before they reply.', 'FiZap');
  }

  {
    const s = content(OL, 'Free month caps & repeat pauses', 'Free months only work while FC is active.',
      'Caps: standard (most cases) 3 months, the default ceiling. Pausing: 6 months; anything beyond 6 requires Support approval. Bulk charges (Job 2, overdue-MF scenario): as many as the bulk charges require. Free months and weeks only work while FC is active; once FC is cancelled there\'s nothing to add them to. Guardrail when a customer asks to pause again right after a previous pause: fewer than 2 MFs charged and paid since the last pause ended, Support approval needed no matter how long they ask for. You can give the pause normally, within the caps, only if all three are true: at least 2 MFs charged and paid since the last pause ended; the customer has actually used the service (booked or completed cleanings) since then; and they\'re not also asking for a refund right now. If any one isn\'t true, Support approval: it may be abuse. Why: pausing then immediately pausing again with no MF in between looks like avoiding the MF and the ETF. Payment alone doesn\'t prove it\'s genuine.');
    const w = (9.1 - 0.3) / 3;
    stat(s, 0.45, 1.45, w, 1.25, '3 months', 'Standard cap');
    stat(s, 0.45 + w + 0.15, 1.45, w, 1.25, '6 months', 'Pausing. More needs Support', C.gold);
    stat(s, 0.45 + 2 * (w + 0.15), 1.45, w, 1.25, 'As needed', 'Bulk overdue charges');
    eyebrow(s, 'Asks to pause again? All three must be true', 0.45, 2.88, 7, C.soft);
    const g = [['FiCreditCard', '2+ MFs charged and paid', 'since the last pause ended'], ['FiCheckCircle', 'Actually using the service', 'booked or completed cleanings'], ['FiSlash', 'Not asking for a refund', 'at the same time']];
    for (const [i, [ico, t, b]] of g.entries()) {
      const x = 0.45 + i * (w + 0.15);
      box(s, x, 3.12, w, 0.72, C.white, C.border);
      await iconDot(s, x + 0.14, 3.3, ico, 0.36);
      T(s, [{ text: t, options: { bold: true, fontFace: HEAD, fontSize: 10, breakLine: true } }, { text: b, options: { fontSize: 8.5, color: C.soft } }], { x: x + 0.6, y: 3.12, w: w - 0.7, h: 0.72, valign: 'middle' });
    }
    await warn(s, 4.0, 'Any one missing?', 'Support approval first. It could be abuse, not a real need for a pause.');
  }

  {
    const s = content(OL, 'They cancel anyway', 'Invalidate the retention offer, with two exemptions.',
      'A retention offer exists to keep the membership. If we offered one and the customer cancelled anyway, invalidate it as part of processing the cancellation. Two exemptions, where the customer is entitled to what they hold: 1) The admin courtesy voucher for unused hours where the Legacy DHJ voucher was only partly used (a 4-hour voucher against a 2-hour cleaning): the system issues an admin courtesy voucher for the balance. Leave it; those hours were bought and paid for. 2) Credits accepted within MCT where the customer then completed the MCT: they tried to self-cancel inside MCT, the system offered credits, they accepted, stayed, and fulfilled the full MCT. The system strips those credits on cancellation; that\'s system behaviour, not our policy. If they already booked and were charged in full with no credit applied, issue an admin refund equal to the reverted credits. If they haven\'t booked yet and want to use them, issue the credits. The principle: did the customer pay for it, or earn it by staying? Either way it survives the cancellation.');
    box(s, 0.45, 1.45, 9.1, 0.6, C.ink);
    T(s, 'Default: invalidate the retention offer when you process the cancellation.', { x: 0.7, y: 1.45, w: 8.7, h: 0.6, fontFace: HEAD, bold: true, fontSize: 12, color: C.white, valign: 'middle' });
    await cards(s, 2.2, 1.6, [
      { ico: 'FiShoppingBag', title: 'Paid for: keep it', body: 'Admin courtesy voucher for the unused hours of a partly used Legacy DHJ voucher.' },
      { ico: 'FiAward', title: 'Earned: keep it', body: 'Credits accepted inside MCT, then MCT completed. Refund or reissue what the system stripped.' },
    ]);
    await tip(s, 4.0, 'Ask:', 'did the customer pay for it, or earn it by staying? Then it survives the cancellation.', 'FiHelpCircle');
  }

  {
    const s = content(OL, 'Exit offers: MCT or ETF reduction?', 'Pick whichever gets the customer out for less.',
      'MCT reduction (Job 2): once a customer isn\'t retainable, doesn\'t qualify for an ETF waiver, and reducing the MCT would cost them less than reducing the ETF, use the MCT reduction. You\'re using it because it\'s the most cost-effective way to close the membership out. Same calculation base and existing-offer check as the retention MCT reduction. ETF reduction: first check the first job\'s hours. The ETF is $35/hr × first-cleaning hours, so an overcharge inflates the ETF. Resolve the OCH first, then present the corrected ETF. Example: first job invoiced 4 hours, the CP\'s own messages show 2.5: refund 1.5 hours and the ETF drops from $140 to $87.50. That\'s the ETF you present. When to offer it: if unawareness (or "I only wanted a one-time cleaning") is the only reason, don\'t offer an ETF reduction proactively; wait for a hint they\'d pay a reduced ETF. If any other reason is mixed in, ETF reduction is part of the standard exit strategy once retention is exhausted. Exception, extreme financial hardship: go straight to exit, offer 50% off the ETF; if they decline, waive it entirely.');
    box(s, 0.45, 1.45, 4.45, 1.5, C.white, C.border);
    T(s, [{ text: 'Fix the first job first', options: { bold: true, fontFace: HEAD, fontSize: 12, color: C.teal, breakLine: true } },
      { text: 'ETF = $35/hr × first-cleaning hours', options: { bold: true, fontSize: 10.5, breakLine: true } },
      { text: 'Invoiced 4 hrs, CP messages show 2.5 → refund 1.5 hrs', options: { fontSize: 9.5, color: C.soft, breakLine: true } },
      { text: 'ETF $140 → $87.50', options: { bold: true, fontSize: 12, color: C.teal } }], { x: 0.65, y: 1.45, w: 4.1, h: 1.5, valign: 'middle' });
    headPanel(s, 5.1, 1.45, 4.45, 1.5, 'Which tool?', ['Not retainable, no ETF waiver', 'Compare MCT vs ETF reduction', [{ text: 'Use whichever costs the customer less', options: { bold: true } }]], C.ink, C.white);
    table(s, ['Why they\'re cancelling', 'ETF reduction'], [
      ['Unawareness only (or "one-time cleaning")', 'Don\'t offer proactively. Wait for a hint they\'d pay'],
      ['Any other reason mixed in', 'Standard exit tool once retention is exhausted'],
      ['Extreme financial hardship', 'Straight to 50% off. Declines? Waive the ETF'],
    ], { y: 3.1, colW: [4.0, 5.1], fontSize: 9.5, rowH: [0.32, 0.38, 0.38, 0.38] });
  }

  {
    const s = content(OL, 'ETF reduction: the 50% cap', 'Deductions change how you get there, not the ceiling.',
      'What can be deducted: the MF and/or voucher price is the standard deduction. Don\'t stack them automatically; only when justified (the customer contests paying the ETF on top of what they already paid, or explicitly asks for both). Additional-hours charge: if the job that triggered FC ran longer than the voucher\'s covered hours and the customer was charged for that overage on their card, it can also be deducted. Check for it on jobs where they booked more hours than the voucher covered. The cap is always 50% off the original ETF, whichever deductions apply. Note the customer\'s Lifetime Net Revenue (LTNR) in your internal note when reducing the ETF, like LMC/Lockout. It\'s documentation, not a limiter: it doesn\'t change what you can offer, and the 50% reduction must still be maxed out before Graceful Closure. Existing offer? Always calculate from the original ETF, never an already-discounted amount, and check for an existing reduction first.');
    stat(s, 0.45, 1.45, 2.6, 2.3, '50%', 'max off the original ETF, whatever is deducted');
    eyebrow(s, 'What can be deducted', 3.25, 1.45, 5, C.soft);
    const d = [['FiDollarSign', 'MF and/or voucher price', 'Standard. Stack only when justified.'], ['FiClock', 'Additional-hours charge', 'If the FC-triggering job ran over the voucher\'s hours.']];
    for (const [i, [ico, t, b]] of d.entries()) {
      const y = 1.72 + i * 0.75;
      box(s, 3.25, y, 6.3, 0.66, C.white, C.border);
      await iconDot(s, 3.38, y + 0.15, ico, 0.36);
      T(s, [{ text: t, options: { bold: true, fontFace: HEAD, fontSize: 10.5, breakLine: true } }, { text: b, options: { fontSize: 9, color: C.soft } }], { x: 3.85, y, w: 5.6, h: 0.66, valign: 'middle' });
    }
    await tip(s, 3.95, 'Internal note:', 'record the customer\'s LTNR. Context only, not a limit.', 'FiFileText');
    await tip(s, 4.47, 'Always from the original ETF:', 'never an already-discounted amount. Check for an existing reduction.', 'FiRefreshCw');
  }

  await knowledgeCheck(OL, [
    ['A customer complains about a poor cleaning but hasn\'t mentioned cancelling. Can you offer a voucher?', 'Yes. You don\'t need cancellation intent if an offer fixes the problem.'],
    ['Customer is on $49 (a $10 reduction) and now qualifies for $15. New fee?', '$44: always from $59. $59 − $15.'],
    ['FC was cancelled, but RC ran and charged a Lockout fee. What kind of offer is the refund?', 'Make-Right: we caused it, whatever their status.'],
    ['Cancelling only because they were unaware of FC. Offer an ETF reduction?', 'Not proactively. Wait for a hint they\'d pay a reduced ETF.'],
  ], 'Ask each question and let trainees answer before revealing. Each click reveals the next answer.', KC_TITLE, '');

  await practice(OL, [
    ['What problem is the customer solving?', 'And which kind of offer fits: Retention, Exit or Make-Right?'],
    ['What actions are you going to take?', 'Which offer, how much, and any existing offer?'],
    ['5-min writing challenge', 'Create your comms for the customer.'],
  ], 'TRAINER: share a live ticket where an offer is (or isn\'t) appropriate. Give trainees a few minutes to review it on their own, then work through it together. 1) What problem is the customer actually solving, and which kind of offer fits? Run the quick test: account or system error, Make-Right; not retainable, Exit; otherwise Retention. 2) What actions are you going to take? Pick the tool built for that problem (voucher/credit, MF reduction from $59, MCT reduction sized to the objection, free months by job, ETF reduction capped at 50%). Check for an existing offer first. 3) 5-minute writing challenge: each trainee writes the comms. Read a few aloud and compare. Then open the floor for questions.');

  // ================= 3. CANCELLATION INTENT IN C/CP COMMS =================
  const CI = 'Cancellation Intent';
  await topic('Cancellation Intent in C/CP Comms', 'The customer told their CP they want to cancel. They haven\'t told us yet.',
    ['Why we reach out first', 'Opening the conversation the right way', 'The 5-step process', 'Routing to the right playbook'],
    'ill30.png',
    'Cancellation Intent Surfaced in C/CP Comms or Reported by CP. Category: Retention and Membership Cancellation. This covers a cancellation conversation that starts not because the customer told us, but because we saw a cancellation intent the customer sent to the CP while reading C/CP messages for a different reason: a service issue, a scheduling question, anything. If the customer reaches out to us about cancelling, go straight to the relevant Retention Playbook instead.');

  {
    const s = content(CI, 'What is this?', 'We spotted it in C/CP messages while investigating something else.',
      'Customers don\'t always realize their cleaner can\'t cancel a membership on their behalf: a CP can cancel a job, but not the FC relationship itself. If we sit on that information and do nothing, the most likely outcome is the customer gets frustrated days or weeks later when nothing has changed and the membership fee hits again. Reaching out first heads that off, and it reads as attentive rather than intrusive, but only if you open the conversation the right way. This page is about the moment before a normal cancellation conversation: how to open it when the customer doesn\'t know we\'ve seen it yet.');
    const w = 2.75, gap = 0.425;
    const flow = [['C tells the CP: "I want to cancel"', 'end'], ['CP can cancel a job, not FC', 'decision'], ['The MF hits again. Frustration', 'penalty']];
    for (let i = 0; i < flow.length; i++) {
      const x = 0.45 + i * (w + gap);
      await flowBox(s, x, 1.65, w, 0.7, flow[i][0], flow[i][1]);
      if (i < flow.length - 1) flowLine(s, [[x + w, 2.0], [x + w + gap, 2.0]], true);
    }
    eyebrow(s, 'If we do nothing', 0.45, 1.42, 3, C.soft);
    box(s, 0.45, 2.6, 9.1, 0.95, C.teal);
    s.addImage({ data: await icon('FiSend', C.white), x: 0.75, y: 2.88, w: 0.38, h: 0.38 });
    T(s, [{ text: 'So we reach out first', options: { bold: true, fontFace: HEAD, fontSize: 15, breakLine: true } }, { text: 'It reads as attentive, not intrusive, if we open it the right way.', options: { fontSize: 10.5 } }],
      { x: 1.35, y: 2.6, w: 8.0, h: 0.95, color: C.white, valign: 'middle' });
    await tip(s, 3.75, 'Customer contacted us about cancelling?', 'that\'s a normal retention conversation. Go to the Retention Playbook.', 'FiCornerUpRight');
  }

  {
    const s = content(CI, 'Get the opening right', 'Know something, not everything.',
      'There are two ways to get the opening wrong, in opposite directions. 1) Pretend you don\'t know anything and wait for the customer to bring it up: it risks looking like we ignored a clear signal, and delays a resolution the customer already wants. 2) Open by reciting back everything the customer told the CP, as if you\'d been reading over their shoulder: it can feel surveillance-y, even though checking C/CP comms is normal investigation work. The right move sits between: acknowledge that we know something, without acting like we know everything. It removes the awkwardness of pretending ignorance and still leaves the customer room to tell their own side in their own words, which makes a retainable conversation more likely.');
    const w = (9.1 - 0.3) / 3;
    const cols = [
      ['FiEyeOff', 'Pretend you don\'t know', 'Looks like we ignored a clear signal, and delays what they want.', PINK, RED, '✕'],
      ['FiCheckCircle', 'Know something, not everything', 'No awkward pretending, and they still tell it in their own words.', C.tealSoft, C.teal, '✓'],
      ['FiEye', 'Recite everything they told the CP', 'Feels like we were reading over their shoulder.', PINK, RED, '✕'],
    ];
    for (const [i, [ico, t, b, soft, col, mark]] of cols.entries()) {
      const x = 0.45 + i * (w + 0.15);
      box(s, x, 1.45, w, 2.2, soft, col);
      await iconDot(s, x + 0.2, 1.65, ico, 0.42, C.white);
      T(s, mark, { x: x + w - 0.55, y: 1.62, w: 0.4, h: 0.45, fontFace: HEAD, bold: true, fontSize: 20, color: col, align: 'center' });
      T(s, t, { x: x + 0.2, y: 2.2, w: w - 0.4, h: 0.5, fontFace: HEAD, bold: true, fontSize: 12, color: C.ink, valign: 'top' });
      T(s, b, { x: x + 0.2, y: 2.75, w: w - 0.4, h: 0.8, fontSize: 9.5, color: C.soft });
    }
    quote(s, 0.45, 3.85, 9.1, 0.6, 'Hey [Name], your cleaner mentioned you might be looking to cancel your membership. Wanted to check in directly.', 'OPEN WITH');
  }

  {
    const s = content(CI, 'The process', 'Reach out, let them explain, then it\'s a normal retention conversation.',
      'Step 1, reach out proactively: don\'t wait for the customer to raise it. Open with neutral framing that signals "we know something", not "we know everything": "Hey [Name], your cleaner mentioned you might be looking to cancel your membership. Wanted to check in directly." Step 2, probe, don\'t assume: give the customer the chance to explain in their own words rather than resolving based on whatever fragment showed up in the CP thread: "Do you mind sharing why? That\'ll help me figure out the best way to handle this for you." This mirrors the General Retention Playbook\'s "Probing, even when you already know the reason". Step 3, explain FC and ETF terms: since the customer has a cancellation intent, if the ETF applies, explain it. Don\'t skip this even if they sound set on cancelling; it\'s often the first time they hear the terms explained clearly. Step 4, handle upcoming appointments and RC: by default cancel upcoming appointments and RC, unless there\'s a clear signal they want to keep them. Step 5, address other pain points, then route to the right playbook. Note: there is no macro for this opening yet; it\'s been flagged for the Comms Kit.');
    const st = [['FiSend', 'Reach out', 'Don\'t wait for them to raise it.'], ['FiHelpCircle', 'Probe', 'Let them explain in their own words.'], ['FiFileText', 'Explain FC & ETF', 'If the ETF applies. Don\'t skip it.'], ['FiCalendar', 'Appointments & RC', 'Cancel by default, unless they want to keep them.'], ['FiGitBranch', 'Route', 'Address pain points, then the right playbook.']];
    const w = (9.1 - 0.4) / 5;
    for (const [i, [ico, t, b]] of st.entries()) {
      const x = 0.45 + i * (w + 0.1);
      box(s, x, 1.45, w, 1.85, C.white, C.border);
      badge(s, x + 0.12, 1.57, i + 1);
      await iconDot(s, x + w - 0.5, 1.55, ico, 0.36);
      T(s, t, { x: x + 0.12, y: 2.05, w: w - 0.24, h: 0.4, fontFace: HEAD, bold: true, fontSize: 11, valign: 'top' });
      T(s, b, { x: x + 0.12, y: 2.47, w: w - 0.24, h: 0.8, fontSize: 9, color: C.soft });
    }
    quote(s, 0.45, 3.5, 9.1, 0.6, 'Do you mind sharing why? That\'ll help me figure out the best way to handle this for you.', 'STEP 2: PROBE');
    await tip(s, 4.25, 'No macro yet:', 'use the Step 1–2 language. It\'s flagged for the Comms Kit.', 'FiMessageSquare');
  }

  {
    const s = content(CI, 'Route to the right playbook', 'Once they\'ve told you why, follow the root cause.',
      'Once the customer has told you why, this becomes a normal retention conversation: the "how we found out" part is done. Membership-terms objection (unaware, cost, MCT length): Membership Retention Playbook. Service issue (quality, reliability): Poor Cleaning Quality Retention Playbook or Service Reliability Retention Playbook. General retainability assessment: General Retention Playbook. For explaining FC and ETF terms: ForeverClean Membership. Washington customers: use Washington State: Temporary FC Cancellation Handling, Step 5. Reach out, send the self-cancellation link and never mention the ETF.');
    const rows = [
      ['FiFileText', 'Membership terms', 'Unaware, cost, MCT length', 'Membership Retention Playbook'],
      ['FiStar', 'Service issue', 'Quality or reliability', 'PCQ or Service Reliability Retention Playbook'],
      ['FiCompass', 'Is this customer retainable?', 'The five key factors', 'General Retention Playbook'],
    ];
    for (const [i, [ico, t, b, pb]] of rows.entries()) {
      const y = 1.45 + i * 0.78;
      box(s, 0.45, y, 4.1, 0.68, C.white, C.border);
      await iconDot(s, 0.58, y + 0.16, ico, 0.36);
      T(s, [{ text: t, options: { bold: true, fontFace: HEAD, fontSize: 11, breakLine: true } }, { text: b, options: { fontSize: 9, color: C.soft } }], { x: 1.05, y, w: 3.4, h: 0.68, valign: 'middle' });
      flowLine(s, [[4.55, y + 0.34], [5.0, y + 0.34]], true);
      s.addText(pb, { shape: pres.shapes.ROUNDED_RECTANGLE, x: 5.0, y, w: 4.55, h: 0.68, rectRadius: 0.08, fill: { color: C.teal }, line: { color: C.teal },
        fontFace: HEAD, bold: true, fontSize: 10.5, color: C.white, align: 'center', valign: 'middle', margin: 0 });
    }
    await warn(s, 3.85, 'Washington customer?', 'use the WA process: send the self-cancel link and never mention the ETF.');
  }

  await knowledgeCheck(CI, [
    ['While checking a PCQ, you see the C told the CP they want to cancel. Wait for them to contact us?', 'No. Reach out proactively.'],
    ['How do you open the conversation?', 'Know something, not everything: "Your cleaner mentioned you might be looking to cancel…"'],
    ['The customer sounds set on cancelling. Skip explaining the ETF?', 'No. If it applies, explain FC and ETF terms.'],
    ['What happens to upcoming appointments and RC?', 'Cancel by default, unless they clearly want to keep them.'],
  ], 'Ask each question and let trainees answer before revealing. Each click reveals the next answer.', KC_TITLE, '');

  await practice(CI, [
    ['Where did the cancellation intent come from?', 'And what is the customer\'s real reason?'],
    ['What actions are you going to take?', 'Opening, probe, terms, appointments, playbook.'],
    ['5-min writing challenge', 'Create your outreach message.'],
  ], 'TRAINER: share a live ticket where the customer told the CP they want to cancel. Give trainees a few minutes to review it on their own, then work through it together. 1) Where did the intent come from, and what is the real reason? Don\'t resolve from the fragment in the CP thread. 2) What actions are you going to take? Reach out proactively (know something, not everything), probe, explain FC/ETF terms if the ETF applies, cancel upcoming appointments and RC by default, then route to the right playbook. WA customer: use the WA process. 3) 5-minute writing challenge: each trainee writes the outreach message. Read a few aloud and compare. Then open the floor for questions.');

  // ================= 4. GENERAL RETENTION PLAYBOOK =================
  const GR = 'Retention Playbook';
  await topic('General Retention Playbook', 'The thinking behind every retention conversation.',
    ['Fix the real problem, not just offer something', 'The 5 key factors: weigh, don\'t count', 'Matching and presenting the offer', 'When retention isn\'t the right call'],
    'ill10.png',
    'General Retention Playbook. Category: Retention and Membership Cancellation. This is the foundational thinking behind every retention conversation: the mindset and assessment process the Poor Cleaning Quality, Standard Service Issue and Membership Retention Playbooks all build on. Read this first; the other playbooks assume you already know how to reason through a retention conversation, and just add the scenario-specific offers. Why it matters: two agents facing the identical situation can produce very different outcomes if one follows a checklist mechanically and the other understands why the checklist says what it says.');

  {
    const s = content(GR, 'Fix the real problem', 'Not just offer something.',
      'Every resolution should go straight to the root cause of why the customer is thinking about cancelling, not a generic offer thrown at the problem. This is a deliberate departure from the old "retention ladder" (start small, escalate only if the customer pushes back). Here, you match the offer to what the customer actually needs and show its value clearly. The goal is for the customer to feel understood and like their specific problem is being solved, not like they\'re being managed with incentives.');
    headPanel(s, 0.45, 1.45, 4.45, 2.1, 'Old: the retention ladder', ['Start small', 'Escalate only if they push back', 'Feels like being managed with incentives'], C.ink, C.white, 10.5);
    headPanel(s, 5.1, 1.45, 4.45, 2.1, 'Now: root cause first', ['Find why they want to leave', 'Match the offer to that reason', 'Show its value clearly'], C.teal, C.tealSoft, 10.5);
    await tip(s, 3.8, 'The goal:', 'the customer feels understood, and their specific problem gets solved.', 'FiTarget');
  }

  {
    const s = content(GR, 'The 5 key factors', 'Look at all five together.',
      'When judging whether a situation is retainable, look at all five factors together. Intent: is the customer seeking help or clarification, or have they already decided? "Can you help me understand this charge?" is open; "Cancel my membership immediately" is decided, but not automatically unretainable. Issue type: can the problem be fixed? Cleaner quality, scheduling mistakes, membership confusion are operational and fixable; repeated failures, misinformation and broken promises are harder but still worth attempting. Emotion: mild frustration vs strong, final language ("I want nothing to do with this company"), which is a signal to slow down and test openness, not proof. Escalation risk: frustrated with no material legal threat, de-escalate with clear, empathetic comms; AG threat, formal legal letter, lawyer CC\'d, class action, law firm email: retention is low, cancel FC and waive ETF instead. History: a first occurrence is easier; repeated bad experiences are harder but don\'t automatically mean don\'t try. Only stop once there are genuinely two consecutive job problems and the customer has declined assistance.');
    s.addTable([
      [{ text: 'Factor', options: { bold: true, color: C.teal, fill: { color: C.tealSoft }, fontFace: HEAD } },
        { text: 'Retainable signal', options: { bold: true, color: C.teal, fill: { color: C.tealSoft }, fontFace: HEAD } },
        { text: 'Harder to retain', options: { bold: true, color: RED, fill: { color: PINK }, fontFace: HEAD } }],
      ...[
        ['Intent', 'Asking for help or clarification', 'Already decided. Not automatically unretainable'],
        ['Issue type', 'Fixable: quality, scheduling, confusion', 'Repeated failures, broken promises. Still try'],
        ['Emotion', 'Mild frustration, wants it fixed', 'Strong, final language: slow down, test openness'],
        ['Escalation risk', 'Frustrated, no material legal threat', 'AG threat, legal letter, lawyer: cancel + waive ETF'],
        ['History', 'First occurrence', 'Repeated bad experiences. Still worth a try'],
      ].map(r => r.map((t, i) => ({ text: t, options: { color: i === 0 ? C.ink : C.soft, bold: i === 0, fill: { color: C.white } } })))
    ], { x: 0.45, y: 1.42, w: 9.1, colW: [1.6, 3.5, 4.0], rowH: 0.42, fontFace: BODY, fontSize: 9.5, valign: 'middle', border: { type: 'solid', pt: 0.75, color: C.border }, margin: [0.03, 0.1, 0.03, 0.1] });
    await tip(s, 4.15, 'History:', 'only stop after two consecutive job problems and the customer declined help.', 'FiClock');
  }

  {
    const s = content(GR, 'Weigh, don\'t count', 'Read the situation, not the scoreboard.',
      'It\'s not about counting how many factors lean one way; it\'s about weighing how much each present factor changes the overall picture. Some factors are minor; others, on their own, are enough to make retention the wrong call. If 3 of 5 factors point to "harder to retain", that does not automatically mean the customer isn\'t retainable, and the reverse is just as true. A customer could have a material legal threat and still have a first-time, easily fixable issue with mild frustration. That\'s not "1 of 5, still retainable": a material legal threat by itself makes this a "don\'t attempt retention" case. Confirmed abuse toward a cleaner works the same way. Compare Emotion alone: strong, final language is not an automatic override; it\'s a signal to slow down and test openness. The skill: knowing which signals are "weigh this alongside everything else" and which are "this alone changes the answer".');
    headPanel(s, 0.45, 1.45, 4.45, 2.0, 'Hard overrides: this alone decides', ['Material legal threat', 'Confirmed abuse toward a cleaner', 'Retention is off the table, whatever the other factors say'], RED, PINK, 10);
    headPanel(s, 5.1, 1.45, 4.45, 2.0, 'Weigh it: a signal to slow down', ['Strong, final language', 'Already decided to cancel', 'Repeated failures', 'Test openness. Not a disqualifier'], C.teal, C.tealSoft, 10);
    await tip(s, 3.65, '3 of 5 "harder"?', 'can still be worth a real attempt. 4 of 5 "retainable" can still be a no.', 'FiSliders');
  }

  {
    const s = content(GR, 'Four steps', 'Acknowledge, understand, align, then offer with purpose.',
      'Step 1, acknowledge the customer\'s feedback: make sure they feel heard and share empathy for a frustrating experience. If they request an action (cancellation, refund) after a disappointing experience, acknowledge the request, then probe if their first message was vague. Step 2, understand the root cause first: if they want to cancel because they were unaware of the membership, start by explaining FC details and ETF terms; don\'t jump to a discount. Probe their actual experience to uncover underlying issues. Don\'t offer MF or MCT reductions in an unawareness case: they don\'t address the root cause. Step 3, align the offer with the root cause: service-related (e.g. PCQ), a voucher or credit on first response is appropriate; cost or commitment concerns (MF or MCT), a reduction right away makes sense. Step 4, offers are tools, not defaults: MF/MCT reductions, vouchers, free months, credits shouldn\'t be handed out reflexively. Always ask: does this offer answer the specific reason they want to leave?');
    await steps(s, 1.45, 1.7, [
      ['Acknowledge', 'Make them feel heard. Acknowledge the request, then probe.'],
      ['Find the root cause', 'Unaware of FC? Explain FC & ETF first. No discount yet.'],
      ['Align the offer', 'Match the offer to the reason they want to leave.'],
      ['Tools, not defaults', 'Does this offer answer their specific reason?'],
    ]);
    table(s, ['Root cause', 'Right move on first response'], [
      ['Service-related (e.g. PCQ)', 'Offer a voucher or credit right away'],
      ['Cost or commitment (MF or MCT)', 'Offer a reduction right away'],
    ], { y: 3.35, colW: [3.6, 5.5], fontSize: 10, rowH: [0.34, 0.38, 0.38] });
    await tip(s, 4.6, 'Unaware of the membership?', 'MF/MCT reductions don\'t fix that. Explain first.', 'FiInfo');
  }

  {
    const s = content(GR, 'Why the ladder fails', 'Any stronger offer must be tied to their reason, not to their "no".',
      'Offering something small, the customer declines, offering more, repeat: this creates a transactional, pushy experience that reduces trust. It also teaches customers that declining gets them a better offer, which isn\'t fair to customers who accept the first reasonable one. The fix isn\'t "never increase an offer"; any stronger offer must be tied to the customer\'s specific reason for wanting to cancel, not offered just because they said no to the last one. Also: an offer that doesn\'t address the customer\'s actual concern is unlikely to retain them and can cheapen the perceived value of the membership.');
    const w = 2.1, gap = 0.233;
    const flow = [['Small offer', 'end'], ['They decline', 'decision'], ['Bigger offer', 'end'], ['Repeat…', 'penalty']];
    for (let i = 0; i < flow.length; i++) {
      const x = 0.45 + i * (w + gap);
      await flowBox(s, x, 1.5, w, 0.55, flow[i][0], flow[i][1]);
      if (i < flow.length - 1) flowLine(s, [[x + w, 1.775], [x + w + gap, 1.775]], true);
    }
    await cards(s, 2.3, 1.45, [
      { ico: 'FiThumbsDown', title: 'Feels pushy', body: 'Transactional, and it reduces trust.' },
      { ico: 'FiAward', title: 'Rewards saying no', body: 'Unfair to customers who accept the first reasonable offer.' },
      { ico: 'FiTrendingDown', title: 'Cheapens FC', body: 'An offer that misses their concern lowers its value.' },
    ]);
    await tip(s, 3.95, 'Stronger offer?', 'only when it\'s tied to their specific reason, not because they said no.', 'FiLink');
  }

  {
    const s = content(GR, 'Presenting an offer', 'Present it and wait, or issue it right away. Just be clear which.',
      'Two ways to deliver an offer. Apply it immediately: fast, but can feel transactional, "here\'s your solution" without involving the customer. Present the offer and wait for a response: encourages collaboration, gives the customer a sense of choice, shows you care about their experience rather than closing a ticket. General rule: it\'s up to you whether to present first or issue right away; just make it clear in your comms whether the offer has been issued on the account yet. Issue right away when (not exhaustive): the customer asked for it (voucher, credits, free months); or the monthly fee is overdue or the next MF is almost due, because presenting first risks them being charged, which aggravates a customer with cancellation intent. Use a "no… but" approach when a request is beyond policy (a full refund, an ETF waiver they don\'t qualify for): "I\'m not able to [X], but what I can do is [Y]". It stays empathetic and solution-focused within policy, without a flat refusal.');
    headPanel(s, 0.45, 1.45, 4.45, 1.45, 'Present and wait', ['Collaborative: gives them a choice', 'Shows you care about their experience'], C.teal, C.tealSoft);
    headPanel(s, 5.1, 1.45, 4.45, 1.45, 'Issue right away when…', ['The customer asked for it', 'The MF is overdue or almost due'], C.gold, C.goldSoft);
    quote(s, 0.45, 3.05, 9.1, 0.6, 'I\'m not able to [X], but what I can do is [Y].', 'BEYOND POLICY? USE "NO… BUT"');
    await tip(s, 3.85, 'Either way:', 'say clearly whether the offer is already on the account or not.', 'FiCheckSquare');
  }

  {
    const s = content(GR, 'Who isn\'t retainable?', 'The list is narrower than you think.',
      'Look at the full picture. Emotional escalation, distrust or frustration does not automatically mean a customer is non-retainable; many calm down once their concerns are acknowledged and addressed. As a general rule, attempt thoughtful retention unless there are clear signals the customer is no longer a fit. Affordability concerns alone are not a reason to stop: a customer who says they can\'t afford the fee may still be retainable through an MF/MCT reduction, pause, or free months. The genuine non-retainable cases are narrower: no ongoing need for recurring cleanings at all (no longer have a home, moved to a care facility); severe, extreme financial hardship where even a reduced fee isn\'t reasonable (handle with empathy, not a sales-style reduction); a material legal threat (filing with the Attorney General, joining a litigation, formal legal proceedings). Bottom line: everyone else, even an angry customer or someone who says the fee is too expensive, may still be a fit. Even if you don\'t retain them today, a great Care experience can leave the door open for them to return.');
    await cards(s, 1.45, 1.75, [
      { ico: 'FiHome', title: 'No ongoing need', body: 'No longer has a home, or moved to a care facility.' },
      { ico: 'FiAlertOctagon', title: 'Extreme hardship', body: 'Even a reduced fee isn\'t reasonable. Empathy, not a sales pitch.' },
      { ico: 'FiFileText', title: 'Material legal threat', body: 'AG complaint, litigation, formal legal proceedings.' },
    ]);
    headPanel(s, 0.45, 3.35, 9.1, 1.1, 'Still worth a real attempt', ['Angry, distrustful or frustrated customers: many calm down once heard', '"I can\'t afford it": MF/MCT reduction, a pause or free months may work'], C.teal, C.tealSoft);
  }

  {
    const s = content(GR, 'Usually not a retention target', 'Common patterns, not automatic answers.',
      'Beyond the general non-retainable signals, these situations are in most cases also not good retention targets: history of filing disputes; bankruptcy or extreme financial hardship; military deployment; ADA accommodation needs; severe weather or natural disaster; severe illness; death of a household or family member; relocation outside of coverage; you\'ve made a genuine, well-positioned retention attempt addressing the root issue and they still want to cancel; multiple unused MFs (beyond MCT) and firm, escalating language about cancelling; a material legal threat; abuse toward a cleaner (see Customer Abusive Behavior Towards a Cleaner: retention is explicitly off the table, and only exit-focused offers, ETF/MCT reduction, remain). Why these are different from "just difficult": their situation isn\'t actually about the service, so pushing retention risks looking tone-deaf.');
    const items = ['History of filing disputes', 'Bankruptcy / extreme hardship', 'Military deployment', 'ADA accommodation needs', 'Severe weather / disaster', 'Severe illness', 'Death in the household', 'Moving outside coverage', 'Genuine attempt already declined', 'Multiple unused MFs + firm', 'Material legal threat', 'Abusive toward a cleaner'];
    const w = (9.1 - 0.3) / 3;
    items.forEach((t, i) => {
      const x = 0.45 + (i % 3) * (w + 0.15), y = 1.45 + Math.floor(i / 3) * 0.5;
      const hard = i >= 10;
      chip(s, x, y, w, t, hard ? PINK : C.white, hard ? RED : C.ink, 0.4);
      if (!hard) s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.4, rectRadius: 0.08, fill: { type: 'none' }, line: { color: C.border, width: 0.75 } });
    });
    await tip(s, 3.6, 'Abusive toward a cleaner:', 'retention is off the table. Only exit offers (ETF/MCT reduction) remain.', 'FiSlash');
    await tip(s, 4.12, 'Why:', 'their situation isn\'t about the service. Pushing retention looks tone-deaf.', 'FiInfo');
  }

  {
    const s = content(GR, 'Context decides', 'The same situation can point to a different call. Investigate first.',
      'This isn\'t a black-and-white list; each one depends on what the customer is actually asking for and why. History of filing disputes: if they now want to cancel or get a refund, it\'s a genuine non-retainable case, since there\'s a real chance they\'ll dispute with their bank again. But if they still want to use the service, we don\'t push them out, unless something points to fraud. Severe illness: if they want to cancel because of it, cancel out of compassion and waive the ETF as a courtesy; no retention. But if the illness is why they use the service, attempt to retain, since cancelling removes help they need. Military deployment: deployed and want to cancel now, a straightforward non-retainable case. Just a heads-up that they\'ll deploy in a couple of months and still using the service: don\'t cancel and waive on the spot; set a Timed Reminder to cancel FC once the deployment happens. The pattern: these lean non-retainable, but the call depends on what the customer is asking for right now.');
    const rows = [
      ['FiAlertCircle', 'History of disputes', 'Wants to cancel or a refund → non-retainable', 'Wants to keep using it → let them (unless fraud)'],
      ['FiActivity', 'Severe illness', 'Cancelling because of it → cancel, waive ETF', 'Needs the help because of it → retain'],
      ['FiFlag', 'Military deployment', 'Deploying now → cancel', 'Deploying in a few months → Timed Reminder'],
    ];
    T(s, 'Leans non-retainable', { x: 2.75, y: 1.38, w: 3.3, h: 0.22, fontSize: 8.5, bold: true, color: RED });
    T(s, 'But…', { x: 6.25, y: 1.38, w: 3.3, h: 0.22, fontSize: 8.5, bold: true, color: C.teal });
    for (const [i, [ico, t, a, b]] of rows.entries()) {
      const y = 1.62 + i * 0.8;
      box(s, 0.45, y, 2.2, 0.7, C.white, C.border);
      await iconDot(s, 0.55, y + 0.17, ico, 0.36);
      T(s, t, { x: 1.0, y, w: 1.6, h: 0.7, fontFace: HEAD, bold: true, fontSize: 10.5, valign: 'middle' });
      box(s, 2.75, y, 3.4, 0.7, PINK);
      T(s, a, { x: 2.9, y, w: 3.15, h: 0.7, fontSize: 9.5, valign: 'middle' });
      box(s, 6.25, y, 3.3, 0.7, C.tealSoft);
      T(s, b, { x: 6.4, y, w: 3.05, h: 0.7, fontSize: 9.5, valign: 'middle' });
    }
    await tip(s, 4.15, 'Ask:', 'what is the customer asking for right now? Not just which situation is on the account.', 'FiSearch');
  }

  {
    const s = content(GR, 'Clear signs they\'re retainable', 'The objection is about structure, not value.',
      'Satisfied with the cleaning itself ("The cleaner did a good job but I don\'t want a membership"): the core service worked; the objection is about structure, not value. A temporary or addressable financial or situational concern ("I can\'t afford the monthly fee right now"): may be resolvable via fee adjustment or scheduling flexibility. Use judgment: genuine severe financial distress (job loss, bankruptcy) may not be appropriate for retention; consider pausing or assisting with cancellation instead. Policy or membership confusion ("I didn\'t realize I signed up for a 6-month membership"): frustration comes from a misunderstanding, not the value; clarify calmly, probe for the actual service experience, and continue retention if they calm down.');
    const c = [
      ['FiSmile', 'Happy with the cleaning', 'The cleaner did a good job but I don\'t want a membership.', 'The service worked. It\'s about structure.'],
      ['FiDollarSign', 'Temporary money concern', 'I can\'t afford the monthly fee right now.', 'Fee adjustment or flexibility. Severe distress? Pause or help cancel.'],
      ['FiHelpCircle', 'Membership confusion', 'I didn\'t realize I signed up for 6 months.', 'Clarify calmly, probe the service, keep going.'],
    ];
    const w = (9.1 - 0.3) / 3;
    for (const [i, [ico, t, q, b]] of c.entries()) {
      const x = 0.45 + i * (w + 0.15);
      box(s, x, 1.45, w, 2.75, C.white, C.border);
      await iconDot(s, x + 0.18, 1.62, ico, 0.42);
      T(s, t, { x: x + 0.18, y: 2.12, w: w - 0.36, h: 0.32, fontFace: HEAD, bold: true, fontSize: 12 });
      box(s, x + 0.18, 2.5, w - 0.36, 0.75, C.tealSoft);
      T(s, '“' + q + '”', { x: x + 0.28, y: 2.5, w: w - 0.56, h: 0.75, fontSize: 9.5, italic: true, valign: 'middle' });
      T(s, b, { x: x + 0.18, y: 3.35, w: w - 0.36, h: 0.8, fontSize: 9.5, color: C.soft });
    }
  }

  {
    const s = content(GR, 'Genuine unhappiness', 'Don\'t assume they\'re non-retainable. Be human first.',
      'Extreme emotional escalation ("This is a scam, cancel this NOW"): don\'t open with a retention offer; it can escalate the anger. Be human first and probe gently: "I\'m so sorry to hear that, I\'m here to help! What happened?" Trust broken by repeated failures ("I gave your service more than one chance and keep being disappointed"): acknowledge the pattern honestly, offer to waive fees given the history, and only then gently offer one more attempt. Full sample: "I\'m so sorry to hear about the terrible experience. Of course, I can process your cancellation, and given your experience we\'ll waive all fees. I know I\'m totally pushing my luck here, but I really hate to see you go like this. I\'d love to try and set you up with someone great and more reliable. But again, I totally understand if you\'d prefer to go, just say the word." The customer feels deceived or trapped ("You never told me about the membership… I no longer trust your company"): explain FC details clearly and factually, then shift to probing their cleaning experience rather than re-litigating the disclosure. If their tone shifts from accusatory to practical, retention may still be viable. Careful: pushing retention prematurely in a trust-broken conversation increases dispute risk.');
    const rows = [
      ['FiZap', 'Extreme escalation', '"This is a scam, cancel this NOW"', 'No offer yet. Be human, probe gently.', 'I\'m so sorry to hear that, I\'m here to help! What happened?'],
      ['FiRepeat', 'Repeated failures', '"I keep being disappointed"', 'Waive fees first, then gently offer one more try.', 'Given your experience we\'ll waive all fees… I\'d love to set you up with someone great.'],
      ['FiLock', 'Feels deceived or trapped', '"You never told me about the membership"', 'Explain FC factually, then probe the cleaning.', 'Clarity and empathy before any solution.'],
    ];
    for (const [i, [ico, t, said, move, q]] of rows.entries()) {
      const y = 1.42 + i * 1.0;
      box(s, 0.45, y, 9.1, 0.9, C.white, C.border);
      await iconDot(s, 0.6, y + 0.24, ico, 0.42);
      T(s, [{ text: t, options: { bold: true, fontFace: HEAD, fontSize: 11, breakLine: true } }, { text: said, options: { italic: true, fontSize: 9, color: C.soft } }], { x: 1.15, y, w: 2.75, h: 0.9, valign: 'middle' });
      T(s, move, { x: 3.95, y, w: 2.4, h: 0.9, fontSize: 9.5, bold: true, color: C.teal, valign: 'middle' });
      box(s, 6.45, y + 0.1, 3.0, 0.7, C.tealSoft);
      T(s, i < 2 ? '“' + q + '”' : q, { x: 6.55, y: y + 0.1, w: 2.85, h: 0.7, fontSize: 8.5, italic: i < 2, valign: 'middle' });
    }
    await tip(s, 4.45, 'Trust broken?', 'pushing retention too soon raises dispute risk.', 'FiAlertTriangle');
  }

  {
    const s = content(GR, 'The Newspaper Test', 'If this interaction were published, would we stand by it?',
      'When you\'re unsure whether a decision is right, ask: if this interaction were published publicly, would we stand by how we handled it? Worked example: a customer on a fixed income (e.g. a disabled veteran) wants to cancel because they can\'t afford ongoing service, even though they were satisfied with the cleaning. Strict enforcement: "Customer didn\'t read the terms, enforce the ETF." If this became public, would it look fair, or like taking advantage of someone already struggling? Retention-aware: recognize the financial limitation, apply compassion where appropriate, consider an ETF waiver. The test in one line: if enforcing a rule would feel gross, deceptive or shameful to explain publicly, don\'t enforce it. That\'s exactly why ETF waivers exist for compassion and extreme hardship cases.');
    box(s, 0.45, 1.45, 3.4, 2.65, C.teal);
    s.addImage({ data: await icon('FiBookOpen', C.white), x: 0.75, y: 1.75, w: 0.5, h: 0.5 });
    T(s, 'If enforcing a rule would feel gross, deceptive or shameful to explain publicly, don\'t enforce it.', { x: 0.75, y: 2.4, w: 2.85, h: 1.55, fontFace: HEAD, bold: true, fontSize: 13, color: C.white, valign: 'top' });
    eyebrow(s, 'Example: fixed income, happy with the cleaning, can\'t afford it', 4.05, 1.45, 5.5, C.soft);
    headPanel(s, 4.05, 1.75, 2.65, 2.35, 'Strict enforcement', ['"Didn\'t read the terms: enforce the ETF"', 'Looks like taking advantage of someone struggling'], RED, PINK);
    headPanel(s, 6.9, 1.75, 2.65, 2.35, 'Retention-aware', ['Recognize the limitation', 'Apply compassion', 'Consider an ETF waiver'], C.teal, C.tealSoft);
    await tip(s, 4.3, 'That\'s why', 'ETF waivers exist for compassion and extreme hardship cases.', 'FiHeart');
  }

  {
    const s = content(GR, 'Probe, even when you know why', 'Let the customer say it to you directly.',
      'If a customer already mentioned their cancellation reason to someone else (told the CP, not you), still probe. This isn\'t ignoring information you have; it gives the customer the chance to express it directly to you. Customers given the opportunity to explain their own situation are more likely to be retained. Weak (assumption-based): "I\'m sorry to hear you want to cancel, do you mind sharing why?" Strong (customer-led): "Your cleaner mentioned you want to cancel the service, but I\'d like to hear directly from you. Would you mind sharing what led you to consider canceling?"');
    quote(s, 0.45, 1.45, 4.45, 1.25, 'I\'m sorry to hear you want to cancel, do you mind sharing why?', 'WEAK: ASSUMPTION-BASED');
    quote(s, 5.1, 1.45, 4.45, 1.25, 'Your cleaner mentioned you want to cancel the service, but I\'d like to hear directly from you. Would you mind sharing what led you to consider canceling?', 'STRONG: CUSTOMER-LED');
    chip(s, 0.45, 2.85, 1.0, '✕  Weak', PINK, RED);
    chip(s, 5.1, 2.85, 1.0, '✓  Strong', C.tealSoft, C.teal);
    await tip(s, 3.4, 'Why:', 'customers who explain their own situation are more likely to stay.', 'FiMessageCircle');
  }

  {
    const s = content(GR, 'Don\'t close too soon', 'Graceful Closure can\'t be reopened. Make sure you\'ve explored everything.',
      'Graceful Closure is for when a customer declines your resolution and keeps insisting. Before you invoke it, ask: 1) Have I actually explored all the options? 2) Do I really know the full picture? Once you move to Graceful Closure you can\'t reopen the conversation or offer new concessions (the No Re-Opening Rule). A common mistake: moving to closure the moment a customer declines, without probing why, or treating a non-response to a probing question as "no further options". If they aren\'t responding to probing, reframe why you\'re asking: "I understand you\'d like to cancel. Would you mind sharing what led to your decision? I just want to make sure I\'ve explored all available options and that we\'re not missing anything that might better fit your needs." If they still won\'t share a reason: keep the conversation open and keep gently asking (no cap while they\'re engaging), but don\'t move to Graceful Closure; you can\'t honestly say "we\'ve reviewed all your options". Instead explain the ETF still applies, on its own terms: "I understand you\'d like to cancel. The early termination fee still applies here; it helps make sure your cleaner is fairly paid for the discounted/free first cleaning. Just to set expectations: your membership will remain active and the monthly fee will continue until the cancellation is processed. That said, if there\'s something specific going on, let me know; it may open up more I can do."');
    eyebrow(s, 'Before Graceful Closure, ask', 0.45, 1.42, 5, C.soft);
    checkRows(s, 1.65, [['Explored all options?', 'Not just the one they declined.'], ['Know the full picture?', 'No reply to a probe ≠ no options left.']], 0.5, 0.08);
    quote(s, 0.45, 2.85, 9.1, 0.75, 'I just want to make sure I\'ve explored all available options and that we\'re not missing anything that might better fit your needs.', 'NOT RESPONDING? REFRAME');
    await warn(s, 3.75, 'Still no reason?', 'keep asking, don\'t close. Explain the ETF still applies, on its own terms.', 0.5);
  }

  {
    const s = content(GR, 'What not to do', 'Four mistakes that cost trust or money.',
      'Generally we don\'t refund an MF and waive the ETF at the same time: both cover the same thing (the discounted first cleaning), so removing both is a real financial loss and a policy inconsistency. If a customer asks for both: "I\'m unable to refund the membership fee since the account was active, but I can apply that recent payment toward reducing your Early Termination Fee." Exemptions: extenuating circumstances, material legal threats, a customer with a history of filing a dispute; see the MF & ETF Refund Request article. Don\'t introduce membership-related offers if the customer hasn\'t raised membership as a concern; stay on their stated issue and don\'t manufacture a cancellation conversation. Don\'t repeat the same offer or reopen closed negotiations; see Graceful Closure. Related articles: Poor Cleaning Quality Retention Framework, Standard Service Issue Retention Playbook, Membership Retention Playbook, ETF Waivers, Graceful Closure, Customer Abusive Behavior Towards a Cleaner.');
    await ztpLike(s);
    async function ztpLike(s) {
      const items = [
        ['FiCopy', 'Refund the MF and waive the ETF', 'Both cover the discounted first cleaning. Apply the MF toward the ETF instead.'],
        ['FiMessageSquare', 'Raise membership unprompted', 'Stay on their stated issue. Don\'t manufacture a cancellation.'],
        ['FiRepeat', 'Repeat the same offer', 'Or reopen closed negotiations. See Graceful Closure.'],
        ['FiSkipForward', 'Close too soon', 'Probe first. You can\'t reopen after Graceful Closure.'],
      ];
      const w = (9.1 - 0.15) / 2, h = 1.05;
      for (const [i, [ico, t, b]] of items.entries()) {
        const x = 0.45 + (i % 2) * (w + 0.15), y = 1.45 + Math.floor(i / 2) * (h + 0.12);
        box(s, x, y, w, h, C.white, C.border);
        await iconDot(s, x + 0.18, y + 0.18, ico, 0.4, PINK);
        T(s, [{ text: 'Don\'t: ' + t, options: { bold: true, fontFace: HEAD, fontSize: 11, breakLine: true } }, { text: b, options: { fontSize: 9.5, color: C.soft } }], { x: x + 0.72, y: y + 0.08, w: w - 0.85, h: h - 0.16, valign: 'middle' });
      }
    }
    await tip(s, 3.85, 'Exemptions to MF + ETF:', 'extenuating circumstances, material legal threats, a history of disputes.', 'FiInfo');
  }

  await knowledgeCheck(GR, [
    ['A customer threatens to file with the Attorney General, but the issue is first-time and fixable. Retain?', 'No. A material legal threat alone overrides. Cancel FC and waive the ETF.'],
    ['"I can\'t afford this." Stop attempting retention?', 'No. Affordability alone isn\'t a reason. Try an MF/MCT reduction, pause or free months.'],
    ['They declined your offer. Offer a bigger one?', 'Only if it\'s tied to their specific reason, not just because they said no.'],
    ['They want an MF refund and an ETF waiver. Both?', 'Generally no. Apply the MF toward reducing the ETF (unless an exemption applies).'],
  ], 'Ask each question and let trainees answer before revealing. Each click reveals the next answer.', KC_TITLE, '');

  await practice(GR, [
    ['Is the customer retainable? Why?', 'Weigh the five key factors. Any hard override?'],
    ['What actions are you going to take?', 'Root cause, the matching offer, how you\'ll present it.'],
    ['5-min writing challenge', 'Create your comms for the customer.'],
  ], 'TRAINER: share a live cancellation ticket. Give trainees a few minutes to review it on their own, then work through it together. 1) Is the customer retainable? Weigh the five key factors (Intent, Issue type, Emotion, Escalation risk, History); check for hard overrides (material legal threat, abuse toward a cleaner) and the "usually not a retention target" situations, then confirm what the customer is asking for right now. 2) What actions are you going to take? Acknowledge, find the root cause, align the offer with it, decide whether to present it or issue it right away, and avoid the ladder. Run the Newspaper Test if unsure. 3) 5-minute writing challenge: each trainee writes the comms. Read a few aloud and compare. Then open the floor for questions.');

  // ================= 5. UNUSED DHJ VOUCHER =================
  const DV = 'Unused DHJ Voucher';
  await topic('Unused DHJ Voucher', 'They bought a DHJ voucher, but haven\'t used it. FC hasn\'t started yet.',
    ['Reading the account at this stage', 'Sorting the request into 4 cases', 'Retention offers for this customer', 'Refunds and closing the account'],
    'ill_pay.png',
    'Unused DHJ Voucher. Category: Retention and Membership Cancellation. This covers a customer who purchased a DHJ voucher but hasn\'t used it on a job yet: no cleaning has happened, so their FC membership hasn\'t technically started. This page is for refund and/or cancellation requests from these customers. If the customer has no refund or cancellation intent at all, only a gentle reminder is needed (case 1).');

  {
    const s = content(DV, 'What is this?', 'Bought, not used. The membership starts only after a completed job.',
      'For Legacy DHJ, the membership only starts once the voucher is applied to a completed job. If the customer only bought the voucher and hasn\'t booked yet, the voucher shows on the account but there\'s no FC table. The FC table (with active status) only appears once the customer books an appointment, and it may show as "active with 0 paid months". That doesn\'t mean FC has started. This distinction matters constantly: it\'s exactly what leads a customer to think they\'ve already committed to something they haven\'t, or an agent to assume a membership exists when it doesn\'t yet.');
    const w = 2.1, gap = 0.233;
    const flow = [['Buys the voucher', 'end', 'Voucher shows. No FC table'], ['Books a cleaning', 'decision', 'FC table appears: "active, 0 paid months"'], ['Cleaning completed', 'end', 'Voucher applied to the job'], ['Membership starts', 'action', 'Now FC has really begun']];
    for (let i = 0; i < flow.length; i++) {
      const x = 0.45 + i * (w + gap);
      await flowBox(s, x, 1.5, w, 0.55, flow[i][0], flow[i][1]);
      if (i < flow.length - 1) flowLine(s, [[x + w, 1.775], [x + w + gap, 1.775]], true);
      T(s, flow[i][2], { x, y: 2.12, w, h: 0.45, fontSize: 9, color: C.soft, align: 'center' });
    }
    await warn(s, 2.85, '"Active, 0 paid months"', 'doesn\'t mean FC has started. Not until the voucher is used on a completed job.', 0.5);
    await tip(s, 3.55, 'Why it matters:', 'customers think they\'ve committed when they haven\'t, and agents assume a membership exists.', 'FiInfo');
  }

  {
    const s = content(DV, 'What does the customer want?', 'Sort the request into one of four cases first. They lead to very different places.',
      'How you handle this depends entirely on what the customer actually wants. Sort into one of the cases before doing anything else. 1) No refund or cancellation intent: they reached out about something unrelated. 2) Refund or cancellation intent. 3) Wants the account closed or deactivated, with no refund intent. 4) Wants a refund, but already self-refunded the voucher before or during their first outreach.');
    const c = [
      ['FiMessageCircle', 'No refund or cancel intent', 'Help with their concern. Remind them once about the membership.', C.tealSoft, C.teal],
      ['FiRotateCcw', 'Wants a refund or to cancel', 'Probe, address the issue, present the right offer.', C.goldSoft, C.gold],
      ['FiUserX', 'Wants the account closed', 'Treat it as cancellation intent. Then close it properly.', C.white, C.ink],
      ['FiCheckCircle', 'Already self-refunded', 'Accept it. Give a clean, hassle-free cancellation.', PINK, RED],
    ];
    const w = (9.1 - 0.15) / 2, h = 1.3;
    for (const [i, [ico, t, b, soft, col]] of c.entries()) {
      const x = 0.45 + (i % 2) * (w + 0.15), y = 1.45 + Math.floor(i / 2) * (h + 0.15);
      box(s, x, y, w, h, soft, col);
      badge(s, x + 0.18, y + 0.2, i + 1);
      await iconDot(s, x + w - 0.6, y + 0.15, ico, 0.42, C.white);
      T(s, t, { x: x + 0.6, y: y + 0.15, w: w - 1.3, h: 0.42, fontFace: HEAD, bold: true, fontSize: 12.5, valign: 'middle' });
      T(s, b, { x: x + 0.6, y: y + 0.62, w: w - 0.9, h: 0.55, fontSize: 10, color: C.soft });
    }
  }

  {
    const s = content(DV, 'Case 1: No refund or cancel intent', 'Don\'t make it bigger than it needs to be.',
      'If the customer reaches out for something unrelated and shows no sign of wanting to cancel or get a refund, don\'t make this bigger than it needs to be. Aside from addressing their concern, gently remind them, once, that the voucher they purchased comes with a membership that starts once they use it, so they\'re not caught off guard later. If they\'re already clearly aware of this, skip the reminder. Including the terms is fine but not mandatory, unless the terms are part of what they\'re raising. For example, "I didn\'t realize this came with a monthly fee" is no longer a courtesy reminder: walk them through the actual terms (MF amount, MCT length, ETF), since that\'s the real issue.');
    await cards(s, 1.45, 1.35, [
      { ico: 'FiCheck', title: 'Help with their concern', body: 'That\'s why they reached out.' },
      { ico: 'FiBell', title: 'Remind them once', body: 'Skip it if they clearly already know.' },
      { ico: 'FiFileText', title: 'Terms are the issue?', body: 'Walk through MF, MCT and ETF.' },
    ]);
    quote(s, 0.45, 3.0, 9.1, 0.65, 'Just a quick reminder, your voucher comes with a membership which will start once it\'s used.', 'THIS IS ENOUGH');
    await tip(s, 3.85, '"I didn\'t realize there\'s a monthly fee"?', 'that\'s no longer a reminder. Explain the actual terms.', 'FiAlertCircle');
  }

  {
    const s = content(DV, 'Case 2: Refund or cancel intent', 'Know what you\'re solving before you offer anything.',
      'Check if the customer already gave a reason. If they did, go to the Retention Strategy Table for the applicable offer. If not, probe first; don\'t skip straight to offers. Then: acknowledge as necessary. Address the actual issue: a service issue means blocking the C/CP pairing and coaching or penalizing the CP as needed; a membership-terms issue means checking the Retention Strategy Table. Present retention offers where applicable. Remind them the voucher is valid for a year, so they can use it later. Remind them the voucher comes with a membership. If they have refund intent specifically, provide the self-refund link, and only then: tell them to enter the voucher code before clicking "Refund Voucher", and double-check the code in your message, since a wrong code breaks the refund. Check upcoming appointments and an active RC plan; if there\'s no sign they want to keep them, cancel via CRM with the correct cancellation reason code. If they decline and insist on a refund: refund the voucher through CRM (CRM > voucher > refund DHJ/FC) and invalidate it right away, advise the 5–10 business day timeframe, and deactivate the FC table if there is one.');
    const st = ['Reason given? If not, probe', 'Acknowledge', 'Fix the actual issue', 'Present the right offer', 'Valid for a year. Comes with a membership', 'Refund intent? Send the self-refund link', 'Cancel appointments and RC if not wanted'];
    st.forEach((t, i) => {
      const y = 1.42 + i * 0.44;
      box(s, 0.45, y, 4.8, 0.38, C.white, C.border);
      badge(s, 0.53, y + 0.04, i + 1);
      T(s, t, { x: 0.95, y, w: 4.2, h: 0.38, fontSize: 9.5, valign: 'middle', bold: i === 5 });
    });
    headPanel(s, 5.45, 1.42, 4.1, 1.45, 'Self-refund link', ['Only if they want a refund', 'Enter the voucher code before "Refund Voucher"', 'Double-check the code in your message'], C.gold, C.goldSoft, 9.5);
    headPanel(s, 5.45, 3.0, 4.1, 1.45, 'Still insists on a refund?', ['CRM > voucher > refund DHJ/FC, then invalidate', 'Refund in 5–10 business days', 'Deactivate the FC table if there is one'], C.ink, C.white, 9.5);
  }

  {
    const s = content(DV, 'Case 2: Which offer?', 'The Retention Strategy Table for unused DHJ vouchers.',
      'Retention Strategy Table. Service: service-related issues (CP no-show, CP cancellation, no CP claim): offer $10–$20 credits. Service limitations or personal reasons (found a different cleaner, wants same-day cleaning, moved, lack of phone support, will self-clean): no offers necessary. Membership-related: MF is expensive, offer to reduce the MF by $10–$15 off the original MF; MCT is too long, offer to reduce the MCT by 2 months from the original MCT; doesn\'t want a membership in general, no offers necessary, advise them to contact us if they want more flexible terms, and a one-time cleaning (full price, charged upfront) can be offered as applicable. One-time cleaning: offer at full price, charged upfront. Trial Cleaning: offer at regular price, charged upfront. Macros: Unused DHJ Voucher: Refund Request; Wants to "Try it Out" (Trial Cleaning); Wants One Time Cleaning.');
    s.addTable([
      ['Root cause', 'Condition', 'Offer'].map(t => ({ text: t, options: { bold: true, color: C.teal, fill: { color: C.tealSoft }, fontFace: HEAD } })),
      ...[
        ['Service', 'CP no-show, CP cancellation, no CP claim', '$10–$20 credits'],
        ['Service', 'Found another cleaner, wants same-day, moved, no phone support, will self-clean', 'No offer needed'],
        ['Membership', 'MF is expensive', 'MF −$10–$15 off the original MF'],
        ['Membership', 'MCT is too long', 'MCT −2 months from the original'],
        ['Membership', 'Doesn\'t want a membership at all', 'No offer. Mention flexible terms; OTC if it fits'],
        ['Alternative', 'Wants one cleaning only', 'One-time cleaning: full price, upfront'],
        ['Alternative', 'Wants to try it out', 'Trial Cleaning: regular price, upfront'],
      ].map(r => r.map((t, i) => ({ text: t, options: { color: i === 1 ? C.soft : C.ink, bold: i !== 1, fill: { color: C.white } } })))
    ], { x: 0.45, y: 1.42, w: 9.1, colW: [1.5, 4.5, 3.1], rowH: 0.38, fontFace: BODY, fontSize: 9.5, valign: 'middle', border: { type: 'solid', pt: 0.75, color: C.border }, margin: [0.03, 0.1, 0.03, 0.1] });
  }

  {
    const s = content(DV, 'Case 3: Close the account', 'No refund intent. Treat the first reach-out as cancellation intent.',
      'Treat the initial reach-out as a cancellation intent: acknowledge, probe for the reason, unsubscribe them and let them know they\'re off the mailing list, remind them the voucher is valid for a year and tied to a membership, and check upcoming appointments as in case 2. If they come back and insist on closing the account, first find out whether they also want their data deleted. Either way: invalidate the voucher (C Do > Voucher Invalidate); attempt to cancel FC and check whether the system flags that the voucher would be refunded (if it does, escalate to a TL to cancel and invalidate without a refund, then continue); send the closure email before deactivating, since you can\'t email them after; deactivate (CRM Do > Deactivate). If they want data deleted: credit card removal, escalate to a TL to remove the card from Stripe permanently; personal data removal, log it to the Data right to know/delete request tracker. If they ask for their personal information deleted specifically, clarify whether they want the membership cancelled, the data deleted, or both. With an unused voucher and an explicit data-deletion request, log the tracker ticket immediately. Macro: Deactivated (send before deactivating).');
    eyebrow(s, 'First reach-out', 0.45, 1.42, 4, C.soft);
    T(s, bullets(['Acknowledge, probe for the reason', 'Unsubscribe them from emails', 'Voucher valid for a year, tied to a membership', 'Check upcoming appointments']), { x: 0.45, y: 1.68, w: 3.0, h: 1.4, fontSize: 9.5, paraSpaceAfter: 4 });
    eyebrow(s, 'Still insists? Close it in this order', 3.7, 1.42, 6, C.soft);
    const st = [['Invalidate the voucher', 'C Do > Voucher Invalidate'], ['Try to cancel FC', 'System says the voucher would be refunded? TL cancels without refund'], ['Send the closure email', 'Before deactivating. You can\'t email them after'], ['Deactivate the account', 'CRM Do > Deactivate']];
    st.forEach(([t, b], i) => {
      const y = 1.68 + i * 0.5;
      box(s, 3.7, y, 5.85, 0.44, i === 2 ? C.goldSoft : C.white, i === 2 ? C.gold : C.border);
      badge(s, 3.8, y + 0.07, i + 1);
      T(s, [{ text: t + '  ', options: { bold: true, fontFace: HEAD, fontSize: 10 } }, { text: b, options: { fontSize: 8.5, color: C.soft } }], { x: 4.2, y, w: 5.25, h: 0.44, valign: 'middle' });
    });
    headPanel(s, 0.45, 3.8, 9.1, 0.9, 'Wants their data deleted too?', [[{ text: 'Card removal: ', options: { bold: true } }, { text: 'TL removes it from Stripe.  ' }, { text: 'Personal data: ', options: { bold: true } }, { text: 'log to the Data right to know/delete tracker (right away if they asked explicitly).' }]], C.teal, C.tealSoft, 9.5);
  }

  {
    const s = content(DV, 'Case 4: Already self-refunded', 'The retention attempt already ended. Make the exit clean.',
      'Most of the time, customers want to take matters into their own hands. If the customer already initiated a self-refund, the system might leave a voucher behind as a last-ditch effort to retain them. If that happens: invalidate the system-generated voucher, if the system hasn\'t already; cancel FC, the recurring cleaning plan and any upcoming appointments, if applicable; address any remaining concerns. Since the refund was initiated before or during their first outreach, the retention attempt has technically already failed before we could intervene. Accept the customer\'s decision and focus on a clean, hassle-free cancellation.');
    await cards(s, 1.45, 1.6, [
      { ico: 'FiXSquare', title: 'Invalidate the leftover voucher', body: 'The system may leave one behind to retain them.' },
      { ico: 'FiCalendar', title: 'Cancel FC, RC & appointments', body: 'Whatever is still active.' },
      { ico: 'FiMessageCircle', title: 'Address what\'s left', body: 'Any remaining concerns.' },
    ]);
    await tip(s, 3.3, 'Accept their decision:', 'focus on a clean, hassle-free cancellation. No retention push.', 'FiCheckCircle');
  }

  await knowledgeCheck(DV, [
    ['The FC table shows "active, 0 paid months". Has the membership started?', 'No. It starts only when the voucher is used on a completed job.'],
    ['A customer with an unused voucher asks about something unrelated. Mention the membership?', 'Yes, once, gently. Skip it if they clearly already know.'],
    ['They say the MF is too expensive. What can you offer?', 'Reduce the MF by $10–$15 off the original MF.'],
    ['When do you send the self-refund link?', 'Only when they want a refund. Enter the voucher code before "Refund Voucher".'],
  ], 'Ask each question and let trainees answer before revealing. Each click reveals the next answer.', KC_TITLE, '');

  await practice(DV, [
    ['Which of the 4 cases is this? Why?', 'And is there an FC table yet?'],
    ['What actions are you going to take?', 'Offer, refund link, refund, or close the account?'],
    ['5-min writing challenge', 'Create your comms for the customer.'],
  ], 'TRAINER: share a live ticket from a customer with an unused DHJ voucher. Give trainees a few minutes to review it on their own, then work through it together. 1) Which case is it (no intent, refund/cancel intent, close the account, already self-refunded), and what does the account show? 2) What actions are you going to take? Probe for the reason, fix the actual issue, use the Retention Strategy Table, remind them the voucher is valid for a year and comes with a membership, self-refund link only with refund intent, cancel appointments/RC if not wanted; or the account-closure order. 3) 5-minute writing challenge: each trainee writes the comms. Read a few aloud and compare. Then open the floor for questions.');

  // ================= 6. FCF PRE-INVOICE CANCELLATION =================
  const FP = 'FCF Pre-Invoice';
  await topic('FCF Pre-Invoice Cancellation', 'Cancelling FCF before the first job is invoiced, and whether to refund.',
    ['Why this only happens on FCF', 'The internal 24-hour window', 'The refund cheat sheet', 'How to process each refund'],
    'ill24.png',
    'FCF Pre-Invoice Cancellation and Refund Requests. Category: Retention and Membership Cancellation. This covers a customer wanting to cancel their FCF membership before their first job has been invoiced or completed. It\'s unique to FCF, since FCF membership starts immediately at sign-up rather than after a completed cleaning. A Legacy DHJ customer can\'t be in this situation: there\'s no membership to cancel until their first job happens.');

  {
    const s = content(FP, 'What is this?', 'FCF starts at sign-up, so customers can cancel before any cleaning.',
      'FCF customers can self-cancel from their dashboard even with zero completed jobs, something DHJ customers were never able to do, so this comes up often. The refund answer depends on a combination of factors (how long they\'ve had it, which cleaning length they chose, whose fault any failure was), not a single rule. Picking the wrong row means refunding something that shouldn\'t be refunded, or denying a refund the customer is entitled to. Retention comes first, always: only use the refund table once all reasonable retention efforts have been exhausted and the customer has decided to cancel. Pricing: $19 for the first month and $59 each month after, charged upfront before the first cleaning, renewing automatically whether or not they book.');
    headPanel(s, 0.45, 1.45, 4.45, 1.5, 'Legacy DHJ', ['Membership starts after a completed job', 'Nothing to cancel before the first cleaning'], C.ink, C.white, 10);
    headPanel(s, 5.1, 1.45, 4.45, 1.5, 'FCF', ['Membership starts at sign-up', 'Can self-cancel with zero completed jobs', '$19 first month, then $59, auto-renewing'], C.teal, C.tealSoft, 10);
    await warn(s, 3.15, 'Retention comes first, always.', 'Use the refund table only once they\'ve decided to cancel.', 0.5);
    await tip(s, 3.85, 'The refund depends on:', 'how long they\'ve had it, the cleaning length they chose, and whose fault any failure was.', 'FiLayers');
  }

  {
    const s = content(FP, 'The 24-hour window', 'Internal only. Never tell the customer.',
      'INTERNAL ONLY: the 24-hour window is something we use to decide what to do, not a rule we tell the customer. Never say "you have 24 hours to cancel" or mention any deadline. If a customer asks why they got a refund (or didn\'t), explain the specific outcome, not the rule. It only applies if they haven\'t had a cleaning yet: zero completed jobs. After even one cleaning, handle it as a normal cancellation. Why it exists: we charge the MF upfront ($19 first month, $59 after) before the first cleaning, and it renews automatically, so we give customers about a day to change their mind. How the window works: 1st MF, from sign-up to when they first contacted us to request a cancellation or refund, or to when they self-cancelled. Succeeding MFs: from when that MF was charged to when they asked. Use reasonable judgment: it\'s a guideline, not a strict cutoff; signing up one day and contacting us the next can still count. Signed up and changed their mind fast: treat it like they never really committed. Waited days: that\'s a real cancellation. No request, no refund: even inside the window, don\'t offer one if they didn\'t ask.');
    box(s, 0.45, 1.45, 3.0, 2.4, RED);
    s.addImage({ data: await icon('FiLock', C.white), x: 0.7, y: 1.68, w: 0.42, h: 0.42 });
    T(s, [{ text: 'Internal only', options: { bold: true, fontFace: HEAD, fontSize: 15, breakLine: true } }, { text: 'Never say "you have 24 hours to cancel". Explain the outcome, not the rule.', options: { fontSize: 10.5 } }], { x: 0.7, y: 2.2, w: 2.55, h: 1.5, color: C.white, valign: 'top' });
    const r = [['FiAlertCircle', 'Zero completed jobs only', 'After one cleaning, it\'s a normal cancellation.'], ['FiClock', '1st MF: from sign-up', 'To their first request, or to their self-cancel.'], ['FiRepeat', 'Later MFs: from that charge', 'To when they asked.'], ['FiSliders', 'A guideline, not a cutoff', 'Signed up one day, asked the next? Can still count.']];
    for (const [i, [ico, t, b]] of r.entries()) {
      const y = 1.45 + i * 0.61;
      box(s, 3.6, y, 5.95, 0.54, C.white, C.border);
      await iconDot(s, 3.7, y + 0.09, ico, 0.36);
      T(s, [{ text: t + '  ', options: { bold: true, fontFace: HEAD, fontSize: 10.5 } }, { text: b, options: { fontSize: 9, color: C.soft } }], { x: 4.18, y, w: 5.3, h: 0.54, valign: 'middle' });
    }
    await tip(s, 4.05, 'No request, no refund:', 'even inside the window, don\'t offer one if they didn\'t ask.', 'FiSlash');
  }

  {
    const s = content(FP, 'Refund cheat sheet: no service issue', 'Timing and the cleaning length decide it.',
      'No service issue, no completed job yet. 2–3 hour cleaning, within 24 hours: refund the 1st MF ($19). Use the Refund Voucher option in CRM (no tracker entry), then cancel FC in CRM and invalidate the voucher. 2–3 hour cleaning, after 24 hours: no refund unless there\'s a compelling reason (e.g. compassion, the customer doesn\'t have a card); cancel FC and invalidate the voucher; with a compelling reason, refund via the Refund Voucher option first. 4–6 hour cleaning, within 24 hours: refund the 1st MF and the add-on ($19 for 4 hours, $59 for 6 hours); log to the FCF Refund tracker for a Stripe FULL refund, then cancel FC and invalidate the voucher. 4–6 hour cleaning, after 24 hours: no refund of the 1st MF unless compelling; refund the add-on; log to the FCF Refund tracker for a Stripe PARTIAL refund, then cancel FC and invalidate the voucher.');
    s.addTable([
      ['Cleaning', 'Timing', '1st MF ($19)', 'Add-on ($19 4hr / $59 6hr)', 'How'].map(t => ({ text: t, options: { bold: true, color: C.teal, fill: { color: C.tealSoft }, fontFace: HEAD } })),
      ...[
        ['2–3 hr', 'Within 24 hrs', '✓ Refund', 'N/A', 'Refund Voucher option in CRM'],
        ['2–3 hr', 'After 24 hrs', '✗ Unless compelling', 'N/A', 'Compelling? Refund Voucher option'],
        ['4–6 hr', 'Within 24 hrs', '✓ Refund', '✓ Refund', 'FCF Refund tracker: Stripe FULL'],
        ['4–6 hr', 'After 24 hrs', '✗ Unless compelling', '✓ Refund', 'FCF Refund tracker: Stripe PARTIAL'],
      ].map(r => r.map((t, i) => ({ text: t, options: { color: t.startsWith('✓') ? C.teal : t.startsWith('✗') ? RED : C.ink, bold: i === 0 || t.startsWith('✓') || t.startsWith('✗'), fill: { color: C.white } } })))
    ], { x: 0.45, y: 1.42, w: 9.1, colW: [1.0, 1.4, 1.75, 2.15, 2.8], rowH: 0.42, fontFace: BODY, fontSize: 9.5, valign: 'middle', border: { type: 'solid', pt: 0.75, color: C.border }, margin: [0.03, 0.1, 0.03, 0.1] });
    await tip(s, 3.7, 'Every row ends with:', 'Cancel FC in CRM → Invalidate the voucher.', 'FiCheckSquare');
    await tip(s, 4.22, 'Compelling reason:', 'e.g. compassion, or the customer doesn\'t have a card.', 'FiHeart');
  }

  {
    const s = content(FP, 'Refund cheat sheet: service issues', 'Whose fault was it?',
      'Service failure, company\'s or CP\'s fault: timing doesn\'t matter; refund the 1st MF and the add-on. Service failure, customer at fault, within 24 hours: refund both. Customer at fault (except biohazard), after 24 hours: no refund of the 1st MF unless there\'s a compelling reason; refund the add-on. Biohazard reported by the CP: timing doesn\'t matter; refund both. How to process: 2–3 hour cleaning, use the Refund Voucher option in CRM (no tracker entry); 4–6 hour cleaning, log to the FCF Refund tracker for a Stripe FULL refund. Then cancel FC in CRM and invalidate the voucher. Agent Do: for 4–6 hour cleanings, don\'t use the Refund Voucher option: it only refunds the $19 MF, not the add-on. Log to the FCF Refund tab so a TL issues the full Stripe refund.');
    s.addTable([
      ['Situation', 'Timing', '1st MF', 'Add-on'].map(t => ({ text: t, options: { bold: true, color: C.teal, fill: { color: C.tealSoft }, fontFace: HEAD } })),
      ...[
        ['Company or CP fault', 'Doesn\'t matter', '✓ Refund', '✓ Refund'],
        ['Customer at fault', 'Within 24 hrs', '✓ Refund', '✓ Refund'],
        ['Customer at fault (not biohazard)', 'After 24 hrs', '✗ Unless compelling', '✓ Refund'],
        ['Biohazard reported by the CP', 'Doesn\'t matter', '✓ Refund', '✓ Refund'],
      ].map(r => r.map((t, i) => ({ text: t, options: { color: t.startsWith('✓') ? C.teal : t.startsWith('✗') ? RED : C.ink, bold: i === 0 || t.startsWith('✓') || t.startsWith('✗'), fill: { color: C.white } } })))
    ], { x: 0.45, y: 1.42, w: 9.1, colW: [3.4, 1.9, 1.9, 1.9], rowH: 0.4, fontFace: BODY, fontSize: 9.5, valign: 'middle', border: { type: 'solid', pt: 0.75, color: C.border }, margin: [0.03, 0.1, 0.03, 0.1] });
    headPanel(s, 0.45, 3.6, 4.45, 1.0, '2–3 hr cleaning', ['Refund Voucher option in CRM. No tracker entry'], C.teal, C.tealSoft, 9.5);
    headPanel(s, 5.1, 3.6, 4.45, 1.0, '4–6 hr cleaning', ['FCF Refund tracker: TL issues a Stripe FULL refund'], C.gold, C.goldSoft, 9.5);
    await warn(s, 4.72, '4–6 hr? Not the Refund Voucher option:', 'it only refunds the $19 MF, not the add-on.', 0.42);
  }

  {
    const s = content(FP, '2+ MFs paid, no jobs done', 'No refund by default. Any one condition qualifies.',
      '2+ MFs paid with no jobs done: no refund by default (timing doesn\'t matter), unless one of these applies: 1) they reached out within 24 hours of being charged their most recent MF; 2) they tried to book within a specific MF\'s coverage period but didn\'t get a cleaning because of CP or platform fault: refund that specific MF; 3) they reached out within about 3 days of their most recent MF charge AND the root cause is CP or platform fault (not a strict 72-hour cutoff; before the 4th day can still qualify); 4) biohazard reported by the CP: refund all unused MFs in full if the customer asks (the 120-day timeframe still applies), regardless of timing, since they never had a real shot at using the membership. The add-on is refunded. If any apply, for the 1st charge: 2–3 hour cleaning, Refund Voucher option in CRM; 4–6 hour cleaning, FCF Refund tab for a Stripe FULL refund; then cancel FC in CRM and invalidate the voucher. Macros: FCF Refund (No Completed Cleaning): Cleaning Charge Only; Courtesy (Membership Fee); Courtesy (Service Issue); Partial MF Refund (1 of 2+ Charges).');
    const c = [
      ['FiClock', 'Within 24 hrs of the latest MF', 'They reached out fast after the charge.'],
      ['FiCalendar', 'Tried to book, CP/platform fault', 'Refund that specific MF.'],
      ['FiAlertTriangle', '~3 days + CP/platform fault', 'Before the 4th day can still count.'],
      ['FiAlertOctagon', 'Biohazard reported by the CP', 'All unused MFs, if they ask (120 days).'],
    ];
    const w = (9.1 - 0.15) / 2, h = 0.95;
    for (const [i, [ico, t, b]] of c.entries()) {
      const x = 0.45 + (i % 2) * (w + 0.15), y = 1.45 + Math.floor(i / 2) * (h + 0.12);
      box(s, x, y, w, h, C.white, C.border);
      await iconDot(s, x + 0.18, y + 0.25, ico, 0.42);
      T(s, [{ text: t, options: { bold: true, fontFace: HEAD, fontSize: 11, breakLine: true } }, { text: b, options: { fontSize: 9.5, color: C.soft } }], { x: x + 0.75, y, w: w - 0.9, h, valign: 'middle' });
    }
    await tip(s, 3.65, 'Then, for the 1st charge:', '2–3 hr, Refund Voucher option. 4–6 hr, FCF Refund tab for a Stripe FULL refund.', 'FiCreditCard');
    await tip(s, 4.17, 'Finish with:', 'Cancel FC in CRM → Invalidate the voucher.', 'FiCheckSquare');
  }

  {
    const s = content(FP, 'Changing the first cleaning\'s length', 'Keep FC in place. Settle the difference at sign-up pricing.',
      'Related scenario, FCF first-cleaning duration change: FCF customers occasionally want to change the duration they picked at sign-up. Approved handling: keep FC in place, manually update the voucher duration in Django, and settle the price difference at sign-up-tier pricing: +$19 for 4 hours, +$59 for 6 hours; refund the difference on a downgrade. Agent Do: confirm the current and target tier, get the customer\'s agreement on the price difference, then update the voucher duration in Django and charge or refund accordingly.');
    const w = 2.1, gap = 0.233;
    const flow = [['Confirm current & target tier', 'end'], ['Customer agrees to the difference', 'decision'], ['Update voucher duration in Django', 'action'], ['Charge or refund the difference', 'action']];
    for (let i = 0; i < flow.length; i++) {
      const x = 0.45 + i * (w + gap);
      await flowBox(s, x, 1.5, w, 0.7, flow[i][0], flow[i][1]);
      if (i < flow.length - 1) flowLine(s, [[x + w, 1.85], [x + w + gap, 1.85]], true);
    }
    const sw = (9.1 - 0.3) / 3;
    stat(s, 0.45, 2.5, sw, 1.3, '+$19', 'to move up to 4 hours');
    stat(s, 0.45 + sw + 0.15, 2.5, sw, 1.3, '+$59', 'to move up to 6 hours');
    stat(s, 0.45 + 2 * (sw + 0.15), 2.5, sw, 1.3, 'Refund', 'the difference on a downgrade', C.gold);
    await tip(s, 4.0, 'Keep FC in place:', 'this isn\'t a cancellation.', 'FiCheckCircle');
  }

  await knowledgeCheck(FP, [
    ['A customer asks how long they have to cancel for a refund. Tell them 24 hours?', 'No. The window is internal only. Explain the outcome, not the rule.'],
    ['FCF, 2–3 hr cleaning, cancels 3 hours after sign-up and asks for a refund. What do you do?', 'Refund via the Refund Voucher option, cancel FC, invalidate the voucher.'],
    ['FCF, 6-hr cleaning, within 24 hrs. Use the Refund Voucher option?', 'No. It only refunds $19. Log to the FCF Refund tracker for a Stripe FULL refund.'],
    ['The customer had one completed cleaning. Does the 24-hour window apply?', 'No. Handle it as a normal cancellation.'],
  ], 'Ask each question and let trainees answer before revealing. Each click reveals the next answer.', KC_TITLE, '');

  await practice(FP, [
    ['Which row of the cheat sheet applies? Why?', 'Completed jobs, timing, cleaning length, fault.'],
    ['What actions are you going to take?', 'Refund how, then cancel FC and invalidate the voucher.'],
    ['5-min writing challenge', 'Create your comms. No 24-hour talk!'],
  ], 'TRAINER: share a live FCF ticket where the customer wants to cancel before their first completed job. Give trainees a few minutes to review it on their own, then work through it together. 1) Which row applies? Check there are zero completed jobs, retention was attempted, the timing (internal 24-hour window), the cleaning length (2–3 vs 4–6 hours), and whose fault any failure was. 2) What actions are you going to take? 2–3 hours: Refund Voucher option in CRM; 4–6 hours: FCF Refund tracker for a Stripe full or partial refund; then cancel FC and invalidate the voucher. 3) 5-minute writing challenge: each trainee writes the comms, without mentioning the 24-hour window. Read a few aloud and compare. Then open the floor for questions.');

  // ================= 7. TRIAL / ONE-TIME CLEANING =================
  const TC = 'Trial & One-Time Cleaning';
  await topic('Trial Cleaning & One-Time Cleaning', 'Alternatives to a full ForeverClean membership.',
    ['Trial Cleaning vs One-Time Cleaning', 'Pricing: what gets deducted', 'Converting an unused voucher or FCF', 'Cancellations and refunds afterwards'],
    'ill27.png',
    'Trial Cleaning (TC) / One-Time Cleaning (OTC). This covers two non-membership (or limited-membership) alternatives to a full ForeverClean membership. Both come up in two ways: 1) Conversion: a customer with an unused DHJ voucher, or an FCF membership with no completed job, can be converted to TC/OTC instead of using or refunding the voucher or continuing the membership. 2) Direct purchase: a customer can buy TC/OTC directly through Sales without ever signing up for ForeverClean (full regular rate from the FC Sales Calculator). This applies to both DHJ and FCF. Why it matters: not every customer wants an ongoing membership. Steering them into a refund by default, when a TC/OTC would get them what they want, is a missed opportunity.');

  {
    const s = content(TC, 'Trial Cleaning vs One-Time Cleaning', 'They look alike, but the consequences are completely different.',
      'Both are one-off paid cleanings charged upfront, but they have completely different downstream consequences. Trial Cleaning (TC): triggers an FC table (membership). The customer has a 30-day window from the appointment to cancel with no obligation. If no cancellation request comes in during that window, the membership activates for real: $59/month, 6-month MCT, $99 ETF going forward. Use case: wants to try the membership before committing. One-Time Cleaning (OTC): no FC table; nothing further, no ongoing relationship. Use case: wants a single cleaning (e.g. a move-out clean) with no interest in a membership.');
    headPanel(s, 0.45, 1.45, 4.45, 2.6, 'Trial Cleaning (TC)', ['Wants to try the membership first', 'Creates an FC table', 'Charged upfront', '30 days from the appointment to cancel, no obligation', 'Then $59/mo, 6-month MCT, $99 ETF'], C.teal, C.tealSoft, 10);
    headPanel(s, 5.1, 1.45, 4.45, 2.6, 'One-Time Cleaning (OTC)', ['Wants a single cleaning (e.g. move-out)', 'No FC table', 'Charged upfront', 'Nothing after: no membership to track'], C.gold, C.goldSoft, 10);
    await tip(s, 4.25, 'Don\'t default to a refund:', 'a TC or OTC may be exactly what they want.', 'FiGift');
  }

  {
    const s = content(TC, 'Pricing: what gets deducted', 'Get the rate from the FC Sales Calculator first.',
      'Use the (Transformation) FC SALES CALCULATOR to get the correct TC/OTC rate before applying either deduction. DHJ: deduct the voucher purchase price they already paid. FCF: deduct all membership fees paid to date (within 120 days), not just the initial $19 first MF, plus the add-on if they chose a 4-hour ($19) or 6-hour ($59) cleaning. If the deduction exceeds the TC/OTC price (FCF): no refund of the excess; the customer pays $0, but nothing comes back to them. Examples: 2 MFs + $59 add-on = $137 paid; a 6-hour clean at $378 means they pay the remaining $241. $150 in MFs paid vs a $120 OTC: $0 due, and the $30 is not refunded.');
    headPanel(s, 0.45, 1.45, 4.45, 1.3, 'DHJ: deduct', ['The voucher price they already paid'], C.ink, C.white, 10);
    headPanel(s, 5.1, 1.45, 4.45, 1.3, 'FCF: deduct', ['All MFs paid to date (within 120 days)', 'Plus the add-on: $19 (4 hr) or $59 (6 hr)'], C.teal, C.tealSoft, 10);
    const ex = [['6-hr clean at $378', '2 MFs + $59 add-on = $137 paid', 'Pays $241'], ['OTC at $120', '$150 in MFs already paid', 'Pays $0. $30 is not refunded']];
    ex.forEach(([a, b, c], i) => {
      const x = 0.45 + i * 4.65;
      box(s, x, 2.95, 4.45, 1.0, C.white, C.border);
      T(s, [{ text: a, options: { bold: true, fontFace: HEAD, fontSize: 11, breakLine: true } }, { text: b, options: { fontSize: 9.5, color: C.soft, breakLine: true } }, { text: '→ ' + c, options: { bold: true, fontSize: 11, color: C.teal } }], { x: x + 0.2, y: 2.95, w: 4.1, h: 1.0, valign: 'middle' });
    });
    await tip(s, 4.15, 'First:', 'get the TC/OTC rate from the (Transformation) FC Sales Calculator.', 'FiDollarSign');
  }

  {
    const s = content(TC, 'Offering it', 'Explain what they signed up for, then offer the better fit.',
      'Acknowledge and empathize. DHJ only: if they have refund intent, provide the self-refund link, advise entering the voucher code before "Refund Voucher", and include the code in your email (not applicable to FCF: they signed up for a membership, not a voucher). Explain what they signed up for and offer the alternative. DHJ OTC: "The voucher you purchased comes with a membership, so if you\'re only looking for a one-time cleaning, using it wouldn\'t be the best route. What I can do instead is offer a one-time cleaning at $[X], that\'s the regular rate minus the $[X] you already paid for the voucher, charged upfront. Just let me know your preferred date and time and I\'ll get that booked for you." FCF OTC: "What you signed up for is a membership, so if you\'re only looking for a one-time cleaning…" DHJ/FCF TC: "I totally get that you might not want to commit to a membership right away. We actually have a Trial Cleaning option, which could be a better fit. It lets you try things out first… a standard [X]-hour clean is normally $[X], but we can apply what you\'ve already paid, bringing your total to $[X], charged upfront." For TC, set expectations: 30 days from the appointment to contact us if they don\'t want to continue, cancelled with no obligation; otherwise the $59/month membership starts automatically, with a 6-month MCT, and a $99 ETF if cancelled after the 30 days and before completing the 6 months. Check upcoming appointments and RC: cancel via CRM if there\'s no sign they want to keep them. Macros: "DHJ - Wants to \'Try it Out\' (Trial Cleaning)", "DHJ - Wants One Time Cleaning".');
    quote(s, 0.45, 1.45, 4.45, 1.45, 'The voucher you purchased comes with a membership, so if you\'re only looking for a one-time cleaning, using it wouldn\'t be the best route. What I can do instead is offer a one-time cleaning at $[X]…', 'ONE-TIME CLEANING');
    quote(s, 5.1, 1.45, 4.45, 1.45, 'We actually have a Trial Cleaning option, which could be a better fit. It lets you try things out first… we can apply what you\'ve already paid, bringing your total to $[X], charged upfront.', 'TRIAL CLEANING');
    headPanel(s, 5.1, 3.05, 4.45, 1.5, 'Trial: set expectations', ['30 days from the appointment to cancel, no obligation', 'Then $59/mo starts automatically', '6-month MCT. $99 ETF if cancelled early'], C.teal, C.white, 9);
    headPanel(s, 0.45, 3.05, 4.45, 1.5, 'Also', ['DHJ with refund intent: send the self-refund link', 'Paid upfront: get their preferred date and time', 'Cancel unwanted appointments and RC'], C.ink, C.white, 9);
  }

  {
    const s = content(TC, 'They accept: set it up', 'The FC table and voucher steps differ by model.',
      'If the customer accepts: confirm the account info (complete address, a valid card on file, appointment date, time and duration). If they still want a refund: DHJ, refund the voucher via CRM and invalidate it; FCF, log the 1st charge (1st MF + add-on, if any) to the FCF Refund log so a TL refunds it via Stripe. If they no longer want a refund, apply the deduction and invalidate the voucher (DHJ voucher price; FCF all MFs paid + add-on). Charge upfront using the FC Sales Calculator. Book the appointment (CRM if applicable, otherwise the customer dashboard). Handle the existing FC table: OTC (DHJ or FCF), invalidate the voucher and deactivate the FC table. DHJ TC: leave the FC table as is, since DHJ MFs only start after a completed job. FCF TC: deactivate the FC table, since FCF starts at sign-up and auto-renews. Prevent a double charge: issue an admin_courtesy voucher (don\'t notify the customer). TC only: click the voucher in Legacy CRM and select "ForeverClean force" (triggers FC once the job is completed) and "1 free month" (delays the next MF by 30 days); add the customer to the FC table (FCF); update the ETF to $99. Confirm to the customer: charged upfront, appointment booked, notified once the CP confirms.');
    const st = ['Confirm address, card, date, time, duration', 'Still wants a refund? DHJ: CRM refund. FCF: FCF Refund log', 'No refund? Deduct what they paid, invalidate the voucher', 'Charge upfront (FC Sales Calculator), book it', 'admin_courtesy voucher, no notification', 'Confirm: charged, booked, CP will confirm'];
    st.forEach((t, i) => {
      const y = 1.42 + i * 0.5;
      box(s, 0.45, y, 4.7, 0.44, C.white, C.border);
      badge(s, 0.53, y + 0.07, i + 1);
      T(s, t, { x: 0.93, y, w: 4.15, h: 0.44, fontSize: 9, valign: 'middle' });
    });
    tableAt(s, 5.3, 4.25, ['Existing FC table', 'Do'], [
      ['OTC (DHJ or FCF)', 'Deactivate it'],
      ['DHJ Trial', 'Leave it as is'],
      ['FCF Trial', 'Deactivate, then force'],
    ], { y: 1.42, colW: [2.0, 2.25], fontSize: 9, rowH: [0.32, 0.36, 0.36, 0.36] });
    headPanel(s, 5.3, 2.95, 4.25, 1.55, 'Trial only', ['Legacy CRM voucher: "ForeverClean force" + "1 free month"', 'FCF: add them to the FC table', 'Update the ETF to $99'], C.teal, C.tealSoft, 9);
  }

  {
    const s = content(TC, 'They decline the offer', 'Refund what they\'re owed and close things out.',
      'If the customer declines the TC/OTC offer. Refund: DHJ, if they request a refund within 1 year of purchase, refund the DHJ voucher through CRM and invalidate it. FCF: check the Refund Eligibility Cheat Sheet (No Completed Job Yet) in the FCF Pre-Invoice Cancellation article. FC table: deactivate it in every case. Advise the customer: DHJ, refund processed, 5–10 business days; FCF, depends on the cheat sheet. CRM account: if they want their account closed, deactivate it, but send comms before deactivating.');
    await cards(s, 1.45, 1.75, [
      { ico: 'FiRotateCcw', title: 'Refund', body: 'DHJ: within 1 year, refund via CRM and invalidate. FCF: use the pre-invoice cheat sheet.' },
      { ico: 'FiXSquare', title: 'Deactivate the FC table', body: 'In every case: DHJ or FCF, TC or OTC.' },
      { ico: 'FiMail', title: 'Closing the account?', body: 'Send comms first, then deactivate the CRM account.' },
    ]);
    await tip(s, 3.45, 'DHJ refund:', 'tell them it\'s processed and takes 5–10 business days.', 'FiClock');
  }

  {
    const s = content(TC, 'Cancelling a TC or OTC', 'Changed their mind, or we let them down?',
      'Applies however the customer got here (Care or Sales). Changed their mind, or no clear reason: give a full admin refund of what they paid; invalidate the voucher tied to the order; if it was a Trial Cleaning, cancel the FC table too (easy to forget: OTC never had a table, but a TC did, and it doesn\'t go away just because you refunded); double-check the appointment itself is cancelled; if they say why, acknowledge warmly and say they\'re welcome back; if they don\'t, probe, since a vague reason might be fixable. If they come back to try again, don\'t charge upfront again; just help them book, and it\'s invoiced normally. Cancelled because of a service issue (no-show, cancellation, no CP claim): same refund steps, plus real acknowledgment: "I understand it\'s frustrating when your cleaning gets cancelled. I\'m here to help." Fix the root cause: block the CP/customer pairing and coach or penalize the CP. If they\'re open to it, offer a different, reliable cleaner, with priority booking if you can: "While your first cleaner cancelled, we can connect you with another trusted cleaner if you\'d like to try again." A small incentive ($10–$25 in credits, or a free hour) goes a long way: we\'re asking for a second chance. Completed the Trial and wants to cancel within the 30-day window: cancel FC, no ETF, whatever the reason. Even if the reason was a bad cleaning, cancel as asked; you can offer a second chance (a different cleaner plus around $20 in credits) if they seem open. If they come back later, follow the usual FC Reactivations.');
    headPanel(s, 0.45, 1.45, 2.95, 2.75, 'Changed their mind', ['Full admin refund', 'Invalidate the voucher', 'Trial? Cancel the FC table too', 'Check the appointment is cancelled', 'No reason given? Probe'], C.teal, C.tealSoft, 9);
    headPanel(s, 3.525, 1.45, 2.95, 2.75, 'We let them down', ['Same refund steps', 'Acknowledge: it wasn\'t on them', 'Block the pairing, coach the CP', 'Offer a reliable cleaner', '$10–$25 credits or a free hour'], RED, PINK, 9);
    headPanel(s, 6.6, 1.45, 2.95, 2.75, 'Trial done, within 30 days', ['Cancel FC. No ETF', 'Whatever the reason', 'Bad cleaning? Still cancel', 'Open to it? New cleaner + ~$20 credits'], C.gold, C.goldSoft, 9);
    await tip(s, 4.4, 'Coming back to try again?', 'don\'t charge upfront again. Just help them book.', 'FiRepeat');
  }

  {
    const s = content(TC, 'Only part of the cleaning was done', 'They don\'t know the voucher exists. They expect real money back.',
      'The customer paid upfront for a TC/OTC but only received part of the cleaning. Key point: the customer has no idea a voucher exists on their account; it\'s an internal way to avoid double-charging when the CP invoices. They paid upfront and expect real money back. If the CP already refunded part of it (as a voucher) and the customer wants cash: the CP already confirmed something went wrong, so invalidate that voucher, admin refund the matching amount from the original charge, and let the customer know. Example: a $252 OTC for 4 hours with 1 hour refunded: refund $63 ($252 / 4). If nothing has been confirmed and it looks like an overcharge: run the standard Overcharged Hours process first, and only refund once it\'s confirmed (real evidence, or the CP never responded). Match the refund to what wasn\'t done: paid 4 hours, got 2, half back. Steps: 1) Block automated comms (C CRM > Do > Deactivate block comms) so the customer doesn\'t get a confusing "your voucher was refunded" message, then refund from the CP dashboard so the CP is only paid for the hours worked. 2) Once you\'ve confirmed the message didn\'t go out, remove the block and invalidate the voucher. 3) Do the math and issue an admin refund for the TC/OTC manual charge.');
    headPanel(s, 0.45, 1.45, 4.45, 1.55, 'CP already refunded (as a voucher)', ['Invalidate that voucher', 'Admin refund the matching amount', 'E.g. $252 for 4 hrs, 1 hr back = $63'], C.teal, C.tealSoft, 9.5);
    headPanel(s, 5.1, 1.45, 4.45, 1.55, 'Nothing confirmed yet', ['Run the Overcharged Hours process first', 'Refund only once it\'s confirmed', 'Paid 4 hrs, got 2? Half back'], C.ink, C.white, 9.5);
    eyebrow(s, 'Issuing the refund', 0.45, 3.15, 4, C.soft);
    const w = 2.9, gap = 0.2;
    const flow = [['Block comms, then\nrefund on the CP dashboard', 'decision'], ['No message sent? Unblock, invalidate voucher', 'end'], ['Admin refund the TC/OTC charge', 'action']];
    for (let i = 0; i < flow.length; i++) {
      const x = 0.45 + i * (w + gap);
      await flowBox(s, x, 3.4, w, 0.62, flow[i][0], flow[i][1]);
      if (i < flow.length - 1) flowLine(s, [[x + w, 3.71], [x + w + gap, 3.71]], true);
    }
    await tip(s, 4.2, 'Block comms first:', 'C CRM > Do > Deactivate block comms. They never knew the voucher existed.', 'FiVolumeX');
  }

  await knowledgeCheck(TC, [
    ['A customer wants a move-out clean only, no membership. TC or OTC?', 'OTC. No FC table, nothing after.'],
    ['FCF: 2 MFs + $59 add-on paid. 6-hr OTC is $378. What do they pay?', '$241 ($378 − $137).'],
    ['FCF: $150 paid, OTC is $120. Refund the $30?', 'No. They pay $0; the excess isn\'t refunded.'],
    ['Completed a Trial, cancels on day 20 over a bad cleaning. ETF?', 'No ETF. Cancel FC as asked; offer a second chance if they\'re open.'],
  ], 'Ask each question and let trainees answer before revealing. Each click reveals the next answer.', KC_TITLE, '');

  await practice(TC, [
    ['Is TC or OTC the better fit? Why?', 'And DHJ or FCF: what gets deducted?'],
    ['What actions are you going to take?', 'Price, charge, book, FC table, voucher.'],
    ['5-min writing challenge', 'Create your offer for the customer.'],
  ], 'TRAINER: share a live ticket where a TC or OTC could be offered (or a TC/OTC cancellation). Give trainees a few minutes to review it on their own, then work through it together. 1) TC or OTC, and why? DHJ or FCF? 2) What actions are you going to take? Get the rate from the FC Sales Calculator, deduct what was paid (voucher, or MFs + add-on within 120 days; no refund of any excess), charge upfront, book, handle the FC table by model, issue the admin_courtesy voucher, and for a TC force FC + 1 free month and set the ETF to $99. 3) 5-minute writing challenge: each trainee writes the offer, including TC expectations if relevant. Read a few aloud and compare. Then open the floor for questions.');

  // ---------- Close ----------
  {
    const s = pres.addSlide({ masterName: 'CONTENT' });
    eyebrow(s, 'Wrap-up', 0.62, 1.9);
    T(s, 'Thank you!', { x: 0.62, y: 2.15, w: 5, h: 0.75, fontFace: HEAD, bold: true, fontSize: 34 });
    T(s, 'When in doubt, check the Knowledge Library or ask your Trainer or Team Lead.', { x: 0.62, y: 2.95, w: 4.6, h: 0.5, fontSize: 12, color: C.soft });
    await panel(s, 'ill9.png');
    s.addNotes('Close the module. Recap: Washington State (confirm residency and MCT 0 / ETF $0, cancel immediately, no retention, never mention the ETF; MF refunds default no). Offer Logic (no cancellation intent needed; Retention, Exit or Make-Right; pick the offer built for the problem; check existing offers; caps and guardrails). Cancellation Intent in C/CP Comms (reach out first; know something, not everything; probe; explain terms; route to the right playbook). General Retention Playbook (fix the real problem; weigh the five factors, don\'t count; hard overrides; match and present the offer; the Newspaper Test; don\'t close too soon). Unused DHJ Voucher (FC hasn\'t started; sort into four cases; self-refund link only with refund intent; close accounts in order). FCF Pre-Invoice (retention first; internal 24-hour window; 2–3 hr Refund Voucher vs 4–6 hr Stripe via tracker; cancel FC and invalidate). Trial and One-Time Cleaning (TC creates an FC table with a 30-day window, OTC doesn\'t; deduct what was paid; never refund the excess).');
  }

  if (topicSlides.length * 3 !== agendaLinks.length) throw new Error(`agenda has ${agendaLinks.length / 3} rows but the deck has ${topicSlides.length} topics`);
  agendaLinks.forEach((l, i) => { agendaSlide._rels.find(r => r.rId === l._rId).Target = String(topicSlides[Math.floor(i / 3)]); });
  await pres.writeFile({ fileName: OUT });
  console.log('wrote', OUT);
})();
