// Builds the "Customer Foundations & C CRM" training deck.
// Usage: node deck.js out.pptx [imageDir]
const path = require('path');
const pptxgen = require('pptxgenjs');
const React = require('react');
const RDS = require('react-dom/server');
const sharp = require('sharp');
const Fi = require('react-icons/fi');

const OUT = process.argv[2] || 'deck.pptx';
const IMG = process.argv[3] || path.join(__dirname, '../img6');

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
pres.title = 'Incorrect Info, LMC & Lockouts';

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
const TOPICS = 3;
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
async function knowledgeCheck(eb, qa, notes, title = 'Knowledge check') {
  const s = content(eb + '  ·  Knowledge check', title, 'Ask each question, then click to reveal the answer.', notes);
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

(async () => {
  // ---------- Cover ----------
  {
    const s = pres.addSlide({ masterName: 'COVER' });
    eyebrow(s, 'Homeaglow  ·  New Hire Care Training', 0.62, 0.6);
    T(s, 'Incorrect Info,\nLMC & Lockouts', { x: 0.62, y: 1.0, w: 4.9, h: 1.6, fontFace: HEAD, bold: true, fontSize: 32 });
    T(s, 'When our own words set the wrong expectation, and how to handle last-minute cancellation and lockout charges.', { x: 0.62, y: 2.7, w: 4.6, h: 0.7, fontSize: 12.5, color: C.soft });
    T(s, 'Internal training material', { x: 0.62, y: 5.0, w: 3, h: 0.2, fontSize: 7, color: C.soft });
    await panel(s, 'ill6.png', 0.4, 4.82);
    s.addNotes('Welcome. This module covers three topics from the Knowledge Library: General Guidance: Incorrect Information Provided; AG-Refund: Last-Minute Cancellation (LMC); and AG-Refund: Lockout (LO). The common thread: validate first, then decide how generous to be. Each topic ends with a knowledge check and time for questions.');
  }

  // ---------- Agenda ----------
  // Each row links to its topic's title slide; the targets are filled in once every slide exists.
  const agendaLinks = [], topicSlides = [];
  let agendaSlide;
  const addTopic = topic;
  topic = async (...args) => { const t = await addTopic(...args); topicSlides.push(t._slideNum); return t; };
  {
    const s = agendaSlide = content('Overview', 'Agenda', 'Three topics. Click a topic to jump to it.',
      'Topic 1 is general guidance that applies whenever our own communication caused the problem. Topics 2 and 3 are action guides for two charges customers contest often: the Last-Minute Cancellation fee and the Lockout charge. Both follow the same two-step shape: is the charge valid, and if so, which customer tier applies. In slideshow mode, click any topic to jump straight to it.');
    const rows = [
      ['General Guidance: Incorrect Information Provided', 'We promised something we can\'t deliver. Fix it, own it, and don\'t charge an ETF for it.'],
      ['AG-Refund: Last-Minute Cancellation (LMC)', 'The $40 fee for cancelling within 6 hours: is it valid, and how much do we refund?'],
      ['AG-Refund: Lockout (LO)', 'The CP couldn\'t get in. Validate the claim, then refund by customer tier.'],
    ];
    const h = 1.0, gy = 0.14;
    for (let i = 0; i < rows.length; i++) {
      const [title, line] = rows[i], y = 1.5 + i * (h + gy);
      const link = { slide: 1, tooltip: 'Go to ' + title };
      agendaLinks.push(link);
      // One shape per row, so a click anywhere on it follows the link.
      s.addText([
        { text: title, options: { color: C.ink, bold: true, fontFace: HEAD, fontSize: 14, breakLine: true, hyperlink: link } },
        { text: line, options: { color: C.soft, fontSize: 10.5, hyperlink: link } },
      ], { shape: pres.shapes.ROUNDED_RECTANGLE, x: 0.45, y, w: 9.1, h, rectRadius: 0.06, fill: { color: C.white }, line: { color: C.border, width: 0.75 },
        fontFace: BODY, valign: 'middle', margin: [62, 40, 0, 0], hyperlink: link });
      const num = { slide: 1, tooltip: 'Go to ' + title };
      agendaLinks.push(num);
      s.addText(String(i + 1).padStart(2, '0'), { x: 0.6, y, w: 0.6, h, fontFace: HEAD, bold: true, fontSize: 18, color: C.teal, align: 'center', valign: 'middle', hyperlink: num });
      const arrow = { slide: 1, tooltip: 'Go to ' + title };
      agendaLinks.push(arrow);
      s.addText('›', { x: 9.0, y, w: 0.4, h, fontFace: HEAD, bold: true, fontSize: 22, color: C.teal, align: 'center', valign: 'middle', hyperlink: arrow });
    }
  }

  // ================= 1. INCORRECT INFORMATION PROVIDED =================
  const INC = 'Incorrect Information';
  await topic('Incorrect Information Provided', 'A customer signed up because of a promise we can\'t keep, and now wants out.',
    ['What counts as a broken promise', 'The exact test: "at" vs "around"', 'The process', 'Worked examples'],
    'ill30.png',
    'Category: Job-Related Charges & Refund Requests. This covers a customer who signed up for a membership and is now trying to cancel because of a guaranteed promise Homeaglow can\'t fulfill. It isn\'t limited to pricing. AG – Refund: Processing Fee is one worked-out example of this general rule.');

  {
    const s = content(INC, 'What is this?', 'Our own words set an expectation we can\'t meet.',
      'Three shapes of the same problem. (1) A rate or term stated as fixed or guaranteed by a Sales or CS agent, or by a system-generated message. (2) A service coverage that was promised but isn\'t part of what the platform offers. (3) Guarantee-pairing: a rep mentions an unrelated guarantee, like the Happiness Guarantee, back-to-back with cancellation or ETF language, without separating "we\'ll fix the job" from "you can cancel penalty-free". The Happiness Guarantee only covers cleaning quality; Sales confirmed it\'s never meant to be tied to charges, refunds or compensation. Sales treats pairing it with cancellation language as a protocol violation on their end; we treat it as misrepresentation on ours. Why it matters: when our communication is the reason expectations don\'t match reality, it\'s our error, not a customer misunderstanding. The financial impact gets corrected whether the customer stays or goes, and if the broken promise is why they want to leave, they shouldn\'t pay an ETF for it.');
    await cards(s, 1.45, 2.0, [
      { ico: 'FiDollarSign', title: 'Rate or price claim', body: 'A rate or term stated as fixed or guaranteed, by an agent or a system message.' },
      { ico: 'FiTool', title: 'Service coverage', body: 'A tool or service confirmed as available that the platform doesn\'t offer.' },
      { ico: 'FiLink', title: 'Guarantee-pairing', body: 'The Happiness Guarantee said right next to cancellation or ETF terms.' },
    ]);
    await tip(s, 3.65, 'Our error:', 'not a customer misunderstanding. If it\'s why they leave, no ETF.', 'FiAlertCircle');
  }

  {
    const s = content(INC, 'The exact test', 'Did we make an affirmative, guaranteed claim that turned out to be false?',
      'This isn\'t about whether the customer asked good questions. It\'s about whether we made an affirmative, guaranteed claim. For rates, the word "at" is the tell: "at $19/hour", "at $23/hour" is an exact figure, a guarantee we can\'t make because CP rates vary and cleaners are independent contractors. "Around", "approximately", "starting at" is an estimate: a different actual rate isn\'t misrepresentation, it\'s a normal expectation gap. Verify against the sales call recording or notes before applying the test. Don\'t rely on the customer\'s recollection alone.');
    await twoCol(s,
      { ico: 'FiAlertTriangle', title: 'Guarantee: misrepresentation', items: ['"At $19/hour"', '"At $23/hour"', 'An exact figure, not a range'] },
      { ico: 'FiCheckCircle', title: 'Estimate: not misrepresentation', items: ['"Around $20/hour"', '"Approximately", "roughly"', '"Starting at"'] },
      1.45, 2.05);
    await tip(s, 3.7, 'Verify first:', 'check the sales call recording or notes. Don\'t rely on the customer\'s memory alone.', 'FiHeadphones');
  }

  {
    const s = content(INC, 'The process', 'Verify first. Only verified misinformation moves on to steps 2–6.',
      '1) Verify: confirm it\'s genuinely a guaranteed promise, not an estimate or a gap in what was proactively mentioned. Look for the specific language used, not just whether the customer was surprised. To verify, run a Sales call review: submit a Sales call review request (link icon on the pop-up opens the request sheet), add an internal note with your escalation link, and don\'t resolve yet: change the status to Waiting on CSQ. If misinformation is not verified, clarify the information and address any confusion. If it is verified, continue: 2) Refund any overpayment that resulted, where applicable: e.g. promised a $49/month MF but charged $59, refund the difference. Not every case has an overpayment; service-capability and guarantee-pairing cases often don\'t. 3) Acknowledge the error directly. Don\'t frame it as a misunderstanding on the customer\'s side. 4) Correct the charges or expectations going forward, and clearly explain what changed. 5) Assess retainability as for any other case. If retainable, attempt retention with one light offer, not a push: up to $20 credit or a 1-hr courtesy voucher. While you wait for their reply, set a TR before the next MF. 6) If the customer declines the offer, or isn\'t retainable, waive the ETF and cancel the FC.' +
      ' TRAINER: click once (Sales call review) to open the escalation steps; click again to close them. In slideshow mode, the link icon on the pop-up opens the Sales call review request sheet.');
    const stepBox = (x, y, w, n, title, sub, kind = 'action') => {
      const dark = kind === 'penalty';
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.62, fill: { color: dark ? C.ink : C.teal }, line: { color: dark ? C.ink : C.teal, width: 1 }, rectRadius: 0.08 });
      T(s, [{ text: `${n}  ${title}`, options: { bold: true, fontFace: HEAD, fontSize: 10.5, breakLine: true } }, { text: sub, options: { fontSize: 8 } }],
        { x: x + 0.05, y, w: w - 0.1, h: 0.62, color: C.white, align: 'center', valign: 'middle' });
    };
    const lbl = (t, x, y, w = 0.5) => T(s, t, { x, y, w, h: 0.2, fontSize: 8, bold: true, color: C.soft, align: 'center' });
    // Row 1: verify, the decision, then steps 2-3.
    const y1 = 1.42, h = 0.62, m1 = y1 + h / 2;
    stepBox(0.45, y1, 1.9, 1, 'Verify', 'Is it a real promise?');
    flowLine(s, [[2.35, m1], [2.6, m1]], true);
    await flowBox(s, 2.6, y1, 2.05, h, 'Misinformation\nverified?', 'decision');
    flowLine(s, [[4.65, m1], [4.95, m1]], true); lbl('Yes', 4.55, m1 - 0.22);
    stepBox(4.95, y1, 2.0, 2, 'Refund the gap', 'Charged more than promised?');
    flowLine(s, [[6.95, m1], [7.25, m1]], true);
    stepBox(7.25, y1, 2.3, 3, 'Own the error', 'Ours, not "a misunderstanding"');
    // Not verified: normal handling.
    flowLine(s, [[3.625, y1 + h], [3.625, 2.3]], true); lbl('No', 3.65, y1 + h + 0.02, 0.3);
    await flowBox(s, 2.6, 2.3, 2.05, 0.5, 'Clarify the information and address any confusion', 'end');
    // Under Verify: the button that opens the Sales call review steps.
    s.addText('Sales call review  ›', { shape: pres.shapes.ROUNDED_RECTANGLE, x: 0.45, y: 2.2, w: 1.9, h: 0.4, rectRadius: 0.08,
      fill: { color: C.white }, line: { color: C.teal, width: 1.5 }, fontFace: HEAD, bold: true, fontSize: 10, color: C.teal, align: 'center', valign: 'middle' });
    flowLine(s, [[1.4, y1 + h], [1.4, 2.2]]);
    // Row 1 to row 2.
    const y2 = 3.05, m2 = y2 + h / 2;
    flowLine(s, [[8.4, y1 + h], [8.4, 2.88], [1.45, 2.88], [1.45, y2]], true);
    stepBox(0.45, y2, 2.0, 4, 'Correct it', 'Fix charges, explain what changed');
    flowLine(s, [[2.45, m2], [2.8, m2]], true);
    stepBox(2.8, y2, 2.5, 5, 'One light offer', 'Up to $20 credit or a 1-hr voucher');
    flowLine(s, [[5.3, m2], [6.65, m2]], true);
    T(s, 'Declined or\nnot retainable', { x: 5.3, y: m2 - 0.42, w: 1.35, h: 0.38, fontSize: 8, bold: true, color: C.soft, align: 'center', valign: 'bottom' });
    stepBox(6.65, y2, 2.9, 6, 'Waive the ETF, cancel FC', 'No ETF: the error was ours', 'penalty');
    await tip(s, 3.9, 'Waiting on a reply?', 'set a TR before the next MF.', 'FiClock');
    // Pop-up: Sales call review steps open on click 1 and close on click 2.
    const px = 0.45, py = 2.7, pw = 5.6, ph = 1.72, SHEET = 'https://docs.google.com/spreadsheets/d/1uW45nsR5hVwTl9x32EWYaOevp_LeXd61mLAM7qUTv2I/edit?gid=0#gid=0';
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: px, y: py, w: pw, h: ph, fill: { color: C.white }, line: { color: C.teal, width: 1.5 }, rectRadius: 0.06,
      shadow: { type: 'outer', color: '000000', opacity: 0.18, blur: 6, offset: 2, angle: 90 }, objectName: 'pop1_2Panel' });
    T(s, 'Sales call review', { x: px + 0.15, y: py + 0.06, w: 3, h: 0.3, fontFace: HEAD, bold: true, fontSize: 11, color: C.teal, valign: 'middle', objectName: 'pop1_2Title' });
    const rows = ['Sales call review request', 'Add an internal note with your escalation link', 'Don\'t resolve yet. Change the status to Waiting on CSQ'];
    for (let i = 0; i < rows.length; i++) {
      const ry = py + 0.42 + i * 0.33;
      s.addText(String(i + 1), { shape: pres.shapes.OVAL, x: px + 0.18, y: ry + 0.03, w: 0.24, h: 0.24, fill: { color: C.tealSoft }, line: { color: C.tealSoft },
        fontFace: HEAD, bold: true, fontSize: 8.5, color: C.teal, align: 'center', valign: 'middle', margin: 0, objectName: `pop1_2Num${i + 1}` });
      T(s, rows[i], { x: px + 0.52, y: ry, w: i === 0 ? 2.3 : pw - 0.7, h: 0.3, fontSize: 10, color: C.ink, bold: i === 2, valign: 'middle', objectName: `pop1_2Row${i + 1}` });
    }
    // Clickable link icon: opens the Sales call review request sheet.
    const link = { url: SHEET, tooltip: 'Open the Sales call review request sheet' };
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: px + 2.45, y: py + 0.43, w: 0.34, h: 0.28, rectRadius: 0.06, fill: { color: C.gold }, line: { color: C.gold }, hyperlink: link, objectName: 'pop1_2LinkBg' });
    s.addImage({ data: await icon('FiLink', C.white), x: px + 2.54, y: py + 0.47, w: 0.17, h: 0.17, hyperlink: link, objectName: 'pop1_2LinkIcon' });
    T(s, 'Verified? Continue to step 2.', { x: px + 0.18, y: py + ph - 0.32, w: pw - 0.36, h: 0.26, fontSize: 8.5, italic: true, color: C.soft, valign: 'middle', objectName: 'pop1_2Foot' });
  }

  {
    const s = content(INC, 'Quoted a rate, left out the fees', 'Never just "a gap in what was mentioned".',
      'If the rep quoted a specific rate and left out the processing fee, tax or Premium, the customer had no way to know the real price. That\'s misrepresentation. Honor the total they were told: admin refund so the job comes to what they expected. If someone disclosed a charge before it happened and the customer agreed to it (e.g. the CP explained the extra-hours rate), refund only the part that was never disclosed. Worked example: FCR told the customer "$20/hr" and didn\'t mention the processing fee, tax or Premium. Admin refund so the job totals $20/hr × hours worked. Set the max CP rate to $20 and the Cleaner Experience Level to Standard, and explain the fees going forward. Recommended macro: Misinformation: address incorrect claims about charges.');
    const w = 2.1, gap = 0.23;
    const flow = [
      ['Told "$20/hr"', 'penalty', 'No mention of processing fee, tax or Premium.'],
      ['Admin refund', 'action', 'So the job totals $20/hr × hours worked.'],
      ['Set the account', 'action', 'Max CP rate $20. Cleaner Experience Level: Standard.'],
      ['Explain the fees', 'end', 'So the next invoice isn\'t a surprise.'],
    ];
    for (let i = 0; i < flow.length; i++) {
      const x = 0.45 + i * (w + gap);
      await flowBox(s, x, 1.55, w, 0.6, flow[i][0], flow[i][1]);
      if (i < flow.length - 1) flowLine(s, [[x + w, 1.85], [x + w + gap, 1.85]], true);
      T(s, flow[i][2], { x, y: 2.25, w, h: 0.7, fontSize: 9.5, color: C.soft, align: 'center' });
    }
    await tip(s, 3.35, 'Disclosed and agreed?', 'refund only the part that was never disclosed.', 'FiInfo');
    await macroFlag(s, 0.45, 3.95, 'Recommended Macro:', 'Misinformation: address incorrect claims about charges', 6.2);
  }

  {
    const s = content(INC, 'Worked examples', 'Same shape of problem, three ways. Plus what doesn\'t count.',
      'Rate guarantee: signed up through Sales after being told "at $23/hour", not "around". If now charged more, refund the difference; if they want to cancel because of that guarantee, waive the ETF. Service capability: customer asked specifically about a rotary floor buffer for engineered hardwood floors; an agent confirmed we could provide it, but cleaners on the platform don\'t carry buffers. If they cancel because of this, they\'re eligible for the ETF waiver. Happiness Guarantee paired with cancellation language: the rep explained cancellation/ETF terms, then immediately said the Happiness Guarantee verbatim ("you\'re covered... you\'re not satisfied until you are") without separating that it only covers cleaning quality. The customer reasonably heard "if I\'m not happy, I can walk away for free". ETF waiver eligible. Contrast, not misrepresentation: "starting around $20/hour" or "roughly $23" is an estimate; and the Happiness Guarantee mentioned only in response to a cleaning-quality concern, with no cancellation/ETF language nearby, is the guarantee working as intended.');
    table(s, ['Scenario', 'What was said', 'Outcome'], [
      ['Rate guarantee', '"At $23/hour", not "around"', 'Refund the difference. Cancelling over it? Waive ETF.'],
      ['Service capability', 'Agent confirmed a rotary floor buffer', 'Platform doesn\'t offer it. Cancelling? ETF waiver.'],
      ['Guarantee-pairing', 'Happiness Guarantee right after ETF terms', 'Reads as "cancel free". ETF waiver.'],
      ['Not misrepresentation', '"Starting around $20" · HG for a quality issue only', 'Estimate or intended use. Standard handling.'],
    ], { y: 1.45, colW: [2.1, 3.4, 3.6], fontSize: 9.5, rowH: 0.6 });
  }

  await knowledgeCheck(INC, [
    ['Sales told the customer "around $20/hour". They\'re charged $24. Misrepresentation?', 'No. "Around" is an estimate. Handle as a normal expectation gap.'],
    ['FCR quoted "$20/hr" and never mentioned processing fee, tax or Premium. What do you do?', 'Admin refund to $20/hr × hours. Max CP rate $20, Experience Level Standard, explain fees.'],
    ['The customer is retainable. What can you offer, and how many times?', 'One light offer: up to $20 credit or a 1-hr voucher. Declined? Waive ETF and cancel.'],
    ['The rep said the Happiness Guarantee right after explaining the ETF. Customer wants to cancel. ETF?', 'Waive it. Pairing the guarantee with cancellation terms is misrepresentation.'],
  ], 'Ask each question and let trainees answer before revealing. Each click reveals the next answer.', 'Quick check');

  {
    const s = content(INC + '  ·  Practice', 'Let\'s investigate together!', 'Review the live ticket and investigate.',
      'TRAINER: share a live Incorrect Information ticket. Give trainees a few minutes to review it on their own, then work through it together. 1) Is there verified misinformation? What exactly was said, and by whom? "At $23/hour" is a guarantee; "around" or "starting at" is an estimate. Would this need a Sales call review (request, internal note with the escalation link, status Waiting on CSQ)? 2) What was the misinformation, and how do we correct it? Refund the gap, own the error, correct the charges or expectations, one light offer; declined or not retainable: waive the ETF and cancel the FC. Which macro? 3) 5-minute writing challenge: each trainee writes the comms they would send the customer. Read a few aloud and compare. Then open the floor for questions on Incorrect Information Provided.');
    const top = 1.42, bottom = 4.3, gap = 0.15;
    // Left: the live ticket the trainer provides.
    const lw = 3.7;
    box(s, 0.45, top, lw, bottom - top, C.white, C.border);
    await iconDot(s, 0.65, top + 0.2, 'FiInbox', 0.42);
    T(s, 'Live ticket', { x: 0.65, y: top + 0.78, w: lw - 0.4, h: 0.32, fontFace: HEAD, bold: true, fontSize: 13 });
    T(s, 'Your trainer will provide a live ticket. Review it and investigate before you decide anything.', { x: 0.65, y: top + 1.15, w: lw - 0.4, h: 0.8, fontSize: 10.5, color: C.soft });
    s.addText('Trainer provides', { shape: pres.shapes.ROUNDED_RECTANGLE, x: 0.65, y: bottom - 0.58, w: 1.5, h: 0.36, rectRadius: 0.08,
      fill: { color: C.goldSoft }, line: { color: C.goldSoft }, fontFace: BODY, bold: true, fontSize: 9.5, color: C.gold, align: 'center', valign: 'middle', margin: 0 });
    // Right: three prompts, in order.
    const rx = 0.45 + lw + gap, rw = 9.55 - rx, rh = (bottom - top - 2 * gap) / 3;
    const prompts = [
      ['Is there verified misinformation?', 'What exactly was said: a promise or an estimate?'],
      ['What are those, and how do we correct them?', 'Refund the gap, own it, correct it, one light offer.'],
      ['5-min writing challenge', 'Create your comms for the customer.'],
    ];
    for (let n = 0; n < prompts.length; n++) {
      const y = top + n * (rh + gap);
      box(s, rx, y, rw, rh, C.white, C.border);
      badge(s, rx + 0.2, y + (rh - 0.3) / 2, n + 1);
      T(s, prompts[n][0], { x: rx + 0.7, y: y + 0.14, w: rw - 0.9, h: 0.32, fontFace: HEAD, bold: true, fontSize: 12.5, valign: 'middle' });
      T(s, prompts[n][1], { x: rx + 0.7, y: y + 0.48, w: rw - 0.9, h: 0.28, fontSize: 10, color: C.soft, valign: 'middle' });
    }
    // Questions strip.
    box(s, 0.45, 4.45, 9.1, 0.55, C.teal);
    s.addImage({ data: await icon('FiMessageCircle', C.white), x: 0.7, y: 4.585, w: 0.28, h: 0.28 });
    T(s, [{ text: 'Questions?  ', options: { bold: true, fontFace: HEAD, fontSize: 14 } }, { text: 'Open the floor before we move to the next topic.', options: { fontSize: 11 } }],
      { x: 1.15, y: 4.45, w: 8.2, h: 0.55, color: C.white, valign: 'middle' });

  }

  // ================= 2. LAST-MINUTE CANCELLATION =================
  const LMC = 'Last-Minute Cancellation';
  await topic('Last-Minute Cancellation (LMC)', 'The fee for cancelling within 6 hours of the start time.',
    ['What the LMC fee is', 'When it\'s automatically waived', 'Refund by customer category', 'Failed charges'],
    'ill27.png',
    'Category: Job-Related Charges & Refund Requests. LMC refund requests are common, and getting them wrong fails in two opposite directions: too generous gives away money on a valid charge; too strict frustrates a customer over a fee that shouldn\'t have applied (e.g. the CP caused the disruption). Two separate questions: is the charge valid, and if so, how generous to be. Mixing them up is where most mistakes happen.');

  {
    const s = content(LMC, 'What is the LMC fee?', 'Charged automatically. You check it, you don\'t apply it.',
      'A Last-Minute Cancellation (LMC) fee is a $40 charge applied when a customer cancels a job within 6 hours of its start time. The matched CP is paid $20 for the inconvenience of losing the booking so close to the start. The system charges it automatically: you\'re not deciding whether it applies, only whether it was applied correctly and whether it should be refunded.');
    const w = (9.1 - 0.3) / 3;
    stat(s, 0.45, 1.45, w, 1.7, '$40', 'Charged to the customer');
    stat(s, 0.45 + w + 0.15, 1.45, w, 1.7, '< 6 hrs', 'Cancelled this close to the start time');
    stat(s, 0.45 + 2 * (w + 0.15), 1.45, w, 1.7, '$20', 'Paid to the matched CP', C.gold);
    await tip(s, 3.4, 'Two questions:', 'is the charge valid? Only then: how generous should you be?', 'FiHelpCircle');
  }

  {
    const s = content(LMC, 'Step 1: is the charge valid?', 'LMC is automatically waived in any of these cases.',
      'LMC is automatically waived (it shouldn\'t have been charged, or should be reversed if it was) when: the job was still Submitted (no CP matched, so no one to compensate); the customer\'s area had severe weather; the CP rescheduled less than 24 hours before the cancellation and the new time wasn\'t requested by the customer; the CP claimed an Alternate Start Time added less than 48 hours before the cancellation; the customer asked for the CP\'s ETA and the CP didn\'t check in within 30 minutes of the start time; or all of these: the customer hasn\'t had a prior LMC waiver, the CP didn\'t check in more than 30 min before start, and the job was claimed or rescheduled in the last 24 hours with either an active FC or a start time that wasn\'t the customer\'s preferred one.');
    const items = [
      ['FiInbox', 'Job still Submitted', 'No CP matched: no one to compensate.'],
      ['FiCloudRain', 'Severe weather', 'In the customer\'s area.'],
      ['FiRepeat', 'CP rescheduled < 24 hrs before', 'And the customer didn\'t ask for the new time.'],
      ['FiClock', 'Alternate Start Time claimed', 'Added < 48 hrs before the cancellation.'],
      ['FiMapPin', 'ETA asked, no check-in', 'CP didn\'t check in within 30 min of start.'],
      ['FiLayers', 'All of these together', 'No prior LMC waiver · CP didn\'t check in > 30 min before · claimed or rescheduled in the last 24 hrs (active FC or not their preferred time).'],
    ];
    const w = (9.1 - 0.15) / 2, h = 0.98;
    for (let i = 0; i < items.length; i++) {
      const x = 0.45 + (i % 2) * (w + 0.15), y = 1.42 + Math.floor(i / 2) * (h + 0.1);
      box(s, x, y, w, h, C.white, C.border);
      await iconDot(s, x + 0.15, y + 0.15, items[i][0], 0.36);
      T(s, items[i][1], { x: x + 0.62, y: y + 0.12, w: w - 0.75, h: 0.3, fontFace: HEAD, bold: true, fontSize: 11, valign: 'middle' });
      T(s, items[i][2], { x: x + 0.62, y: y + 0.44, w: w - 0.75, h: h - 0.5, fontSize: 9, color: C.soft });
    }
  }

  {
    const s = content(LMC, 'Invalid charge: reverse both sides', 'E.g. the customer cancelled because the CP wanted to reschedule or cancel, or the CP no-showed.',
      'If the charge is invalid: 1) Refund the LMC charge via the Manual Charges section. 2) Create a Cleaner Payment Holdback for the $20 the CP received: CP CRM > Do > CleanerPaymentHoldback > create (the amount is entered in cents, so $20 = 2000). Special case: the customer was charged LMC but the job was later completed anyway. Refund the LMC charge, create a $20 holdback with invalid_lmc as the holdback type, log the ticket, then recreate the job and invoice it normally so it still triggers the customer\'s FC correctly. TRAINER: in slideshow mode, click Refund the LMC charge or Hold back the CP\'s $20 to play that how-to video.');
    const w = 4.0, cx1 = 0.45, cx2 = 0.45 + w + 0.4;
    eyebrow(s, 'Charge was invalid', cx1, 1.5, 4, C.soft);
    // Clickable: each step opens its how-to video in Drive.
    const videoBtn = (text, y, url, tooltip) => s.addText(text, { shape: pres.shapes.ROUNDED_RECTANGLE, x: cx1, y, w, h: 0.55, rectRadius: 0.08,
      fill: { color: C.teal }, line: { color: C.teal, width: 1 }, fontFace: HEAD, bold: true, fontSize: 10, color: C.white, align: 'center', valign: 'middle',
      hyperlink: { url, tooltip } });
    videoBtn('▶  Refund the LMC charge', 1.8, 'https://drive.google.com/file/d/1dGh4N34uNXt2b5cvOZ1YuZrXYro7ufBa/view?usp=sharing', 'Play: how to refund the LMC fee');
    T(s, 'From the Manual Charges section.', { x: cx1, y: 2.4, w, h: 0.3, fontSize: 9.5, color: C.soft, align: 'center' });
    flowLine(s, [[cx1 + w / 2, 2.7], [cx1 + w / 2, 2.9]], true);
    videoBtn('▶  Hold back the CP\'s $20', 2.9, 'https://drive.google.com/file/d/1gqoWAA__DzuEuV7R4xlQCMbLl--T5nq6/view?usp=sharing', 'Play: how to create an LMC fee holdback');
    T(s, 'CP CRM › Do › CleanerPaymentHoldback › create.\nAmount in cents: 2000.', { x: cx1, y: 3.5, w, h: 0.5, fontSize: 9.5, color: C.soft, align: 'center' });
    // Label wording as the trainer edited it in Google Slides (kept as typed, not uppercased).
    T(s, 'SPECIAL CASE: LMC was charged but JOB COMPLETED later anyway', { x: cx2, y: 1.5, w: 4.6, h: 0.2, fontFace: HEAD, bold: true, fontSize: 8, color: C.soft, charSpacing: 0.5 });
    box(s, cx2, 1.8, 9.55 - cx2, 2.2, C.white, C.border);
    T(s, bullets(['Refund the LMC charge', '$20 holdback, type invalid_lmc', 'Log the ticket', 'Recreate the job and invoice it normally, so the customer\'s FC still triggers']),
      { x: cx2 + 0.2, y: 1.95, w: 9.55 - cx2 - 0.4, h: 1.95, fontSize: 10.5, paraSpaceAfter: 6 });
  }

  {
    const s = content(LMC, 'Step 2: valid charge, check the customer category', 'Not a ladder: go straight to the right offer.',
      'Use this table only when the customer explicitly requests a refund, or contests the charge expecting it to be reversed or credited. If they\'re only asking for an explanation and haven\'t expressed dissatisfaction or asked for removal, don\'t use it. It\'s deliberately not a ladder: don\'t start small and escalate only if they push back. New customer (0-2 invoiced jobs): refund the LMC fee outright. They\'re brand new; charging $40 this early risks losing them before they see the membership\'s value. Low to moderate expected future LTNR (3-4 invoiced jobs): offer credits up to 100% of the LMC fee, not cash. If they say no but don\'t threaten to cancel, tell them this is the most we can offer. If they get upset, send them to Service Recovery so someone can explain by phone. If they still want to cancel, put the LMC fee toward their ETF. Very high-value customer (more than 4 invoiced jobs, high retention): refund the whole $40. Macros: Last-minute cancel: high-value customer; Last-minute cancel: no refund (standard customer).');
    table(s, ['Customer', 'Invoiced jobs', 'What to do'], [
      ['New customer', '0–2', 'Refund the LMC fee outright.'],
      ['Low to moderate LTNR', '3–4', 'Credits up to 100% of the fee, not cash. Upset? Service Recovery. Cancelling? Apply it to the ETF.'],
      ['Very high-value', 'More than 4, high retention', 'Refund the full $40.'],
    ], { y: 1.45, colW: [2.2, 2.0, 4.9], fontSize: 10, rowH: [0.34, 0.5, 0.62, 0.5] });
    await tip(s, 3.55, 'Only if asked:', 'use this when they want the charge reversed, not when they just want it explained.', 'FiMessageSquare');
    await macroFlag(s, 0.45, 4.15, 'Recommended Macros:', 'Last-minute cancel: high-value customer · Last-minute cancel: no refund (standard customer)', 9.1);
  }

  {
    const s = content(LMC, 'Finding the charge, and when it fails', 'A failed LMC charge is finished. Leave it.',
      'Where to find LMC charges: on the CRM under Manual Charges, or C CRM > View > ManualCustomerPaymentTxn. When issuing credits for an LMC, use "lmc" as the internal reason. If the LMC charge fails to collect, let it go: the system makes a couple of attempts and then stops. We don\'t collect it again, and we don\'t pay the cleaner from it. There\'s nothing to void and nothing to run through the manual-charge refund flow: the charge never landed. This applies to LMC specifically: other fees don\'t all behave this way (a failed Premium fee keeps retrying until it succeeds), so check the fee before assuming a failed charge is finished. The cleaner is paid only if they ask: a CP who reaches out for compensation on a last-minute cancellation is handled on its own merits through the usual CP-side path. A failed charge looks like an open loop and the instinct is to chase it: void it, refund it, re-run it. All three are wrong.');
    await twoCol(s,
      { ico: 'FiSearch', title: 'Where to find it', items: ['CRM › Manual Charges', 'C CRM › View › ManualCustomerPaymentTxn', 'Credits reason code: lmc'] },
      { ico: 'FiSlash', title: 'Charge failed?', items: ['Let it go: the system stops retrying', 'Don\'t void, refund or re-run it', 'CP paid only if they ask (CP-side path)'] },
      1.45, 2.1);
    await tip(s, 3.75, 'LMC only:', 'other fees behave differently. A failed Premium fee keeps retrying.', 'FiAlertCircle');
  }

  await knowledgeCheck(LMC, [
    ['A customer cancels 3 hours before start. The job was still Submitted. LMC?', 'No. Automatically waived: no CP was matched.'],
    ['The CP no-showed and the customer cancelled. LMC was charged. What do you do?', 'Refund it from Manual Charges and create a $20 CleanerPaymentHoldback.'],
    ['Valid LMC, 3 invoiced jobs, customer wants a refund. What do you offer?', 'Credits up to 100% of the fee, not cash.'],
    ['The LMC charge failed to collect. What now?', 'Nothing. Let it go: don\'t void, refund or re-run it.'],
  ], 'Ask each question and let trainees answer before revealing. Each click reveals the next answer.');

  await wrapUp(LMC, [
    ['Validate first', 'Check the six automatic waivers before anything else.'],
    ['Invalid? Both sides', 'Refund the customer, hold back the CP\'s $20.'],
    ['Valid? Not a ladder', 'Go straight to the category\'s offer.'],
  ], 'TRAINER: recap the three takeaways, then open the floor for questions on Last-Minute Cancellation.');

  // ================= 3. LOCKOUT =================
  const LO = 'Lockout';
  await topic('Lockout (LO)', 'The CP showed up but couldn\'t get in. Is the charge fair?',
    ['What a lockout is', 'Four questions to validate it', 'Refund by customer category', 'Special scenarios and lockout pay'],
    'ill7.png',
    'Category: Job-Related Charges & Refund Requests. The tension is the same as LMC: the CP showed up and deserves to be paid for that time, but the charge is only fair if the CP met their end of the bargain (arrived on time, tried to get in, stayed the required time). Validate the claim before deciding anything about a refund: refunding a lockout that was never valid isn\'t generosity, it\'s correcting an error.');

  {
    const s = content(LO, 'What is a lockout?', 'A claim the CP submits when they can\'t get in.',
      'A Lockout is a claim a CP submits when they arrive at the customer\'s home at the scheduled time but can\'t get in to complete the cleaning, despite trying to reach the customer. It exists so the CP is still compensated for the time, travel and effort of showing up. Once submitted, the job status changes to Invoiced and the customer is charged the full job cost. That\'s why lockout tickets almost always arrive as a customer contesting a charge for a cleaning that never happened. This topic also covers what to do when a CP contacts us about a lockout (lockout pay).');
    await steps(s, 1.45, 2.0, [
      ['CP can\'t get in', 'Arrived at the scheduled time and tried to reach the customer.'],
      ['Lockout submitted', 'The job changes to Invoiced.'],
      ['Customer charged', 'The full job cost, for a cleaning that didn\'t happen.'],
      ['Customer contests', 'That\'s the ticket you\'ll see.'],
    ]);
    await tip(s, 3.7, 'Validate first:', 'refunding an invalid lockout isn\'t generosity. It\'s fixing an error.', 'FiShield');
  }

  {
    const s = content(LO, 'Step 1: validate the claim', 'Ask in order. Any "no" usually means the lockout is invalid.',
      'Ask these four questions in order. Any "no" typically means the lockout is invalid. Always leave an internal note documenting your findings: it\'s your reference if either party disputes the outcome later. Next slide: what evidence to check for each question.');
    const w = 4.6, x = 0.45, h = 0.44, gap = 0.16;
    const qs = ['Did the customer want the appointment?', 'Did the CP arrive on time?', 'Did the CP arrive at the address?', 'Did the CP try to complete it?'];
    for (let i = 0; i < qs.length; i++) {
      const y = 1.42 + i * (h + gap);
      await flowBox(s, x, y, w, h, `${i + 1}  ${qs[i]}`, 'decision');
      if (i < qs.length - 1) flowLine(s, [[x + w / 2, y + h], [x + w / 2, y + h + gap]], true);
      T(s, 'Yes', { x: x + w / 2 + 0.08, y: y + h, w: 0.5, h: gap, fontSize: 8, bold: true, color: C.soft, valign: 'middle' });
      flowLine(s, [[x + w, y + h / 2], [6.2, y + h / 2]]);
      T(s, 'No', { x: x + w + 0.1, y: y + h / 2 - 0.2, w: 0.5, h: 0.2, fontSize: 8, bold: true, color: C.soft });
    }
    // Every "no" feeds one outcome box.
    flowLine(s, [[6.2, 1.42 + h / 2], [6.2, 1.42 + 3 * (h + gap) + h / 2]]);
    flowLine(s, [[6.2, 2.54], [6.5, 2.54]], true);
    await flowBox(s, 6.5, 2.24, 3.05, 0.6, 'Invalid: refund via CP Dashboard', 'penalty');
    const yEnd = 1.42 + 3 * (h + gap) + h;
    flowLine(s, [[x + w / 2, yEnd], [x + w / 2, yEnd + 0.18]], true);
    await flowBox(s, x, yEnd + 0.18, w, 0.45, 'All yes: valid. Check the customer category', 'action');
    await tip(s, 4.45, 'Always:', 'leave an internal note with your findings.', 'FiEdit3');
  }

  {
    const s = content(LO, 'What to check for each question', 'Evidence that\'s hard to fake beats one side\'s word.',
      '1) Customer\'s messages to the CP. Mentioned cancelling or rescheduling more than 6 hours before start: invalid, the CP shouldn\'t have gone. Within 6 hours: if the CP acknowledged the appointment wouldn\'t happen, still invalid. 2) CPs may arrive 30 minutes before up to 30 minutes after the start time. Check call/text timestamps, CP-to-Care comms, check-in messages, GPS screenshots, CP App "approaching" notifications. More than 30 min late: invalid. Exception: the evidence shows the customer had already decided not to have the appointment (e.g. told Support but the job wasn\'t cancelled and the CP wasn\'t told; or told the CP after start time that they\'d cancelled long before). The CP still has to show they were on site and tried to reach the customer. Coach lateness separately. 3) CP messages confirming arrival, CTJ logs, pre-evidence photos, or a selfie at the home. A photo of the house without the CP in frame: check it isn\'t reused or reverse-searchable. A photo is only required for CP Dashboard submissions, not the CP App. Faked (not just weak) evidence: stop and go to AG – Fraud › CP Document Fraud. 4) At least 2 contact attempts, at least 5 minutes apart, staying at least 15 minutes, and following any entry notes. Same rule for CP App and CP Dashboard. Confirm with timestamps: contact attempts, the lockout action in Job History, CP messages to the customer or CS. Waived if the CP says the customer refused entry or asked to cancel on arrival: then look for other proof the CP was on site on time.');
    await checkRows(s, 1.42, [
      ['Wanted the appointment', 'Customer asked to cancel or reschedule > 6 hrs before? Invalid. Within 6 hrs and the CP acknowledged it? Invalid.'],
      ['Arrived on time', '30 min before to 30 min after start. Timestamps, check-ins, GPS, CP App "approaching". Later? Invalid, unless the customer had already decided.'],
      ['At the address', 'CTJ, arrival messages, photos, a selfie at the home. A house photo alone: check it isn\'t reused. Faked? AG – Fraud.'],
      ['Tried to complete it', '2+ contact attempts, 5+ min apart, 15+ min on site, entry notes followed. Same rule for App and Dashboard.'],
    ], 0.74, 0.1);
  }

  {
    const s = content(LO, 'Invalid: refund via the CP Dashboard', 'The claim was the error, so the CP pays for it.',
      'If the lockout doesn\'t hold up under Step 1, refund the charge through the CP Dashboard refund flow, not an admin refund. Also check for a Premium rate charge on the same job and refund that too if applicable. Why CP Dashboard: it charges the cost back to the CP, because the CP made the mistake (an invalid lockout claim). An admin refund comes out of Homeaglow\'s pocket as a goodwill gesture, but the customer did nothing wrong, so there\'s nothing to smooth over. An admin refund would make Homeaglow pay for the CP\'s error.');
    await twoCol(s,
      { ico: 'FiCheckCircle', title: 'CP Dashboard refund', items: ['Charges the cost back to the CP', 'The CP made the error', 'Also refund Premium, if charged'] },
      { ico: 'FiXCircle', title: 'Not an admin refund', items: ['Comes out of Homeaglow\'s pocket', 'A goodwill gesture: not needed here', 'Would make us pay for the CP\'s error'] },
      1.45, 2.3);
  }

  {
    const s = content(LO, 'Step 3: valid lockout, check the customer category', 'Same as LMC: go straight to the right offer.',
      'Same non-ladder principle as LMC. New customer (1-2 invoiced jobs): admin refund the full job cost; refunding now and educating protects a relationship that\'s still forming. Low to moderate expected future LTNR (3-4 invoiced jobs): credits up to 75% of the job cost, keeping at least $40 to cover the CP\'s costs. Declined but not escalating: this is the maximum. Escalates: route to Service Recovery. Insists on cancelling: apply the job cost toward the ETF (max reduction still 50%). Very high-value customer (more than 4 invoiced jobs, high retention): offer two options: the full LO fee as account credit, or a refund of the LO fee minus $40 (the CP cost). Explain why the fee exists either way. Applies to customers with an active FC in most cases; a customer actively using the service without an FC table is an edge case, use judgment. Deactivated FC, RC still active (cancelled FC but the Recurring Cleaning plan stayed on and a job still happened): admin refund the full charge and deactivate the RC plan. That\'s a product gap: FC self-cancel doesn\'t cancel RC. If vague, probe: check for a legitimate explanation (e.g. a messaging bug) before assuming the customer is at fault. It\'s fine to open with the full job cost as credits as a goodwill gesture before working through the tiers. Macros: Lockout fee: high-value customer; Lockout fee: no refund (standard customer).');
    table(s, ['Customer', 'Invoiced jobs', 'What to do'], [
      ['New customer', '1–2', 'Admin refund the full job cost.'],
      ['Low to moderate LTNR', '3–4', 'Credits up to 75% of the job, keep $40 for the CP. Escalates? Service Recovery. Cancelling? Toward ETF (max 50% off).'],
      ['Very high-value', 'More than 4, high retention', 'Choice: full fee as credit, or refund minus $40. Explain the fee.'],
      ['Deactivated FC, RC still on', 'Any', 'Admin refund in full. Deactivate the RC plan.'],
    ], { y: 1.42, colW: [2.2, 1.9, 5.0], fontSize: 9.5, rowH: 0.56 });
    await macroFlag(s, 0.45, 4.62, 'Recommended Macros:', 'Lockout fee: high-value customer · Lockout fee: no refund (standard customer)', 9.1);
  }

  {
    const s = content(LO, 'Special scenarios', 'When the miss isn\'t the customer\'s.',
      'Customer denied entry because of the CP or a Trust & Safety concern (showed up unprofessional, brought an unauthorized guest, didn\'t have cleaning supplies): issue an admin refund for the lockout fee, still pay the CP $40 for their time and travel, document the incident clearly in internal notes, and coach the CP on why the customer denied entry. The customer\'s safety and trust come first, but the CP still showed up and is owed something. Customer cancelled through Support, but we didn\'t act on it (the job wasn\'t cancelled and the CP wasn\'t told): the miss is ours. Don\'t charge the customer: refund or cancel the job on the C-side. Pay the CP the standard lockout amount through MCP, as long as they show they were on site and tried to reach the customer. Their lateness doesn\'t matter here.');
    await twoCol(s,
      { ico: 'FiShield', title: 'Entry denied: CP fault or T&S', items: ['Admin refund the lockout fee', 'Still pay the CP $40', 'Document it in internal notes', 'Coach the CP'] },
      { ico: 'FiPhoneMissed', title: 'Cancelled via Support, not actioned', items: ['Don\'t charge the customer', 'Refund or cancel on the C-side', 'Pay the CP via MCP, if on site and tried', 'Lateness doesn\'t matter here'] },
      1.45, 2.55);
  }

  {
    const s = content(LO, 'Lockout pay: when the CP contacts us', 'Treat any lockout report as a request for lockout pay review.',
      'When a CP reports a lockout, whether they ask for lockout pay directly or only describe what happened, treat it as a request for lockout pay review. Signals: the CP mentions time waited, travel, gas or cost, or lost earnings, or asks what happens with the job. The exception is a CP who clearly only wants the job cancelled. Validate using Step 1, including the lateness exception. If the customer had cancelled through Support and we didn\'t act, follow that special scenario. Outcomes: valid lockout the CP didn\'t submit (didn\'t know how, app issue): pay the standard lockout amount. Invalid lockout but evidence the CP was on site: one-time courtesy MCP if the CP has no cp_false_check_in or false invoice flags, isn\'t suspended, and hasn\'t had courtesy lockout pay in the last 90 days. No evidence the CP was on site: no pay; explain the lockout requirements. MCPs over $40 still need Support approval. Why: requiring specific words would reward CPs who know our terms over CPs who were actually locked out. The courtesy conditions are checkable, so every agent and QA reviewer reaches the same answer; the 90-day cap protects against abuse.');
    table(s, ['Situation', 'Outcome'], [
      ['Valid lockout, CP didn\'t submit it (didn\'t know how, app issue)', 'Pay the standard lockout amount.'],
      ['Invalid lockout, but evidence the CP was on site', 'One-time courtesy MCP: no cp_false_check_in or false invoice flags, not suspended, none in the last 90 days.'],
      ['No evidence the CP was on site', 'No pay. Explain the lockout requirements.'],
    ], { y: 1.42, colW: [4.0, 5.1], fontSize: 9.5, rowH: [0.34, 0.55, 0.62, 0.5] });
    await tip(s, 3.55, 'Counts as a request:', 'time waited, travel, gas, lost earnings, or "what happens with the job?"', 'FiMessageCircle');
    await tip(s, 4.1, 'MCP over $40?', 'still needs Support approval.', 'FiLock');
  }

  await knowledgeCheck(LO, [
    ['The customer asked the CP to reschedule 8 hours before start. The CP went anyway. Valid lockout?', 'No. The customer had already cancelled more than 6 hours before. Invalid.'],
    ['The CP arrived 45 minutes late and the customer had stepped out. Valid?', 'No. More than 30 minutes late is invalid (unless the customer had already decided).'],
    ['The lockout is invalid. How do you refund?', 'Through the CP Dashboard, not an admin refund. Refund Premium too.'],
    ['Valid lockout, 3 invoiced jobs, wants a refund. What do you offer?', 'Credits up to 75% of the job cost, keeping $40 for the CP.'],
  ], 'Ask each question and let trainees answer before revealing. Each click reveals the next answer.');

  await wrapUp(LO, [
    ['Four questions', 'Wanted it? On time? At the address? Tried to get in?'],
    ['Invalid? CP pays', 'Refund through the CP Dashboard, not admin.'],
    ['Valid? By category', 'Straight to the tier\'s offer. $40 CP cost stays.'],
  ], 'TRAINER: recap the three takeaways, then open the floor for questions on Lockouts.');

  // ---------- Close ----------
  {
    const s = pres.addSlide({ masterName: 'CONTENT' });
    eyebrow(s, 'Wrap-up', 0.62, 1.9);
    T(s, 'Thank you!', { x: 0.62, y: 2.15, w: 5, h: 0.75, fontFace: HEAD, bold: true, fontSize: 34 });
    T(s, 'When in doubt, check the Knowledge Library or ask your Trainer or Team Lead.', { x: 0.62, y: 2.95, w: 4.6, h: 0.5, fontSize: 12, color: C.soft });
    await panel(s, 'ill9.png');
    s.addNotes('Close the module. Recap: Incorrect Information (test the exact words, own the error, one light offer, else waive ETF), LMC (validate against the automatic waivers, then refund by category, not a ladder; leave failed charges alone), and Lockouts (four validation questions; invalid goes through the CP Dashboard; valid goes by category, keeping $40 for the CP).');
  }

  // Point every agenda row (card, number and arrow) at its topic's title slide.
  if (topicSlides.length * 3 !== agendaLinks.length) throw new Error(`agenda has ${agendaLinks.length / 3} rows but the deck has ${topicSlides.length} topics`);
  // pptxgenjs fixes a link's target when it's added, so update the stored relationship.
  agendaLinks.forEach((l, i) => { agendaSlide._rels.find(r => r.rId === l._rId).Target = String(topicSlides[Math.floor(i / 3)]); });
  await pres.writeFile({ fileName: OUT });
  console.log('wrote', OUT);
})();
