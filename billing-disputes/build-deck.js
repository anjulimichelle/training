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
const TOPICS = 9;
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
    T(s, 'Charges for time not delivered, off-platform payments, and cleaners who don\'t show, cancel or reschedule.', { x: 0.62, y: 2.7, w: 4.6, h: 0.7, fontSize: 12.5, color: C.soft });
    T(s, 'Internal training material', { x: 0.62, y: 5.0, w: 3, h: 0.2, fontSize: 7, color: C.soft });
    await panel(s, 'ill6.png', 0.4, 4.82);
    s.addNotes('Welcome. This module covers seven topics: Overcharged Hours, Unauthorized Addition of Hours, False Invoice, Disintermediation (Cash Payment), Cleaner Didn\'t Show, Cleaner Cancellation, and Unauthorized Reschedule. Each topic ends with a knowledge check and a wrap-up with time for questions.');
  }

  // ---------- Agenda ----------
  // Each row links to its topic's title slide. Slide numbers aren't known yet,
  // so the links are filled in from topicSlides once every slide is built.
  const agendaLinks = [], topicSlides = [];
  let agendaSlide;
  const addTopic = topic;
  topic = async (...args) => { const t = await addTopic(...args); topicSlides.push(t._slideNum); return t; };
  {
    const s = agendaSlide = content('Overview', 'Agenda', 'Nine topics. Click a topic to jump to it.',
      'The first three topics are the "unworked hours" family: they all look the same to the customer ("I was charged for time I didn\'t get"), but each has a different question, different evidence and a different resolution. The fourth topic covers customers paying cleaners off the platform. Topics 5 to 7 cover reliability issues: no-shows, cleaner cancellations, and unauthorized reschedules. Topics 8 and 9 cover the specialized escalation teams: Trust and Safety, and Service Recovery. In slideshow mode, click any topic to jump straight to it.');
    const rows = [
      ['Overcharged Hours (OCH)', 'The cleaner billed more time than they worked'],
      ['Unauthorized Addition of Hours', 'The cleaner added time without the customer\'s OK'],
      ['False Invoice', 'The cleaner billed a job they never did'],
      ['Disintermediation: Cash Payment', 'The customer paid the cleaner directly'],
      ['Cleaner Didn\'t Show', 'The cleaner claimed the job but never came'],
      ['Cleaner Cancellation', 'The cleaner cancelled a claimed job'],
      ['Unauthorized Reschedule', 'The cleaner moved the job without the customer\'s OK'],
      ['Trust and Safety', 'Safety, property or conduct risk: escalate it'],
      ['Service Recovery', 'Legal, reputational or relationship risk: hand off'],
    ];
    const gap = 0.15, w = (9.1 - gap) / 2, h = 0.62, gy = 0.09;
    rows.forEach(([title, line], i) => {
      const col = i < 5 ? 0 : 1, row = i < 5 ? i : i - 5;
      const x = 0.45 + col * (w + gap), y = 1.45 + row * (h + gy);
      const link = { slide: 1, tooltip: 'Go to ' + title };
      agendaLinks.push(link);
      // One shape per row, so a click anywhere on it follows the link. The runs share
      // the link object, so pptxgenjs registers one relationship for the whole card.
      s.addText([
        { text: title, options: { color: C.ink, bold: true, fontFace: HEAD, fontSize: 11.5, breakLine: true, hyperlink: link } },
        { text: line, options: { color: C.soft, fontSize: 9, hyperlink: link } },
      ], { shape: pres.shapes.ROUNDED_RECTANGLE, x, y, w, h, rectRadius: 0.06, fill: { color: C.white }, line: { color: C.border, width: 0.75 },
        fontFace: BODY, valign: 'middle', margin: [45, 36, 0, 0], hyperlink: link });  // points: left, right, bottom, top
      const num = { slide: 1, tooltip: 'Go to ' + title };
      agendaLinks.push(num);
      s.addText(String(i + 1).padStart(2, '0'), { x: x + 0.1, y, w: 0.45, h, fontFace: HEAD, bold: true, fontSize: 12, color: C.teal, align: 'center', valign: 'middle', hyperlink: num });
      const arrow = { slide: 1, tooltip: 'Go to ' + title };
      agendaLinks.push(arrow);
      s.addText('›', { x: x + w - 0.42, y, w: 0.3, h, fontFace: HEAD, bold: true, fontSize: 18, color: C.teal, align: 'center', valign: 'middle', hyperlink: arrow });
    });
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
      { ico: 'FiMessageSquare', title: 'Worked but increased the hours', body: 'Unauthorized Addition of Hours' },
      { ico: 'FiClock', title: 'Billed more than worked', body: 'Overcharged Hours: this topic' },
    ]);
    await tip(s, 3.6, 'Why it matters:', 'each one needs different evidence and leads to a different resolution and penalty.', 'FiInfo');
  }

  {
    const s = content(OCH, 'Policies', null,
      'Claim window: only jobs completed in the last 120 days. Exception: the window does not apply when the overcharge is already clearly proven when the customer reports it (CTJ, message history or job timestamps make it obvious; "likely" or "probably" is not enough). Then refund the difference via the CP Dashboard regardless of job age or the cleaner\'s status, and follow the rest of the process including the penalty ladder. Example: CTJ shows a 45-minute visit but 2 hours invoiced, reported 150 days later: refund and apply the ladder. If not clearly proven and beyond 120 days: don\'t request documentation or refund; close the ticket as unresolved instead of escalating (this includes cases where the only reason to refund is the CP\'s Suspended + DNR history). Evidence-based refund: refund only the hours the evidence supports, not automatically the full disputed amount. CP Suspended + DNR: if the cleaner is already suspended + DNR for repeated false invoice/overcharge offenses: within 120 days, refund the excess even without clear evidence, with no need to request evidence from the CP (a documented pattern removes the benefit of the doubt); beyond 120 days, refund the excess only when the issue is clearly proven.');
    await cards(s, 1.4, 2.7, [
      { ico: 'FiCalendar', title: '120-day claim window', body: 'Unless the overcharge is clearly proven when reported. Then the job\'s age doesn\'t matter.' },
      { ico: 'FiSliders', title: 'Refund what\'s proven', body: 'Only the hours the evidence supports, not the full disputed amount.' },
      { ico: 'FiAlertOctagon', title: 'CP Suspended + DNR', body: [
        { text: 'Already DNR for repeat offenses?', options: { breakLine: true } },
        { text: 'Within 120 days: ', options: { bold: true, color: C.ink } }, { text: 'refund the excess even without clear evidence. No need to request evidence from the CP.', options: { breakLine: true } },
        { text: 'Beyond 120 days: ', options: { bold: true, color: C.ink } }, { text: 'refund the excess only if the issue is clearly proven.' },
      ] },
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
      'Branch A (clear evidence present): a straight line once you\'re in it: refund, check for a Premium fee, then split on the 30-minute rule. Branch B (no clear evidence present): takes longer and may span more than one contact: request documentation from the CP and resume when the Timed Reminder triggers. If CTJ matches the hours invoiced, that is clear evidence no overcharge occurred: stop, no refund, address other concerns. Known bug (Trello): do not treat an abnormally short CTJ interval (30 minutes or less between ARRIVED AT and COMPLETED AT) on its own as clear evidence; it\'s inconclusive unless other evidence proves the CP only stayed that long.');
    const branch = (x, w, title) => {
      box(s, x, 1.15, w, 0.62, C.teal);
      T(s, [{ text: title, options: { bold: true, fontFace: HEAD, fontSize: 11.5 } }], { x: x + 0.14, y: 1.15, w: w - 0.28, h: 0.62, color: C.white, valign: 'middle' });
    };
    const grid = (x, y, colW, rows) => {
      const H = ['What you find', 'What it means'].map(t => ({ text: t, options: { bold: true, color: C.teal, fill: { color: C.tealSoft }, fontFace: HEAD } }));
      const R = rows.map(r => r.map((t, i) => ({ text: t, options: { color: i === 0 ? C.ink : C.soft, bold: i === 0, fill: { color: C.white } } })));
      s.addTable([H, ...R], { x, y, w: colW[0] + colW[1], colW, fontFace: BODY, fontSize: 8, valign: 'middle', border: { type: 'solid', pt: 0.75, color: C.border }, margin: [0.03, 0.08, 0.03, 0.08] });
    };
    branch(0.45, 5.45, 'Branch A: clear evidence present');
    grid(0.45, 1.85, [2.35, 3.1], [
      ['CTJ shorter than hours invoiced', 'Clear evidence of overcharge → go to Branch A'],
      ['Customer supplies valid evidence of the actual duration (e.g. Ring camera, timestamped photo)', 'Treat it as you would a CTJ record → go to Branch A, refunding only the hours the evidence supports'],
      ['Cleaner acknowledges overcharge in messages', 'Clear evidence → go to Branch A'],
      ['Cleaner already issued a refund in response to the complaint', 'Supports the overcharge claim (people rarely refund for something they didn\'t do) → go to Branch A'],
      ['Customer claims a shorter cleaning duration than what you found during your investigation', 'CTJ confirms there is an overcharge, regardless of whether it matches the hours reported by the customer → go to Branch A (only refund the hours that are clearly supported by the CTJ)'],
    ]);
    branch(6.05, 3.5, 'Branch B: no clear evidence');
    grid(6.05, 1.85, [1.55, 1.95], [
      ['Customer claims a shorter cleaning duration than what you found during your investigation', 'Job history or other information sources (e.g., C/CP message timestamps) → go to Branch B'],
      ['None of the above found', 'No clear evidence → go to Branch B'],
    ]);
    box(s, 6.05, 4.05, 3.5, 0.95, C.goldSoft);
    T(s, [{ text: 'CTJ matches hours invoiced', options: { bold: true, color: C.gold, breakLine: true } }, { text: 'Clear evidence NO overcharge occurred → stop, no refund, address other concerns' }], { x: 6.17, y: 4.05, w: 3.28, h: 0.95, fontSize: 8.5, valign: 'middle' });
  }

  let branchA;
  {
    const s = branchA = content(OCH, 'Branch A: clear evidence present', 'A straight line once you\'re in it: refund, check for a Premium fee, then split on the 30-minute rule.',
      'Step 1: refund the overcharged hour(s) via the CP Dashboard; the evidence already settles it. Step 2: if the customer was also charged a Premium rate, refund it: a confirmed overcharge is itself the service issue, no separate request needed. Step 3: check how much was overcharged. 30 minutes or more: apply the cp_overcharged_hours flag, ban the C/CP pairing, send comms to both, then go to the Penalty section. Under 30 minutes: ban the C/CP pairing, send comms to both, done (no flag, no penalty).');
    const cx = 5.0, w = 2.7, h = 0.46;
    await flowBox(s, cx - w / 2, 1.45, w, h, 'Refund via CP Dashboard', 'action');
    flowLine(s, [[cx, 1.45 + h], [cx, 2.08]], true);
    await flowBox(s, cx - w / 2, 2.08, w, h, 'Check for Premium fee', 'action');
    flowLine(s, [[cx, 2.08 + h], [cx, 2.71]], true);
    await flowBox(s, cx - w / 2, 2.71, w, h, 'Overcharged ≥ 30 min?', 'decision');
    flowSplit(s, cx, 2.71 + h, 3.42, 2.45, 7.55, 3.75, 'Yes', 'No');
    await flowBox(s, 2.45 - w / 2, 3.75, w, h, 'Flag, ban, penalize', 'penalty');
    await flowBox(s, 7.55 - w / 2, 3.75, w, h, 'Ban the pairing', 'end');
    T(s, 'cp_overcharged_hours flag  → norequests status  → ban the pairing  → send comms to C & CP', { x: 2.45 - 1.6, y: 4.27, w: 3.2, h: 0.45, fontSize: 8.5, color: C.soft, align: 'center' });
    T(s, 'Comms to both. Done: no flag, no penalty.', { x: 7.55 - 1.6, y: 4.27, w: 3.2, h: 0.45, fontSize: 8.5, color: C.soft, align: 'center' });
  }

  {
    const s = content(OCH, 'Branch B: no clear evidence present', 'Takes longer: it pauses while we wait on the CP, and picks back up when the Timed Reminder triggers.',
      'If the overcharge amount isn\'t known yet, ask the customer for detail first (B1). Is the CP suspended + DNR due to multiple false invoice or overcharged hours? Yes: skip documentation and treat it as valid: go to Branch A from step 1. No: request documentation. Refund the Premium Upsell fee if charged; create a 72-hour Timed Reminder on the C side; tell the customer we\'ll follow up in 4-5 business days and the CP is blocked from their future requests (not penalized yet); block the pairing and tell the CP documentation is due within 3 days or the alleged hours will be refunded on their behalf. When the reminder triggers: valid proof from the CP: no refund (confirm the CP stays blocked from this customer). No or insufficient proof: refund the alleged hours via the CP Dashboard, check/refund Premium, log to the Ticket Tracker, then the penalty check. Timed Reminder format: Action On: 72 hours from the time of ticket handling. Who Should Act: Any CS. Action: C <C_ID> reported that CP overcharged J <JOB_ID> for <# OF OVERCHARGED HOURS>. Follow When the Timed Reminder Triggers. TRAINER DEMO (click the Trainer demo button on the Timed Reminder box to play the video): show trainees how to add a Timed Reminder in the Legacy C CRM (Do > TimedReminder > Create): set it 72 hours out, assign Any CS, and paste the Action text with the real C ID, job ID and hours.' + ' TRAINER: the penalty check shows when the slide opens (30 minutes or more overcharged: add the cp_overcharged_hours flag + norequests; under 30 minutes: no flag, no penalty). Click once (Request docs) to show the Timed Reminder format; click again (Refund, then penalty check) to go back to the penalty check.');
    // Yes (skip to Branch A) on the left; No (request docs, then the reminder) on the right.
    const lx = 2.05, rx = 6.2, w = 2.8, h = 0.52;
    const button = (text, x, y, bw, bh, fill, line, opts = {}) => s.addText(text, Object.assign({ shape: pres.shapes.ROUNDED_RECTANGLE, x, y, w: bw, h: bh, rectRadius: 0.08,
      fill: { color: fill }, line: { color: line, width: fill === line ? 1 : 2 }, fontFace: HEAD, bold: true, fontSize: 10, color: C.white, align: 'center', valign: 'middle' }, opts));
    await flowBox(s, 5.0 - 1.75, 1.42, 3.5, 0.55, 'CP suspended + DNR due to multiple false invoice or overcharged hours?', 'decision');
    flowSplit(s, 5.0, 1.97, 2.2, lx, rx, 2.45, 'Yes', 'No');
    // Clickable: jumps back to the Branch A slide.
    button('Skip to Branch A  ›', lx - w / 2, 2.45, w, h, C.teal, C.teal, { hyperlink: { slide: branchA._slideNum, tooltip: 'Go to Branch A' } });
    // These two look like buttons: click 1 opens the Timed Reminder format, click 2 returns to the penalty check.
    button('Request docs\n72-hour Timed Reminder  ›', rx - w / 2, 2.45, w, h, C.teal, C.teal);
    flowLine(s, [[rx, 2.45 + h], [rx, 3.12]], true);
    await flowBox(s, rx - w / 2, 3.12, w, h, 'Reminder triggers\nCP provided valid proof?', 'decision');
    flowSplit(s, rx, 3.12 + h, 3.84, 4.95, 7.45, 4.05, 'Yes', 'No');
    await flowBox(s, 4.95 - 1.15, 4.05, 2.3, 0.46, 'No refund', 'end');
    button('Refund, then penalty check  ›', 7.45 - 1.2, 4.05, 2.4, 0.46, C.ink, C.teal);
    // Default view: the penalty check, under the Yes side.
    const px = 0.45, py = 3.12, pw = 3.2, ph = 1.39;
    const panel = (name) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: px, y: py, w: pw, h: ph, fill: { color: C.white }, line: { color: C.teal, width: 1.5 }, rectRadius: 0.06,
      shadow: { type: 'outer', color: '000000', opacity: 0.18, blur: 6, offset: 2, angle: 90 }, objectName: name });
    panel('Penalty panel');
    s.addShape(pres.shapes.RECTANGLE, { x: px, y: py, w: pw, h: 0.32, fill: { color: C.teal }, line: { color: C.teal } });
    T(s, 'Penalty check', { x: px + 0.15, y: py, w: pw - 0.3, h: 0.32, fontFace: HEAD, bold: true, fontSize: 10.5, color: C.white, valign: 'middle' });
    T(s, [
      { text: '≥ 30 minutes overcharged:  ', options: { bold: true, color: C.teal } },
      { text: 'add flag cp_overcharged_hours + norequests', options: { breakLine: true } },
      { text: '< 30 minutes overcharged:  ', options: { bold: true, color: C.teal } },
      { text: 'no flag, no penalty' },
    ], { x: px + 0.15, y: py + 0.42, w: pw - 0.3, h: ph - 0.5, fontSize: 9.5, paraSpaceAfter: 6, valign: 'top' });
    // Pop-up: the Timed Reminder format covers the penalty check on click 1 and closes on click 2.
    panel('pop1_2Panel');
    T(s, 'Timed Reminder format', { x: px + 0.15, y: py + 0.05, w: 2.0, h: 0.28, fontFace: HEAD, bold: true, fontSize: 10, color: C.teal, valign: 'middle', objectName: 'pop1_2Title' });
    // Clickable: opens the Timed Reminder demo video in Drive.
    s.addText('▶  Trainer demo', { shape: pres.shapes.ROUNDED_RECTANGLE, x: px + pw - 1.25, y: py + 0.07, w: 1.15, h: 0.26, rectRadius: 0.06, margin: [2, 2, 0, 0],
      fill: { color: C.gold }, line: { color: C.gold }, fontSize: 8, bold: true, color: C.white, align: 'center', valign: 'middle', objectName: 'pop1_2Tag',
      hyperlink: { url: 'https://drive.google.com/file/d/1iPyoRBTJAQ5KWMQwxaxtWI0m-hbjzQZl/view?usp=drive_link', tooltip: 'Play the Timed Reminder demo' } });
    T(s, [
      { text: 'Action On  ', options: { bold: true, color: C.ink } }, { text: '72 hours from the time of ticket handling', options: { breakLine: true } },
      { text: 'Who Should Act  ', options: { bold: true, color: C.ink } }, { text: 'Any CS', options: { breakLine: true } },
      { text: 'Action  ', options: { bold: true, color: C.ink } }, { text: 'C <C_ID> reported that CP overcharged J <JOB_ID> for <# OF OVERCHARGED HOURS>' },
    ], { x: px + 0.15, y: py + 0.38, w: pw - 0.3, h: ph - 0.45, fontSize: 8.5, color: C.soft, paraSpaceAfter: 2, valign: 'top', objectName: 'pop1_2Body' });
    await macroFlag(s, 0.45, 4.72, 'Recommended macros:', 'Overcharged hrs');
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
      'Cleaner refunds after being notified: treat the refund as a sign the overcharge happened unless evidence proves otherwise; if 30 minutes or more, still apply the flag (a voluntary refund confirms the violation, it doesn\'t erase it); if it happens before the first-response Timed Reminder triggers, invalidate that reminder; tell the customer the cleaner refunded. Spotted a likely overcharge before the customer reported it? Reach out proactively. Sample: "Just checking in to see how your recent cleaning with <<Cleaner Name>> on <<Date>> went. We also wanted to quickly confirm the number of hours worked during the visit, just to make sure everything looks right on your end. Let us know if you have any questions or if anything seems off - we\'re happy to help!"');
    await cards(s, 1.3, 3.0, [
      { ico: 'FiRotateCcw', title: 'Cleaner refunds first', body: [
        { text: 'The cleaner refunds after being told about the report.', options: { breakLine: true } },
        { text: 'Treat the refund as confirmation the overcharge happened, unless evidence proves otherwise.', options: { breakLine: true } },
        { text: '30 min or more: ', options: { bold: true, color: C.ink } }, { text: 'still apply the cp_overcharged_hours flag.', options: { breakLine: true } },
        { text: 'Before the first-response Timed Reminder triggers: ', options: { bold: true, color: C.ink } }, { text: 'invalidate that reminder. Tell the customer the cleaner refunded.' },
      ] },
      { ico: 'FiEye', title: 'You spot it first', body: [
        { text: 'A job looks overcharged, but the customer hasn\'t contacted us.', options: { breakLine: true } },
        { text: 'Don\'t wait: ', options: { bold: true, color: C.ink } }, { text: 'reach out to check how the cleaning went and confirm the number of hours worked.', options: { breakLine: true } },
        { text: 'Why: ', options: { bold: true, color: C.ink } }, { text: 'it builds trust and catches the discrepancy before it becomes a dispute.' },
      ] },
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
    ['The customer self-invoiced 2 hrs for a job the cleaner worked 1 hr 11 min. Penalize the cleaner?', 'No. Just refund the difference. The customer billed the hours, so this is not the cleaner’s fault.'],
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
      'TRAINER: descriptions are hidden. Click Show details (or press the right arrow): each click types out the next step\'s description. ' + 'Step 1: refund only the alleged unauthorized hours via the cleaner\'s dashboard, not the full job; the base appointment was authorized and worked. Step 2: coach the cleaner on asking permission before extending; send comms about the action taken; give them 3 days to provide documentation proving the customer authorized the extra hours. Refunding first protects the customer\'s money without making them wait on an investigation. Step 3: if a Premium fee was charged, refund it; unauthorized hours is itself the service issue. Step 4: tell the customer a refund was processed on their cleaner\'s behalf, remind them to leave clear instructions for future cleaners, and block the C/CP pairing. Macros: Unauthorized Addition of Hours (customer refund); Unauthorized Addition of Hours > CP Coaching.');
    const steps = [
      ['Refund the unauthorized hours only', 'action', 'Refund via the CP Dashboard for the alleged unauthorized hours only, not the full job.'],
      ['Coach CP + request evidence', 'action', 'Send coaching comms to the CP. The CP has 3 days to prove the customer authorized the extra hours.'],
      ['Check for Premium charge', 'decision', 'If a Premium fee was charged, refund it.'],
      ['Close the loop with the customer, done', 'penalty', 'Tell the customer the refund was processed on their cleaner\'s behalf. Remind them to leave clear instructions for future cleaners. Block the C/CP pairing.'],
    ];
    await detailsButton(s, 7.6, 0.95);
    const x = 0.6, w = 3.2, h = 0.6, gap = 0.27;
    for (let i = 0; i < steps.length; i++) {
      const y = 1.5 + i * (h + gap);
      await flowBox(s, x, y, w, h, steps[i][0], steps[i][1]);
      if (i < steps.length - 1) flowLine(s, [[x + w / 2, y + h], [x + w / 2, y + h + gap]], true);
      flowLine(s, [[x + w, y + h / 2], [x + w + 0.3, y + h / 2]]);
      T(s, steps[i][2], { x: x + w + 0.4, y, w: 5.15, h, fontSize: 9.5, color: C.soft, valign: 'middle', objectName: `type${i + 1}Desc` });
    }
  }

  {
    const s = content(UAH, 'Scenario: the helper', 'The CP brought a helper, worked half the duration, and charged in full.',
      'This is the one real fork in this topic. Example: the customer booked 4 hours; the CP arrived with a helper, the two of them finished in 2 hours, and the CP invoiced the full 4 hours. Ask: is there evidence the customer authorized the helper and the reduced hours (for example, a message where the customer agreed)? Yes: advise the customer that records show they authorized it; address other pain points; do not penalize the CP. No: follow the main flow: refund the unauthorized hours, request documentation from the CP (3 days), and coach that helpers require permission and can\'t justify full-rate billing. If the CP\'s document wasn\'t just weak but faked, stop and go to AG - Fraud > CP Document Fraud: invalid evidence means the claim fails; faked evidence is a separate offence with a separate outcome.');
    const cx = 5.0, lx = 2.55, rx = 7.45, w = 3.2;
    await flowBox(s, cx - 1.75, 1.35, 3.5, 0.58, 'Evidence says C authorized\nhelper + reduced hours?', 'decision');
    flowSplit(s, cx, 1.93, 2.25, lx, rx, 2.6, 'Yes', 'No');
    await flowBox(s, lx - w / 2, 2.6, w, 0.55, 'No penalty\nC agreed to this', 'end');
    await flowBox(s, rx - w / 2, 2.6, w, 0.55, 'Request documentation\nSame as the main flow', 'penalty');
    const desc = (x, items) => T(s, items.map(([b, t], i) => [{ text: b, options: { bold: true, color: C.ink } }, { text: t, options: { breakLine: i < items.length - 1 } }]).flat(),
      { x: x - w / 2, y: 3.28, w, h: 1.25, fontSize: 9, color: C.soft, paraSpaceAfter: 3, valign: 'top' });
    desc(lx, [['Tell the customer ', 'records show they authorized the helper and the shorter time.'], ['Address ', 'any other pain points.'], ['Don\'t ', 'penalize the CP.']]);
    desc(rx, [['Follow Branch B.', ''], ['Request ', 'documentation from the CP (3 days).'], ['Coach: ', 'helpers need permission and can\'t justify full-rate billing.']]);
    await tip(s, 4.6, 'Faked documents?', 'stop and go to AG – Fraud (CP Document Fraud). It\'s a separate offence.', 'FiAlertTriangle');
  }

  await knowledgeCheck(UAH, [
    ['Booked 3 hrs. The cleaner worked 5 hrs and invoiced 5 without asking. OCH or Unauthorized Hours?', 'Unauthorized Hours. The work was real, but the customer never gave permission.'],
    ['What do you refund?', 'Only the 2 added hours, via the CP Dashboard. Plus Premium if charged. Not the full job.'],
    ['Do you wait for the cleaner\'s proof before refunding?', 'No. Refund first. The cleaner then has 3 days to prove the customer authorized it.'],
    ['The cleaner is permanently suspended. Do you request documentation?', 'No. There\'s no reactivation path, so resolve directly with the customer.'],
  ], 'Ask each question and let trainees answer before revealing. Each click reveals the next answer.');

  {
    const s = content(UAH + '  ·  Practice', 'Let\'s practice!', 'Review the live ticket and investigate.',
      'TRAINER: provide a live ticket for trainees to review. Give them time to investigate (job history, CTJ, C/CP messages, Premium charge), then discuss. Ask: Is it Unauthorized Hours, Overcharged Hours, or something else? Did the customer authorize the extra time? What actions will you take (refund only the added hours via the CP Dashboard, refund Premium if charged, coach the CP and request evidence within 3 days, block the C/CP pairing)? How will you respond to the customer?');
    await card(s, 0.45, 1.45, 3.9, 3.0, { ico: 'FiInbox', title: 'Live ticket', body: 'Your trainer will provide a live ticket. Review it and investigate before you decide anything.' });
    box(s, 0.65, 3.85, 1.35, 0.3, C.goldSoft);
    T(s, 'Trainer provides', { x: 0.65, y: 3.85, w: 1.35, h: 0.3, fontSize: 8, bold: true, color: C.gold, align: 'center', valign: 'middle' });
    await card(s, 4.5, 1.45, 5.05, 1.42, { n: 1, title: 'What will be your actions?', body: 'Which issue is it, and what will you do on the account?' });
    await card(s, 4.5, 3.03, 5.05, 1.42, { n: 2, title: 'How will you respond to the customer?', body: 'Write the reply you would send.' });
  }

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
    const s = content(FI, 'Confirmed false invoice: the process', null,
      'TRAINER: descriptions are hidden. Click Show details (or press the right arrow): each click types out the next step\'s description. ' + '1) Issue a full refund via the CP Dashboard (its description, How to refund from CP dashboard, links to the how-to video: click it to play). 2) Check the cleaner\'s Flags table for an existing cp_false_invoice flag on this job. 3) If the system already flagged it, validate that the Job ID matches; on the flag\'s Django page set Value to {"admin": "<YOUR CRM NAME>", "admin_flag_comments": "<THE ID TRIPLET>", "cp_false_invoice_job_id": <THE JOB ID>}. If no flag exists, add one: CP CRM > Do > Flag > cp_false_invoice. Always include the ID triplet (Customer ID | CP ID | Job ID) and a short summary. 4) Check the Issues Table for prior false invoice / overcharge history and penalize accordingly. Reason: cpq_false_invoice Customer <C\'S NAME> reported incorrect hours charged. Submit reactivation appeal. 5) Coach the cleaner: don\'t claim jobs you can\'t complete or invoice without working; ask the customer to reschedule, or cancel from your dashboard if they disagree. Guardian Angel or Tier 10 banner with pending items: acknowledge, don\'t clear, reassign to allentolentino. 6) Refund the Premium fee if charged (AG - Refund: Premium Upsell). 7) Block the C/CP pairing unless the customer wants the same cleaner. 8) Tell the customer the refund was issued and arrives in 5-10 business days; address other pain points. Macros: False Invoice (C-facing); False Invoice (With Penalty) (CP-facing). Final step: confirm the refund with the customer and offer a priority booking.');
    const rows = [
      ['FALSE INVOICE (CP-fault)', 'penalty', 'Confirm if it was the CP who invoiced the job.'],
      // Description is a link that plays the how-to video.
      ['Issue a full refund via the CP dashboard', 'action', [{ text: '▶  How to refund from CP dashboard', options: { bold: true, color: C.teal,
        hyperlink: { url: 'https://drive.google.com/file/d/1TuwGdwdOWppyezVsG-aITCSja8J-Z0aJ/view?usp=drive_link', tooltip: 'Play: how to refund from CP dashboard' } } }]],
      ['Check for cp_false_invoice flag (Validate or Add)', 'decision', 'Already flagged? Check the Job ID matches. No flag? CP CRM › Do › Flag › cp_false_invoice. Include the ID triplet (C | CP | Job).'],
      ['Apply penalty', 'action', 'If CP invoiced: penalty (cp facing reason cpq_false_invoice Customer <C\'S NAME> reported incorrect hours charged) + false invoice flag'],
      ['Coach / penalize the CP', 'action', 'Send coaching comms to CP'],
      ['Check for Premium fee and refund accordingly', 'action', 'If the customer paid a Premium fee, refund it (AG – Refund: Premium Upsell).'],
      ['Ban the C/CP pairing', 'action', 'Always, unless the customer wants to keep the same CP.'],
      ['Confirm refund with C + offer priority booking', 'end', 'Refund arrives in 5–10 business days. Address any other pain points.'],
    ];
    await detailsButton(s, 7.6, 0.55);
    const x = 0.5, w = 3.55, h = 0.38, gap = 0.1, y0 = 1.15;
    for (let i = 0; i < rows.length; i++) {
      const y = y0 + i * (h + gap);
      await flowBox(s, x, y, w, h, rows[i][0], rows[i][1]);
      if (i < rows.length - 1) flowLine(s, [[x + w / 2, y + h], [x + w / 2, y + h + gap]], true);
      flowLine(s, [[x + w, y + h / 2], [x + w + 0.2, y + h / 2]]);
      T(s, rows[i][2], { x: x + w + 0.3, y: y - 0.03, w: 5.45, h: h + 0.06, fontSize: 8.5, color: C.soft, valign: 'middle', objectName: `type${i + 1}Desc` });
    }
    await macroFlag(s, 6.35, 4.93, 'Recommended Macro:', 'False Invoice');
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

  {
    const s = content(FI + '  ·  Practice', 'Let\'s practice!', 'Review the live ticket and investigate.',
      'TRAINER: provide a live ticket for trainees to review. Give them time to investigate (job history, CTJ, C/CP messages, Premium charge), then discuss. Ask: Who invoiced the job? Is it a False Invoice, Overcharged Hours, or something else? What actions will you take (full refund via the CP Dashboard, validate or add the cp_false_invoice flag, apply the penalty, coach the CP, refund Premium if charged, block the C/CP pairing)? How will you respond to the customer?');
    await card(s, 0.45, 1.45, 3.9, 3.0, { ico: 'FiInbox', title: 'Live ticket', body: 'Your trainer will provide a live ticket. Review it and investigate before you decide anything.' });
    box(s, 0.65, 3.85, 1.35, 0.3, C.goldSoft);
    T(s, 'Trainer provides', { x: 0.65, y: 3.85, w: 1.35, h: 0.3, fontSize: 8, bold: true, color: C.gold, align: 'center', valign: 'middle' });
    await card(s, 4.5, 1.45, 5.05, 1.42, { n: 1, title: 'What will be your actions?', body: 'Which issue is it, and what will you do on the account?' });
    await card(s, 4.5, 3.03, 5.05, 1.42, { n: 2, title: 'How will you respond to the customer?', body: 'Write the reply you would send.' });
  }

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

  let jobStatus;
  {
    const s = content(CASH, 'Key principles', null,
      'Don\'t penalize customers for paying off-platform: they\'re not breaking the rules, the cleaner is. Don\'t act as an investigator when there\'s nothing to go on: with no cleaner report and no proof from the customer, focus on the customer\'s experience. How the cash payment was confirmed changes the tone of the coaching, not the cleaner\'s status. Remind the customer that payment runs through the platform, not cash, unless it\'s a tip: many don\'t realise the job still has to be invoiced, and a cash payment today can still turn into a charge later.');
    // Titles here run long, so give them two lines before the body.
    const items = [
      { ico: 'FiHeart', title: 'Don’t penalize the  customer', body: 'The cleaner broke the rule, not them.' },
      { ico: 'FiSearch', title: 'Don’t act as investigator', body: 'Nothing to go on? Focus on the customer\'s experience.' },
      { ico: 'FiMessageSquare', title: 'Coaching tone - based on how cash payment was confirmed', body: 'How it was confirmed changes the coaching tone.' },
      { ico: 'FiCreditCard', title: 'Remind the customer where payment runs', body: 'Pay through the platform. Cash is only for tips.' },
    ];
    const cw = (9.1 - 0.45) / 4;
    for (let i = 0; i < items.length; i++) {
      const x = 0.45 + i * (cw + 0.15), y = 1.4;
      box(s, x, y, cw, 2.6, C.white, C.border);
      await iconDot(s, x + 0.18, y + 0.18, items[i].ico, 0.36);
      T(s, items[i].title, { x: x + 0.18, y: y + 0.66, w: cw - 0.36, h: 0.92, fontFace: HEAD, bold: true, fontSize: 12, valign: 'top' });
      T(s, items[i].body, { x: x + 0.18, y: y + 1.65, w: cw - 0.36, h: 0.9, fontSize: 9.5, color: C.soft, valign: 'top' });
    }
  }

  {
    const s = jobStatus = content(CASH, 'The job status decides the action', 'Nobody should pay twice, and nobody should be paid twice.',
      'The customer shouldn\'t pay twice (cash and the platform) and the cleaner shouldn\'t be paid twice (cash and their payout). Pending Invoice: nothing has been charged; cancel via the C CRM with the "does not want the service" reason code so it never gets charged. Invoiced: refund from the CP Dashboard, which reverses both the customer\'s charge and the cleaner\'s payout in one action. Cancelled (the customer says the cleaning happened and they paid cash): nothing was ever charged, so there is nothing to refund. Do not recreate the job: that\'s how the customer ends up paying twice. Exception: the cleaner tells us separately they were never paid; then work it as a normal cash case from the job\'s actual status. Otherwise educate on the cash policy, cancel the RC and upcoming jobs if the customer no longer wants service, and close.');
    await cards(s, 1.4, 2.0, [
      // 📹 runs link to how-to videos in Drive.
      { ico: 'FiClock', title: 'Pending Invoice', body: [
        { text: 'Fully covered? ' }, { text: 'Cancel via the C CRM: "does not want the service".', options: { breakLine: true } },
        { text: 'Partially covered? ' }, { text: '📹Invoice ', options: { hyperlink: { url: 'https://drive.google.com/file/d/1NI_7g4lgQ7n8nVf4-w2sNLksijWYnf-e/view?usp=drive_link', tooltip: 'Play video' } } }, { text: 'the remaining balance.' },
      ] },
      { ico: 'FiDollarSign', title: 'Invoiced', body: [
        { text: 'Full refund → ' }, { text: '📹 Refund from CP Dashboard', options: { hyperlink: { url: 'https://drive.google.com/file/d/1TuwGdwdOWppyezVsG-aITCSja8J-Z0aJ/view?usp=drive_link', tooltip: 'Play video' } } }, { text: '. Reverses both sides.', options: { breakLine: true } },
        { text: 'Partial refund → ' }, { text: '📹Admin refund', options: { hyperlink: { url: 'https://drive.google.com/file/d/1gfRiLUltOEjqDdeLfUkY-7V7u7b0C36n/view?usp=drive_link', tooltip: 'Play video' } } }, { text: ' + ' },
        { text: '📹CP Holdback', options: { hyperlink: { url: 'https://drive.google.com/file/d/1df803P7wbfSjWemzHFR7nFgx-AGHbZXW/view?usp=drive_link', tooltip: 'Play video' } } },
      ] },
      { ico: 'FiXCircle', title: 'Cancelled', body: 'Nothing to refund. Don\'t recreate the job.' },
    ]);
    await tip(s, 3.6, 'Partial cash payment?', 'invoice (or refund) only the difference, based on CP Pay, not the C Price.', 'FiPercent');
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
    // Under "Act on the status": jumps back to the job status slide.
    const cw = (9.1 - 0.45) / 4, bx = 0.45 + 2 * (cw + 0.15);
    s.addText('Job status actions  ›', { shape: pres.shapes.ROUNDED_RECTANGLE, x: bx + 0.18, y: 2.95, w: cw - 0.36, h: 0.32, rectRadius: 0.08,
      fill: { color: C.teal }, line: { color: C.teal }, fontFace: HEAD, bold: true, fontSize: 9.5, color: C.white, align: 'center', valign: 'middle', margin: [2, 2, 0, 0],
      hyperlink: { slide: jobStatus._slideNum, tooltip: 'Go to: The job status decides the action' } });
    await tip(s, 3.6, 'Partial payment, still Pending Invoice?', 'invoice the remaining balance, based on CP Pay.', 'FiPercent');
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
    // Branch B splits on what the CP CRM shows: one flowchart per path.
    const start = [
      ['Customer says they paid the CP cash', 'penalty', 'C-side agent: the customer contacted us first. The CP-side ticket may not be worked yet, or the CP never reported it.'],
    ];
    const paths = [
      ['Branch B: report on file', 'Report found in the CP CRM, not yet actioned: resolve it yourself, no need to wait for the CP-side queue. Amount: did the CP say how much they received in cash? If not, contact the CP and get a number (or a clear "I got paid in full") before doing anything else, same gate as Branch A. Job status: still Pending Invoice: cancel via C CRM, reason code "does not want the service"; if it was only a partial cash payment, invoice the remaining balance instead. Already Invoiced: refund via the CP Dashboard; if partial, admin refund the customer for the cash-paid portion (based on CP Pay, not C Price) and apply a CP Holdback for the same amount, so the CP isn\'t paid twice. Penalty: apply the soft ladder; this is still a CP admission, just one a C-side agent is closing out.', [
        ['Check CP CRM: report found', 'decision', 'A cash-payment report is on file but not yet actioned. Resolve it yourself: no need to wait for the CP-side queue.'],
        ['Get the cash amount', 'action', 'Did the CP say how much they received? If not, contact the CP for a number (or a clear "paid in full") before doing anything else.'],
        ['Still Pending Invoice?', 'decision', 'Fully paid in cash? Cancel via C CRM: Reason code "does not want the service".\nPartial payment? Invoice the remaining balance instead.'],
        ['Already Invoiced?', 'decision', 'Fully paid in cash? Refund via the CP Dashboard.\nPartial payment? Admin refund the cash-paid portion (based on CP Pay, not C Price) + a CP Holdback for the same amount.'],
        ['Apply the soft ladder', 'end', 'It\'s a CP admission, just one a C-side agent is closing out. Use the soft ladder on the penalty ladder slide.'],
      ]],
      ['Branch B: nothing on file', 'Nothing on file in the CP CRM: nothing is confirmed yet, so treat it like any other unconfirmed claim. Ask the customer for proof (encourage it, don\'t require it). Call the CP if it\'s within business hours (8AM-8PM their local time); no answer during business hours: send an SMS; outside business hours: email instead. Set a Timed Reminder for 48 hours. Customer provides proof, whatever the CP says: resolve per the job status (C CRM cancellation or CP Dashboard refund) and apply the hard ladder; the customer\'s documentation confirmed it, so it\'s treated as getting caught, not disclosing. No proof, but the CP admits it when contacted: resolve per the job status and apply the soft ladder. CP denies and the customer never produces proof: no refund and no penalty, since nothing is confirmed; assess for retention: high-value / retention-likely (multiple completed jobs, high LTNR): voucher covering the full job hours, Internal Reason: cash_payment; not high-value: $20-$50 in platform credits, Internal Reason: cash_payment.', [
        ['Check CP CRM: nothing on file', 'decision', 'No report from CP. Treat it like any other unconfirmed claim.'],
        ['Ask C for proof', 'action', 'Encourage proof, don\'t require it.'],
        ['Contact the CP', 'action', 'Call within business hours (8AM–8PM CP local time). No answer: SMS. Outside business hours: email.'],
        ['Set a 48-hour Timed Reminder', 'action', 'Pick the ticket back up when it triggers.'],
        ['C provides proof → hard ladder', 'action', 'Pending Invoice: cancel from CCRM (full) or invoice remaining balance (partial)\nInvoiced: Refund from CP Dashboard (full) or admin refund + CP holdback (partial)\nHard ladder penalty'],
        ['No proof from C but CP admits it → soft ladder', 'action', 'Resolve per job status above + Soft ladder penalty.\n1st instance - coach\n2nd instance - coach + final warning\n3rd instance - suspended'],
        ['CP denies, no proof → retention', 'end', 'No refund, no penalty. High value: voucher for the full job hours. Otherwise $20–$50 credits. Reason: cash_payment.'],
      ]],
    ];
    for (const [title, notes, steps] of paths) {
      const s = content(CASH, title, null,
        'TRAINER: descriptions are hidden. Click Show details (or press the right arrow): each click types out the next step\'s description. ' + notes);
      await detailsButton(s, 7.6, 0.55);
      const rows = start.concat(steps);
      // Rows grow with multi-line descriptions so nothing overlaps.
      const n = rows.length, gap = 0.1, x = 0.5, w = 3.55;
      const hs = rows.map(r => Math.max(n > 7 ? 0.34 : 0.42, (r[2].split('\n').length) * 0.13 + 0.08));
      let y = 1.2;
      for (let i = 0; i < n; i++) {
        const h = hs[i], multi = rows[i][2].includes('\n');
        await flowBox(s, x, y, w, h, rows[i][0], rows[i][1]);
        if (i < n - 1) flowLine(s, [[x + w / 2, y + h], [x + w / 2, y + h + gap]], true);
        flowLine(s, [[x + w, y + h / 2], [x + w + 0.2, y + h / 2]]);
        T(s, rows[i][2], { x: x + w + 0.3, y: y - 0.04, w: 5.45, h: h + 0.08, fontSize: multi ? 7.5 : 8.5, color: C.soft, valign: 'middle', objectName: `type${i + 1}Desc` });
        y += h + gap;
      }
    }
  }

  await knowledgeCheck(CASH, [
    ['The cleaner reported full cash payment. The job is still Pending Invoice. What do you do?', 'Cancel via the C CRM with "does not want the service". Nothing was charged, so nothing to refund.'],
    ['Same case, but the job was already invoiced.', 'Refund from the CP Dashboard. It reverses the customer\'s charge and the cleaner\'s payout.'],
    ['The job shows Cancelled, but the customer says they paid cash. Recreate the job?', 'No. Nothing was charged. Recreating it would make the customer pay twice.'],
    ['The cleaner denies it and the customer has no proof. What now?', 'No refund, no penalty. Assess retention: voucher (high value) or $20–$50 credits, reason cash_payment.'],
  ], 'Ask each question and let trainees answer before revealing. Each click reveals the next answer.');

  {
    const s = content(CASH + '  ·  Practice', 'Let\'s practice!', 'Review the live ticket and investigate.',
      'TRAINER: provide a live ticket for trainees to review. Give them time to investigate (job history, CTJ, C/CP messages, Premium charge), then discuss. Ask: Did the CP already report the cash payment (check the CP CRM)? How much was paid in cash, full or partial? What is the job status? What actions will you take (cancel via C CRM, invoice the balance, refund via CP Dashboard, or admin refund + CP Holdback), and which ladder applies? How will you respond to the customer?');
    await card(s, 0.45, 1.45, 3.9, 3.0, { ico: 'FiInbox', title: 'Live ticket', body: 'Your trainer will provide a live ticket. Review it and investigate before you decide anything.' });
    box(s, 0.65, 3.85, 1.35, 0.3, C.goldSoft);
    T(s, 'Trainer provides', { x: 0.65, y: 3.85, w: 1.35, h: 0.3, fontSize: 8, bold: true, color: C.gold, align: 'center', valign: 'middle' });
    await card(s, 4.5, 1.45, 5.05, 1.42, { n: 1, title: 'What will be your actions?', body: 'Which issue is it, and what will you do on the account?' });
    await card(s, 4.5, 3.03, 5.05, 1.42, { n: 2, title: 'How will you respond to the customer?', body: 'Write the reply you would send.' });
  }

  await wrapUp(CASH, [
    ['Status decides the action', 'Pending Invoice: cancel. Invoiced: refund. Cancelled: don\'t recreate.'],
    ['Get the amount first', 'Full vs partial is based on CP Pay, not the C Price.'],
    ['Never penalize the customer', 'Coach the cleaner. C-side never suspends or adds DNR.'],
  ], 'TRAINER: recap the three takeaways, then open the floor for questions on Cash Payment.');

  // ================= 5. CLEANER DIDN'T SHOW =================
  const NS = 'Cleaner Didn\'t Show';
  await topic('Cleaner Didn\'t Show', 'The CP claimed the job, then never communicated or never showed up.',
    ['When it\'s valid to report a no-show', 'Reschedule intent vs no-show', 'Handling by job status', 'Resolution and penalty'],
    'ill7.png',
    'Category: Service Issues > Reliability. A no-show is a serious reliability failure, but only report it when you\'re actually confident it happened. Reporting a no-show that didn\'t occur penalizes a CP unfairly, based on nothing more than a customer\'s assumption. The bar is certainty, not just a complaint.');

  {
    const s = content(NS, 'Certainty, not a complaint', 'Only report a no-show via the Job Admin page if at least one of these is true.',
      'Only report a CP as a no-show via the Job Admin page if at least one is true: the customer says the CP was a no-show AND you have sufficient proof; you\'re 100% sure from customer/CP messages that the CP genuinely didn\'t show; or the customer left a review explicitly saying the CP was a no-show. Signals the system may pick up automatically (they don\'t confirm a no-show on their own, but are worth checking): the customer asked for the CP\'s ETA, or the CP sent an SMS with cancellation/reschedule intent. Reporting via the Job Admin page only works while the CP is still attached to the job (Claimed, Pending Invoice). Once the job reverts to Submitted or moves to Cancelled, there\'s no CP-job link to report against. If the CP was a no-show and the job is Invoiced, follow the False Invoice process.');
    await cards(s, 1.4, 1.75, [
      { ico: 'FiMessageCircle', title: 'Reported + proven', body: 'The customer says it, and you have sufficient proof.' },
      { ico: 'FiCheckCircle', title: '100% sure', body: 'Customer/CP messages show the CP genuinely didn\'t show.' },
      { ico: 'FiStar', title: 'In a review', body: 'The customer\'s review explicitly says the CP was a no-show.' },
    ]);
    await tip(s, 3.3, 'Signals to check:', 'the customer asked for the CP\'s ETA, or the CP texted cancel/reschedule intent.', 'FiSearch');
    await tip(s, 3.85, 'No show via Admin Page:', 'Claimed or Pending Invoice. Already Invoiced? Follow the False Invoice process.', 'FiAlertCircle');
  }

  {
    const s = content(NS, 'Reschedule intent or no-show?', 'If C reports a no-show, check the C/CP messages first: when did the CP say they wanted to reschedule?',
      'If the CP showed an intent to reschedule before the scheduled start time, don\'t report it as a No Show in Job Admin. Cancel the appointment through C-CRM using the CP No Show/Reschedule code. Coaching for a no-show would normally encourage the CP to reschedule rather than cancel, and the CP already tried to, so the penalty can be waived. Before start: cancel via C-CRM (no penalty + coaching). After start: report as no-show via Job Admin (penalty + coaching). At the start time: assess whether the CP should be penalized based on the reason for rescheduling.');
    await cards(s, 1.55, 2.05, [
      { ico: 'FiRewind', title: 'Before the start time', body: 'Cancel via C-CRM with the CP No Show/Reschedule code. No penalty + coaching.' },
      { ico: 'FiClock', title: 'At the start time', body: 'Assess whether a penalty fits, based on the reason for rescheduling.' },
      { ico: 'FiFastForward', title: 'After the start time', body: 'Report as a no-show via Job Admin. Penalty + coaching.' },
    ]);
  }

  {
    const s = content(NS, 'Handling by job status', 'Then, for every status: check the resolution table, fix the unreliability issue, and send comms.',
      'Applies to all statuses: before choosing a resolution, check the Quick Reference table (next slide) for credits and free-month eligibility. Weigh the customer\'s emotional state, whether they\'re asking for a refund or hinting at cancelling, and whether they already have recent credits/vouchers. If the customer doesn\'t seem upset and isn\'t asking for a refund, consider non-credit options first (Priority Booking, blocking the C/CP pairing). Penalize per CP penalty guidance. Claimed or Pending Invoice: if the CP should be penalized, report the no-show (Job Admin Page > No Show); after reporting, the job resolves to Cancelled or Submitted, so follow that column. If no penalty applies (e.g. reschedule intent before the start time), cancel the job from C-CRM with the CP No Show/Reschedule code.');
    // Flowchart: one no-show, three job statuses, one shared set of next steps.
    await flowBox(s, 5.0 - 1.4, 1.45, 2.8, 0.42, 'CLEANER NO SHOW', 'penalty');
    const cols = [
      [1.75, 'SUBMITTED', 'Advise that we\'re matching them with a new cleaner.'],
      [5.0, 'CLAIMED / PENDING INVOICE', 'Penalize CP? Report the CP as a no-show: Job Admin Page › No Show.\nNo penalty? Cancel the job from C-CRM with the CP No Show/Reschedule code.'],
      [8.25, 'CANCELLED', 'Offer Priority Booking.'],
    ];
    const cw = 3.0;
    flowLine(s, [[5.0, 1.87], [5.0, 2.07]]);
    flowLine(s, [[1.75, 2.07], [8.25, 2.07]]);
    for (const [cx, head, body] of cols) {
      flowLine(s, [[cx, 2.07], [cx, 2.25]], true);
      await flowBox(s, cx - cw / 2, 2.25, cw, 0.38, head, 'action');
      box(s, cx - cw / 2, 2.63, cw, 0.92, C.tealSoft, C.teal);
      T(s, body, { x: cx - cw / 2 + 0.1, y: 2.63, w: cw - 0.2, h: 0.92, fontSize: body.length > 80 ? 8.5 : 9.5, color: C.ink, align: 'center', valign: 'middle' });
      flowLine(s, [[cx, 3.55], [cx, 3.72]]);
    }
    flowLine(s, [[1.75, 3.72], [8.25, 3.72]]);
    flowLine(s, [[5.0, 3.72], [5.0, 3.88]], true);
    box(s, 1.6, 3.88, 6.8, 1.15, C.goldSoft, C.gold);
    T(s, [
      { text: 'See the quick reference for credits / free months', options: { bullet: true, breakLine: true } },
      { text: 'Address the unreliability issue: ban the C/CP pairing, penalize as necessary, offer to book a Priority Booking', options: { bullet: true, breakLine: true } },
      { text: 'Send comms to C and CP', options: { bullet: true } },
    ], { x: 1.8, y: 3.9, w: 6.45, h: 1.1, fontSize: 9.5, color: C.ink, paraSpaceAfter: 3, valign: 'middle' });
  }

  {
    const s = content(NS, 'Resolution for CP issues', 'Credits and free months. The same table applies to cleaner cancellations.',
      'Free months reminder. With service cancellation intent: agents may issue free months for unused MFs (max 3, capped at unused paid MFs or months requested, whichever is lowest). Offering is discretionary unless C names a pain point tied to the unused MFs ("I\'ve been paying but haven\'t gotten a cleaning", cost concerns); then it\'s expected and flagged in QA if skipped. With no cancellation intent: a service issue still gets the standard response (credits + help rebooking); upgrading to a free month for unused MFs only if C names them as a pain point. Example: C mentions a no-show, has 3 unused MFs, doesn\'t raise them: credits + priority rebooking. Example: C adds "I\'ve got months just sitting there unused": issue the free month instead of (or alongside) credits.');
    table(s, ['Situation', 'Credits', 'Free month(s)'], [
      ['Non-FC member', '$10–20 (encourage booking)', '—'],
      ['FCF, no completed job', 'None (free/discounted cleaning already available)', 'Unused MFs: free months for those, max 3. Otherwise 1 if next MF is within 2 weeks'],
      ['FCF, with completed job', '$10–20 (encourage booking)', 'Unused MFs: free months for those, max 3. Otherwise 1 if next MF is within 2 weeks'],
      ['Unused DHJ', '$10–20 (encourage booking)', 'None: FC hasn\'t started yet'],
      ['DHJ, FC started', '$10–20 (encourage booking)', 'Unused MFs: free months for those, max 3. Otherwise 1 if next MF is within 1 week'],
      ['Late (all cases above)', '$10–25, for all', 'Unused MFs: never more than asked for, max 3. Otherwise FCF: 1 if next MF within 2 weeks; DHJ: within 1 week'],
    ], { y: 1.4, colW: [2.0, 2.6, 4.5], fontSize: 8.5, rowH: 0.42 });
    await tip(s, 4.45, 'Free months for unused MFs:', 'expected only when C names them as a pain point. Otherwise optional.', 'FiCalendar');
  }

  {
    const s = content(NS, 'Changing CP status', 'Reporting the no-show (Job Admin › No Show) is what triggers the penalty.',
      'The system usually sets norequests or suspended on its own. If the system penalized the CP, leave it: coach the CP and tell them about the penalty and how to appeal. If it didn\'t, lean towards coaching; set norequests yourself only if the CP already has a couple of no-shows or cancellations. C-side agents never suspend or add DNR for a no-show (see CP Penalty: C-Side Handling). CP-facing reason code: cpq_noshow Your customer <C\'S NAME> reported you didn\'t show up to their <DAY, DATE> job. Submit reactivation appeal through dashboard. Macros: No show (customer-facing); No Show/Cancellation (No Request) (CP-facing, <10% threshold).');
    const w = (9.1 - 0.15) / 2;
    await card(s, 0.45, 1.4, w, 1.6, { ico: 'FiCpu', title: 'System penalized the CP', body: 'Leave it. Coach the CP, explain the penalty and how to appeal.' });
    await card(s, 0.45 + w + 0.15, 1.4, w, 1.6, { ico: 'FiMessageSquare', title: 'System didn\'t', body: 'Lean towards coaching. norequests only if they already have a couple of no-shows or cancellations.' });
    box(s, 0.45, 3.15, 9.1, 0.62, C.white, C.border);
    T(s, [{ text: 'CP-facing reason  ', options: { bold: true, color: C.teal, fontFace: HEAD } }, { text: 'cpq_noshow Your customer <C\'S NAME> reported you didn\'t show up to their <DAY, DATE> job. Submit reactivation appeal through dashboard.' }], { x: 0.65, y: 3.15, w: 8.8, h: 0.62, fontSize: 9.5, valign: 'middle' });
    await tip(s, 3.95, 'Never:', 'C-side agents never suspend or add DNR for a no-show.', 'FiLock');
  }

  await knowledgeCheck(NS, [
    ['C says the CP didn\'t show. Messages show the CP asked to move it to tomorrow, 2 hours before the start. Report a no-show?', 'No. Reschedule intent before the start time: cancel via C-CRM with the CP No Show/Reschedule code. No penalty, coach the CP.'],
    ['The CP never came, but the job is already Invoiced. What now?', 'You can\'t report a no-show here. Follow the False Invoice process.'],
    ['The job is Claimed and you\'re sure the CP didn\'t show. First step?', 'Job Admin Page › No Show. Then follow the status it resolves to: Submitted or Cancelled.'],
    ['The system didn\'t penalize the CP. It\'s their first no-show. Set norequests?', 'Lean towards coaching. norequests only if they already have a couple of no-shows or cancellations. Never suspend or DNR.'],
  ], 'Ask each question and let trainees answer before revealing. Each click reveals the next answer.');

  await wrapUp(NS, [
    ['Be certain first', 'Proof, messages or a review. Not just an assumption.'],
    ['Reschedule intent matters', 'Before the start: cancel via C-CRM, no penalty. After: report the no-show.'],
    ['Status decides the steps', 'Submitted, Claimed/Pending Invoice or Cancelled. Credits per the table.'],
  ], 'TRAINER: recap the three takeaways, then open the floor for questions on Cleaner Didn\'t Show before moving on.');

  // ================= 6. CLEANER CANCELLATION =================
  const CC = 'Cleaner Cancellation';
  await topic('Cleaner Cancellation', 'The CP cancelled a claimed job, from their dashboard or by texting the customer.',
    ['How the system treats a cancellation', 'What happens to the job', 'The handling process', 'When the customer can reschedule'],
    'ill30.png',
    'Category: Service Issues > Reliability. Cleaner Cancellation is when a CP cancels a claimed job via their dashboard, or sends the customer a cancellation-intent SMS (which the system also detects). A CP with cancellation intent is supposed to try rescheduling with the customer first; cancelling outright is the fallback if the customer declines a new time. Related: Cleaner Didn\'t Show is the distinct scenario where the CP silently doesn\'t show.');

  {
    const s = content(CC, 'Not every cancellation counts the same', 'How the CP cancelled decides whether their Keep Rate is affected.',
      'Cancelling through the dashboard normally affects a CP\'s Keep Rate. Cancelling by reporting an emergency, a safety concern, or "client wants to cancel" through the CP App does not. Treating every cancellation as equally penalizable unfairly hits CPs who were upfront about a legitimate reason. Unlike a no-show, you don\'t report anything manually: the job history logs the cancellation, so the Keep Rate impact lands whatever job status results. The system only penalizes a CP cancellation within 7 days of a 1st C/CP match, or within 2 days of a repeat match. Cancelling via CRM (you, on the CP\'s behalf) never penalizes the CP.');
    await cards(s, 1.4, 1.85, [
      { ico: 'FiMonitor', title: 'CP Dashboard', body: 'Counts toward Keep Rate, within 7 days (1st match) or 2 days (repeat match).' },
      { ico: 'FiSmartphone', title: 'CP App Report Issue', body: 'Emergency, safety concern or "client wants to cancel": no Keep Rate impact.' },
      { ico: 'FiShield', title: 'Cancelled via CRM', body: 'You cancel on the CP\'s behalf: never penalizes the CP.' },
    ]);
    await tip(s, 3.45, 'No manual report needed:', 'the job history logs the cancellation, so the Keep Rate impact lands either way.', 'FiList');
  }

  {
    const s = content(CC, 'What happens to the job', 'It depends on timing and alternate start times.',
      'If the job start time is already in the past and there are no other alternate start times: the job becomes Cancelled. If the start time is still in the future and there\'s still time to match a new CP: the job becomes Submitted again. If the customer has other available alternate start times: the job is rescheduled to one of them and returns to Submitted. The same rules apply if the CP cancelled via the CP App\'s Report Issue and chose "I have an emergency" or "My client wants to cancel."');
    table(s, ['Situation', 'Resulting job status'], [
      ['Start time already passed, no other alternate start times', 'Cancelled'],
      ['Start time still in the future, time to match a new CP', 'Submitted'],
      ['Customer has other alternate start times', 'Rescheduled to one of them: Submitted'],
    ], { y: 1.45, colW: [5.6, 3.5], fontSize: 10.5, rowH: 0.55 });
    await tip(s, 3.85, 'Same rules:', 'when the CP cancels via CP App Report Issue ("emergency" or "client wants to cancel").', 'FiSmartphone');
  }

  {
    const s = content(CC, 'The handling process', 'Then, for every status: check the resolution table, fix the unreliability issue, and send comms.',
      '1) Attempt to rematch first, not cancel outright: if there\'s still a chance to match another CP, take that route so the customer still gets their cleaning. 2) Check the Quick Reference table for credits and free months, with the same judgment as any reliability issue: emotional state, cancellation/refund intent, existing credits or vouchers. 3) Offer a Priority Booking if the customer doesn\'t get the outcome they wanted: no reschedule option, an auto-rescheduled time they didn\'t choose, or the job goes to Cancelled. 4) Block the C/CP pairing. 5) Send comms to customer and CP: coach the CP on reliability, penalize as necessary. 6) CP status: cancelling from the CP Dashboard is how the system penalizes the CP; it usually sets norequests or suspended itself. If it did, leave it, coach the CP and explain the penalty. If it didn\'t, lean towards coaching; set norequests only if the CP already has a couple of cancellations or no-shows. C-side agents never suspend or add DNR. Reason code: cpq_cancel Your customer <C\'S NAME> reported you cancelled their <DAY, DATE> job. Submit reactivation appeal through dashboard. Macros: CP Cancel > Job Cancelled (customer-facing); No Show/Cancellation (No Request) (CP-facing); No Show/Cancellation (Permanent Deactivation) (CP-facing).');
    // Flowchart: one cancellation, three job statuses, one shared set of next steps.
    await flowBox(s, 5.0 - 1.4, 1.45, 2.8, 0.42, 'CLEANER CANCELLATION', 'penalty');
    const cols = [
      [1.75, 'SUBMITTED', 'Advise that we\'re matching them with a new cleaner.'],
      [5.0, 'CLAIMED / PENDING INVOICE', 'Cancel via the CP Dashboard or rematch the CP.'],
      [8.25, 'CANCELLED', 'Offer Priority Booking.'],
    ];
    const cw = 3.0;
    flowLine(s, [[5.0, 1.87], [5.0, 2.07]]);
    flowLine(s, [[1.75, 2.07], [8.25, 2.07]]);
    for (const [cx, head, body] of cols) {
      flowLine(s, [[cx, 2.07], [cx, 2.25]], true);
      await flowBox(s, cx - cw / 2, 2.25, cw, 0.38, head, 'action');
      box(s, cx - cw / 2, 2.63, cw, 0.92, C.tealSoft, C.teal);
      T(s, body, { x: cx - cw / 2 + 0.1, y: 2.63, w: cw - 0.2, h: 0.92, fontSize: 9.5, color: C.ink, align: 'center', valign: 'middle' });
      flowLine(s, [[cx, 3.55], [cx, 3.72]]);
    }
    flowLine(s, [[1.75, 3.72], [8.25, 3.72]]);
    flowLine(s, [[5.0, 3.72], [5.0, 3.88]], true);
    box(s, 1.6, 3.88, 6.8, 1.15, C.goldSoft, C.gold);
    T(s, [
      { text: 'See the quick reference for credits / free months', options: { bullet: true, breakLine: true } },
      { text: 'Address the unreliability issue: ban the C/CP pairing, penalize as necessary, offer Priority Booking', options: { bullet: true, breakLine: true } },
      { text: 'Send comms to C and CP', options: { bullet: true } },
    ], { x: 1.8, y: 3.9, w: 6.45, h: 1.1, fontSize: 9.5, color: C.ink, paraSpaceAfter: 3, valign: 'middle' });
  }

  {
    const s = content(CC, 'When the customer can reschedule', 'The CP\'s Keep Rate is impacted either way: the job history is logged.',
      'If the customer has the option to reschedule: customer reschedules: Submitted; customer doesn\'t reschedule: Cancelled. Either way the CP\'s Keep Rate is still impacted since the job history is logged, which can still lead to norequests or suspended. When the customer doesn\'t have the option (e.g. alternate times were already added to their appointment), the job may auto-reschedule to the latest alternate time. If no alternate times exist and rescheduling isn\'t offered, the job changes to Cancelled, with the same Keep Rate impact. Free months reminder: not mandatory and won\'t fail QA either way; hold off unless there\'s cancellation intent or C brings up the unused months.');
    const w = (9.1 - 0.15) / 2;
    eyebrow(s, 'Customer can reschedule', 0.48, 1.45, 4, C.soft);
    table(s, ['Customer action', 'Job status'], [['Reschedules', 'Submitted'], ['Doesn\'t reschedule', 'Cancelled']], { y: 1.72, colW: [2.6, 1.87], fontSize: 10.5, rowH: 0.45 });
    await card(s, 0.45 + w + 0.15, 1.45, w, 1.75, { ico: 'FiCalendar', title: 'No option to reschedule', body: 'Alternate times already added? It may auto-reschedule to the latest one. None? The job is Cancelled.' });
    await tip(s, 3.45, 'Free months:', 'not mandatory. Hold off unless there\'s cancellation intent or C raises unused months.', 'FiInfo');
  }

  await knowledgeCheck(CC, [
    ['The CP cancelled from the CP App with "My client wants to cancel." Is their Keep Rate affected?', 'No. Emergency, safety concern or "client wants to cancel" via the CP App doesn\'t count.'],
    ['You cancel the job in the CRM on the CP\'s behalf. Is the CP penalized?', 'No. Cancelling via CRM never penalizes the CP.'],
    ['The CP cancels; the start time is still days away. What happens, and what\'s your first move?', 'The job goes back to Submitted. Try to rematch first so the customer still gets their cleaning.'],
    ['The job auto-rescheduled to an alternate time the customer didn\'t pick. What do you offer?', 'A Priority Booking. Block the pairing, and check the table for credits.'],
  ], 'Ask each question and let trainees answer before revealing. Each click reveals the next answer.');

  await wrapUp(CC, [
    ['How they cancelled matters', 'Dashboard counts. CP App emergency or client-cancel doesn\'t. CRM never does.'],
    ['Rematch first', 'Then Priority Booking if the outcome isn\'t what the customer wanted.'],
    ['Coach before you pause', 'System penalty? Leave it. Never suspend or DNR.'],
  ], 'TRAINER: recap the three takeaways, then open the floor for questions on Cleaner Cancellation before moving on.');

  // ================= 7. UNAUTHORIZED RESCHEDULE =================
  const UR = 'Unauthorized Reschedule';
  await topic('Unauthorized Reschedule', 'The CP rescheduled the appointment without the customer\'s consent.',
    ['Two checks before you act', 'Still assigned: the start-time table', 'No longer assigned', 'Penalty and reason code'],
    'ill22.png',
    'Category: Service Issues > Reliability. Different from the routine, agreed-upon reschedule covered in Rescheduling (Bookings): here the customer never agreed to the new time.');

  {
    const s = content(UR, 'Handling guide', null,
      'Step 1: is the job still assigned to the CP who rescheduled it? No, no longer assigned: review the CP\'s account for prior instances and update their CP Step if warranted; send the CP coaching comms using the applicable Comms Kit macro and ban (block) the C/CP pairing; offer the customer a Priority Booking and address any other pain points. Yes, still assigned: add the cp_unauthorized_reschedule flag to the CP\'s account, update their CP Step if warranted, coach and ban the CP, then check how far away the original start time is. >48 hours away: reschedule the job back to the original start time via the Customer Dashboard. 48 hours or less: check the customer\'s Priority Markup cohort (C CRM > View). If the cohort is 30_to_50_pct or 50_to_100_pct, cancel the job via CRM using CP no-show/reschedule as the reason code, then rebook via CRM. Add an internal note: cp_unauthorized_reschedule cancelled J <JOB ID> to avoid priority_fee and recreated J <JOB ID>. Tell the customer their appointment was rebooked to their original preferred schedule and offer to add Alternate Start Times. If no priority fee will apply, reschedule via the Customer Dashboard instead. Already in the past: follow AG: Pending Invoice\'s completion-verification steps, whether the job is Pending Invoice or Claimed. Job already cancelled/invoiced: offer a Priority Booking and address any other pain points. The flag can still be added even if the CP is no longer assigned: it documents the CP\'s own conduct.');
    // Spine: issue > decision > Yes action, with the No branch off to the right.
    const cx = 4.0;
    await flowBox(s, cx - 1.4, 1.1, 2.8, 0.34, 'UNAUTHORIZED RESCHEDULE', 'penalty');
    flowLine(s, [[cx, 1.44], [cx, 1.6]], true);
    await flowBox(s, cx - 1.75, 1.6, 3.5, 0.5, 'Is the job still assigned to the CP\nwho rescheduled it?', 'decision');
    flowLine(s, [[cx + 1.75, 1.85], [6.45, 1.85]], true);
    T(s, 'NO', { x: cx + 1.85, y: 1.62, w: 0.5, h: 0.22, fontSize: 8.5, bold: true, color: C.soft });
    box(s, 6.45, 1.15, 3.1, 1.4, C.white, C.border);
    T(s, [
      { text: 'Update CP Step if warranted', options: { bullet: true, breakLine: true } },
      { text: 'Coach and ban the CP', options: { bullet: true, breakLine: true } },
      { text: 'Offer the C a Priority Booking and address any other pain points', options: { bullet: true } },
    ], { x: 6.55, y: 1.15, w: 2.95, h: 1.4, fontSize: 9, color: C.ink, paraSpaceAfter: 3, valign: 'middle' });
    flowLine(s, [[cx, 2.1], [cx, 2.3]], true);
    T(s, 'YES', { x: cx + 0.08, y: 2.08, w: 0.5, h: 0.22, fontSize: 8.5, bold: true, color: C.soft });
    box(s, cx - 2.1, 2.3, 4.2, 0.55, C.tealSoft, C.teal);
    T(s, [
      { text: 'Add the ' }, { text: 'cp_unauthorized_reschedule', options: { italic: true } },
      { text: ' flag, update their CP Step if warranted, coach and ban the CP.' },
    ], { x: cx - 2.0, y: 2.3, w: 4.0, h: 0.55, fontSize: 9, color: C.ink, align: 'center', valign: 'middle' });
    // Four outcomes by where the original start time sits now.
    const cols = [
      ['Original start time is\n>48 hours away', 'Reschedule back to the original job start time.'],
      ['Original start time is\n≤48 hours away', 'Cancel and rebook the job via the C CRM, or reschedule via the C Dashboard if no priority fee will apply.'],
      ['Original start time is\nin the past', 'Follow AG: Pending Invoice completion verification steps.'],
      ['Cancelled / Invoiced', 'Offer a Priority Booking and address other concerns.'],
    ];
    const cw = 2.15, gap = 0.15, x0 = 0.45, xs = cols.map((_, i) => x0 + i * (cw + gap) + cw / 2);
    flowLine(s, [[cx, 2.85], [cx, 3.0]]);
    flowLine(s, [[xs[0], 3.0], [xs[3], 3.0]]);
    for (let i = 0; i < cols.length; i++) {
      flowLine(s, [[xs[i], 3.0], [xs[i], 3.15]], true);
      await flowBox(s, xs[i] - cw / 2, 3.15, cw, 0.5, cols[i][0], 'action');
      box(s, xs[i] - cw / 2, 3.65, cw, 1.2, C.tealSoft, C.teal);
      T(s, cols[i][1], { x: xs[i] - cw / 2 + 0.08, y: 3.65, w: cw - 0.16, h: 1.2, fontSize: 9, color: C.ink, align: 'center', valign: 'middle' });
    }
  }

  {
    const s = content(UR, 'Penalty and reason code', 'A first unauthorized reschedule is coaching only.',
      '1st instance: warning macro, no status change. 2nd+ instance: norequests plus the warning macro. Before pausing a CP, check their messages and flags for an earlier warning. C-side agents never suspend or add DNR; repeat offenders are the CP-side team\'s call. CP-facing reason code: cpq_reschedule Your customer <C\'S NAME> reported you rescheduled their <DAY, DATE> job without permission. Submit reactivation appeal through dashboard. Macros: Unapproved reschedule (sample 1) (customer-facing, includes a $20 credit); Unapproved reschedule (sample 2) (customer-facing, no credit: apology + coordination note); Unauthorized Reschedule (Warning) (CP-facing, 1st instance).');
    const w = (9.1 - 0.15) / 2;
    await card(s, 0.45, 1.4, w, 1.55, { n: 1, title: '1st instance', body: 'Warning macro. No status change.' });
    await card(s, 0.45 + w + 0.15, 1.4, w, 1.55, { n: 2, title: '2nd+ instance', body: 'norequests plus the warning macro. Check messages and flags for an earlier warning first.' });
    box(s, 0.45, 3.1, 9.1, 0.62, C.white, C.border);
    T(s, [{ text: 'CP-facing reason  ', options: { bold: true, color: C.teal, fontFace: HEAD } }, { text: 'cpq_reschedule Your customer <C\'S NAME> reported you rescheduled their <DAY, DATE> job without permission. Submit reactivation appeal through dashboard.' }], { x: 0.65, y: 3.1, w: 8.8, h: 0.62, fontSize: 9.5, valign: 'middle' });
    await tip(s, 3.9, 'Never:', 'C-side agents never suspend or add DNR. Repeat offenders are the CP-side team\'s call.', 'FiLock');
  }

  await knowledgeCheck(UR, [
    ['What\'s the first thing you check?', 'Whether the job is still assigned to the CP who rescheduled it.'],
    ['Still assigned. The original start time is 4 days away. What do you do?', 'Add the cp_unauthorized_reschedule flag, then reschedule back to the original time via the Customer Dashboard.'],
    ['Still assigned, original start in 30 hours, cohort 50_to_100_pct. What do you do?', 'Cancel via CRM (CP no-show/reschedule), rebook via CRM, add the internal note, tell the customer, offer Alternate Start Times.'],
    ['It\'s the CP\'s first unauthorized reschedule. Change their status?', 'No. Warning macro only. 2nd+ instance: norequests plus the warning.'],
  ], 'Ask each question and let trainees answer before revealing. Each click reveals the next answer.');

  // Practice slide the trainer added in Google (a copy of the Cash Payment one).
  {
    const s = content(CASH + '  ·  Practice', 'Let\'s practice!', 'Review the live ticket and investigate.',
      'TRAINER: provide a live ticket for trainees to review. Give them time to investigate (job history, CTJ, C/CP messages, Premium charge), then discuss. Ask: Did the CP already report the cash payment (check the CP CRM)? How much was paid in cash, full or partial? What is the job status? What actions will you take (cancel via C CRM, invoice the balance, refund via CP Dashboard, or admin refund + CP Holdback), and which ladder applies? How will you respond to the customer?');
    await card(s, 0.45, 1.45, 3.9, 3.0, { ico: 'FiInbox', title: 'Live ticket', body: 'Your trainer will provide a live ticket. Review it and investigate before you decide anything.' });
    box(s, 0.65, 3.85, 1.35, 0.3, C.goldSoft);
    T(s, 'Trainer provides', { x: 0.65, y: 3.85, w: 1.35, h: 0.3, fontSize: 8, bold: true, color: C.gold, align: 'center', valign: 'middle' });
    await card(s, 4.5, 1.45, 5.05, 1.42, { n: 1, title: 'What will be your actions?', body: 'Which issue is it, and what will you do on the account?' });
    await card(s, 4.5, 3.03, 5.05, 1.42, { n: 2, title: 'How will you respond to the customer?', body: 'Write the reply you would send.' });
  }

  await wrapUp(UR, [
    ['Check before you fix', 'Still assigned? How far is the original start?'],
    ['Avoid a new Priority Fee', '48 hours or less: check the cohort before rebooking.'],
    ['Coach first', '1st: warning only. 2nd+: norequests + warning.'],
  ], 'TRAINER: recap the three takeaways, then open the floor for questions on Unauthorized Reschedule.');

  // ================= 8. TRUST AND SAFETY =================
  const TS = 'Trust and Safety';
  // Bulleted list inside a white panel with a small heading.
  const listPanel = (s, x, y, w, h, head, items, fontSize = 9.5) => {
    box(s, x, y, w, h, C.white, C.border);
    eyebrow(s, head, x + 0.2, y + 0.16, w - 0.4);
    T(s, bullets(items), { x: x + 0.2, y: y + 0.45, w: w - 0.4, h: h - 0.55, fontSize, color: C.ink, paraSpaceAfter: 3 });
  };
  // Horizontal flow: boxes joined by arrows, with a description under each.
  const hFlow = async (s, y, items) => {
    const n = items.length, gap = 0.3, w = (9.1 - gap * (n - 1)) / n;
    for (let i = 0; i < n; i++) {
      const x = 0.45 + i * (w + gap), [label, kind, desc] = items[i];
      await flowBox(s, x, y, w, 0.62, label, kind);
      if (i < n - 1) flowLine(s, [[x + w, y + 0.31], [x + w + gap, y + 0.31]], true);
      T(s, desc, { x, y: y + 0.72, w, h: 1.3, fontSize: 9, color: C.soft, valign: 'top', align: 'center' });
    }
  };

  await topic('Trust and Safety', 'Reports that put a person\'s safety, property or wellbeing at risk.',
    ['What T&S covers', 'What always escalates', 'When the customer must ask', 'Who decides refunds, ETF and retention'],
    'ill11.png',
    'Category: Specialized Escalation Teams. Most Service Issues are things Care can fully own and close out: refund it, re-clean it, coach the CP, done. Trust & Safety cases are different: they involve real risk to a person\'s physical safety, property, or wellbeing, not just a bad service experience. Getting one wrong (moving too fast, resolving it like a routine complaint, or missing that it qualifies) creates a liability a refund can\'t fix.');

  {
    const s = content(TS, 'What is Trust and Safety?', 'Any report of behavior that could compromise one of these three.',
      'T&S covers any situation where a customer or cleaner reports behavior that could compromise personal safety, property security, or professional conduct. If a report suggests any risk (physical, emotional, or property-related), the case gets escalated. Only the T&S team is authorized to handle these situations directly. Care\'s responsibility is to recognize the issue and escalate it, not investigate or resolve it. T&S is a separate, specially-trained track because these situations need a level of investigation and judgment a standard support resolution isn\'t built for. That\'s also why T&S, not Care, gets final say on refunds, ETF and retention once a case is theirs. If you\'re unsure whether a case meets the bar but it feels risky, you can still raise it through the T&S channel: you don\'t need certainty to escalate.');
    await cards(s, 1.45, 1.7, [
      { ico: 'FiShield', title: 'Personal safety', body: 'Physical or emotional risk to the customer or the cleaner.' },
      { ico: 'FiHome', title: 'Property security', body: 'Theft, damage, or a home left unsecured.' },
      { ico: 'FiUserCheck', title: 'Professional conduct', body: 'Harassment, discrimination, unwanted contact, substances.' },
    ]);
    await tip(s, 3.4, 'Your role:', 'recognize the issue and escalate it. Don\'t investigate it or resolve it.', 'FiFlag');
    await tip(s, 3.95, 'Unsure?', 'if it feels risky, escalate. You don\'t need certainty.', 'FiHelpCircle');
  }

  {
    const s = content(TS, 'The handling process', 'Escalate as soon as a T&S indicator is confirmed.',
      'Escalate immediately once a T&S indicator is confirmed. Don\'t attempt retention: T&S leads the investigation and resolution. Don\'t issue refunds or contact the cleaner directly: T&S evaluates liability, coordinates the investigation and determines the resolution (including a possible full ETF waiver). Agents must send an email to the customer when escalating a ticket to T&S. Macro: Escalate report to T&S. Related articles: ETF Waivers; Service Recovery Overview.');
    await hFlow(s, 1.6, [
      ['T&S indicator confirmed', 'penalty', 'Matches an always-escalate item, or the customer asked (conditional cases).'],
      ['Escalate to T&S', 'action', 'Right away. Macro: Escalate report to T&S.'],
      ['Email the customer', 'action', 'Required the first time you escalate.'],
      ['Hands off', 'end', 'No refund, no contact with the cleaner, no retention attempt.'],
    ]);
    await tip(s, 3.3, 'Why:', 'T&S evaluates liability and may waive the full ETF. Acting first can undercut that.', 'FiInfo');
  }

  {
    const s = content(TS, 'Who decides', 'Once a case is T&S, they have final say.',
      'All T&S concerns (once escalation applies) must be routed to the T&S team. Care agents must not make refund or compensation decisions for T&S cases, waive or reduce ETF unless directed by T&S, or attempt retention or offer membership-related incentives. T&S independently investigates and owns the final resolution: refund handling is determined solely by T&S after investigation; T&S may waive up to the full ETF following review; no retention attempt by Care, defer to the T&S outcome. If T&S declines the case ("this is not a T&S concern"), the ticket returns to the original agent and is handled as BAU. The T&S constraints lift with it: run the normal playbook for the underlying issue and apply the ordinary ETF and retention rules.');
    await twoCol(s,
      { ico: 'FiSlash', title: 'Care must not', items: ['Decide refunds or compensation', 'Waive or reduce the ETF (unless T&S directs)', 'Attempt retention or offer membership incentives', 'Contact the cleaner directly'] },
      { ico: 'FiShield', title: 'T&S owns', items: ['Refund scope, after investigation', 'ETF: may waive up to the full amount', 'Retention and membership outcome'] },
      1.45, 2.35);
    await tip(s, 4.0, 'T&S declines?', 'it\'s yours again as BAU: normal playbook, ordinary ETF and retention rules.', 'FiCornerDownLeft');
  }

  {
    const s = content(TS, 'Safety concerns', 'Always escalate.',
      'Every item on this slide is an always-escalate item. Law enforcement requests also go to T&S: never share customer or cleaner information yourself.');
    const items = [
      ['FiAlertOctagon', 'On-platform death'],
      ['FiAlertTriangle', 'Physical violence or threats of violence'],
      ['FiUserX', 'Unwanted physical contact'],
      ['FiEyeOff', 'Any sexually related complaint'],
      ['FiUsers', 'Reports of human trafficking'],
      ['FiEye', 'Reports of stalking'],
      ['FiFileText', 'Law enforcement requests'],
    ];
    const w = (9.1 - 0.15) / 2, rowH = 0.5;
    for (let i = 0; i < items.length; i++) {
      const x = 0.45 + (i % 2) * (w + 0.15), y = 1.45 + Math.floor(i / 2) * (rowH + 0.1);
      box(s, x, y, w, rowH, C.white, C.border);
      await iconDot(s, x + 0.12, y + (rowH - 0.32) / 2, items[i][0], 0.32);
      T(s, items[i][1], { x: x + 0.55, y, w: w - 0.65, h: rowH, fontFace: HEAD, bold: true, fontSize: 10.5, valign: 'middle' });
    }
    T(s, 'Sexually related: solicitation, indecent exposure, inappropriate photos, sexual requests about attire, sexual messaging.', { x: 0.45 + w + 0.15, y: 3.25, w, h: 0.5, fontSize: 8.5, color: C.soft, valign: 'middle' });
    await tip(s, 4.0, 'No judgment call:', 'every item here goes straight to T&S.', 'FiArrowUpRight');
  }

  {
    const s = content(TS, 'Escalating: the first report', 'The first time the customer reaches out about a T&S concern.',
      'First time the C reaches out: 1) Respond to the C using the New CRM macro "Escalate Report to T&S". 2) Post in the #new-ts-safety-reports Slack channel using the format: Ticket Link: / CID | Job ID | CP ID / Background: (a short summary of what the customer reported). 3) Leave an internal note on the C account containing the Slack link of your escalation. 4) Don\'t resolve the ticket: change the queue in New CRM to T&S Safety Reports. The screenshot shows a sample post (IDs are training placeholders).');
    // Four steps across the top; the sample post and its format below.
    const rows = [
      ['New CRM macro:\n"Escalate Report to T&S"', 'action'],
      ['Post in\n#new-ts-safety-reports', 'action'],
      ['Internal note on the C account\nwith your Slack post link', 'action'],
      ['Don\'t resolve. Change queue\nto T&S Safety Reports', 'end'],
    ];
    const gap = 0.25, w = (9.1 - gap * 3) / 4;
    for (let i = 0; i < rows.length; i++) {
      const x = 0.45 + i * (w + gap);
      await flowBox(s, x, 1.5, w, 0.72, rows[i][0], rows[i][1]);
      badge(s, x - 0.08, 1.38, i + 1);
      if (i < rows.length - 1) flowLine(s, [[x + w, 1.86], [x + w + gap, 1.86]], true);
    }
    await screen(s, 'ts_safety_report.png', 0.45, 2.45, 6.1, 1.95, 'Sample post in #new-ts-safety-reports');
    box(s, 6.7, 2.45, 2.85, 1.95, C.white, C.border);
    eyebrow(s, 'Post format', 6.9, 2.62, 2.5);
    T(s, [
      { text: 'Ticket Link:', options: { breakLine: true } },
      { text: 'CID | Job ID | CP ID', options: { breakLine: true } },
      { text: 'Background:' },
    ], { x: 6.9, y: 2.95, w: 2.55, h: 1.2, fontFace: 'Courier New', fontSize: 10.5, color: C.ink, paraSpaceAfter: 8, valign: 'top' });
  }

  {
    const s = content(TS, 'Policy violations and conduct', 'Also always escalate.',
      'Unwanted communication means personal, repeated, or continuing after being asked to stop. A single job-related follow-up (confirming arrival time, asking about a forgotten item) is routine and doesn\'t qualify on its own. Verbal dispute: a disagreement about the cleaning itself ("you missed a spot") isn\'t a verbal dispute for T&S purposes. It only qualifies when the interaction itself becomes heated, aggressive, or intimidating. It\'s different from harassment, which implies one party targeting the other; a verbal dispute is mutual or situational escalation between both parties. Workplace safety excludes homes with biohazardous conditions. Hurt during cleaning: unless the incident was caused by either party. Every item is always-escalate, except the conditional pattern on the next slides: don\'t assume anything else is conditional.');
    const w = (9.1 - 0.15) / 2;
    listPanel(s, 0.45, 1.45, w, 2.45, 'Conduct', [
      'Discrimination (race, age, nationality, etc.)',
      'Unwanted communication, on or off platform*',
      'C/CP harassment (more than just rude)',
      'Verbal dispute: raised voices, aggressive confrontation†',
      'Unauthorized return to C\'s residence',
      'C felt unsafe: CP brought a helper without permission',
    ], 9.5);
    listPanel(s, 0.45 + w + 0.15, 1.45, w, 2.45, 'Safety and substances', [
      'Drugs or weapons present at the workplace',
      'Possessing or using illegal substances',
      'Distributing (or trying to) illegal substances',
      'Soliciting or drinking alcohol',
      'CP failed to secure C\'s residence',
      'Animal involvement (bite, injury, lost animal)',
      'C/CP hurt during cleaning',
    ], 9.5);
    T(s, [
      { text: '* Not a single job-related follow-up (arrival time, forgotten item).', options: { breakLine: true } },
      { text: '† Not a disagreement about the cleaning itself, unless it turns heated or intimidating.' },
    ], { x: 0.45, y: 3.98, w: 9.1, h: 0.45, fontSize: 8.5, color: C.soft, italic: true });
  }

  {
    const s = content(TS, 'Theft and damage', 'Which claims are covered.',
      'Theft and damage claims go to T&S. Covered claims: the job was paid in full on the platform; the requester\'s account is in good standing with no outstanding balances (no failed or disputed charges); the claim was reported within 30 days of the cleaning; the requester hasn\'t violated the Terms of Service. Excluded: cleaning done outside the platform; losses of cash, third-party gift cards/vouchers and securities, as well as fine arts, antiques and jewelry; losses of pets, personal liability, damage to common areas, and sentimental or undocumented intangible value; items already recovered by the police or replaced by the cleaner; items that still work (minor cosmetic damage, scratches, ordinary wear and tear).');
    await twoCol(s,
      { ico: 'FiCheckCircle', title: 'Included', items: ['Job paid in full on the platform', 'Account in good standing: no failed or disputed charges', 'Reported within 30 days of the cleaning', 'Requester hasn\'t violated the Terms of Service'] },
      { ico: 'FiXCircle', title: 'Excluded', items: ['Cleaning done off the platform', 'Cash, gift cards, securities, fine art, antiques, jewelry', 'Pets, liability, common areas, sentimental value', 'Already recovered by police or replaced by the CP', 'Items that still work (scratches, wear and tear)'] },
      1.45, 3.0);
  }

  {
    const s = content(TS, 'Escalating: conduct, theft and damage', 'Policy violations, unprofessional conduct, theft and damage: the first report.',
      'Use this flow for policy violations and unprofessional conduct, and for theft and damage. First time the C reaches out: 1) Respond to the C using the New CRM macro "Escalate Report to T&S". 2) Post in the #new-ts-nonsafety-reports Slack channel using the format: Ticket Link: / CID | Job ID | CP ID / Background: (a short summary of what the customer reported). 3) Leave an internal note on the C account containing the Slack link of your escalation. 4) Don\'t resolve the ticket: change the queue in New CRM to T&S Non-Safety Reports. Safety concerns go to #new-ts-safety-reports instead (see the first-report slide). The screenshot shows a sample post (IDs are training placeholders).');
    // Same four steps as the safety flow, with the non-safety channel and queue.
    const rows = [
      ['New CRM macro:\n"Escalate Report to T&S"', 'action'],
      ['Post in\n#new-ts-nonsafety-reports', 'action'],
      ['Internal note on the C account\nwith your Slack post link', 'action'],
      ['Don\'t resolve. Queue:\nT&S Non-Safety Reports', 'end'],
    ];
    const gap = 0.25, w = (9.1 - gap * 3) / 4;
    for (let i = 0; i < rows.length; i++) {
      const x = 0.45 + i * (w + gap);
      await flowBox(s, x, 1.5, w, 0.72, rows[i][0], rows[i][1]);
      badge(s, x - 0.08, 1.38, i + 1);
      if (i < rows.length - 1) flowLine(s, [[x + w, 1.86], [x + w + gap, 1.86]], true);
    }
    await screen(s, 'ts_nonsafety_report.png', 0.45, 2.45, 6.1, 2.2, 'Sample post in #new-ts-nonsafety-reports');
    box(s, 6.7, 2.45, 2.85, 2.2, C.white, C.border);
    eyebrow(s, 'Post format', 6.9, 2.62, 2.5);
    T(s, [
      { text: 'Ticket Link:', options: { breakLine: true } },
      { text: 'CID | Job ID | CP ID', options: { breakLine: true } },
      { text: 'Background:' },
    ], { x: 6.9, y: 2.95, w: 2.55, h: 1.2, fontFace: 'Courier New', fontSize: 10.5, color: C.ink, paraSpaceAfter: 8, valign: 'top' });
  }

  {
    const s = content(TS, 'Conditional escalation', 'A few real concerns escalate only if the customer explicitly asks.',
      'A small number of situations are genuine T&S concerns, but only get escalated to the T&S team if the customer explicitly asks to be escalated for that reason. If the customer doesn\'t ask, handle it as BAU, but it still qualifies as a T&S scenario for ETF waiver purposes, because the underlying safety issue is real either way. Known examples: CP brought a minor to the appointment; CP brought or sent an unauthorized/third-party person. The list isn\'t exhaustive: other situations may fit the same pattern (a genuine safety concern a customer mentions in passing). Before treating a new situation this way, confirm it\'s genuinely comparable: a real safety concern, not just something vaguely safety-adjacent. When unsure, escalate to T&S rather than deciding solo. Note the difference from the always-escalate item: if C says they felt unsafe because the CP brought a helper without permission, escalate.');
    const lx = 2.55, rx = 7.45, w = 3.6;
    await flowBox(s, 5.0 - 2.3, 1.45, 4.6, 0.5, 'CP brought a minor, or brought/sent an unauthorized person', 'penalty');
    flowLine(s, [[5.0, 1.95], [5.0, 2.12]], true);
    await flowBox(s, 5.0 - 1.9, 2.12, 3.8, 0.45, 'Did the customer explicitly ask to escalate?', 'decision');
    flowSplit(s, 5.0, 2.57, 2.77, lx, rx, 3.0, 'No', 'Yes');
    await flowBox(s, lx - w / 2, 3.0, w, 0.45, 'Handle as BAU', 'end');
    await flowBox(s, rx - w / 2, 3.0, w, 0.45, 'Escalate to T&S', 'action');
    T(s, 'Still counts as a T&S scenario for an ETF waiver.', { x: lx - w / 2, y: 3.5, w, h: 0.3, fontSize: 9, color: C.soft, align: 'center' });
    T(s, 'T&S owns the case from here.', { x: rx - w / 2, y: 3.5, w, h: 0.3, fontSize: 9, color: C.soft, align: 'center' });
    await tip(s, 4.1, 'Not exhaustive:', 'a new case must be a real safety concern. Unsure? Escalate.', 'FiHelpCircle');
  }

  {
    const s = content(TS, 'Escalating: follow-ups', 'Updates, a report made by mistake, or a triggered T&S reminder.',
      'Use this flow when the C is asking for an update, the C says they made a mistake in reporting, or the action item is just a triggered timed reminder from T&S. 1) Check if a post for the ticket has already been made in the #trust-and-safety-follow-ups Slack channel. Yes: add a response to the original thread and tick "Also send to #trust-and-safety-follow-ups" to bump the message. No: create a new post using the format: Link to original escalation: / Link to the new comms received (if any): / Background: , then leave an internal note on the account containing the Slack link of your follow-up escalation. 2) Don\'t resolve the ticket: change the queue in New CRM to T&S Follow-ups. Screenshots: a sample follow-up post, and a triggered timed reminder from T&S.');
    const cx = 3.1, lx = 1.75, rx = 4.45, w = 2.55;
    await flowBox(s, 0.45, 1.4, 5.3, 0.62, 'C asks for an update, says they reported by mistake,\nor a timed reminder from T&S triggers', 'penalty');
    flowLine(s, [[cx, 2.02], [cx, 2.17]], true);
    await flowBox(s, cx - 1.75, 2.17, 3.5, 0.5, 'Already posted in\n#trust-and-safety-follow-ups?', 'decision');
    flowSplit(s, cx, 2.67, 2.82, lx, rx, 2.98, 'Yes', 'No');
    await flowBox(s, lx - w / 2, 2.98, w, 0.85, 'Reply in the original thread.\nTick "Also send to\n#trust-and-safety-follow-ups"', 'action');
    await flowBox(s, rx - w / 2, 2.98, w, 0.85, 'New post (format at right).\nInternal note with the\nSlack link of your post', 'action');
    flowLine(s, [[lx, 3.83], [lx, 4.0]]);
    flowLine(s, [[rx, 3.83], [rx, 4.0]]);
    flowLine(s, [[lx, 4.0], [rx, 4.0]]);
    flowLine(s, [[cx, 4.0], [cx, 4.15]], true);
    await flowBox(s, 0.45, 4.15, 5.3, 0.48, 'Don\'t resolve. Change the queue to T&S Follow-ups', 'end');
    await screen(s, 'ts_followup.png', 5.95, 1.4, 3.6, 1.45, 'Follow-up post format');
    await screen(s, 'ts_reminder.png', 5.95, 3.0, 3.6, 1.68, 'Triggered timed reminder from T&S');
  }

  await knowledgeCheck(TS, [
    ['C mentions in passing that the CP brought their teenager along, but doesn\'t ask to escalate. What do you do?', 'Handle as BAU. It still counts as a T&S scenario for an ETF waiver.'],
    ['C says "you missed a spot" and the CP disagreed calmly. Is this a T&S verbal dispute?', 'No. A disagreement about the cleaning only qualifies if it turns heated, aggressive or intimidating.'],
    ['C reports the CP sent sexual messages and asks for a refund. What do you offer?', 'Nothing yourself. Escalate to T&S and email the customer; T&S decides the refund.'],
    ['T&S replies: "This is not a T&S concern." Now what?', 'It\'s yours again as BAU: normal playbook, ordinary ETF and retention rules.'],
  ], 'Ask each question and let trainees answer before revealing. Each click reveals the next answer.');

  await wrapUp(TS, [
    ['Recognize and escalate', 'Care doesn\'t investigate or resolve T&S cases.'],
    ['Hands off', 'No refunds, ETF changes, retention or cleaner contact.'],
    ['Unsure? Escalate', 'You don\'t need certainty to raise a risk.'],
  ], 'TRAINER: recap the three takeaways, then open the floor for questions on Trust and Safety.');

  // ================= 9. SERVICE RECOVERY =================
  const SR = 'Service Recovery';
  await topic('Service Recovery', 'Cases with legal, reputational or relationship risk.',
    ['What Service Recovery handles', 'Material vs. general legal threats', 'How to escalate', 'When a case comes back'],
    'ill10.png',
    'Category: Specialized Escalation Teams. T&S exists because some situations carry real physical or safety risk. Service Recovery exists for a different reason: some situations carry real legal, reputational, or relationship risk. A customer unhappy enough to leave negative feedback and ask for a callback, or one invoking legal action or the FTC, needs higher-touch handling: the stakes go beyond one bad interaction. Care\'s role, same as with T&S, is recognition and handoff, not resolution.');

  {
    const s = content(SR, 'What Service Recovery handles', 'Care recognizes and hands off. SR decides the resolution.',
      'Service Recovery handles: customers who left negative CSAT survey feedback and opted in to be contacted by phone or email; and high-risk general legal threats (e.g. "I will contact a lawyer") or FTC threats, when the risk of escalation is significantly high. SR determines refund eligibility, scope and any ETF waiver for the cases they own. BBB: a BBB threat alone is handled as BAU, not escalated to SR. An actual filed BBB complaint follows the HDR Ticket Handling Process, not Service Recovery.');
    const w = (9.1 - 0.15) / 2;
    await card(s, 0.45, 1.45, w, 1.7, { ico: 'FiThumbsDown', title: 'Negative CSAT, opted in', body: 'The customer left negative survey feedback and asked to be contacted by phone or email.' });
    await card(s, 0.45 + w + 0.15, 1.45, w, 1.7, { ico: 'FiAlertTriangle', title: 'High-risk legal or FTC threat', body: '"I\'ll get a lawyer" or an FTC mention, when the risk of escalation is significantly high.' });
    await tip(s, 3.35, 'SR decides:', 'refund eligibility, scope and any ETF waiver for the cases they own.', 'FiCheckCircle');
    await tip(s, 3.9, 'BBB:', 'a threat is BAU. A filed complaint follows the HDR Ticket Handling Process.', 'FiInfo');
  }

  {
    const s = content(SR, 'Material vs. general legal threats', 'Only some legal language needs Service Recovery.',
      'Material legal threats (Attorney General threat, formal legal letter, lawyer CC\'d on comms, class action threat, email from a law firm, threat to join ongoing litigation, small claims court threat): no SR escalation needed. The agent waives the ETF and cancels FC immediately. See ETF Waivers > Material Legal Threats for the full reasoning. General legal threats ("I\'ll get a lawyer," FTC mention, social media threat): handle as BAU; escalate to SR only if the risk of escalation is significant. BBB is a special case: a threat is BAU with no escalation; an actual filed complaint follows the HDR Ticket Handling Process, not SR.');
    table(s, ['', 'Material legal threat', 'General legal threat', 'BBB (special case)'], [
      ['Examples', 'Attorney General, formal legal letter, lawyer CC\'d, class action, email from a law firm, joining litigation, small claims court', '"I\'ll get a lawyer," FTC mention, social media threat', '"I\'ll report this to the BBB" vs. an actual filed complaint'],
      ['What you do', 'No SR needed. Waive the ETF and cancel FC immediately.', 'Handle as BAU. Escalate to SR only if the risk is significant.', 'Threat: BAU, no escalation. Filed complaint: HDR ticket process.'],
    ], { y: 1.45, colW: [1.3, 2.6, 2.6, 2.6], fontSize: 9.5, rowH: [0.4, 1.25, 0.85] });
    await tip(s, 4.25, 'Remember:', 'material threats skip SR. The waiver is immediate.', 'FiZap');
  }

  {
    const s = content(SR + '  ·  Practice', 'Material or general?', 'Read each message, decide, then click to reveal.',
      'TRAINER: read each message aloud and have trainees call it: material, general, or BBB. Each click reveals the next answer. Material: waive the ETF and cancel FC immediately, no SR. General: BAU, escalate to SR only if the risk is significant. BBB threat: BAU.');
    const qa = [
      ['"I\'m going to file a complaint with the Attorney General about this."', 'Material. Waive the ETF, cancel FC. No SR.'],
      ['"If this doesn\'t get fixed, I\'m getting a lawyer involved."', 'General. BAU; SR only if the risk is significant.'],
      ['"My lawyer has been CC\'d on this email and will be following up."', 'Material. Waive the ETF, cancel FC. No SR.'],
      ['"I\'m reporting this to the Better Business Bureau."', 'BBB threat. BAU, no escalation.'],
      ['"If this isn\'t resolved, I\'m filing in small claims court."', 'Material. Waive the ETF, cancel FC. No SR.'],
      ['"I\'ll be filing a complaint with the FTC."', 'General. BAU; SR only if the risk is significant.'],
    ];
    const n = qa.length, h = (3.75 - 0.08 * (n - 1)) / n;
    for (let i = 0; i < n; i++) {
      const y = 1.4 + i * (h + 0.08);
      box(s, 0.45, y, 5.05, h, C.white, C.border);
      badge(s, 0.6, y + (h - 0.3) / 2, i + 1);
      T(s, qa[i][0], { x: 1.05, y, w: 4.35, h, fontSize: 9.5, italic: true, valign: 'middle' });
      box(s, 5.6, y, 3.95, h, C.tealSoft);
      T(s, qa[i][1], { x: 5.75, y, w: 3.7, h, fontSize: 9.5, valign: 'middle', objectName: `step${i + 1}Ans` });
    }
  }

  {
    const s = content(SR, 'How to escalate', null,
      'First time this issue is being escalated: submit it through the Transformation to SR Reassignment Form (the green button opens the form). Not the first time, because the C is responding to SR: no form needed. DM the previous SR agent in Slack to let them know the customer responded; if they\'re offline, any SR agent online (currently Avegaile "Avie" Gaco, Liz, LorraineYvone, Ann; check the SR schedule sheet linked in the Knowledge Library). Either way, don\'t resolve the ticket: change the queue to SWAT/Service Recovery in New CRM. Cases marked Handle as BAU or Invalid Escalation are sent back automatically via the Slack channel with SWAT notes: the original agent takes ownership back, reviews the notes (Open SWAT notes button), and proceeds accordingly.');
    const FORM = 'https://forms.gle/LzPKYjneeVsAXpbf9';
    const SWAT = 'https://docs.google.com/spreadsheets/d/1Fn3KZLtrux1RKbrATji5G06gs3uVLljK0xJe1rbQFfk/edit?resourcekey=&gid=213718827#gid=213718827';
    // Clickable pill that opens a link when presenting.
    const linkButton = (x, y, w, label, url) => s.addText(label, { shape: pres.shapes.ROUNDED_RECTANGLE, x, y, w, h: 0.34, rectRadius: 0.08,
      fill: { color: C.teal }, line: { color: C.teal }, fontFace: HEAD, bold: true, fontSize: 10, color: C.white, align: 'center', valign: 'middle',
      hyperlink: { url, tooltip: label } });
    const lx = 2.55, rx = 7.45, w = 3.6;
    await flowBox(s, 5.0 - 1.4, 1.15, 2.8, 0.38, 'ESCALATE TO SR', 'penalty');
    flowLine(s, [[5.0, 1.53], [5.0, 1.68]], true);
    await flowBox(s, 5.0 - 1.95, 1.68, 3.9, 0.45, '1st time this issue is being escalated?', 'decision');
    flowSplit(s, 5.0, 2.13, 2.3, lx, rx, 2.5, 'Yes', '');
    T(s, 'No: C is responding to SR', { x: rx + 0.08, y: 2.32, w: 2.4, h: 0.22, fontSize: 8.5, bold: true, color: C.soft });
    await flowBox(s, lx - w / 2, 2.5, w, 0.55, 'Submit through the Transformation\nto SR Reassignment Form', 'action');
    linkButton(lx - 0.95, 3.13, 1.9, 'Open the form  ↗', FORM);
    await flowBox(s, rx - w / 2, 2.5, w, 0.55, 'DM the previous SR agent in Slack,\nor any SR agent online', 'action');
    T(s, 'No need to submit the form.', { x: rx - w / 2, y: 3.12, w, h: 0.34, fontSize: 9, color: C.soft, align: 'center', valign: 'middle' });
    for (const cx of [lx, rx]) {
      flowLine(s, [[cx, 3.47], [cx, 3.62]], true);
      await flowBox(s, cx - w / 2, 3.62, w, 0.5, 'Don\'t resolve. Change the queue to\nSWAT/Service Recovery in New CRM', 'end');
    }
    await tip(s, 4.35, 'Sent back?', '"Handle as BAU" or "Invalid Escalation": it\'s yours again. Review the SWAT notes.', 'FiCornerDownLeft');
    linkButton(7.75, 4.39, 1.7, 'Open SWAT notes  ↗', SWAT);
  }

  await knowledgeCheck(SR, [
    ['C left negative CSAT feedback and opted in to a call. Where does it go?', 'Service Recovery, via the Transformation to SR Reassignment Form.'],
    ['C writes: "I\'ll report this to the BBB." Escalate to SR?', 'No. A BBB threat is BAU. A filed complaint follows the HDR ticket process.'],
    ['C\'s email CCs their lawyer. Do you escalate to SR?', 'No. It\'s a material threat: waive the ETF and cancel FC immediately.'],
    ['The ticket says "C responded to SR." Do you fill out the form?', 'No. DM the previous SR agent in Slack, or any SR agent online.'],
  ], 'Ask each question and let trainees answer before revealing. Each click reveals the next answer.');

  await wrapUp(SR, [
    ['Recognize and hand off', 'SR decides refunds, scope and ETF waivers.'],
    ['Know the threat type', 'Material: waive now. General: BAU unless high risk. BBB threat: BAU.'],
    ['Use the form', 'Except "C responded to SR": DM the last SR agent.'],
  ], 'TRAINER: recap the three takeaways, then open the floor for questions on Service Recovery.');

  // ---------- Close ----------
  {
    const s = pres.addSlide({ masterName: 'CONTENT' });
    eyebrow(s, 'Wrap-up', 0.62, 1.9);
    T(s, 'Thank you!', { x: 0.62, y: 2.15, w: 5, h: 0.75, fontFace: HEAD, bold: true, fontSize: 34 });
    T(s, 'When in doubt, check the Knowledge Library or ask your Trainer or Team Lead.', { x: 0.62, y: 2.95, w: 4.6, h: 0.5, fontSize: 12, color: C.soft });
    await panel(s, 'ill9.png');
    s.addNotes('Close the module. Recap: OCH (billed more than worked), Unauthorized Hours (worked but never asked), False Invoice (never showed, check who invoiced first), Cash Payment (job status decides the action), Cleaner Didn\'t Show (be certain; reschedule intent matters), Cleaner Cancellation (how they cancelled matters; rematch first), Unauthorized Reschedule (still assigned? how far is the start?), Trust and Safety (recognize and escalate, hands off), and Service Recovery (know the threat type; use the form).');
  }

  // Point every agenda row (card, number and arrow) at its topic's title slide.
  if (topicSlides.length * 3 !== agendaLinks.length) throw new Error(`agenda has ${agendaLinks.length / 3} rows but the deck has ${topicSlides.length} topics`);
  // pptxgenjs fixes a link's target when it's added, so update the stored relationship.
  agendaLinks.forEach((l, i) => { agendaSlide._rels.find(r => r.rId === l._rId).Target = String(topicSlides[Math.floor(i / 3)]); });
  await pres.writeFile({ fileName: OUT });
  console.log('wrote', OUT);
})();
