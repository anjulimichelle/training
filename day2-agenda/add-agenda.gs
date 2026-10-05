/**
 * Adds a clickable Agenda slide (slide 2) to the Day 2 Training presentation.
 *
 * Each agenda topic links to that topic's title slide, so clicking it in
 * slideshow mode jumps straight there. Nothing else in the deck is changed.
 *
 * How to run:
 *   1. Go to https://script.google.com and click "New project".
 *   2. Delete the sample code, paste this whole file, and click Save.
 *   3. Pick "addAgenda" in the function dropdown and click Run.
 *   4. Approve the permission prompt (it only needs access to your Slides).
 *
 * Safety: the script first finds every destination slide. If any topic slide
 * is missing or ambiguous, or an Agenda slide already exists, it stops before
 * changing anything.
 */
const PRESENTATION_ID = '1O3sZWqlug6jj3Fu_TmGZQNE2GYuVErRj97D2uF7ADdw';

// [agenda label, title text on that topic's title slide]
const ITEMS = [
  ['Cleaner Journey', 'Cleaner Journey'],
  ['Intro to CP CRM (Legacy)', 'Introduction to CP CRM (Legacy)'],
  ['Intro to CP Dashboard', 'Introduction to CP Dashboard'],
  ['CP Profile Status', 'CP Profile Status'],
  ['CP-Facing Deactivation', 'CP-Facing Deactivation Reasons'],
  ['Operating Models', 'Operating Models'],
  ['CP Tiering', 'CP Tiering'],
  ['Job Status', 'Job Statuses'],
  ['AG: Pending Invoice', 'AG: Pending Invoice'],
];

// The slide whose label, title and slide number the Agenda reuses.
const TEMPLATE_TITLE = 'The ideal cleaner journey';

// The deck's own palette.
const TEAL = '#0B7A6F', TEAL_SOFT = '#E3F4F1', INK = '#1D1D1F', SOFT = '#6E6E73', BORDER = '#DCE1E7', WHITE = '#FFFFFF';

function norm(t) { return t.replace(/\s+/g, ' ').trim(); }

function shapeText(shape) {
  try { return norm(shape.getText().asString()); } catch (e) { return ''; }
}

function slideTexts(slide) { return slide.getShapes().map(shapeText); }

function addAgenda() {
  const pres = SlidesApp.openById(PRESENTATION_ID);
  const slides = pres.getSlides();

  if (slides.some(s => { const t = slideTexts(s); return t.includes('Agenda') && t.includes('OVERVIEW'); })) {
    throw new Error('An Agenda slide already exists. Nothing was changed.');
  }

  // Find each topic's title slide: its title text plus a "TOPIC ..." label.
  const dest = ITEMS.map(([label, title]) => {
    const found = slides.filter(s => {
      const t = slideTexts(s);
      return t.includes(title) && t.some(x => /^TOPIC \d/.test(x));
    });
    if (found.length !== 1) {
      throw new Error(`Expected one topic slide titled "${title}" (for "${label}"), found ${found.length}. Nothing was changed.`);
    }
    return found[0];
  });

  // Start from a plain content slide so the background, label, title and footer match the deck.
  const template = slides.find(s => slideTexts(s).includes(TEMPLATE_TITLE));
  if (!template) throw new Error(`Could not find the "${TEMPLATE_TITLE}" slide to copy the design from. Nothing was changed.`);
  const eyebrowText = slideTexts(template).find(t => t && t === t.toUpperCase() && /[A-Z]/.test(t));

  const s = template.duplicate();
  s.move(1); // becomes slide 2, right after the title slide

  // Keep the label, the title and any placeholder (the slide number); remove everything else.
  let titleShape = null, eyebrowShape = null;
  s.getPageElements().forEach(el => {
    let shape = null;
    try { shape = el.asShape(); } catch (e) { /* not a shape: picture, table, line */ }
    if (shape && shape.getPlaceholderType() !== SlidesApp.PlaceholderType.NONE) return;
    const t = shape ? shapeText(shape) : '';
    if (t === TEMPLATE_TITLE) titleShape = shape;
    else if (t === eyebrowText) eyebrowShape = shape;
    else el.remove();
  });
  titleShape.getText().setText('Agenda');
  eyebrowShape.getText().setText('OVERVIEW');
  const HEAD = titleShape.getText().getTextStyle().getFontFamily();

  const sub = s.insertTextBox('Click a topic to jump to it.', 32, 72, 655, 20);
  sub.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setForegroundColor(SOFT);

  // Two columns: topics 1–5 on the left, 6–9 on the right.
  const colW = 320, gapX = 15, rowH = 46, gapY = 8, top = 102;
  ITEMS.forEach(([label], i) => {
    const col = i < 5 ? 0 : 1, row = i < 5 ? i : i - 5;
    const x = 32 + col * (colW + gapX), y = top + row * (rowH + gapY);
    const target = dest[i];

    const card = s.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, y, colW, rowH);
    card.getFill().setSolidFill(WHITE);
    card.getBorder().setWeight(0.75).getLineFill().setSolidFill(BORDER);

    const badge = s.insertShape(SlidesApp.ShapeType.ELLIPSE, x + 12, y + 11, 24, 24);
    badge.getFill().setSolidFill(TEAL_SOFT);
    badge.getBorder().setTransparent();
    badge.getText().setText(String(i + 1));
    badge.getText().getTextStyle().setFontFamily(HEAD).setFontSize(10).setBold(true).setForegroundColor(TEAL);
    badge.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    badge.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

    const text = s.insertTextBox(label, x + 46, y, colW - 76, rowH);
    text.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

    const arrow = s.insertTextBox('›', x + colW - 30, y, 22, rowH);
    arrow.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

    // Link every piece of the row, so a click anywhere on it jumps to the topic.
    [card, badge, text, arrow].forEach(el => el.setLinkSlide(target));
    // Text links default to blue + underline; restyle them to match the deck.
    text.getText().getTextStyle().setLinkSlide(target).setFontFamily(HEAD).setFontSize(12).setBold(true).setUnderline(false).setForegroundColor(INK);
    arrow.getText().getTextStyle().setLinkSlide(target).setFontFamily(HEAD).setFontSize(18).setBold(true).setUnderline(false).setForegroundColor(TEAL);
  });

  s.getNotesPage().getSpeakerNotesShape().getText().setText(
    'Walk through today\'s agenda. In slideshow mode, click any topic to jump straight to it.');

  const ids = pres.getSlides().map(x => x.getObjectId());
  Logger.log('Agenda added as slide 2. Links: ' + ITEMS.map(([l], i) => `${l} -> slide ${ids.indexOf(dest[i].getObjectId()) + 1}`).join('; '));
}
