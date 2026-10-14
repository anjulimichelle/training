"""Day 4 deck follow-up.

Slide 10: Timed Reminder card shows by default (drop its click animation).
Slide 11: recolor the reminder diagram to the deck theme (amber decision node,
teal action node, white/grey outcome).
Slide 12: rebuild Penalty section with the full who-invoiced + penalty check detail.

Shapes are found by text, not id: Google renumbers ids on every export.
Usage: python3 slides10-12-theme-penalty.py in.pptx out.pptx
"""
import sys
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

TEAL, TEAL_PALE, AMBER, AMBER_PALE = '0B7A6F', 'E3F4F1', '9A6410', 'FFF3D6'
INK, MUTED, LINE = '1D1D1F', '6E6E73', 'DCE1E7'

prs = Presentation(sys.argv[1])
s10, s11, s12 = prs.slides[9], prs.slides[10], prs.slides[11]


def txt(sh):
    return sh.text_frame.text if sh.has_text_frame else ''


def find(slide, pred):
    hits = [sh for sh in slide.shapes if pred(sh)]
    assert len(hits) == 1, [txt(h)[:30] for h in hits]
    return hits[0]


def geom(sh, x, y, w, h):
    sh.left, sh.top, sh.width, sh.height = Pt(x), Pt(y), Pt(w), Pt(h)


def fill(sh, fill_hex, line_hex=None, line_w=None):
    sh.fill.solid(); sh.fill.fore_color.rgb = RGBColor.from_string(fill_hex)
    if line_hex:
        sh.line.fill.solid(); sh.line.color.rgb = RGBColor.from_string(line_hex)
    if line_w:
        sh.line.width = Pt(line_w)


def recolor_runs(sh, hex_, para=None):
    ps = sh._element.findall('.//' + qn('a:p'))
    for p in ps if para is None else [ps[para]]:
        for tag in ('a:rPr', 'a:endParaRPr'):
            for r in p.iter(qn(tag)):
                sf = r.find(qn('a:solidFill'))
                if sf is not None:
                    c = sf.find(qn('a:srgbClr'))
                    if c is None:
                        sf.clear(); c = sf.makeelement(qn('a:srgbClr'), {}); sf.append(c)
                    c.set('val', hex_)


def write(sh, paras, anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT, space=0, hang=0):
    """paras: list of list of (text, font, size, bold, color)."""
    tf = sh.text_frame
    body = tf._txBody
    for p in body.findall(qn('a:p'))[1:]:
        body.remove(p)
    p0 = body.find(qn('a:p'))
    for child in list(p0):
        if child.tag != qn('a:pPr'):
            p0.remove(child)
    tf.vertical_anchor = anchor
    tf.word_wrap = True
    for i, runs in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if i and space:
            p.space_before = Pt(space)
        if hang:
            pPr = p._p.get_or_add_pPr()
            pPr.set('marL', str(int(hang * 12700)))
            pPr.set('indent', '0' if runs[0][0].startswith(' ') else str(int(-hang * 12700)))
        for text, font, size, bold, color in runs:
            r = p.add_run(); r.text = text
            f = r.font; f.name = font; f.size = Pt(size); f.bold = bold
            f.color.rgb = RGBColor.from_string(color)


# ---------- slide 10: card visible by default ----------
assert txt(s10.shapes[1]).startswith('Branch B: no clear evidence')
t = s10._element.find(qn('p:timing'))
if t is not None:
    s10._element.remove(t)

# ---------- slide 11: theme colors ----------
assert txt(s11.shapes[1]).startswith('Branch B: when the reminder')
node = find(s11, lambda s: txt(s).startswith('Reminder triggers'))
fill(node, AMBER_PALE, AMBER, 1)
recolor_runs(node, AMBER)
no = find(s11, lambda s: txt(s) == 'No refund')
fill(no, 'FFFFFF', LINE, 1)
recolor_runs(no, INK)
ref = find(s11, lambda s: txt(s).startswith('Refund, then penalty check'))
ref.fill.solid(); ref.fill.fore_color.rgb = RGBColor.from_string(TEAL)
recolor_runs(ref, 'FFFFFF', 0)
if len(ref.text_frame.paragraphs) > 1:  # trainer's sub-line: keep, make it readable on teal
    recolor_runs(ref, TEAL_PALE, 1)
for sh in s11.shapes:
    if sh.shape_type == 1 and sh._element.find('.//' + qn('a:ln')) is not None:
        c = sh._element.find(qn('p:spPr')).find(qn('a:ln')).find('.//' + qn('a:srgbClr'))
        if c is not None and c.get('val') == '666666':
            c.set('val', LINE)
        elif c is not None and c.get('val') == 'E06666':
            c.set('val', TEAL)

# ---------- slide 12: penalty section ----------
sh = list(s12.shapes)
assert txt(sh[1]) == 'Penalty section' and txt(sh[6]) == 'Who invoiced?'
(_, _, sub,
 a_bg, a_badge, a_icon, a_title, a_body,
 b_bg, b_badge, b_icon, b_title, b_body,
 c_bg, c_badge, c_icon, c_title, c_body,
 r_bg, r_txt, n_bg, n_icon, n_txt, _num) = sh

# Row 1: who invoiced
geom(a_bg, 32, 102, 655, 58)
geom(a_badge, 44, 118, 26, 26); geom(a_icon, 50.5, 124.5, 13, 13)
geom(a_title, 78, 102, 110, 58); a_title.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
# reuse the "note the history" card + text as the two answer pills
geom(c_bg, 190, 109, 240, 44); fill(c_bg, 'F5F6F8', LINE, 0.75)
c_bg._element.find(qn('p:spPr')).find(qn('a:prstGeom')).set('prst', 'roundRect')
geom(c_body, 200, 109, 222, 44)
write(c_body, [[('Customer invoiced', 'Play', 10, True, INK)],
               [('→  Refund, no cleaner penalty', 'Arial', 9.5, False, MUTED)]], MSO_ANCHOR.MIDDLE, space=2)
geom(a_body, 448, 109, 223, 44)
geom(c_badge, 438, 109, 241, 44); fill(c_badge, TEAL_PALE, TEAL, 0.75)
c_badge._element.find(qn('p:spPr')).find(qn('a:prstGeom')).set('prst', 'roundRect')
write(a_body, [[('Cleaner invoiced', 'Play', 10, True, TEAL)],
               [('→  Proceed with penalty check ↓', 'Arial', 9.5, False, INK)]], MSO_ANCHOR.MIDDLE, space=2)
c_title._element.getparent().remove(c_title._element)
c_icon._element.getparent().remove(c_icon._element)
# z-order: pills above row card
tree = s12.shapes._spTree
for el in (c_bg._element, c_badge._element, c_body._element, a_body._element):
    tree.remove(el); tree.append(el)

# Row 2 left: under 30 min (uses the flag card shapes)
geom(b_bg, 32, 170, 236, 164)
geom(b_badge, 44, 181, 26, 26); geom(b_icon, 50.5, 187.5, 13, 13)
geom(b_title, 78, 181, 180, 26); b_title.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
write(b_title, [[('Less than 30 min', 'Play', 12, True, INK)]], MSO_ANCHOR.MIDDLE)
geom(b_body, 44, 216, 214, 118)
item = lambda t, bold=False: [('•  ', 'Arial', 10, True, TEAL), (t, 'Arial', 10, bold, INK)]
write(b_body, [item('No penalty', True), item('Ban the C/CP pairing'), item('CP coaching')], space=6, hang=11)

# Row 2 right: 30 min or more (reuse the first card's frame for the bg)
geom(r_bg, 278, 170, 409, 164); r_bg.adjustments[0] = 0.06
fill(r_bg, 'FFFFFF', TEAL, 1.25)
geom(r_txt, 290, 181, 385, 26)
write(r_txt, [[('30 min or more', 'Play', 12, True, INK)]], MSO_ANCHOR.MIDDLE)
body = s12.shapes.add_textbox(Pt(290), Pt(212), Pt(385), Pt(116))
write(body, [
    item('Ban the C/CP pairing'),
    item('Add the cp_overcharged_hours flag'),
    item('Set norequests, with this CP-facing reason:'),
    [(' cpq_overcharge Customer <C\'S NAME> reported incorrect hours charged. '
      'Submit reactivation appeal.', 'Courier New', 8.5, False, AMBER)],
    item('Deactivation internal note on the CP account: include past OCH and FI '
         'if you see a pattern of incorrect charging'),
], space=4, hang=11)
# flag badge belongs to the 30+ card; the under-30 card is plain
geom(b_badge, 290, 181, 26, 26); geom(b_icon, 296.5, 187.5, 13, 13)
geom(r_txt, 324, 181, 350, 26)
geom(b_title, 44, 181, 214, 26)
for el in (b_badge._element, b_icon._element):
    tree.remove(el); tree.append(el)

# Never bar
geom(n_bg, 32, 342, 655, 26)
geom(n_icon, 44, 348.5, 13, 13)
geom(n_txt, 64, 342, 615, 26); n_txt.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

prs.save(sys.argv[2])
print('saved', sys.argv[2])

# speaker notes: add the new penalty-check detail (appended once)
from pptx import Presentation as _P
prs2 = _P(sys.argv[2])
nt = prs2.slides[11].notes_slide.notes_text_frame
ADD = ("Penalty checks once the cleaner invoiced. Under 30 minutes: no penalty; ban the C/CP pairing "
       "and coach the CP. 30 minutes or more: ban the C/CP pairing, add the cp_overcharged_hours flag, "
       "set norequests with the CP-facing reason above, and leave a deactivation internal note on the CP "
       "account that includes past OCH and FI if you see a pattern of incorrect charging.")
if ADD not in nt.text:
    nt.add_paragraph().text = ADD
prs2.save(sys.argv[2])
