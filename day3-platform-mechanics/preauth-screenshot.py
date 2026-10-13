"""Day 3 original deck: slide 34 gets a "how to spot it" panel with a pre-authorization payment
screenshot; its "Watch out" reminder moves under the table on slide 35. Patches the live export."""
import copy, sys
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn

src, dst, img = sys.argv[1:4]
C = RGBColor.from_string
TEAL, INK, SOFT, BORDER, WHITE = C('0B7A6F'), C('1D1D1F'), C('6E6E73'), C('DCE1E7'), C('FFFFFF')
prs = Presentation(src)
s34, s35 = prs.slides[33], prs.slides[34]
txt = lambda s: ' '.join(sh.text_frame.text for sh in s.shapes if sh.has_text_frame)
assert 'What is a pre-auth?' in txt(s34) and 'Standard vs exception' in txt(s35)
sh34 = list(s34.shapes)
watch = [sh for sh in sh34 if sh.top >= Inches(3.45) and sh.top < Inches(4.0)]
assert any(sh.has_text_frame and sh.text_frame.text.startswith('Watch out') for sh in watch), [s.name for s in watch]

# ---- move Watch out (bar, icon, text) to slide 35, just under the table ----
table = next(sh for sh in s35.shapes if sh.has_table)
tbl_bottom = table.top + sum(r.height for r in table.table.rows)
dy = tbl_bottom + Inches(0.18) - min(sh.top for sh in watch)
for sh in watch:
    el = copy.deepcopy(sh._element)
    blip = el.find('.//' + qn('a:blip'))
    if blip is not None:
        part = s34.part.related_part(blip.get(qn('r:embed')))
        blip.set(qn('r:embed'), s35.part.relate_to(part, RT.IMAGE))
    s35.shapes._spTree.append(el)
    new = s35.shapes[-1]; new.top = new.top + dy
    sh._element.getparent().remove(sh._element)

# ---- slide 34: "How to spot it" panel with the screenshot ----
X, Y, W, H = 0.45, 3.5, 9.1, 1.38
card = s34.shapes.add_shape(5, Inches(X), Inches(Y), Inches(W), Inches(H))
st = card._element.find(qn('p:style'))
if st is not None: card._element.remove(st)
card.adjustments[0] = 0.06
card.fill.solid(); card.fill.fore_color.rgb = WHITE
card.line.color.rgb = BORDER; card.line.width = Emu(9525)

def text(x, y, w, h, runs, size):
    tb = s34.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)); tf = tb.text_frame
    tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.TOP
    for m in ('margin_left', 'margin_right', 'margin_top', 'margin_bottom'): setattr(tf, m, 0)
    for t, bold, color, font in runs:
        r = tf.paragraphs[0].add_run(); r.text = t; f = r.font
        f.size = Pt(size); f.bold = bold; f.color.rgb = color; f.name = font
text(X + 0.18, Y + 0.14, 8.7, 0.25, [('How to spot a pre-auth', True, INK, 'Play')], 12)
text(X + 0.18, Y + 0.42, 8.7, 0.22, [('The payment shows ', False, SOFT, 'Arial'), ('Reason: Pre-Authorization', True, INK, 'Arial'),
                                     (' with a Job ID. You can check it in Stripe.', False, SOFT, 'Arial')], 9.5)
pic = s34.shapes.add_picture(img, Inches(X + 0.18), Inches(Y + 0.76), width=Inches(8.74))
pic.line.color.rgb = BORDER; pic.line.width = Emu(9525)
print('screenshot height in', round(pic.height / 914400, 2), 'panel bottom', round(Y + 0.76 + pic.height / 914400, 2))

n = s34.notes_slide.notes_text_frame
n.text = n.text + ("\nHow to spot a pre-auth: in the payment list it shows Reason: Pre-Authorization with a Job ID "
                   "(in the example it is canceled, so the hold was released). You can look it up in Stripe to confirm it is a hold, not a charge.")
prs.save(dst)
