"""Refunds deck: merge three new topics (built by deck9b_body.js) into the live Google export before the closing
slide, renumber 'Topic 0X of 04' to 'of 07', and re-lay the agenda as 7 linked rows. Everything else stays as is."""
import copy, io, re, sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn

live, new, dst = sys.argv[1:4]
prs = Presentation(live); src = Presentation(new)
S = list(prs.slides)
assert S[-1].shapes[1].text_frame.text.strip() == 'Thank you!'
ref = S[3]; layout = ref.slide_layout
num = [x for x in ref.shapes if x.is_placeholder][0]

added = []
for ss in src.slides:
    d = prs.slides.add_slide(layout)
    for ph in list(d.placeholders): ph._element.getparent().remove(ph._element)
    tree = d.shapes._spTree
    for el in ss.shapes._spTree.iterchildren():
        if el.tag in (qn('p:nvGrpSpPr'), qn('p:grpSpPr')): continue
        el = copy.deepcopy(el)
        for blip in el.iter(qn('a:blip')):
            part = ss.part.related_part(blip.get(qn('r:embed')))
            _, rid = d.part.get_or_add_image_part(io.BytesIO(part.blob))
            blip.set(qn('r:embed'), rid)
        assert not list(el.iter(qn('a:hlinkClick'))), 'unexpected link in new slide'
        tree.append(el)
    ph = copy.deepcopy(num._element); ph.find('.//' + qn('p:cNvPr')).set('id', '9999'); tree.append(ph)
    t = ss._element.find(qn('p:timing'))
    if t is not None: d._element.append(copy.deepcopy(t))
    d.notes_slide.notes_text_frame.text = ss.notes_slide.notes_text_frame.text
    added.append(d)
lst = prs.slides._sldIdLst; ids = list(lst)
close = ids[len(S) - 1]
for el in ids[len(S):]:
    lst.remove(el); close.addprevious(el)

# Topic counters.
for s in prs.slides:
    sh = s.shapes[0]
    if sh.has_text_frame and re.fullmatch(r'TOPIC 0\d OF 04', sh.text_frame.text.strip().upper()):
        for r in sh.text_frame.paragraphs[0].runs: r.text = re.sub(r'(?i)of 04', 'of 07', r.text)

# Agenda: 7 rows, two columns (4 + 3).
ag = prs.slides[1]
sub = ag.shapes[2]; r = sub.text_frame.paragraphs[0].runs; r[0].text = 'Seven topics. Click a topic to jump to it.'
for x in r[1:]: x._r.getparent().remove(x._r)
rows = []
for card in [x for x in ag.shapes if x.shape_type == 1 and x.width > Inches(4)]:
    group = [x for x in ag.shapes if x.shape_type == 1 and x.top == card.top and card.left <= x.left < card.left + card.width]
    rows.append(sorted(group, key=lambda x: x.left))
rows.sort(key=lambda g: (g[0].top, g[0].left))
assert len(rows) == 4 and all(len(g) == 3 for g in rows)
topics = [s for s in prs.slides if s.shapes[0].has_text_frame and s.shapes[0].text_frame.text.strip().upper().startswith('TOPIC 0')]
assert len(topics) == 7
H, GAP = 0.8, 0.1
def place(g, k):
    col, row = (0, k) if k < 4 else (1, k - 4)
    x0 = Inches(0.45 + col * 4.63); y = Inches(1.42 + row * (H + GAP))
    dx = x0 - g[0].left
    for shp in g:
        shp.left = shp.left + dx; shp.top = y; shp.height = Inches(H)
for k, g in enumerate(rows): place(g, k)
for size_shape in [g[0] for g in rows]:
    for p in size_shape.text_frame.paragraphs:
        for rr in p.runs:
            if rr.font.size and rr.font.size >= Pt(13): rr.font.size = Pt(13)
new_rows = [('ETF Charge Declined', 'C agreed to pay the ETF; the charge failed.'),
            ('Failed Charge or Overdue MF', 'What to check before raising the card.'),
            ('Graceful Closure', 'When the answer is no, and C keeps pushing.')]
for j, (title, line) in enumerate(new_rows):
    target = topics[4 + j]
    rid = ag.part.relate_to(target.part, RT.SLIDE)
    g = []
    for shp in rows[0]:
        el = copy.deepcopy(shp._element)
        el.find('.//' + qn('p:cNvPr')).set('id', str(9000 + j * 10 + len(g)))
        for h in el.iter(qn('a:hlinkClick')): h.set(qn('r:id'), rid)
        ag.shapes._spTree.append(el)
        g.append(ag.shapes[-1])
    card, numb, arrow = g
    paras = card.text_frame.paragraphs
    paras[0].runs[0].text = title; [x._r.getparent().remove(x._r) for x in paras[0].runs[1:]]
    paras[1].runs[0].text = line; [x._r.getparent().remove(x._r) for x in paras[1].runs[1:]]
    numb.text_frame.paragraphs[0].runs[0].text = f'{5 + j:02d}'
    place(g, 4 + j)
# Make sure every agenda link lands on its topic.
got = sorted({list(prs.slides).index(x.click_action.target_slide) + 1 for x in ag.shapes if x.click_action.target_slide is not None})
print('agenda targets', got, 'new slides', len(added))
prs.save(dst)
