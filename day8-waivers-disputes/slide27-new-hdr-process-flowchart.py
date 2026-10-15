"""Day 8 deck slide 27: replace the "Which track am I on?" table with the New HDR ticket
handling process as a flow chart (greater vs less than the card-based threshold, refund
intent vs cancellation + refund intent, 1st and 2nd responses, 48-hour rule). A
"Sample internal note" button opens a pop-up on click 1 and closes it on click 2.

Usage: python3 slide27-new-hdr-process-flowchart.py in.pptx out.pptx
"""
import sys
from lxml import etree
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

TEAL, TEAL_PALE, AMBER, AMBER_PALE = '0B7A6F', 'E3F4F1', '9A6410', 'FFF3D6'
INK, MUTED, WHITE, GREY = '1D1D1F', '6E6E73', 'FFFFFF', '6E6E73'

prs = Presentation(sys.argv[1])
s = prs.slides[26]
txt = lambda sh: sh.text_frame.text if sh.has_text_frame else ''
assert txt(s.shapes[1]) == 'Which track am I on?', txt(s.shapes[1])
old_notes = s.notes_slide.notes_text_frame.text
sl = list(s.shapes)
keep = {id(x._element) for x in sl[:3]} | {id(x._element) for x in sl if x.is_placeholder}
for sh in sl:
    if id(sh._element) not in keep:
        sh._element.getparent().remove(sh._element)


def set_text(shape, t):
    p = shape.text_frame.paragraphs[0]
    p.runs[0].text = t
    for r in p.runs[1:]:
        r._r.getparent().remove(r._r)


set_text(s.shapes[1], 'New HDR ticket handling process')
set_text(s.shapes[2], 'Start by comparing the charge C wants refunded with the card-based threshold.')


def strip(sh):
    st = sh._element.find(qn('p:style'))
    if st is not None:
        sh._element.remove(st)


def box(x, y, w, h, fill, line, paras, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=5):
    sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(x), Pt(y), Pt(w), Pt(h)); strip(sh)
    sh.adjustments[0] = min(0.3, 8 / min(w, h))
    sh.fill.solid(); sh.fill.fore_color.rgb = RGBColor.from_string(fill)
    if line:
        sh.line.color.rgb = RGBColor.from_string(line); sh.line.width = Pt(1)
    else:
        sh.line.fill.background()
    tf = sh.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Pt(margin); tf.margin_top = tf.margin_bottom = Pt(2)
    for i, runs in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align; p.space_after = Pt(2.5)
        for t, size, bold, color in runs:
            r = p.add_run(); r.text = t; f = r.font
            f.name = 'Play' if bold else 'Arial'; f.size = Pt(size); f.bold = bold
            f.color.rgb = RGBColor.from_string(color)
    return sh


def line(x1, y1, x2, y2, head=True):
    c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Pt(x1), Pt(y1), Pt(x2), Pt(y2)); strip(c)
    c.line.color.rgb = RGBColor.from_string(GREY); c.line.width = Pt(1)
    if head:
        ln = c._element.find('.//' + qn('a:ln'))
        etree.SubElement(ln, qn('a:tailEnd'), type='triangle', w='med', len='med')
    return c


B = lambda t, size=8, color=INK: (t, size, True, color)
A = lambda t, size=8, color=INK: (t, size, False, color)

# column 1: start
box(32, 180, 100, 100, TEAL, None,
    [[A('Check if the charge C wants refunded is ', 8.5, WHITE), B('greater or less than', 8.5, WHITE),
      A(' the card-based threshold', 8.5, WHITE)]])
line(132, 230, 140, 230, head=False)
line(140, 135, 140, 290, head=False)
line(140, 135, 148, 135); line(140, 290, 148, 290)

# column 2: greater / less
box(148, 115, 92, 40, AMBER_PALE, AMBER, [[B('Greater than', 8.5, AMBER)], [A('the card-based threshold', 7.5)]])
box(148, 270, 92, 40, TEAL_PALE, TEAL, [[B('Less than', 8.5, TEAL)], [A('the card-based threshold', 7.5)]])

# greater branch
line(240, 135, 255, 135)
box(255, 110, 120, 50, AMBER_PALE, AMBER, [[A('Follow the usual ', 8), B('Transformation handling process', 8, AMBER), A('.', 8)]])
line(375, 135, 390, 135)
box(390, 98, 297, 76, AMBER_PALE, AMBER, [
    [B('2nd response: ', 7.5, AMBER), A('if C comes back still insisting on a refund, ', 7.5),
     B('do not issue a refund', 7.5), A(', unless another policy supersedes this handling and makes C eligible (e.g. a material legal threat). '
                                     "First check for applicable offers or existing policies that can address C's concern.", 7.5)],
    [A('If the case leads to a dispute, the Dispute Response Team responds to the disputed charge.', 7.5, MUTED)]],
    align=PP_ALIGN.LEFT, margin=8)

# less branch
line(240, 290, 247, 290, head=False)
line(247, 222, 247, 320, head=False)
line(247, 222, 255, 222); line(247, 320, 255, 320)
box(255, 204, 88, 36, TEAL_PALE, TEAL, [[A('C has ', 8), B('refund intent only', 8, TEAL)]])
box(255, 298, 88, 44, TEAL_PALE, TEAL, [[A('C has ', 8), B('cancellation and refund intent', 8, TEAL)]])
line(343, 222, 353, 222); line(343, 320, 353, 320)
box(353, 197, 122, 50, WHITE, TEAL, [[B('1st response', 8, TEAL)], [A('Follow Transformation Guidelines for Refund Requests', 7.5)]])
box(353, 295, 122, 50, WHITE, TEAL, [[B('1st response', 8, TEAL)], [A('Follow Transformation Retention Guidelines', 7.5)]])
line(475, 222, 485, 222); line(475, 320, 485, 320)


def second(y, h, extra=None):
    paras = [[B('2nd response: ', 7.5, TEAL), A('if C comes back insisting on a refund, check if it has been more than 48 hours since C was flagged HDR.', 7.5)],
             [B('≤ 48 hours: ', 7.5, AMBER), A('leave an internal note of your card-based threshold assessment and refund the charge.' + (' ' + extra if extra else ''), 7.5)],
             [B('> 48 hours: ', 7.5, TEAL), A('handle as usual.', 7.5)]]
    return box(485, y, 202, h, TEAL_PALE, TEAL, paras, align=PP_ALIGN.LEFT, margin=7)


second(184, 76)
second(272, 90, 'Handle the cancellation intent as usual.')

# button + pop-up
btn = box(32, 330, 100, 32, AMBER, None, [[B('Sample internal note', 8, WHITE)], [A('click to open ▸', 7, AMBER_PALE)]])
btn.name = 'SampleNoteButton'
pop = []
px, py, pw, ph = 140, 112, 440, 200
card = box(px, py, pw, ph, WHITE, AMBER, [[A('')]])
card.line.width = Pt(1.5)
sp = card._element.find(qn('p:spPr'))
eff = etree.SubElement(sp, qn('a:effectLst'))
sh_ = etree.SubElement(eff, qn('a:outerShdw'), blurRad='76200', dist='25400', dir='5400000', algn='t', rotWithShape='0')
clr = etree.SubElement(sh_, qn('a:srgbClr'), val='000000'); etree.SubElement(clr, qn('a:alpha'), val='25000')
pop.append(card)
pop.append(box(px, py, pw, 28, AMBER, None, [[B('Sample internal note', 11, WHITE), A('    card-based threshold assessment', 8.5, AMBER_PALE)]],
               align=PP_ALIGN.LEFT, margin=12))
note = ['FC Active | ETF Not waived',
        'Root cause: Moving out of the country | PCQ Unaware | Does not want membership',
        'Disputable amount: $49',
        'Card & threshold: Mastercard <58',
        'Assessment: Less',
        'Action: Refund']
pop.append(box(px + 14, py + 38, pw - 28, ph - 66, AMBER_PALE, INK,
               [[A(l, 9.5)] for l in note], align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE, margin=12))
pop.append(box(px + 14, py + ph - 22, pw - 28, 14, WHITE, None, [[A('Click to close', 7.5, MUTED)]], align=PP_ALIGN.RIGHT, margin=0))


def effect(cid, spid, node):
    tgt = f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
    return (f'<p:par><p:cTn id="{cid}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="{node}">'
            f'<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
            f'<p:set><p:cBhvr><p:cTn id="{cid+1}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>{tgt}'
            f'<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
            f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{cid+2}" dur="400"/>{tgt}</p:cBhvr></p:animEffect>'
            f'</p:childTnLst></p:cTn></p:par>')


def leave(cid, spid, node):
    tgt = f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
    return (f'<p:par><p:cTn id="{cid}" presetID="1" presetClass="exit" presetSubtype="0" fill="hold" grpId="1" nodeType="{node}">'
            f'<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
            f'<p:set><p:cBhvr><p:cTn id="{cid+1}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>{tgt}'
            f'<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="hidden"/></p:to></p:set>'
            f'</p:childTnLst></p:cTn></p:par>')


def click(cid, inner):
    return (f'<p:par><p:cTn id="{cid}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>'
            f'<p:par><p:cTn id="{cid+1}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
            f'{inner}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>')


ids = [sh.shape_id for sh in pop]
cid, c1 = 3, ''
for i, spid in enumerate(ids):
    c1 += effect(cid + 2, spid, 'clickEffect' if i == 0 else 'withEffect'); cid += 3
click1 = click(3, c1); cid += 2
start2, c2 = cid, ''
cid += 2
for i, spid in enumerate(ids):
    c2 += leave(cid, spid, 'clickEffect' if i == 0 else 'withEffect'); cid += 2
click2 = click(start2, c2)
bld = ''.join(f'<p:bldP spid="{i}" grpId="0"/><p:bldP spid="{i}" grpId="1"/>' for i in ids)
NS = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
tim = ('<p:timing ' + NS + '><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
       '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
       + click1 + click2 +
       '</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
       '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>'
       '</p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst>' + bld + '</p:bldLst></p:timing>')
old = s._element.find(qn('p:timing'))
if old is not None:
    s._element.remove(old)
t_el = etree.fromstring(tim)
ext = s._element.find(qn('p:extLst'))
(ext.addprevious(t_el) if ext is not None else s._element.append(t_el))

s.notes_slide.notes_text_frame.text = (
    'New HDR ticket handling process. Start by checking whether the charge C wants refunded is greater or less than the card-based threshold.\n'
    'Greater than the threshold: follow the usual Transformation handling process. If C comes back still insisting on a refund, do not issue a refund, '
    'unless another policy supersedes this handling and makes C eligible for a refund (for example a material legal threat). First check for any applicable '
    "offers or existing policies that can address C's concern. If the case eventually leads to a dispute, the Dispute Response Team responds to the disputed charge.\n"
    'Less than the threshold, refund intent only: 1st response, follow the Transformation Guidelines for Refund Requests. '
    'Less than the threshold, cancellation and refund intent: 1st response, follow the Transformation Retention Guidelines.\n'
    'On either path, 2nd response if C comes back insisting on a refund: check whether more than 48 hours have passed since C was flagged HDR. '
    'Within 48 hours: leave an internal note of your card-based threshold assessment and refund the charge (and handle any cancellation intent as usual). '
    'After 48 hours: handle as usual.\n'
    'TRAINER: click once to open the sample internal note, click again to close it. Sample: FC Active | ETF Not waived; Root cause: Moving out of the country | '
    'PCQ Unaware | Does not want membership; Disputable amount: $49; Card & threshold: Mastercard <58; Assessment: Less; Action: Refund.\n'
    'Which track applies (from the previous version of this slide): ' + old_notes)
prs.save(sys.argv[2]); print('saved', sys.argv[2])
