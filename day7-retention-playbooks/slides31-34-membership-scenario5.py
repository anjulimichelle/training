"""Day 7 deck, Membership Retention Scenario 5 (multiple unused MFs due to inactivity, or forgot about FC).

Slide 31: overview (what it looks like, scope check, probe-first guidance with a sample line).
New 32:   within MCT flow chart (root cause -> genuinely inactive offer -> reduced ETF hint -> exit / closure).
New 33:   beyond MCT flow chart with the "When the TR is triggered" pop-up (click 1 shows, click 2 hides).
New 34:   refund requests after long inactivity (last 120 days beyond MCT; 12-paid / 6-MCT example).
Helpers come from slides23-25-membership-scenario1.py.

Usage: python3 slides31-34-membership-scenario5.py in.pptx out.pptx
"""
import copy, sys
from lxml import etree
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

TEAL, TEAL_PALE, AMBER, AMBER_PALE = '0B7A6F', 'E3F4F1', '9A6410', 'FFF3D6'
INK, MUTED, WHITE, LINE = '1D1D1F', '6E6E73', 'FFFFFF', 'DCE1E7'

prs = Presentation(sys.argv[1])
txt = lambda sh: sh.text_frame.text if sh.has_text_frame else ''
S = list(prs.slides)
assert txt(S[30].shapes[1]) == 'Scenario 5: inactivity or "forgot" about FC', txt(S[30].shapes[1])
assert txt(S[31].shapes[1]) == 'Scenario 6: moving', txt(S[31].shapes[1])
s = None
s8_notes = "The same Timed Reminder rules apply in every beyond-MCT branch. If the next MF is more than 5 days away, set the reminder at least 3 days before it. If it is 5 days or less, give 1 free month first and set the reminder 3 days before the new MF date. When it triggers with no reply and no engagement, cancel FC; the auto comms are enough unless an unresolved pain point needs a manual email. If there is no reply but the customer is using the service, such as a new booking, don't cancel. Confirm with them, push the due date out 1 to 2 weeks, and set a new reminder 3 days before it."


def clear(slide):
    sl = list(slide.shapes)
    keep = {id(e._element) for e in sl[:3]} | {id(e._element) for e in sl if e.is_placeholder}
    for sh in sl:
        if id(sh._element) not in keep:
            sh._element.getparent().remove(sh._element)


def set_text(sh, t, size=None):
    runs = sh.text_frame.paragraphs[0].runs
    runs[0].text = t
    for r in runs[1:]:
        r._r.getparent().remove(r._r)
    for p in sh.text_frame.paragraphs[1:]:
        p._p.getparent().remove(p._p)
    if size:
        runs[0].font.size = Pt(size)


def strip(sh):
    st = sh._element.find(qn('p:style'))
    if st is not None:
        sh._element.remove(st)


def fill_paras(tf, paras, align):
    for i, runs in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        for t, font, size, bold, color in runs:
            r = p.add_run(); r.text = t
            f = r.font; f.name = font; f.size = Pt(size); f.bold = bold
            f.color.rgb = RGBColor.from_string(color)
    return tf


def hang(tf):
    for p in tf.paragraphs:
        lead = p.runs[0].text if p.runs else ''
        if lead.startswith('•'):
            p._p.get_or_add_pPr().set('marL', str(11 * 12700)); p._p.get_or_add_pPr().set('indent', str(-11 * 12700))
        elif lead.startswith('–'):
            p._p.get_or_add_pPr().set('marL', str(24 * 12700)); p._p.get_or_add_pPr().set('indent', str(-10 * 12700))


def box(x, y, wd, h, fill, line, paras, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, lw=1):
    sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(x), Pt(y), Pt(wd), Pt(h))
    strip(sh)
    sh.adjustments[0] = min(0.3, 8 / h)
    sh.fill.solid(); sh.fill.fore_color.rgb = RGBColor.from_string(fill)
    if line:
        sh.line.color.rgb = RGBColor.from_string(line); sh.line.width = Pt(lw)
    else:
        sh.line.fill.background()
    tf = sh.text_frame
    tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Pt(7); tf.margin_top = tf.margin_bottom = Pt(3)
    fill_paras(tf, paras, align)
    hang(tf)
    return sh


def text(x, y, wd, h, paras, align=PP_ALIGN.LEFT):
    tb = s.shapes.add_textbox(Pt(x), Pt(y), Pt(wd), Pt(h))
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    fill_paras(tf, paras, align)
    return tb


def arrow(x1, y1, x2, y2, head=True):
    c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Pt(x1), Pt(y1), Pt(x2), Pt(y2))
    strip(c)
    c.line.color.rgb = RGBColor.from_string(MUTED); c.line.width = Pt(1)
    if head:
        ln = c._element.find('.//' + qn('a:ln'))
        etree.SubElement(ln, qn('a:tailEnd'), type='triangle', w='med', len='med')
    return c


P = lambda t, size=9, bold=True, color=WHITE: (t, 'Play', size, bold, color)
A = lambda t, size=8.5, bold=False, color=INK: (t, 'Arial', size, bold, color)
B = lambda t, size=8.5: [('•  ', 'Arial', size, True, TEAL), A(t, size)]
B2 = lambda t, size=8.5: [('–  ', 'Arial', size, False, MUTED), A(t, size)]


A = lambda t, size=8.5, bold=False, color=INK: (t, 'Arial', size, bold, color)
B = lambda t, size=8.5: [('•  ', 'Arial', size, True, TEAL), A(t, size)]
B2 = lambda t, size=8.5: [('–  ', 'Arial', size, False, MUTED), A(t, size)]

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


def apply_timing(pop):
    ids = [sh.shape_id for sh in pop]
    cid = 5
    c1 = ''
    for i, spid in enumerate(ids):
        c1 += effect(cid, spid, 'clickEffect' if i == 0 else 'withEffect'); cid += 3
    start2 = cid; cid += 2
    c2 = ''
    for i, spid in enumerate(ids):
        c2 += leave(cid, spid, 'clickEffect' if i == 0 else 'withEffect'); cid += 2
    bld = ''.join(f'<p:bldP spid="{i}" grpId="0"/><p:bldP spid="{i}" grpId="1"/>' for i in ids
                  if s.shapes._spTree.xpath(f'.//p:sp[p:nvSpPr/p:cNvPr/@id="{i}"]'))
    NS = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    tim = ('<p:timing ' + NS + '><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
           '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
           + click(3, c1) + click(start2, c2) +
           '</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
           '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>'
           '</p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst>' + bld + '</p:bldLst></p:timing>')
    old = s._element.find(qn('p:timing'))
    if old is not None:
        s._element.remove(old)
    t_el = etree.fromstring(tim)
    ext = s._element.find(qn('p:extLst'))
    (ext.addprevious(t_el) if ext is not None else s._element.append(t_el))



def add_note(slide, text_):
    nt = slide.notes_slide.notes_text_frame
    if text_[:40] in nt.text:
        return
    p = copy.deepcopy(nt.paragraphs[-1]._p)
    for r in p.findall(qn('a:r'))[1:]:
        p.remove(r)
    p.findall(qn('a:r'))[0].find(qn('a:t')).text = text_
    nt.paragraphs[-1]._p.addnext(p)


# ======================= slide 31: overview =======================
s = S[30]
clear(s)
set_text(s.shapes[1], 'Scenario 5: Multiple Unused MFs Due to Inactivity, or "Forgot" About FC', 15)
set_text(s.shapes[2], 'C has been paying MFs for several months without using the service.')
box(32, 102, 400, 70, TEAL_PALE, TEAL, [
    [P('What this looks like', 10, True, TEAL)],
    [A('C forgot about the membership, believed cancelling their recurring cleaning plan also cancelled FC, or simply went inactive.', 8.5)],
    [A('The goal: make sure they get real value from what they paid for, while keeping the door open to retain them.', 8.5, True, INK)],
], anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT)
box(444, 102, 243, 70, AMBER_PALE, AMBER, [
    [P('Scope check', 10, True, AMBER)],
    [A('Not the same as an FCF customer who has never had a completed job. This scenario assumes C used the service at some point, then went quiet.', 8.5)],
], anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT)
box(32, 182, 655, 66, WHITE, LINE, [
    [P('Agent guidance:  probe first', 10, True, TEAL)],
    [A('Before doing anything else, ask why C wants to cancel or why they haven\'t used the service. Check for an unresolved issue keeping them from getting value out of the membership.', 8.5)],
    [A('"Were there any issues with your cleaners? I just want to check if anything was unsatisfactory, or if there\'s anything I can do. We\'d love to keep you and make sure your experience goes well."', 8.5, False, MUTED)],
], anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT)
for p in s.shapes[-1].text_frame.paragraphs[2:3]:
    for r in p.runs:
        r.font.italic = True
cols = [('Within MCT', 'next slide', TEAL), ('Beyond MCT', 'slide after', TEAL), ('Refund requests', 'after long inactivity', AMBER)]
for i, (a, b, c) in enumerate(cols):
    x = 32 + i * 222
    box(x, 258, 211, 26, c, None, [[P(a + '  ›  ', 9, True), P(b, 8, False, TEAL_PALE if c == TEAL else AMBER_PALE)]])
add_note(s, 'Scenario 5 overview. The customer has been paying membership fees for several months without using the service, either because they forgot about the membership, believed cancelling their recurring cleaning plan also cancelled FC, or simply went inactive. '
            'The goal is to make sure they get real value from what they paid for, while keeping the door open to retain them. '
            'Scope check: this is not the same as an FCF customer who has never had a completed job; this scenario assumes the customer used the service at some point and then went quiet. '
            'Agent guidance: probe first why the customer wants to cancel or why they haven\'t used the service, before doing anything else. It is worth checking whether an unresolved issue is keeping them from getting value out of the membership. '
            'Example: "Were there any issues with your cleaners? I just want to check if anything was unsatisfactory, or if there\'s anything I can do. We\'d love to keep you and make sure your experience goes well."')

# ======================= new slide 32: within MCT =======================
ref = S[30]
s = prs.slides.add_slide(ref.slide_layout)
for ph in list(s.placeholders):
    ph._element.getparent().remove(ph._element)
num = [x for x in ref.shapes if x.is_placeholder][0]
for el in (ref.shapes[0]._element, ref.shapes[1]._element, ref.shapes[2]._element, num._element):
    s.shapes._spTree.append(copy.deepcopy(el))
lst = prs.slides._sldIdLst; el = list(lst)[-1]; lst.remove(el); list(lst)[30].addnext(el)
set_text(s.shapes[1], 'Scenario 5: Within MCT', 28)
set_text(s.shapes[2], 'Multiple unused MFs due to inactivity, or "forgot" about FC. The ETF applies.')
CX = 400
box(CX - 200, 99, 400, 22, TEAL, None,
    [[P('FC retention – Long inactivity  |  Forgot about the FC', 8.5)], [P('(ETF applies)', 7.5, False, TEAL_PALE)]])
arrow(CX, 121, CX, 127)
box(CX - 60, 127, 120, 18, AMBER_PALE, AMBER, [[P('Root cause?', 9, True, AMBER)]])
arrow(CX - 60, 136, 230, 136)
box(32, 120, 198, 66, WHITE, LINE, [
    [P('Follow the appropriate playbook for:', 8, True, TEAL)],
    B('PCQ', 7.5), B('Unaware of the FC', 7.5), B('Service Reliability', 7.5), B('2 consecutive issue jobs', 7.5)],
    align=PP_ALIGN.LEFT)
arrow(CX, 145, CX, 153)
box(CX - 150, 153, 300, 60, TEAL_PALE, TEAL, [
    [P('Genuinely inactive, or forgot about their FC', 8.5, True, TEAL)],
    B('Explain FC terms (MF + ETF)', 8), B('Offer free months (max 3)', 8),
    B('Apply free weeks/months to move the billing date as a one-time courtesy', 8)], align=PP_ALIGN.LEFT)
arrow(CX, 213, CX, 220)
box(CX - 130, 220, 260, 20, INK, None, [[P('C rejects the resolution and insists on cancelling the FC', 8, False)]])
arrow(CX, 240, CX, 247)
box(CX - 100, 247, 200, 20, AMBER_PALE, AMBER, [[P('Any hint they\'d pay a reduced ETF?', 8.5, True, AMBER)]])
arrow(CX + 100, 257, 535, 257)
text(CX + 106, 246, 30, 9, [[P('YES', 7.5, True, INK)]])
box(535, 244, 152, 26, TEAL, None, [[P('EXIT Strategy', 9)], [P('ETF reduction, max 50%', 7.5, False, TEAL_PALE)]])
arrow(611, 270, 611, 278)
box(535, 278, 152, 22, INK, None, [[P('C declines, insists on cancelling', 7.5, False)]])
arrow(535, 289, CX + 70, 289)
arrow(CX, 267, CX, 278)
text(CX + 5, 267, 30, 9, [[P('NO', 7.5, True, INK)]])
box(CX - 70, 278, 140, 22, AMBER_PALE, AMBER, [[P('GRACEFUL CLOSURE', 9, True, AMBER)]])
box(32, 312, 655, 24, WHITE, LINE, [[P('Probe first:  ', 8.5, True, TEAL),
    A('confirm the real reason before anything else. Only genuine inactivity or forgetting stays in this flow.', 8)]], align=PP_ALIGN.LEFT)
s.notes_slide.notes_text_frame.text = (
    'Flow (within MCT, ETF applies). Find the root cause first. If it is PCQ, unawareness of the FC, Service Reliability, or two consecutive issue jobs, follow that playbook instead. '
    'If they are genuinely inactive or forgot about their FC: explain the FC terms (MF and ETF), offer free months (max 3), and move the billing date with free weeks/months as a one-time courtesy. '
    'If C rejects and insists on cancelling, check for any hint they would pay a reduced ETF: if yes, Exit Strategy with ETF reduction up to 50%; if they decline that too, or there is no hint, Graceful Closure.')

# ======================= new slide 33: beyond MCT =======================
ref = S[30]
s = prs.slides.add_slide(ref.slide_layout)
for ph in list(s.placeholders):
    ph._element.getparent().remove(ph._element)
num = [x for x in ref.shapes if x.is_placeholder][0]
for el in (ref.shapes[0]._element, ref.shapes[1]._element, ref.shapes[2]._element, num._element):
    s.shapes._spTree.append(copy.deepcopy(el))
lst = prs.slides._sldIdLst; el = list(lst)[-1]; lst.remove(el); list(lst)[31].addnext(el)
set_text(s.shapes[1], 'Scenario 5: Beyond MCT', 28)
set_text(s.shapes[2], 'Multiple unused MFs due to inactivity, or "forgot" about FC. The ETF does not apply.')
s.notes_slide.notes_text_frame.text = (
    'Flow (beyond MCT, ETF does not apply): is the customer still retainable? If it is not reasonable to retain, for example only a few jobs against many paid months, address any pain points, cancel the FC, and refund the MFs if C has requested a refund and is eligible (see the next slide). '
    'If retainable, probe for any underlying issues, offer up to 3 free months for paid but unused MFs (not more than asked), explicitly confirm no further MF will be charged unless they agree to continue, and set a TR. '
    'If C rejects the resolution and insists on cancelling, cancel the FC.')
L, LW = 32, 400
CX = L + LW / 2
R, RW = 456, 231
RX = R + RW / 2

box(L, 98, LW, 22, TEAL, None,
    [[P('FC retention – Long inactivity  |  Forgot about the FC', 8.5)], [P('(ETF does not apply)', 8, False, TEAL_PALE)]])
arrow(CX, 120, CX, 127)
box(CX - 110, 127, 220, 26, AMBER_PALE, AMBER,
    [[P('Is the C still retainable?', 9.5, True, AMBER)],
     [P('Use the Five Key Factors', 7.5, False, AMBER)]])
arrow(CX + 110, 140, R, 140)
text(CX + 120, 129, 30, 10, [[P('NO', 8, True, INK)]])
box(R, 118, RW, 44, TEAL_PALE, TEAL,
    [[A('Address any pain points, cancel the FC, and refund the MFs if C has requested a refund and is eligible.', 8.5, False, TEAL)]])
box(R, 170, RW, 44, WHITE, LINE,
    [[P('Not reasonable to retain?  ', 8.5, True, TEAL),
      A('e.g. only a few jobs against many paid months: cancel.', 8)]],
    align=PP_ALIGN.LEFT)

box(R, 222, RW, 40, AMBER_PALE, AMBER,
    [[P('Refund request?  ', 8.5, True, AMBER), A('Only MFs from the last 120 days that are beyond MCT. See the next slide.', 8)]],
    align=PP_ALIGN.LEFT)
arrow(CX, 153, CX, 163)
text(CX + 5, 153, 30, 10, [[P('YES', 8, True, INK)]])
box(L, 163, LW, 88, TEAL_PALE, TEAL, [
    B('Probe for any underlying issues'),
    B('Offer free months for unused paid MFs (max 3, not more than asked)'),
    B('Explicitly confirm no further MF will be charged unless they agree to continue'),
    B('Set a TR 3 days before the next MF date. If the next MF is within 5 days, issue a free month first'),
], anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.LEFT)

# split into two branches
LB, RB = L + 100, L + LW - 100
arrow(CX, 251, CX, 260, head=False)
arrow(LB, 260, RB, 260, head=False)
arrow(LB, 260, LB, 269)
arrow(RB, 260, RB, 269)
box(LB - 92, 269, 184, 26, INK, None, [[P('C rejects the resolution', 9, False)]])
arrow(LB, 295, LB, 304)
box(LB - 70, 304, 140, 24, AMBER_PALE, AMBER, [[P('Cancel the FC', 9.5, True, AMBER)]])
box(RB - 92, 269, 184, 26, INK, None, [[P('No reply, but the TR was triggered', 9, False)]])
arrow(RB, 295, RB, 304)
box(RB - 70, 304, 140, 24, TEAL, None, [[P('What to do next  ▸', 9, True)]])
text(RB - 92, 331, 184, 10, [[A('click to see both outcomes', 7, False, MUTED)]], PP_ALIGN.CENTER)

# ---- pop-up: when the TR is triggered (from slide 8) ----
pop = []
card = box(110, 102, 500, 218, WHITE, TEAL, [[A('')]], lw=1.5)
sp = card._element.find(qn('p:spPr'))
eff = etree.SubElement(sp, qn('a:effectLst'))
sh_ = etree.SubElement(eff, qn('a:outerShdw'), blurRad='76200', dist='25400', dir='5400000', algn='t', rotWithShape='0')
clr = etree.SubElement(sh_, qn('a:srgbClr'), val='000000'); etree.SubElement(clr, qn('a:alpha'), val='25000')
pop.append(card)
pop.append(box(250, 114, 220, 30, AMBER_PALE, AMBER, [[P('When the TR is triggered', 11, True, AMBER)]]))
PL, PR = 245, 475
pop.append(arrow(360, 144, 360, 154, head=False))
pop.append(arrow(PL, 154, PR, 154, head=False))
pop.append(arrow(PL, 154, PL, 164))
pop.append(arrow(PR, 154, PR, 164))
pop.append(box(PL - 110, 164, 220, 30, AMBER, None, [[P('No reply and no new bookings', 9.5)]]))
pop.append(box(PL - 110, 200, 220, 84, AMBER_PALE, AMBER,
               [[A('Cancel the FC.', 9, True)],
                [A("Send comms only if there's an unresolved pain point the auto comms wouldn't cover.", 9)]]))
pop.append(box(PR - 110, 164, 220, 30, TEAL, None, [[P('No reply, but C is using the service', 9.5)]]))
pop.append(box(PR - 110, 200, 220, 84, TEAL_PALE, TEAL, [
    B("Don't cancel the FC", 9), B('Contact C to confirm the cancellation', 9),
    B('Add 1–2 weeks to move the due date, and set a new TR', 9)], align=PP_ALIGN.LEFT))
pop.append(text(130, 296, 460, 12, [[A('Click to close', 7.5, False, MUTED)]], PP_ALIGN.RIGHT))




apply_timing(pop)
add_note(s, 'TRAINER: click once to open the "When the TR is triggered" pop-up, click again to close it. ' + s8_notes)

# ======================= new slide 34: refund requests =======================
ref = S[30]
s = prs.slides.add_slide(ref.slide_layout)
for ph in list(s.placeholders):
    ph._element.getparent().remove(ph._element)
num = [x for x in ref.shapes if x.is_placeholder][0]
for el in (ref.shapes[0]._element, ref.shapes[1]._element, ref.shapes[2]._element, num._element):
    s.shapes._spTree.append(copy.deepcopy(el))
lst = prs.slides._sldIdLst; el = list(lst)[-1]; lst.remove(el); list(lst)[32].addnext(el)
set_text(s.shapes[1], 'Refund Requests After Long Inactivity', 28)
set_text(s.shapes[2], 'When C pushes back on "donating" money because they were unaware of the FC or believed they had already cancelled it.', 11)
box(32, 104, 655, 34, TEAL, None, [[P('Only the MFs charged in the last 120 days that are beyond C\'s MCT can be refunded.', 11)]])
box(32, 146, 250, 30, AMBER_PALE, AMBER, [[P('Formula:  ', 9, True, AMBER), P('refund = paid months − C\'s MCT', 9, True, INK)]])
text(292, 146, 395, 30, [[A('...and only the charges from the last 120 days count. Not for a customer who simply changed their mind about wanting the service.', 8, False, MUTED)]])
text(32, 186, 655, 14, [[P('EXAMPLE:  12 PAID MONTHS AGAINST A 6-MONTH MCT', 8.5, True, AMBER)]])
G = 4; W = (655 - 11 * G) / 12
for i in range(12):
    x = 32 + i * (W + G)
    if i < 6:
        fill, line, col, lab = 'F5F6F8', LINE, MUTED, 'MCT'
    elif i < 8:
        fill, line, col, lab = AMBER_PALE, AMBER, AMBER, '>120 days'
    else:
        fill, line, col, lab = TEAL_PALE, TEAL, TEAL, 'refund'
    box(x, 204, W, 34, fill, line, [[P('M' + str(i + 1), 9, True, col)], [A(lab, 6.5, False, col)]])
text(32, 244, 655, 30, [[A('Six months sit beyond the MCT, but only the charges from the last 120 days are refundable, which is about four of them. ', 8.5),
                         P('The refund is 4 months, not 6.', 9, True, TEAL)]])
box(32, 278, 655, 30, AMBER_PALE, AMBER, [[P('Don\'t steer:  ', 9, True, AMBER),
    A('if C explicitly says they don\'t want a refund and just wants to cancel, don\'t steer them toward one. Proceed with the cancellation.', 8.5)]],
    align=PP_ALIGN.LEFT)
s.notes_slide.notes_text_frame.text = (
    'Refund requests after long inactivity. If the customer pushes back on "donating" money, specifically because they were unaware of the membership or believed they had already cancelled it (not because they simply changed their mind about wanting the service), '
    'only the MFs charged in the last 120 days that are beyond the customer\'s MCT can be refunded. Formula: refund = paid months minus the customer\'s MCT. '
    'Example: 12 paid months against a 6-month MCT. Six months sit beyond the MCT, but only the charges from the last 120 days are refundable, which is about four of them. The refund is 4 months, not 6. '
    'If the customer explicitly says they don\'t want a refund and just wants to cancel, don\'t steer them toward one; proceed with the cancellation per the branches on the previous slides.')

prs.save(sys.argv[2]); print('saved', sys.argv[2])
