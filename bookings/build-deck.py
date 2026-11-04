"""Build the "Bookings" training deck in the New Hire Care Training theme.
Template: an export of the Day 8 deck (cover, topic divider, content header and Questions slides
are reused for the theme). Content: the Care Knowledge Library "Bookings" section, eight articles:
Appointments and Scheduling, Rescheduling, Cancellation (Job-Level), Pairing (C/CP Blocking),
Priority Booking, No Claim, Cleaner Swap, Common Questions - Pre-Booking & Service FAQs.
Each topic gets a "TOPIC 0N OF 08" divider so the Training Agenda Library lists it.

Usage: python3 build-deck.py day8-template.pptx out.pptx
"""
import copy, sys
from lxml import etree
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

C = {'TEAL': '0B7A6F', 'TEAL_PALE': 'E3F4F1', 'AMBER': '9A6410', 'AMBER_PALE': 'FFF3D6',
     'INK': '1D1D1F', 'MUTED': '6E6E73', 'WHITE': 'FFFFFF', 'LINE': 'DCE1E7', 'GREY': '6E6E73', 'ALT': 'F5F6F8'}
X, W = 32.4, 655.2

prs = Presentation(sys.argv[1])
txt = lambda sh: sh.text_frame.text if sh.has_text_frame else ''
S = list(prs.slides)
lst = prs.slides._sldIdLst
el_of = {id(s): e for s, e in zip(S, list(lst))}
cover = S[0]
divider = next(s for s in S if txt(s.shapes[0]).startswith('TOPIC') and len(s.shapes) > 20)
questions = next(s for s in S if txt(s.shapes[1]) == 'Questions?')
content_ref = next(s for s in S if txt(s.shapes[0]) == 'ASSESSING DISPUTE LIKELIHOOD' and txt(s.shapes[1]) == 'What is this?')
keep = {id(cover), id(divider), id(questions)}


def set_runs(shape, t):
    tf = shape.text_frame
    p0 = tf.paragraphs[0]
    lines = t.split('\n')
    for p in tf.paragraphs[1:]:
        p._p.getparent().remove(p._p)
    for r in p0.runs[1:]:
        r._r.getparent().remove(r._r)
    p0.runs[0].text = lines[0]
    for ln in lines[1:]:
        np_ = copy.deepcopy(p0._p); p0._p.getparent().append(np_)
        rr = np_.findall(qn('a:r'))
        rr[0].find(qn('a:t')).text = ln
    return shape



def new_slide(eyebrow, title, sub):
    s = prs.slides.add_slide(content_ref.slide_layout)
    for ph in list(s.placeholders):
        ph._element.getparent().remove(ph._element)
    num = [x for x in content_ref.shapes if x.is_placeholder][0]
    for e in (content_ref.shapes[0]._element, content_ref.shapes[1]._element, content_ref.shapes[2]._element, num._element):
        s.shapes._spTree.append(copy.deepcopy(e))
    set_runs(s.shapes[0], eyebrow); set_runs(s.shapes[1], title); set_runs(s.shapes[2], sub)
    order.append(s)
    return s


def strip(sh):
    st = sh._element.find(qn('p:style'))
    if st is not None:
        sh._element.remove(st)


def fill_tf(tf, paras, anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT, after=2):
    tf.word_wrap = True; tf.vertical_anchor = anchor
    for i, item in enumerate(paras):
        num, runs = item if isinstance(item, tuple) else ('', item)
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(after); p.alignment = align
        if num:
            pPr = p._p.get_or_add_pPr()
            pPr.set('marL', str(int(12 * 12700))); pPr.set('indent', str(-10 * 12700))
            runs = [(num + ' ', runs[0][1], True, 'TEAL' if num != '•' else 'MUTED')] + runs
        for t, size, bold, color in runs:
            r = p.add_run(); r.text = t; f = r.font
            f.name = 'Play' if bold else 'Arial'; f.size = Pt(size); f.bold = bold
            f.color.rgb = RGBColor.from_string(C[color])


def box(s, x, y, w, h, bg, line=None, paras=(), align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE, margin=9,
        shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=6, lw=0.75):
    sh = s.shapes.add_shape(shape, Pt(x), Pt(y), Pt(w), Pt(h)); strip(sh)
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        sh.adjustments[0] = min(0.5, radius / min(w, h))
    sh.fill.solid(); sh.fill.fore_color.rgb = RGBColor.from_string(C[bg])
    if line:
        sh.line.color.rgb = RGBColor.from_string(C[line]); sh.line.width = Pt(lw)
    else:
        sh.line.fill.background()
    tf = sh.text_frame
    tf.margin_left = tf.margin_right = Pt(margin); tf.margin_top = tf.margin_bottom = Pt(3)
    fill_tf(tf, paras, anchor, align)
    return sh


def text(s, x, y, w, h, paras, anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT, after=2, name=None):
    tb = s.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(h))
    tf = tb.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    fill_tf(tf, paras, anchor, align, after)
    if name:
        tb.name = name
    return tb


def arrow(s, x1, y1, x2, y2, head=True):
    c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Pt(x1), Pt(y1), Pt(x2), Pt(y2)); strip(c)
    c.line.color.rgb = RGBColor.from_string(C['GREY']); c.line.width = Pt(1)
    if head:
        ln = c._element.find('.//' + qn('a:ln'))
        etree.SubElement(ln, qn('a:tailEnd'), type='triangle', w='med', len='med')


def dot(s, x, y, label, bg='TEAL', d=20, sz=9):
    return box(s, x, y, d, d, bg, paras=[[(label, sz, True, 'WHITE')]], align=PP_ALIGN.CENTER, margin=0, shape=MSO_SHAPE.OVAL)


B = lambda t, c='INK', sz=8.5: (t, sz, True, c)
A = lambda t, c='INK', sz=8.5: (t, sz, False, c)


def tile(s, x, y, w, h, n, head, body, col='TEAL', pale='TEAL_PALE', sz=8.5):
    box(s, x, y, w, h, 'WHITE', 'LINE')
    dot(s, x + 10, y + 10, n, col, 22)
    text(s, x + 40, y + 12, w - 50, 18, [[B(head, col, 10.5)]])
    text(s, x + 12, y + 40, w - 24, h - 46, body, after=3)



def reveal_timing(s, groups):
    """One click per group: fade every shape in the group in together."""
    def effect(cid, spid, node):
        tgt = f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
        return (f'<p:par><p:cTn id="{cid}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="{node}">'
                f'<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
                f'<p:set><p:cBhvr><p:cTn id="{cid+1}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>{tgt}'
                f'<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
                f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{cid+2}" dur="400"/>{tgt}</p:cBhvr></p:animEffect>'
                f'</p:childTnLst></p:cTn></p:par>')
    cid, clicks = 3, ''
    for ids in groups:
        start, inner = cid, ''
        cid += 2
        for i, spid in enumerate(ids):
            inner += effect(cid, spid, 'clickEffect' if i == 0 else 'withEffect'); cid += 3
        clicks += (f'<p:par><p:cTn id="{start}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>'
                   f'<p:par><p:cTn id="{start+1}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
                   f'{inner}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>')
    bld = ''.join(f'<p:bldP spid="{i}" grpId="0" animBg="1"/>' for g in groups for i in g)
    NS = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    tim = ('<p:timing ' + NS + '><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
           '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>' + clicks +
           '</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
           '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>'
           '</p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst>' + bld + '</p:bldLst></p:timing>')
    ext = s._element.find(qn('p:extLst'))
    t_el = etree.fromstring(tim)
    (ext.addprevious(t_el) if ext is not None else s._element.append(t_el))



def dup_slide(src):
    """Copy a slide (shapes, pictures and links) onto a new slide with the same layout."""
    s = prs.slides.add_slide(src.slide_layout)
    for ph in list(s.placeholders):
        ph._element.getparent().remove(ph._element)
    for sh in src.shapes:
        el = copy.deepcopy(sh._element)
        for node in el.iter():
            for attr in (qn('r:embed'), qn('r:link'), qn('r:id')):
                rid = node.get(attr)
                if rid and rid in src.part.rels:
                    rel = src.part.rels[rid]
                    new = (s.part.relate_to(rel.target_ref, rel.reltype, is_external=True) if rel.is_external
                           else s.part.relate_to(rel.target_part, rel.reltype))
                    node.set(attr, new)
        s.shapes._spTree.append(el)
    return s


N_TOPICS = 8
dividers = []


def topic_divider(n, eyebrow_area, title, desc, learn, notes):
    s = divider if n == 1 else dup_slide(divider)
    set_runs(s.shapes[0], f'TOPIC {n:02d} OF {N_TOPICS:02d}')
    set_runs(s.shapes[1], title)
    set_runs(s.shapes[2], desc)
    for k, t in zip((6, 9, 12, 15), learn):
        set_runs(s.shapes[k], t)
    s.notes_slide.notes_text_frame.text = notes
    order.append(s)
    dividers.append(s)
    return s


def kc(area, qa, notes):
    s = new_slide(area + '  ·  KNOWLEDGE CHECK', 'Let’s See What You’ve Got!', 'Ask each question, then click to reveal the answer.')
    n = len(qa); gap = 6; h = min(56, (240 - gap * (n - 1)) / n)
    groups = []
    for i, (q, a) in enumerate(qa):
        y = 100 + i * (h + gap)
        box(s, X, y, 364, h, 'WHITE', 'LINE')
        box(s, X + 10, y + (h - 20) / 2, 20, 20, 'TEAL_PALE', paras=[[(str(i + 1), 9, True, 'TEAL')]], align=PP_ALIGN.CENTER, margin=0, shape=MSO_SHAPE.OVAL)
        text(s, X + 40, y, 316, h, [[A(q, 'INK', 8.5)]], anchor=MSO_ANCHOR.MIDDLE)
        b = box(s, X + 371, y, W - 371, h, 'TEAL_PALE')
        t = text(s, X + 381, y, W - 391, h, [[A(a, 'INK', 8.5)]], anchor=MSO_ANCHOR.MIDDLE)
        groups.append([b.shape_id, t.shape_id])
    reveal_timing(s, groups)
    s.notes_slide.notes_text_frame.text = notes
    return s


def panel(s, x, y, w, h, head, paras, col='TEAL', bg='WHITE', line='LINE', after=3):
    box(s, x, y, w, h, bg, line, anchor=MSO_ANCHOR.TOP)
    text(s, x + 12, y + 10, w - 24, h - 16, [[B(head, col, 10)]] + list(paras), after=after)


def table(s, x, y, widths, heights, rows, head_bg='TEAL', zebra=None):
    gt = s.shapes.add_table(len(rows), len(widths), Pt(x), Pt(y), Pt(sum(widths)), Pt(sum(heights)))
    tbl = gt.table
    for j, w_ in enumerate(widths):
        tbl.columns[j].width = Pt(w_)
    tblPr = tbl._tbl.tblPr
    for a in ('bandRow', 'firstRow'):
        tblPr.set(a, '0')
    sid = tblPr.find(qn('a:tableStyleId'))
    if sid is not None:
        tblPr.remove(sid)
    for i, row in enumerate(rows):
        tbl.rows[i].height = Pt(heights[i])
        for j, v in enumerate(row):
            c = tbl.cell(i, j)
            c.margin_left = c.margin_right = Pt(8); c.margin_top = c.margin_bottom = Pt(3)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            bg = head_bg if i == 0 else ((zebra or {}).get(i) or 'WHITE')
            c.fill.solid(); c.fill.fore_color.rgb = RGBColor.from_string(C[bg])
            tcPr = c._tc.get_or_add_tcPr()
            for tag in ('a:lnL', 'a:lnR', 'a:lnT', 'a:lnB'):
                ln = tcPr.makeelement(qn(tag), {'w': '9525'})
                sf = ln.makeelement(qn('a:solidFill'), {}); sf.append(sf.makeelement(qn('a:srgbClr'), {'val': C['LINE']})); ln.append(sf)
                tcPr.append(ln)
            fe = tcPr.find(qn('a:solidFill')); tcPr.remove(fe); tcPr.insert(0, fe)
            p = c.text_frame.paragraphs[0]; c.text_frame.word_wrap = True
            for r in list(p.runs):
                r._r.getparent().remove(r._r)
            runs = v if isinstance(v, list) else [(v, 9 if i == 0 else 8.5, i == 0 or j == 0, 'WHITE' if i == 0 else ('TEAL' if j == 0 else 'INK'))]
            for t_, sz, bold, col in runs:
                r = p.add_run(); r.text = t_
                r.font.size = Pt(sz); r.font.bold = bold; r.font.name = 'Play' if bold else 'Arial'
                r.font.color.rgb = RGBColor.from_string(C[col])
    return tbl


# ---------------- cover ----------------
set_runs(cover.shapes[0], 'HOMEAGLOW  ·  NEW HIRE CARE TRAINING')
set_runs(cover.shapes[1], 'Bookings')
set_runs(cover.shapes[2], 'Scheduling, rescheduling, cancelling, pairing, priority bookings, unclaimed jobs, cleaner swaps and common pre-booking questions.')
cover.notes_slide.notes_text_frame.text = ('Bookings. Source: the Care Knowledge Library, Bookings section (eight articles). '
                                           'Pairing and Priority Booking have moved to Notion; these slides follow the last Knowledge Library version.')
order = [cover]

# ---------------- overview / agenda ----------------
s_over = new_slide('OVERVIEW', 'Bookings', 'Eight topics on how jobs get booked, moved, matched and cancelled. Click a topic to jump to it.')
topics = [('Appointments and Scheduling', 'Cleaner requests, strict schedules, add-on hours.'),
          ('Rescheduling', 'Who can move a job, and when.'),
          ('Cancellation (Job-Level)', 'One job, not the membership.'),
          ('Pairing', 'Blocking a C/CP match.'),
          ('Priority Booking', 'A goodwill job you create for C.'),
          ('No Claim', 'No CP has claimed the job yet.'),
          ('Cleaner Swap', 'What you can and can\'t say.'),
          ('Common Questions', 'Pre-booking and service FAQs.')]
tw, th, tg = (W - 10) / 2, 52, 7
over_tiles = []
for i, (name, sub) in enumerate(topics):
    c_, r_ = divmod(i, 4)
    x, y = X + c_ * (tw + 10), 102 + r_ * (th + tg)
    shapes = [box(s_over, x, y, tw, th, 'WHITE', 'LINE')]
    shapes.append(text(s_over, x + 14, y, 30, th, [[B(f'{i + 1:02d}', 'TEAL', 14)]], anchor=MSO_ANCHOR.MIDDLE))
    shapes.append(text(s_over, x + 52, y + 8, tw - 80, th - 12, [[B(name, 'INK', 10.5)], [A(sub, 'MUTED', 8.5)]], anchor=MSO_ANCHOR.MIDDLE, after=1))
    shapes.append(text(s_over, x + tw - 22, y, 12, th, [[B('›', 'TEAL', 14)]], anchor=MSO_ANCHOR.MIDDLE))
    over_tiles.append(shapes)
s_over.notes_slide.notes_text_frame.text = (
    'Bookings covers the conversations around how a cleaning gets booked and what happens to it before it is done: requests for a specific cleaner or schedule, '
    'rescheduling, cancelling a single job, blocking a customer and cleaner from being matched again, creating a priority booking, jobs nobody has claimed, '
    'cleaner swaps, and the quick pre-booking questions customers ask. Click a topic to jump to it.')

# ================= 01 Appointments and Scheduling =================
EB = 'APPOINTMENTS AND SCHEDULING'
topic_divider(1, EB, 'Appointments and Scheduling', 'Requests about how or when a cleaning gets booked: a specific cleaner, a time window, a scheduling preference.',
              ('Protected-trait requests', 'A specific cleaner, every time', 'Strict schedules', 'Add-on hours and duration'),
              'These tickets look simple, but each request has a specific mechanical answer. Some have a real fix in the system (the strict-schedule flag); '
              'some have none (filtering by gender). Knowing which is which keeps you from overpromising on one and under-delivering on the other.')

s = new_slide(EB, 'Requests for a specific cleaner', 'Some requests have a fix in the system. Some don\'t, and we don\'t make one up.')
LW = 360
panel(s, X, 100, LW, 244, 'C asks for a cleaner by gender or another protected trait', [
    [A('There\'s no system filter for it, and '), B('we don\'t do the rematch'), A('.')],
    ('•', [B('Don\'t swap the CP'), A(', even if the job is already claimed.')]),
    ('•', [B('Don\'t add the preference'), A(' to cleaning or internal notes. If someone already did, remove it.')]),
    ('•', [A('Point C to their dashboard: '), B('Manage Appointment › Change My Cleaner'), A(', up to 6 hours before start.')]),
    ('•', [A('Under 6 hours: still don\'t swap. C can keep the appointment or cancel under the normal rules.')]),
    [A('Macro: Protected-Trait Preference – Change Your Own Cleaner (new CRM)', 'MUTED', 8)],
    [B('Why: ', 'AMBER', 8.5), A('if we use Change CP, the CP is told "Homeaglow customer service removed you". For a protected trait, that puts Homeaglow on record '
                                  'removing a CP for who they are. If C changes it, the CP just sees that C cancelled.', 'INK', 8.5)]],
      col='AMBER', bg='AMBER_PALE', line='AMBER')
RX, RW = X + LW + 10, W - LW - 10
panel(s, RX, 100, RW, 244, 'C asks for the same cleaner for future appointments', [
    [B('Only if C explicitly asks.'), A(' Never set it as a nice gesture.')],
    ('1', [A('Open the '), B('Job Admin Page')]),
    ('2', [A('Select the icon next to '), B('ID')]),
    ('3', [A('Enter the cleaner\'s CP ID in '), B('Requested CP')]),
    ('4', [A('Untick '), B('Can Substitute')]),
    ('5', [A('Scroll down and click '), B('SAVE')]),
    [A('See CP Tiering for how Requested CP affects claim eligibility and swap protection.', 'MUTED', 8)]],
      bg='TEAL_PALE', line=None, after=4)
s.notes_slide.notes_text_frame.text = (
    'Requests for a specific cleaner. If a customer asks for a cleaner by gender, race or another protected trait: there is no system option to filter by it, and we do not do the rematch. '
    'Point the customer to their own rematch, even when the job is already claimed. The reason: the CP gets a message either way, but a different one. When the customer changes '
    'their cleaner from the dashboard, the CP sees "<Customer> canceled their appointment for <date>." When we use Change CP, the CP sees "Homeaglow customer service removed you from '
    '<Customer>\'s <date> appointment." For a request based on a protected trait, that second message puts Homeaglow on record as removing the CP for who they are.\n'
    'So: do not swap the CP, and do not add the preference to the cleaning notes or internal notes (remove it if another agent added it). Explain briefly that we do not match cleaners '
    'on that trait and point them to Manage Appointment > Change My Cleaner, up to 6 hours before start. Under 6 hours, still do not swap; they can keep the appointment or cancel under '
    'the normal cancellation rules. Use the Protected-Trait Preference – Change Your Own Cleaner macro in the new CRM.\n'
    'A specific cleaner for future appointments: only change the Requested CP if the customer explicitly asks, never proactively. Job Admin Page, select the icon next to ID, enter '
    'the cleaner\'s CP ID in Requested CP, untick Can Substitute, scroll down and click SAVE. See CP Tiering for how Requested CP interacts with claim eligibility and swap protection.')

s = new_slide(EB, 'Strict schedules and add-on hours', 'One flag stops unwanted alternate times. Each add-on adds one hour.')
panel(s, X, 100, 260, 244, 'C says only certain days or times work', [
    [A('By default, a job still '), B('unclaimed 24 hours before start'), A(' gets alternate start times added automatically. That can book C into a slot they never wanted.')],
    [A('Add the flag so it stops for their future requests:')],
    [B('C CRM › Do › do_not_auto_add_alternate_start_times', 'TEAL', 8.5)]], bg='TEAL_PALE', line=None, after=6)
RX, RW = X + 270, W - 270
panel(s, RX, 100, RW, 112, 'C asks about the extra time', [
    [A('Each add-on adds '), B('1 hour'), A(': inside of oven, inside of refrigerator, fold/wash laundry, interior windows. It isn\'t a bug or an upsell.')],
    [A('Bookings run '), B('2 to 10 hours'), A('. Add-ons stack inside that range.')],
    [B('C wants a different duration? '), A('Submitted: update it in the C Dashboard. Claimed: tell the CP that C is asking to change the duration.')]])
panel(s, RX, 220, RW, 124, 'C pushes back ("I booked 6 hours, why add 1 more for the oven?")', [
    ('1', [A('Check which add-ons are on this booking, and that each is one of the four.')]),
    ('2', [A('Explain the extra time is '), B('for the add-on'), A(', not an increase to their base cleaning.')]),
    ('3', [A('Don\'t want the time? C can '), B('remove the add-on'), A(' instead.')]),
    ('4', [A('Over 10 hours? Remove an add-on or reduce the base hours.')])], col='AMBER', bg='AMBER_PALE', line=None)
s.notes_slide.notes_text_frame.text = (
    'Strict schedules: by default, the system auto-adds alternate start times to a job that is still unclaimed 24 hours before it starts. That can create a booking the customer never '
    'wanted if their availability is genuinely fixed. If a customer tells you only certain days or times work, add the flag so the system stops doing that for their future requests: '
    'C CRM > Do > do_not_auto_add_alternate_start_times.\n'
    'Hours and duration: when a customer selects an add-on, the system prompts them to add the time needed, and each add-on adds 1 hour: clean inside of oven, inside of refrigerator, '
    'fold/wash laundry, interior windows. It is not a bug or an upsell. Minimum booking is 2 hours, maximum 10; add-ons stack within that range. To change a job\'s duration: if it is '
    'Submitted, update it via the C Dashboard; if Claimed, tell the CP their client is asking to change the duration.\n'
    'If the customer pushes back: check which add-ons they selected and confirm each is one of the four (a flat +1 hour, no exceptions); explain the extra time is for the add-on, not '
    'their base cleaning; if they do not want the time, they can remove the add-on; if an add-on would push the total past 10 hours, they need to remove an add-on or reduce the base duration.')

kc('APPOINTMENTS AND SCHEDULING', [
    ('C asks for a female cleaner. The job is already claimed. Do you swap the CP?', 'No. Point C to Manage Appointment › Change My Cleaner (up to 6 hours before). Don\'t add it to notes.'),
    ('C loves their cleaner but hasn\'t asked to keep them. Set them as Requested CP?', 'No. Only change Requested CP when C explicitly asks.'),
    ('C says only Saturdays work for them. What do you add?', 'The do_not_auto_add_alternate_start_times flag (C CRM › Do).'),
    ('C booked 9 hours, then adds the oven and fridge add-ons. What happens?', 'That\'s 11 hours, over the 10-hour max. Remove an add-on or reduce the base hours.')],
   'Knowledge check: Appointments and Scheduling. Each click reveals the next answer. 1) Protected trait: never swap, point to Change My Cleaner, keep it out of notes. '
   '2) Requested CP only on an explicit ask. 3) Strict schedule flag. 4) Bookings max out at 10 hours.')

# ================= 02 Rescheduling =================
EB = 'RESCHEDULING'
topic_divider(2, EB, 'Rescheduling', 'Moving an already-claimed appointment to a different date or time, with everyone\'s agreement.',
              ('When C can reschedule', 'When the CP can reschedule', 'Inside the 6-hour window', '"Why was it rescheduled?"'),
              'Rescheduling mistakes usually look like one of two things: telling a customer they cannot reschedule when there is still a path through the CP, or treating a legitimate '
              'reschedule like an unauthorized one and investigating a CP who did nothing wrong. Know which flow you are in.')

s = new_slide(EB, 'Who is rescheduling?', 'The customer\'s window closes at 6 hours. The CP\'s never does, as long as C agrees.')
cw = (W - 20) / 3
panel(s, X, 100, cw, 196, 'C wants to reschedule', [
    ('•', [A('From the dashboard, '), B('6 hours or more'), A(' before start.')]),
    ('•', [A('New slot under 48 hours away? A '), B('Priority Fee'), A(' may apply.')]),
    ('•', [A('Don\'t apply the 12-hour new-booking rule to a reschedule.')]),
    [B('Under 6 hours: ', 'AMBER'), A('the dashboard option is gone. Ask C to contact their CP, who can reschedule for them.')]], bg='TEAL_PALE', line=None)
panel(s, X + cw + 10, 100, cw, 196, 'The CP wants to reschedule', [
    ('•', [A('CPs are encouraged to '), B('reschedule rather than cancel'), A('.')]),
    ('•', [B('No fixed time limit'), A(' as long as C agrees, even after the original start time.')]),
    ('•', [A('Same when the CP reschedules for C inside the 6-hour window.')]),
    [B('No agreement? ', 'AMBER'), A('The CP cancels from their dashboard. See Cancellation (Job-Level).')]], bg='TEAL_PALE', line=None)
panel(s, X + 2 * (cw + 10), 100, cw, 196, '"Why was my cleaning rescheduled?"', [
    [A('It may not have been. The job may have been claimed on an '), B('auto-added alternate start time'), A('.')],
    [B('Check the job history first.'), A(' The fix and explanation differ.')],
    [A('C doesn\'t want auto-added times? Add the '), B('do_not_auto_add_alternate_start_times'), A(' flag.')]], col='AMBER', bg='AMBER_PALE', line=None, after=5)
box(s, X, 306, W, 38, 'WHITE', 'LINE', margin=14, paras=[
    [B('CP moved it without C\'s consent?  ', 'AMBER', 9), A('That\'s not this flow. Use Unauthorized Reschedule (Service Issues), which has its own investigation and penalty process.', 'INK', 9)]])
s.notes_slide.notes_text_frame.text = (
    'Rescheduling covers a customer or CP moving an already-claimed appointment to a different date or time, the routine authorized case. If a CP rescheduled without the '
    'customer\'s consent, that is Unauthorized Reschedule under Service Issues, with its own investigation and penalty process.\n'
    'Customer reschedules: available from the dashboard any time 6 hours or more before start. If the new slot is less than 48 hours away, a Priority Fee may apply (AG – Refund – '
    'Priority Fee if they contest it). This is a different window from new bookings, which need 12 hours lead time, so do not apply the 12-hour rule to a reschedule. Under 6 hours, '
    'the dashboard option is gone, but they are not out of options: advise them to contact their CP, who can reschedule on their behalf.\n'
    'CP reschedules: CPs are encouraged to reschedule rather than cancel, so the job still gets done. There is no fixed timing restriction as long as the customer agrees, even after '
    'the original time has passed, and this also applies when the CP reschedules for the customer inside the 6-hour window. If they cannot agree, the CP cancels from their dashboard.\n'
    '"Why was my cleaning rescheduled?" Sometimes the job was actually claimed on one of its auto-added Alternate Start Times. Check the job history before responding. If the customer '
    'does not want alternate start times added automatically, add the do_not_auto_add_alternate_start_time(s) flag to the account.')

kc('RESCHEDULING', [
    ('C wants to move today\'s cleaning. It starts in 3 hours. Can they?', 'Not from the dashboard. Ask C to contact their CP, who can reschedule with C\'s agreement.'),
    ('C reschedules into a slot 30 hours away. Could a fee apply?', 'Yes. A Priority Fee may apply when the new slot is under 48 hours away.'),
    ('C asks why their cleaning was rescheduled. What do you check first?', 'The job history. It may have been claimed on an auto-added alternate start time.'),
    ('A CP moved the job without asking C. Is this a reschedule ticket?', 'No. Use Unauthorized Reschedule (Service Issues).')],
   'Knowledge check: Rescheduling. 1) Under 6 hours, the CP can still reschedule. 2) Under 48 hours, a Priority Fee may apply. 3) Check job history for alternate start times. '
   '4) No consent means Unauthorized Reschedule.')

# ================= 03 Cancellation (Job-Level) =================
EB = 'CANCELLATION (JOB-LEVEL)'
topic_divider(3, EB, 'Cancellation (Job-Level)', 'Cancelling a single job or appointment, not the customer\'s ForeverClean membership.',
              ('Job vs membership', 'Who cancels, and the job status', 'The $40 LMC fee', 'Reason codes and follow-up'),
              'The right action depends on who started the cancellation, what status the job is in, and whether a fee applies. The reason code also matters later: it decides '
              'whether a CP is penalized, whether an LMC fee applies, and how the incident shows up in account-history reviews such as the two-consecutive-serious-issue ETF waiver check.')

s = new_slide(EB, 'Who is cancelling?', 'Cancelling a job doesn\'t cancel the membership, and cancelling the membership doesn\'t cancel the job.')
table(s, X, 100, [130, W - 130], [20, 30, 30, 30, 26], [
    ('Who', 'Quick reference'),
    ('Customer', 'From the dashboard, any time before the job is invoiced. May trigger an LMC fee.'),
    ('Cleaner (CP)', 'From the dashboard or app, any time before invoiced. Affects Keep Rate, except when reported as an emergency, a safety concern, or "client wants to cancel".'),
    ('Admin (you, in CRM)', 'C CRM › Do › Job › Cancel. Lets you waive LMC fees, and never penalizes the CP.'),
    ('System', 'Auto-cancels under the conditions in Job Statuses › Cancelled.')])
cw = (W - 10) / 2
panel(s, X, 254, cw, 90, 'Check the job status first', [
    [B('Pending Invoice'), A(' doesn\'t prove the job happened. Follow AG: Pending Invoice\'s completion checks before cancelling.')],
    [B('Invoiced'), A(' jobs can only become Cancelled through a full refund.')]], bg='TEAL_PALE', line=None)
panel(s, X + cw + 10, 254, cw, 90, 'Job vs membership', [
    [A('C can cancel one job and '), B('keep their FC'), A('. C can cancel FC while a job still exists, so cancel upcoming jobs and RC as a separate step.')],
    [A('Ending the membership? See Retention and Cancellation.', 'MUTED', 8)]], col='AMBER', bg='AMBER_PALE', line=None)
s.notes_slide.notes_text_frame.text = (
    'Cancellation here means a single job or appointment, not the ForeverClean membership. The word does double duty: a customer can cancel one job and keep their membership fully '
    'active, and a customer can cancel their membership while a scheduled job still exists, which is why the FC cancellation playbooks remind you to cancel upcoming appointments and RC '
    'as a separate step. Do not assume one implies the other.\n'
    'Step 1, who is cancelling. Customer: from the dashboard any time before invoiced; may trigger an LMC fee. Cleaner: from the dashboard or app any time before invoiced; affects '
    'Keep Rate except when reported as an emergency, a safety concern or "client wants to cancel". Admin, you in the CRM: C CRM > Do > Job > Cancel; lets you waive LMC fees and never '
    'penalizes the CP. System: auto-cancels per Job Statuses > Cancelled.\n'
    'Step 2, check the status. Pending Invoice does not confirm the job happened, so follow AG: Pending Invoice\'s completion-verification steps before cancelling one. Invoiced jobs '
    'can only become Cancelled via a full refund, not a direct cancellation.')

s = new_slide(EB, 'The LMC fee, reason codes and follow-up', 'C cancels under 6 hours before start: a $40 Last-Minute Cancellation fee applies by default.')
LW = 352
panel(s, X, 100, LW, 244, 'The LMC is waived automatically when', [
    ('•', [A('The job was still '), B('Submitted'), A(' (no CP matched yet)')]),
    ('•', [A('C\'s area had '), B('severe weather')]),
    ('•', [A('The CP rescheduled under 24 hours before the cancellation, to a time C didn\'t ask for')]),
    ('•', [A('The CP claimed an alternate start time added under 48 hours before the cancellation')]),
    ('•', [A('C asked for the CP\'s ETA and the CP didn\'t check in within 30 minutes of start')]),
    ('•', [A('All of: no prior LMC waiver for C, the CP didn\'t check in over 30 minutes before start, and the job was claimed or rescheduled in the last 24 hours with an active FC or a start time that isn\'t C\'s preferred one')]),
    [B('Charge not valid? ', 'AMBER'), A('(e.g. the CP wanted to cancel, or no-showed) Refund C '), B('and'), A(' create a $20 Cleaner Payment Holdback for the CP\'s LMC share.')]],
      bg='TEAL_PALE', line=None)
RX, RW = X + LW + 10, W - LW - 10
panel(s, RX, 100, RW, 132, 'Pick the right reason code', [
    [A('It decides what shows in the job history and whether the CP is penalized. The full table is in '), B('AG: Pending Invoice › Cancellation Reason Codes'), A('.')],
    [A('Most used here: Customer wants one time, Customer doesn\'t want service, CP Quality, CP no-show/reschedule, CP cancellation, No CP Claimed, Customer schedule, '
       'Customer moved, Customer double-booked.', 'MUTED', 8)]])
panel(s, RX, 240, RW, 104, 'Follow up', [
    ('•', [A('Cancelled because of the CP? Consider a '), B('Pairing'), A(' block or '), B('CP Profile Status'), A(' coaching or penalty.')]),
    ('•', [A('FC member who really wants out? Check '), B('Retention and Cancellation'), A(' instead of handling it only at job level.')])],
      col='AMBER', bg='AMBER_PALE', line=None)
s.notes_slide.notes_text_frame.text = (
    'Step 3, the LMC fee. If the customer cancelled less than 6 hours before start, a $40 Last-Minute Cancellation fee applies by default. Before charging or refunding it, confirm '
    'whether it is valid. It is waived automatically when: the job was still Submitted; the customer\'s area had severe weather; the CP rescheduled less than 24 hours before the '
    'cancellation and the new time was not requested by the customer; the CP claimed an Alternate Start Time added less than 48 hours before the cancellation; the customer asked for the '
    'CP\'s ETA and the CP did not check in within 30 minutes of start; or all of: no prior LMC waiver for this customer, the CP did not check in more than 30 minutes before start, and the '
    'job was claimed or rescheduled in the last 24 hours with either an active FC or a start time that is not the customer\'s preferred one.\n'
    'If the LMC charge is invalid, for example the customer only cancelled because the CP wanted to cancel or reschedule, or the CP was a no-show: refund the customer and create a '
    'Cleaner Payment Holdback for the $20 the CP would have received. Do not just refund one side.\n'
    'Step 4, the reason code decides what shows in the job history and whether the CP is penalized. AG: Pending Invoice > Cancellation Reason Codes is the full table; check it rather '
    'than guessing from the shortlist.\n'
    'Step 5: if cancelling because of a CP issue, consider whether a Pairing block or CP Profile Status coaching or penalty is also warranted. If an FC member\'s cancellation is the '
    'first sign of a bigger "I want out" conversation, check Retention and Cancellation.')

kc('CANCELLATION (JOB-LEVEL)', [
    ('C cancels 4 hours before start. No CP had claimed the job yet. LMC fee?', 'No. It\'s waived automatically while the job is still Submitted.'),
    ('The LMC wasn\'t valid because the CP no-showed. What do you do?', 'Refund C and create a $20 Cleaner Payment Holdback for the CP. Don\'t fix just one side.'),
    ('You cancel a job from C CRM. Does it penalize the CP?', 'No. Admin cancellations never penalize the CP, and you can waive the LMC.'),
    ('C wants an invoiced job cancelled. Can you cancel it directly?', 'No. An invoiced job only becomes Cancelled through a full refund.')],
   'Knowledge check: Cancellation (Job-Level). 1) Submitted jobs have no LMC. 2) An invalid LMC needs both the refund and the $20 holdback. 3) Admin cancels never penalize the CP. '
   '4) Invoiced jobs need a full refund.')

# ================= 04 Pairing =================
EB = 'PAIRING'
topic_divider(4, EB, 'Pairing', 'Whether a specific customer and cleaner can be matched with each other again.',
              ('What a block does', 'When to block', 'How to block or unblock', 'Common mistakes'),
              'Blocking a pairing is a standard step in almost every Service Issue and Retention ticket: PCQ, Overcharged Hours, False Invoice, CP no-shows, unauthorized reschedules. '
              'Get it wrong and the same problem can come back on the customer\'s very next booking.')

s = new_slide(EB, 'Blocking a C/CP pairing', 'A block stops one customer and one cleaner being matched again. Both can still work with everyone else.')
cw = (W - 10) / 2
panel(s, X, 100, cw, 132, 'When to block', [
    [B('Default: block'), A(', unless C explicitly wants the same cleaner again. You need a reason not to block, not a reason to.')],
    [B('Don\'t block when', 'AMBER')],
    ('•', [A('C asks to keep or retry that cleaner')]),
    ('•', [A('The issue was fully C\'s fault, or a misunderstanding')]),
    [A('Telling C is optional: do it silently, or mention it as reassurance.', 'MUTED', 8)]], bg='TEAL_PALE', line=None)
panel(s, X + cw + 10, 100, cw, 132, 'How to block', [
    ('1', [A('Open the '), B('C/CP comms'), A(' of the CP you want to ban')]),
    ('2', [A('Use the dropdown at the lower right of a C/CP thread and select '), B('CP Ban List')]),
    ('3', [A('Enter the '), B('Customer ID'), A(' and click '), B('Submit')]),
    [A('Always start from that C/CP thread: it pre-fills the right CP. To check or undo a block, use the same CP Ban List.', 'MUTED', 8)]], after=4)
mw = (W - 20) / 3
for i, (h, body) in enumerate([
        ('Wrong direction', 'It blocks one customer from one CP. It doesn\'t stop the CP working with others, or C booking other CPs.'),
        ('One block per CP', 'Issues with two different CPs across two tickets need two separate blocks.'),
        ('Telling C isn\'t required', 'It\'s optional messaging. Don\'t hold up the resolution over the wording.')]):
    panel(s, X + i * (mw + 10), 242, mw, 102, h, [[A(body)]], col='AMBER', bg='AMBER_PALE', line=None)
s.notes_slide.notes_text_frame.text = (
    'Pairing is whether a specific customer and cleaner can be matched with each other going forward. Blocking a pairing tells the system never to match that customer with that '
    'cleaner again; it does not affect either party working with anyone else.\n'
    'When to block: block by default unless the customer explicitly wants the same cleaner again. Across PCQ, reliability issues, overcharged hours and most other service failures, '
    'blocking is the standard action: you need a reason not to. Do not block when the customer asks to keep or retry the same cleaner, or when the issue was fully the customer\'s fault '
    'or a misunderstanding. Telling the customer is optional: do it silently, or mention it as reassurance ("I\'ve made sure you won\'t be matched with them again").\n'
    'How: go to the C/CP comms of the CP you want to ban, click the dropdown at the lower right of a comms thread between the customer and that CP, select CP Ban List, enter the '
    'Customer ID and click Submit. Always start from the specific thread so the CP context is pre-filled. Use the same CP Ban List to check or reverse a block.\n'
    'Mistakes: blocking in the wrong direction (it blocks one customer from one CP only); forgetting it is per CP, not per ticket (two CPs need two blocks); and treating telling '
    'the customer as required. Disclosing internal pairing or ban mechanics to a customer or CP falls under the Zero Tolerance Policy (Disclosing Internal Information).\n'
    'Note: this article has moved to Notion (Pairing (C/CP Blocking)); this slide follows the last Knowledge Library version.')

kc('PAIRING', [
    ('PCQ ticket. C doesn\'t mention the cleaner at all. Block the pairing?', 'Yes. Blocking is the default unless C wants the same cleaner again.'),
    ('C says they\'d like to give the same cleaner another try. Block?', 'No. Don\'t block when C asks to keep or retry the cleaner.'),
    ('C had issues with two different CPs on two tickets. How many blocks?', 'Two. Blocks are per CP, not per ticket.'),
    ('Do you have to tell C you blocked the pairing?', 'No. It\'s optional. Use it as reassurance if it helps.')],
   'Knowledge check: Pairing. 1) Block by default. 2) Respect a request to retry. 3) One block per CP. 4) Telling C is optional.')

# ================= 05 Priority Booking =================
EB = 'PRIORITY BOOKING'
topic_divider(5, EB, 'Priority Booking', 'A job you create and book for a customer as a goodwill or retention gesture.',
              ('Three kinds of "priority"', 'Priority Booking vs Priority Fee', 'When to offer it', 'How to create one'),
              'PCQ, LMC, Lockout, Unauthorized Reschedule and Retention pages all say "offer a priority booking". There are three different "priority" ideas in the system, and mixing '
              'them up leads to real mistakes.')

s = new_slide(EB, 'Three kinds of "priority"', 'Priority Booking is something you give. A Priority Fee is something C paid.')
table(s, X, 100, [140, 330, W - 470], [20, 30, 38, 30], [
    ('Term', 'What it is', 'Who triggers it'),
    ('Priority Booking', 'You manually create and book a job for C as a goodwill gesture. No system flag: just a manually created job.', 'You, the agent'),
    ('Priority Request', 'An automatic upgrade for a submitted job still unclaimed 24–48 hours before start: more frequent CP notifications, and CPs up to 14% above C\'s max rate can claim it.', 'The system, automatically'),
    ('Priority Fee / Markup', 'A charge to C for booking within 48 hours of the start time.', 'The system, based on booking timing')],
      zebra={1: 'TEAL_PALE'})
cw = (W - 20) / 3
panel(s, X, 232, cw, 112, 'Booking or fee?', [
    [A('"You charged me for priority and now it\'s free?" '), B('Check which one is on the account'), A(' before answering.')],
    [A('Disputing a fee? Use AG – Refund – Priority Fee.', 'MUTED', 8)]], col='AMBER', bg='AMBER_PALE', line=None)
panel(s, X + cw + 10, 232, cw, 112, 'When to offer it', [
    ('•', [A('After a PCQ complaint, with a different cleaner')]),
    ('•', [A('As part of an LMC or Lockout resolution')]),
    ('•', [A('After a CP no-show or unauthorized reschedule')]),
    ('•', [A('Instead of cancelling an unclaimed job')])], bg='TEAL_PALE', line=None)
panel(s, X + 2 * (cw + 10), 232, cw, 112, 'How to create it', [
    [B('C CRM › Do › Job › Create'), A(', then fill in the appointment details.')],
    [A('No "priority" flag. It follows normal claiming, like any other job.')]], bg='TEAL_PALE', line=None)
s.notes_slide.notes_text_frame.text = (
    'Priority Booking is when you manually create and book a job for a customer as a goodwill or retention gesture, usually while resolving a service issue.\n'
    'Three "priority" concepts get mixed up. Priority Booking (this topic): you manually create a job, with no distinct system flag. Priority Request: an automatic system upgrade a '
    'submitted job gets 24 to 48 hours before start if still unclaimed, with more frequent CP notifications and CPs up to 14% above the customer\'s max requested rate allowed to claim. '
    'Priority Fee or Priority Markup: a charge to the customer for booking within 48 hours of start, applied automatically.\n'
    'Priority Booking and Priority Fee are opposites in practice: one is something you give for free, the other is something the customer paid for short notice. If a customer says '
    '"you charged me for priority and now you\'re offering it for free?", check which is on their account first. A disputed Priority Fee is a different workflow: AG – Refund – Priority Fee.\n'
    'Offer it after a PCQ complaint (priority cleaning with a different cleaner), as part of an LMC or Lockout resolution, after a CP no-show or unauthorized reschedule, or as an '
    'alternative to cancelling an unclaimed job. Create it in C CRM > Do > Job > Create; it then follows normal claiming like any job (see CP Tiering).\n'
    'Note: this article has moved to Notion (Priority Booking); this slide follows the last Knowledge Library version.')

kc('PRIORITY BOOKING', [
    ('"You charged me for priority, now you\'re offering it free?" What do you do first?', 'Check which is on the account: a Priority Fee C paid, or a Priority Booking you\'re giving.'),
    ('How do you create a Priority Booking?', 'C CRM › Do › Job › Create. There\'s no special flag.'),
    ('A submitted job is unclaimed 30 hours before start. Who makes it a Priority Request?', 'The system, automatically, 24–48 hours before start.')],
   'Knowledge check: Priority Booking. 1) Check which priority it is. 2) It is just a manually created job. 3) Priority Requests are automatic.')

# ================= 06 No Claim =================
EB = 'NO CLAIM'
topic_divider(6, EB, 'No Claim', 'A submitted job no CP has claimed yet. Nothing has gone wrong with a cleaning, because none has happened.',
              ('What the system is already doing', 'Answering "will it be claimed?"', 'Cancel, or priority booking?', 'The No CP Claimed code'),
              'The instinct is to reassure the customer and wait. But the system is already working on a timeline, and knowing it lets you give an accurate answer instead of a vague one.')

s = new_slide(EB, 'When nobody has claimed the job', 'Know where the job sits on the timeline before you answer.')
steps = [('24–48 hours before start', 'Upgrades to a ', 'Priority Request', ': more frequent CP notifications, and CPs up to 14% above C\'s max rate can claim.'),
         ('24 hours before start', 'Adds ', 'alternate start times', ' (10 AM–2 PM the same day, plus the next 2 days) to widen the pool.'),
         ('0, 12 or 24 hours before start', 'With no time left to match, the system ', 'auto-cancels', ' the job.')]
sw = (W - 2 * 22) / 3
for i, (when, a1, b1, a2) in enumerate(steps):
    x = X + i * (sw + 22)
    box(s, x, 100, sw, 92, 'TEAL_PALE' if i < 2 else 'AMBER_PALE')
    dot(s, x + 10, 109, str(i + 1), 'TEAL' if i < 2 else 'AMBER')
    text(s, x + 38, 111, sw - 46, 16, [[B(when, 'TEAL' if i < 2 else 'AMBER', 9.5)]])
    text(s, x + 12, 134, sw - 22, 56, [[A(a1), B(b1), A(a2)]])
    if i < 2:
        arrow(s, x + sw + 3, 146, x + sw + 19, 146)
text(s, X, 196, W, 12, [[A('You don\'t trigger any of this. Full mechanics: Job Statuses › Submitted.', 'MUTED', 8)]])
cw = (W - 10) / 2
panel(s, X, 214, cw, 130, 'C asks if or when it will be claimed', [
    [A('The most common version. Check the timeline: "it will upgrade within 24–48 hours" is a different answer from "it already has, and CPs are being notified more often now".')],
    [B('Don\'t promise a claim time.'), A(' Explain what the system is doing and when the next step kicks in.')]], bg='WHITE')
panel(s, X + cw + 10, 214, cw, 130, 'C wants to cancel because nobody claimed it', [
    [A('Would C still want the cleaning if it matched faster? Offer a '), B('Priority Booking'), A(' first.')],
    [A('Cancelling? Use the '), B('No CP Claimed'), A(' reason code (AG: Pending Invoice › Cancellation Reason Codes).')]], col='AMBER', bg='AMBER_PALE', line=None)
s.notes_slide.notes_text_frame.text = (
    'No Claim is a Submitted job no CP has picked up yet: the customer is asking whether or when it will be claimed, or wants to cancel because it has not been. It is not a service '
    'failure.\n'
    'The system is already working through a timeline for any unclaimed submitted job, and you do not trigger any of it. 24 to 48 hours before start, still unclaimed: the job '
    'auto-upgrades to a Priority Request, with more frequent CP notifications, and CPs at rates up to 14% above the customer\'s previous max requested rate can claim it. 24 hours before '
    'start, still unclaimed: the system auto-adds Alternate Start Times (10 AM to 2 PM the same day, plus the next 2 days) to widen the pool. 0, 12 or 24 hours before start, still '
    'unclaimed with no time left to match: the system auto-cancels the job. Full mechanics are in Job Statuses > Submitted.\n'
    'Match anxiety is the most common version. Check where the job sits before responding. Do not promise a specific claim time: explain what the system is doing and when the next '
    'automatic step kicks in.\n'
    'If the customer wants to cancel because nobody claimed it, use the No CP Claimed reason code. Before cancelling, consider whether a Priority Booking is a better offer, especially '
    'if they would still want the cleaning if it matched faster.')

kc('NO CLAIM', [
    ('C asks, "Will someone claim my cleaning by tonight?" What do you say?', 'Explain where the job is on the timeline and what happens next. Don\'t promise a claim time.'),
    ('The job is unclaimed and starts in 20 hours. What has the system already done?', 'Upgraded it to a Priority Request and added alternate start times.'),
    ('C wants to cancel because nobody claimed it, but still wants a cleaning. What do you offer?', 'A Priority Booking. If C still cancels, use the No CP Claimed reason code.')],
   'Knowledge check: No Claim. 1) Explain the timeline, never promise a time. 2) Under 24 hours, both automatic steps have run. 3) Offer a Priority Booking before cancelling.')

# ================= 07 Cleaner Swap =================
EB = 'CLEANER SWAP'
topic_divider(7, EB, 'Cleaner Swap', 'A higher-tier CP claims a job a lower-tier CP had already claimed, and takes it over.',
              ('What a swap is', 'What to tell the customer', 'What to tell the CP', 'Swap protection'),
              'Swaps are invisible to both sides by design: the customer sees a different cleaner, and the swapped-out CP is told the customer cancelled. That only works if you '
              'don\'t break it. Telling a cleaner a swap happened, or how tiering and swapping work, is a Zero Tolerance Policy violation.')

s = new_slide(EB, 'Cleaner swaps: what you can say', 'Talk about the outcome. Never the mechanism.')
table(s, X, 100, [190, W - 190], [20, 30, 30, 36], [
    ('Situation', 'What to say'),
    ('C asks why their CP changed', 'You\'ve been matched with a different, well-suited cleaner for this job. Don\'t mention tiers, swapping or "someone else took it".'),
    ('A CP asks why they were removed', 'The customer cancelled that booking. Don\'t mention tiers, swapping or being outranked by another CP.'),
    ('Either asks "Was I swapped?" or "Does tiering exist?"', 'Don\'t confirm or explain it. Point back to the outcome: a new CP was matched, or the booking is no longer available.')],
      zebra={})
cw = (W - 10) / 2
panel(s, X, 230, cw, 114, 'Zero Tolerance Policy', [
    [A('Disclosing that a swap happened, or how tiering and swapping work, '), B('to cleaners is a ZTP violation'), A('. It isn\'t a judgment call.')],
    [A('This is the CP Swap / Tiering example named under Disclosing Internal Information.', 'MUTED', 8)]], col='AMBER', bg='AMBER_PALE', line='AMBER')
panel(s, X + cw + 10, 230, cw, 114, 'Swap protection', [
    [A('A CP can\'t be swapped out if they\'ve '), B('sent and received at least one non-system message'), A(' with C in the past month.')],
    [A('Swap disputed? Check this first. CP Tiering › Tiering Model has the 24–48 hour windows by tier.', 'MUTED', 8)]], bg='TEAL_PALE', line=None)
s.notes_slide.notes_text_frame.text = (
    'A Cleaner Swap happens when a higher-tier CP claims a job a lower-tier CP had already claimed: the original CP is swapped out and a different, usually better-suited CP takes '
    'the job. The tier mechanics live in CP Tiering; this topic is what to do when someone asks.\n'
    'Swapping is invisible by default: the customer just sees a different cleaner, and the swapped-out CP is told the customer cancelled. That framing is deliberate. Disclosing that '
    'a swap happened, or how tiering or swapping works, to cleaners is a Zero Tolerance Policy violation; it is the CP Swap/Tiering example named under Disclosing Internal Information.\n'
    'Customer asks why their CP changed: they have been matched with a different, well-suited cleaner for this job. CP asks why they were removed: the customer cancelled that booking. '
    'If either asks directly "was I swapped?" or "does tiering exist?", do not confirm or explain the mechanic; redirect to the outcome.\n'
    'Swap protection: a CP is protected if they have sent and received at least one non-system-generated message with the customer in the past month. If a swap is disputed, check that '
    'first; CP Tiering > Tiering Model has the 24 to 48 hour windows by tier.')

kc('CLEANER SWAP', [
    ('A CP asks why they were taken off a job. What do you say?', '"The customer cancelled that booking." No mention of tiers or another CP.'),
    ('C asks why they have a different cleaner now. What do you say?', 'They\'ve been matched with a different, well-suited cleaner for this job.'),
    ('A CP asks, "Does tiering exist?" Do you explain it?', 'No. Don\'t confirm or explain it. Disclosing it to cleaners is a ZTP violation.'),
    ('A CP disputes a swap. What do you check first?', 'Swap protection: at least one non-system message sent and received with C in the past month.')],
   'Knowledge check: Cleaner Swap. 1) and 2) Talk about the outcome. 3) Never confirm the mechanism. 4) Check swap protection first.')

# ================= 08 Common Questions =================
EB = 'COMMON QUESTIONS'
topic_divider(8, EB, 'Common Questions', 'Pre-booking and service FAQs: quick answers, not decision trees.',
              ('Coverage and booking basics', 'Reviews and account', 'Payment method issues', 'Scheduling conflicts'),
              'These questions come up constantly but do not fit the decision-heavy pages. This topic gives agents the answer and the macro without digging through several mechanics pages.')


def faq_grid(s, y0, h, items, cols=4):
    fw = (W - 8 * (cols - 1)) / cols
    rows = (len(items) + cols - 1) // cols
    fh = (h - 8 * (rows - 1)) / rows
    for i, (q, body, macro) in enumerate(items):
        r_, c_ = divmod(i, cols)
        x, y = X + c_ * (fw + 8), y0 + r_ * (fh + 8)
        paras = [[B(q, 'TEAL', 8.5)]] + [[A(b, 'INK', 8)] if isinstance(b, str) else b for b in body]
        if macro:
            paras.append([A('Macro: ' + macro, 'MUTED', 7.5)])
        box(s, x, y, fw, fh, 'WHITE', 'LINE')
        text(s, x + 9, y + 8, fw - 18, fh - 12, paras, after=2)


s = new_slide(EB, 'Coverage and booking basics', 'Straight answers, plus the macro to use.')
faq_grid(s, 100, 244, [
    ('"Do you cover my area?"', ['Yes: confirm, point C to cleaners\' profiles, reviews and availability, and offer to help book.',
                                 [A('No: apologize, say the network grows weekly. ', 'INK', 8), B('Unused voucher? Refund it.', 'INK', 8)],
                                 'Check: homeaglow.com/coverage'], 'Coverage (Yes) / Coverage (No)'),
    ('"What\'s the minimum booking?"', ['2 hours, so the cleaner isn\'t rushed. If it takes less, C pays only for time spent.'], 'Minimum hours'),
    ('"How many hours should I book?"', ['2 hrs: light kitchen/bathroom. 3: most 1BR. 4: standard 2BR/2BA. 6: smaller homes, larger spaces or deep first cleans.',
                                         'The booking flow also recommends a duration.'], 'Recommended Hours'),
    ('"Same-day bookings?"', ['We accept requests up to 12 hours before the appointment, so requested cleaners have time to review and accept.'], None),
    ('"Insured or bonded?"', ['Cleaners pass a criminal background check, and ratings and reviews show on their profile. Homeaglow doesn\'t provide insurance; some cleaners carry their own.'], 'Bonding & Insurance'),
    ('"What\'s included?"', ['May include bathrooms, dusting, vacuuming, mopping, dishes, trash, light organizing, laundry and wash-and-fold, depending on C\'s needs. Full list: Help Center.'], 'What\'s Included'),
    ('"Is tip included?"', ['No. Tipping is never required. C can tip and review from the dashboard after the appointment.'], 'Tip not included'),
    ('"I moved. Do you cover my new area?"', ['Congratulate them, confirm nationwide coverage, ask for the new address. Need time to settle? Offer to pause the membership for a month.'], 'Moved')])
s.notes_slide.notes_text_frame.text = (
    'Common questions, part 1: coverage and booking basics. Coverage: if yes, confirm and point them to cleaners\' profiles, reviews and availability, and offer to help book if they '
    'share availability. If no, apologize, note the network grows with new cleaners weekly, and invite them to check back; if they have an unused purchased voucher, refund it rather '
    'than leave them holding a voucher they cannot use. Check coverage at homeaglow.com/coverage. Macros: Coverage (Yes) / Coverage (No).\n'
    'Minimum booking: 2 hours, so the cleaner can do a thorough job without rushing; if it takes less, the customer is only charged for actual time (Minimum hours). How many hours: '
    '2 for a light kitchen or bathroom, 3 for most 1BR apartments, 4 for a standard 2BR/2BA, 6 for smaller homes, larger spaces or deep first-time cleanings; the booking flow also '
    'recommends a duration (Recommended Hours). Same-day: we accept requests up to 12 hours before the appointment. Insurance: cleaners complete a criminal background check and their '
    'ratings and reviews are visible; Homeaglow does not provide insurance, though individual cleaners may carry their own (Bonding & Insurance). What\'s included: may include '
    'bathrooms, dusting, vacuuming, mopping, dishes, trash removal, light organizing, laundry, wash-and-fold and more depending on what they select; point them to the Help Center '
    '(What\'s Included). Tip: never required; they can tip and review from the dashboard after the appointment (Tip not included). Moved: congratulate them, confirm nationwide coverage, '
    'ask for the new address; if they need time to settle, offer to pause their membership for a month (Moved).')

s = new_slide(EB, 'Account, payment and scheduling questions', 'Quick answers. Route anything bigger to its own playbook.')
faq_grid(s, 100, 244, [
    ('"Poor reviews" / bad experience', ['Acknowledge it genuinely and offer to make it right. A specific incident goes to its Service Issues playbook.'], 'Poor Reviews'),
    ('Unsubscribe from marketing', ['Unsubscribe right away and confirm. Don\'t make C cancel their account or membership to stop emails.'], 'Unsubscribed'),
    ('Fully deactivate (non-FC C)', ['All in the first contact: unsubscribe, cancel upcoming jobs and RC, send the Deactivated macro, then deactivate (C CRM › Do › Deactivate).',
                                     'Active FC or unused voucher? Handle that first. Up to $20 courtesy credit is optional.'], 'Deactivated (send first)'),
    ('Two jobs too close together', ['Flag it to C proactively. Cancel one to avoid a mix-up, and offer to reopen it if C meant to keep both.'], 'Double book (cancelled)'),
    ('Charged on a new card', ['Banks can push replacement card details to merchants. Check whether Stripe auto-updated the card.'], 'C was charged on new card'),
    ('Charged on an old or deleted card', ['The current card failed, so the system fell back to a saved card. Offer to help update it.'], 'C was charged on old/deleted card'),
    ('Remove my card', ['Not possible with an active FC, RC or upcoming job. Ask what\'s behind the request: it may be an FC cancellation.'], 'Delete my credit card'),
    ('"Why do I keep being told to update my card?"', ['Every card is Capital One? MFs are skipped with no failed charge. Check Stripe first; then see Capital One Card on File.'],
     'Capital One – Add a Different Card (No Cleaning Yet)')])
s.notes_slide.notes_text_frame.text = (
    'Common questions, part 2. Bad experience or poor reviews concern: acknowledge genuinely; most customers have a great experience, but when something goes wrong it is taken '
    'seriously. Offer to make it right, and route a specific incident to the relevant Service Issues playbook (Poor Reviews). Unsubscribe from marketing: do it immediately and confirm; '
    'do not require them to cancel their account or membership (Unsubscribed).\n'
    'Fully deactivate, non-FC customer: do it all in the first contact, whether they say close or deactivate: unsubscribe (C CRM > Do > Deactivate > Unsubscribe), cancel upcoming '
    'appointments and RC, send the Deactivated macro, then deactivate (C CRM > Do > Deactivate). You may offer up to $20 courtesy credit if they create a new account later; it is '
    'optional. If they have an active FC or an unused DHJ or FCF voucher, handle that first. If they ask to delete personal data, follow the data deletion article.\n'
    'Payment: charged on a new card, banks sometimes push updated card info to merchants, so check whether Stripe auto-updated it. Charged on an old or deleted card, the current card '
    'failed and the system fell back to a saved one; offer to help update it. Remove my card: not possible with an active FC, RC or upcoming appointment; ask what is prompting it. '
    'Capital One: if every card is Capital One, MFs are skipped with no failed charge; check Stripe first because some MFs have gone through; if it went through, tell them the account '
    'is current; otherwise follow Capital One Card on File.\n'
    'Double-booked: proactively flag it, cancel one of the two, and offer to reopen it if they meant to keep both (Double book (cancelled)).')

kc('COMMON QUESTIONS', [
    ('C\'s area isn\'t covered, and they have an unused voucher. What do you do?', 'Refund the voucher. Don\'t leave them holding one they can\'t use.'),
    ('C wants a cleaning today, in 8 hours. Can they book it?', 'No. We accept requests up to 12 hours before the appointment.'),
    ('C with an active FC asks to remove their card. Can you?', 'No. A card must stay on file. Ask what\'s prompting the request.'),
    ('C has two appointments booked too close together. What do you do?', 'Flag it, cancel one, and offer to reopen it if C meant to keep both.')],
   'Knowledge check: Common Questions. 1) Refund an unused voucher when there is no coverage. 2) 12 hours minimum notice. 3) A card must stay on file with FC, RC or an upcoming job. '
   '4) Handle double bookings proactively.')

# overview tiles jump to each topic divider
for shapes, target in zip(over_tiles, dividers):
    for sh in shapes:
        sh.click_action.target_slide = target

order.append(questions)
set_runs(questions.shapes[1], 'Thank you!')
questions.notes_slide.notes_text_frame.text = 'Open the floor for questions. When in doubt, check the Knowledge Library (Bookings) or ask your Trainer or Team Lead.'

# keep only the slides we built, in order
want = [id(s) for s in order]
all_slides = list(prs.slides)
els = list(lst)
pairs = dict((id(sl), e) for sl, e in zip(all_slides, els))
for sl, e in zip(all_slides, els):
    if id(sl) not in want:
        rid = e.get(qn('r:id')); lst.remove(e); prs.part.drop_rel(rid)
for e in list(lst):
    lst.remove(e)
for sl in order:
    lst.append(pairs[id(sl)])
prs.save(sys.argv[2]); print('saved', sys.argv[2], len(prs.slides), 'slides')
