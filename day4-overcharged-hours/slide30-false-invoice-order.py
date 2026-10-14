"""Day 4 slide 30 (Confirmed false invoice: the process).

Reorder the steps, make clear a customer-invoiced job gets no CP penalty and no
cp_false_invoice flag, point Apply penalty at the penalty ladder, and fold the
macro tile into the step descriptions (coaching link on Coach, macro on Confirm).
Click-to-reveal order follows the new order.

Usage: python3 slide30-false-invoice-order.py in.pptx out.pptx
"""
import copy, re, sys
from pptx import Presentation
from pptx.util import Pt
from pptx.oxml.ns import qn

prs = Presentation(sys.argv[1])
s = prs.slides[29]
shapes = list(s.shapes)
txt = lambda sh: sh.text_frame.text if sh.has_text_frame else ''
assert txt(shapes[1]).startswith('Confirmed false invoice')

LABELS = ['FALSE INVOICE (CP-fault)', 'Issue a full refund', 'Check for cp_false_invoice flag',
          'Apply penalty', 'Coach / penalize', 'Check for Premium fee', 'Ban the C/CP pairing',
          'Confirm refund with C']
rows = []
for lab in LABELS:
    i = next(k for k, sh in enumerate(shapes) if txt(sh).startswith(lab))
    box = shapes[i - 1]
    grp = [box, shapes[i]]
    j = i + 1
    while shapes[j].shape_type == 9:  # connector lines
        grp.append(shapes[j]); j += 1
    desc = shapes[j]
    grp.append(desc)
    rows.append(dict(top=box.top, grp=grp, desc=desc, has_down=len(grp) == 5))
tops = [r['top'] for r in rows]

NEW = [0, 1, 5, 6, 2, 3, 4, 7]  # old row index for each new position
for pos, old in enumerate(NEW):
    dy = tops[pos] - rows[old]['top']
    for sh in rows[old]['grp']:
        sh.top = sh.top + dy
assert rows[NEW[-1]]['has_down'] is False and all(rows[o]['has_down'] for o in NEW[:-1])

# click-to-reveal order
timing = s._element.find(qn('p:timing'))
main = next(c for c in timing.iter(qn('p:cTn')) if c.get('nodeType') == 'mainSeq')
lst = main.find(qn('p:childTnLst'))
pars = {par.find('.//' + qn('p:spTgt')).get('spid'): par for par in lst.findall(qn('p:par'))}
for p in pars.values():
    lst.remove(p)
for old in NEW:
    lst.append(pars[str(rows[old]['desc'].shape_id)])


def grey_rpr(desc):
    for r in desc._element.iter(qn('a:r')):
        c = r.find(qn('a:rPr')).find('.//' + qn('a:srgbClr'))
        if c is not None and c.get('val') == '6E6E73':
            return copy.deepcopy(r.find(qn('a:rPr')))
    raise ValueError


def set_runs(desc, runs):
    """runs: list of (text, bold, color or None, hlink rId or None)."""
    base = grey_rpr(desc)
    p = desc._element.find('.//' + qn('a:p'))
    for extra in desc._element.findall('.//' + qn('a:p'))[1:]:
        extra.getparent().remove(extra)
    for r in p.findall(qn('a:r')):
        p.remove(r)
    end = p.find(qn('a:endParaRPr'))
    for text, bold, color, rid in runs:
        r = p.makeelement(qn('a:r'), {})
        rpr = copy.deepcopy(base)
        rpr.set('b', '1' if bold else '0')
        if color:
            rpr.find('.//' + qn('a:srgbClr')).set('val', color)
        if rid:
            h = rpr.makeelement(qn('a:hlinkClick'), {qn('r:id'): rid})
            rpr.append(h)
        r.append(rpr)
        t = r.makeelement(qn('a:t'), {}); t.text = text; r.append(t)
        if end is not None:
            end.addprevious(r)
        else:
            p.append(r)


INK, TEAL, AMBER = '1D1D1F', '0B7A6F', '9A6410'
d = [r['desc'] for r in rows]
set_runs(d[0], [('Confirm the CP invoiced the job. ', False, None, None),
                ('Customer invoiced it? ', True, INK, None),
                ('Refund only: no CP penalty and no cp_false_invoice flag.', True, INK, None)])
set_runs(d[2], [('CP-invoiced jobs only. ', True, INK, None),
                ('Already flagged? Check the Job ID matches. No flag? CP CRM › Do › Flag › '
                 'cp_false_invoice. Include the ID triplet (C | CP | Job).', False, None, None)])
set_runs(d[3], [('CP-invoiced jobs only. ', True, INK, None),
                ('Check the penalty ladder (next slide). ', True, TEAL, None),
                ("CP-facing reason: cpq_false_invoice Customer <C'S NAME> reported incorrect hours charged.",
                 False, None, None)])
coach_rid = next(k for k, r in s.part.rels.items() if r.is_external and 'spreadsheets' in r._target)
set_runs(d[4], [('Send coaching comms to CP.  ', False, None, None),
                ('CP Coaching: ', True, AMBER, None),
                ('False Invoice 🔗', True, TEAL, coach_rid)])
set_runs(d[7], [('Refund arrives in 5–10 business days. Address any other pain points.  ', False, None, None),
                ('Recommended macro: ', True, AMBER, None),
                ('False Invoice', True, INK, None)])

# macro tile is now redundant
tile_txt = next(sh for sh in shapes if txt(sh).startswith('Recommended Macro'))
k = shapes.index(tile_txt)
for sh in shapes[k - 2:k + 1]:
    sh._element.getparent().remove(sh._element)

# speaker notes follow the new order
nt = s.notes_slide.notes_text_frame
old = nt.text
parts = re.split(r'\s(\d)\) ', old)
head, steps = parts[0], {int(parts[i]): parts[i + 1].strip() for i in range(1, len(parts), 2)}
if len(steps) == 8 and 'Customer invoiced' not in old:
    tail_split = steps[8].split(' Macros:', 1)
    steps[8] = tail_split[0]
    tail = ' Macros:' + tail_split[1] if len(tail_split) > 1 else ''
    steps[4] = steps[4].replace('penalize accordingly', 'penalize accordingly using the penalty ladder (next slide)')
    order = [1, 6, 7, 2, 3, 4, 5, 8]
    new = (head + ' First confirm who invoiced: if the customer invoiced the job, refund only. '
           'No CP penalty and no cp_false_invoice flag.')
    for n, o in enumerate(order, 1):
        new += f' {n}) {steps[o]}'
    nt.text = new + tail

prs.save(sys.argv[2])
print('saved', sys.argv[2])
