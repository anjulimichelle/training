"""Add on-click builds Google Slides keeps:
- a shape named reveal*: its paragraphs fade in one per click (line by line);
- shapes named stepN*: everything with the same N appears together, one step per click, in N order;
- shapes named popA_B*: appear on click A and disappear again on click B (a pop-up that closes);
- shapes named typeN*: one per click, typed in letter by letter (PowerPoint typewriter; Google may show it whole)."""
import re, sys, zipfile

src, dst = sys.argv[1], sys.argv[2]


def effect(cid, tgt, node):
    return (f'<p:par><p:cTn id="{cid}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="{node}">'
            f'<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
            f'<p:set><p:cBhvr><p:cTn id="{cid+1}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>{tgt}'
            f'<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
            f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{cid+2}" dur="600"/>{tgt}</p:cBhvr></p:animEffect>'
            f'</p:childTnLst></p:cTn></p:par>')


def leave(cid, tgt, node):
    """Disappear: the exit counterpart of an appear."""
    return (f'<p:par><p:cTn id="{cid}" presetID="1" presetClass="exit" presetSubtype="0" fill="hold" grpId="1" nodeType="{node}">'
            f'<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
            f'<p:set><p:cBhvr><p:cTn id="{cid+1}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>{tgt}'
            f'<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="hidden"/></p:to></p:set>'
            f'</p:childTnLst></p:cTn></p:par>')


def typed(cid, tgt, node):
    """Appear, iterated by letter (40 ms apart): the PowerPoint typewriter build."""
    return (f'<p:par><p:cTn id="{cid}" presetID="1" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="{node}">'
            f'<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:iterate type="lt"><p:tmAbs val="40"/></p:iterate><p:childTnLst>'
            f'<p:set><p:cBhvr><p:cTn id="{cid+1}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>{tgt}'
            f'<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
            f'</p:childTnLst></p:cTn></p:par>')


def click(cid, inner):
    return (f'<p:par><p:cTn id="{cid}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>'
            f'<p:par><p:cTn id="{cid+1}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
            f'{inner}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>')


def timing(clicks, bld):
    return ('<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
            '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
            + ''.join(clicks) +
            '</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
            '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>'
            '</p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst>' + ''.join(bld) + '</p:bldLst></p:timing>')


zin = zipfile.ZipFile(src)
zout = zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED)
for item in zin.infolist():
    data = zin.read(item.filename)
    if re.match(r'ppt/slides/slide\d+\.xml$', item.filename):
        x = data.decode('utf8')
        # Linked buttons: the link lives on the shape, so drop pptxgenjs's duplicate
        # run-level link and its forced underline.
        if '<a:hlinkClick' in x:
            def clean(m):
                sp = m.group(0)
                if not re.search(r'<p:cNvPr[^>]*>\s*<a:hlinkClick', sp):
                    # Run-level link only: keep the run's own colour (not the theme's blue link colour).
                    return re.sub(r'<a:hlinkClick ([^>]*?)/>', lambda h: '<a:hlinkClick ' + h.group(1) + '><a:extLst><a:ext uri="{A12FA001-AC4F-418D-AE19-62706E023703}"><ahyp:hlinkClr xmlns:ahyp="http://schemas.microsoft.com/office/drawing/2018/hyperlinkcolor" val="tx"/></a:ext></a:extLst></a:hlinkClick>', sp)
                sp = re.sub(r'(<a:rPr[^>]*>)(.*?)<a:hlinkClick[^>]*?(?:/>|>.*?</a:hlinkClick>)', r'\1\2', sp, flags=re.S)
                sp = re.sub(r'(<a:rPr[^>]*?) u="sng"', r'\1 u="none"', sp)
                return re.sub(r'\s*<a:extLst>\s*<a:ext uri="\{A12FA001-AC4F-418D-AE19-62706E023703\}">.*?</a:extLst>', '', sp, flags=re.S)
            x = re.sub(r'<p:sp>.*?</p:sp>', clean, x, flags=re.S)
            data = x.encode('utf8')
        clicks, bld, cid = [], [], 3
        m = re.search(r'<p:cNvPr id="(\d+)" name="reveal', x)
        if m:
            spid = m.group(1)
            start = x.index('name="reveal')
            body = x[start:x.index('</p:sp>', start)]
            n = body.count('<a:p>') + body.count('<a:p ')
            for i in range(n):
                tgt = f'<p:tgtEl><p:spTgt spid="{spid}"><p:txEl><p:pRg st="{i}" end="{i}"/></p:txEl></p:spTgt></p:tgtEl>'
                clicks.append(click(cid, effect(cid + 2, tgt, 'clickEffect')))
                cid += 5
            bld.append(f'<p:bldP spid="{spid}" grpId="0" build="p"/>')
        # Click number -> shapes that appear ('in') or disappear ('out') on that click.
        steps = {}
        for sid, n in re.findall(r'<p:cNvPr id="(\d+)" name="step(\d+)[^"]*"', x):
            steps.setdefault(int(n), []).append((sid, 'in'))
        for sid, a, b in re.findall(r'<p:cNvPr id="(\d+)" name="pop(\d+)_(\d+)[^"]*"', x):
            steps.setdefault(int(a), []).append((sid, 'in'))
            steps.setdefault(int(b), []).append((sid, 'out'))
        built = set()
        for k in sorted(steps):
            inner, first = '', True
            for sid, kind in steps[k]:
                tgt = f'<p:tgtEl><p:spTgt spid="{sid}"/></p:tgtEl>'
                node = 'clickEffect' if first else 'withEffect'
                inner += effect(cid + 2, tgt, node) if kind == 'in' else leave(cid + 2, tgt, node)
                first = False
                cid += 3
            clicks.append(click(cid, inner))
            cid += 2
            for sid, kind in steps[k]:
                grp = '0' if kind == 'in' else '1'
                # build entries only for sp shapes (pictures don't take one)
                if (sid, grp) not in built and re.search(rf'<p:sp>(?:(?!</p:sp>).)*?<p:cNvPr id="{sid}"', x, re.S):
                    bld.append(f'<p:bldP spid="{sid}" grpId="{grp}" animBg="1"/>')
                    built.add((sid, grp))
        for sid, n in sorted(re.findall(r'<p:cNvPr id="(\d+)" name="type(\d+)[^"]*"', x), key=lambda t: int(t[1])):
            clicks.append(click(cid, typed(cid + 2, f'<p:tgtEl><p:spTgt spid="{sid}"/></p:tgtEl>', 'clickEffect')))
            cid += 4
            bld.append(f'<p:bldP spid="{sid}" grpId="0" animBg="1"/>')
        if clicks:
            x = x.replace('</p:sld>', timing(clicks, bld) + '</p:sld>')
            data = x.encode('utf8')
            print('animated', item.filename, 'clicks', len(clicks))
    zout.writestr(item, data)
zout.close()
