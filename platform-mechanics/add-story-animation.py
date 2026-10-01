"""Give the shape named storyText an on-click, paragraph-by-paragraph fade-in, so the story builds line by line."""
import re, sys, zipfile, shutil
src, dst = sys.argv[1], sys.argv[2]
zin = zipfile.ZipFile(src)
zout = zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED)
for item in zin.infolist():
    data = zin.read(item.filename)
    if re.match(r'ppt/slides/slide\d+\.xml$', item.filename) and b'name="storyText"' in data:
        x = data.decode('utf8')
        spid = re.search(r'<p:cNvPr id="(\d+)" name="storyText"', x).group(1)
        start = x.index('name="storyText"')
        sp_end = x.index('</p:sp>', start)
        n = x[start:sp_end].count('<a:p>') + x[start:sp_end].count('<a:p ')
        pars, cid = [], 3
        for i in range(n):
            tgt = f'<p:tgtEl><p:spTgt spid="{spid}"><p:txEl><p:pRg st="{i}" end="{i}"/></p:txEl></p:spTgt></p:tgtEl>'
            pars.append(
                f'<p:par><p:cTn id="{cid}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>'
                f'<p:par><p:cTn id="{cid+1}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
                f'<p:par><p:cTn id="{cid+2}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect">'
                f'<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
                f'<p:set><p:cBhvr><p:cTn id="{cid+3}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>{tgt}'
                f'<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
                f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{cid+4}" dur="600"/>{tgt}</p:cBhvr></p:animEffect>'
                f'</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>')
            cid += 5
        timing = ('<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
                  '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
                  + ''.join(pars) +
                  '</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
                  '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>'
                  '</p:childTnLst></p:cTn></p:par></p:tnLst>'
                  f'<p:bldLst><p:bldP spid="{spid}" grpId="0" build="p"/></p:bldLst></p:timing>')
        x = x.replace('</p:sld>', timing + '</p:sld>') if '<p:timing>' not in x else x
        # p:timing must come after p:clrMapOvr and before p:extLst
        if '<p:extLst>' in x.split('<p:timing>')[0].split('</p:cSld>')[-1]:
            pass
        data = x.encode('utf8')
        print('animated', item.filename, 'spid', spid, 'paragraphs', n)
    zout.writestr(item, data)
zout.close()
