"""Build compressed "download copies" of training decks that are too large for the Google Drive
connector (16 MB response limit), so the Training Agenda Library can still combine their slides.
Each copy is the deck's own .pptx export with large images losslessly-looking quantized (256-colour
PNG), stored unconverted in the copies folder and titled with the source id and source edit time,
which the page checks so it never uses a copy older than the deck. Usage: make-download-copies.py <deck id>..."""
import io, json, os, re, subprocess, sys, tempfile, zipfile
from PIL import Image

FOLDER = os.environ['COPIES_FOLDER']
API = 'https://www.googleapis.com/drive/v3'
PPTX = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'

def curl(*args, binary=False):
    out = subprocess.run(['curl', '-sS', '--fail-with-body', *args], capture_output=True, check=True).stdout
    return out if binary else json.loads(out)

def shrink(data):
    zin = zipfile.ZipFile(io.BytesIO(data)); buf = io.BytesIO()
    zout = zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED)
    for it in zin.infolist():
        b = zin.read(it.filename)
        if it.filename.startswith('ppt/media/') and it.filename.lower().endswith('.png') and len(b) > 80_000:
            im = Image.open(io.BytesIO(b)); o = io.BytesIO()
            q = im.convert('RGBA').quantize(256, method=Image.FASTOCTREE) if im.mode in ('RGBA', 'LA', 'P') else im.convert('RGB').quantize(256)
            q.save(o, 'PNG', optimize=True)
            if o.tell() < len(b): b = o.getvalue()
        zout.writestr(it, b)
    zout.close(); return buf.getvalue()

existing = {}
for f in curl(f"{API}/files?q='{FOLDER}'+in+parents+and+trashed=false&fields=files(id,name)&pageSize=100")['files']:
    m = re.search(r'\| id (\S+)\]', f['name'])
    if m: existing[m.group(1)] = f['id']

for deck in sys.argv[1:]:
    meta = curl(f'{API}/files/{deck}?fields=name,modifiedTime')
    raw = curl(f'{API}/files/{deck}/export?mimeType={PPTX}', binary=True)
    slim = shrink(raw)
    title = f"{meta['name'][:120]} [download copy | src {meta['modifiedTime']} | id {deck}]"
    with tempfile.NamedTemporaryFile(suffix='.pptx', delete=False) as t: t.write(slim)
    if deck in existing:
        fid = existing[deck]
        curl('-X', 'PATCH', '-H', f'Content-Type: {PPTX}', '--data-binary', f'@{t.name}', f'https://www.googleapis.com/upload/drive/v3/files/{fid}?uploadType=media&fields=id')
        curl('-X', 'PATCH', '-H', 'Content-Type: application/json', '-d', json.dumps({'name': title}), f'{API}/files/{fid}?fields=id')
    else:
        with tempfile.NamedTemporaryFile('w', suffix='.json', delete=False) as m: json.dump({'name': title, 'mimeType': PPTX, 'parents': [FOLDER]}, m)
        fid = curl('-X', 'POST', '-F', f'metadata=@{m.name};type=application/json;charset=UTF-8', '-F', f'file=@{t.name};type={PPTX}',
                   'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id')['id']
    print(f"{meta['name'][:40]}: {len(raw)/1e6:.1f} MB -> {len(slim)/1e6:.1f} MB  copy {fid}  src {meta['modifiedTime']}")
