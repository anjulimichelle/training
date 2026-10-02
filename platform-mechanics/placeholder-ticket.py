"""Placeholder name on the triggered-ticket screenshot; copy the voucher panel as is."""
import sys; sys.path.insert(0, '.')
from PIL import Image, ImageDraw, ImageFont
from retext_cp import replace, fill_rows, fit, ink
B = '/tmp/claude-0/-workspace-training/3ce5cff6-cccf-5f33-b0ee-0b147a55120e/'
FIG = {w: f'/root/.fonts/Figtree-{w}.ttf' for w in (500, 600, 700)}
im = Image.open(B + 'images/27.png').convert('RGB')
im = im.resize((im.width * 2, im.height * 2), Image.LANCZOS)  # work at 2x so new text is crisp
(x0, y0, x1, y1), _ = ink(im, (68, 16, 144, 56))
f = fit(FIG[600], 'Brittany', x1 - x0 + 1)
dx = int(round(f.getlength('Anjuli Kintanar') - f.getlength('Brittany')))
strip = im.crop((146, 8, 1424 - dx, 60))
fill_rows(im, (146, 8, 1423, 60), 1423)
im.paste(strip, (146 + dx, 8))
replace(im, (68, 16, 144, 56), 'Brittany', 'Anjuli Kintanar', FIG[600])
d = ImageDraw.Draw(im)
d.ellipse((10, 12, 52, 54), fill=(97, 95, 255))
d.text((31, 33), 'AK', font=ImageFont.truetype(FIG[600], 17), fill='white', anchor='mm')
im.save('img3/ticket_nonlogged.png')
v = Image.open(B + 'images/28.png').convert('RGB')
v.resize((v.width * 2, v.height * 2), Image.LANCZOS).save('img3/vouchers_newcrm.png')
