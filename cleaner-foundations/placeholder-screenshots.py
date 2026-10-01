"""Swap the real cleaner's details in the CP CRM / CP Dashboard screenshots for training placeholders."""
from PIL import Image, ImageDraw, ImageFont
import numpy as np

B = '/tmp/claude-0/-workspace-training/3ce5cff6-cccf-5f33-b0ee-0b147a55120e/'
SRC, OUT = B + 'images/', B + 'scratchpad/img2/'
LIB = '/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf'

NAME, SHORT, FIRST = 'Truffle Wuffle', 'Truffle W', 'Truffle'
EMAIL, PHONE, HANDLE = 'truffle_wuffle@gmail.com', '555-012345678', 'trufflewuffle10'


def ink(im, box, thr=70):
    a = np.asarray(im.crop(box)).astype(int)
    bg = a[0, 0]
    d = np.abs(a - bg).sum(axis=2)
    ys, xs = np.where(d > thr)
    iy, ix = np.unravel_index(d.argmax(), d.shape)
    return (box[0] + xs.min(), box[1] + ys.min(), box[0] + xs.max(), box[1] + ys.max()), tuple(a[iy, ix])


def fit(fontfile, text, width):
    best = None
    for s10 in range(60, 400):
        f = ImageFont.truetype(fontfile, s10 / 10)
        b = f.getbbox(text)
        err = abs((b[2] - b[0]) - width)
        if best is None or err < best[0]:
            best = (err, f)
    return best[1]


def fill_rows(im, box, refx):
    d = ImageDraw.Draw(im)
    for y in range(box[1], box[3] + 1):
        d.line((box[0], y, box[2], y), fill=im.getpixel((refx, y)))


def replace(im, box, old, new, fontfile=LIB, color=None):
    (x0, y0, x1, y1), inkc = ink(im, box)
    f = fit(fontfile, old, x1 - x0 + 1)
    fill_rows(im, (x0 - 1, y0 - 1, x1 + 1, y1 + 1), box[0])
    ob = f.getbbox(old)
    ImageDraw.Draw(im).text((x0 - ob[0], y0 - ob[1]), new, font=f, fill=color or inkc)
    return x0 - ob[0] + f.getlength(new), x1


def move(im, box, dx):
    """Move a strip horizontally by dx (either direction), backfilling with the row background."""
    strip = im.crop(box)
    fill_rows(im, box, box[0] - 1)
    im.paste(strip, (box[0] + dx, box[1]))


def face(src, crop, size, shape='ellipse'):
    f = Image.open(src).convert('RGB').crop(crop).resize(size, Image.LANCZOS)
    m = Image.new('L', size, 0)
    getattr(ImageDraw.Draw(m), shape)((0, 0, size[0] - 1, size[1] - 1), fill=255)
    return f, m


def replace_centered(im, box, old, new, fontfile=LIB):
    (x0, y0, x1, y1), inkc = ink(im, box)
    f = fit(fontfile, old, x1 - x0 + 1)
    fill_rows(im, (x0 - 1, y0 - 1, x1 + 1, y1 + 1), box[0])
    ob, nb = f.getbbox(old), f.getbbox(new)
    cx = (x0 + x1) / 2
    ImageDraw.Draw(im).text((cx - (nb[2] - nb[0]) / 2 - nb[0], y0 - ob[1]), new, font=f, fill=inkc)


CARTOON = (SRC + '11.webp', (938, 300, 1010, 372))   # green-cap cleaner from the onboarding graphic

# ---------------- Legacy CP CRM ----------------
C = Image.open(SRC + '7.png').convert('RGB')
end, _ = replace(C, (440, 80, 562, 104), 'Jennifer Salazar', NAME)
move(C, (563, 82, 830, 104), int(end) + 8 - 563)  # pull the badges in
for y in (479, 568, 807):                                   # sender column
    replace(C, (376, y - 5, 422, y + 9), 'Jennifer S', SHORT)
for y in (534, 602, 636, 670, 704, 738, 773):               # recipient column
    replace(C, (867, y - 5, 911, y + 9), 'Jennifer S', SHORT)
# Notes link: name, then pull "| claimed" left to close the gap.
end, x1 = replace(C, (805, 836, 885, 853), 'Jennifer Salazar', NAME)
move(C, (x1 + 2, 836, x1 + 80, 853), int(end - x1) + 2)
replace(C, (1201, 166, 1300, 186), 'jennifersalazar10', HANDLE)
replace(C, (1203, 857, 1410, 875), 'jenbaby210@gmail.com    (512) 676-7276', EMAIL + '    ' + PHONE)
f, m = face(*CARTOON, (116, 91))
C.paste(f, (317, 82), m)
fill_rows(C, (88, 864, 205, 892), 85)                      # drop the screenshot's red label
C.save(OUT + 'cpcrm.png')

# ---------------- CP Dashboard ----------------
D = Image.open(SRC + '8.png').convert('RGB')
replace(D, (1773, 34, 1819, 52), 'Jennifer', FIRST)
f, m = face(*CARTOON, (16, 16))
D.paste(f, (1820, 35), m)
replace_centered(D, (736, 496, 828, 513), '4009 Sabio Dr', '123 Sample St')      # customer addresses
replace_centered(D, (1099, 560, 1208, 577), '1701 oak hill lane', '456 Training Ave')
fill_rows(D, (90, 818, 200, 846), 85)
D.save(OUT + 'cpdash.png')
print('done')
