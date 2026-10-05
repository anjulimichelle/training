"""Placeholder names on the pending-invoice C/CP comms screenshot (image 37)."""
import sys; sys.path.insert(0, '.')
from PIL import Image, ImageDraw
from retext_cp import ink, fit, replace, move, fill_rows, LIB
B = '/tmp/claude-0/-workspace-training/3ce5cff6-cccf-5f33-b0ee-0b147a55120e/'
LIBB = LIB.replace('Regular', 'Bold')
im = Image.open(B + 'images/37.png').convert('RGB')


def swap(box, old, new, font=LIB, tail_end=None):
    """Replace old->new in box, shifting the rest of the line (up to tail_end) to keep spacing."""
    (x0, y0, x1, y1), _ = ink(im, box)
    if tail_end:
        f = fit(font, old, x1 - x0 + 1)
        dx = int(round(f.getlength(new) - f.getlength(old)))
        if dx > 0:
            move(im, (x1 + 2, box[1], tail_end - dx, box[3]), dx)
        elif dx < 0:
            move(im, (x1 + 2, box[1], tail_end, box[3]), dx)
    replace(im, box, old, new, font)


swap((90, 27, 140, 44), 'Chase L', 'Anjuli K', LIBB)
swap((461, 27, 503, 44), 'Scott L', 'Truffle W')
# Notes line: redraw everything from "C 4390823" to the end of the line.
OLD = 'C 4390823 Chase Lilly | CP 1570164 Scott Lopez Flores | claimed'
GREY, BLUE = (133, 132, 133), (56, 127, 187)
f = fit(LIB, OLD, 789 - 369 + 1)
fill_rows(im, (366, 77, 800, 96), 365)
ob = f.getbbox(OLD)
x, y = 369 - ob[0], 80 - ob[1]
d = ImageDraw.Draw(im)
for text, col in [('C ', GREY), ('4390823 Anjuli Kintanar', BLUE), (' | CP ', GREY), ('1570164 Truffle Wuffle', BLUE), (' | claimed', GREY)]:
    d.text((x, y), text, font=f, fill=col)
    x += f.getlength(text)
swap((97, 160, 142, 178), 'Scott L', 'Truffle W', LIBB)
swap((453, 160, 503, 178), 'Chase L', 'Anjuli K')
swap((72, 183, 108, 203), 'Scott', 'Truffle', tail_end=560)
fill_rows(im, (790, 9, 890, 30), 785)  # comment-toolbar bubble
im = im.crop((18, 14, 1156, 279))  # drop the yellow page frame and comment toolbar
im.resize((im.width * 2, im.height * 2), Image.LANCZOS).save('img2/pi_comms.png')
