"""Clean the Callbox / Google Voice screenshots: drop the red labels, swap the customer's name on the call button."""
from PIL import Image, ImageDraw, ImageFont
import numpy as np
B = '/tmp/claude-0/-workspace-training/3ce5cff6-cccf-5f33-b0ee-0b147a55120e/'
LB = '/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf'


def drop_red_label(im, min_y):
    a = np.asarray(im).astype(int)
    red = (a[:, :, 0] > 180) & (a[:, :, 1] < 90) & (a[:, :, 2] < 90)
    red[:min_y] = False
    ys, xs = np.where(red)
    x0, x1, y0, y1 = xs.min() - 3, xs.max() + 3, ys.min() - 3, ys.max() + 3
    d = ImageDraw.Draw(im)
    for y in range(y0, y1 + 1):
        d.line((x0, y, x1, y), fill=im.getpixel((x0 - 5, y)))


def rename_call_button(im, name):
    a = np.asarray(im).astype(int)
    green = (a[:, :, 1] > 150) & (a[:, :, 0] < 130) & (a[:, :, 2] < 130)
    ys, xs = np.where(green)
    x0, x1, y0, y1 = xs.min(), xs.max(), ys.min(), ys.max()
    g = tuple(a[(y0 + y1) // 2, x0 + 3])
    d = ImageDraw.Draw(im)
    d.rectangle((x0 + 2, y0 + 2, x1 - 2, y1 - 2), fill=g)
    f = ImageFont.truetype(LB, 8.5)
    d.text(((x0 + x1) / 2, (y0 + y1) / 2 + 0.5), 'Call ' + name, font=f, fill=(255, 255, 255), anchor='mm')


for src, out, min_y in [('23.png', 'call_primary.png', 150), ('24.png', 'call_secondary.png', 130), ('25.png', 'call_gvoice.png', 400)]:
    im = Image.open(B + 'images/' + src).convert('RGB')
    drop_red_label(im, min_y)
    if src != '25.png':
        rename_call_button(im, 'Anjuli K')
    im.save(B + 'scratchpad/img3/' + out)
print('done')
