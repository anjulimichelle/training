"""Cut the per-type voucher snippets out of the KB composites (images 35, 36) and drop the KB's red highlight boxes."""
from PIL import Image
import numpy as np
B = '/tmp/claude-0/-workspace-training/3ce5cff6-cccf-5f33-b0ee-0b147a55120e/images/'
CROPS = {
    '35': {'vt_free_new': (68, 82, 545, 119), 'vt_free_legacy': (67, 207, 245, 221),
           'vt_dhj_new': (69, 346, 551, 383), 'vt_dhj_legacy': (72, 487, 311, 504),
           'vt_gift': (76, 613, 440, 630)},
    '36': {'vt_groupon': (61, 116, 250, 133), 'vt_second': (64, 209, 384, 229),
           'vt_admin_new': (65, 302, 665, 346), 'vt_admin_legacy': (81, 377, 416, 395),
           'vt_giftcard': (73, 447, 285, 468)},
}


def unred(im):
    a = np.asarray(im).astype(int)
    r, g, b = a[..., 0], a[..., 1], a[..., 2]
    red = ((r > 150) & (g < 140) & (b < 140) & (r - g > 70) & (r - b > 70)) | ((r > 150) & (r - g > 25) & (np.abs(g - b) < 20))
    out = a.copy()
    for y, x in zip(*np.where(red)):
        for d in range(1, 10):  # nearest non-red pixel in any direction
            hit = [(y + dy, x + dx) for dy, dx in ((0, d), (0, -d), (d, 0), (-d, 0))
                   if 0 <= y + dy < a.shape[0] and 0 <= x + dx < a.shape[1] and not red[y + dy, x + dx]]
            if hit:
                out[y, x] = a[hit[0]]; break
    return Image.fromarray(out.astype('uint8'))


for src, crops in CROPS.items():
    im = Image.open(B + src + '.png').convert('RGB')
    for name, box in crops.items():
        c = im.crop(box)
        if name in ('vt_gift', 'vt_giftcard', 'vt_groupon'):  # only these carry the KB's red boxes
            c = unred(c)
        c.resize((c.width * 3, c.height * 3), Image.LANCZOS).save(f'img3/{name}.png')
