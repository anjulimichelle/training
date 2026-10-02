"""Placeholder customer name on the Update Max Rate screenshots."""
import sys; sys.path.insert(0, '.')
from PIL import Image
from retext_cp import replace
B = '/tmp/claude-0/-workspace-training/3ce5cff6-cccf-5f33-b0ee-0b147a55120e/'
up = lambda im: im.resize((im.width * 2, im.height * 2), Image.LANCZOS)
f = Image.open(B + 'images/34.png').convert('RGB')
replace(f, (140, 58, 250, 80), 'Rafaella Nicoletti', 'Anjuli Kintanar')
f.crop((0, 0, 963, 265)).save('img3/maxrate_form.png')
up(Image.open(B + 'images/32.png').convert('RGB')).save('img3/maxrate_do_btn.png')
Image.open(B + 'images/33.png').convert('RGB').crop((0, 50, 597, 325)).save('img3/maxrate_do_menu.png')
