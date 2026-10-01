"""Placeholder customer details for the New CRM customer-info screenshot (GVoice / Callbox slide)."""
exec(open('/tmp/claude-0/-workspace-training/3ce5cff6-cccf-5f33-b0ee-0b147a55120e/scratchpad/retext_cp.py').read().split('# ---------------- Legacy CP CRM')[0])
FIG = '/root/.fonts/Figtree-500.ttf'


def with_icon(im, textbox, iconbox, old, new, gap=4):
    ic = im.crop(iconbox)
    fill_rows(im, iconbox, iconbox[0] - 1)
    end, _ = replace(im, textbox, old, new, FIG)
    im.paste(ic, (int(end) + gap, iconbox[1]))


I = Image.open(SRC + '22.png').convert('RGB')
replace(I, (51, 138, 127, 155), 'Catherine Garon', 'Anjuli Kintanar', FIG)
with_icon(I, (52, 158, 122, 175), (124, 158, 136, 175), '(917) 609-7742', '555-012345678')
with_icon(I, (63, 178, 235, 195), (237, 178, 248, 196), '9333 Merlot Cir, Breinigsville, PA 18031', '432 HG Street, Meridian, ID')
with_icon(I, (347, 138, 467, 155), (467, 138, 480, 155), 'catherineagill@yahoo.com', 'anjuli_training@gmail.com')
I.save(B + 'scratchpad/img3/callback_crm.png')
print('done')
