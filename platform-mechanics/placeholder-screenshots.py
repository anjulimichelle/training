"""Placeholder names for the job history screenshots: customer Anjuli Kintanar, cleaner Truffle Wuffle."""
exec(open('/tmp/claude-0/-workspace-training/3ce5cff6-cccf-5f33-b0ee-0b147a55120e/scratchpad/retext_cp.py').read().split('# ---------------- Legacy CP CRM')[0])
OUT3 = B + 'scratchpad/img3/'

# Jobs panel with the "i" icon that opens the job history
J = Image.open(SRC + '20.png').convert('RGB')
replace(J, (785, 93, 822, 106), 'Lilliam S', SHORT)
for y in (211, 239, 266, 301):
    replace(J, (570, y - 3, 608, y + 9), 'Lilliam S', SHORT)
replace(J, (676, 286, 705, 299), 'Lilliam', FIRST)
J.save(OUT3 + 'jh_open.png')

# The job history table itself
H = Image.open(SRC + '21.png').convert('RGB')
for r in (114, 192, 270, 347, 425, 503):
    replace(H, (404, r + 36, 456, r + 49), 'Brittany', 'Anjuli')
    replace(H, (404, r + 49, 456, r + 62), 'Simpson', 'Kintanar')
for r in (114, 192, 270, 347, 425):
    replace(H, (827, r + 10, 878, r + 23), 'Lilliam', FIRST)
    replace(H, (827, r + 23, 878, r + 36), 'Soto', 'Wuffle')
H.save(OUT3 + 'jh_table.png')
print('done')
