"""Placeholder details for the extra CP CRM screenshots (status panel, Change Step menu, header, Change CP step modal)."""
exec(open('/tmp/claude-0/-workspace-training/3ce5cff6-cccf-5f33-b0ee-0b147a55120e/scratchpad/retext_cp.py').read().split('# ---------------- Legacy CP CRM')[0])

# Status / reliability panel
P = Image.open(SRC + '14.png').convert('RGB')
replace(P, (130, 94, 185, 113), 'regianen', HANDLE)
P.save(OUT + 'cp_reliability.png')

# Change Step menu: drop the customer names listed behind the menu
M = Image.open(SRC + '15.png').convert('RGB')
replace(M, (123, 13, 178, 33), 'regianen', HANDLE)
M = M.crop((0, 0, 756, 356))
fill_rows(M, (168, 286, 755, 355), 755)
fill_rows(M, (0, 284, 19, 355), 0)
M.save(OUT + 'cp_changestep.png')

# CP header: name, photo; crop off the cut-off "Call ..." button row
H = Image.open(SRC + '16.png').convert('RGB')
replace(H, (109, 19, 224, 42), 'Regiane Nunes', NAME)
f, m = face(*CARTOON, (70, 92))
H.paste(f, (30, 21), m)
H = H.crop((0, 14, 795, 371))
H.save(OUT + 'cp_header.png')
H.crop((0, 0, 560, 76)).save(OUT + 'cp_header_top.png')  # name and banners only

# Change CP step modal: sample customer name in the reason, crop to the modal
X = Image.open(SRC + '17.png').convert('RGB')
end, x1 = replace(X, (172, 71, 220, 89), 'Meynard Q', 'Alice')
move(X, (221, 71, 480, 89), int(end) + 4 - 222)
X = X.crop((36, 8, 660, 800))
X.save(OUT + 'cp_stepmodal.png')
print('done')
