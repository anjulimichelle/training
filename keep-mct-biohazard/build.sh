#!/bin/sh
# Assemble deck8.js (shared helpers + deck2 + deck6 helpers + deck7 helpers + this deck), build, then add click animations.
set -e
cd "$(dirname "$0")"
python3 - <<'PY'
base = open('deck.js').read(); base = base[:base.index('(async () => {')]
base = base.replace("pres.title = 'Customer Foundations & the C CRM';", "pres.title = 'Keep MCT & Biohazard Reports';")
base = base.replace("path.join(__dirname, '../img')", "path.join(__dirname, '../img8')")
base = base.replace("const TOPICS = 6;", "const TOPICS = 2;")
d2 = open('deck2_body.js').read(); d2 = d2[:d2.index('(async () => {')]
d6 = open('deck6_body.js').read(); d6 = d6[:d6.index('(async () => {')]
d7 = open('deck7_main.js').read(); d7 = d7[:d7.index('(async () => {')]
open('deck8.js', 'w').write(base + d2 + d6 + d7 + open('deck8_body.js').read())
PY
node deck8.js /workspace/training/.qa/r8_0.pptx
python3 add_story_anim.py /workspace/training/.qa/r8_0.pptx /workspace/training/.qa/r8.pptx
