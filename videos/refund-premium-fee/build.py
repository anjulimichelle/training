import json, subprocess, soundfile as sf
from kokoro_onnx import Kokoro
T='../tts/'
k=Kokoro(T+'kokoro_fp32.onnx', T+'voices.npz')
segs=[
 (0.3, "Here's how to refund a premium fee."),
 (5.8, "First, open Payment history in the customer's CRM, and find the premium fee you need to refund."),
 (12.6, "Here, it's the twenty-dollar premium surcharge. Click its link to open the refund page."),
 (20.8, "On the next page, double-check the amount, then type in your reason for the refund."),
 (34.2, "Next, scroll to the bottom of the page and click the green Refund Charge button."),
 (42.4, "That's it. The premium fee is refunded automatically."),
 (50.2, "To double-check, refresh the CRM and open Payment history again."),
 (59.2, "The refunded premium fee now shows as zero dollars."),
 (64.8, "Thanks for watching!"),
]
out=[]; prev_end=0
for i,(t,txt) in enumerate(segs):
    a,sr=k.create(txt, voice='am_puck', speed=1.0, lang='en-us')
    sf.write(f'r{i}.wav',a,sr); d=len(a)/sr
    flag='  OVERLAP' if t<prev_end+0.15 else ''
    print(f'{t:5.1f} {d:5.2f} ends {t+d:5.2f}{flag}  {txt[:50]}')
    out.append([t,d,txt]); prev_end=t+d
json.dump(out,open('segs.json','w'))
