import json, subprocess, soundfile as sf
from kokoro_onnx import Kokoro
T='../tts/'
k=Kokoro(T+'kokoro_fp32.onnx', T+'voices.npz')
segs=[
 (0.3, "Here's how to invoice a pending invoice job from the customer's dashboard."),
 (5.0, "First, find the job you want to invoice. Here, it's the March 2 appointment with CP Miaka, under Pending Invoice Jobs."),
 (13.0, "Check that the cleaning duration the customer confirmed matches what was booked."),
 (18.4, "Once that's confirmed, click Login as Customer to copy their login link, and open it in an incognito window."),
 (35.6, "On the customer's dashboard, under Past cleanings, select the appointment you want to invoice."),
 (44.2, "On the next page, scroll down and click Report an issue."),
 (50.4, "You'll see a prompt about this appointment. Click the yellow button that says, My cleaning was completed."),
 (60.4, "That's it. This automatically invoices the customer for the cleaning."),
 (69.5, "To double-check, head back to the customer CRM and open Completed jobs."),
 (78.0, "The March 2 appointment now shows as invoiced."),
 (85.0, "Thanks for watching!"),
]
out=[]; prev_end=0
for i,(t,txt) in enumerate(segs):
    a,sr=k.create(txt, voice='am_puck', speed=1.0, lang='en-us')
    sf.write(f'r{i}.wav',a,sr); d=len(a)/sr
    flag='  OVERLAP' if t<prev_end+0.15 else ''
    print(f'{t:5.1f} {d:5.2f} ends {t+d:5.2f}{flag}  {txt[:50]}')
    out.append([t,d,txt]); prev_end=t+d
json.dump(out,open('segs.json','w'))
