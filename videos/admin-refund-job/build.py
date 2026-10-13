import json, subprocess, soundfile as sf
from kokoro_onnx import Kokoro
T='../tts/'
k=Kokoro(T+'kokoro_fp32.onnx', T+'voices.npz')
segs=[
 (0.5, "Here's a quick look at how to admin refund a job charge."),
 (4.5, "First, open Completed jobs in the customer's CRM, and find the job that needs a refund."),
 (11.0, "Click its job ID. This opens the job admin page."),
 (22.0, "From here, click the Do button, and select Refund Job."),
 (29.3, "Tick both agent acknowledgements."),
 (37.0, "Under Refund Dollar Amount, type in the amount."),
 (45.0, "Select the options that apply, then add a reason for the admin refund."),
 (57.5, "Then scroll down and click the red Issue Refund button. The refund processes automatically."),
 (70.0, "To double-check, refresh the CRM. The refund should show on the job admin page."),
 (78.5, "And that's it. Thanks for watching!"),
]
out=[]; prev_end=0
for i,(t,txt) in enumerate(segs):
    a,sr=k.create(txt, voice='am_puck', speed=1.0, lang='en-us')
    sf.write(f'r{i}.wav',a,sr); d=len(a)/sr
    flag='  OVERLAP' if t<prev_end+0.15 else ''
    print(f'{t:5.1f} {d:5.2f} ends {t+d:5.2f}{flag}  {txt[:50]}')
    out.append([t,d,txt]); prev_end=t+d
json.dump(out,open('segs.json','w'))
