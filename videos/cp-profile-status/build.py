import json, subprocess, soundfile as sf
from kokoro_onnx import Kokoro
T='../tts/'
k=Kokoro(T+'kokoro_fp32.onnx', T+'voices.npz')
segs=[
 (0.3, "Here's how to change a CP's profile status."),
 (3.2, "First, make sure you've checked whether the CP should be penalized, since their status affects whether they can claim new jobs."),
 (12.2, "Once you're sure, open their profile in the CP CRM, and click their status tag at the top."),
 (18.6, "Under Change Step, pick the new status. Here, it's no requests, suspend, or CP opt out."),
 (24.9, "On the Change CP step page, type in the CP-facing deactivation reason code."),
 (30.5, "Under how many hours to deactivate, choose Do not auto-reactivate CP."),
 (35.6, "Then click the yellow button, no requests and notify CP."),
 (40.2, "And that's it. The CP's status changes to no requests automatically, and you'll see a note about it in their messages."),
 (49.0, "Thanks for watching!"),
]
out=[]; prev_end=0
for i,(t,txt) in enumerate(segs):
    a,sr=k.create(txt, voice='am_puck', speed=1.0, lang='en-us')
    sf.write(f'r{i}.wav',a,sr); d=len(a)/sr
    flag='  OVERLAP' if t<prev_end+0.15 else ''
    print(f'{t:5.1f} {d:5.2f} ends {t+d:5.2f}{flag}  {txt[:50]}')
    out.append([t,d,txt]); prev_end=t+d
json.dump(out,open('segs.json','w'))
