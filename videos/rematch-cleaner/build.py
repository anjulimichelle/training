import json, subprocess, soundfile as sf
from kokoro_onnx import Kokoro
T='../tts/'
k=Kokoro(T+'kokoro_fp32.onnx', T+'voices.npz')
segs=[
 (0.5, "Here's a quick look at how to rematch a cleaner."),
 (4.3, "In the customer's CRM, find the job you want to rematch. It can be a claimed job, or a pending invoice job."),
 (12.0, "Here, we'll open Claimed and Submitted, then click the job ID to open the job admin page."),
 (21.5, "From the job admin page, click the Do button."),
 (29.5, "Scroll to Change CP, and choose whether you want to ban the customer and cleaner pair."),
 (38.8, "Then click the blue Change CP button."),
 (44.5, "This removes the matched cleaner automatically. Back on the job admin page, the status now shows Submitted, so available cleaners can claim it."),
 (54.2, "Thanks for watching!"),
]
out=[]; prev_end=0
for i,(t,txt) in enumerate(segs):
    a,sr=k.create(txt, voice='am_puck', speed=1.0, lang='en-us')
    sf.write(f'r{i}.wav',a,sr); d=len(a)/sr
    flag='  OVERLAP' if t<prev_end+0.15 else ''
    print(f'{t:5.1f} {d:5.2f} ends {t+d:5.2f}{flag}  {txt[:50]}')
    out.append([t,d,txt]); prev_end=t+d
json.dump(out,open('segs.json','w'))
