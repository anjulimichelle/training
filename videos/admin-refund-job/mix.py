import json,subprocess,sys
out=sys.argv[1]
segs=json.load(open('segs.json'))
inp=['-i','src.mp4']; fl=[]
for i,(t,d,_) in enumerate(segs):
    inp+=['-i',f'r{i}.wav']; ms=int(t*1000); fl.append(f'[{i+1}:a]aresample=48000,adelay={ms}|{ms}[a{i}]')
fl.append(''.join(f'[a{i}]' for i in range(len(segs)))+f'amix=inputs={len(segs)}:normalize=0,volume=0.89,apad[aout]')
subprocess.run(['ffmpeg','-v','error','-y',*inp,'-filter_complex',';'.join(fl),'-map','0:v','-map','[aout]','-c:v','libx264','-preset','medium','-crf','20','-pix_fmt','yuv420p','-c:a','aac','-b:a','160k','-shortest','-movflags','+faststart',out],check=True)
