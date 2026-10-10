import json,os,io,urllib.request,time
from concurrent.futures import ThreadPoolExecutor
from PIL import Image
d=json.load(open('files.json'))
ids=[i for v in d.values() for i,t in v]
os.makedirs('/home/user/Chaplygin-Illya/docs/photos',exist_ok=True)
def get(i):
    big=f'/home/user/Chaplygin-Illya/docs/photos/{i}.jpg'; sm=f'/home/user/Chaplygin-Illya/docs/photos/{i}_s.jpg'
    if os.path.exists(sm): return 1
    for a in range(4):
        try:
            b=urllib.request.urlopen(urllib.request.Request(f'https://drive.google.com/thumbnail?id={i}&sz=w1600',headers={'User-Agent':'Mozilla/5.0'}),timeout=60).read()
            im=Image.open(io.BytesIO(b)).convert('RGB')
            im.save(big,quality=80,optimize=True,progressive=True)
            s=im.copy(); s.thumbnail((640,640)); s.save(sm,quality=78,optimize=True,progressive=True)
            return 1
        except Exception as e: time.sleep(2*(a+1))
    print('FAIL',i,flush=True); return 0
with ThreadPoolExecutor(6) as ex: ok=sum(ex.map(get,ids))
print('ok',ok,'of',len(ids))
