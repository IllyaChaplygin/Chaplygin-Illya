import re,csv,json,urllib.request,time,sys
rows=[r for r in list(csv.reader(open('data/sheet.csv',encoding='utf-8')))[2:] if any(r)]
fid=lambda u: re.search(r'folders/([\w-]+)',u).group(1)
cache={}
def ls(f):
    if f in cache: return cache[f]
    for _ in range(3):
        try:
            h=urllib.request.urlopen(urllib.request.Request('https://drive.google.com/embeddedfolderview?id='+f,headers={'User-Agent':'Mozilla/5.0'}),timeout=30).read().decode()
            break
        except Exception as e: time.sleep(2); h=''
    items=[]
    for m in re.finditer(r'<a href="https://drive.google.com/(file/d|drive/folders)/([\w-]+)[^"]*"(.*?)</a>',h,re.S):
        t=re.search(r'flip-entry-title">([^<]*)',m.group(3))
        items.append(('d' if m.group(1)!='file/d' else 'f',m.group(2),(t.group(1) if t else '')))
    cache[f]=items; return items
def walk(f,depth=0):
    out=[]
    for k,i,t in ls(f):
        if k=='f': out.append((i,t))
        elif depth<4: out+=walk(i,depth+1)
    return out
res={}
for f in sorted({r[22] for r in rows if r[22]}):
    res[f]=walk(fid(f)); print(len(res[f]),f[-45:-30],flush=True)
json.dump(res,open('files.json','w'),ensure_ascii=False)
