import requests, urllib.parse, re, html, json, time, sys
H={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36","Accept-Language":"uk"}
S=requests.Session(); S.headers.update(H)
TERMS=["ottogi осьминог рис","ottogi рис восьминіг","ottogi гамбурзький стейк рис",
       "ottogi jjampong рис","ottogi чжамппонг","ottogi кімчі рис 310",
       "ben's original bio basmati 240","бенс оріджинал 240","ben's sticky bowl",
       "ben's original липкий рис","ben's original 240 г рис","ben's original express bio",
       "рис ottogi стакан","ottogi cooked rice","edison lee ben's"]
def search(term):
    tx=S.get("https://prom.ua/ua/search?search_term="+urllib.parse.quote(term),timeout=45).text
    out=[]
    for blk in re.split(r'data-qaid="product_block"',tx)[1:]:
        blk=blk[:9000]
        hr=re.findall(r'href="(/ua/p\d+[^"]+)"',blk)
        nm=re.search(r'data-qaid="product_name"[^>]*>(.*?)</',blk)
        if hr and nm:
            out.append((hr[0].split("?")[0], html.unescape(re.sub("<[^>]+>","",nm.group(1))).strip()))
    return out
urls={}
for t in TERMS:
    try:
        for u,n in search(t): urls.setdefault(u,(t,n))
    except Exception as e: print("ERR",t,str(e)[:60],file=sys.stderr)
    time.sleep(0.3)
print("карток:",len(urls),file=sys.stderr)
KEEP=re.compile(r"ottogi|оттогі|ben'?s|бенс|uncle ben|bibigo|бібіго",re.I)
rows=[]
for u,(term,nm) in urls.items():
    if not KEEP.search(nm): continue
    try: r=S.get("https://prom.ua"+u,timeout=45); t=r.text
    except Exception: continue
    m=re.search(r"<title[^>]*>(.*?)</title>",t,re.S); title=html.unescape(m.group(1)) if m else ""
    mp=re.search(r"ціна\s*([\d\s]+(?:[.,]\d+)?)\s*₴",title)
    price=float(mp.group(1).replace(" ","").replace(" ","").replace(",",".")) if mp else None
    seller=re.search(r'"seller":\{"@type":"Organization","name":"([^"]+)"',t)
    img=re.findall(r'https://images\.prom\.ua/[^"\\ ]+?_w\d+_[^"\\ ]+?\.(?:jpg|jpeg|png|webp)',t)
    rows.append({"url":r.url,"name":nm,"price":price,
                 "seller":html.unescape(seller.group(1)) if seller else None,
                 "img":img[0] if img else None,"term":term})
    time.sleep(0.2)
json.dump(rows,open("recheck.json","w"),ensure_ascii=False,indent=1)
for x in sorted(rows,key=lambda r:(r['name'] or '')):
    print(x["price"],"|",(x["name"] or "")[:78],"|",x["seller"])
