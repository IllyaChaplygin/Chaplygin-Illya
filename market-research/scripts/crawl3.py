"""Сторінка товару Prom → роздрібна ціна за 1 шт, оптові пороги, «N+ купили», відгуки,
бренд, країна, вага, упаковка, імпортер. Вхід: crawl_urls.json (+ extra_urls.json)."""
import requests, re, html, json, time, sys, os
H={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36","Accept-Language":"uk"}
S=requests.Session(); S.headers.update(H)
urls=json.load(open("crawl_urls.json"))
if os.path.exists("extra_urls.json"): urls.update(json.load(open("extra_urls.json")))
def parse(u,t):
    rec={"url":u}
    for m in re.findall(r'<script type="application/ld\+json">(.*?)</script>',t,re.S):
        try: j=json.loads(m)
        except Exception: continue
        if isinstance(j,dict) and j.get("@type")=="Product":
            rec["name"]=html.unescape(j.get("name") or ""); rec["desc"]=html.unescape(j.get("description") or "")
            rec["brand"]=html.unescape((j.get("brand") or {}).get("name") or "")
            o=j.get("offers") or {}
            if isinstance(o,list): o=o[0]
            rec["ld_price"]=o.get("price"); rec["seller"]=html.unescape((o.get("seller") or {}).get("name") or "")
            imgs=j.get("image") or []; rec["img"]=imgs[0] if imgs else None
    m=re.search(r'"minimumOrderQuantity":"","priceCurrencyLocalized":"[^"]*","price":"([\d.]+)"',t)
    rec["price1"]=float(m.group(1)) if m else (float(rec["ld_price"]) if rec.get("ld_price") else None)
    rec["tiers"]=sorted({(int(q),float(p)) for p,q in re.findall(r'"price":"([\d.]+)","priceCurrencyLocalized":"[^"]*","minimumOrderQuantity":"(\d+)"',t)})
    txt=re.sub(r"\s+"," ",html.unescape(re.sub(r"<[^>]+>"," ",t)))
    m=re.search(r"(\d[\d\s]*)\+?\s*купил",txt); rec["bought"]=int(m.group(1).replace(" ","")) if m else 0
    m=re.search(r"(\d\.\d)\s*\((\d+)\)",txt[:20000]); rec["rating"],rec["reviews"]=(float(m.group(1)),int(m.group(2))) if m else (None,0)
    k=txt.find("Характеристики та опис"); ch=txt[k:k+1200] if k>=0 else ""
    rec["chars"]=ch
    def field(lb):
        mm=re.search(lb+r" (.+?)(?= Виробник| Країна| Вага| Об'єм| Склад| Тип| Упаковка| Термін| Вид| Смак| Особливості| Кількість| Користувацькі|$)",ch)
        return mm.group(1).strip()[:60] if mm else None
    rec["maker"]=field("Виробник"); rec["country"]=field("Країна виробник"); rec["weight"]=field("Вага")
    rec["pack"]=field("Упаковка"); rec["shelf"]=field("Термін придатності")
    imp=re.findall(r"[Іі]мпортер[^.;\n]{0,120}",rec.get("desc","")+" "+txt)
    rec["importer"]=imp[:3]
    return rec
out=[]
for i,(u,term) in enumerate(urls.items()):
    try:
        r=S.get("https://prom.ua"+u if u.startswith("/") else u,timeout=45)
        rec=parse(r.url,r.text); rec["term"]=term; out.append(rec)
    except Exception as e:
        print("ERR",u,e,file=sys.stderr)
    if i%40==0: print(i,file=sys.stderr)
    time.sleep(0.25)
json.dump(out,open("listings3.json","w"),ensure_ascii=False,indent=1)
print("done",len(out))
