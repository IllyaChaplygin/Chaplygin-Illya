import requests, urllib.parse, re, html, json, time
H={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36","Accept-Language":"uk"}

def cards(term):
    tx=requests.get("https://prom.ua/ua/search?search_term="+urllib.parse.quote(term),
                    headers=H,timeout=40).text
    out=[]
    for blk in re.split(r'data-qaid="product_block"', tx)[1:]:
        blk=blk[:6000]
        nm=re.search(r'data-qaid="product_name"[^>]*>(.*?)</', blk)
        if not nm: continue
        name=html.unescape(re.sub("<[^>]+>","",nm.group(1))).strip()
        pr=re.search(r'data-qaid="product_price"[^>]*>(.*?)</span>', blk, re.S)
        if not pr:
            pr=re.search(r'>(\d[\d\s ]{1,9})\s*(?:грн|₴)', blk)
        price=None
        if pr:
            t=html.unescape(re.sub("<[^>]+>","",pr.group(1)))
            d=re.sub(r"[^\d]","",t)
            price=int(d) if d else None
        sel=re.search(r'data-qaid="company_name"[^>]*>(.*?)</', blk)
        seller=html.unescape(re.sub("<[^>]+>","",sel.group(1))).strip() if sel else None
        out.append({"name":name,"price":price,"seller":seller})
    return out

TERMS=["готовий рис ottogi","рис ottogi відварений","стакан відвареного рису ottogi",
       "bibigo рис 210","clearspring рис tamari","саморозігрівальний рис haidilao",
       "mo xiao xian рис","instant rice китай 275","travellunch рис",
       "ben's original рис пауч","готовий рис пауч 250"]
res={}
for t in TERMS:
    try:
        res[t]=cards(t)
        print(f"--- {t}  ({len(res[t])})")
        for c in res[t][:10]:
            print(f"    {str(c['price']):>6} грн | {str(c['seller'])[:18]:18} | {c['name'][:78]}")
    except Exception as e: print(t,"ERR",str(e)[:60])
    time.sleep(0.4)
json.dump(res, open("prom_cards.json","w"), ensure_ascii=False, indent=1)
