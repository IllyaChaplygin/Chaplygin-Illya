"""Жива перевірка: готовий рис кімнатного зберігання під мікрохвильовку.
Для кожної картки — роздрібна ціна за 1 шт, маса, продавець, URL, дата."""
import requests, urllib.parse, re, html, json, time, sys
H={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36","Accept-Language":"uk"}
S=requests.Session(); S.headers.update(H)
TERMS=["ben's original рис","бенс оріджинал рис","ben's original express","рис express 220",
       "ottogi рис готовий","стакан відвареного рису ottogi","ottogi 오뚜기 밥",
       "bibigo рис 210","cj hetbahn рис","hetbahn","готовий рис корея 210",
       "clearspring рис тамарі","clearspring brown wild rice",
       "portion рис з куркою","маркел каша рисова","adventure menu chicken korma рис",
       "рис готовий до вживання мікрохвильовка","готовий рис реторт пауч",
       "рис у реторт-пакеті готовий","рис 90 секунд мікрохвильовка","microwave rice 2 minutes",
       "veetee рис","tilda рис мікрохвильовка","uncle bens express"]
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
    except Exception as e: print("ERR",t,str(e)[:40],file=sys.stderr)
    time.sleep(0.3)
print("карток знайдено:",len(urls),file=sys.stderr)
rows=[]
for i,(u,(term,nm)) in enumerate(urls.items()):
    try: r=S.get("https://prom.ua"+u,timeout=45); t=r.text
    except Exception: continue
    m=re.search(r"<title>(.*?)</title>",t,re.S); title=html.unescape(m.group(1)) if m else ""
    mp=re.search(r"ціна\s*([\d\s]+(?:[.,]\d+)?)\s*₴",title)
    price=float(mp.group(1).replace(" ","").replace(",",".")) if mp else None
    txt=re.sub(r"\s+"," ",html.unescape(re.sub(r"<[^>]+>"," ",t)))
    k=txt.find("Характеристики та опис"); ch=txt[k:k+1400] if k>=0 else ""
    def f(lb):
        mm=re.search(lb+r" (.+?)(?= Виробник| Країна| Вага| Склад| Тип| Упаковка| Термін| Вид| Смак| Особлив| Кількість| Користув|$)",ch)
        return mm.group(1).strip()[:40] if mm else None
    brand=re.search(r'"brand":\{"@type":"Brand","name":"([^"]+)"',t)
    seller=re.search(r'"seller":\{"@type":"Organization","name":"([^"]+)"',t)
    tiers=sorted({(int(q),float(p)) for p,q in re.findall(r'"price":"([\d.]+)","priceCurrencyLocalized":"[^"]*","minimumOrderQuantity":"(\d+)"',t)})
    rows.append({"url":r.url,"name":nm,"price1":price,"tiers":tiers,
                 "brand":html.unescape(brand.group(1)) if brand else None,
                 "seller":html.unescape(seller.group(1)) if seller else None,
                 "maker":f("Виробник"),"country":f("Країна виробник"),"weight":f("Вага"),
                 "pack":f("Упаковка"),"shelf":f("Термін придатності"),
                 "micro":bool(re.search(r"мікрохвил|microwave|мікрохвильов",txt,re.I)),
                 "boil":bool(re.search(r"залийте окропом|залити окропом|окропом",txt,re.I)),
                 "term":term})
    if i%20==0: print(i,file=sys.stderr)
    time.sleep(0.22)
json.dump(rows,open("verify.json","w"),ensure_ascii=False,indent=1)
print("готово",len(rows))
