import requests, urllib.parse, re, html, json, time, sys
H={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/140.0.0.0 Safari/537.36","Accept-Language":"uk"}
S=requests.Session(); S.headers.update(H)
TERMS=["плов реторт пакет 350","плов готовий реторт","рис з овочами реторт пакет","каша рисова реторт пакет","готова страва плов 350 г",
"haidilao рис саморозігрів","hai di lao self heating rice","рис саморозігрів","рис з нагрівачем","самонагрівальний рис",
"henan рис швидкого приготування","рис швидкого приготування 144 г","рис у чаші під окріп","рис швидкого приготування чаша",
"james cook рис","сублімований рис","sublimate рис плов","сублімат плов","adventure menu рис","travellunch рис","mountain house рис",
"харчі рис","їжа в похід рис","fest плов","qiaoshanmei","молочний склад ходорівський каша рисова","макро каша рисова","верес каша рисова","маркел рис",
"bibigo рис","ottogi рис","hetbahn","cj рис готовий","bonduelle рис","tilda рис","seeds of change рис","uncle bens рис 250","ben's original рис",
"рис басматі у пакеті для мікрохвильовки","рис мікрохвильовка пакет 250","рис готовий у пакетику","рис для боулів готовий","рис для суші готовий",
"yatekomo","gallina blanca рис","рис по-корейськи готовий","бібімбап рис","чапче рис","кімчі рис готовий","рис кімчі чаша","смажений рис реторт",
"nasi goreng готовий","рис з курячою грудкою готовий","каша гречана реторт пакет","булгур реторт","кускус готовий реторт"]
def search(term):
    tx=S.get("https://prom.ua/ua/search?search_term="+urllib.parse.quote(term),timeout=45).text
    out=[]
    for blk in re.split(r'data-qaid="product_block"',tx)[1:]:
        blk=blk[:9000]
        hr=re.findall(r'href="(/ua/p\d+[^"]+)"',blk)
        nm=re.search(r'data-qaid="product_name"[^>]*>(.*?)</',blk)
        pr=re.search(r'data-qaid="product_price"[^>]*>(.*?)</',blk)
        cp=re.search(r'data-qaid="company_name"[^>]*>(.*?)</',blk)
        im=re.search(r'<img[^>]+src="(https://images\.prom\.ua[^"]+)"',blk)
        cl=lambda m: html.unescape(re.sub("<[^>]+>","",m.group(1))).strip() if m else None
        if hr and nm: out.append((hr[0].split("?")[0], cl(nm), cl(pr), cl(cp), im.group(1) if im else None))
    return out
seen={}
for t in TERMS:
    try:
        for u,n,p,c,im in search(t): seen.setdefault(u,(n,p,c,t,im))
    except Exception as e: print("ERR",t,str(e)[:60],file=sys.stderr)
    time.sleep(0.2)
json.dump(seen,open('sw4.json','w'),ensure_ascii=False,indent=0)
print(len(seen))
