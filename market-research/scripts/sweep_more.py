import requests, urllib.parse, re, html, json, time, sys
H={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/140.0.0.0 Safari/537.36","Accept-Language":"uk"}
S=requests.Session(); S.headers.update(H)
TERMS=["хетбан рис","hetbahn 햇반","cj hetbahn","готовий рис у чаші","рис у стакані готовий",
 "дончжан рис готовий","пулмуоне рис","dongwon рис","рис덮밥","덮밥 컵밥","컵밥 рис",
 "рис готовий корея стакан","рис з м'ясом реторт","каша рисова реторт пакет",
 "рис готовий 350 г пауч","рис з куркою реторт-пакет","готова страва рис пауч мікрохвильовка",
 "рис швидкого приготування чаша","instant rice cup","ready to eat rice ukraine",
 "рис готовий до вживання 250 г","veetee рис готовий","tilda steamed basmati",
 "riso scotti pronto","uncle bens express рис 250","рис express мікрохвильовка 2 хвилини",
 "рис із овочами готовий пауч","рис басматі готовий пауч","рис жасмин готовий пауч",
 "рис з карі готовий пауч","гарнір рис готовий упаковка"]
def search(term):
    tx=S.get("https://prom.ua/ua/search?search_term="+urllib.parse.quote(term),timeout=45).text
    out=[]
    for blk in re.split(r'data-qaid="product_block"',tx)[1:]:
        blk=blk[:9000]
        hr=re.findall(r'href="(/ua/p\d+[^"]+)"',blk)
        nm=re.search(r'data-qaid="product_name"[^>]*>(.*?)</',blk)
        pr=re.search(r'data-qaid="product_price"[^>]*>(.*?)</',blk)
        cp=re.search(r'data-qaid="company_name"[^>]*>(.*?)</',blk)
        cl=lambda m: html.unescape(re.sub("<[^>]+>","",m.group(1))).strip() if m else None
        if hr and nm: out.append((hr[0].split("?")[0], cl(nm), cl(pr), cl(cp)))
    return out
seen={}
for t in TERMS:
    try:
        for u,n,p,c in search(t): seen.setdefault(u,(n,p,c,t))
    except Exception as e: print("ERR",t,str(e)[:40],file=sys.stderr)
    time.sleep(0.25)
KNOWN=re.compile(r"ottogi|оттогі|ben'?s|бенс|uncle ben|bibigo|clearspring|adventure menu|"
                 r"маркел|мартел|верес|макро|ходорів|portion|hapay|appetit|м'ясторія|food fabrika",re.I)
RICE=re.compile(r"рис|rice|밥|плов|каша рисова",re.I)
BAD=re.compile(r"окроп|сублім|freeze|самонагр|саморозігр|заморож|крупа|сирий|0[,.]5 ?кг|1 ?кг|"
               r"вармішель|локшина|локшин|локшинка|локшину|м'ясо|рамен|лапша|noodle|соус|приправ",re.I)
out=[]
for u,(n,p,c,t) in seen.items():
    if not n or not RICE.search(n): continue
    if KNOWN.search(n): continue
    if BAD.search(n): continue
    out.append((p,n,c,u))
json.dump(out,open("sweep_more.json","w"),ensure_ascii=False,indent=1)
print("всього карток:",len(seen),"нових кандидатів:",len(out))
for p,n,c,u in sorted(out): print(p,"|",n[:80],"|",c)
