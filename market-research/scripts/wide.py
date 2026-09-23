"""Широкий пошук готового рису кімнатного зберігання під розігрів."""
import requests, urllib.parse, re, html, json, time, sys
H={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36","Accept-Language":"uk"}
S=requests.Session(); S.headers.update(H)
TERMS=[
 # гарнір
 "готовий рис пауч","рис готовий до вживання","рис express мікрохвильовка","рис 2 хвилини пауч",
 "рис у пакеті готовий 250","microwave rice pouch","рис припущений готовий","рис відварений пауч",
 "ben's original","vitasia рис","riso gallo рис готовий","sonko рис готовий","oryza рис готовий",
 "tilda рис готовий","veetee рис готовий","rizopia","рис жасмин готовий пауч",
 # чаша
 "готовий рис стакан","рис у чаші готовий","cj hetbahn","hetbahn 210","bibigo рис",
 "ottogi рис","sajo рис","pulmuone рис","рис корея готовий 210","рис японський готовий",
 # український реторт
 "каша рисова реторт","каша рисова з м'ясом 350","рис з м'ясом реторт пакет","рис з куркою реторт",
 "рис зі свининою реторт","рис з яловичиною реторт","готова страва рис реторт","розігрій та їж рис",
 "сухпайок рис готовий","ірп рис","сніданок туриста рис","армійський пайок рис",
 "каша рисова верес","каша рисова макро","каша рисова маркел","каша рисова пак","каша рисова гурман",
 "каша рисова ходорівський","каша рисова м'ясна","рисова каша консерва розігрій",
 "food for you рис","portion рис","їdlo рис","добра їжа рис","смачна допомога рис",
 "рис по-домашньому готовий","плов реторт пакет готовий","плов готовий пауч 350",
 # імпорт реторт
 "adventure menu рис ready to eat","clearspring рис","natura рис готовий","tanoshi рис",
 "rice ready to eat ukraine","реторт пакет рис імпорт",
]
def cards(term,page=1):
    u="https://prom.ua/ua/search?search_term="+urllib.parse.quote(term)+(f"&page={page}" if page>1 else "")
    tx=S.get(u,timeout=45).text
    out=[]
    for blk in re.split(r'data-qaid="product_block"',tx)[1:]:
        blk=blk[:9000]
        nm=re.search(r'data-qaid="product_name"[^>]*>(.*?)</',blk)
        pr=re.search(r'data-qaid="product_price"[^>]*>(.*?)</span>',blk,re.S)
        sel=re.search(r'data-qaid="company_name"[^>]*>(.*?)</',blk)
        hr=re.findall(r'href="(/ua/p\d+[^"]+)"',blk)
        img=re.findall(r'(https://images\.prom\.ua/[^"\'\\ )]+_w400_[^"\'\\ )]+)',blk)
        if not nm: continue
        n=html.unescape(re.sub("<[^>]+>","",nm.group(1))).strip()
        p=re.sub(r"[^\d]","",html.unescape(re.sub("<[^>]+>","",pr.group(1)))) if pr else None
        s=html.unescape(re.sub("<[^>]+>","",sel.group(1))).strip() if sel else None
        out.append({"name":n,"price":int(p) if p else None,"seller":s,
                    "url":hr[0].split("?")[0] if hr else None,"img":img[0] if img else None,"term":term})
    return out
res={}
for t in TERMS:
    for pg in (1,2):
        try: got=cards(t,pg)
        except Exception as e: print("ERR",t,str(e)[:40],file=sys.stderr); got=[]
        for c in got: res[(c["name"],c["seller"])]=c
        time.sleep(0.28)
        if len(got)<20: break
    print(f"{t:38} {len(res)}",file=sys.stderr)
json.dump(list(res.values()),open("wide.json","w"),ensure_ascii=False,indent=1)
print("усього карток",len(res))
