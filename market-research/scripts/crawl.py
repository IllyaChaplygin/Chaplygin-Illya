"""Повний обхід карток Prom.ua: URL → сторінка товару → бренд, ціна за 1 шт,
країна, вага, упаковка, кількість, імпортер, відгуки. Результат: listings.json"""
import requests, urllib.parse, re, html, json, time, sys, os
H={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36","Accept-Language":"uk"}
S=requests.Session(); S.headers.update(H)
TERMS=["готовий рис","рис швидкого приготування","рис готовий до вживання","рис у чашці","рис у стакані",
 "стакан відвареного рису","відварений рис","готовий рис ottogi","ottogi рис","bibigo рис","cj рис",
 "ben's original рис","uncle bens рис","рис для мікрохвильовки","рис 2 хвилини","рис у пакеті готовий",
 "henan рис","рис henan","xiao guo zao","yatekomo","gallina blanca рис",
 "саморозігрівальний рис","саморозігрівальний горщик","самонагрівальний рис","haidilao","dacia рис",
 "mo xiao xian","zihaiguo","qiaoshanmei","self heating rice","instant rice",
 "сублімований рис","рис сублімований","сублімована їжа рис","travellunch","trek'n eat","mountain house",
 "adventure menu","adventure food","sublimate","james cook рис","їжа в похід рис","харчі рис","fest food mission",
 "маркел рис","portion рис","пирятинський делікатес рис","clearspring рис","рис з куркою готовий",
 "рис з м'ясом готовий","плов готовий реторт","плов сублімований","каша рисова з м'ясом готова",
 "tilda рис","veetee рис","riso gallo","oryza готовий рис","kupiec рис","sonko рис","hetbahn"]
def search(term,page=1):
    u="https://prom.ua/ua/search?search_term="+urllib.parse.quote(term)+(f"&page={page}" if page>1 else "")
    tx=S.get(u,timeout=45).text; out=[]
    for blk in re.split(r'data-qaid="product_block"',tx)[1:]:
        blk=blk[:9000]
        hrefs=[h for h in re.findall(r'href="(/ua/p\d+[^"]+)"',blk)]
        if hrefs: out.append(hrefs[0].split("?")[0])
    return out
urls={}
for t in TERMS:
    for pg in (1,2):
        try:
            got=search(t,pg)
        except Exception as e:
            print("ERR",t,e,file=sys.stderr); got=[]
        for u in got: urls.setdefault(u,t)
        time.sleep(0.3)
        if len(got)<20: break
    print(f"{t:36} всього {len(urls)}",file=sys.stderr)
json.dump(urls,open("crawl_urls.json","w"),ensure_ascii=False,indent=0)
