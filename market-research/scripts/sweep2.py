"""Поглиблений добір: Henan, Yatekomo, сублімація, саморозігрів — з фото."""
import requests, urllib.parse, re, html, json, time
H={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36","Accept-Language":"uk"}
def cards(term, page=1):
    u="https://prom.ua/ua/search?search_term="+urllib.parse.quote(term)
    if page>1: u+=f"&page={page}"
    tx=requests.get(u,headers=H,timeout=45).text
    out=[]
    for blk in re.split(r'data-qaid="product_block"',tx)[1:]:
        blk=blk[:9000]
        nm=re.search(r'data-qaid="product_name"[^>]*>(.*?)</',blk)
        if not nm: continue
        name=html.unescape(re.sub("<[^>]+>","",nm.group(1))).strip()
        pr=re.search(r'data-qaid="product_price"[^>]*>(.*?)</span>',blk,re.S); price=None
        if pr:
            d=re.sub(r"[^\d]","",html.unescape(re.sub("<[^>]+>","",pr.group(1))))
            price=int(d) if d else None
        sel=re.search(r'data-qaid="company_name"[^>]*>(.*?)</',blk)
        seller=html.unescape(re.sub("<[^>]+>","",sel.group(1))).strip() if sel else None
        imgs=re.findall(r'(https://images\.prom\.ua/[^"\'\\ )]+_w400_[^"\'\\ )]+)',blk)
        out.append({"name":name,"price":price,"seller":seller,"img":imgs[0] if imgs else None,"term":term})
    return out

TERMS={
 "henan":["henan рис 174","рис henan гребінці","рис henan курка","рис henan яловичина",
          "рис henan пібімбап","henan cantonese sausage","henan braised chicken",
          "henan pepper minced chicken","henan scallops mushroom","henan гостра курка 144",
          "рис швидкого приготування henan","henan ковбаса рис"],
 "yatekomo":["yatekomo рис","gallina blanca yarroz","рис yatekomo теріякі","yarroz yatekomo яловичина"],
 "jamescook":["james cook рис","james cook машкічірі","james cook рис з м'ясом"],
 "yizhavpohid":["їжа в похід рис","їжа в похід рис з куркою"],
 "adventure":["adventure food рис карі","adventure menu рис басматі","mountain house рис карі"],
 "kharchi":["харчі рис зі свининою","харчі сублімований рис"],
 "skarb":["самонагрівний рис скарб","self heating taiwanese braised pork рис","self heating double pepper beef"],
 "haidilao2":["haidilao свиняча грудинка гірчична зелень рис","haidilao рис 1190"],
 "travellunch2":["travellunch курка карі з рисом 125"],
}
res={}
for b,ts in TERMS.items():
    seen,items=set(),[]
    for t in ts:
        try:
            for c in cards(t):
                k=(c["name"],c["seller"])
                if k in seen: continue
                seen.add(k); items.append(c)
        except Exception as e: print(b,t,"ERR",str(e)[:50])
        time.sleep(0.35)
    res[b]=items
    print(f"{b:14} {len(items):3} карток, з фото {sum(1 for i in items if i['img'])}")
json.dump(res,open("sweep2.json","w"),ensure_ascii=False,indent=1)
