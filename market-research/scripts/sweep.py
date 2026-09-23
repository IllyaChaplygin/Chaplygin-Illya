"""Широкий пошук по Prom.ua: які ще бренди готового рису є на ринку."""
import requests, urllib.parse, re, html, json, time, collections
H={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36","Accept-Language":"uk"}

def cards(term, page=1):
    u = "https://prom.ua/ua/search?search_term=" + urllib.parse.quote(term)
    if page > 1: u += f"&page={page}"
    tx = requests.get(u, headers=H, timeout=45).text
    out=[]
    for blk in re.split(r'data-qaid="product_block"', tx)[1:]:
        blk = blk[:9000]
        nm = re.search(r'data-qaid="product_name"[^>]*>(.*?)</', blk)
        if not nm: continue
        name = html.unescape(re.sub("<[^>]+>","",nm.group(1))).strip()
        pr = re.search(r'data-qaid="product_price"[^>]*>(.*?)</span>', blk, re.S)
        price=None
        if pr:
            d = re.sub(r"[^\d]","", html.unescape(re.sub("<[^>]+>","",pr.group(1))))
            price = int(d) if d else None
        sel = re.search(r'data-qaid="company_name"[^>]*>(.*?)</', blk)
        seller = html.unescape(re.sub("<[^>]+>","",sel.group(1))).strip() if sel else None
        imgs = re.findall(r'(https://images\.prom\.ua/[^"\'\\ )]+_w400_[^"\'\\ )]+)', blk)
        out.append({"name":name,"price":price,"seller":seller,"img":imgs[0] if imgs else None,"term":term})
    return out

TERMS = ["рис швидкого приготування", "готовий рис", "рис готовий до вживання",
         "instant rice", "рис у чашці", "рис у стакані", "рис з м'ясом готовий",
         "henan рис", "рис henan", "рис зі смаком гребінців",
         "саморозігрівальний рис", "self heating rice", "рис самонагрівальний",
         "сублімований рис", "рис сублімований страва", "рис быстрого приготовления",
         "рис в пакеті готовий 90 секунд", "microwave rice", "рис мікрохвильовка готовий",
         "корейський готовий рис", "китайський готовий рис", "hetbahn", "cj hetbahn рис",
         "nongshim рис", "samyang рис", "paldo рис", "рис доширак", "uncle bens рис"]
res=[]
for t in TERMS:
    try:
        c = cards(t); res += c
        print(f"{t:38} {len(c)}")
    except Exception as e:
        print(t, "ERR", str(e)[:60])
    time.sleep(0.4)
json.dump(res, open("sweep.json","w"), ensure_ascii=False, indent=1)
print("\nвсього карток", len(res))
