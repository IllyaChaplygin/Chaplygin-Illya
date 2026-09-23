"""Збір карток SKU з Prom.ua: назва, ціна, продавець і пакшот."""
import requests, urllib.parse, re, html, json, os, time
H={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36","Accept-Language":"uk"}
os.makedirs("sku", exist_ok=True)

def cards(term):
    tx = requests.get("https://prom.ua/ua/search?search_term=" + urllib.parse.quote(term),
                      headers=H, timeout=45).text
    out = []
    for blk in re.split(r'data-qaid="product_block"', tx)[1:]:
        blk = blk[:9000]
        nm = re.search(r'data-qaid="product_name"[^>]*>(.*?)</', blk)
        if not nm:
            continue
        name = html.unescape(re.sub("<[^>]+>", "", nm.group(1))).strip()
        pr = re.search(r'data-qaid="product_price"[^>]*>(.*?)</span>', blk, re.S)
        price = None
        if pr:
            d = re.sub(r"[^\d]", "", html.unescape(re.sub("<[^>]+>", "", pr.group(1))))
            price = int(d) if d else None
        sel = re.search(r'data-qaid="company_name"[^>]*>(.*?)</', blk)
        seller = html.unescape(re.sub("<[^>]+>", "", sel.group(1))).strip() if sel else None
        imgs = re.findall(r'(https://images\.prom\.ua/[^"\'\\ )]+_w400_[^"\'\\ )]+)', blk)
        out.append({"name": name, "price": price, "seller": seller,
                    "img": imgs[0] if imgs else None})
    return out

TERMS = {
    "ottogi":      ["готовий рис ottogi", "рис ottogi відварений", "стакан відвареного рису ottogi",
                    "ottogi рис бібімбап", "ottogi рис свинина", "ottogi рис яловичина"],
    "haidilao":    ["саморозігрівальний рис haidilao", "haidilao рис", "саморозігрівальний рис"],
    "travellunch": ["travellunch рис"],
    "moxiaoxian":  ["mo xiao xian рис", "instant rice китай 275"],
    "portion":     ["portion рис з куркою", "пирятинський делікатес рис"],
    "sublimate":   ["sublimate рис"],
    "bibigo":      ["bibigo рис 210"],
    "clearspring": ["clearspring рис tamari", "clearspring brown wild rice"],
    "bens":        ["ben's original рис", "бенс оріджинал рис"],
}
res = {}
for brand, terms in TERMS.items():
    seen, items = set(), []
    for t in terms:
        try:
            for c in cards(t):
                k = (c["name"], c["seller"])
                if k in seen:
                    continue
                seen.add(k); items.append(c)
        except Exception as e:
            print(brand, t, "ERR", str(e)[:50])
        time.sleep(0.35)
    res[brand] = items
    print(f"{brand:12} {len(items)} карток, з фото {sum(1 for i in items if i['img'])}")
json.dump(res, open("sku_cards.json", "w"), ensure_ascii=False, indent=1)
