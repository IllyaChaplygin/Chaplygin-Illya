import json, time, urllib.parse, sys, requests
UA=("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")
def H(chain): return {"User-Agent":UA,"Accept":"application/json","Accept-Language":"uk",
                      "Origin":f"https://{chain}.zakaz.ua","Referer":f"https://{chain}.zakaz.ua/"}
stores=json.load(open("zk.json"))
chains={}
for s in stores: chains.setdefault(s["retail_chain"], s["id"])
SKIP={"winetime","alcohub","masterzoo","biotus"}
TERMS=["рис","готовий рис","ready rice","рис швидкого приготування","рис для мікрохвильовки",
       "смажений рис","hetbahn","bibigo","uncle ben","ben's original","veetee","tilda",
       "рис у соусі","fried rice","рис готовий","онігірі","ottogi","cj","рис басматі",
       "каша швидкого приготування","локшина швидкого приготування"]
def norm(it):
    disc=it.get("discount") or {}
    return {"title":it.get("title"),"ean":it.get("ean"),"sku":it.get("sku"),
            "price":(it.get("price") or 0)/100 or None,
            "old_price":(disc.get("old_price") or 0)/100 or None,
            "on_sale":disc.get("status"),
            "weight":it.get("weight"),"unit":it.get("unit"),
            "volume":it.get("volume"),
            "producer":(it.get("producer") or {}).get("trademark") if isinstance(it.get("producer"),dict) else it.get("producer"),
            "country":it.get("country"),
            "price_per":it.get("price_per_unit") or it.get("unit_price"),
            "url":it.get("web_url")}
out={}
for chain,sid in chains.items():
    if chain in SKIP: continue
    for t in TERMS:
        u=f"https://stores-api.zakaz.ua/stores/{sid}/products/search/?q={urllib.parse.quote(t)}"
        try:
            r=requests.get(u,headers=H(chain),timeout=30)
            if r.status_code!=200: out.setdefault(chain,{})[t]={"http":r.status_code}; continue
            d=r.json()
            out.setdefault(chain,{})[t]={"count":d.get("count"),
                                         "items":[norm(x) for x in (d.get("results") or [])[:60]]}
        except Exception as e:
            out.setdefault(chain,{})[t]={"error":str(e)[:70]}
        time.sleep(0.2)
    print(chain,"done",file=sys.stderr)
json.dump(out,open("zakaz_results.json","w"),ensure_ascii=False,indent=1)
print("saved",file=sys.stderr)
