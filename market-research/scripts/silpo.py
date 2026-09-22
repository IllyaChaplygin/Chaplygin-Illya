import requests, urllib.parse, json, time
UA=("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")
H={"User-Agent":UA,"Accept":"application/json","Accept-Language":"uk","Referer":"https://silpo.ua/"}
B="1ed43e73-051b-6842-a111-a5ad042eb496"
TERMS=["рис","готовий рис","ready rice","рис швидкого приготування","смажений рис",
       "hetbahn","bibigo","uncle ben","ben's original","veetee","tilda","рис у соусі",
       "fried rice","ottogi","рис для мікрохвильовки","онігірі","рис басматі","кімчі"]
out={}
for t in TERMS:
    u=(f"https://sf-ecom-api.silpo.ua/v1/uk/branches/{B}/products?limit=60&offset=0"
       f"&search={urllib.parse.quote(t)}")
    try:
        r=requests.get(u,headers=H,timeout=30); d=r.json()
        items=[{"title":x.get("title"),"price":x.get("price"),"oldPrice":x.get("oldPrice"),
                "brand":x.get("brandTitle"),"ratio":x.get("ratio"),
                "displayPrice":x.get("displayPrice"),"displayRatio":x.get("displayRatio"),
                "section":x.get("sectionSlug"),"slug":x.get("slug")}
               for x in d.get("items",[])]
        out[t]={"total":d.get("total"),"items":items}
        print(f"{t:32} total={d.get('total')}")
    except Exception as e:
        out[t]={"error":str(e)[:80]}; print(t,"ERR",str(e)[:60])
    time.sleep(0.3)
json.dump(out,open("silpo_results.json","w"),ensure_ascii=False,indent=1)
