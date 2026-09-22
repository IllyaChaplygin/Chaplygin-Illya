import json, urllib.parse, time, requests
UA=("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")
H={"User-Agent":UA,"Accept":"application/json","Accept-Language":"uk"}
S="https://search.rozetka.com.ua/ua/search/api/v6/?front-type=xl&country=UA&lang=ua&text="
D="https://common-api.rozetka.com.ua/v2/goods/get-details?front-type=xl&country=UA&lang=ua&with_groups=1&with_docket=1&goods_group_href=1&product_ids="

def ids(term, limit=40):
    d=requests.get(S+urllib.parse.quote(term),headers=H,timeout=30).json()
    return [g["id"] for g in d["data"]["goods"][:limit]], d["data"]["quantities"]["goods_quantity_total_found"]

def details(idlist):
    out=[]
    for i in range(0,len(idlist),20):
        chunk=",".join(str(x) for x in idlist[i:i+20])
        r=requests.get(D+chunk,headers=H,timeout=40)
        try: dd=r.json().get("data",[])
        except Exception: dd=[]
        for x in dd:
            out.append({"id":x.get("id"),"title":x.get("title"),
                        "price":x.get("price"),"old_price":x.get("old_price"),
                        "status":x.get("sell_status"),"seller":x.get("seller",{}).get("title") if isinstance(x.get("seller"),dict) else x.get("seller_id"),
                        "brand":x.get("brand"),"href":x.get("href"),
                        "category":(x.get("category") or {}).get("title") if isinstance(x.get("category"),dict) else None})
        time.sleep(0.3)
    return out

TERMS=["готовий рис","ready rice","рис швидкого приготування","hetbahn","хетбан","bibigo",
       "uncle ben's","ben's original","veetee","рис в соусі","смажений рис готовий",
       "ottogi рис","рис пауч готовий","рис готовий до вживання","instant rice","рис для мікрохвильової печі"]
res={}
for t in TERMS:
    try:
        il,total=ids(t)
        res[t]={"total":total,"items":details(il)}
        print(f"{t:34} total={total:5} got={len(res[t]['items'])}")
    except Exception as e:
        print(t,"ERR",str(e)[:70]); res[t]={"error":str(e)[:90]}
    time.sleep(0.3)
json.dump(res,open("rozetka_results.json","w"),ensure_ascii=False,indent=1)
