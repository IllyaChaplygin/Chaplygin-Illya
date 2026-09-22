import requests, urllib.parse, re, html, json, time
UA=("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36")
H={"User-Agent":UA,"Accept-Language":"uk"}
TERMS=["готовий рис","рис готовий до вживання","veetee","hetbahn","хетбан",
       "ben's original","uncle ben's рис","рис у мікрохвильовку","рис пауч готовий",
       "ready rice","instant rice","рис варений","готовий рис корея","bibigo рис",
       "рис в стаканчику готовий","смажений рис готовий"]
out={}
for t in TERMS:
    u="https://prom.ua/ua/search?search_term="+urllib.parse.quote(t)
    try:
        r=requests.get(u,headers=H,timeout=35)
        txt=r.text
        # product cards are embedded as JSON-LD / inline state
        names=re.findall(r'"name"\s*:\s*"([^"]{6,120})"', txt)
        prices=re.findall(r'"price"\s*:\s*"?([\d.]+)"?', txt)
        cnt=re.search(r'(\d[\d\s]*)\s*(?:товар|пропозиц)', html.unescape(re.sub(r"<[^>]+>"," ",txt)))
        out[t]={"http":r.status_code,"names":[html.unescape(n) for n in names[:40]],
                "prices":prices[:40],"count_hint":cnt.group(0) if cnt else None}
        print(f"{t:30} {r.status_code} names={len(names)} {out[t]['count_hint']}")
    except Exception as e:
        print(t,"ERR",str(e)[:60]); out[t]={"error":str(e)[:80]}
    time.sleep(0.5)
json.dump(out,open("prom_results.json","w"),ensure_ascii=False,indent=1)
