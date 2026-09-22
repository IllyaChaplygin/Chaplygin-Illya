import requests, urllib.parse, re, html, json, time
H={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36","Accept-Language":"uk"}
TERMS=[
 "готовий рис","рис готовий до вживання","рис у пауці","рис пауч готовий",
 "ben's original рис","uncle ben's рис","veetee рис","tilda рис","sunrice",
 "clearspring рис","bibigo рис","ottogi рис","рис express 2 хвилини",
 "рис мікрохвильовка готовий","ready to eat rice","microwave rice pouch",
 "рис басматі готовий пауч","рис жасмин готовий","рис з овочами готовий пауч",
 "рис у соусі пауч","рис швидкого приготування пауч","готовий гарнір рис",
 "riso pronto","рис 250 г готовий","рис 220 г готовий","рис бібімбап",
 "сенсой рис","sen soy рис","рис для сушi готовий","рис вarilla",
]
KEY=re.compile(r"рис|rice|bap\b|bibimbap", re.I)
BAD=re.compile(r"локшин|вермішель|лапш|корм|папір|оцет|борошн|крекер|чіпс|соус до|молок|"
               r"каша|сироп|олі|насін|сир\b|хліб|печив|цукерк|вино|кав|чай|тарілк|"
               r"рисоварк|контейнер|мило|шампун|крем", re.I)
found={}
for t in TERMS:
    u="https://prom.ua/ua/search?search_term="+urllib.parse.quote(t)
    try:
        tx=requests.get(u,headers=H,timeout=35).text
        names=[html.unescape(n) for n in re.findall(r'"name"\s*:\s*"([^"]{6,140})"', tx)]
        for n in dict.fromkeys(names):
            if KEY.search(n) and not BAD.search(n) and not n.isupper():
                found.setdefault(n, set()).add(t)
        print(f"{t:34} {len(names)} назв")
    except Exception as e:
        print(t,"ERR",str(e)[:50])
    time.sleep(0.35)
json.dump({k:sorted(v) for k,v in found.items()}, open("prom_hunt.json","w"),
          ensure_ascii=False, indent=1)
print("\n=== КАНДИДАТИ ===")
for n in sorted(found): print("  ", n[:110])
