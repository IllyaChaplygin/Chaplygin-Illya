import requests, urllib.parse, json, re, os, html
UA=("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36")
H={"User-Agent":UA,"Accept-Language":"uk"}
os.makedirs("img", exist_ok=True)
manifest=[]

def grab(url, name, src, meta=None):
    try:
        r=requests.get(url, headers=H, timeout=60)
        if r.status_code!=200 or len(r.content)<3000:
            print(f"  skip {name}: {r.status_code} {len(r.content)}B"); return
        ext = ".png" if r.content[:4]==b"\x89PNG" else ".jpg"
        p=f"img/{name}{ext}"
        open(p,"wb").write(r.content)
        manifest.append({"file":p,"source":src,"url":url, **(meta or {})})
        print(f"  ok   {p}  {len(r.content)//1024}KB")
    except Exception as e:
        print(f"  ERR  {name}: {str(e)[:60]}")

# ── A. the two ambient RTE SKUs actually on sale in Ukraine ──────────────
grab("https://www.smak-korea.com.ua/wp-content/uploads/vareniy-ris-210gr-2.jpg",
     "ua_hetbahn", "smak-korea.com.ua",
     {"title":"CJ Hetbahn «Варений рис» 210 г","price":135.0})

C="challenge_passed=ec5771e8cbb89de94cf2cacd37791f002083fdae11fb2acf1aeece2448ba8afc"
r=requests.get("https://asia-goods.com.ua/hotovyi-rys-z-korei/",headers={**H,"Cookie":C},timeout=40)
imgs=re.findall(r'(https://asia-goods\.com\.ua/[^"\']*?\.(?:jpg|jpeg|png|webp))', r.text)
imgs=[i for i in imgs if "logo" not in i.lower() and "icon" not in i.lower()]
print("asia-goods candidate images:", len(imgs))
for i,u in enumerate(dict.fromkeys(imgs)):
    if i>=3: break
    grab(u, f"ua_asiagoods_{i}", "asia-goods.com.ua",
         {"title":"«Готовий рис з Кореї», арт. 10296","price":190.0})

# ── B. substitutes, high-res from the zakaz CDN ─────────────────────────
ZH={**H,"Accept":"application/json","Origin":"https://novus.zakaz.ua","Referer":"https://novus.zakaz.ua/"}
WANT=[("рис","Рис Novus довгозернистий пропарений 5x80г","sub_rice_bags"),
      ("локшина","Shin Cup","sub_noodle_cup"),
      ("токпоккі","Yopokki","sub_tteok"),
      ("каша","каша","sub_porridge")]
for store in ["482010105","48246401","48215610"]:
    for q, needle, name in WANT:
        if any(m["file"].startswith(f"img/{name}") for m in manifest): continue
        u=f"https://stores-api.zakaz.ua/stores/{store}/products/search/?q={urllib.parse.quote(q)}"
        try: res=requests.get(u,headers=ZH,timeout=30).json().get("results",[])
        except Exception: continue
        for it in res:
            if needle.lower() in (it.get("title") or "").lower():
                im=(it.get("img") or {}).get("s1350x1350")
                if im:
                    grab(im, name, it.get("web_url",""),
                         {"title":it["title"],"price":(it.get("price") or 0)/100})
                    break

json.dump(manifest, open("img/manifest.json","w"), ensure_ascii=False, indent=1)
print("\nmanifest entries:", len(manifest))
