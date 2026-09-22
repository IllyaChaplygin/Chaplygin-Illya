import requests, re, sys
UA=("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")
H={"User-Agent":UA}
SHOPS=["https://sushipovar.ua","https://yaponskiy-kvartal.com","https://candymeow.com.ua",
       "https://asiafoods.com.ua","https://rozetka.com.ua"]
KW=re.compile(r"(gotov|hotov|varen|hetbahn|bahn|ready-rice|ready_rice|instant-rice|gotoviy|готов)", re.I)
RICE=re.compile(r"(ris|rys|rice|рис)", re.I)
for base in SHOPS:
    urls=set()
    for sm in ["/sitemap.xml","/sitemap_index.xml","/sitemap-index.xml"]:
        try:
            r=requests.get(base+sm,headers=H,timeout=35)
            if r.status_code!=200: continue
            locs=re.findall(r"<loc>\s*(.*?)\s*</loc>", r.text)
            subs=[u for u in locs if u.endswith(".xml")]
            if subs:
                for s in subs[:14]:
                    try:
                        rr=requests.get(s,headers=H,timeout=35)
                        urls |= set(re.findall(r"<loc>\s*(.*?)\s*</loc>", rr.text))
                    except Exception: pass
            else:
                urls |= set(locs)
            break
        except Exception as e:
            pass
    cand=[u for u in urls if RICE.search(u) and KW.search(u)]
    print(f"{base:34} sitemap urls={len(urls):6}  rice+ready={len(cand)}")
    for u in sorted(cand)[:25]: print("     ", u)
    sys.stdout.flush()
