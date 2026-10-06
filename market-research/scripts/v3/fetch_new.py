import requests, re, json, html, os, sys
H={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/140.0.0.0 Safari/537.36","Accept-Language":"uk"}
NEW={
 'soc_jasmine':'/ua/p2309195374-seeds-change-organicheskij.html',
 'soc_brownwild':'/ua/p2309055293-seeds-change-organicheskij.html',
 'soc_cilantro':'/ua/p2309055294-seeds-change-organicheskij.html',
 'soc_brownbasmati':'/ua/p2312760393-seeds-change-organicheskij.html',
 'soc_spanish':'/ua/p2415838708-seeds-change-organicheskij.html',
 'soc_quinoagarlic':'/ua/p2070289595-ris-seeds-change.html',
 'forestia':'/ua/p2898788202-forestia-gotova-strava.html',
 'hd_spicy360':'/ua/p3214732523-kitajskij-samorazogrevayuschijsya-ris.html',
 'af_cashew':'/ua/p3012612040-sublimirovannaya-eda-adventure.html',
 'kh_plov_veal':'/ua/p2175144464-100-sublimat-plov.html',
 'kh_plov_rice':'/ua/p2859362555-plov-bagato-risu.html',
 'qs_chicken':'/ua/p3169921940-ris-bystrogo-prigotovleniya.html',
 'bens_basmati220_es':'/ua/p3222916784-ris-basmati-proparennyj.html',
}
res={}
for k,u in NEW.items():
    try:
        t=requests.get('https://prom.ua'+u,headers=H,timeout=40).text
    except Exception as e:
        print(k,'ERR',e); continue
    og=re.search(r'<meta property="og:image" content="([^"]+)"',t)
    pr=re.search(r'"price"\s*:\s*"?([\d.]+)',t) or re.search(r'data-qaid="product_price"[^>]*>([^<]+)',t)
    ti=re.search(r'<title>([^<]+)',t)
    seller=re.search(r'data-qaid="company_name"[^>]*>([^<]+)',t)
    avail=re.search(r'data-qaid="product_presence"[^>]*>([^<]+)',t)
    im=og.group(1) if og else None
    res[k]=dict(url=u,img=im,price=pr.group(1) if pr else None,title=html.unescape(ti.group(1))[:90] if ti else None,
                seller=html.unescape(seller.group(1)).strip() if seller else None, avail=avail.group(1).strip() if avail else None)
    if im:
        b=requests.get(im,headers=H,timeout=40).content
        open(f'{k}.jpg','wb').write(b)
    print(k,res[k])
json.dump(res,open('new.json','w'),ensure_ascii=False,indent=1)
