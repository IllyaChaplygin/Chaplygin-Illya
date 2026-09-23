"""Завантаження пакшотів під кожен SKU + нормалізація."""
import json, os, re, requests
import numpy as np
from PIL import Image, ImageFilter
H={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36"}
os.makedirs("sku", exist_ok=True)
cards=json.load(open("sku_cards.json"))

# (файл, бренд, фрагмент назви для пошуку картки)
WANT=[
 ("ot_bibimbap","ottogi","бибимбап"),        ("ot_pork_spicy","ottogi","Гостра свинина"),
 ("ot_pork310","ottogi","свининою 310"),      ("ot_beef320","ottogi","Гостра яловичина"),
 ("ot_bulgogi","ottogi","бульгогі"),          ("ot_chicken_rib","ottogi","курячими ребрами"),
 ("ot_tuna","ottogi","тунцем і майонезом"),   ("ot_octopus","ottogi","осьминога 280"),
 ("hd_stew_chicken187","haidilao","тушкованою куркою, 187"),
 ("hd_teriyaki181","haidilao","куркою теріяки, 181"),
 ("hd_pork_fish187","haidilao","подрібненою свининою"),
 ("hd_spicy_chicken175","haidilao","гострим курячим соусом"),
 ("hd_canton187","haidilao","кантонському стилі"),
 ("hd_beef_curry272","haidilao","яловичиною та карі"),
 ("hd_beef_stew272","haidilao","тушкованою яловичиною, 272"),
 ("hd_tomato_beef272","haidilao","томатами та яловичою"),
 ("hd_stewed_chicken165","haidilao","Stewed Chicken 165"),
 ("tl_nasi125","travellunch","Nasi Goreng 125"),
 ("tl_nasi250","travellunch","Nasi Goreng 250"),
 ("tl_strog125","travellunch","бефстроганів з рисом 125"),
 ("tl_strog250","travellunch","Бефстроганов з рисом 250"),
 ("tl_beef_pepper125","travellunch","яловичиною та перцем 125"),
 ("tl_beef_pepper250","travellunch","Яловичиною та Соусом Перцю 250"),
 ("mx_beef_bamboo","moxiaoxian","пагонами бамбука"),
 ("mx_bacon_peas","moxiaoxian","беконом та зеленим"),
 ("sm_chicken_fruit","sublimate","куркою, фруктами"),
 ("sm_green_curry","sublimate","зелене карі"),
 ("sm_tomato_soup","sublimate","Томатний суп"),
 ("cs_brown_wild","clearspring","Тамарі органічна"),
]

def trim(im, thr=246, pad=0.02):
    a=np.asarray(im.convert("RGB")).min(axis=2)
    ys,xs=np.where(a<thr)
    if not len(xs): return im
    x0,x1,y0,y1=xs.min(),xs.max(),ys.min(),ys.max()
    m=int(pad*max(x1-x0,y1-y0))
    return im.crop((max(0,x0-m),max(0,y0-m),min(im.width,x1+m),min(im.height,y1+m)))

ok=0
for fn, brand, frag in WANT:
    hit=next((c for c in cards.get(brand,[]) if frag.lower() in c["name"].lower() and c["img"]), None)
    if not hit:
        print(f"  MISS {fn:22} ({frag})"); continue
    url=hit["img"].replace("_w400_","_w640_").replace("_h400_","_h640_")
    try:
        r=requests.get(url,headers=H,timeout=45)
        if r.status_code!=200 or len(r.content)<4000:
            r=requests.get(hit["img"],headers=H,timeout=45)
        im=trim(Image.open(__import__("io").BytesIO(r.content)).convert("RGB"))
        im=im.resize((max(1,round(im.width*760/im.height)),760), Image.LANCZOS)
        im=im.filter(ImageFilter.UnsharpMask(radius=1.2,percent=55,threshold=3))
        im.save(f"sku/{fn}.jpg", quality=94); ok+=1
        print(f"  ok   {fn:22} {im.size}  {hit['price']} грн")
    except Exception as e:
        print(f"  ERR  {fn:22} {str(e)[:50]}")
print("завантажено:", ok, "із", len(WANT))
