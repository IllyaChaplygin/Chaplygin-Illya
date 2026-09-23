"""Пакшоти для нових брендів: Henan, Yatekomo, Qiaoshanmei, сублімати, Маркел."""
import json, os, requests
import numpy as np
from PIL import Image, ImageFilter
H={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36"}
os.makedirs("sku", exist_ok=True)
pool=json.load(open("pool.json"))
WANT=[("hn_braised174","henan",0),("hn_scallop174","henan",1),("hn_sausage174","henan",2),
      ("hn_pepperbeef174","henan",3),("hn_pepperchicken174","henan",14),("hn_spicychicken144","henan",11),
      ("hn_sweetsourbeef144","henan",13),("hn_eggplant144","henan",18),
      ("hn_braised174_dcm","henan",8),("hn_pepperchicken174_dcm","henan",9),
      ("yt_beef84","yatekomo",0),("yt_teriyaki84","yatekomo",1),
      ("qs_sh_pepperbeef","qiaoshan",0),("qs_sh_braisedpork","qiaoshan",1),
      ("qs_spicybeef","qiaoshan",2),("qs_braisedpork","qiaoshan",3),
      ("hd_porkbelly360","haidilao",12),("hd_pork_garlic","haidilao",17),("hd_pork_fish","haidilao",6),
      ("tl_curry125","travellunch",1),("sm_plov140","sublimate",4),
      ("jc_meatveg","jamescook",3),("jc_curry","jamescook",4),("jc_mashkichiri","jamescook",0),("jc_veg","jamescook",7),
      ("yp_chicken85","yizha",0),("yp_veg85","yizha",1),
      ("kh_pork","kharchi",0),("kh_kichri","kharchi",1),("kh_plovxl","kharchi",2),("kh_ricemeat","kharchi",10),
      ("am_korma110","advmenu",5),("am_tikka110","advmenu",9),("am_meatballs","advmenu",16),
      ("am_korma400","advmenu",1),("am_wild","advmenu",8),
      ("af_curry146","advfood",0),("mh_curry","mh",0),("te_fish","trek",0),("te_strog","trek",1),
      ("fest_plov","fest",1),("mk_rice350","markel",1)]
def trim(im, thr=246, pad=0.02):
    a=np.asarray(im.convert("RGB")).min(axis=2); ys,xs=np.where(a<thr)
    if not len(xs): return im
    x0,x1,y0,y1=xs.min(),xs.max(),ys.min(),ys.max(); m=int(pad*max(x1-x0,y1-y0))
    return im.crop((max(0,x0-m),max(0,y0-m),min(im.width,x1+m),min(im.height,y1+m)))
ok=0
for fn,b,i in WANT:
    c=pool[b][i]; url=c["img"]
    for u in (url.replace("_w400_","_w640_"), url):
        try:
            r=requests.get(u,headers=H,timeout=40)
            if r.status_code==200 and len(r.content)>3000:
                open("tmp.img","wb").write(r.content)
                im=trim(Image.open("tmp.img").convert("RGB"))
                h=760; im=im.resize((max(1,round(im.width*h/im.height)),h),Image.LANCZOS)
                im=im.filter(ImageFilter.UnsharpMask(radius=1.1,percent=50,threshold=3))
                im.save(f"sku/{fn}.jpg",quality=92); ok+=1; print("OK ",fn,im.size); break
        except Exception as e: print("ERR",fn,str(e)[:60])
    else: print("FAIL",fn)
print(ok,"/",len(WANT))
