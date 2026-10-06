# -*- coding: utf-8 -*-
"""Спільний модуль колоди «Корейський рамен в Україні» за моделлю рисової колоди."""
import sys,json,os,re,math,statistics as st,collections as co
SC="/tmp/claude-0/-home-user-Chaplygin-Illya/8bc0034a-4b53-5022-9bf5-e5697c987387/scratchpad"
sys.path.insert(0,SC)
from deckkit import *
from deckkit import _rect,_S
from pptx.enum.text import PP_ALIGN as A
import calc120 as M
L=A.LEFT; R=A.RIGHT; C=A.CENTER

# ---------------- дані ----------------
F=json.load(open(f"{SC}/facts.json",encoding='utf-8'))
SEG=json.load(open(f"{SC}/cards_seg.json",encoding='utf-8'))
CK=json.load(open(f"{SC}/cards_kor.json",encoding='utf-8'))
CM=json.load(open(f"{SC}/cards_mass.json",encoding='utf-8'))
SS=json.load(open(f"{SC}/segstats.json",encoding='utf-8'))
_SJ=json.load(open(f"{SC}/segmented.json",encoding='utf-8')); SD_ROWS=_SJ['rows']
BT={r['tm']:r for r in json.load(open(f"{SC}/brandtiers.json",encoding='utf-8'))}
if "без тм" in CK: CK["Choi's"]=CK.pop("без тм")

# ---- v5: без рисових галушок (сегмент 4) — прибрано з усіх розрахунків ----
SD_ROWS=[r for r in SD_ROWS if r['seg']!=4]
SEG.pop('4',None); SS.pop('4',None)
_ZR=json.load(open(f"{SC}/zakaz_raw.json",encoding='utf-8'))
_EANS={r['ean']:r for r in SD_ROWS}
ZREC=[z for z in _ZR if z['ean'] in _EANS]                       # записи «магазин × SKU» для 89 позицій
F['n_sku']=len(SD_ROWS); F['n_brands']=len({r['tm'] for r in SD_ROWS})
F['n_chains']=len({c for r in SD_ROWS for c in r['chains']}|{'silpo'})
def _med(a): return st.median(a) if a else 0
_bands=[(40,90,"40–90 г"),(90,110,"90–110 г"),(110,125,"110–125 г"),(125,145,"125–145 г"),(145,900,"145 г +")]
F['bands']=[]
for a_,b_,lab in _bands:
    v=[r for r in SD_ROWS if a_<=r['w']<b_]
    F['bands'].append({"lab":lab,"n":len(v),"share":round(len(v)/len(SD_ROWS)*100),
                       "pmed":round(_med([r['pmed'] for r in v])),"g":round(_med([r['per100'] for r in v]))})
F['matrix']=[[sum(1 for r in SD_ROWS if r['seg'] in fm and a_<=r['w']<b_) for a_,b_,_ in _bands] for fm in ((1,2),(3,))]
F['chain_rows']=[]
for ch_,v in sorted(co.defaultdict(list,{c:[z for z in ZREC if z['chain']==c] for c in {z['chain'] for z in ZREC}}).items(),key=lambda x:-len({z['ean'] for z in x[1]})):
    pr=[z['price'] for z in v]
    F['chain_rows'].append({"chain":ch_,"n":len({z['ean'] for z in v}),"med":round(_med(pr)),"lo":round(min(pr)),"hi":round(max(pr))})
TOP2=[b for b in ("Samyang","Nongshim")]
N_TOP2=sum(1 for r in SD_ROWS if r['tm'] in TOP2)
def g(b): return CK.get(b,[])
def tl(tm,label=None):
    r=BT[tm]; p=f"{SC}/img/{r['best']['ean']}.jpg"
    return {"img":p if os.path.exists(p) else None,"name":label or tm,
            "price":f"{r['pmed']} грн","sub":f"{r['g']} грн/100 г · {r['nch']} мереж · {r['sku']} SKU"}
def img_for(ean):
    p=f"{SC}/img/{ean}.jpg"; return p if os.path.exists(p) else None

# ---------------- форматування ----------------
def d1(x): return f"{x:.1f}".replace('.',',')
def d2(x): return f"{x:.2f}".replace('.',',')
def d0(x): return f"{x:.0f}"
def th(x): return f"{x:,.0f}".replace(',',' ')
def th1(x): return f"{x:,.1f}".replace(',',' ').replace('.',',')
def eur(x,n=2): return "€"+f"{x:.{n}f}".replace('.',',')
def pct(x,n=1): return (f"{x:.{n}f}".replace('.',',').replace('-','−'))+" %"
def sgn(x,n=1): return (f"{x:+.{n}f}".replace('.',',').replace('-','−'))+" %"
NM={'auchan':'Ашан','novus':'NOVUS','metro':'METRO','cosmos':'«Космос»','vostorg':'«Восторг»',
    'tavriav':'«Таврія В»','torba':'«Торба»','zaraz':'«Зараз»','grono':'Grono','onde':'Onde',
    'chudomarket':'«ЧудоМаркет»','kharkiv':'«Клас»','silpo':'«Сільпо»'}
NATIONAL={'auchan','novus','metro','zaraz','silpo'}

# сегменти
SNAME={'1':'РАМЕН У ПАКЕТІ · З БУЛЬЙОНОМ','2':'РАМЕН У ПАКЕТІ · ІЗ СОУСОМ','3':'РАМЕН У СТАКАНІ','4':'РИСОВІ ГАЛУШКИ ТА ЛОКШИНА'}
SSHORT={'1':'Пакет · з бульйоном','2':'Пакет · із соусом','3':'Стакан','4':'Рисові'}
SCOL={'1':NAVY,'2':DARKAMBER,'3':PURPLE,'4':TEAL}
STITLE={'1':'Пакет з бульйоном','2':'Пакет із соусом','3':'Стакан','4':'Рисові галушки'}
SPREP={'1':'Варіння 4–5 хв · подається з бульйоном','2':'Варіння 4–5 хв · воду злити, змішати із соусом',
       '3':'Залити окропом 3–4 хв · без плити','4':'Рисова основа · окремий привід споживання'}
PB=[(0,80,"до 80"),(80,95,"80–95"),(95,110,"95–110"),(110,130,"110–130"),(130,150,"130–150"),(150,180,"150–180"),(180,9e9,"180 +")]
def band_of(p):
    for i,(lo,hi,_) in enumerate(PB):
        if lo<=p<hi: return i
    return len(PB)-1

PACK=[r for r in SD_ROWS if r['seg'] in (1,2) and r.get('pmed')]
PACK_P=sorted(r['pmed'] for r in PACK)
PACK_MED=st.median(PACK_P)
def quart(arr):
    """Квартили методом statistics.quantiles (exclusive) — як у статистиці сегментів."""
    return st.quantiles(sorted(arr),n=4)
PACK_Q=quart(PACK_P)

# ---------------- базові примітиви колоди ----------------
def estat(s,l,t,label,big,note,acc=DARKAMBER,w=3.88,h=1.46,bigsz=26):
    _rect(s,l,t,w,h,LIGHT); _rect(s,l,t,0.045,h,acc)
    txt(s,l+0.28,t+0.17,w-0.52,0.14,label,7.6,True,GREY)
    txt(s,l+0.28,t+0.40,w-0.52,0.46,big,bigsz,True,acc)
    ny=(t+h-0.52) if h>=1.40 else (t+0.40+bigsz*1.25/72+0.06)
    txt(s,l+0.28,ny,w-0.52,0.40,note,8.4,False,GREY,lh=1.3)

def ecall(s,l,t,w,h,head,body,acc=AMBER,hs=10,bs=8.8):
    _rect(s,l,t,w,h,LIGHT); _rect(s,l,t,0.045,h,acc)
    txt(s,l+0.30,t+0.18,w-0.60,0.18,head,hs,True,NAVY)
    txt(s,l+0.30,t+0.48,w-0.60,h-0.66,body,bs,False,GREY,lh=1.42)

def estrip(s,t,head,body,acc=NAVY,h=0.60):
    _rect(s,0.62,t,12.10,h,LIGHT); _rect(s,0.62,t,0.045,h,acc)
    txt(s,0.92,t+0.10,11.50,0.16,head,9.2,True,NAVY)
    txt(s,0.92,t+0.32,11.50,0.20,body,8.6,False,GREY,lh=1.3)

def verdict(s,t,head,body,acc=RED,h=0.92):
    _rect(s,0.62,t,12.10,h,NAVY); _rect(s,0.62,t,0.06,h,acc)
    txt(s,0.96,t+0.14,11.50,0.22,head,11,True,GOLD)
    txt(s,0.96,t+0.42,11.50,0.40,body,13.5,True,WHITE,lh=1.22)

def cardrow(s,t,items,acc=BLUE,n=7):
    for i,c in enumerate(items[:n]):
        card(s,CARD_X[i],t,c.get("img"),c["name"],c["w"],c["price"],c["sellers"],acc)

def minibar(s,l,t,w,rows,xmax,acc=BLUE,lab_w=1.05,val_w=0.62,rowh=0.205):
    bw=w-lab_w-val_w
    for lab,v,extra in rows:
        txt(s,l,t+0.02,lab_w-0.08,0.14,lab,7.8,False,NAVY)
        _rect(s,l+lab_w,t+0.045,max(bw*v/xmax,0.02),0.105,acc)
        txt(s,l+w-val_w,t+0.02,val_w,0.14,extra,7.8,True,GREY,PP_ALIGN.RIGHT)
        t+=rowh
    return t

def tint(c,k=0.55):
    r,g_,b=c[0],c[1],c[2]
    return RGBColor(int(r+(255-r)*k),int(g_+(255-g_)*k),int(b+(255-b)*k))
AMBLIGHT=RGBColor(0xFD,0xF3,0xDF)

def rangebars(s,l,t,w,rows,xmax,ticks,rowh=0.285,lab_w=2.60,w_col=1.05,n_col=0.45,val_w=1.15,band=None,bandlab=None,
              head=None,headcol=NAVY):
    """Діапазонні смуги за брендами: тонка лінія min–max, блок — медіана. rows=[dict(label,wt,n,lo,med,hi,col,bold)]."""
    px=l+lab_w+w_col+n_col+0.12; pw=w-(lab_w+w_col+n_col+0.12)-val_w
    def X(v): return px+pw*min(max(v,0),xmax)/xmax
    H=len(rows)*rowh
    if head:
        _rect(s,l,t,w,0.30,headcol)
        txt(s,l+0.14,t+0.07,lab_w,0.18,head,8.2,True,WHITE)
        txt(s,l+lab_w,t+0.08,w_col,0.14,"ВАГА",6.6,True,WHITE)
        txt(s,l+lab_w+w_col,t+0.08,n_col,0.14,"SKU",6.6,True,WHITE,PP_ALIGN.RIGHT)
        t+=0.36
    if band:
        _rect(s,X(band[0]),t-0.02,X(band[1])-X(band[0]),H+0.04,AMBLIGHT)
    for tv in ticks:
        _rect(s,X(tv),t-0.02,0.006,H+0.04,HAIR)
    y=t
    for rw in rows:
        b=rw.get('bold',False); col=rw['col']
        if b: _rect(s,l,y,w-0.02,rowh-0.02,AMBLIGHT)
        txt(s,l+0.10,y+rowh/2-0.095,lab_w-0.12,0.19,rw['label'],8.4,True,DARKAMBER if b else NAVY)
        txt(s,l+lab_w,y+rowh/2-0.085,w_col,0.17,rw.get('wt',''),7.6,False,GREY)
        txt(s,l+lab_w+w_col,y+rowh/2-0.085,n_col,0.17,str(rw.get('n','')),8,True,GREY,PP_ALIGN.RIGHT)
        lo,med,hi=rw['lo'],rw['med'],rw['hi']
        if hi>lo: _rect(s,X(lo),y+rowh/2-0.03,max(X(hi)-X(lo),0.02),0.06,tint(col,0.45) if not b else tint(col,0.25))
        _rect(s,X(med)-0.07,y+rowh/2-0.095,0.14,0.19,col)
        txt(s,X(hi)+0.10,y+rowh/2-0.09,val_w+0.3,0.18,rw.get('txt') or (f"{lo:.0f}–{hi:.0f}" if hi>lo+0.5 else f"{med:.0f}"),8.2,True,DARKAMBER if b else NAVY)
        y+=rowh
    for tv in ticks:
        txt(s,X(tv)-0.35,y+0.02,0.70,0.14,f"{tv}",7.2,False,GREY,PP_ALIGN.CENTER)
    if bandlab:
        txt(s,X(band[0])-1.6,y+0.18,X(band[1])-X(band[0])+3.2,0.14,bandlab,7.2,True,DARKAMBER,PP_ALIGN.CENTER)
    return y+0.34

def brand_rows(rows_,col,cap=None):
    """Групування SKU за брендом → рядки для rangebars (за зростанням медіани)."""
    d=co.defaultdict(list)
    for r in rows_:
        if r.get('pmed'): d[r['tm']].append(r)
    out=[]
    for tm,v in d.items():
        pm=[x['pmed'] for x in v]; ws=sorted(set(round(x['w']) for x in v))
        wt=(f"{ws[0]} г" if len(ws)==1 else f"{ws[0]}–{ws[-1]} г")
        out.append(dict(label=tm,wt=wt,n=len(v),lo=min(pm),med=st.median(pm),hi=max(pm),col=col))
    out.sort(key=lambda x:x['med'])
    return out[:cap] if cap else out

def slide_tail_sources(s,text,y=7.12):
    txt(s,0.62,y,12.10,0.30,text,7.2,False,PALE,lh=1.25)

prs=new_deck(); N=[0]
def num(): N[0]+=1; return f"{N[0]:02d}"
TOP2CH={t:len({c for r in SD_ROWS if r['tm']==t for c in r['chains']}|({'silpo'} if t in('Samyang','Nongshim') else set())) for t in TOP2}

# ============ 01 ОБКЛАДИНКА ============
s=prs.slides.add_slide(prs.slide_layouts[6]); N[0]=1
_rect(s,0,0,13.333,7.5,NAVY); _rect(s,0,7.20,13.333,0.30,AMBER)
s.shapes.add_picture(LOGO,Inches(0.90),Inches(0.80),Inches(2.10),Inches(0.91))
txt(s,0.92,2.52,11.0,0.25,"ДОСЛІДЖЕННЯ РИНКУ ТА ПОЗИЦІОНУВАННЯ · ЖОВТЕНЬ 2026",10.5,True,GOLD)
txt(s,0.90,2.88,11.5,1.60,"Корейський рамен\nв Україні",46,True,WHITE,lh=1.05)
txt(s,0.92,4.98,11.0,0.70,
 "Оцінка шести SKU SIAS — рамену в пакеті — проти ринкової полиці та ціна закупівлі для полиці 120 грн:\n"
 f"{F['n_sku']} SKU корейського рамену · {F['n_brands']} брендів · {F['n_chains']} мереж із {F['chains_checked']} · зріз каталогів 29.09.2026 · митна база 2025–2026",
 11,False,PALE,lh=1.5)
_rect(s,0.92,5.96,0.05,0.92,AMBER)
txt(s,1.18,5.99,11.0,0.88,
 "Усі ціни — роздріб за одну пачку, з ПДВ, на дату зрізу. Опт, ящики та мультипаки до медіан не входять.\n"
 "Фото товарів — з каталогів мереж, як ідентифікація SKU. Фінансова модель — RAMEN_FINMODEL_4_TABS_PRESENTATION.xlsx.",
 9,False,PALE,lh=1.45)

# ============ 02 ШІСТЬ ЦИФР ============
cust_x=838.480003/239.67149
s=slide(prs,"ГОЛОВНЕ","Шість цифр, які описують ринок",num(),
 "Усі ціни в колоді — за одну упаковку, з ПДВ. Опт і ящики до розрахунку не входять.")
estat(s,0.62,1.78,"ПОЗИЦІЙ У ПРОДАЖУ",str(F['n_sku']),f"{F['n_brands']} брендів · 3 сегменти · {F['n_chains']} мереж із 18. Кожна позиція має ціну в каталозі.",DARKAMBER,3.88,1.42,26)
estat(s,4.72,1.78,"ІМПОРТ ЗА МИТНИЦЕЮ","838,5 т","Січень–травень 2026, код УКТ ЗЕД 1902 30 10 00. За весь 2025 рік — 711,6 т.",TEAL,3.88,1.42,26)
estat(s,8.82,1.78,"2026 ПРОТИ 2025",f"×{d1(cust_x)}","Ті самі п’ять місяців: 239,7 → 838,5 т (+249,8 %); інвойс $1,02 → $3,20 млн.",PURPLE,3.90,1.42,26)
estat(s,0.62,3.34,"ІМПОРТ З КОРЕЇ","118,5 т","Січень–травень 2026: +116,3 % до 54,8 т торік. Це 14 % коду; лідер — В’єтнам, 166,7 т.",NAVY,3.88,1.42,26)
estat(s,4.72,3.34,"РОЗРИВ ЦІН ФОРМАТІВ",f"×{d1(SS['3']['g']/SS['1']['g'])}",f"Ціна за 100 г: стакан {SS['3']['g']} грн проти {SS['1']['g']} грн у пакеті з бульйоном.",DARKAMBER,3.88,1.42,26)
estat(s,8.82,3.34,"ЦІЛЬОВА ПОЛИЦЯ SIAS","120 грн",f"P75 серед {len(PACK)} пакетів: +{(120/PACK_MED-1)*100:.0f} % до медіани {d1(PACK_MED)} грн; дорожчих — 18 із 71.",RED,3.90,1.42,26)
ecall(s,0.62,4.94,5.92,1.80,"ЩО ЦЕ ОЗНАЧАЄ ДЛЯ ПОЛИЦІ",
 "1. Категорія повністю імпортна і швидко росте: ввіз за п’ять місяців 2026 року вже більший за весь 2025.\n"
 "2. Корея росте повільніше за код в цілому: її частка впала з 23 % до 14 %.\n"
 f"3. Полиця концентрована: Samyang і Nongshim дають {N_TOP2} позицій з {F['n_sku']} і стоять у {min(TOP2CH.values())}–{max(TOP2CH.values())} мережах.",AMBER)
ecall(s,6.80,4.94,5.92,1.80,"ЩО ПОТРІБНО ВІД SIAS",
 "1. Полиця 120 грн за моделлю з маржею 30 % вимагає закупівлі €0,54 (повна машина) або €0,32 (6 палет).\n"
 "2. Зараз SIAS пропонує €0,65–0,75: потрібне зниження — на 17–28 % для повної машини і на 51–57 % для 6 палет.\n"
 "3. Тому торг іде про ціну закупівлі, а не про ціну на полиці.",RED)
slide_tail_sources(s,f"Джерела: каталоги мереж 29.09.2026 ({F['n_sku']} SKU); митна база користувача, код 1902 30 10 00, січень–травень 2025 і 2026; фінмодель RAMEN_FINMODEL_4_TABS (вкладки «Цена 120 грн»).",6.92)

# ============ 03 ОБСЯГ РИНКУ · МИТНА БАЗА ============
s=slide(prs,"ОБСЯГ РИНКУ · МИТНА БАЗА","Скільки локшини ввозять в Україну",num(),
 "Код УКТ ЗЕД 1902 30 10 00: обсяг товарного коду, а не роздрібні продажі рамену. Однакові п’ять місяців для порівняння.")
txt(s,0.62,1.78,6.0,0.16,"ОБСЯГ ІМПОРТУ, ТОНН",8.6,True,NAVY)
hbars(s,0.62,2.10,6.20,[("2025 · січень–травень",239.67,GREYBAR,"239,7 т"),
                       ("2026 · січень–травень",838.48,AMBBAR,"838,5 т"),
                       ("2025 · увесь рік",711.6,NAVYBAR,"711,6 т")],
      900,[0,300,600,900],lab_w=1.95,val_w=1.20,rowh=0.50,bar_h=0.26)
txt(s,0.62,4.02,6.0,0.16,"ІНВОЙС, МЛН ДОЛАРІВ",8.6,True,NAVY)
hbars(s,0.62,4.34,6.20,[("2025 · січень–травень",1.02459,GREYBAR,"$1,02 млн"),
                       ("2026 · січень–травень",3.20062,AMBBAR,"$3,20 млн")],
      3.6,[0,1,2,3],lab_w=1.95,val_w=1.20,rowh=0.50,bar_h=0.26)
estat(s,7.10,1.78,"ОБСЯГ","+249,8 %","239,7 → 838,5 т за січень–травень: ×3,5.",AMBER,2.72,1.30,21)
estat(s,10.00,1.78,"ІНВОЙС","+212,4 %","$1,02 → $3,20 млн: ×3,1.",TEAL,2.72,1.30,21)
ecall(s,7.10,3.22,5.62,1.78,"ЩО КАЖЕ СЕРЕДНЯ ЦІНА ІМПОРТУ",
 "Митна вартість за кілограм: $4,27 у 2025 → $3,82 у 2026, тобто −10,7 %. Ввіз росте швидше за гроші: ймовірно, в обсязі додається дешевший товар.\n"
 "Для SIAS це означає: категорія розширюється за рахунок цінового низу, а не премії.",RED)
ecall(s,7.10,5.12,5.62,1.62,"ЧОГО ЦЕ НЕ ПОКАЗУЄ",
 "Код охоплює товар усіх країн походження, не лише корейський. Частки брендів і форматів у митній базі не видно — вони взяті з каталогів мереж.",NAVY)
ecall(s,0.62,5.80,6.20,0.96,"ЯК ЧИТАТИ ЦИФРИ",
 "Порівняння — за однакові п’ять місяців (01–05). За увесь 2025 рік — 711,6 т: за п’ять місяців 2026-го ввезено вже на 18 % більше.",AMBER)
slide_tail_sources(s,"Митна база користувача, код УКТ ЗЕД 1902 30 10 00, січень–травень 2025 і 2026; повний 2025 рік — 711,6 т. Митна вартість на кілограм — власний розрахунок: інвойс ÷ вага.",6.92)

# ============ 04 МИТНА БАЗА · ПОХОДЖЕННЯ ============
s=slide(prs,"МИТНА БАЗА · ПОХОДЖЕННЯ","Звідки їде: Корея — друга за обсягом країна",num(),
 "Січень–травень 2026, код УКТ ЗЕД 1902 30 10 00. Країни за вагою нетто, тонн.")
ctry=[("В’єтнам",166.687505),("Корея",118.497356),("Румунія",86.919064),("Велика Британія",58.564844),
      ("Латвія",57.30452),("Чехія",50.800755),("Польща",48.761462)]
oth=838.480003-sum(v for _,v in ctry)
txt(s,0.62,1.78,6.0,0.16,"ПО КРАЇНАХ ПОХОДЖЕННЯ, ТОНН",8.6,True,NAVY)
hbars(s,0.62,2.10,6.30,[(k,v,(AMBBAR if k=="Корея" else GREYBAR),f"{th1(v)} т") for k,v in ctry]+
      [("Інші країни разом",oth,NAVYBAR,f"{th1(oth)} т")],300,[0,100,200,300],lab_w=1.95,val_w=1.20,rowh=0.42,bar_h=0.22)
txt(s,7.10,1.78,5.6,0.16,"КОРЕЯ → УКРАЇНА · ПОРІВНЯННЯ ОДНАКОВИХ ПЕРІОДІВ",8.6,True,NAVY)
table(s,2.10,[("СІЧЕНЬ–ТРАВЕНЬ",7.10,1.90,L),("2025",9.00,1.20,R),("2026",10.30,1.20,R),("ЗМІНА",11.60,1.12,R)],
 [["Обсяг, т","54,8","118,5","+116,3 %"],
  ["Інвойс, тис. $","348,6","809,2","+132,1 %"],
  ["$ за кг","6,36","6,83","+7,4 %"],
  ["Частка в імпорті","22,9 %","14,1 %","−8,7 п.п."]],
 head_size=7.4,row_size=10,rh=0.38,left=7.10,bandw=5.62)
ecall(s,7.10,3.98,5.62,1.62,"КОРЕЯ ДОРОЖЧА ЗА СЕРЕДНЄ В 1,8 РАЗА",
 "Митна вартість корейського ввозу — $6,83/кг проти $3,82 в цілому по коду. Покупець у категорії вже платить премію за корейський товар.",TEAL)
ecall(s,7.10,5.78,5.62,1.00,"АЛЕ ЧАСТКА КОРЕЇ ПАДАЄ",
 "Код в цілому виріс на 250 %, Корея — на 116 %: різницю забрав дешевший ввіз.",AMBER)
ecall(s,0.62,5.78,6.30,1.00,"СКІЛЬКИ ЦЕ В ПАЧКАХ",
 "118,5 т при середній вазі пачки 120 г — близько 0,99 млн пачок за п’ять місяців. Це оцінка порядку величини, не продажі.",NAVY)
slide_tail_sources(s,"Митна база користувача, код УКТ ЗЕД 1902 30 10 00, січень–травень. Частка Кореї, $/кг і кількість пачок — власний розрахунок. Країна походження за митною декларацією.",6.92)
# ============ 05 СЕГМЕНТАЦІЯ ============
def pick_photo(k,must=None):
    arr=[c for c in SEG[k] if c.get('img') and os.path.exists(c['img']) and (not must or re.search(must,c['name'],re.I))]
    arr.sort(key=lambda c:-(c.get('nch') or 0)); return arr[0] if arr else None
PH={'1':pick_photo('1','Shin Ramyun'),'2':pick_photo('2','Carbonara'),'3':pick_photo('3')}
s=slide(prs,"СЕГМЕНТАЦІЯ","Сегментація за способом приготування",num(),
 f"Три сегменти, {F['n_sku']} позицій: спершу спосіб приготування, далі вага, ціна й бренди. Належність наших шести SKU до сегментів 1 чи 2 потребує інструкцій приготування.")
_x=[0.62,4.72,8.82]
for i,k in enumerate(("1","2","3")):
    v=SS[k]; l=_x[i]; col=SCOL[k]
    _rect(s,l,1.78,3.90,0.40,col)
    txt(s,l+0.14,1.86,3.62,0.26,SNAME[k],8.2,True,WHITE,lh=1.1)
    _rect(s,l+0.07,2.26,3.76,1.14,LIGHT)
    ph=PH[k]
    if ph:
        try:
            from PIL import Image
            im=Image.open(ph['img']); ar=im.width/im.height; h=1.0; w=min(h*ar,3.6)
            s.shapes.add_picture(ph['img'],Inches(l+0.07+(3.76-w)/2),Inches(2.26+(1.14-h)/2),Inches(w),Inches(h))
        except Exception: pass
    txt(s,l+0.14,3.46,1.30,0.40,f"{v['n']}",26,True,col)
    txt(s,l+1.10,3.62,2.60,0.20,f"позицій · {round(v['n']*100/F['n_sku'])} %",8.6,False,GREY)
    for j,(lab,val) in enumerate((("ВАГА",f"{v['wlo']}–{v['whi']} г"),("МЕДІАНА ЦІНИ ЗА ПАЧКУ",f"{v['med']} грн"),("ЦІНА ЗА 100 Г",f"{v['g']} грн"),("БРЕНДІВ",f"{v['brands']}"))):
        yy=4.00+j*0.27
        txt(s,l+0.14,yy,2.00,0.16,lab,7,True,GREY)
        txt(s,l+2.00,yy-0.02,1.78,0.20,val,9.6,True,NAVY,PP_ALIGN.RIGHT)
txt(s,0.62,5.14,8.0,0.16,"СЕГМЕНТ × ЦІНОВА СМУГА · ПОЗИЦІЙ, грн за пачку",7.8,True,NAVY)
_grid=[[sum(1 for r in SD_ROWS if r['seg']==int(k) and r.get('pmed') and lo<=r['pmed']<hi) for lo,hi,_ in PB] for k in ("1","2","3")]
heat(s,0.62,5.30,[p[2] for p in PB],[SSHORT[k] for k in ("1","2","3")],_grid,rl=1.60,cw=1.10,ch=0.34,
     tail=[f"{SS[k]['n']} SKU · медіана {SS[k]['med']} грн" for k in ("1","2","3")],tail_w=2.30,tail_head="УСЬОГО")
slide_tail_sources(s,"Сегмент визначено за способом приготування й форматом упаковки на фото картки; медіана — за медіанними цінами позицій у каталогах 18 мереж, 29.09.2026. Сільпо (4 позиції без ваги) до сегментів не віднесено.",6.98)

# ============ 06 ЦІНОВІ ПОВЕРХИ ============
s=slide(prs,"ЦІНОВІ ПОВЕРХИ","Три поверхи локшини швидкого приготування",num(),
 "Поверх визначено за ціною за 100 г — це єдиний коректний вимір, коли пачки важать від 50 до 200 г. Межі взяті з фактичного розподілу брендів, а не задані наперед.")
tierpanel(s,0.62,1.74,3.90,3.40,BLUE,"ПОВЕРХ 1 · ЕКОНОМ","9–25 грн  ·  14–37 грн/100 г",
 "Українське виробництво і приватні марки мереж · 15 брендів",
 [tl("Мівіна"),tl("Reeva"),tl("Роллтон"),tl("Щедро")])
tierpanel(s,4.72,1.74,3.90,3.40,TEAL,"ПОВЕРХ 2 · МІДЛ","38–124 грн  ·  58–86 грн/100 г",
 "Азія та імпорт середньої ціни — сюди потрапляє й дешева Корея · 9 брендів",
 [tl("O'Food"),tl("Ottogi"),tl("Paldo"),tl("Yum Yum")])
tierpanel(s,8.82,1.74,3.90,3.40,DARKAMBER,"ПОВЕРХ 3 · ІМПОРТНИЙ ПРЕМІУМ","34–198 грн  ·  93–146 грн/100 г",
 "Samyang, Nongshim та імпортна азійська локшина · 9 брендів",
 [tl("Samyang"),tl("Nongshim"),tl("Мама"),tl("Micoem")])
_rect(s,0.62,5.24,12.10,0.76,LIGHT); _rect(s,0.62,5.24,0.045,0.76,RED)
per_t=[120/w*100 for w in (112.5,113,123,131)]
txt(s,0.90,5.32,3.50,0.14,"ДЕ ОПИНЯЮТЬСЯ НАШІ ШІСТЬ SKU",8.8,True,RED)
txt(s,0.90,5.52,3.50,0.40,f"{min(per_t):.0f}–{max(per_t):.0f} грн/100 г при полиці 120 грн;\n124–132 за поточною ціною",9.0,True,NAVY,lh=1.25)
txt(s,4.60,5.32,7.90,0.66,
 f"При полиці 120 грн наші SKU стають між Samyang (93) і Nongshim (107 грн/100 г) — двома брендами, що тримають {N_TOP2} позицій з {F['n_sku']}. "
 "Поточна розрахункова полиця 140–162 грн виводить їх вище за обох.",8.6,False,NAVY,lh=1.38)
ecall(s,0.62,6.14,12.10,1.28,"ЩО ЦЕ ЗМІНЮЄ В КАРТИНІ КОНКУРЕНЦІЇ",
 "Корейський рамен — не однорідний преміум. Ottogi (75), Paldo (77), O’Food (67) і Bibigo (70 грн/100 г) стоять у МІДЛІ поруч із Yum Yum, а не поруч із Samyang і Nongshim; саме ця дешева половина (67–89 грн за пачку) — прямий ціновий конкурент новому бренду.\n"
 "На пропозицію впливає й фабрика Nestlé в Смолигові (Волинь): запущена в липні 2025, близько $50 млн; ≈ 75 % випуску — на експорт до ЄС під Maggi, решта — на внутрішній ринок (Latifundist, Just Food).",NAVY,9.4,8.2)

# ============ БЛОКИ СЕГМЕНТІВ ============
SEGROWS={k:[r for r in SD_ROWS if r['seg']==int(k) and r.get('pmed')] for k in "123"}
NOW_LOW=M.forward_current("33")[2]['shelf']      # 140,55  (4 SKU, 33 палети, маржа 30 %)
NOW_HIGH=M.forward_current("33")[0]['shelf']     # 162,18  (2 SKU)
def n_above(arr,x): return sum(1 for p in arr if p>x)
def sample(k,n=7):
    arr=sorted([c for c in SEG.get(k,[]) if c.get('pmed')],key=lambda c:c['pmed']); out=[]
    for i in range(n):
        j=round(i*(len(arr)-1)/(n-1)) if len(arr)>1 else 0
        if arr[j] not in out: out.append(arr[j])
    return out
def band_text(k):
    v=SS[k]; return f"СЕГМЕНТ {k} · {SNAME[k]}   ·   {SPREP[k]}   ·   вага {v['wlo']}–{v['whi']} г"

def segoverview(k,title,sub,note,keys=None):
    v=SS[k]; col=SCOL[k]
    s=slide(prs,f"СЕГМЕНТ {k} · {SNAME[k]}",title,num(),sub)
    estat(s,0.62,1.78,"ПОЗИЦІЙ У СЕГМЕНТІ",f"{v['n']} SKU",f"{v['brands']} брендів у {v['chains']} мережах.\n{round(v['n']*100/F['n_sku'])} % усієї категорії.",col,3.10,1.22,22)
    estat(s,3.88,1.78,"МЕДІАНА ПОЛИЦІ",f"{v['med']} грн",f"P25–P75: {v['p25']}–{v['p75']} грн.\nДіапазон: {v['lo']}–{v['hi']} грн.",col,3.10,1.22,22)
    estat(s,7.14,1.78,"ЦІНА ЗА 100 Г",f"{v['g']} грн","за 100 г — медіана сегмента.\nНаші SKU при полиці 120 грн: 92–107.",col,3.10,1.22,22)
    estat(s,10.40,1.78,"ВАГОВИЙ ДІАПАЗОН",f"{v['wlo']}–{v['whi']} г","фактичні межі сегмента\nза зрізом 29.09.2026.",col,2.32,1.22,19)
    txt(s,0.62,3.18,5.60,0.14,"ХТО ТРИМАЄ СЕГМЕНТ · SKU НА БРЕНД",8,True,NAVY)
    mx=max(n for _,n in v['blist'])
    minibar(s,0.62,3.42,5.30,[(b,n,f"{n}") for b,n in v['blist'][:6]],mx,col,lab_w=1.30,val_w=0.50,rowh=0.205)
    txt(s,6.60,3.18,6.12,0.14,"ЩО ЦЕ ОЗНАЧАЄ ДЛЯ НАС",8,True,NAVY)
    txt(s,6.60,3.42,6.12,1.30,note,8.8,False,GREY,lh=1.45)
    y=section_head(s,4.82,"ПРЕДСТАВНИКИ СЕГМЕНТА · ВІД НАЙДЕШЕВШОЇ ДО НАЙДОРОЖЧОЇ ПОЗИЦІЇ","фото — з каталогів мереж")
    cardrow(s,y+0.06,sample(k),col)
    return s

def catalog(k,items,i,m,extra_band=None):
    v=SS[k]; col=SCOL[k]
    pm=[c['pmed'] for c in items]; g_=[c['per100'] for c in items]
    lo,hi=min(pm),max(pm)
    ws=sorted(set(c['w'] for c in items))
    title=f"{STITLE[k]}: каталог {i} з {m} — {lo:.0f}–{hi:.0f} грн"
    s=slide(prs,f"СЕГМЕНТ {k} · {SNAME[k]}   ·   {v['n']} ПОЗИЦІЙ",title,num(),
            f"{len(items)} позицій за медіанною ціною {lo:.0f}–{hi:.0f} грн ({min(g_):.0f}–{max(g_):.0f} грн/100 г). Відсортовано від найдешевшої позиції.")
    band(s,1.68,band_text(k),col)
    cards(s,items,col)
    brands=co.Counter(c['tm'] for c in items).most_common(3)
    bt=", ".join(f"{b} ({n})" for b,n in brands)
    a=sum(1 for p in pm if p<=120); b_=len(pm)-a
    if 120<lo: pos=f"Уся ця частина дорожча за нашу ціль 120 грн (мінімум {lo:.0f} грн)."
    elif 120>hi: pos=f"Уся ця частина дешевша за нашу ціль 120 грн (максимум {hi:.0f} грн)."
    else: pos=f"Наша ціль 120 грн потрапляє в цю частину: дешевших за неї {a} із {len(pm)}, дорожчих — {b_}."
    mg=st.median(g_); ratio=99/mg
    gtxt=(f"Ціна за 100 г у цій частині — медіана {mg:.0f} грн; наші SKU при полиці 120 грн мають 92–107 — "+
          (f"у {d1(ratio)} раза більше." if ratio>=1.15 else ("на рівні." if ratio>0.87 else f"на {(1-ratio)*100:.0f} % менше.")))
    ecall(s,0.62,6.05,12.10,0.90,"ЩО ЦЕ ОЗНАЧАЄ ДЛЯ ПОЛИЦІ",
          f"Бренди частини: {bt}. {pos} {gtxt}",col,9.4,8.4)
    return s
def catalog_series(k,items=None):
    items=items if items is not None else SEG[k]
    m=math.ceil(len(items)/14); size=math.ceil(len(items)/m)
    for i in range(m): catalog(k,items[i*size:(i+1)*size],i+1,m)

ALLCH=[("auchan","Ашан",1),("novus","NOVUS",1),("metro","METRO",1),("zaraz","«Зараз»",1),("silpo","«Сільпо»",1),
       ("megamarket","МегаМаркет",1),("ultramarket","Ultramarket",1),("ekomarket","ЕКО",1),("epicentr","Епіцентр",1),
       ("grono","Grono",0),("onde","Onde",0),("kharkiv","«Клас»",0),("vostorg","«Восторг»",0),("cosmos","«Космос»",0),
       ("tavriav","«Таврія В»",0),("chudomarket","«ЧудоМаркет»",0),("torba","«Торба»",0),("ideal","«Ідеал»",0)]
SILPO={'1':1,'2':3,'3':0,'4':0}
def chipflow(s,x,y,w,items,col,chip_h=0.30,gap=0.10):
    cx=x; cy=y
    for name,cnt,star in items:
        cw=0.24+0.078*len(name)+(0.46 if cnt is not None else 0.0)
        if cx+cw>x+w: cx=x; cy+=chip_h+0.08
        _rect(s,cx,cy,cw,chip_h,LIGHT)
        txt(s,cx+0.10,cy+0.065,cw-(0.5 if cnt is not None else 0.14),0.18,name,8.4,False,INK)
        if cnt is not None:
            _rect(s,cx+cw-0.40,cy+0.04,0.34,0.22,col)
            txt(s,cx+cw-0.40,cy+0.065,0.34,0.17,(f"{cnt}*" if star else f"{cnt}"),8,True,WHITE,PP_ALIGN.CENTER)
        cx+=cw+gap
    return cy+chip_h
def wheresold(keys,title,kick,note):
    col=SCOL[keys[0]]
    cnt=co.Counter()
    for r in SD_ROWS:
        if str(r['seg']) in keys:
            for ch in r['chains']: cnt[ch]+=1
    star=set()
    for k in keys:
        if SILPO[k]: cnt['silpo']+=SILPO[k]; star.add('silpo')
    tot=sum(SS[k]['n'] for k in keys)
    present=[(c,n) for c,n in cnt.most_common() if n>0]
    s=slide(prs,kick,title,num(),
            f"{len(present)} мереж із 18 мають позиції цього сегмента; число поруч — скільки позицій сегмента є в мережі.")
    txt(s,0.62,1.80,5.0,0.16,"ПОЗИЦІЙ СЕГМЕНТА В МЕРЕЖІ",8.2,True,NAVY)
    nm={c:n for c,n,_ in ALLCH}
    rows=[(nm[c],n,col if c not in star else tint(col,0.35),(f"{n}*" if c in star else f"{n}")) for c,n in present]
    hbars(s,0.62,2.12,5.70,rows,max(n for _,n in present)*1.1,[0],lab_w=1.55,val_w=0.70,rowh=min(0.34,4.45/len(rows)),bar_h=0.19)
    # праворуч — групи
    nat=[(n_,cnt.get(c,0),c in star) for c,n_,t in ALLCH if t==1 and cnt.get(c,0)>0]
    reg=[(n_,cnt.get(c,0),c in star) for c,n_,t in ALLCH if t==0 and cnt.get(c,0)>0]
    zero=[(n_,None,False) for c,n_,t in ALLCH if cnt.get(c,0)==0]
    y=1.80
    def grp(y,head,tag,tagcol,items,h):
        _rect(s,6.60,y,0.045,h,tagcol)
        txt(s,6.82,y+0.06,3.80,0.20,head,10.5,True,NAVY)
        _rect(s,11.10,y+0.07,1.62,0.22,tagcol); txt(s,11.10,y+0.10,1.62,0.16,tag,6.4,True,WHITE,PP_ALIGN.CENTER)
        chipflow(s,6.82,y+0.42,5.88,items,tagcol)
    ny=0.42+0.38*math.ceil(max(len(nat),1)*1.9/5.9)+0.15
    grp(y,f"Національні мережі · {len(nat)}","ЦІЛЬОВИЙ КАНАЛ",RED,nat,1.52); y+=1.52+0.10
    grp(y,f"Регіональні мережі · {len(reg)}","РЕГІОНАЛЬНИЙ ВХІД",TEAL,reg,1.52); y+=1.52+0.10
    grp(y,f"Без позицій сегмента · {len(zero)}","ВІЛЬНА ПОЛИЦЯ",PALE,zero,1.62)
    slide_tail_sources(s,note,6.98)
    return s

# ============ СЕГМЕНТ 1 ============
_p1=[r['pmed'] for r in SEGROWS['1']]; _p2=[r['pmed'] for r in SEGROWS['2']]
S1,S2,S3=SS['1'],SS['2'],SS['3']
segoverview("1","Пакет з бульйоном: найбільший сегмент",
 "Основна частина корейської полиці: локшина й порошковий бульйон, вариться в каструлі. Показано сім позицій, рівновіддалених по ціні.",
 f"Сюди найімовірніше потраплять чотири наші SKU по 112,5–113 г — якщо їх спосіб приготування варіння з бульйоном (не підтверджено).\n\n"
 f"Ціль 120 грн — на {(120/S1['med']-1)*100:.0f} % вище за медіану сегмента ({S1['med']} грн) і вище за P75 ({S1['p75']} грн): дорожчих за 120 лише {n_above(_p1,120)} із {S1['n']}. "
 f"Поточна розрахункова полиця {d2(NOW_LOW)} грн дорожча за всіх, крім {n_above(_p1,NOW_LOW)}.")
catalog_series("1")
wheresold(["1"],"Пакет з бульйоном: де продається","СЕГМЕНТ 1 · ДЕ ПРОДАЄТЬСЯ",
 "Суцільна перевірка каталогів 18 мереж, 29.09.2026. * «Сільпо» — позиція без ваги в каталозі, віднесена за назвою (Nongshim Shin Ramyun).")

# ============ СЕГМЕНТ 2 ============
segoverview("2","Пакет із соусом: 30 позицій, медіана 106 грн",
 "Локшина, з якої воду зливають, а смак дає соус: Buldak, Toomba, Chapaghetti, Volcano, Carbonara. Показано сім позицій, рівновіддалених по ціні.",
 f"Сюди найімовірніше потраплять Gouda Cheese 123 г і Carbonara Spicy 131 г — якщо їх спосіб приготування це підтвердить.\n\n"
 f"Ціль 120 грн — на {(120/S2['med']-1)*100:.0f} % вище за медіану сегмента ({S2['med']} грн), але нижче P75 ({S2['p75']} грн): дорожчих за 120 — {n_above(_p2,120)} із {S2['n']}. "
 f"Прямі аналоги за вагою: Paldo Volcano Carbonara 130 г — 114 грн, Choi’s Carbonara 131 г — 125,20 грн. Поточна полиця {d2(NOW_HIGH)} грн дорожча за {S2['n']-n_above(_p2,NOW_HIGH)} із {S2['n']}.")
catalog_series("2")
wheresold(["2"],"Пакет із соусом: де продається","СЕГМЕНТ 2 · ДЕ ПРОДАЄТЬСЯ",
 "Суцільна перевірка каталогів 18 мереж, 29.09.2026. * «Сільпо» — три позиції Samyang Buldak без ваги в каталозі, віднесені за назвою.")

# ============ СЕГМЕНТ 3 І 4 ============
segoverview("3","Стакан: найдорожчий формат за 100 г",
 "Заливається окропом, готується без плити: формат для споживання поза домом. Наші шість SKU — пакет, тому це ринковий контекст, а не пряма конкуренція.",
 f"Ринок вже платить премію за зручність: {S3['g']} грн/100 г проти {S1['g']} у пакеті з бульйоном — у {S3['g']/S1['g']:.1f} раза більше.\n\n"
 f"Сегмент вузький: {S3['n']} позицій у {S3['chains']} мережах і лише {S3['brands']} бренди. Якщо в лінійці SIAS з’явиться стакан, це єдиний формат, де премія за грам уже прийнята ринком; у поточній пропозиції стакана немає.")
catalog_series("3")


# ============ 20 ЗВЕДЕННЯ · БРЕНДИ × МЕРЕЖІ ============
SV_CH=[("auchan","Ашан"),("grono","Grono"),("onde","Onde"),("metro","METRO"),("zaraz","«Зараз»"),("kharkiv","«Клас»"),
       ("vostorg","«Восторг»"),("cosmos","«Космос»"),("tavriav","«Таврія В»"),("novus","NOVUS"),("chudomarket","«ЧудоМ.»"),
       ("torba","«Торба»"),("silpo","«Сільпо»")]
SV_BR=["Nongshim","Samyang","Paldo","Ottogi","O'Food","Bibigo","Otoki","fire bull","Choi's"]
SILPO_P={"Samyang":[89.99,89.99,199.0],"Nongshim":[139.0]}      # «Сільпо»: ціни з каталогу, без EAN/ваги
SILPO_BR={b:(len(v),st.median(v)) for b,v in SILPO_P.items()}
def _cell(b,ch):
    if ch=="silpo": return SILPO_BR.get(b,(0,None))
    v=[z for z in ZREC if z['chain']==ch and _EANS[z['ean']]['tm']==b]
    return (len({z['ean'] for z in v}),st.median([z['price'] for z in v]) if v else None)
SV=[[_cell(b,ch) for ch,_ in SV_CH] for b in SV_BR]
def _tot(b):
    rows_b=[r for r in SD_ROWS if r['tm']==b]
    chs={c for r in rows_b for c in r['chains']}|({"silpo"} if b in SILPO_BR else set())
    n=len(rows_b)+(SILPO_BR[b][0] if b in SILPO_BR else 0)
    pm=[r['pmed'] for r in rows_b if r.get('pmed')]
    return n,len(chs),(st.median(pm) if pm else None)
s=slide(prs,"ЗВЕДЕННЯ · БРЕНДИ × МЕРЕЖІ","Усі бренди в усіх мережах: скільки позицій і за скільки",num(),
 "У кожній клітинці — кількість позицій бренду в мережі (колір) і медіанна ціна за пачку, грн. Зріз 29.09.2026, 13 мереж, де є категорія.")
L0=0.62; RL=1.28; CW=0.655; TW=0.62; T0=1.70; HH=0.42; RH=0.395
x_tail=L0+RL+CW*13+0.10
for j,(ch,nm) in enumerate(SV_CH):
    _rect(s,L0+RL+j*CW+0.02,T0,CW-0.04,HH-0.04,NAVY)
    txt(s,L0+RL+j*CW+0.02,T0+0.13,CW-0.04,0.18,nm,6.8,True,WHITE,PP_ALIGN.CENTER)
for j,hd in enumerate(("SKU","МЕРЕЖ","МЕДІАНА")):
    _rect(s,x_tail+j*TW,T0,TW-0.03,HH-0.04,DARKAMBER)
    txt(s,x_tail+j*TW,T0+0.13,TW-0.03,0.18,hd,6.8,True,WHITE,PP_ALIGN.CENTER)
mx=max(n for row in SV for n,_ in row)
y=T0+HH
for i,b in enumerate(SV_BR):
    nm_=("Choi’s (SIAS)" if b=="Choi's" else b)
    txt(s,L0,y+RH/2-0.09,RL-0.10,0.18,nm_,8.6,True,NAVY,PP_ALIGN.RIGHT)
    for j,(n,p) in enumerate(SV[i]):
        x=L0+RL+j*CW
        step=0 if n==0 else 1+int(round((len(SEQT)-2)*(n/mx)))
        _rect(s,x+0.02,y+0.02,CW-0.04,RH-0.04,SEQT[step])
        if n:
            col=WHITE if step>=4 else NAVY
            txt(s,x+0.02,y+0.04,CW-0.04,0.18,str(n),10,True,col,PP_ALIGN.CENTER)
            txt(s,x+0.02,y+0.23,CW-0.04,0.14,f"{p:.0f}",7.2,False,col,PP_ALIGN.CENTER)
        else:
            txt(s,x+0.02,y+0.12,CW-0.04,0.18,"—",9,False,PALE,PP_ALIGN.CENTER)
    n,nc,pm=_tot(b)
    for j,val in enumerate((str(n),str(nc),f"{pm:.0f} грн" if pm else "—")):
        _rect(s,x_tail+j*TW,y+0.02,TW-0.03,RH-0.04,LIGHT)
        txt(s,x_tail+j*TW,y+0.12,TW-0.03,0.18,val,9,True,NAVY,PP_ALIGN.CENTER)
    y+=RH
# підсумковий рядок
_rect(s,L0+RL,y+0.04,CW*13+TW*3+0.10,0.012,NAVY)
txt(s,L0,y+0.14,RL-0.10,0.18,"УСЬОГО",8.6,True,DARKAMBER,PP_ALIGN.RIGHT)
for j,(ch,_) in enumerate(SV_CH):
    n=sum(SV[i][j][0] for i in range(len(SV_BR)))
    pr=[z['price'] for z in ZREC if z['chain']==ch] if ch!="silpo" else SILPO_P["Samyang"]+SILPO_P["Nongshim"]
    txt(s,L0+RL+j*CW,y+0.10,CW,0.18,str(n),10,True,DARKAMBER,PP_ALIGN.CENTER)
    txt(s,L0+RL+j*CW,y+0.28,CW,0.14,f"{st.median(pr):.0f}",7.2,False,GREY,PP_ALIGN.CENTER)
tn=len(SD_ROWS)+4
for j,val in enumerate((str(tn),"13",f"{st.median([r['pmed'] for r in SD_ROWS]):.0f} грн")):
    txt(s,x_tail+j*TW,y+0.18,TW-0.03,0.18,val,9,True,DARKAMBER,PP_ALIGN.CENTER)
y+=0.50
nong=[sum(v[0] for v in SV[SV_BR.index(b)]) for b in ("Nongshim","Samyang")]
estrip(s,y+0.02,"ЩО ПОКАЗУЄ ЗВЕДЕННЯ",
 f"Nongshim і Samyang — {N_TOP2+4} із {tn} позицій; перший у 11 мережах із 13, другий у 12. Choi’s (SIAS) — одна позиція у двох мережах («Космос», «Таврія В»), 125 грн. Решта шість брендів — від 1 до 14 позицій, переважно в «Ашані».",NAVY,0.70)
slide_tail_sources(s,"Джерела: каталоги мереж (zakaz.ua) і «Сільпо», 29.09.2026. Клітинка = унікальні EAN × медіана цін по магазинах мережі; «Сільпо» — 4 позиції без EAN (3 Samyang, з них 2 за 89,99 — акція, і 1 Nongshim); 93 = 89 позицій з EAN + 4 «Сільпо». «Медіана» в підсумку — за медіанними цінами позицій по мережах. Сегменти 1–3.",7.04)

# ============ ФОРМАТ × ВАГА ============
s=slide(prs,"ФОРМАТ × ВАГА","Де насправді щільна полиця",num(),
 "Кожна клітинка — скільки позицій корейського рамену має цей формат у цій ваговій смузі. Насиченість кольору пропорційна кількості.")
txt(s,0.62,1.78,12.10,0.15,"МАТРИЦЯ ФОРМАТ × ВАГА · КІЛЬКІСТЬ SKU",8,True,NAVY)
y=matrix(s,0.62,2.02,[b['lab'] for b in F['bands']],["Пакет (пачка)","Стакан"],F['matrix'],cw=1.70,ch=0.54,rl=2.30)
WB={'40–90 г':'лише стакани: дорого за 100 г, дешево за чек','90–110 г':'перехідна смуга, мало позицій',
 '110–125 г':'найщільніша смуга ринку — Shin Ramyun, Ottogi, Paldo','125–145 г':'смуга преміальних лінійок і соусних SKU',
 '145 г +':'поодинокі позиції'}
y2=section_head(s,y+0.30,"ВАГОВІ СМУГИ РИНКУ","частка SKU, медіана полиці та ціна за 100 г")
cols=[("ВАГОВА СМУГА",0.62,2.10,L),("SKU",2.76,0.70,R),("ЧАСТКА",3.56,1.00,R),("МЕДІАНА",4.66,1.20,R),
      ("ГРН/100 Г",5.96,1.20,R),("ЩО ЦЕ ЗА СМУГА",7.32,5.40,L)]
table(s,y2,cols,[[b['lab'],str(b['n']),f"{b['share']} %",f"{b['pmed']} грн",str(b['g']),WB[b['lab']]] for b in F['bands']],
      head_size=7.4,row_size=9.6,rh=0.345)
ecall(s,0.62,6.46,5.92,0.92,"ВАГА НАШИХ SKU — РИНКОВА",
 f"112,5–131 г потрапляє у дві найщільніші смуги: {F['bands'][2]['share']+F['bands'][3]['share']} % усіх позицій ринку. За форматом і грамажем заперечень немає.",TEAL)
ecall(s,6.80,6.46,5.92,0.92,"ЦІНА ЗА 100 Г — ВИЩА ЗА ВІДПОВІДНІ СМУГИ",
 f"При полиці 120 грн наші SKU — 92–107 грн/100 г проти {F['bands'][2]['g']}–{F['bands'][3]['g']} у їхніх смугах: на 7–37 % вище. За поточною розрахунковою ціною 124–132 грн/100 г.",RED)

# ============ НАЙПОПУЛЯРНІШІ ФОРМАТИ ============
_cc=co.defaultdict(list)
for _r in SD_ROWS:
    _f="стакан" if _r['seg']==3 else "пакет"
    _cc[(_f,round(_r['w']))].append(_r)
_combos=sorted([(k,v) for k,v in _cc.items() if len(v)>=2],key=lambda x:-len(x[1]))
_cov=sum(len(v) for _,v in _combos)
s=slide(prs,"НАЙПОПУЛЯРНІШІ ФОРМАТИ","Топ поєднань формат + вага",num(),
 f"Поєднання з двома і більше позиціями. У {len(_combos)} таких клітинках лежать {_cov} позицій із {len(SD_ROWS)} — решта розсіяна по унікальних грамажах.")
_top=_combos[0][1]; _top_med=st.median([x['pmed'] for x in _top])
estat(s,0.62,1.76,"УПАКОВКА №1 НА РИНКУ","пакет 120 г",f"{len(_combos[0][1])} позицій, медіана {_top_med:.0f} грн, {st.median([x['per100'] for x in _top]):.0f} грн/100 г.\nNongshim, Ottogi, Paldo — 11 мереж.",DARKAMBER,3.88,1.28,24)
estat(s,4.72,1.76,"УПАКОВКА №2","пакет 130 г","9 позицій, медіана 114 грн, 87 грн/100 г.\nСюди лягає наша Carbonara 131 г.",TEAL,3.88,1.28,24)
estat(s,8.82,1.76,"НАЙДЕШЕВШИЙ ФОРМАТ","стакан 80 г","4 позиції, медіана 70 грн.\nSamyang Buldak у стакані.",PURPLE,3.88,1.28,24)
estat(s,0.62,3.18,"ВАГА ЧОТИРЬОХ НАШИХ SKU","113 г","Жодної позиції ринку в цій вазі.\nНайближча щільна клітинка — 120 г.",RED,3.88,1.28,24)
estat(s,4.72,3.18,"РІЗНИЦЯ ДО СТАНДАРТУ","113 проти 120 г","На 7 г легше за стандарт категорії.\nЗа однакової ціни це −6 % продукту.",RED,3.88,1.28,24)
estat(s,8.82,3.18,"ПОКРИТТЯ СІТКИ",f"{_cov} з {len(SD_ROWS)}","позицій лежать у повторюваних\nпоєднаннях формату й ваги.",NAVY,3.88,1.28,24)
y=section_head(s,4.64,"ТОП ПОЄДНАНЬ ФОРМАТ × ВАГА","хто їх тримає і чи є там наші SKU")
_CH={"пакет:120":"ні — наші 112,5–113 г","пакет:130":"так — Carbonara 131 г","пакет:140":"ні",
     "пакет:105":"ні","стакан:67":"ні","стакан:80":"ні",
     "пакет:110":"ні","стакан:68":"ні","стакан:62":"ні","пакет:122":"поруч — Gouda 123 г","пакет:137":"ні"}
cols=[("ФОРМАТ",0.62,1.70,L),("ВАГА",2.40,0.80,R),("SKU",3.30,0.60,R),("МЕРЕЖ",4.00,0.80,R),
      ("МЕДІАНА",4.90,1.10,R),("ГРН/100 Г",6.10,1.10,R),("ХТО ТРИМАЄ",7.34,3.30,L),("ЧИ Є ТУТ НАШІ SKU",10.70,2.02,L)]
rows=[]
for (f,w),v in _combos[:5]:
    br=co.Counter(x['tm'] for x in v).most_common(3)
    rows.append([f,f"{w} г",str(len(v)),str(len(set(c for x in v for c in x['chains']))),
        f"{st.median([x['pmed'] for x in v]):.0f} грн",f"{st.median([x['per100'] for x in v]):.0f}",
        " · ".join(f"{b} ({n})" for b,n in br),_CH.get(f"{f}:{w}","ні")])
table(s,y,cols,rows,head_size=7.4,row_size=9.6,rh=0.31)
estrip(s,6.74,"ГОЛОВНЕ ПРО УПАКОВКУ",
 f"Стандарт категорії — пакет 120 г. Чотири наші SKU на 7 г легші; при полиці 120 грн вони на {(120/_top_med-1)*100:.0f} % дорожчі за медіану цієї клітинки ({_top_med:.0f} грн), за поточної розрахункової 140,55 грн — на {(NOW_LOW/_top_med-1)*100:.0f} %.",RED,0.62)


NOWS={"33":M.forward_current("33"),"6":M.forward_current("6")}

# ============ ЦІНИ НА РІВНІ SKU ============
def _clean_name(r):
    t=r['title']; tm=r['tm']
    pat=re.compile(r"nong\s?shim" if tm=="Nongshim" else re.escape(tm).replace("'","['’]?"),re.I)
    m=pat.search(t); rest=t[m.end():] if m else t
    rest=re.sub(r"\s*\d+[.,]?\d*\s*(г|гр|g)\b\s*$","",rest.strip(),flags=re.I)
    rest=re.sub(r"швидкого приготування","",rest,flags=re.I)
    rest=re.sub(r"^\s*(локшина|рамен|рамьон)\s+","",rest,flags=re.I)
    rest=re.sub(r"\s+"," ",rest).strip(" ,.")
    tmd=tm.replace("Choi's","Choi’s")
    nm=f"{tmd} · {rest}" if rest else tmd
    return nm
def _fw(w): return (f"{w:g}".replace('.',',')+" г")
def _short(x,n=42): return x if len(x)<=n else x[:n-1].rstrip()+"…"
MKT=[]
for r in SD_ROWS:
    if r['seg'] in (1,2) and r.get('pmed'):
        MKT.append(dict(label=_short(_clean_name(r))+f" · {_fw(r['w'])}",price=r['pmed'],nch=r['nch'],ours=False))
NOWS={"33":M.forward_current("33"),"6":M.forward_current("6")}
OURS_NAMES=[(nm,wt) for nm,wt,_,_ in M.SKUS]
def ours_rows(prices):
    return [dict(label=f"Choi’s SIAS · {nm} · {wt}",price=p,nch=None,ours=True) for (nm,wt),p in zip(OURS_NAMES,prices)]
def ladder2(sl,l,top,w,rows,xmax,rowh,line=120):
    lab_w=2.95; bar_l=l+lab_w+0.05; bar_w=w-lab_w-0.05-0.55-0.38
    X=lambda v: bar_l+bar_w*v/xmax
    H=rowh*len(rows)
    _rect(sl,bar_l,top-0.04,0.006,H+0.08,HAIR)
    _rect(sl,X(line)-0.006,top-0.08,0.014,H+0.14,RED)
    txt(sl,X(line)-0.4,top-0.24,0.8,0.14,"120 грн",7,True,RED,PP_ALIGN.CENTER)
    y=top
    for r in rows:
        o=r['ours']
        if o: _rect(sl,l,y,w,rowh-0.02,AMBLIGHT)
        txt(sl,l+0.06,y+rowh/2-0.085,lab_w-0.06,0.17,r['label'],7.8,o,(DARKAMBER if o else NAVY))
        _rect(sl,bar_l,y+rowh/2-0.075,max(bar_w*r['price']/xmax,0.02),0.15,AMBBAR if o else NAVYBAR)
        txt(sl,X(r['price'])+0.07,y+rowh/2-0.085,0.60,0.17,f"{r['price']:.0f}" if not o else d2(r['price']).rstrip('0').rstrip(','),8.2,True,(DARKAMBER if o else NAVY))
        if r['nch']: txt(sl,l+w-0.38,y+rowh/2-0.08,0.38,0.16,f"{r['nch']} м.",7,False,GREY,PP_ALIGN.RIGHT)
        y+=rowh
def ladder_pair(sl,rows,xmax,top=1.98):
    rows=sorted(rows,key=lambda r:(r['price'],not r['ours']))
    half=math.ceil(len(rows)/2)
    while half<len(rows) and rows[half-1]['ours'] and rows[half]['ours'] and rows[half-1]['price']==rows[half]['price']: half+=1
    rowh=min(0.34,4.05/half)
    ladder2(sl,0.62,top,5.95,rows[:half],xmax,rowh); ladder2(sl,6.77,top,5.95,rows[half:],xmax,rowh)
    return top+rowh*half
# --- слайд А: навколо 120 ---
_lo,_hi=96,144
A_m=[r for r in MKT if _lo<=r['price']<=_hi]
A_rows=A_m+ours_rows([120.0]*6)
nA_low=sum(1 for r in MKT if r['price']<_lo); nA_hi=sum(1 for r in MKT if r['price']>_hi)
s=slide(prs,"ЦІНА НА ПОЛИЦІ · SKU-РІВЕНЬ · 1 З 2","Хто стоїть навколо 120 грн: конкретні позиції",num(),
 f"{len(A_m)} позицій конкурентів у коридорі {_lo}–{_hi} грн і шість наших SKU при ціні 120 грн. Пакетний рамен, ціна за одну пачку з ПДВ.")
yb=ladder_pair(s,A_rows,150)
estrip(s,max(yb+0.12,6.10),"ЩО ПОКАЗУЄ КОРИДОР",
 f"Конкурентів з ціною до 120 грн включно — {sum(1 for r in A_m if r['price']<=120)} з {len(A_m)}. Нижче за коридор ще {nA_low} позицій (від {min(r['price'] for r in MKT):.0f} грн), вище — {nA_hi}. Наші шість SKU при 120 грн стоять поруч з Nongshim Chapaghetti (120), Paldo Volcano Chicken (123) та Choi’s Карбон (125).",NAVY,0.70)
slide_tail_sources(s,"Джерело цін конкурентів: каталоги мереж 29.09.2026, медіана ціни позиції (EAN) по мережах, де вона є; «м.» — число мереж. Наші SKU: ціль 120 грн — полиця за фінмоделлю при закупівлі €0,54 (33 палети). Червона лінія — 120 грн.",7.04)
# --- слайд Б: вище 120 ---
B_m=[r for r in MKT if r['price']>=120]
f33=NOWS["33"]
B_rows=B_m+ours_rows([f['shelf'] for f in f33])
s=slide(prs,"ЦІНА НА ПОЛИЦІ · SKU-РІВЕНЬ · 2 З 2","Де ми стоїмо зараз: від 120 грн і наша полиця",num(),
 f"{len(B_m)} позицій конкурентів від 120 грн і шість SKU SIAS за поточною закупівлею (33 палети, маржа 30 %): {d2(f33[2]['shelf'])} грн — чотири SKU, {d2(f33[0]['shelf'])} грн — Gouda і Carbonara.")
yb=ladder_pair(s,B_rows,210)
_below=lambda x: sum(1 for r in MKT if r['price']<x)
estrip(s,max(yb+0.12,6.10),"ЩО ПОКАЗУЄ СПИСОК",
 f"При поточній закупівлі чотири SKU ({d2(f33[2]['shelf'])} грн) були б дорожчими за {_below(f33[2]['shelf'])} із {len(MKT)} пакетних позицій, Gouda і Carbonara ({d2(f33[0]['shelf'])} грн) — за {_below(f33[0]['shelf'])}. При 120 грн — за {_below(120.0)}.",RED,0.70)
slide_tail_sources(s,"Джерело цін конкурентів: каталоги мереж 29.09.2026, медіана ціни позиції по мережах. Наші SKU «зараз» = Self-Cost СС ÷ (1 − бонус 25 % − маржа 30 %) × націнка 1,40; курс 51,20 грн/€. Червона лінія — 120 грн.",7.04)

# ============ РИТЕЙЛ-АУДИТ · МЕРЕЖІ ============
SCALE={'auchan':'національна','novus':'національна','metro':'національна · C&C','zaraz':'національна'}
BRANDS_IN={'auchan':'Samyang · Nongshim · Paldo · Ottogi · O’Food · Bibigo','grono':'Samyang · Nongshim · Paldo',
 'onde':'Samyang · Nongshim · Ottogi','metro':'Samyang · Nongshim','zaraz':'Samyang · Nongshim · Ottogi',
 'vostorg':'Samyang · Nongshim','kharkiv':'Samyang · Nongshim','cosmos':'Samyang · Nongshim · Ottogi · Choi’s',
 'tavriav':'Samyang · Nongshim · Ottogi · Choi’s','novus':'Samyang · Nongshim','chudomarket':'Samyang · Nongshim','torba':'Samyang'}
s=slide(prs,"РИТЕЙЛ-АУДИТ · МЕРЕЖІ","Що з корейського рамену стоїть у мережах",num(),
 "Суцільна перевірка онлайн-каталогів 18 мереж, зріз 29.09.2026. SKU — унікальні позиції корейського рамену в каталозі мережі.")
cols=[("МЕРЕЖА",0.62,2.60,L),("ПОКРИТТЯ",3.22,1.80,L),("SKU",5.02,0.70,R),("МЕДІАНА",5.82,1.10,R),
      ("ДІАПАЗОН",7.02,1.50,R),("БРЕНДИ НА ПОЛИЦІ",8.70,4.02,L)]
rows=[]
for r in F['chain_rows']:
    rows.append([NM.get(r['chain'],r['chain']),SCALE.get(r['chain'],'регіональна'),str(r['n']),
                 f"{r['med']} грн",f"{r['lo']}–{r['hi']}",BRANDS_IN.get(r['chain'],'')])
rows.append(["«Сільпо»","національна · №1","4","114 грн","90–199","Samyang (3) · Nongshim (1)"])
rows.append(["МегаМаркет · Ultramarket · ЕКО","національна","0","—","—","категорії немає"])
rows.append(["Епіцентр · «Ідеал»","національна / регіон.","0","—","—","категорії немає"])
table(s,1.86,cols,rows,head_size=7.4,row_size=9.2,rh=0.282)
_cs={z['ean'] for z in ZREC if z['chain']=='cosmos'}; _ts={z['ean'] for z in ZREC if z['chain']=='tavriav'}
_n_cos=len(_cs); _n_same=len(_cs&_ts)
ecall(s,0.62,6.52,3.88,0.90,"ДІРКА №1 · «СІЛЬПО»",
 "Найбільша мережа країни тримає 4 позиції. Полиця відкрита, але вхід там 90–199 грн.",AMBER,9.2,8.4)
ecall(s,4.72,6.52,3.88,0.90,"ДІРКА №2 · П’ЯТЬ МЕРЕЖ З НУЛЕМ",
 "МегаМаркет, Ultramarket, ЕКО, Епіцентр, «Ідеал» — жодної позиції у зрізі.",AMBER,9.2,8.4)
ecall(s,8.82,6.52,3.90,0.90,"ОДИН ПОСТАЧАЛЬНИК НА ДВІ МЕРЕЖІ",
 f"«Космос» і «Таврія В» — по {_n_cos} позицій; {_n_same} із них збігаються за EAN. Найімовірніше, один постачальник на дві мережі.",NAVY,9.2,8.4)


# ============ КАНАЛИ ПРОДАЖУ ============
s=slide(prs,"КАНАЛИ ПРОДАЖУ","Де продається категорія і кому",num(),
 "Ліворуч — скільки точок у кожному каналі перевірено у зрізі. Праворуч — хто туди приходить і чи це наш покупець.")
txt(s,0.62,1.80,5.60,0.14,"ТОЧОК ПРОДАЖУ В КАНАЛІ · ПЕРЕВІРЕНО У ЗРІЗІ",8,True,NAVY)
hbars(s,0.62,2.12,5.90,[("Продуктові мережі",18,NAVYBAR,"18  ·  13 мають рамен"),("Азійські фудшопи",5,TEALBAR,"5"),("Маркетплейси",4,AMBBAR,"4")],
      20,[0,5,10,15,20],lab_w=1.95,val_w=1.70,rowh=0.46,bar_h=0.22)
ecall(s,0.62,4.10,5.90,2.70,"КАНАЛ І ПОКУПЕЦЬ — РІЗНІ ОСІ",
 "Продуктові мережі — єдиний канал з обсягом, але саме там уже стоять Samyang і Nongshim.\n\n"
 "Маркетплейси й азійські фудшопи дають орієнтир ціни, але не обсяг: медіана на Prom — 135 грн, у фудшопах 61–158 грн. Умови бонусу в цих каналах не перевірялися: у моделі закладено 25 % для мережевого роздробу.\n\n"
 "На відміну від готового рису, outdoor- і мілітарі-рітейл для рамену не канал: рамен потребує окропу, тому в туристичних магазинах його немає.\n\n"
 "Практичний вхід — «Сільпо» і п’ять нацмереж із нулем.",NAVY,10,8.8)
def chcard(t2,tag,tagcol,head,sub,who,body):
    _rect(s,6.80,t2,5.92,1.26,LIGHT); _rect(s,6.80,t2,0.045,1.26,tagcol)
    txt(s,7.06,t2+0.10,3.60,0.20,head,11,True,NAVY)
    _rect(s,10.98,t2+0.10,1.64,0.22,tagcol); txt(s,10.98,t2+0.13,1.64,0.16,tag,6.2,True,WHITE,PP_ALIGN.CENTER)
    txt(s,7.06,t2+0.36,5.42,0.13,sub,6.8,True,tagcol)
    txt(s,7.06,t2+0.54,5.42,0.44,body,7.6,False,NAVY,lh=1.25)
    txt(s,7.06,t2+1.02,5.42,0.13,"ПОКУПЕЦЬ:  "+who,6.8,True,GREY)
chcard(1.80,"ЦІЛЬОВИЙ КАНАЛ",RED,"Продуктовий роздріб","МЕРЕЖІ НАЦІОНАЛЬНОГО ПОКРИТТЯ · 18 ПЕРЕВІРЕНО","Масовий покупець · щоденне харчування",
 "13 мереж мають корейський рамен, 5 — жодної позиції. «Сільпо» тримає 4 позиції на всю мережу.")
chcard(3.18,"ЦІЛЬОВИЙ КАНАЛ",DARKAMBER,"Онлайн-маркетплейси","УНІВЕРСАЛЬНІ МАЙДАНЧИКИ · 4 МАЙДАНЧИКИ","Масовий покупець · планована закупівля",
 "Prom · Rozetka · MAUDAU · Bigl. Медіана в категорії «локшина рамен» на Prom — 135 грн, мінімум 80 грн.")
chcard(4.56,"ВХІД",TEAL,"Азійські фудшопи","СПЕЦІАЛІЗОВАНИЙ ФУДРІТЕЙЛ · 5 МАГАЗИНІВ","Покупець азійської кухні · знає категорію",
 "OMG! Asia · Тайякі Март · Суші Повар · Asia Foods · Panda Asian. Корейський рамен 61–158 грн, ядро 92–135 грн.")
chcard(5.94,"НЕ НАШ КАНАЛ",PALE,"Outdoor- і мілітарі-рітейл","ТУРИСТИЧНІ ТА ВІЙСЬКОВІ МАГАЗИНИ · 0 ПОЗИЦІЙ","Турист і військовий · автономне харчування",
 "Рамену там немає: він потребує окропу й каструлі.")

# ============ СУМІЖНА ПОЛИЦЯ ============
s=slide(prs,"СУМІЖНА ПОЛИЦЯ","Не-корейська локшина: підлога, від якої рахує покупець",num(),
 "Чотирнадцять найпоширеніших не-корейських позицій локшини швидкого приготування — по одній на бренд. Вони формують уявлення про «нормальну» ціну в категорії.")
band(s,1.74,"УКРАЇНА ТА АЗІЯ БЕЗ КОРЕЇ   ·   Варіння або окріп   ·   50–200 г   ·   9–144 грн   ·   18–146 грн/100 г",BLUE)
cards(s,CM,BLUE)
ecall(s,0.62,6.08,5.92,1.26,"ЩО ЦЕ ОЗНАЧАЄ ДЛЯ ЦІНИ",
 "Український мас-маркет — 9–22 грн за 50–60 г, тобто 18–33 грн/100 г. Корейський рамен коштує у 2,5–3 рази дорожче за 100 г, і ринок це приймає. Але прийнята премія має межу.",BLUE)
ecall(s,6.80,6.08,5.92,1.26,"ДЕ ЦЯ МЕЖА ПРОХОДИТЬ",
 "У корейському рамені в пакеті ціна за 100 г — від 49 до 158 грн, медіана 80. При полиці 120 грн наші SKU — 92–107 грн/100 г; на рівні 119 грн/100 г і вище стоять лише 6 із 71 пакетних позицій.",RED)

# ============ CHOI'S ============
CH_NEED=M.max_cc_uah(125.2)
CH_P33=M.max_price_eur("33",shelf=125.2); CH_P6=M.max_price_eur("6",shelf=125.2)
s=slide(prs,"ОКРЕМИЙ КОНКУРЕНТ","Choi’s: ринковий орієнтир, окремий продукт",num(),
 "Choi’s Carbonara 131 г — позиція вихідного зрізу 29.09.2026. Це окремий конкурент, а не наш товар SIAS.")
ch=g("Choi's")
if ch: card(s,0.62,1.86,ch[0].get("img"),ch[0]["name"],ch[0]["w"],ch[0]["price"],ch[0]["sellers"],RED)
estat(s,2.50,1.86,"РОЗДРІБНА ЦІНА НА ПОЛИЦІ","125,20 грн","«Таврія В» і «Космос» на 29.09.\nЗа V3 (05–06.10) у «Таврії В» — 131,50 грн.",NAVY,3.26,1.86,25)
estat(s,5.98,1.86,"НАШІ SIAS · GOUDA / CARBONARA","120 грн","Ціль за моделлю: на 4 % нижче за Choi’s 125,20.\nЗараз за моделлю — 162,18 грн (+30 %).",RED,3.40,1.86,25)
estat(s,9.60,1.86,"СС ДЛЯ ПАРИТЕТУ 125,20",f"{d2(CH_NEED)} грн",f"Закупівля SIAS: {eur(CH_P33)} (машина), {eur(CH_P6)} (6 палет) — при маржі 30 %.",DARKAMBER,3.12,1.86,25)
y=section_head(s,4.06,"ЩО З ЦЬОГО МОЖНА І ЧОГО НЕ МОЖНА ВИВЕСТИ","")
ecall(s,0.62,y,3.88,2.30,"1 · ЦЕ ПОРІВНЯННЯ РІЗНИХ ПРОДУКТІВ",
 "Choi’s Carbonara коштує 125,20–131,50 грн. Наші SIAS Gouda / Carbonara за поточною собівартістю — 162,18 грн за моделлю, тобто на 30 % дорожче.\n\n"
 f"Щоб стати на 125,20 грн з маржею 30 %, потрібна СС {d2(CH_NEED)} грн. Це ціновий сценарій, а не ціна закупівлі конкурента.",AMBER)
ecall(s,4.72,y,3.88,2.30,"2 · ЩО ПЕРЕВІРИТИ ДЛЯ SIAS",
 "Для наших шести SKU потрібно підтвердити:\n\n· склад та інструкції приготування;\n· що саме входить у собівартість — товар і всі витрати;\n· умови поставки й канал продажу;\n· підтримку запуску і попит за нашою ціною.",NAVY)
ecall(s,8.82,y,3.90,2.30,"3 · МЕЖА ВИСНОВКУ",
 "Ціна Choi’s не розкриває його собівартості, бонусу мережі чи маржі.\n\nЗ неї не можна вивести закупівельну ціну SIAS, спільного імпортера або доступність ексклюзиву для нашого продукту.",RED)
slide_tail_sources(s,"Джерело: каталоги tavriav.zakaz.ua і cosmos.zakaz.ua, «Локшина Choi’s рамен карбон 131 г»; 05–06.10 — за V3. Паритет — власний розрахунок за формулою вкладок 3–4.",6.94)

# ============ CHOI'S У «ТАВРІЇ В» · ЦІНА ПРОДАЖУ ============
EU_PROMO=1.80*M.RATE; EU_REG=2.20*M.RATE
_dearer=lambda x: sum(1 for p in PACK_P if p>x)
s=slide(prs,"ОКРЕМИЙ КОНКУРЕНТ · ЦІНА ПРОДАЖУ","Choi’s у «Таврії В»: по чому продавали",num(),
 "Той самий бренд SIAS уже стоїть на українській полиці. Три дати й три рівні цін — щоб порівняти з нашою ціллю та поточною моделлю.")
txt(s,0.62,1.80,6.2,0.16,"CHOI’S CARBONARA 131 Г · ЦІНА ЗА ПАЧКУ, ГРН",8.4,True,NAVY)
hbars(s,0.62,2.10,6.60,[
  ("ЄС · dotasia.eu · акція €1,80",EU_PROMO,GREYBAR,f"{d1(EU_PROMO)} грн"),
  ("ЄС · dotasia.eu · звичайна €2,20",EU_REG,GREYBAR,f"{d1(EU_REG)} грн"),
  ("«Таврія В» · 29.09.2026",125.2,NAVYBAR,"125,20 грн"),
  ("«Космос» · 29.09.2026",125.2,NAVYBAR,"125,20 грн"),
  ("«Таврія В» · 05–06.10 (за V3)",131.5,NAVYBAR,"131,50 грн"),
  ("НАШ SIAS · ціль за моделлю",120.0,AMBBAR,"120,00 грн"),
  ("НАШ SIAS · зараз, 33 палети",NOWS["33"][0]['shelf'],tint(AMBBAR,0.35),f"{d2(NOWS['33'][0]['shelf'])} грн"),
  ("НАШ SIAS · зараз, 6 палет",NOWS["6"][0]['shelf'],tint(AMBBAR,0.35),f"{d2(NOWS['6'][0]['shelf'])} грн"),
 ],220,[0,40,80,120,160,200],lab_w=2.80,val_w=1.10,rowh=0.40,bar_h=0.23)
txt(s,7.55,1.80,5.2,0.16,"ПОРІВНЯННЯ · «ЗАРАЗ» = 33 ПАЛЕТИ, МАРЖА 30 %",8.0,True,NAVY)
cols=[("",7.55,1.70,L),("CHOI’S",9.25,1.15,R),("ЦІЛЬ 120",10.40,1.15,R),("ЗАРАЗ",11.55,1.17,R)]
_c=NOWS["33"][0]['shelf']
tab=[["Ціна, грн","125,20 / 131,50","120,00",d2(_c)],
     ["За 100 г, грн","95,6 / 100,4","91,6",d1(_c/131*100)],
     ["До Choi’s 125,20",  "—", sgn((120/125.2-1)*100), sgn((_c/125.2-1)*100)],
     ["До медіани пакета 102,8",sgn((125.2/PACK_MED-1)*100),sgn((120/PACK_MED-1)*100),sgn((_c/PACK_MED-1)*100)],
     ["Дорожчих у категорії*",f"{_dearer(125.2)} з 71",f"{_dearer(120)} з 71",f"{_dearer(_c)} з 71"]]
table(s,2.10,cols,tab,head_size=7.0,row_size=9.0,rh=0.44,left=7.55,bandw=5.17)
ecall(s,7.55,4.80,5.17,2.06,"ЩО ЦЕ ЗНАЧИТЬ ДЛЯ ПЕРЕГОВОРІВ",
 "1. Наша ціль 120 грн — на 4–9 % нижче за те, по чому цей бренд уже продають у «Таврії В». Якщо це той самий EAN 3760344350847, діючий імпортер матиме ціну вищу за нашу.\n"
 "2. За 29.09 → 05–06.10 ціна в «Таврії В» зросла на 5 %, акційної ціни немає (стара ціна дорівнює новій).\n"
 "3. Історії цін немає — лише дві дати. Її варто запросити в «Таврії В» і SIAS.",RED,9.6,8.2)
ecall(s,0.62,5.74,6.60,1.14,"ЄВРОПЕЙСЬКИЙ ОРІЄНТИР — ТІЛЬКИ ДЛЯ ПОРІВНЯННЯ",
 "dotasia.eu продає ту саму позицію за €1,80 замість €2,20 (−18 %); Gouda Cheese 123 г — так само. Це інтернет-магазин в ЄС, ПДВ на сторінці не вказано, товар «немає в наявності» — прямо порівнювати з українською полицею не можна.",AMBER,9.4,8.2)
slide_tail_sources(s,"* З 71 пакетної позиції зрізу 29.09.2026. Курс для перерахунку EUR→грн — 51,20. Джерела: tavriav.zakaz.ua, cosmos.zakaz.ua; ціна «Таврії В» 05–06.10 — за V3; dotasia.eu (Choi’s Carbonara 131 г, Gouda 123 г).",7.00)

# ============ АКЦІЇ І МІНІМАЛЬНІ ЦІНИ ============
Z=json.load(open(f"{SC}/zakaz_raw.json",encoding='utf-8'))
_eans={r['ean']:r for r in SD_ROWS}
_rec=[z for z in Z if z['ean'] in _eans and _eans[z['ean']]['seg'] in (1,2)]
KEY=[('08801073113428','Samyang Buldak 2× spicy','140 г'),('08801073110502','Samyang Buldak Hot Chicken','140 г'),
     ('08801073113381','Samyang Buldak тушкована курка','145 г'),('08801073116474','Samyang Buldak з сиром','130 г'),
     ('08801043150620','Nongshim Shin Ramyun','120 г'),('08801043157742','Nongshim Kimchi Ramyun','120 г'),
     ('08801043157728','Nongshim Chapaghetti','140 г'),('08801043018470','Nongshim Toomba Spicy & Creamy','137 г'),
     ('00648436310685','Paldo Volcano Carbonara','130 г')]
krow=[]
for e,lab,wt in KEY:
    v=[z for z in _rec if z['ean']==e]
    krow.append((lab,wt,len({z['chain'] for z in v}),min(z['price'] for z in v),max(z['price'] for z in v)))
_ap=[z for z in _rec if z['chain']=='auchan']
_app=[z for z in _ap if z.get('old_price') and z['price']<z['old_price']-0.005]
s=slide(prs,"АКЦІЇ ТА МІНІМАЛЬНІ ЦІНИ","Ключові конкуренти продаються нижче 120 грн",num(),
 "Одна й та сама позиція конкурента коштує по-різному в різних мережах і під час акцій. Для SIAS це діапазон конкурентного тиску на полиці.")
cols=[("ПОЗИЦІЯ КОНКУРЕНТА",0.62,3.10,L),("ВАГА",3.72,0.80,R),("МЕРЕЖ",4.52,0.70,R),("НАЙНИЖЧА, грн",5.22,1.30,R),("НАЙВИЩА, грн",6.52,1.30,R),("РІЗНИЦЯ",7.82,0.90,R)]
table(s,1.86,cols,[[l,w,str(n),d2(lo),d2(hi),pct((1-lo/hi)*100,0)] for l,w,n,lo,hi in krow],head_size=7.4,row_size=9.6,rh=0.33,bandw=8.30)
estat(s,9.10,1.78,"НИЖЧІ ЗА 120 ГРН",f"{sum(1 for k in krow if k[3]<120)} з {len(krow)}","ключових позицій конкурентів\nмають мінімальну ціну нижче 120 грн.",RED,3.62,1.22,24)
estat(s,9.10,3.14,"РОЗКИД МІЖ МЕРЕЖАМИ","34 %","медіанний розкид найдешевшої й найдорожчої\nмережі для 18 позицій у трьох і більше мережах.",DARKAMBER,3.62,1.22,24)
estat(s,9.10,4.50,"АКЦІЯ В «АШАНІ»",f"{len(_app)} з {len(_ap)}","записів пакетного рамену на акції;\nзнижка найчастіше 20–23 %, максимум 39 %.",TEAL,3.62,1.22,24)
ecall(s,0.62,5.46,8.30,1.38,"СКІЛЬКИ КОШТУЄ АКЦІЯ",
 "«Сільпо»: Samyang Buldak зі смаком курки та сиру і Buldak Carbonara — 89,99 грн замість 174 грн (−48 %). Samyang Tangle: від 77,50 до 157,00 грн залежно від мережі (−51 %).\n"
 "Отже, конкурент може виставити ту саму пачку за 86–105 грн на акції; наша ціна 120 грн — вище за цей рівень, але нижче за повну ціну більшості позицій.",NAVY,9.6,8.6)
ecall(s,9.10,5.86,3.62,0.98,"ОБМЕЖЕННЯ ДАНИХ",
 "Стару ціну віддають лише деякі мережі (в основному «Ашан»): частка акцій тут — нижня межа.",AMBER,8.6,7.8)
slide_tail_sources(s,"Каталоги мереж 29.09.2026: ціна за пачку і стара ціна. «Різниця» = 1 − найнижча ÷ найвища ціна позиції по мережах. «Сільпо» — ціна й стара ціна в каталозі.",7.00)


# ============ ФІНАНСИ 1 · ПОЛИЦЯ ЗАРАЗ ============
F33=M.forward_current("33"); F6=M.forward_current("6")
_below=lambda x: sum(1 for p in PACK_P if p<x)
s=slide(prs,"ФІНМОДЕЛЬ · 1 З 2",f"Ціна на полиці зараз: {d0(F33[2]['shelf'])}–{d0(F6[0]['shelf'])} грн замість цільових 120",num(),
 "Шість SKU SIAS за поточною закупівлею €0,65–0,75 за пачку. Ціна на полиці — з ПДВ, за одну пачку, маржа 30 %, бонус мережі 25 %.")
flow(s,1.76,[("€0,65 · €0,75","закупівля SIAS зараз за пачку"),
             (f"{d2(F33[2]['cc'])} · {d2(F33[0]['cc'])} грн","СС (Self-Cost): товар, логістика, ПДВ · 33 палети"),
             (f"{d2(F33[2]['partner'])} · {d2(F33[0]['partner'])} грн","ціна партнеру = СС ÷ 0,45"),
             (f"{d2(F33[2]['shelf'])} · {d2(F33[0]['shelf'])} грн","полиця = ціна партнеру × 1,40")],bw=2.80,gap=0.30,bh=0.86)
cols=[("ПОЗИЦІЯ SIAS",0.70,2.60,L),("ВАГА",3.30,0.80,R),("ЗАКУПІВЛЯ, €/пач.",4.20,1.35,R),("СС, ГРН · 33 ПАЛ.",5.65,1.35,R),
      ("ПОЛИЦЯ · 33 ПАЛЕТИ",7.10,1.55,R),("ПОЛИЦЯ · 6 ПАЛЕТ",8.75,1.45,R),("ЦІЛЬ",10.30,0.70,R),("ВІДРИВ ВІД 120 · 33 ПАЛ.",11.05,1.62,R)]
rows=[]
for (nm,wt,w_,cur),a,b in zip(M.SKUS,F33,F6):
    rows.append([f"Choi’s SIAS · {nm}",wt,eur(cur),d2(a['cc']),d2(a['shelf']),d2(b['shelf']),"120,00",f"+{a['shelf']-120:.2f}".replace('.',',')+f" грн · +{(a['shelf']/120-1)*100:.0f} %"])
table(s,2.80,cols,rows,head_size=7.0,row_size=9.6,rh=0.35,left=0.62,bandw=12.10)
estat(s,0.62,5.22,"33 ПАЛЕТИ · ЧОТИРИ SKU",f"{d2(F33[2]['shelf'])} грн",f"+{(F33[2]['shelf']/120-1)*100:.0f} % до цілі 120. Дорожче за {_below(F33[2]['shelf'])} із {len(PACK_P)} пакетних позицій ринку.",TEAL,3.88,1.42,26)
estat(s,4.72,5.22,"33 ПАЛЕТИ · GOUDA, CARBONARA",f"{d2(F33[0]['shelf'])} грн",f"+{(F33[0]['shelf']/120-1)*100:.0f} % до цілі 120. Дорожче за {_below(F33[0]['shelf'])} із {len(PACK_P)} пакетних позицій.",DARKAMBER,3.88,1.42,26)
estat(s,8.82,5.22,"6 ПАЛЕТ · УСІ SKU",f"{d0(F6[2]['shelf'])}–{d0(F6[0]['shelf'])} грн",f"+{(F6[2]['shelf']/120-1)*100:.0f}…{(F6[0]['shelf']/120-1)*100:.0f} % до цілі 120. Вище за весь ринок пакетів (макс. {max(PACK_P):.0f} грн).",RED,3.90,1.42,24)
slide_tail_sources(s,"Фінмодель SIAS: вкладки «Финмодель 6 паллет» і «Финмодель 33 паллет», колонки маржі 30 %. СС — Self-Cost_Sias.xlsx. Курс 51,20 грн/€; ціна партнеру — вже з ПДВ. Коефіцієнти: 1 − 0,25 − 0,30 = 0,45; націнка 1,40.",7.00)

# ============ ФІНАНСИ 2 · ПОТРІБНА ЗАКУПІВЛЯ ============
x120=120/M.MARKUP; ccmax=x120*(1-M.BONUS-M.MARGIN); e1=ccmax/M.RATE
A33=M.ask_price("33"); A6=M.ask_price("6")
s=slide(prs,"ФІНМОДЕЛЬ · 2 З 2","Яка закупівля потрібна, щоб стати на полицю за 120 грн",num(),
 "Обернений розрахунок за тією ж моделлю: від ціни на полиці 120 грн униз до ціни, яку ми можемо заплатити SIAS за пачку.")
flow(s,1.76,[("120,00 грн","цільова ціна на полиці з ПДВ"),
             (f"{d2(x120)} грн","ціна партнеру = 120 ÷ 1,40"),
             (f"{d2(ccmax)} грн","макс. СС = ціна партнеру × 45 %"),
             (f"€{e1:.3f}".replace('.',','),"макс. СС в євро = грн ÷ 51,20"),
             (f"{eur(A33)} · {eur(A6)}","закупівля: 33 палети · 6 палет")],bw=2.0,gap=0.525,bh=0.86)
cols=[("ПОЗИЦІЯ SIAS",0.70,2.50,L),("ВАГА",3.20,0.80,R),("ЗАРАЗ, €/пач.",4.05,1.15,R),("ПОТРІБНО · 33 ПАЛЕТИ",5.25,1.70,R),("ЗНИЖЕННЯ",7.00,1.00,R),
      ("ПОТРІБНО · 6 ПАЛЕТ",8.10,1.65,R),("ЗНИЖЕННЯ",9.80,1.00,R)]
rows=[[f"Choi’s SIAS · {nm}",wt,eur(cur),eur(A33),"−"+pct((1-A33/cur)*100,0),eur(A6),"−"+pct((1-A6/cur)*100,0)] for nm,wt,w_,cur in M.SKUS]
yb=table(s,2.80,cols,rows,head_size=7.0,row_size=9.6,rh=0.32,left=0.62,bandw=10.40)
estat(s,11.00,2.78,"33 ПАЛЕТИ",f"−{(1-A33/0.65)*100:.0f}…{(1-A33/0.75)*100:.0f} %","",TEAL,1.72,1.00,14)
estat(s,11.00,3.88,"6 ПАЛЕТ",f"−{(1-A6/0.65)*100:.0f}…{(1-A6/0.75)*100:.0f} %","",RED,1.72,1.00,14)
_rect(s,0.62,5.02,6.50,1.88,LIGHT); _rect(s,0.62,5.02,0.045,1.88,AMBER)
txt(s,0.92,5.12,6.0,0.18,"ЯКЩО ЗАКУПІВЛЮ НЕ ЗНИЗИТИ: ЩО ВИТРИМАЄ ПОЛИЦЯ 120 ГРН",9,True,NAVY)
mg=[0.30,0.25,0.20,0.15]
mp=lambda sc,m: math.floor(M.max_price_eur(sc,shelf=120,margin=m)*100+1e-9)/100
table(s,5.40,[("НАША МАРЖА",0.92,1.50,L),("МАКС. ЗАКУПІВЛЯ · 33 ПАЛ.",2.45,2.25,R),("МАКС. ЗАКУПІВЛЯ · 6 ПАЛ.",4.75,2.20,R)],
      [[pct(m*100,0),eur(mp("33",m)),eur(mp("6",m))] for m in mg],head_size=6.8,row_size=9.4,rh=0.27,left=0.84,bandw=6.20)
_m33=[M.margin_at_shelf(F33[0]['cc'])*100,M.margin_at_shelf(F33[2]['cc'])*100]; _m6=[M.margin_at_shelf(F6[0]['cc'])*100,M.margin_at_shelf(F6[2]['cc'])*100]
txt(s,0.92,6.64,6.10,0.18,f"При поточній закупівлі полиця 120 грн дає маржу {_m33[0]:.0f}–{_m33[1]:.0f} % (33 палети) і {_m6[0]:.0f}…{_m6[1]:.0f} % (6 палет).".replace('-','−'),7.8,False,GREY)
ecall(s,7.30,5.02,5.42,1.88,"УМОВИ РОЗРАХУНКУ",
 "Бонус мережі 25 %, наша маржа 30 %, націнка магазину 1,40, курс 51,20 грн/€, мито 0 %, ПДВ 20 %. Витрати партії (фрахт, брокер) — на пачку: 33 палети €0,09, 6 палет €0,34.\n"
 "Не враховано: лістинг-фі, промо-підтримка мережі, відстрочка. Згоди SIAS на ці ціни немає — це ціна для переговорів.",NAVY,9.4,8.2)
slide_tail_sources(s,"Фінмодель SIAS: вкладки «Цена 120 грн — 6 паллет» і «— 33 паллет». Запит — гранична ціна, округлена вниз до євроцента. «Зараз» — з Self-Cost_Sias (€0,65 — чотири SKU, €0,75 — Gouda і Carbonara).",7.12)

out="/home/user/Chaplygin-Illya/SIAS_MARKET_RESEARCH_UA_V5.pptx"
prs.save(out); print("saved",out,"slides:",len(prs.slides._sldIdLst))
