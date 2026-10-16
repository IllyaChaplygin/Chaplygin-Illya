# -*- coding: utf-8 -*-
"""V10: правки поверх файлу користувача (22af1e38…UA_.pptx): нова структура ціни (бонус 35 %, маржа 20 %), слайди 27–29, легенда слайда 23."""
import sys,os,re,math
sys.path.insert(0,"/tmp/claude-0/-home-user-Chaplygin-Illya/8bc0034a-4b53-5022-9bf5-e5697c987387/scratchpad/v5")
import kit2
from kit2 import *
from pptx import Presentation
from pptx.util import Inches,Emu
from deckkit import _rect
import openpyxl as _ox
XLSX=os.environ.get("XLSX","/home/user/Chaplygin-Illya/RAMEN_FINMODEL_INTERNAL_UA.xlsx")
UP="/root/.claude/uploads/8bc0034a-4b53-5022-9bf5-e5697c987387/22af1e38-SIAS_MARKET_RESEARCH_UA_.pptx"
OUT=os.environ.get("OUT","/home/user/Chaplygin-Illya/SIAS_MARKET_RESEARCH_UA_V12.pptx")
prs=Presentation(UP)
OLD=[sl.slide_id for sl in prs.slides]
def rr(s,l,t,w,h,fill,rad=0.08):
    sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(l),Inches(t),Inches(w),Inches(h))
    sh.line.fill.background(); sh.shadow.inherit=False; sh.fill.solid(); sh.fill.fore_color.rgb=fill
    sh.adjustments[0]=rad; return sh
def newslide(kicker,title,sub=None): return slide(prs,kicker,title,"00",sub)
_XL=_ox.load_workbook(XLSX,data_only=True)
def pc(x): return f"{x*100:.0f} %"
def _fmt2(x): return f"{x:,.2f}".replace(',',' ').replace('.',',')
def _eur(x,n=3): return "€"+f"{x:.{n}f}".replace('.',',')
KEYT=RGBColor(0xE7,0xED,0xF5)
C_CC=RGBColor(0x46,0x53,0x82); C_BON=DARKAMBER
def final_slide(sheet,pal,col,askrow):
    w=_XL[sheet]; g=lambda a: w[a].value
    ask=g(f'B{askrow}'); cur4=g('B5'); cur2=g('B7'); bon=g('D14'); mar=g('C14'); X=g('G14') or 0
    ccaf=g('C8'); n7,n8,n9=g('H4'),g('H6'),g('H8')
    share=1-bon-mar
    s=newslide(f"ФІНМОДЕЛЬ · {pal} {'ПАЛЕТ' if pal==6 else 'ПАЛЕТИ'} · ФІНАЛЬНІ ЦИФРИ",f"{pal} {'палет' if pal==6 else 'палети'}: закупівля €{cur4:.2f}–{cur2:.2f} → {_eur(ask,2)}, полиця 120 ₴".replace('.',','))
    tiles=[("ЗАРАЗ · ПОЛИЦЯ",f"{_fmt2(n7)} / {_fmt2(n8)} ₴",RED,20,f"закупівля €{cur4:.2f} / €{cur2:.2f} · +{(n7/120-1)*100:.0f}…{(n8/120-1)*100:.0f} % до цілі 120 ₴".replace('.',',')),
           ("ЗАПИТ ДО SIAS",_eur(ask,2),col,22,f"−{(1-ask/cur4)*100:.0f} % / −{(1-ask/cur2)*100:.0f} % від теперішньої закупівлі (одна ціна для 6 SKU)"),
           ("ПІСЛЯ · ПОЛИЦЯ",f"{_fmt2(n9)} ₴",TEAL,20,f"бонус мережі {pc(bon)} · СС {_fmt2(ccaf)} ₴ · запас {_fmt2(120-n9)} ₴")]
    for i,(lab,big,acc,sz,sub) in enumerate(tiles):
        l=0.62+i*4.10
        _rect(s,l,1.42,3.90,1.02,LIGHT); _rect(s,l,1.42,0.05,1.02,acc)
        txt(s,l+0.25,1.50,3.5,0.16,lab,8,True,GREY)
        txt(s,l+0.25,1.70,3.5,0.4,big,sz,True,acc)
        txt(s,l+0.25,2.15,3.55,0.26,sub,8,False,GREY,lh=1.15)
    heads=[("Закупівля SIAS","ціна пачки, €",TEAL),("Собівартість","повна, в Україні" if X==0 else f"повна + {_fmt2(X)} ₴ витрат",C_CC),("Ціна партнеру","від собівартості",NAVY),("Бонус мережі",f"{pc(bon)} ціни партнеру",C_BON),("Залишається нам","витрати та маржа",DARKAMBER),("Націнка","магазину · +40 %",RGBColor(0x7C,0x86,0xA8)),("ПОЛИЦЯ з ПДВ","партнеру × 1,40",RED)]
    X0=3.45; CW=1.28; GP=0.04; HY=2.62
    for i,(h1,h2,cc) in enumerate(heads):
        x=X0+i*(CW+GP); _rect(s,x,HY,CW,0.56,cc)
        txt(s,x+0.04,HY+0.10,CW-0.08,0.2,h1,9,True,WHITE,C); txt(s,x+0.04,HY+0.33,CW-0.08,0.16,h2,7.4,False,tint(cc,0.7) if cc!=NAVY else PALE,C)
    cols="BCDEFGH"
    rows=[(4,"ДО · 4 SKU",f"Vegetable, Spicy, Beef, Chicken · закупівля €{cur4:.2f}".replace('.',','),None),
          (6,"ДО · 2 SKU",f"Gouda, Carbonara · закупівля €{cur2:.2f}".replace('.',','),None),
          (8,"ПІСЛЯ · запит SIAS",f"усі 6 SKU · закупівля {_eur(ask,2)}",AMBLIGHT)]
    y=HY+0.62; RH=0.60
    for r,l1,l2,bg in rows:
        _rect(s,0.62,y,2.78,RH-0.04,bg or LIGHT)
        txt(s,0.72,y+0.08,2.6,0.2,l1,10,True,NAVY if bg else INK); txt(s,0.72,y+0.33,2.62,0.18,l2,7.4,False,GREY)
        for i,c_ in enumerate(cols):
            x=X0+i*(CW+GP); last=(i==6)
            _rect(s,x,y,CW,RH-0.04,bg or (KEYT if last else WHITE))
            if not bg and not last:
                _rect(s,x,y,CW,0.01,HAIR); _rect(s,x,y+RH-0.05,CW,0.01,HAIR)
            txt(s,x,y+0.06,CW,0.26,_fmt2(g(f"{c_}{r}"))+" ₴",12.5 if not last else 14,True,RED if (last and r!=8) else (TEAL if (last and r==8) else NAVY),C)
            txt(s,x,y+0.34,CW,0.16,_eur(g(f"{c_}{r+1}")),8.4,False,GREY,C)
        y+=RH
    _rect(s,0.62,y,2.78,0.40,RGBColor(0xF8,0xE9,0xE9)); txt(s,0.72,y+0.11,2.6,0.2,"ЗМІНА · ПІСЛЯ − ДО (4 SKU)",8.4,True,RED)
    for i,c_ in enumerate(cols):
        x=X0+i*(CW+GP); _rect(s,x,y,CW,0.40,RGBColor(0xF8,0xE9,0xE9))
        txt(s,x,y+0.04,CW,0.18,f"{g(f'{c_}10'):+,.2f}".replace(',',' ').replace('.',',').replace('-','−')+" ₴",9.4,True,RED,C)
        d=g(f'{c_}11'); txt(s,x,y+0.22,CW,0.15,("+" if d>0 else "−")+_eur(abs(d)),7.4,False,GREY,C)
    y+=0.52
    txt(s,0.62,y,6,0.16,"ЗАКУПІВЛЯ SIAS ПО ПОЗИЦІЯХ · ЗАРАЗ → ЗАПИТ",8.2,True,NAVY)
    y+=0.24
    ORDER=[2,3,4,5,0,1]
    for k,idx in enumerate(ORDER):
        nm,wt,_,cur=M.SKUS[idx]; l=0.62+k*2.025
        _rect(s,l,y,1.93,0.80,LIGHT); _rect(s,l,y,1.93,0.04,col)
        txt(s,l+0.10,y+0.10,1.75,0.18,nm,8.8,True,NAVY); txt(s,l+0.10,y+0.28,1.75,0.14,wt,7.4,False,GREY)
        txt(s,l+0.10,y+0.46,1.2,0.22,f"€{cur:.2f} → €{ask:.2f}".replace('.',','),10.5,True,INK)
        rr(s,l+1.36,y+0.48,0.50,0.22,col,0.3); txt(s,l+1.36,y+0.51,0.50,0.16,f"−{(1-ask/cur)*100:.0f} %",8.2,True,WHITE,C)
    extra="" if X==0 else f" До собівартості додано {_fmt2(X)} ₴ витрат на пачку."
    txt(s,0.62,7.05,12.1,0.3,f"Курс 51,20 ₴/€. Ціна партнеру = СС ÷ (1 − бонус мережі {pc(bon)} − наша частка {pc(mar)}), полиця = ціна партнеру × 1,40; усі ціни з ПДВ, за пачку.{extra} Собівартість «до» — з Self-Cost; «після» — за запитом {_eur(ask,2)}. Джерело: RAMEN_FINMODEL_4_TABS, вкладка «Цена 120 грн — {'6' if pal==6 else '33'} паллет». Згоди SIAS немає — ціна для переговорів.",7,False,PALE,lh=1.2)
    return s
f33=final_slide('Цена 120 грн — 33 паллет',33,TEAL,19)
f6=final_slide('Цена 120 грн — 6 паллет',6,RED,20)
BON=_XL['Цена 120 грн — 33 паллет']['D14'].value; MAR=_XL['Цена 120 грн — 33 паллет']['C14'].value; XX=_XL['Цена 120 грн — 33 паллет']['G14'].value or 0
# ============ ЧУТЛИВІСТЬ (бонус з моделі, маржа 30…10 %) ============
SHELVES=[105,110,120,130,140]; MARGINS=[0.30,0.25,0.20,0.15,0.10]
def gridsc(sc): return [[math.floor(M.max_price_eur(sc,shelf=sh,margin=m,bonus=BON)*100+1e-9)/100 for sh in SHELVES] for m in MARGINS]
def cellcol(v):
    if v>=0.75: return (TEAL,WHITE)
    if v>=0.65: return (tint(TEAL,0.55),NAVY)
    if v>=0.50: return (tint(AMBER,0.55),NAVY)
    return (tint(RED,0.55),NAVY)
s=newslide("ЧУТЛИВІСТЬ","Яка закупівля потрібна при іншій полиці й іншій нашій частці",
 f"Максимальна ціна закупівлі, €/пачку, при бонусі мережі {pc(BON)}. Колір — порівняння з поточною ціною SIAS: €0,75 (Gouda, Carbonara) і €0,65 (решта).")
def sgrid(l,sc,head,acc):
    G_=gridsc(sc)
    txt(s,l,1.78,5.9,0.18,head,9.6,True,acc)
    cw=0.96; rl=1.00; t=2.12
    txt(s,l,t,rl,0.16,"НАМ \\ ПОЛИЦЯ",6.4,True,GREY)
    for j,sh in enumerate(SHELVES):
        txt(s,l+rl+j*cw,t-0.02,cw,0.20,f"{sh} грн",8.6,True,(RED if sh==120 else NAVY),PP_ALIGN.CENTER)
    y=t+0.26
    for i,m in enumerate(MARGINS):
        txt(s,l,y+0.12,rl-0.1,0.22,f"{int(round(m*100))} %",10,True,(RED if abs(m-MAR)<1e-9 else NAVY),PP_ALIGN.RIGHT)
        for j,sh in enumerate(SHELVES):
            v=G_[i][j]; bg,tc=cellcol(v)
            _rect(s,l+rl+j*cw+0.03,y+0.03,cw-0.06,0.50,bg)
            txt(s,l+rl+j*cw+0.03,y+0.16,cw-0.06,0.26,f"€{v:.2f}".replace('.',','),12,True,tc,PP_ALIGN.CENTER)
            if sh==120 and abs(m-MAR)<1e-9:
                fx=l+rl+j*cw+0.03; fy=y+0.03; fw=cw-0.06; fh=0.50
                for (a_,b_,c_,d_) in ((fx,fy,fw,0.04),(fx,fy+fh-0.04,fw,0.04),(fx,fy,0.04,fh),(fx+fw-0.04,fy,0.04,fh)): _rect(s,a_,b_,c_,d_,RED)
        y+=0.56
    return y
sgrid(0.62,"33","33 ПАЛЕТИ · ПОВНА МАШИНА",TEAL)
sgrid(6.80,"6","6 ПАЛЕТ",RED)
y0=5.16; lx=0.76
for lab,col in (("≥ €0,75 — досяжно для всіх SKU",TEAL),("€0,65–0,75 — для чотирьох SKU",tint(TEAL,0.55)),("€0,50–0,65 — потрібне зниження",tint(AMBER,0.55)),("< €0,50 — суттєве зниження",tint(RED,0.55))):
    _rect(s,lx,y0,0.18,0.14,col); txt(s,lx+0.24,y0-0.02,2.6,0.16,lab,7.8,False,NAVY); lx+=2.98
txt(s,0.76,y0+0.20,11.8,0.16,f"Червона рамка — ціль: полиця 120 грн, бонус мережі {pc(BON)}, частка «залишається нам» {pc(MAR)}.",7.8,True,RED)
slide_tail_sources(s,f"Розрахунок за формулою вкладок «Цена 120 грн» (курс 51,20; бонус {pc(BON)}; націнка 1,40; фінансування 2,5 %; ПДВ 20 %). Ціна вниз до євроцента.",7.06)
sens=s

# ============ З ЧОГО СКЛАДАЄТЬСЯ СОБІВАРТІСТЬ: БЛОКИ ============
import json as _json
BL=_json.load(open("/tmp/claude-0/-home-user-Chaplygin-Illya/8bc0034a-4b53-5022-9bf5-e5697c987387/scratchpad/v10/blocks.json"))
def cost_blocks_slide():
    s=newslide("СОБІВАРТІСТЬ · БЛОКИ","З чого складається собівартість в Україні","₴ на пачку: закупівля у SIAS + доставка + імпортний ПДВ + фінансування = собівартість")
    heads=[("Закупівля у SIAS","ціна пачки × курс",TEAL),("Доставка","фрахт, брокер, термінал",C_CC),("Імпортний ПДВ","20 % з товару й фрахту",DARKAMBER),("Фінансування","2,5 % від суми",RGBColor(0x5B,0x4B,0x9A)),("Собівартість","повна, в Україні",NAVY)]
    X0=3.20; CW=1.62; GP=0.22
    def panel(pal,title,y0,col,cur4,cur2,ask):
        d=BL[pal]
        txt(s,0.62,y0,6,0.18,title,9.6,True,col)
        hy=y0+0.26
        for i,(h1,h2,cc) in enumerate(heads):
            x=X0+i*(CW+GP); _rect(s,x,hy,CW,0.48,cc)
            txt(s,x+0.04,hy+0.06,CW-0.08,0.2,h1,9,True,WHITE,C); txt(s,x+0.04,hy+0.28,CW-0.08,0.16,h2,7.2,False,tint(cc,0.7) if cc!=NAVY else PALE,C)
            if i<4: txt(s,x+CW,hy+0.12,GP,0.26,"+" if i<3 else "=",14,True,GREY,C)
        rows=[("ДО · 4 SKU","Vegetable, Spicy, Beef, Chicken",d['before4'],None),
              ("ДО · 2 SKU","Gouda, Carbonara",d['before2'],None),
              ("ПІСЛЯ · запит SIAS",f"усі 6 SKU · закупівля {_eur(ask,2)}",d['after'],AMBLIGHT)]
        y=hy+0.54; RH=0.52
        for l1,l2,vals,bg in rows:
            _rect(s,0.62,y,2.46,RH-0.04,bg or LIGHT)
            txt(s,0.72,y+0.06,2.3,0.2,l1,10,True,NAVY if bg else INK); txt(s,0.72,y+0.29,2.34,0.18,l2,7.2,False,GREY)
            for i,v in enumerate(vals):
                x=X0+i*(CW+GP); last=(i==4)
                _rect(s,x,y,CW,RH-0.04,bg or (KEYT if last else WHITE))
                if not bg and not last: _rect(s,x,y,CW,0.01,HAIR); _rect(s,x,y+RH-0.05,CW,0.01,HAIR)
                txt(s,x,y+0.04,CW,0.26,_fmt2(v)+" ₴",13 if last else 12,True,(TEAL if (last and bg) else NAVY),C)
                txt(s,x,y+0.31,CW,0.16,_eur(v/51.2),8.2,False,GREY,C)
            y+=RH
        return y
    a=_XL['Цена 120 грн — 33 паллет']['B19'].value; b=_XL['Цена 120 грн — 6 паллет']['B20'].value
    y=panel("33","33 ПАЛЕТИ · ПОВНА МАШИНА",1.72,TEAL,0.65,0.75,a)
    panel("6","6 ПАЛЕТ",y+0.16,RED,0.65,0.75,b)
    txt(s,0.62,7.0,12.1,0.4,"Доставка — фрахт, брокер, термінал і сертифікація партії, у розрахунку на пачку; чим менша партія, тим дорожча доставка. Мито 0 %. ПДВ 20 % рахується з вартості товару й 70 % фрахту. Курс 51,20 ₴/€. «До» — за Self-Cost (логістика розподілена за вартістю), «після» — рівномірно на пачку.",7,False,PALE,lh=1.2)
    return s
cb=cost_blocks_slide()
# ============ ТЕКСТОВІ ПРАВКИ В СЛАЙДАХ КОРИСТУВАЧА ============
def settext(sh,new):
    p=sh.text_frame.paragraphs[0]; p.runs[0].text=new
    for r in p.runs[1:]: r.text=""
REPL=[("МАРЖА 30 %",f"БОНУС МЕРЕЖІ {pc(BON)}"),("маржа 30 %",f"бонус мережі {pc(BON)}")]
sl_old=list(prs.slides)
for i,sl in enumerate(sl_old,1):
    if i in (27,28,29): continue
    for sh in sl.shapes:
        if not sh.has_text_frame: continue
        t=sh.text_frame.text
        if len(sh.text_frame.paragraphs)==1 and sh.text_frame.paragraphs[0].runs:
            n=t
            for a,b in REPL:
                if a in n: n=n.replace(a,b)
            if i==23:
                if n=="100–120": n="100–120 ₴"
                if n=="120–140": n="120–140 ₴"
            if n!=t and i==23 and "₴" in n and len(n)<12: sh.width=Inches(1.1)
            if n!=t: settext(sh,n); print('slide',i,'|',t[:60],'->',n[:70])
# ============ ЗБІРКА ============
lst=prs.slides._sldIdLst
els={e.id:e for e in lst}
new_ids=[f33.slide_id,f6.slide_id,cb.slide_id,sens.slide_id]
order=OLD[:26]+new_ids
for e in list(lst): lst.remove(e)
for sid in order: lst.append(els[sid])
for sid in OLD[26:]:
    prs.part.drop_rel(els[sid].rId)
for i,s_ in enumerate(prs.slides,1):
    for sh in s_.shapes:
        if sh.has_text_frame and re.fullmatch(r"\d\d",sh.text_frame.text.strip()) and abs(Emu(sh.left).inches-11.55)<0.05:
            sh.text_frame.paragraphs[0].runs[0].text=f"{i:02d}"
prs.save(OUT); print("saved",OUT,len(prs.slides._sldIdLst))
