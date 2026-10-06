# -*- coding: utf-8 -*-
"""Збирає v5_*.py з v4_*.py правками (без сегмента 4) + нові слайди."""
import re,sys
SC="/tmp/claude-0/-home-user-Chaplygin-Illya/8bc0034a-4b53-5022-9bf5-e5697c987387/scratchpad"
def rd(n): return open(f"{SC}/{n}",encoding='utf-8').read()
def wr(n,t): open(f"{SC}/v5/{n}",'w',encoding='utf-8').write(t)
def rep(t,a,b,cnt=1):
    assert t.count(a)==cnt,(a[:60],t.count(a)); return t.replace(a,b)

# ---------- A: обкладинка, шість цифр, митниця ----------
a=rd("v4_a.py")
a=rep(a,'"ПОЗИЦІЙ У ПРОДАЖУ","96","11 брендів · 4 сегменти · 13 мереж із 18. Кожна позиція має ціну в каталозі."',
        '"ПОЗИЦІЙ У ПРОДАЖУ",str(F[\'n_sku\']),f"{F[\'n_brands\']} брендів · 3 сегменти · {F[\'n_chains\']} мереж із 18. Кожна позиція має ціну в каталозі."')
a=rep(a,'"РОЗРИВ ЦІН ФОРМАТІВ","×1,8","За 100 г: стакан 154 грн проти 86 грн у пакеті з бульйоном."',
        '"РОЗРИВ ЦІН ФОРМАТІВ",f"×{d1(SS[\'3\'][\'g\']/SS[\'1\'][\'g\'])}",f"Ціна за 100 г: стакан {SS[\'3\'][\'g\']} грн проти {SS[\'1\'][\'g\']} грн у пакеті з бульйоном."')
a=rep(a,'"3. Полиця концентрована: Samyang і Nongshim дають 54 позиції з 96 і стоять у 10–11 мережах."',
        'f"3. Полиця концентрована: Samyang і Nongshim дають {N_TOP2} позицій з {F[\'n_sku\']} і стоять у {min(TOP2CH.values())}–{max(TOP2CH.values())} мережах."')
a=rep(a,'каталоги мереж 29.09.2026 (96 SKU)','каталоги мереж 29.09.2026 ({F[\'n_sku\']} SKU)')
a=a.replace('slide_tail_sources(s,"Джерела: каталоги мереж 29.09.2026 ({F','slide_tail_sources(s,f"Джерела: каталоги мереж 29.09.2026 ({F')
wr("v5_a.py",a)

# ---------- B: сегментація, поверхи ----------
b=rd("v4_b.py"); i6=b.index("# ============ 06 ЦІНОВІ ПОВЕРХИ")
seg5='''# ============ 05 СЕГМЕНТАЦІЯ ============
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

'''
rest=b[i6:]
rest=rep(rest,'"34–204 грн  ·  93–146 грн/100 г"','"34–198 грн  ·  93–146 грн/100 г"')
rest=rep(rest,"Samyang, Nongshim і в'єтнамська рисова локшина · 8 брендів","Samyang, Nongshim та імпортна азійська локшина · 9 брендів")
rest=rep(rest,'tl("Kool Cung Dinh","Cung Dinh Kool")','tl("Micoem")')
rest=rep(rest,'двома брендами, що тримають 54 позиції з 96. "','двома брендами, що тримають {N_TOP2} позицій з {F[\'n_sku\']}. "')
rest=rep(rest,' "При полиці 120 грн наші SKU стають між',' f"При полиці 120 грн наші SKU стають між')
wr("v5_b.py",seg5+rest)

# ---------- C: блоки сегментів ----------
c=rd("v4_c.py")
c=rep(c,'"ЦІНА ЗА ГРАМ"','"ЦІНА ЗА 100 Г"')
c=rep(c,"for k in \"1234\"}","for k in \"123\"}")
wr("v5_c.py",c)

# ---------- D: сегменти 1–3 ----------
d=rd("v4_d.py")
d=rep(d,"S1,S2,S3,S4=SS['1'],SS['2'],SS['3'],SS['4']","S1,S2,S3=SS['1'],SS['2'],SS['3']")
i3=d.index("_cups=SEG['3']")
d=d[:i3]+'catalog_series("3")\n\n'
wr("v5_d.py",d)

# ---------- E: формат×вага, формати, (ціни — нові) ----------
e=rd("v4_e.py")
e=rep(e,"'145 г +':'рисові галушки та нестандарт, інший привід споживання'","'145 г +':'поодинокі позиції'")
e=rep(e,'"ЦІНА ЗА ГРАМ — ВИЩА ЗА ВІДПОВІДНІ СМУГИ"','"ЦІНА ЗА 100 Г — ВИЩА ЗА ВІДПОВІДНІ СМУГИ"')
e=rep(e,'''    _f="стакан" if _r['seg']==3 else ("рисові галушки" if _r['seg']==4 else "пакет")''','''    _f="стакан" if _r['seg']==3 else "пакет"''')
e=rep(e,'"22 позиції, медіана 93 грн, 77 грн/100 г.\\nNongshim, Ottogi, Paldo — 11 мереж."','f"{len(_combos[0][1])} позицій, медіана {_top_med:.0f} грн, {st.median([x[\'per100\'] for x in _top]):.0f} грн/100 г.\\nNongshim, Ottogi, Paldo — 11 мереж."')
ip=e.index("# ============ ЦІНА ЗА ОДНУ УПАКОВКУ"); e_head=e[:ip]
nows=[l for l in e[ip:].split("\n") if l.startswith("NOWS=")][0]
wr("v5_e1.py",e_head+"\n"+nows+"\n")
e_tail=rd("v4_f.py")
e_tail=rep(e_tail,"Samyang · Nongshim · Paldo · Yopokki","Samyang · Nongshim · Paldo")
e_tail=rep(e_tail,'"«Космос» і «Таврія В» мають ідентичні 12 позицій і той самий діапазон цін."','f"«Космос» і «Таврія В» — по {_n_cos} позицій; {_n_same} із них збігаються за EAN. Найімовірніше, один постачальник на дві мережі."')
e_tail=rep(e_tail,'ecall(s,0.62,6.52,3.88,0.90,"ДІРКА №1','_cs={z[\'ean\'] for z in ZREC if z[\'chain\']==\'cosmos\'}; _ts={z[\'ean\'] for z in ZREC if z[\'chain\']==\'tavriav\'}\n_n_cos=len(_cs); _n_same=len(_cs&_ts)\necall(s,0.62,6.52,3.88,0.90,"ДІРКА №1')
# теплову карту «бренд × мережа» (слайд 27) прибрано: її заміняє зведення на слайді 20
ih=e_tail.index("# ============ РИТЕЙЛ-АУДИТ · ПОЗИЦІЇ")
e_tail=e_tail[:ih]
wr("v5_e2.py",e_tail)

# ---------- G: канали, суміжна, Choi's, ... ----------
g_=rd("v4_g.py")
g_=rep(g_,"ціна за грам — від 49 до 158 грн/100 г","ціна за 100 г — від 49 до 158 грн")
wr("v5_g.py",g_)
# ---------- H: акції (без аргументів) ----------
h=rd("v4_h.py"); ih=h.index("# ============ АРГУМЕНТИ ДЛЯ SIAS"); wr("v5_h.py",h[:ih])
print("ok")
