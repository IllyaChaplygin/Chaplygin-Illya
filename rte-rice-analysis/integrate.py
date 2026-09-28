# -*- coding: utf-8 -*-
"""Merge the part-2 analysis slides into the original market research deck."""
import re, copy
from pptx import Presentation
from pptx.util import Emu, Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
import data as d, deckkit as K, slides as S

ORIG="/root/.claude/uploads/df3fafbd-cd5a-5689-8909-9c95fa4cbf11/c731572d-RTE_Rice_Market_Research..pptx"
prs=Presentation(ORIG)
n_orig=len(prs.slides._sldIdLst)
assert n_orig==19, n_orig

def badge_shape(sl):
    """The page-number textbox: top-right corner, two digits."""
    for sh in sl.shapes:
        if sh.has_text_frame and re.fullmatch(r"\d\d", sh.text_frame.text.strip()):
            if Emu(sh.left).inches>11.5 and Emu(sh.top).inches<0.9:
                return sh
    return None

def set_badge(sl, num):
    sh=badge_shape(sl)
    if sh is None: return False
    p=sh.text_frame.paragraphs[0]
    if p.runs: p.runs[0].text=f"{num:02d}"
    for r in p.runs[1:]: r.text=""
    return True


# ── сегментні заголовки для карткових слайдів оригіналу ───────────
# (індекс у ОРИГІНАЛІ, ключ сегмента, kicker, заголовок, підзаголовок)
CARD_SEG=[
 (7,  "S1","BEN'S ORIGINAL · 13 ПОЗИЦІЙ",
      "Пауч 220–250 г — єдиний бренд на мережевій полиці",
      "Найглибша лінійка сегмента і єдина, що дійшла до продуктового роздрібу."),
 (8,  "S1","OTTOGI · 10 ПОЗИЦІЙ",
      "Чаша 217–320 г — корейський лідер азійського каналу",
      "Готова страва з наповнювачем. Найдорожча упаковка сегмента розігріву."),
 (9,  "S2","HENAN · 8 ПОЗИЦІЙ",
      "Чаша 144 і 174 г — найдешевша чаша ринку",
      "Сухий рис у чаші, заливається окропом на 8 хвилин. Маса — до заливання."),
 (10, "S1","ПООДИНОКІ БРЕНДИ · 4 БРЕНДИ, 4 ПОЗИЦІЇ",
      "Bibigo, Clearspring, Portion, Маркел — по одній позиції",
      "Той самий сегмент розігріву, але різні країни, ваги й канали. "
      "Gallina Blanca у правій частині — інший сегмент, рис під окріп."),
 (11, "S3","HAIDILAO · 10 ПОЗИЦІЙ",
      "Коробка 165–360 г — лідер саморозігріву",
      "Хімічний нагрівач у коробці: готується без мікрохвильовки й окропу."),
 (12, "S3","КИТАЙСЬКИЙ САМОРОЗІГРІВ · РЕШТА БРЕНДІВ",
      "Mo Xiao Xian, Zihaiguo, Rongcheng Haoji — 4 позиції",
      "Усі три бренди в одному магазині. Qiaoshanmei у правій частині — "
      "інший сегмент: рис у пакеті під окріп, не саморозігрів."),
 (13, "S4","ІМПОРТНИЙ СУБЛІМАТ · 4 БРЕНДИ, 12 ПОЗИЦІЙ",
      "Travellunch, Trek'n Eat, Mountain House, Adventure Food",
      "Німеччина, США, Нідерланди. Залити окропом на 8–10 хвилин."),
 (14, "S4","ГОТОВА СТРАВА І СУБЛІМАТ · 2 БРЕНДИ, 8 ПОЗИЦІЙ",
      "Adventure Menu і SubliMate",
      "Adventure Menu — єдиний бренд із двома технологіями: реторт 400 г і сублімат 110 г."),
 (15, "S4","УКРАЇНСЬКИЙ СУБЛІМАТ · 4 БРЕНДИ, 11 ПОЗИЦІЙ",
      "James Cook, Їжа в Похід, Харчі, !FEST",
      "Усе українське виробництво категорії. Канал — outdoor- і мілітарі-рітейл."),
]

def _put(sh, val):
    p=sh.text_frame.paragraphs[0]
    if p.runs:
        p.runs[0].text=val
        for r in p.runs[1:]: r.text=""

def retitle(sl, seg_key, kicker, title, sub):
    """Rewrite kicker/title/subtitle and stamp the segment label on the subtitle line."""
    kick_sh=title_sh=sub_sh=None
    for sh in sl.shapes:
        if not (sh.has_text_frame and sh.text_frame.text.strip()): continue
        x,y=Emu(sh.left).inches, Emu(sh.top).inches
        if 2.0<x<3.0 and 0.25<y<0.45: kick_sh=sh
        elif 2.0<x<3.0 and 0.50<y<0.75: title_sh=sh
        elif x<1.0 and 1.30<y<1.60: sub_sh=sh
    if kick_sh is not None:  _put(kick_sh,kicker)
    if title_sh is not None: _put(title_sh,title)

    name,col,tech,mass,occ,med = d.SEGMENTS[seg_key]
    # segment label + attributes occupy the left of the subtitle line
    lbl=f"{name}   ·   {tech}   ·   {mass}   ·   привід: {occ}"
    K.rect(sl,0.66,1.245,0.045,0.215,fill=col)
    K.text(sl,0.82,1.265,7.30,0.19,lbl,size=6.8,color=col,bold=True)
    K.text(sl,8.30,1.265,4.32,0.19,f"МЕДІАНА СЕГМЕНТА: {med.upper()}",
           size=6.8,color=d.MUTED,bold=True,align="r")
    # the original subtitle drops below the segment line
    if sub_sh is not None:
        _put(sub_sh,sub)
        sub_sh.top=Inches(1.50); sub_sh.left=Inches(0.66); sub_sh.width=Inches(11.90)


# ── section divider ────────────────────────────────────────────────
def divider(prs, kicker, title, sub, items):
    s=K.slide(prs, d.INK)
    K.rect(s,0,0,K.W,0.09,fill=d.C_BRAND)
    K.text(s,0.9,2.35,9,0.3,kicker,size=9,color=d.C_BRAND,bold=True)
    K.text(s,0.9,2.78,10.6,1.35,title,size=36,color=d.WHITE,bold=True,spacing=1.06)
    K.text(s,0.9,4.42,9.6,0.7,sub,size=11.5,color="#B9C2DA",spacing=1.5)
    for i,(k,v) in enumerate(items):
        x=0.9+i*2.95
        K.text(s,x,5.52,2.7,0.2,k,size=7,color="#717890",bold=True,caps=True)
        K.text(s,x,5.74,2.7,0.42,v,size=19,color=d.C_BRAND,bold=True)
    return s

# ── build the new slides onto the original deck ────────────────────
built={}
for fn in ("sl_02","sl_03","sl_04","sl_05","sl_06","sl_07","sl_08",
           "sl_09","sl_11","sl_12","sl_13","sl_14","sl_15"):
    built[fn]=getattr(S,fn)(prs)
DIV=divider(prs,"ПОЗИЦІОНУВАННЯ",
  "Де ми стоїмо\nі що заводити",
  "Зіставлення фінансової моделі з полицею, рекомендація формату й ціни, цільовий FOB.",
  [("НАШИХ SKU","32"),("ТОВАРНИХ ГРУП","7"),("ПОСТАЧАЛЬНИКІВ","3"),("СЦЕНАРІЇВ ЦІНИ","3")])

ids=list(prs.slides._sldIdLst)                 # 0..18 original, 19.. new
orig=ids[:n_orig]
def nid(fn): return ids[n_orig+list(built).index(fn)] if fn in built else None
new={fn:ids[n_orig+i] for i,fn in enumerate(built)}
new["DIV"]=ids[-1]

# ── apply segmentation to the original card slides ─────────────────
for oi,key,kick,ttl,sub in CARD_SEG:
    retitle(prs.slides[oi], key, kick, ttl, sub)

# ── target narrative order ─────────────────────────────────────────
order=[
 ("o",0,None),      # 01 титул
 ("o",1,2),         # 02 головне
 ("o",2,3),         # 03 обсяг ринку
 ("o",3,4),         # 04 карта ринку — 4 сегменти
 ("n","sl_04",5),   # 05 матриця брендів 1/2
 ("n","sl_05",6),   # 06 матриця брендів 2/2
 ("o",4,7),         # 07 формати упаковки
 ("n","sl_06",8),   # 08 аналітика форматів
 ("o",5,9),         # 09 грамаж
 ("n","sl_15",10),  # 10 модель формат × грамаж
 ("o",6,11),        # 11 цінові сходи брендів
 ("o",7,12),("o",8,13),("o",9,14),("o",10,15),   # 12–15 сегмент 1 + 2
 ("o",11,16),("o",12,17),                        # 16–17 сегмент 3
 ("o",13,18),("o",14,19),("o",15,20),            # 18–20 сегмент 4
 ("o",16,21),       # 21 суміжна полиця
 ("n","sl_03",22),  # 22 канали: точка продажу × покупець
 ("o",18,23),       # 23 висновки дослідження
 ("n","DIV",None),  # 24 роздільник
 ("n","sl_02",25),  # 25 шість висновків
 ("n","sl_07",26),  # 26 ранжир за упаковку
 ("n","sl_08",27),  # 27 ціна за 100 г
 ("n","sl_09",28),  # 28 фінмодель проти ринку
 ("n","sl_11",29),  # 29 рекомендація формату
 ("n","sl_12",30),  # 30 рекомендація ціни
 ("n","sl_13",31),  # 31 постачальник
 ("n","sl_14",32),  # 32 план дій
]
assert len(order)==32, len(order)

sldIdLst=prs.slides._sldIdLst
seq=[orig[r] if k=="o" else new[r] for k,r,_ in order]
for el in list(sldIdLst): sldIdLst.remove(el)
for el in seq: sldIdLst.append(el)

# ── renumber every badge to its new position ───────────────────────
missing=[]
for i,(sl,(kind,ref,badge)) in enumerate(zip(prs.slides,order),1):
    if badge is None: continue
    if not set_badge(sl,badge): missing.append(i)
if missing: print("badge not found on:",missing)

# ── refresh the title slide for the expanded scope ─────────────────
t=prs.slides[0]
for sh in t.shapes:
    if sh.has_text_frame and "ДОСЛІДЖЕННЯ РИНКУ" in sh.text_frame.text:
        p=sh.text_frame.paragraphs[0]
        if p.runs:
            p.runs[0].text="ДОСЛІДЖЕННЯ РИНКУ ТА ПОЗИЦІОНУВАННЯ · ВЕРЕСЕНЬ 2026"
            for r in p.runs[1:]: r.text=""

prs.save("RTE_Rice_Market_Research_FULL.pptx")
print("saved:",len(prs.slides._sldIdLst),"slides")
