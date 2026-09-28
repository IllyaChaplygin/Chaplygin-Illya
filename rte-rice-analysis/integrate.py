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
    if kick_sh is not None:  _put(kick_sh, d.SEGMENTS[seg_key][0] + "   ·   " + kicker)
    if title_sh is not None: _put(title_sh,title)

    name,col,tech,mass,occ,med = d.SEGMENTS[seg_key]
    lbl=f"{tech}   ·   {mass}   ·   привід: {occ}   ·   медіана сегмента {med.split()[0]} грн"
    K.rect(sl,0.66,1.245,0.045,0.215,fill=col)
    K.text(sl,0.82,1.265,11.6,0.19,lbl,size=6.8,color=col,bold=True)
    # the original subtitle drops below the segment line
    if sub_sh is not None:
        _put(sub_sh,sub)
        sub_sh.top=Inches(1.50); sub_sh.left=Inches(0.66); sub_sh.width=Inches(11.90)



# ── формат кожної товарної картки ──────────────────────────────────
FMT_COL={"ПАУЧ":d.C_AMBER,"ЧАША":d.C_TEAL,"ДОЙПАК":d.C_PURPLE,
         "КОРОБКА":d.C_PINK,"ПАКЕТ":d.C_TEAL}
# індекс оригіналу -> (формат за замовчуванням, [(маркер у тексті, формат, мітка сегмента)])
CARD_FMT={
 7:  ("ПАУЧ",    []),
 8:  ("ЧАША",    []),
 9:  ("ЧАША",    []),
 10: ("ПАУЧ",    [("BIBIGO","ЧАША",None),("GALLINA","ЧАША","СЕГ. 2")]),
 11: ("КОРОБКА", []),
 12: ("КОРОБКА", [("MO XIAO XIAN","ЧАША",None),("QIAOSHANMEI","ПАКЕТ","СЕГ. 2")]),
 13: ("ДОЙПАК",  []),
 14: ("ДОЙПАК",  [("400 г","ПАУЧ","СЕГ. 1")]),
 15: ("ДОЙПАК",  []),
}

def product_cards(sl):
    """Every product card, found by its text block: (name_x, name_y, width, text)."""
    bx=[]
    for sh in sl.shapes:
        if not (sh.has_text_frame and sh.text_frame.text.strip()): continue
        x,y,w=Emu(sh.left).inches, Emu(sh.top).inches, Emu(sh.width).inches
        if y<1.80 or y>6.95 or w>3.2: continue
        bx.append((round(x,2), y, w, sh.text_frame.text.strip()))
    cols={}
    for x,y,w,txt in bx: cols.setdefault(x,[]).append((y,w,txt))
    clusters=[]
    for x,items in cols.items():
        items.sort()
        cur=[items[0]]
        for it in items[1:]:
            if it[0]-cur[-1][0] > 0.45: clusters.append((x,cur)); cur=[it]
            else: cur.append(it)
        clusters.append((x,cur))
    out=[]
    for x,c in clusters:
        ny,nw,ntxt = c[0]
        if ny < 2.30: continue                      # шапка бренду, не картка
        if len(c) < 3: continue                     # коментар-блок, не картка
        if len(ntxt) > 45: continue                 # текст висновку, не назва
        if not any("грн" in t for _,_,t in c): continue
        body=" ".join(t for _,_,t in c)
        body=" ".join(body.split()).upper()
        out.append((x, ny, nw, body))
    return sorted(out, key=lambda a:(a[1],a[0]))

def stamp_formats(sl, oi):
    default, rules = CARD_FMT[oi]
    n=0
    for x,y,w,body in product_cards(sl):
        fmt, segmark = default, None
        for marker, f, sm in rules:
            if marker.upper() in body:
                fmt, segmark = f, sm
                break
        label = fmt if not segmark else f"{fmt} · {segmark}"
        bw = min(0.112*len(label)**0.88 + 0.17, max(w, 1.0))
        K.rect(sl, x-0.05, y-0.215, bw, 0.175, fill=FMT_COL[fmt], radius=0.45)
        K.text(sl, x-0.05, y-0.195, bw, 0.15, label, size=5.6,
               color=d.WHITE, bold=True, align="c")
        n+=1
    return n


# ── section divider ────────────────────────────────────────────────
def divider(prs, kicker, title, sub, items):
    s=K.slide(prs, d.INK)
    K.rect(s,0,0,K.W,0.09,fill=d.C_BRAND)
    K.text(s,0.9,2.35,9,0.3,kicker,size=9,color=d.C_BRAND,bold=True)
    K.text(s,0.9,2.78,10.6,1.35,title,size=36,color=d.WHITE,bold=True,spacing=1.06)
    K.text(s,0.9,4.42,9.6,0.7,sub,size=11.5,color="#B9C2DA",spacing=1.5)
    K.text(s,11.97,0.44,0.70,0.36,"00",size=17,color="#5B6B8C",bold=True,align="r")
    for i,(k,v) in enumerate(items):
        x=0.9+i*2.95
        K.text(s,x,5.52,2.7,0.2,k,size=7,color="#717890",bold=True,caps=True)
        K.text(s,x,5.74,2.7,0.42,v,size=19,color=d.C_BRAND,bold=True)
    return s

# ── build the new slides onto the original deck ────────────────────
built={}
for fn in ("sl_02","sl_03","sl_04","sl_05","sl_06","sl_07",
           "sl_09","sl_10","sl_11","sl_12","sl_13","sl_14","sl_15","sl_16","sl_17","sl_18"):
    built[fn]=getattr(S,fn)(prs)
DIV=divider(prs,"ПОЗИЦІОНУВАННЯ",
  "Де ми стоїмо\nі що заводити",
  "Зіставлення фінансової моделі з полицею, рекомендація формату й ціни, цільовий FOB.",
  [("НАШИХ SKU","32"),("ТОВАРНИХ ГРУП","7"),("ПОСТАЧАЛЬНИКІВ","3"),("СЦЕНАРІЇВ ЦІНИ","3")])



def strip_per100(sl):
    """Drop every 'X грн/100 г' figure — the client does not use that metric."""
    import re as _re
    kill=[]
    for sh in sl.shapes:
        if not sh.has_text_frame: continue
        txt=sh.text_frame.text.strip()
        if _re.search(r"грн\s*/\s*100\s*г", txt) or txt.upper() in ("ЗА 100 Г","ГРН/100 Г"):
            kill.append(sh)
    for sh in kill:
        sh._element.getparent().remove(sh._element)
    return len(kill)

# ── supplier deck: merge and restyle to the main masthead ──────────

def _send_to_back(sl, shape):
    """Move a shape behind everything else on the slide."""
    tree = sl.shapes._spTree
    el = shape._element
    tree.remove(el)
    tree.insert(2, el)

def restyle_supplier(sl):
    """Give a merged supplier slide the same masthead as the rest of the deck."""
    kick = ttl = sub = None
    for sh in sl.shapes:
        if not (sh.has_text_frame and sh.text_frame.text.strip()): continue
        x, y = Emu(sh.left).inches, Emu(sh.top).inches
        if x < 1.0 and 0.20 < y < 0.45: kick = sh
        elif x < 1.0 and 0.50 < y < 0.75: ttl = sh
        elif x < 1.0 and 1.10 < y < 1.35: sub = sh
    band = K.rect(sl, 0, 0, K.W, 1.16, fill=d.INK2)
    rule = K.rect(sl, 0, 1.16, K.W, 0.06, fill=d.C_BRAND)
    _send_to_back(sl, rule); _send_to_back(sl, band)
    try: K.img(sl, K.LOGO, 0.66, 0.26, w=1.52, h=0.66)
    except Exception: pass
    if kick is not None:
        kick.left, kick.top, kick.width = Inches(2.52), Inches(0.34), Inches(7.40)
        for r in kick.text_frame.paragraphs[0].runs:
            r.font.color.rgb = K.C("#FFC95C"); r.font.size = Pt(9.5); r.font.bold = True
    if ttl is not None:
        ttl.left, ttl.top, ttl.width = Inches(2.52), Inches(0.60), Inches(8.80)
        for r in ttl.text_frame.paragraphs[0].runs:
            r.font.color.rgb = K.C(d.WHITE); r.font.size = Pt(19); r.font.bold = True
    if sub is not None:
        sub.left, sub.top, sub.width = Inches(0.66), Inches(1.37), Inches(11.90)
        for r in sub.text_frame.paragraphs[0].runs:
            r.font.color.rgb = K.C(d.MUTED); r.font.size = Pt(9.5)
    # page badge, added last so it sits on top of the band
    K.text(sl, 11.97, 0.44, 0.70, 0.36, "00", size=17,
           color="#465382", bold=True, align="r")


ids=list(prs.slides._sldIdLst)                 # 0..18 original, 19.. new
orig=ids[:n_orig]
def nid(fn): return ids[n_orig+list(built).index(fn)] if fn in built else None
new={fn:ids[n_orig+i] for i,fn in enumerate(built)}
new["DIV"]=ids[-1]


# ── apply segmentation to the original card slides ─────────────────

# ── дослівні заміни в текстах оригіналу: прибрати метрику «за 100 г» ──
TEXT_FIXES = {
 3:  [("Медіана готового рису для розігріву (30 позицій) — 73 грн за 100 г.",
       "Медіана готового рису для розігріву (30 позицій) — 164 грн за упаковку.")],
 15: [("український реторт 350 г за 94–109 грн: найдешевший грам категорії, 27–31 грн за 100 г.",
       "український реторт 350 г за 94–109 грн: найдешевша упаковка сегмента.")],
 17: [("коробка 440 г за 735 грн: 167 грн за 100 г, дешевше за грам, ніж Haidilao (230–371).",
       "коробка 440 г за 735 грн — удвічі більша порція, ніж у Haidilao за ті самі гроші.")],
 19: [("реторт без води, 79–101 грн за 100 г: у тому ж коридорі, що Ben's Original і Ottogi,",
       "реторт без води, 315–404 грн за упаковку: дорожче за Ben's Original і Ottogi,")],
 21: [("Порівняння за 100 г", "Порівняння за упаковку"),
      ("18,5 грн", "63 грн"), ("47,6 грн", "119 грн"), ("41,0 грн", "90 грн"),
      ("45,7 грн", "153 грн"), ("88,0 грн", "255 грн")],
}

def _font_of(tf):
    """The formatting of the first run, to re-apply after a text rewrite."""
    for p in tf.paragraphs:
        for r in p.runs:
            col = None
            try:
                if r.font.color and r.font.color.type is not None:
                    col = r.font.color.rgb
            except Exception:
                pass
            return dict(size=r.font.size, bold=r.font.bold, italic=r.font.italic,
                        name=r.font.name, color=col)
    return {}


def apply_text_fixes(prs_):
    """Rewrite phrases in the original slides.

    Uses the TextFrame.text setter rather than editing runs by hand: it
    rebuilds the paragraphs cleanly instead of leaving blank runs behind.
    Run-level formatting inside a paragraph is lost, so the first run's font
    is re-applied to keep size, weight and colour.
    """
    n = 0
    for oi, pairs in TEXT_FIXES.items():
        sl = prs_.slides[oi]
        for sh in sl.shapes:
            if not (sh.has_text_frame and sh.text_frame.text.strip()):
                continue
            tf = sh.text_frame
            txt = tf.text
            hits = [(o, w) for o, w in pairs if o in txt]
            if not hits:
                continue
            font = _font_of(tf)
            for o, w in hits:
                txt = txt.replace(o, w); n += 1
            tf.text = txt
            for p in tf.paragraphs:
                for r in p.runs:
                    if font.get("size"):  r.font.size = font["size"]
                    if font.get("name"):  r.font.name = font["name"]
                    r.font.bold = font.get("bold")
                    r.font.italic = font.get("italic")
                    if font.get("color") is not None:
                        r.font.color.rgb = font["color"]
    return n


# сегмент 5 — суміжна консервна полиця
d.SEGMENTS["S5"]=("СЕГМЕНТ 5 · СУМІЖНА ПОЛИЦЯ · КОНСЕРВА",d.C_GREEN,
                  "Готова до вживання","Маса — готової страви","Обід поза домом","119 грн")
CARD_SEG=CARD_SEG+[(16,"S5","КОНСЕРВА НА ПОЛИЦІ · 6 ПОЗИЦІЙ",
   "Рис із м'ясом у мережах уже є — але в бляшанці",
   "Прямий конкурент за той самий привід: гаряча страва з рисом і м'ясом без готування.")]
CARD_FMT[16]=("КОНСЕРВА",[])
FMT_COL["КОНСЕРВА"]=d.C_GREEN

for oi,key,kick,ttl,sub in CARD_SEG:
    retitle(prs.slides[oi], key, kick, ttl, sub)
    _n=stamp_formats(prs.slides[oi], oi)
    _k=strip_per100(prs.slides[oi])
    print(f'  orig{oi+1:02}: {_n} карток, прибрано грн/100г: {_k}')

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
 ("n","sl_16",11),  # 11 найпопулярніші поєднання формат + грамаж
 ("o",6,12),        # 12 цінові сходи брендів
 ("o",7,13),("o",8,14),("o",9,15),("o",10,16),   # 13–16 сегмент 1 + 2
 ("o",11,17),("o",12,18),                        # 17–18 сегмент 3
 ("o",13,19),("o",14,20),("o",15,21),            # 19–21 сегмент 4
 ("o",16,22),       # 22 суміжна полиця · сегмент 5
 ("n","sl_17",23),  # 23 ритейл-аудит мереж
 ("n","sl_18",24),  # 24 азійські бренди в мережах
 ("n","sl_03",25),  # 25 канали: точка продажу × покупець
 ("o",17,26),       # 26 де представлено (оригінал)
 ("o",18,27),       # 27 висновки дослідження
 ("n","DIV",28),   # 28 роздільник
 ("n","sl_02",29),  # 29 шість висновків
 ("n","sl_07",30),  # 30 ранжир за упаковку
 ("n","sl_09",31),  # 31 фінмодель проти ринку
 ("n","sl_10",32),  # 32 аудит фінмоделі
 ("n","sl_11",33),  # 33 рекомендація формату
 ("n","sl_12",34),  # 34 рекомендація ціни
 ("n","sl_13",35),  # 35 постачальник
 ("n","sl_14",36),  # 36 план дій
]
assert len(order)==36, len(order)

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

print("текстових замін:", apply_text_fixes(prs))
prs.save("RTE_Rice_Market_Research_FULL.pptx")
print("saved:",len(prs.slides._sldIdLst),"slides")

# структурна перевірка пакета — щоб биті посилання не проходили тихо
import subprocess, sys as _sys
_r = subprocess.run([_sys.executable, "check_pkg.py", "RTE_Rice_Market_Research_FULL.pptx"],
                    capture_output=True, text=True)
print(_r.stdout.strip())
if "ПРОБЛЕМ" in _r.stdout:
    _sys.exit("ПАКЕТ БИТИЙ — не віддавати файл")
