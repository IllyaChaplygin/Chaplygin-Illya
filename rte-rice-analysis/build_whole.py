# -*- coding: utf-8 -*-
"""The complete deck, assembled in a fresh package.

Every build that failed was Presentation(their_file) saved back out. The one
that opens was built from a fresh presentation. So build the whole thing
fresh and copy their 19 slides into it, shapes and images and all, instead
of re-serialising their package.

Their slides carry explicit colours and fonts (1674 srgbClr, Segoe UI
everywhere); the schemeClr references sit in the <p:style> defaults that the
explicit fills override, so moving them to another theme does not change how
they look.
"""
from pptx import Presentation
import re
from pptx.util import Emu, Inches
import data as d
import deckkit as K
import slides as S
import merge_suppliers as MS

ORIG = "/root/.claude/uploads/df3fafbd-cd5a-5689-8909-9c95fa4cbf11/c731572d-RTE_Rice_Market_Research..pptx"
OUT = "RTE_Rice_Market_Research_FINAL.pptx"


# ── сегмент + формат у шапці кожної картки брендів ─────────────────
d.SEGMENTS["S5"] = ("СЕГМЕНТ 5 · СУМІЖНА ПОЛИЦЯ", d.C_GREEN,
                    "Готова до вживання", "Маса — готової страви",
                    "Обід поза домом", "119 грн")

# індекс у оригіналі -> (сегмент, формати, kicker, заголовок, підзаголовок)
CARDS = {
 7:  ("S1","ПАУЧ","BEN'S ORIGINAL · 13 ПОЗИЦІЙ",
      "Пауч 220–250 г — єдиний бренд на мережевій полиці",
      "Найглибша лінійка сегмента і єдина, що дійшла до продуктового роздрібу."),
 8:  ("S1","ЧАША","OTTOGI · 10 ПОЗИЦІЙ",
      "Чаша 217–320 г — корейський лідер азійського каналу",
      "Готова страва з наповнювачем. Найдорожча упаковка сегмента розігріву."),
 9:  ("S2","ЧАША","HENAN · 8 ПОЗИЦІЙ",
      "Чаша 144 і 174 г — найдешевша чаша ринку",
      "Сухий рис у чаші, заливається окропом на 8 хвилин."),
 10: ("S1","ЧАША + ПАУЧ","ПООДИНОКІ БРЕНДИ · 4 БРЕНДИ",
      "Bibigo — чаша, Clearspring, Portion і Маркел — пауч",
      "Той самий сегмент розігріву в двох форматах. Gallina Blanca праворуч — "
      "інший сегмент: сухий рис у чаші під окріп."),
 11: ("S3","КОРОБКА","HAIDILAO · 10 ПОЗИЦІЙ",
      "Коробка 165–360 г — лідер саморозігріву",
      "Хімічний нагрівач у коробці: готується без мікрохвильовки й окропу."),
 12: ("S3","ЧАША + КОРОБКА","КИТАЙСЬКИЙ САМОРОЗІГРІВ · 3 БРЕНДИ",
      "Mo Xiao Xian — чаша, Zihaiguo і Rongcheng Haoji — коробка",
      "Усі в одному магазині. Qiaoshanmei праворуч — інший сегмент: рис у пакеті під окріп."),
 13: ("S4","ДОЙПАК","ІМПОРТНИЙ СУБЛІМАТ · 4 БРЕНДИ, 12 ПОЗИЦІЙ",
      "Дойпак 110–250 г — Німеччина, США, Нідерланди",
      "Залити окропом на 8–10 хвилин. У продуктовий роздріб ці бренди не йдуть."),
 14: ("S4","ПАУЧ + ДОЙПАК","ДВІ ТЕХНОЛОГІЇ В ОДНОМУ БРЕНДІ · 8 ПОЗИЦІЙ",
      "Adventure Menu — пауч 400 г і дойпак 110 г, SubliMate — дойпак",
      "Єдиний бренд категорії, що робить і готову страву, і сублімат."),
 15: ("S4","ДОЙПАК","УКРАЇНСЬКИЙ СУБЛІМАТ · 4 БРЕНДИ, 11 ПОЗИЦІЙ",
      "Дойпак 80–100 г — усе вітчизняне виробництво категорії",
      "Канал — туристичні та військові магазини."),
 16: ("S5","БЛЯШАНКА","СУМІЖНА ПОЛИЦЯ · 6 ПОЗИЦІЙ",
      "Рис із м'ясом у мережах уже є — але в консервній бляшанці",
      "Прямий конкурент за той самий привід: гаряча страва з рисом і м'ясом без готування."),
}
FMT_COL = {"ПАУЧ": d.C_AMBER, "ЧАША": d.C_TEAL, "ДОЙПАК": d.C_PURPLE,
           "КОРОБКА": d.C_PINK, "БЛЯШАНКА": d.C_GREEN}


def _put(sh, val):
    p = sh.text_frame.paragraphs[0]
    if p.runs:
        p.runs[0].text = val
        for r in p.runs[1:]:
            r.text = ""


def dress_card_slide(sl, seg, fmts, kicker, title, sub):
    """Name the slide by its segment and format, and drop the per-100g figures."""
    kick = ttl = sb = None
    for sh in sl.shapes:
        if not (sh.has_text_frame and sh.text_frame.text.strip()):
            continue
        x, y = Emu(sh.left).inches, Emu(sh.top).inches
        if 2.0 < x < 3.0 and 0.25 < y < 0.45: kick = sh
        elif 2.0 < x < 3.0 and 0.50 < y < 0.75: ttl = sh
        elif x < 1.0 and 1.30 < y < 1.60: sb = sh
    name, col, tech, mass, occ, med = d.SEGMENTS[seg]
    if kick is not None: _put(kick, f"{name}   ·   {fmts}   ·   {kicker}")
    if ttl is not None:  _put(ttl, title)
    if sb is not None:
        _put(sb, sub)
        sb.top, sb.left, sb.width = Inches(1.50), Inches(0.66), Inches(11.90)
    K.rect(sl, 0.66, 1.245, 0.045, 0.215, fill=col)
    K.text(sl, 0.82, 1.265, 11.6, 0.19,
           f"ФОРМАТ: {fmts}   ·   {tech}   ·   {mass}   ·   привід: {occ}",
           size=6.8, color=col, bold=True)
    # прибрати грн/100 г
    kill = [sh for sh in sl.shapes
            if sh.has_text_frame
            and (re.search(r"грн\s*/\s*100\s*г", sh.text_frame.text)
                 or sh.text_frame.text.strip().upper() in ("ЗА 100 Г", "ГРН/100 Г"))]
    for sh in kill:
        sh._element.getparent().remove(sh._element)
    return len(kill)

src = Presentation(ORIG)
out = K.new_deck()

# their slide index -> where it goes, and which of mine follow it
PLAN = [
    (0, []), (1, []), (2, []), (3, ["sl_04", "sl_05"]),
    (4, ["sl_06"]), (5, ["sl_15", "sl_16"]), (6, []),
    (7, []), (8, []), (9, []), (10, []), (11, []), (12, []),
    (13, []), (14, []), (15, []),
    (16, ["sl_17", "sl_18", "sl_03"]),
    (17, []),
    (18, ["sl_02", "sl_07", "sl_09", "sl_11", "sl_12", "sl_13", "sl_14"]),
]

mine = set()
stripped = 0
for oi, after in PLAN:
    copy = MS.copy_slide(src.slides[oi], out)
    if oi in CARDS:
        stripped += dress_card_slide(copy, *CARDS[oi])
    for fn in after:
        sl = getattr(S, fn)(out)
        mine.add(id(sl._element))

# page numbers on my slides; theirs keep whatever they had
for pos, sl in enumerate(out.slides, 1):
    for sh in sl.shapes:
        if (sh.has_text_frame
                and re.fullmatch(r"\d\d", sh.text_frame.text.strip())
                and Emu(sh.left).inches > 11.5 and Emu(sh.top).inches < 0.9):
            p = sh.text_frame.paragraphs[0]
            if p.runs:
                p.runs[0].text = f"{pos:02}"
                for r in p.runs[1:]:
                    r.text = ""

out.save(OUT)
print(f"прибрано грн/100 г: {stripped}")
print(f"{OUT}: {len(out.slides._sldIdLst)} слайдів "
      f"({len(PLAN)} ваших + {len(mine)} нових)")
