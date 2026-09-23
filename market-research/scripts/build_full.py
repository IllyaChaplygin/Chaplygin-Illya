"""Готовий рис · Україна — глибокий розбір ринку по кожному SKU.

Кожен бренд отримує власний розворот: усі його позиції з пакшотом, масою,
ціною за упаковку та ціною за 100 грамів. Без наших позицій.

Фірмовий стиль Morskyi Dim: #303A5D / #F9A50B, хвильовий патерн і логотип.
"""
import statistics as st

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from PIL import Image

W, H = 13.3333, 7.5
M = 0.66
HDR = 1.16

NAVY   = RGBColor(0x30, 0x3A, 0x5D)
NAVY_L = RGBColor(0x46, 0x53, 0x82)
ORANGE = RGBColor(0xF9, 0xA5, 0x0B)
AMBER  = RGBColor(0xFF, 0xC9, 0x5C)
TEAL   = RGBColor(0x1E, 0x9E, 0xA6)
GREEN  = RGBColor(0x37, 0xA1, 0x69)
PLUM   = RGBColor(0x7B, 0x5E, 0xA7)
ROSE   = RGBColor(0xC9, 0x4F, 0x7C)
SLATE  = RGBColor(0x5B, 0x6B, 0x8C)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
MIST   = RGBColor(0xF1, 0xF4, 0xFA)
MIST_D = RGBColor(0xE2, 0xE8, 0xF3)
INK    = RGBColor(0x1C, 0x22, 0x35)
GREY   = RGBColor(0x71, 0x78, 0x90)

FONT = "Segoe UI"
BG_TITLE = "brand/md_bg_title.jpg"
LOGO = "brand/md_logo_white.png"

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(W), Inches(H)
BLANK = prs.slide_layouts[6]
PG = {"n": 0}


# ══ ПРЕДМЕТ ДОСЛІДЖЕННЯ ═══════════════════════════════════════════════════
# Тільки готовий рис, який зберігається за кімнатної температури і
# розігрівається — мікрохвильовка 60–120 с або занурення пауча в окріп.
# НЕ входять: рис під заливання окропом, саморозігрів, сублімація,
# заморозка, охолоджені страви, суха крупа й варильні пакети.
#
# Кожна ціна — роздрібна за ОДНУ упаковку, перевірена на сайті продавця
# 23 вересня 2026 р. Оптові пороги (від 3–20 шт) до розрахунку не входять.

# (файл пакшоту, назва, грамів, ціна_від, ціна_до, канал)
# ── Гарнір: чистий рис у паучі, мікрохвильовка 90 с ───────────────────────
BENS = [
    ("md/bens_longgrain.jpg",     "Long Grain",        250, 45, 45,       "Сільпо"),
    ("md/bens_basmati250.jpg",    "Basmati",           250, 89.99, 89.99, "Сільпо"),
    ("md/bens_mediterran.jpg",    "Mediterran",        250, 89.99, 89.99, "Сільпо"),
    ("md/bens_risibisi.jpg",      "Risi Bisi",         250, 89.99, 89.99, "Сільпо"),
    ("md/bens_curry_indien.jpg",  "Indian Curry",      250, 89.99, 89.99, "Сільпо"),
    ("md/bens_curry_linsen.jpg",  "Curry з сочевицею", 220, 89.99, 89.99, "Сільпо"),
    ("md/bens_mexikanisch.jpg",   "Mexikanisch",       220, 109, 109,     "Сільпо"),
    ("md/bens_curryreis.jpg",     "Curryreis Indien",  220, 144, 144,     "Сільпо"),
    ("md/bens_basmati220.jpg",    "Basmati",           220, 144, 144,     "Сільпо"),
    ("md/bens_sweetchili.jpg",    "Sweet Chili",       220, 179, 179,     "Сільпо"),
    ("sku/bens_lang220_single.jpg", "Long Grain",      220, 139, 139,     "Edison Lee"),
]
# ── Страва в чаші: рис із наповнювачем, мікрохвильовка ────────────────────
OTTOGI = [
    ("sku/ot_bulgogi.jpg",     "Бульгогі",            320, 252, 252, "Тайякі Март"),
    ("sku/ot_pork310.jpg",     "Свинина",             310, 252, 252, "Тайякі Март"),
    ("sku/ot_chicken_rib.jpg", "Гострі курячі ребра", 310, 260, 289, "Gurmissimo · Апетітаріум"),
    ("sku/ot_tuna.jpg",        "Тунець і майонез",    247, 225, 356, "Тайякі · Gurm. · Апетіт."),
    ("sku/ot_bibimbap.jpg",    "Пібімпаб",            269, 252, 334, "Тайякі · Апетітаріум"),
    ("sku/ot_pork_spicy.jpg",  "Гостра свинина",      269, 334, 334, "Апетітаріум"),
    ("sku/ot_beef320.jpg",     "Гостра яловичина",    320, 356, 356, "Апетітаріум"),
]
# (файл, бренд, назва, грамів, від, до, канал, колір)
BOWL_MORE = [
    ("md/bibigo_bowl.jpg", "Bibigo", "Білий рис", 210, 189, 189, "Тайякі Март", GREEN),
]
# ── Страва в реторт-паучі, українське виробництво ─────────────────────────
UA_RETORT = [
    (None, "Portion", "Каша рисова з куркою", 350, 71.4, 71.4,
     "portion.com.ua", ROSE),
    ("sku/ua_markel_soy.jpg", "Маркел", "Рис із соєвим м'ясом", 350, 80, 96,
     "Мартел-shop · СУХПАЙ · UPcompany", SLATE),
    ("sku/ua_pak_pork.jpg", "Маркел", "Каша рисова зі свининою", 350, 99, 99,
     "UPcompany", SLATE),
    ("sku/ua_markel_eat.jpg", "Маркел", "Рис із рослинним фаршем", 350, 127, 144,
     "СУХПАЙ · UPcompany", SLATE),
    ("sku/ua_veres_pork.jpg", "Верес", "Свинина, горошок, кукурудза", 350, 100, 100,
     "Prom", GREEN),
    ("sku/ua_veres_chicken.jpg", "Верес", "Каша рисова з куркою", 350, 108, 134,
     "Козуб · Смачна адреса · Продукт-Shop", GREEN),
    ("sku/ua_makro_chicken.jpg", "МАКРО", "Курятина 35 % і овочі", 350, 101, 101,
     "СУХПАЙ", ORANGE),
    ("sku/ua_makro_pork.jpg", "МАКРО", "Свинина 35 % і овочі", 350, 101, 135,
     "СУХПАЙ · Euro-komplekt", ORANGE),
    ("sku/ua_makro_beef.jpg", "МАКРО", "Яловичина 35 % і овочі", 350, 106, 107,
     "СУХПАЙ · UPcompany", ORANGE),
    ("sku/ua_khodoriv_pork.jpg", "Ходорівський", "Свинина, горошок, кукурудза", 350, 199, 199,
     "Prom", PLUM),
]
# ── Страва в реторт-паучі, імпорт ─────────────────────────────────────────
IMPORT_RETORT = [
    ("md/clearspring.jpg", "Clearspring", "Brown & Wild Rice, тамарі", 250, 446, 446,
     "Скарби Азії", TEAL),
    ("sku/am_wild.jpg", "Adventure Menu", "Курка в томатному соусі", 400, 315, 357,
     "Freeride · MK-Sport · Лєєр", GREEN),
    ("sku/am_meatballs.jpg", "Adventure Menu", "Тефтелі з басматі", 400, 336, 466,
     "Freeride · OXO · Лєєр · MK · Palmer", GREEN),
    ("sku/am_korma400.jpg", "Adventure Menu", "Chicken Korma з рисом", 400, 376, 404,
     "ALANTUR · Modern Shop · MK-Sport", GREEN),
]

N_SKU = len(BENS) + len(OTTOGI) + len(BOWL_MORE) + len(UA_RETORT) + len(IMPORT_RETORT)
BRANDS = (["Ben's Original", "Ottogi"] +
          [b for _, b, *_ in BOWL_MORE + UA_RETORT + IMPORT_RETORT])
N_BRANDS = len(dict.fromkeys(BRANDS))

# ── Тип паковання ─────────────────────────────────────────────────────────
PACK_BY_BRAND = {"Ben's Original": "pouch", "Ottogi": "cup", "Bibigo": "cup",
                 "Portion": "pouch", "Маркел": "pouch", "Верес": "pouch",
                 "МАКРО": "pouch", "Ходорівський": "pouch",
                 "Clearspring": "pouch", "Adventure Menu": "pouch"}
PACK_RU = {"pouch": "ПАУЧ · РЕТОРТ", "cup": "ЧАША · СТАКАН"}


def pack_of(brand, name=None, grams=None):
    return PACK_BY_BRAND[brand]


# ── Суміжна полиця мереж: те саме споживання, але бляшанка ────────────────
CHAIN_SHELF = [
    ("sku/ch_hapay_rice.jpg", "hapay!", "Каша рисова зі свининою", 340, 62.9, 62.9, "Ашан", GREEN),
    ("sku/ch_lappetit_pork.jpg", "L'appetit", "Каша рисова зі свининою", 340, 119.9, 119.9,
     "Ашан", ORANGE),
    ("sku/ch_lappetit_beef.jpg", "L'appetit", "Каша рисова з яловичиною", 340, 125.9, 125.9,
     "Ашан", ORANGE),
    ("sku/ch_hapay_plov.jpg", "hapay!", "Плов з качки та булгуру", 340, 117, 117, "Ашан", GREEN),
    ("sku/ch_foodfabrika.jpg", "Food Fabrika", "Плов з куркою", 250, 118.9, 118.9, "Восторг", PLUM),
    ("sku/ch_myastoria.jpg", "М'ясторія", "Плов з куркою та родзинками", 350, 147.3, 160,
     "Novus · МегаМаркет · Космос", ROSE),
]

# ── примітиви ─────────────────────────────────────────────────────────────
def rect(s, x, y, w, h, fill, line=None, lw=1.0, rounded=False, adj=0.10):
    sh = s.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE,
        Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid(); sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line; sh.line.width = Pt(lw)
    sh.shadow.inherit = False
    if rounded:
        sh.adjustments[0] = adj
    return sh


def dot(s, cx, cy, d, fill):
    return rect(s, cx - d / 2, cy - d / 2, d, d, fill, rounded=True, adj=0.5)


def text(s, x, y, w, h, body, size=11, bold=False, italic=False, color=INK,
         font=FONT, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, line=None):
    box = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    items = body if isinstance(body, list) else [body]
    for i, item in enumerate(items):
        txt, over = item if isinstance(item, tuple) else (item, {})
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = over.get("align", align)
        if line:
            p.line_spacing = line
        for chunk in (txt if isinstance(txt, list) else [txt]):
            ctxt, cover = chunk if isinstance(chunk, tuple) else (chunk, {})
            r = p.add_run(); r.text = ctxt
            f = r.font
            f.name = cover.get("font", over.get("font", font))
            f.size = Pt(cover.get("size", over.get("size", size)))
            f.bold = cover.get("bold", over.get("bold", bold))
            f.italic = cover.get("italic", over.get("italic", italic))
            f.color.rgb = cover.get("color", over.get("color", color))
    return box


def pic(s, path, x, y, w, h):
    iw, ih = Image.open(path).size
    sc = min(w / iw, h / ih)
    pw, ph = iw * sc, ih * sc
    return s.shapes.add_picture(path, Inches(x + (w - pw) / 2), Inches(y + (h - ph) / 2),
                                Inches(pw), Inches(ph))


def cover_pic(s, path, x, y, w, h):
    iw, ih = Image.open(path).size
    sc = max(w / iw, h / ih)
    pw, ph = iw * sc, ih * sc
    p = s.shapes.add_picture(path, Inches(x), Inches(y), Inches(pw), Inches(ph))
    p.crop_left = p.crop_right = max(0.0, (pw - w) / pw / 2)
    p.crop_top = p.crop_bottom = max(0.0, (ph - h) / ph / 2)
    p.left, p.top, p.width, p.height = Inches(x), Inches(y), Inches(w), Inches(h)
    return p


def slide(dark=False):
    s = prs.slides.add_slide(BLANK)
    rect(s, 0, 0, W, H, NAVY if dark else WHITE)
    PG["n"] += 1
    return s


def header(s, eyebrow, title, dek=None):
    rect(s, 0, 0, W, HDR, NAVY)
    rect(s, 0, HDR, W, 0.055, ORANGE)
    s.shapes.add_picture(LOGO, Inches(M), Inches(0.26), Inches(1.52), Inches(0.66))
    text(s, 2.52, 0.34, 7.40, 0.24, eyebrow, size=9.5, bold=True, color=AMBER)
    text(s, 2.52, 0.60, 8.80, 0.40, title, size=19, bold=True, color=WHITE)
    text(s, W - M - 0.70, 0.44, 0.70, 0.36, f"{PG['n']:02d}", size=17, bold=True,
         color=NAVY_L, align=PP_ALIGN.RIGHT)
    if dek:
        text(s, M, HDR + 0.26, 11.9, 0.34, dek, size=12, color=GREY, line=1.20)


def foot(s, src):
    text(s, M, 7.08, 11.9, 0.26, src, size=8, color=GREY, line=1.14)


def pill(s, x, y, label, bg=GREEN, w=None, size=8.5, fg=WHITE):
    w = w or (0.24 + 0.072 * len(label))
    rect(s, x, y, w, 0.26, bg, rounded=True, adj=0.5)
    text(s, x, y, w, 0.26, label, size=size, bold=True, color=fg,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    return w


def num(v, d=0):
    """1190 → «1 190», 89.99 → «90»; десяткова кома."""
    return f"{v:,.{d}f}".replace(",", "\u00a0").replace(".", ",")


def pl(n, one, few, many):
    """Українська множина: 1 позиція, 3 позиції, 5 позицій."""
    n = int(n)
    if n % 10 == 1 and n % 100 != 11:
        return f"{n} {one}"
    if 2 <= n % 10 <= 4 and not 12 <= n % 100 <= 14:
        return f"{n} {few}"
    return f"{n} {many}"


def rng(lo, hi):
    a, b = num(lo), num(hi)
    return a if a == b else f"{a}–{b}"


def money(lo, hi):
    return num(lo) if lo == hi else f"{num(lo)}–{num(hi)}"


def sku_cell(s, x, y, w, h, img, name, grams, lo, hi, chan, c, brand=None):
    """Клітинка SKU: пакшот, назва, маса, ціна, ціна за 100 г, канал."""
    rect(s, x, y, w, h, WHITE, line=MIST_D, lw=1.0, rounded=True, adj=0.07)
    ph = h - 1.30                      # текстовий блок завжди 1,30″
    rect(s, x + 0.08, y + 0.08, w - 0.16, ph, MIST, rounded=True, adj=0.08)
    if img:
        pic(s, img, x + 0.14, y + 0.12, w - 0.28, ph - 0.08)
    else:
        text(s, x + 0.08, y + ph / 2 - 0.04, w - 0.16, 0.24, "фото немає", size=8,
             color=GREY, align=PP_ALIGN.CENTER)
    yy = y + ph + 0.12
    text(s, x + 0.13, yy, w - 0.26, 0.36, name, size=9, bold=True, color=NAVY, line=1.02)
    gt = f"{num(grams, 0 if float(grams).is_integer() else 1)} г" if grams else "маса н/д"
    if brand:
        text(s, x + 0.13, yy + 0.38, w - 0.26, 0.18,
             [([(brand.upper() + "  ·  ", {"bold": True, "color": c}), gt], {})],
             size=7.5, color=GREY)
    else:
        text(s, x + 0.13, yy + 0.38, w - 0.26, 0.18, gt, size=7.5, color=GREY)
    text(s, x + 0.13, yy + 0.56, w - 0.26, 0.26, money(lo, hi) + " грн", size=11.5,
         bold=True, color=c)
    if grams:
        p100 = rng(lo / grams * 100, hi / grams * 100)
        text(s, x + 0.13, yy + 0.83, w - 0.26, 0.18, p100 + " грн/100 г", size=7.5,
             bold=True, color=GREY)
    text(s, x + 0.13, yy + 1.00, w - 0.26, 0.18, chan[:30], size=7, color=GREY)


def brand_head(s, x, y, brand, origin, fmt, n, lo, hi, plo, phi, c):
    rect(s, x, y, 11.98, 0.66, MIST, rounded=True, adj=0.20)
    rect(s, x, y, 0.09, 0.66, c, rounded=True, adj=0.5)
    text(s, x + 0.30, y + 0.07, 3.00, 0.28, brand, size=14, bold=True, color=NAVY)
    text(s, x + 0.30, y + 0.36, 3.00, 0.22, origin, size=8.5, color=GREY)
    for i, (lb, vl) in enumerate([("ФОРМАТ", fmt), ("ПОЗИЦІЙ", f"{n}"),
                                  ("ЗА УПАКОВКУ", f"{money(plo, phi)} грн"),
                                  ("ЗА 100 Г", f"{rng(lo, hi)} грн")]):
        cx = x + 3.60 + i * 2.10
        text(s, cx, y + 0.09, 2.00, 0.18, lb, size=7.5, bold=True, color=GREY)
        text(s, cx, y + 0.29, 2.00, 0.28, vl, size=11.5, bold=True,
             color=c if i >= 2 else INK)


import collections, math

# ══ РОЗРАХУНКИ ════════════════════════════════════════════════════════════
import collections


def as8(items, brand, c):
    return [(im, brand, nm, g, lo, hi, ch, c) for im, nm, g, lo, hi, ch in items]


L_SIDE = as8(BENS, "Ben's Original", ORANGE)                    # гарнір, пауч
L_BOWL = as8(OTTOGI, "Ottogi", PLUM) + BOWL_MORE                # страва в чаші
L_UA = UA_RETORT                                                # страва, пауч UA
L_IMP = IMPORT_RETORT                                           # страва, пауч імпорт
ALL8 = L_SIDE + L_BOWL + L_UA + L_IMP
assert len(ALL8) == N_SKU

COUNTRY = {"Ben's Original": "ЄС", "Ottogi": "Корея", "Bibigo": "Корея",
           "Portion": "Україна", "Маркел": "Україна", "Верес": "Україна",
           "МАКРО": "Україна", "Ходорівський": "Україна",
           "Clearspring": "ЄС", "Adventure Menu": "ЄС"}
BY_PACK = collections.defaultdict(list)
for _x in ALL8:
    BY_PACK[pack_of(_x[1])].append(_x)
PACK_ORDER = ["pouch", "cup"]
N_UA = sum(1 for x in ALL8 if COUNTRY[x[1]] == "Україна")
N_SILPO = sum(1 for x in ALL8 if "Сільпо" in x[6])
PER100 = sorted(x[4] / x[3] * 100 for x in ALL8)
MED100 = st.median(PER100)
MED_UNIT = st.median([x[4] for x in ALL8])

# ── Імпорт готового рису, CN 1904 90 10 «рис приготовлений» ───────────────
# Дзеркальна статистика: експорт ЄС→Україна (Eurostat Comext DS-045409)
# та експорт Азії→Україна (UN Comtrade). Дзеркало не має прогалин
# української митниці 2019–2021 рр.
IMPORT = [(2019, 99.1, 294), (2020, 204.4, 553), (2021, 204.0, 567), (2022, 127.6, 434),
          (2023, 125.8, 428), (2024, 111.0, 475), (2025, 179.4, 823)]
IMPORT_PARTNERS_2025 = [("Польща", 94.4, 239), ("Болгарія", 78.3, 527), ("Італія", 5.6, 46),
                        ("Корея", 2.0, 10), ("Китай", 1.0, 4)]
EUR_UAH, USD_UAH = 47.15, 41.71          # НБУ, середній за 2025 р.
T_EU, V_EU = IMPORT[-1][1], IMPORT[-1][2]
T_ASIA = 3.1
T_TOTAL = T_EU + T_ASIA
T_PREV = IMPORT[-2][1] + 3.5
CIF_UAH = (V_EU * EUR_UAH + 14.0 * USD_UAH) / 1000            # млн грн
AVG_PACK_G = 300                          # медіана маси в межах дослідження
PACKS_K = T_TOTAL * 1000_000 / AVG_PACK_G / 1000              # тис. упаковок на рік
RETAIL_LO, RETAIL_HI = CIF_UAH * 2.2, CIF_UAH * 2.8           # роздріб, млн грн
POP_M = 29.0
EU_PER_CAP_USD = 1.43                     # ринок instant rice ЄС / населення ЄС
UA_PER_CAP_USD = RETAIL_LO * 1e6 / USD_UAH / (POP_M * 1e6)
GAP = EU_PER_CAP_USD / UA_PER_CAP_USD
GROWTH = (T_TOTAL / T_PREV - 1) * 100

import math

# ── примітиви ─────────────────────────────────────────────────────────────
def kpi(s, x, y, w, h, label, value, note, c):
    rect(s, x, y, w, h, MIST, rounded=True, adj=0.08)
    rect(s, x, y, w, 0.09, c, rounded=True, adj=0.5)
    text(s, x + 0.24, y + 0.26, w - 0.48, 0.22, label, size=8.5, bold=True, color=GREY)
    text(s, x + 0.24, y + 0.50, w - 0.48, 0.44, value, size=23, bold=True, color=c)
    text(s, x + 0.24, y + 1.00, w - 0.48, h - 1.08, note, size=9, color=INK, line=1.24)


def insight(s, x, y, w, h, title, body, c=ORANGE):
    rect(s, x, y, w, h, NAVY, rounded=True, adj=0.06)
    rect(s, x, y, 0.09, h, c, rounded=True, adj=0.5)
    if title:
        text(s, x + 0.34, y + 0.18, w - 0.60, 0.28, title, size=12.5, bold=True, color=AMBER)
    text(s, x + 0.34, y + (0.54 if title else 0.22), w - 0.60,
         h - (0.70 if title else 0.40), body, size=10,
         color=RGBColor(0xD5, 0xDB, 0xEA), line=1.30)


def grid(s, items, y, w, h, x0=M, gap=0.10, brand=True):
    for i, (im, b, nm, g, lo, hi, ch, c) in enumerate(items):
        sku_cell(s, x0 + i * (w + gap), y, w, h, im, nm, g, lo, hi, ch, c,
                 brand=b if brand else None)


def brands_of(items):
    return list(dict.fromkeys(x[1] for x in items))


def rng_unit(items):
    return min(x[4] for x in items), max(x[5] for x in items)


def p100(items):
    v = [(lo / g * 100, hi / g * 100) for *_, g, lo, hi, _, _ in items if g]
    return min(a for a, _ in v), max(b for _, b in v)


def med_unit(items):
    return st.median([x[4] for x in items])


def med100(items):
    return st.median([x[4] / x[3] * 100 for x in items])


SEGMENTS = [("ГАРНІР У ПАУЧІ", L_SIDE, ORANGE,
             "Чистий рис без наповнювача. Мікрохвильовка 90 с."),
            ("СТРАВА В ЧАШІ", L_BOWL, PLUM,
             "Рис із м'ясом у жорсткій чаші. Мікрохвильовка 2 хв."),
            ("СТРАВА В ПАУЧІ · УКРАЇНА", L_UA, GREEN,
             "Рис із м'ясом у реторт-паучі. Розігрів у воді або НВЧ."),
            ("СТРАВА В ПАУЧІ · ІМПОРТ", L_IMP, TEAL,
             "Те саме, але завезене: ЄС, органіка й туристична лінійка.")]


# ══ 01 · ОБКЛАДИНКА ═══════════════════════════════════════════════════════
def slide_cover():
    s = slide(dark=True)
    cover_pic(s, BG_TITLE, 0, 0, W, H)
    s.shapes.add_picture(LOGO, Inches(M), Inches(0.52), Inches(1.94), Inches(0.84))
    rect(s, M, 2.02, 0.54, 0.09, ORANGE)
    text(s, M, 2.34, 7.00, 0.34, "ДОСЛІДЖЕННЯ РИНКУ · ВЕРЕСЕНЬ 2026", size=11.5,
         bold=True, color=AMBER)
    text(s, M, 2.78, 7.20, 1.70, "Готовий рис\nв Україні", size=46, bold=True,
         color=WHITE, line=1.04)
    text(s, M, 4.52, 7.00, 0.36, "Кімнатне зберігання · розігрів · без готування",
         size=14.5, color=RGBColor(0xC6, 0xCE, 0xE2))
    rect(s, M, 5.12, 5.90, 0.055, RGBColor(0x55, 0x60, 0x8C))
    text(s, M, 5.34, 6.90, 0.86,
         f"{pl(N_BRANDS, 'бренд', 'бренди', 'брендів')} · "
         f"{pl(N_SKU, 'позиція', 'позиції', 'позицій')} · ціна за одну упаковку\n"
         "Кожну ціну перевірено на сайті продавця 23 вересня 2026 р.",
         size=11.5, color=RGBColor(0x9F, 0xA9, 0xC4), line=1.40)
    for im, x, y, w, h in [("md/bens_basmati250.jpg", 8.26, 0.96, 2.35, 2.40),
                           ("sku/ot_bibimbap.jpg", 10.78, 0.96, 2.35, 2.40),
                           ("sku/ua_veres_chicken.jpg", 8.26, 3.50, 2.35, 2.10),
                           ("sku/ua_makro_beef.jpg", 10.78, 3.50, 2.35, 2.10)]:
        rect(s, x, y, w, h, WHITE, rounded=True, adj=0.06)
        pic(s, im, x + 0.16, y + 0.14, w - 0.32, h - 0.28)
    rect(s, 8.26, 5.76, 4.87, 1.02, RGBColor(0x26, 0x2F, 0x4D), rounded=True, adj=0.14)
    text(s, 8.50, 5.92, 4.44, 0.78, " · ".join(dict.fromkeys(x[1] for x in ALL8)),
         size=9.5, bold=True, color=WHITE, line=1.32)


# ══ 02 · ГОЛОВНЕ ══════════════════════════════════════════════════════════
def slide_summary():
    s = slide()
    header(s, "ГОЛОВНЕ", "Шість цифр про категорію",
           "Предмет: рис, що зберігається за кімнатної температури і лише розігрівається.")
    K = [("ОБСЯГ ІМПОРТУ 2025", f"{num(T_TOTAL)} т", ORANGE,
          f"Готовий рис, CN 1904 90 10.\n+{num(GROWTH)} % до 2024 року."),
         ("ЄМНІСТЬ РОЗДРІБУ", f"{num(RETAIL_LO)}–{num(RETAIL_HI)}", TEAL,
          "млн грн на рік — оцінка\nза імпортом × 2,2–2,8."),
         ("ВІДСТАВАННЯ ВІД ЄС", f"×{num(GAP)}", ROSE,
          f"{num(UA_PER_CAP_USD, 2)} $ на особу проти\n1,43 $ у ЄС."),
         ("ПОЗИЦІЙ У ПРОДАЖУ", f"{N_SKU}", PLUM,
          f"{N_BRANDS} брендів. {N_UA} позицій —\nукраїнського виробництва."),
         ("МЕДІАНА ЗА УПАКОВКУ", f"{num(MED_UNIT)} грн", GREEN,
          f"За 100 г — {num(MED100)} грн.\nДіапазон 45–446 грн."),
         ("МАСОВА ВАГА", "350 г", SLATE,
          "10 позицій. Далі 250 г — 6,\n220 г — 6.")]
    for i, (lb, v, c, note) in enumerate(K):
        kpi(s, M + i * 2.04, 1.88, 1.86, 1.92, lb, v, note, c)
    insight(s, M, 4.02, 5.86, 2.36, "Що показало дослідження",
            "1. Українське виробництво в категорії вже є — і воно найдешевше за грам: "
            "реторт-пауч 350 г за 71–199 грн, це 20–57 грн за 100 г.\n"
            "2. Ben's Original — єдиний імпортний гарнір у мережі: 10 позицій у «Сільпо» "
            "за 45–179 грн.\n"
            "3. Корейська чаша — найдорожчий формат за грам (91 грн/100 г) і продається "
            "лише у трьох азійських магазинах.", ORANGE)
    insight(s, M + 6.12, 4.02, 5.86, 2.36, "Де порожньо",
            "1. Немає українського гарніру — чистого рису в паучі. Усі 10 українських "
            "позицій це рис із м'ясом, тобто повноцінна страва.\n"
            "2. Немає чаші дешевше 189 грн: Bibigo 189, Ottogi 225–356.\n"
            "3. Немає жодної позиції категорії в АТБ, Novus, Metro, Varus, Ашан, Fozzy "
            "та ще 11 мережах — перевірено через API 17 мереж.", TEAL)
    foot(s, "Ціни: сайти продавців, 23.09.2026, за одну упаковку в роздріб. "
            "Імпорт: Eurostat Comext CN 1904 90 10 + UN Comtrade. "
            "Ємність роздрібу — оцінка, метод на слайді 03.")


# ══ 03 · ОБСЯГ РИНКУ ══════════════════════════════════════════════════════
def slide_volume():
    s = slide()
    header(s, "ОБСЯГ РИНКУ",
           f"{num(T_TOTAL)} тонн імпорту, {num(RETAIL_LO)}–{num(RETAIL_HI)} млн грн роздрібу",
           "Імпорт готового рису відновився: 2025 рік — найбільший за сім років за вартістю.")
    rect(s, M, 1.88, 7.30, 3.34, MIST, rounded=True, adj=0.05)
    text(s, M + 0.28, 2.06, 6.00, 0.24, "ІМПОРТ ГОТОВОГО РИСУ З ЄС, ТОНН НА РІК",
         size=9, bold=True, color=GREY)
    bx, by0, bh = M + 0.34, 4.56, 2.02
    mx = max(t for _, t, _ in IMPORT)
    for i, (yr, t, v) in enumerate(IMPORT):
        x = bx + i * 1.00
        h = bh * t / mx
        c = ORANGE if yr == 2025 else (NAVY_L if yr >= 2022 else RGBColor(0xC2, 0xC9, 0xDA))
        rect(s, x, by0 - h, 0.86, h, c, rounded=True, adj=0.10)
        text(s, x - 0.06, by0 - h - 0.24, 0.98, 0.22, num(t), size=8.5, bold=True,
             color=c if yr == 2025 else GREY, align=PP_ALIGN.CENTER)
        text(s, x - 0.06, by0 + 0.06, 0.98, 0.22, str(yr), size=8.5,
             bold=(yr == 2025), color=INK, align=PP_ALIGN.CENTER)
        text(s, x - 0.06, by0 + 0.26, 0.98, 0.20, f"€{num(v)}k", size=7.5, color=GREY,
             align=PP_ALIGN.CENTER)
    text(s, M + 0.34, 4.98, 7.00, 0.20,
         "Вартість 2025 року — 823 тис. € проти 475 тис. € у 2024-му, +73 %.",
         size=8.5, color=GREY)
    rect(s, M + 7.54, 1.88, 4.44, 3.34, WHITE, line=MIST_D, lw=1.2, rounded=True, adj=0.05)
    text(s, M + 7.80, 2.06, 3.92, 0.24, "ЗВІДКИ ЇДЕ, 2025 РІК", size=9, bold=True, color=GREY)
    yy = 2.36
    tot = sum(t for _, t, _ in IMPORT_PARTNERS_2025)
    for nm, t, v in IMPORT_PARTNERS_2025:
        w = 2.90 * t / tot
        text(s, M + 7.80, yy, 1.00, 0.22, nm, size=9, color=INK)
        rect(s, M + 8.80, yy + 0.05, max(w, 0.04), 0.13, ORANGE if t > 50 else TEAL,
             rounded=True, adj=0.5)
        text(s, M + 8.80 + max(w, 0.04) + 0.08, yy - 0.01, 0.92, 0.22, f"{num(t, 1)} т",
             size=8.5, bold=True, color=GREY)
        yy += 0.34
    text(s, M + 7.80, yy + 0.06, 3.92, 0.76,
         "Польща й Болгарія — 95 % тонажу: це заводи ЄС, що пакують рис у пауч.\n"
         "Корея і Китай разом 3 т — чаші Ottogi та Bibigo. Україна свій обсяг "
         "виробляє на місці, в імпорт він не входить.",
         size=8.5, color=GREY, line=1.26)
    rect(s, M, 5.26, 11.98, 1.44, NAVY, rounded=True, adj=0.07)
    rect(s, M, 5.26, 0.09, 1.44, TEAL, rounded=True, adj=0.5)
    text(s, M + 0.34, 5.42, 3.90, 0.26, "Як рахували ємність", size=12.5, bold=True, color=AMBER)
    STEPS = [(f"{num(T_TOTAL)} т", "імпорт 2025"),
             (f"{num(CIF_UAH)} млн грн", "CIF за курсом НБУ"),
             ("× 2,2–2,8", "мито 0 %, ПДВ, маржа"),
             (f"{num(RETAIL_LO)}–{num(RETAIL_HI)} млн", "роздріб на рік"),
             (f"≈ {num(PACKS_K)} тис.", "упаковок на рік")]
    for i, (v, lb) in enumerate(STEPS):
        x = M + 4.30 + i * 1.54
        text(s, x, 5.42, 1.48, 0.28, v, size=12, bold=True, color=WHITE)
        text(s, x, 5.70, 1.48, 0.36, lb, size=8, color=RGBColor(0x9F, 0xA9, 0xC4), line=1.20)
        if i < len(STEPS) - 1:
            text(s, x + 1.32, 5.44, 0.20, 0.24, "→", size=12, bold=True, color=ORANGE)
    text(s, M + 0.34, 5.78, 3.76, 0.80,
         "Це лише імпортна частина. Український реторт-пауч виробляють у країні, "
         "і в митну статистику він не потрапляє.",
         size=8.5, color=RGBColor(0xB9, 0xC2, 0xDA), line=1.26)
    foot(s, "Джерела: Eurostat Comext DS-045409, CN 1904 90 10 «рис приготовлений», "
            "експорт ЄС→Україна 2019–2025; UN Comtrade (Корея, Китай) за 2025 р.; "
            "курс НБУ, середній за 2025 р.: 47,15 грн/€. Ємність роздрібу — ОЦІНКА.")


# ══ 04 · ЧОТИРИ СЕГМЕНТИ ══════════════════════════════════════════════════
def slide_map():
    s = slide()
    header(s, "КАРТА КАТЕГОРІЇ", "Чотири сегменти, десять брендів",
           "Усередині категорії — два формати паковання і дві ролі: гарнір або повна страва.")
    for i, (kind, items, c, note) in enumerate(SEGMENTS):
        x = M + i * 3.07
        lo, hi = rng_unit(items)
        br = brands_of(items)
        rect(s, x, 1.90, 2.82, 3.06, MIST, rounded=True, adj=0.06)
        rect(s, x, 1.90, 2.82, 0.10, c, rounded=True, adj=0.5)
        text(s, x + 0.22, 2.10, 2.40, 0.24, kind, size=9.5, bold=True, color=c)
        text(s, x + 0.22, 2.38, 2.40, 0.32,
             f"{pl(len(items), 'позиція', 'позиції', 'позицій')}", size=14, bold=True, color=NAVY)
        text(s, x + 0.22, 2.72, 2.40, 0.62, " · ".join(br), size=9, color=INK, line=1.22)
        text(s, x + 0.22, 3.38, 2.40, 0.44, note, size=8.5, color=GREY, line=1.20)
        rect(s, x + 0.22, 3.88, 2.38, 0.02, MIST_D)
        text(s, x + 0.22, 3.98, 2.40, 0.20, "ЗА УПАКОВКУ · МЕДІАНА", size=7,
             bold=True, color=GREY)
        text(s, x + 0.22, 4.16, 2.40, 0.28, f"{num(med_unit(items))} грн", size=15,
             bold=True, color=c)
        text(s, x + 0.22, 4.44, 2.40, 0.20, f"{num(lo)} – {num(hi)} грн", size=8.5, color=GREY)
        text(s, x + 0.22, 4.64, 2.40, 0.20, f"за 100 г — {num(med100(items))} грн",
             size=8.5, bold=True, color=INK)
    insight(s, M, 5.14, 11.98, 1.42, "Як читати цю карту",
            "Український реторт-пауч уже виграє за грам: 29 грн за 100 г проти 41 грн у "
            "Ben's Original і 91 грн у корейської чаші. Але це страва з м'ясом, а не гарнір — "
            "інша полиця і інший привід.\n"
            "Ben's Original тримає нішу гарніру сам-один: жодного українського чистого рису "
            "в паучі на ринку немає. Корейська чаша — найдорожчий формат і найвужчий канал: "
            "три азійські магазини.", ORANGE)
    foot(s, f"{N_SKU} позицій, ціни за одну упаковку, 23.09.2026.")

# ══ СЛАЙДИ БРЕНДІВ ════════════════════════════════════════════════════════
def sku_slide(eyebrow, title, dek, items, src, head=None, two_rows=True, note=None):
    s = slide()
    header(s, eyebrow, title, dek)
    y0 = 1.72
    if head:
        brand, origin, fmt, c, plo, phi = head
        lo100, hi100 = p100(items)
        brand_head(s, M, y0, brand, origin, fmt, len(items), lo100, hi100, plo, phi, c)
        y0 = 2.50
    else:
        y0 = 1.88
    extra = 2 if note else 0
    n = len(items)
    if two_rows:
        cols = math.ceil((n + extra) / 2)
        w = (11.98 - 0.10 * (cols - 1)) / cols
        h = 2.22 if head else 2.34
        gap2 = (h + 0.06) if head else (h + 0.10)
        grid(s, items[:cols], y0, w, h)
        rest = items[cols:]
        grid(s, rest, y0 + gap2, w, h)
        if note:
            x0 = M + len(rest) * (w + 0.10)
            insight(s, x0, y0 + gap2, M + 11.98 - x0, h, note[0], note[1], note[2])
    else:
        w = (11.98 - 0.10 * (n - 1)) / n
        grid(s, items, y0, w, 3.44)
        if note:
            insight(s, M, 5.48, 11.98, 1.42, note[0], note[1], note[2])
    foot(s, src)


def slide_bens():
    sku_slide("СЕГМЕНТ 1 · ГАРНІР У ПАУЧІ",
              f"Ben's Original — {pl(len(BENS), 'позиція', 'позиції', 'позицій')}",
              "Єдиний бренд чистого готового рису на ринку. Мікрохвильовка 90 секунд.",
              as8(BENS, "Ben's Original", ORANGE),
              "Ціни каталогу «Сільпо», отримані поштучно з картки кожного товару 23.09.2026; "
              "Long Grain 220 г — Edison Lee, у картці зазначено «ціну вказано за одну порцію». "
              "Позицію Sticky Bowl 220 г із попередньої редакції знято: у каталозі її немає.",
              head=("Ben's Original", "Mars · Франція, Німеччина", "пауч", ORANGE, 45, 179))


def slide_bowls():
    sku_slide("СЕГМЕНТ 2 · СТРАВА В ЧАШІ",
              "Ottogi і Bibigo — вісім позицій",
              "Рис із наповнювачем у жорсткій чаші. Корея, канал — азійські фудшопи.",
              L_BOWL,
              "Ціни з карток Prom.ua 23.09.2026, роздрібні за одну чашу. У «Тайякі Март» "
              "діє нижча ціна від 15–20 шт — вона до розрахунку не бралася. Три позиції "
              "Ottogi з попередньої редакції знято: на сайтах продавців їх більше немає.",
              note=("Що це означає",
                    "Найдорожчий формат категорії: 91 грн за 100 г проти 41 у Ben's і 29 в "
                    "українському паучі. Та сама позиція коштує по-різному: тунець 247 г — "
                    "225 грн у «Тайякі Март» і 356 грн в «Апетітаріум», розкид 58 %.", PLUM))


def slide_ua():
    sku_slide("СЕГМЕНТ 3 · СТРАВА В ПАУЧІ · УКРАЇНА",
              f"{pl(len(L_UA), 'позиція', 'позиції', 'позицій')} п'яти українських виробників",
              "Рис із м'ясом у реторт-паучі 350 г. Зберігання 24 місяці, розігрів без готування.",
              L_UA,
              "Ціни з карток продавців на Prom.ua і з фірмового магазину portion.com.ua, "
              "23.09.2026, за один пауч. Оптові пороги від 3–12 шт не бралися.",
              note=("Що це означає",
                    "Найдешевший грам усієї категорії: 20–57 грн за 100 г. Виробництво "
                    "українське, тому ці позиції не потрапляють в імпортну статистику "
                    "на слайді 03 — реальний обсяг категорії більший за 182 тонни.\n"
                    "Канал вузький: Prom і фірмові магазини, у мережах цих брендів немає.",
                    GREEN))


def slide_import():
    sku_slide("СЕГМЕНТ 4 · СТРАВА В ПАУЧІ · ІМПОРТ",
              "Clearspring і Adventure Menu — чотири позиції",
              "Той самий формат, але завезений: британська органіка й чеська туристична лінійка.",
              L_IMP, "Ціни з карток Prom.ua 23.09.2026 за одну упаковку.",
              two_rows=False,
              note=("Що це означає",
                    "Adventure Menu 400 г READY TO EAT — реторт без води, 79–101 грн за 100 г: "
                    "той самий коридор, що Ben's Original, але канал туристичний. "
                    "Clearspring 250 г за 446 грн — 178 грн за 100 г, найдорожчий пауч ринку: "
                    "органічна сертифікація й дикий рис.", TEAL))

# ══ ФОРМАТ І ГРАМАЖ ═══════════════════════════════════════════════════════
def slide_format():
    s = slide()
    header(s, "ФОРМАТ І ГРАМАЖ", "Пауч — 25 позицій із 33, масова вага 350 г",
           "Два формати паковання. Кожен тримає свою вагу — платформи майже не перетинаються.")
    PICS = {"pouch": "sku/ua_veres_chicken.jpg", "cup": "sku/ot_bibimbap.jpg"}
    NOTE = {"pouch": "Плаский реторт-пауч. Займає менше місця на полиці, дешевший у логістиці, "
                     "тримає 24 місяці. Формат, у якому працюють і Ben's, і всі українці.",
            "cup": "Жорстка чаша: їсти можна прямо з неї, але вона дорожча й об'ємніша. "
                   "На українському ринку — тільки корейський імпорт."}
    COL = {"pouch": ORANGE, "cup": PLUM}
    for i, k in enumerate(PACK_ORDER):
        it = BY_PACK[k]
        x = M + i * 3.07
        c = COL[k]
        pr = sorted(y[4] for y in it)
        ms = sorted(y[3] for y in it)
        rect(s, x, 1.88, 2.82, 4.94, WHITE, line=MIST_D, lw=1.2, rounded=True, adj=0.05)
        rect(s, x, 1.88, 2.82, 0.10, c, rounded=True, adj=0.5)
        rect(s, x + 0.16, 2.08, 2.50, 1.54, MIST, rounded=True, adj=0.06)
        pic(s, PICS[k], x + 0.26, 2.14, 2.30, 1.42)
        text(s, x + 0.24, 3.72, 2.40, 0.24, PACK_RU[k], size=10.5, bold=True, color=c)
        rect(s, x + 0.24, 4.02, 2.34, 0.18, MIST_D, rounded=True, adj=0.5)
        rect(s, x + 0.24, 4.02, 2.34 * len(it) / N_SKU, 0.18, c, rounded=True, adj=0.5)
        text(s, x + 0.24, 4.26, 2.40, 0.24,
             f"{pl(len(it), 'позиція', 'позиції', 'позицій')} · "
             f"{num(len(it) / N_SKU * 100)} % · "
             f"{pl(len(set(y[1] for y in it)), 'бренд', 'бренди', 'брендів')}",
             size=10.5, bold=True, color=NAVY)
        rect(s, x + 0.24, 4.58, 2.34, 0.02, MIST_D)
        for j, (lb, v) in enumerate([("ЗА УПАКОВКУ, МЕДІАНА", f"{num(st.median(pr))} грн"),
                                     ("ЗА 100 Г, МЕДІАНА", f"{num(med100(it))} грн"),
                                     ("ДІАПАЗОН", f"{num(min(pr))} – {num(max(pr))} грн"),
                                     ("МАСА", f"{num(min(ms))} – {num(max(ms))} г")]):
            yy = 4.68 + j * 0.46
            text(s, x + 0.24, yy, 1.90, 0.20, lb, size=7, bold=True, color=GREY)
            text(s, x + 0.24, yy + 0.16, 2.40, 0.22, v, size=10.5, bold=True,
                 color=c if j < 2 else INK)
        text(s, x + 0.24, 6.46, 2.40, 0.30, NOTE[k], size=7.5, color=GREY, line=1.22)
    # грамаж
    cnt = collections.Counter(x[3] for x in ALL8)
    tops = sorted(cnt.items(), key=lambda kv: kv[0])
    text(s, 6.88, 1.88, 6.00, 0.24, "ПОЗИЦІЙ У КОЖНІЙ ВАЗІ", size=9, bold=True, color=GREY)
    bx, by0 = 6.88, 4.46
    mxn = max(n for _, n in tops)
    bw = (5.78 - 0.10 * (len(tops) - 1)) / len(tops)
    for i, (g_, n) in enumerate(tops):
        x = bx + i * (bw + 0.10)
        h = 1.96 * n / mxn
        pk = collections.Counter(pack_of(y[1]) for y in ALL8 if y[3] == g_)
        c = COL[pk.most_common(1)[0][0]]
        rect(s, x, by0 - h, bw, h, c, rounded=True, adj=0.14)
        text(s, x, by0 - h - 0.24, bw, 0.22, str(n), size=11, bold=True, color=c,
             align=PP_ALIGN.CENTER)
        text(s, x, by0 + 0.06, bw, 0.22, f"{num(g_)} г", size=8, bold=True, color=INK,
             align=PP_ALIGN.CENTER)
    text(s, 6.88, 4.76, 5.90, 0.44,
         "350 г — український реторт-пауч. 250 і 220 г — Ben's Original. "
         "210–320 г — корейська чаша. 400 г — Adventure Menu.",
         size=8.5, color=GREY, line=1.24)
    insight(s, 6.88, 5.40, 5.90, 1.42, None,
            "Український пауч стоїть у вазі 350 г, Ben's Original — у 220 і 250 г, "
            "корейська чаша — у 210–320 г. Ваги майже не перетинаються.\n"
            "Вільне вікно — пауч 200–250 г українського виробництва: у цій вазі сьогодні "
            "тільки імпортний Ben's Original за 45–179 грн.", ORANGE)
    foot(s, f"Розподіл {N_SKU} позицій за типом паковання і масою з карток продавців, 23.09.2026.")


# ══ ЦІНА ЗА ОДНУ УПАКОВКУ ═════════════════════════════════════════════════
def slide_ladder():
    s = slide()
    header(s, "ЦІНА", "Кожна позиція: за упаковку і за 100 грамів",
           "Ліворуч — скільки покупець платить у касі. Праворуч — скільки це коштує за грам.")
    rows = sorted(ALL8, key=lambda x: x[4])
    rect(s, M, 1.84, 11.98, 5.06, MIST, rounded=True, adj=0.04)
    AX1, AW1 = 4.34, 3.10
    AX2, AW2 = 8.74, 2.90
    MX1, MX2 = 460.0, 190.0

    def a1(v):
        return AX1 + AW1 * v / MX1

    def a2(v):
        return AX2 + AW2 * v / MX2
    for t in range(0, 461, 115):
        rect(s, a1(t), 2.14, 0.012, 4.24, RGBColor(0xDC, 0xE2, 0xEE))
        text(s, a1(t) - 0.34, 6.46, 0.68, 0.20, num(t), size=7.5, color=GREY,
             align=PP_ALIGN.CENTER)
    for t in range(0, 191, 50):
        rect(s, a2(t), 2.14, 0.012, 4.24, RGBColor(0xDC, 0xE2, 0xEE))
        text(s, a2(t) - 0.34, 6.46, 0.68, 0.20, num(t), size=7.5, color=GREY,
             align=PP_ALIGN.CENTER)
    rect(s, a2(MED100), 2.14, 0.02, 4.24, NAVY_L)
    text(s, M + 0.24, 1.98, 2.60, 0.18, "БРЕНД І ПОЗИЦІЯ", size=6.5, bold=True, color=GREY)
    text(s, M + 2.92, 1.98, 0.46, 0.18, "МАСА", size=6.5, bold=True, color=GREY,
         align=PP_ALIGN.RIGHT)
    text(s, AX1, 1.98, 2.20, 0.18, "ГРН ЗА УПАКОВКУ", size=6.5, bold=True, color=GREY)
    text(s, AX2, 1.98, 2.20, 0.18, "ГРН ЗА 100 Г", size=6.5, bold=True, color=GREY)
    yy = 2.30
    for im, b, nm, g, lo, hi, ch, c in rows:
        label = f"{b} · {nm}"
        text(s, M + 0.24, yy - 0.075, 2.62, 0.17,
             [([(b + " · ", {"bold": True, "color": NAVY}), (nm, {"color": GREY})], {})],
             size=7.5)
        text(s, M + 2.86, yy - 0.07, 0.52, 0.17, f"{num(g)} г", size=7, color=GREY,
             align=PP_ALIGN.RIGHT)
        rect(s, a1(lo), yy - 0.03, max(a1(hi) - a1(lo), 0.03), 0.06, c, rounded=True, adj=0.5)
        dot(s, a1(lo), yy, 0.115, c)
        lbl = num(lo) if lo == hi else f"{num(lo)}–{num(hi)}"
        text(s, a1(hi) + 0.07, yy - 0.085, 0.80, 0.18, lbl, size=7, bold=True, color=c)
        p1, p2 = lo / g * 100, hi / g * 100
        rect(s, a2(p1), yy - 0.03, max(a2(p2) - a2(p1), 0.03), 0.06, c, rounded=True, adj=0.5)
        dot(s, a2(p1), yy, 0.115, c)
        l2 = num(p1) if lo == hi else f"{num(p1)}–{num(p2)}"
        text(s, a2(p2) + 0.07, yy - 0.085, 0.80, 0.18, l2, size=7, bold=True, color=c)
        yy += 0.126
    text(s, a2(MED100) - 0.70, 6.64, 1.60, 0.20, f"медіана {num(MED100)} грн", size=7.5,
         bold=True, color=NAVY_L, align=PP_ALIGN.CENTER)
    foot(s, "Роздрібна ціна за одну упаковку на сайті продавця, 23.09.2026. Де продавців "
            "кілька — показано діапазон. Оптові пороги від 3–20 шт у розрахунок не входять.")


# ══ ПОЛИЦЯ МЕРЕЖ ══════════════════════════════════════════════════════════
def slide_chain_shelf():
    s = slide()
    header(s, "ЩО СТОЇТЬ У МЕРЕЖІ ЗАМІСТЬ НАС", "Рис із м'ясом на полиці є — але в бляшанці",
           "Мережі не мають жодного пауча з готовим рисом. Натомість мають консерву за 63–160 грн.")
    for i, (im, b, nm, g, lo, hi, ch, c) in enumerate(CHAIN_SHELF):
        sku_cell(s, M + i * 2.02, 1.88, 1.92, 2.86, im, nm, g, lo, hi, ch, c, brand=b)
    insight(s, M, 4.94, 5.86, 1.94, "Чому це важливо",
            "Це прямий конкурент за той самий привід: гаряча страва з рисом і м'ясом без "
            "готування, уже на полиці мережі, з українським виробником.\n"
            "Бляшанку не можна поставити в мікрохвильовку — у цьому перевага пауча. "
            "Але орієнтир ціни задає саме вона, а не Ottogi за 252 грн.", ROSE)
    rect(s, M + 6.12, 4.94, 5.86, 1.94, MIST, rounded=True, adj=0.06)
    text(s, M + 6.40, 5.10, 5.30, 0.26, "Порівняння за 100 г", size=12, bold=True, color=NAVY)
    CMP = [("Каша рисова hapay! 340 г · бляшанка", 18.5, GREEN),
           ("Український пауч 350 г · медіана", 29.0, ORANGE),
           ("Ben's Original · медіана", 41.0, PLUM),
           ("Плов М'ясторія 350 г · лоток", 45.7, ROSE),
           ("Ottogi · медіана", 91.0, SLATE)]
    yy = 5.44
    mxv = max(v for _, v, _ in CMP)
    for lb, v, c in CMP:
        text(s, M + 6.40, yy - 0.04, 2.70, 0.20, lb, size=8.5, color=INK)
        rect(s, M + 9.20, yy + 0.01, 1.90 * v / mxv, 0.13, c, rounded=True, adj=0.5)
        text(s, M + 9.20 + 1.90 * v / mxv + 0.08, yy - 0.05, 0.80, 0.20, f"{num(v, 1)} грн",
             size=8, bold=True, color=c)
        yy += 0.28
    foot(s, "Джерело: API 17 мереж zakaz.ua, 23.09.2026. Ці шість позицій — суміжна "
            f"категорія, у {N_SKU} позицій дослідження вони не входять: бляшанку не "
            "розігрівають у мікрохвильовці.")


# ══ КАНАЛИ ════════════════════════════════════════════════════════════════
def slide_channels():
    s = slide()
    header(s, "ДЕ ПРЕДСТАВЛЕНО", "Три типи каналів, 16 продавців",
           "Мережа є лише в одного бренду. Українське виробництво в мережі не представлене взагалі.")
    GROUPS = [("МЕРЕЖЕВИЙ РОЗДРІБ", ORANGE, ["Сільпо"],
               "Ben's Original — 10 позицій. Інших брендів категорії немає."),
              ("АЗІЙСЬКІ ФУДШОПИ", PLUM,
               ["Тайякі Март", "Апетітаріум", "Gurmissimo", "Скарби Азії", "Edison Lee"],
               "Ottogi, Bibigo, Clearspring. Вузький канал, висока ціна."),
              ("PROM ТА ФІРМОВІ МАГАЗИНИ", GREEN,
               ["portion.com.ua", "СУХПАЙ", "UPcompany", "Мартел-shop", "Козуб Маркет",
                "Смачна адреса", "Продукт-Shop", "Euro-komplekt", "ALANTUR", "MK-Sport"],
               "Увесь український реторт-пауч і Adventure Menu.")]
    yy = 1.90
    for gi, (title, c, chans, what) in enumerate(GROUPS):
        rh = 0.86 if gi < 2 else 1.14
        rect(s, M, yy, 11.98, rh, MIST if gi % 2 == 0 else WHITE, rounded=True, adj=0.12)
        rect(s, M, yy, 0.09, rh, c, rounded=True, adj=0.5)
        text(s, M + 0.30, yy + 0.14, 3.70, 0.22, title, size=9.5, bold=True, color=c)
        text(s, M + 0.30, yy + 0.40, 3.70, 0.56, what, size=8.5, color=GREY, line=1.20)
        cx, cy = M + 4.20, yy + 0.16
        for ch in chans:
            w = 0.24 + 0.068 * len(ch)
            if cx + w > M + 11.10:
                cx, cy = M + 4.20, cy + 0.34
            pill(s, cx, cy, ch, c, w=w, size=8.5)
            cx += w + 0.12
        text(s, M + 11.26, yy + rh / 2 - 0.17, 0.56, 0.34, str(len(chans)), size=16,
             bold=True, color=c, align=PP_ALIGN.RIGHT)
        yy += rh + 0.10
    rect(s, M, 5.14, 5.86, 1.02, MIST, rounded=True, adj=0.10)
    text(s, M + 0.26, 5.28, 5.34, 0.24, "Мережі без жодної позиції категорії",
         size=11, bold=True, color=NAVY)
    text(s, M + 0.26, 5.54, 5.34, 0.56,
         "АТБ, Novus, Metro, Varus, Ашан, Fozzy, «Таврія В», МегаМаркет, ЕКО маркет, "
         "Ultramarket та ще 7 мереж — перевірено через API zakaz.ua 23.09.2026.",
         size=8.5, color=INK, line=1.24)
    rect(s, M + 6.12, 5.14, 5.86, 1.02, MIST, rounded=True, adj=0.10)
    text(s, M + 6.38, 5.28, 5.34, 0.24, "Що відомо про походження", size=11, bold=True, color=NAVY)
    text(s, M + 6.38, 5.54, 5.34, 0.56,
         "Ben's Original виробляє Mars, заводи у Франції та Німеччині. Український пауч — "
         "Верес, Ходорівський м'ясокомбінат, «Маркел», «МАКРО», Portion. Імпортера в "
         "картках продавців не вказано.",
         size=8.5, color=INK, line=1.24)
    insight(s, M, 6.28, 11.98, 0.60, None,
            "Головна дірка каналу: український виробник робить найдешевший за грам продукт "
            "категорії — і не має жодної мережевої полиці.", GREEN)


# ══ ВИСНОВКИ ══════════════════════════════════════════════════════════════
def slide_conclusions():
    s = slide()
    header(s, "ВИСНОВКИ", "Чотири факти і три обмеження",
           "Усе нижче спирається на ціни, перевірені поштучно 23 вересня 2026 року.")
    FACTS = [("Категорія мала, але росте", ORANGE,
              f"{num(T_TOTAL)} т імпорту у 2025-му, +{num(GROWTH)} % за рік і рекорд за "
              "вартістю. На особу це у 20 разів менше за ЄС."),
             ("Гарнір тримає один бренд", PLUM,
              "Ben's Original — 10 позицій у «Сільпо» за 45–179 грн. Другого бренду "
              "чистого готового рису на ринку немає."),
             ("Українці виграють за грам", GREEN,
              "Реторт-пауч 350 г за 71–199 грн — це 20–57 грн за 100 г проти 41 у Ben's "
              "і 91 у корейської чаші."),
             ("Мережі закриті для категорії", ROSE,
              "З 33 позицій у мережі продаються 10, і всі — Ben's. Український пауч "
              "живе на Prom і у фірмових магазинах.")]
    for i, (t, c, body) in enumerate(FACTS):
        x = M + (i % 2) * 6.12
        y = 1.88 + (i // 2) * 1.36
        rect(s, x, y, 5.86, 1.26, MIST, rounded=True, adj=0.08)
        rect(s, x, y, 0.09, 1.26, c, rounded=True, adj=0.5)
        text(s, x + 0.30, y + 0.16, 5.30, 0.26, f"{i + 1}. {t}", size=12.5, bold=True, color=NAVY)
        text(s, x + 0.30, y + 0.48, 5.30, 0.70, body, size=9.5, color=INK, line=1.28)
    insight(s, M, 4.72, 11.98, 1.56, "Чого дані не показують",
            "1. Продажів. Ні мережі, ні маркетплейси не віддають обсяги: ємність на слайді 03 "
            "порахована за імпортом, а не заміряна.\n"
            "2. Обсягу українського виробництва. Верес, «МАКРО», «Маркел», Portion "
            "виробляють у країні, тож у митну статистику не потрапляють — реальна категорія "
            "більша за 182 тонни.\n"
            "3. Імпортерів. У картках продавців їх не вказано; перелік дасть лише "
            "вивантаження Держмитслужби за УКТЗЕД 1904 90 10 за 2024–2025 рр.", ORANGE)
    foot(s, f"{N_SKU} позицій, {N_BRANDS} брендів, 16 продавців, 17 мереж. "
            "Зріз 23 вересня 2026 р. Предмет: рис кімнатного зберігання, який лише розігрівають.")


# ══ ЗБІРКА ════════════════════════════════════════════════════════════════
slide_cover()
slide_summary()
slide_volume()
slide_map()
slide_bens()
slide_bowls()
slide_ua()
slide_import()
slide_format()
slide_ladder()
slide_chain_shelf()
slide_channels()
slide_conclusions()

prs.save("Gotovyi_Rys_Rozbir_Brendiv.pptx")
print("saved ·", len(prs.slides._sldIdLst), "slides ·", N_SKU, "SKU ·", N_BRANDS, "brands ·",
      "median/unit", MED_UNIT, "· median/100g", round(MED100, 1))
