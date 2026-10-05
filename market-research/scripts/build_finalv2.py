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
    # 뚝배기불고기밥: у «Тайякі» — «бульгогі», в «Апетітаріум» — «гостра яловичина».
    # Обидві картки — те саме фото пачки. Одна позиція.
    ("sku/ot_bulgogi.jpg",     "Бульгогі",            320, 252, 356, "Тайякі · Апетітаріум"),
    # На пачці — свинина з каракатицею, рис 180 г + соус 130 г = 310 г. Три продавці
    # ведуть її як три різні позиції («свинина» 310 г, «гостра свинина» 269 г,
    # «з соусом з осьминога» 280 г) — на всіх трьох картках те саме фото. Одна позиція.
    ("sku/ot_pork310.jpg",     "Свинина з восьминогом", 310, 252, 334,
     "Тайякі · Gurm. · Апетіт."),
    ("sku/ot_chicken_rib.jpg", "Гострі курячі ребра", 310, 260, 289, "Gurmissimo · Апетітаріум"),
    ("sku/ot_tuna.jpg",        "Тунець і майонез",    247, 225, 356, "Тайякі · Gurm. · Апетіт."),
    ("sku/ot_bibimbap.jpg",    "Пібімпаб",            269, 252, 334, "Тайякі · Апетітаріум"),
]
# (файл, бренд, назва, грамів, від, до, канал, колір)
BOWL_MORE = [
    ("md/bibigo_bowl.jpg", "Bibigo", "Білий рис", 210, 189, 189, "Тайякі Март", GREEN),
]
# ── Страва в реторт-паучі, українське виробництво ─────────────────────────
UA_RETORT = [
    ("sku/ua_portion_pack.jpg", "Portion", "Каша рисова з м'ясом курки", 350, 71.4, 85,
     "portion.com.ua · Prom", ROSE),
    ("sku/ua_myasnytsia.jpg", "М'ясниця", "Свинина 35 %, горошок, кукурудза", 350, 88, 88,
     "Belorfoods", TEAL),
    ("sku/ua_markel_soy.jpg", "Маркел", "Рис із соєвим м'ясом", 350, 80, 96,
     "UP Shop · СУХПАЙ · Мартел-shop", SLATE),
    ("sku/ua_pak_pork.jpg", "МАКРО", "Свинина та овочі", 350, 96.1, 135,
     "UP Shop · СУХПАЙ · Ліхтар", ORANGE),
    ("sku/ua_veres_pork.jpg", "Верес", "Свинина, горошок, кукурудза", 350, 100, 100,
     "МореПродуктів", GREEN),
    ("sku/ua_makro_chicken.jpg", "МАКРО", "Курятина 35 % і овочі", 350, 101, 101,
     "СУХПАЙ", ORANGE),
    ("sku/ua_makro_beef.jpg", "МАКРО", "Яловичина й солодкий перець", 350, 103.85, 107,
     "UP Shop · СУХПАЙ", ORANGE),
    ("sku/ua_khodoriv_350.jpg", "Ходорівський", "Свинина та овочі", 350, 105, 105,
     "Молочний склад", PLUM),
    ("sku/ua_veres_chicken.jpg", "Верес", "Каша рисова з куркою", 350, 108, 134,
     "Козуб · Смачна адреса · Продукт-Shop", GREEN),
    ("sku/ua_markel_eat.jpg", "Маркел", "Рис із рослинним фаршем", 350, 127, 144,
     "UP Shop · СУХПАЙ", SLATE),
    ("sku/ua_khodoriv_pork.jpg", "Ходорівський", "Свинина, горошок, кукурудза", 350, 140, 199,
     "Foodi Shop · Food Shop · МореПродуктів", PLUM),
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
                 "МАКРО": "pouch", "Ходорівський": "pouch", "М'ясниця": "pouch",
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


def weights_of(items, idx=3):
    """Грамажі формату: до трьох — переліком «220, 240 і 250 г», більше — «217–320 г»."""
    ws = sorted({x[idx] for x in items})
    if len(ws) == 1:
        return f"{num(ws[0])} г"
    if len(ws) <= 3:
        return ", ".join(num(w) for w in ws[:-1]) + f" і {num(ws[-1])} г"
    return f"{num(ws[0])}–{num(ws[-1])} г"


def fmt_title(items, fmt):
    """«13 позицій: пауч 220, 240 і 250 г»."""
    return (f"{pl(len(items), 'позиція', 'позиції', 'позицій')}: "
            f"{fmt} {weights_of(items)}")


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
    text(s, x + 0.13, yy + 0.84, w - 0.26, 0.18, chan[:34], size=7.5, color=GREY)


def brand_head(s, x, y, brand, origin, fmt, n, mlo, mhi, plo, phi, c):
    rect(s, x, y, 11.98, 0.66, MIST, rounded=True, adj=0.20)
    rect(s, x, y, 0.09, 0.66, c, rounded=True, adj=0.5)
    text(s, x + 0.30, y + 0.07, 3.00, 0.28, brand, size=14, bold=True, color=NAVY)
    text(s, x + 0.30, y + 0.36, 3.00, 0.22, origin, size=8.5, color=GREY)
    for i, (lb, vl) in enumerate([("ФОРМАТ", fmt), ("ПОЗИЦІЙ", f"{n}"),
                                  ("ЗА УПАКОВКУ", f"{money(plo, phi)} грн"),
                                  ("МАСА", f"{rng(mlo, mhi)} г")]):
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
           "МАКРО": "Україна", "Ходорівський": "Україна", "М'ясниця": "Україна",
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
# Топ-3 грамажі за кількістю позицій
G_TOP = collections.Counter(x[3] for x in ALL8).most_common(3)

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
          f"Діапазон {num(min(x[4] for x in ALL8))}–{num(max(x[5] for x in ALL8))} грн.\n"
          f"Половина позицій — до {num(MED_UNIT)} грн."),
         ("МАСОВА ВАГА", f"{num(G_TOP[0][0])} г", SLATE,
          f"{pl(G_TOP[0][1], 'позиція', 'позиції', 'позицій')}. "
          f"Далі {num(G_TOP[1][0])} г — {G_TOP[1][1]},\n"
          f"{num(G_TOP[2][0])} г — {G_TOP[2][1]}.")]
    for i, (lb, v, c, note) in enumerate(K):
        kpi(s, M + i * 2.04, 1.88, 1.86, 1.92, lb, v, note, c)
    insight(s, M, 4.02, 5.86, 2.36, "Що показало дослідження",
            "1. Українське виробництво в категорії вже є — і воно найдешевше: реторт-пауч "
            "350 г за 71–199 грн, медіана 101 грн за упаковку.\n"
            "2. Ben's Original — єдиний імпортний гарнір у мережі: 10 позицій у «Сільпо» "
            "за 45–179 грн.\n"
            "3. Корейська чаша — найдорожчий сегмент (медіана 252 грн) і продається "
            "лише у трьох азійських магазинах.", ORANGE)
    insight(s, M + 6.12, 4.02, 5.86, 2.36, "Де порожньо",
            f"1. Немає українського гарніру — чистого рису в паучі. Усі {N_UA} українських "
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
    header(s, "КАРТА КАТЕГОРІЇ",
           f"Чотири сегменти, {pl(N_BRANDS, 'бренд', 'бренди', 'брендів')}",
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
        text(s, x + 0.22, 4.64, 2.40, 0.20,
             f"маса {rng(min(y[3] for y in items), max(y[3] for y in items))} г",
             size=8.5, bold=True, color=INK)
    insight(s, M, 5.14, 11.98, 1.42, "Як читати цю карту",
            "Український реторт-пауч уже виграє за ціною: медіана 101 грн проти 90 грн у "
            "Ben's Original — але в паучі 350 г проти 220–250 г, тобто це повна страва "
            "з м'ясом, а не гарнір. Інша полиця й інший привід.\n"
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
        ms = [x[3] for x in items]
        brand_head(s, M, y0, brand, origin, fmt, len(items), min(ms), max(ms), plo, phi, c)
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
    sku_slide("ФОРМАТ 1 · ПАУЧ · ГАРНІР · BEN'S ORIGINAL",
              fmt_title(L_SIDE, "пауч"),
              "Єдиний бренд чистого готового рису на ринку. Мікрохвильовка 90 секунд.",
              L_SIDE,
              "Ціни каталогу «Сільпо», отримані поштучно з картки кожного товару 23.09.2026; "
              "Long Grain 220 г — Edison Lee, у картці зазначено «ціну вказано за одну порцію». "
              "Позицій Sticky Bowl 220 г і Bio Basmati 240 г тут немає: у каталозі «Сільпо» "
              "їх не залишилося, в інших продавців картки зняті — перевірено 23.09.2026.",
              head=("Ben's Original", "Mars · Франція, Німеччина", "пауч", ORANGE, 45, 179))


def slide_bowls():
    sku_slide("ФОРМАТ 2 · ЧАША · OTTOGI ТА BIBIGO",
              fmt_title(L_BOWL, "чаша"),
              "Рис із наповнювачем у жорсткій чаші. Корея, канал — азійські фудшопи.",
              L_BOWL,
              "Ціни з карток Prom.ua 23.09.2026, роздрібні за одну чашу. У «Тайякі Март» "
              "діє нижча ціна від 15–20 шт — вона до розрахунку не бралася. Позиції Ottogi "
              "«Гамбурзький стейк» 315 г, «Чжамппонг» 217,5 г і «Кімчі» 310 г зняті: "
              "їх не продає жоден продавець. «Свинина з восьминогом» показана як одна "
              "позиція: три продавці ведуть її як три різні (310, 280 і 269 г), але на "
              "всіх трьох картках те саме фото пачки — рис 180 г плюс соус 130 г.",
              note=("Що це означає",
                    "Найдорожчий сегмент: медіана 252 грн за чашу проти 90 грн у Ben's і 101 грн "
                    "в українському паучі. Та сама позиція коштує по-різному: тунець 247 г — "
                    "225 грн у «Тайякі Март» і 356 грн в «Апетітаріум», розкид 58 %.", PLUM))


def slide_ua():
    sku_slide("ФОРМАТ 3 · РЕТОРТ-ПАУЧ · УКРАЇНА",
              fmt_title(L_UA, "реторт-пауч"),
              "Рис із м'ясом у реторт-паучі 350 г. Зберігання 24 місяці, розігрів без готування.",
              L_UA,
              "Ціни з карток продавців на Prom.ua і з фірмового магазину portion.com.ua, "
              "23.09.2026, за один пауч. Оптові пороги від 3–12 шт не бралися.",
              note=("Що це означає",
                    "Найдешевший сегмент категорії: 71–199 грн за пауч 350 г, медіана 101 грн. "
                    "Виробництво українське, тому ці позиції не потрапляють в імпортну "
                    "статистику на слайді 03 — реальна категорія більша за 182 тонни.\n"
                    "Канал вузький: Prom і фірмові магазини, у мережах цих брендів немає.",
                    GREEN))


def slide_import():
    sku_slide("ФОРМАТ 4 · РЕТОРТ-ПАУЧ · ІМПОРТ",
              fmt_title(L_IMP, "реторт-пауч"),
              "Той самий формат, але завезений: британська органіка й чеська туристична лінійка.",
              L_IMP, "Ціни з карток Prom.ua 23.09.2026 за одну упаковку.",
              two_rows=False,
              note=("Що це означає",
                    "Adventure Menu 400 г READY TO EAT — найбільша упаковка категорії, 315–466 грн, "
                    "канал туристичний. Clearspring 250 г за 446 грн — найдорожчий пауч ринку: "
                    "органічна сертифікація й дикий рис.", TEAL))


# ══ ФОРМАТ І ГРАМАЖ ═══════════════════════════════════════════════════════
def slide_format():
    s = slide()
    header(s, "ФОРМАТ І ГРАМАЖ",
           f"Пауч — {len(BY_PACK['pouch'])} позицій із {N_SKU}, масова вага 350 г",
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
                                     ("СЕРЕДНЯ МАСА", f"{num(st.median(ms))} г"),
                                     ("ДІАПАЗОН",
                                      f"{num(min(pr))} – {num(max(y[5] for y in it))} грн"),
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
    """Сходи брендів за ціною однієї упаковки. Крапка — медіана бренду."""
    s = slide()
    lo_all = min(x[4] for x in ALL8)
    hi_all = max(x[5] for x in ALL8)
    header(s, "ЦІНА ЗА ОДНУ УПАКОВКУ",
           f"Від {num(lo_all)} до {num(hi_all)} грн: сходи всіх {N_BRANDS} брендів",
           "Крапка — медіана бренду, смуга — від найдешевшої до найдорожчої його позиції.")
    SEGC = {}
    for _nm, _items, _c, _ in SEGMENTS:
        for _x in _items:
            SEGC[_x[1]] = (_c, _nm)
    rows = []
    for b in dict.fromkeys(x[1] for x in ALL8):
        it = [x for x in ALL8 if x[1] == b]
        rows.append((b, min(x[4] for x in it), max(x[5] for x in it),
                     st.median([x[4] for x in it]), len(it), COUNTRY[b], SEGC[b][0]))
    rows.sort(key=lambda r: r[3])
    LO, HI = 0, 500
    AX, AW = 4.70, 7.10

    def ax(v):
        return AX + AW * (v - LO) / (HI - LO)
    rect(s, M, 1.86, 11.98, 4.90, MIST, rounded=True, adj=0.04)
    for t in range(100, 501, 100):
        rect(s, ax(t), 2.10, 0.012, 3.94, RGBColor(0xDC, 0xE2, 0xEE))
        text(s, ax(t) - 0.40, 6.12, 0.80, 0.22, num(t), size=8.5, color=GREY,
             align=PP_ALIGN.CENTER)
    text(s, M + 0.26, 2.06, 2.30, 0.18, "БРЕНД", size=7, bold=True, color=GREY)
    text(s, M + 2.64, 2.06, 0.96, 0.18, "КРАЇНА", size=7, bold=True, color=GREY)
    text(s, M + 3.66, 2.06, 0.34, 0.18, "SKU", size=7, bold=True, color=GREY,
         align=PP_ALIGN.RIGHT)
    text(s, M + 0.26, 6.12, 3.20, 0.22, "ГРН ЗА ОДНУ УПАКОВКУ", size=8.5, bold=True, color=GREY)
    TOP, BOT = 2.46, 5.90
    step = min(0.372, (BOT - TOP) / max(1, len(rows) - 1))
    yy = TOP
    for b_, lo, hi, med, n, cty, c in rows:
        text(s, M + 0.26, yy - 0.11, 2.34, 0.24, b_, size=10.5, bold=True, color=NAVY)
        text(s, M + 2.64, yy - 0.09, 0.96, 0.20, cty, size=8, color=GREY)
        text(s, M + 3.60, yy - 0.09, 0.40, 0.20, str(n), size=8, bold=True, color=GREY,
             align=PP_ALIGN.RIGHT)
        x0, x1 = ax(lo), ax(hi)
        rect(s, x0, yy - 0.045, max(x1 - x0, 0.03), 0.09, c, rounded=True, adj=0.5)
        dot(s, ax(med), yy, 0.19, c)
        lbl = num(med) if lo == hi else f"{num(lo)} – {num(hi)}"
        text(s, x1 + 0.12, yy - 0.115, 1.40, 0.24, lbl, size=9, bold=True, color=c)
        yy += step
    rect(s, ax(45), 6.40, ax(200) - ax(45), 0.06, ORANGE)
    text(s, ax(45), 6.50, 4.20, 0.24, "вікно мережевої полиці: 45–200 грн", size=9,
         bold=True, color=ORANGE)
    lx = M + 0.26
    for _lb, _c in [("Гарнір", ORANGE), ("Чаша", PLUM), ("Пауч UA", GREEN),
                    ("Пауч імпорт", TEAL)]:
        dot(s, lx + 0.06, 6.52, 0.13, _c)
        w = 0.16 + 0.056 * len(_lb)
        text(s, lx + 0.18, 6.42, w + 0.10, 0.22, _lb, size=8, color=INK)
        lx += w + 0.28
    foot(s, "Роздрібна ціна за одну упаковку на сайті продавця, 23.09.2026. Оптові пороги "
            "(у частини продавців діє нижча ціна від 3–20 шт) у розрахунок не входять.")


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
    text(s, M + 6.40, 5.10, 5.30, 0.26, "Порівняння за упаковку", size=12, bold=True, color=NAVY)
    CMP = [("Каша рисова hapay! 340 г · бляшанка", 63, GREEN),
           ("Ben's Original 250 г · медіана", 90, PLUM),
           ("Український пауч 350 г · медіана", 101, ORANGE),
           ("Плов М'ясторія 350 г · лоток", 147, ROSE),
           ("Ottogi 247–320 г · медіана", 252, SLATE)]
    yy = 5.44
    mxv = max(v for _, v, _ in CMP)
    for lb, v, c in CMP:
        text(s, M + 6.40, yy - 0.04, 2.70, 0.20, lb, size=8.5, color=INK)
        rect(s, M + 9.20, yy + 0.01, 1.90 * v / mxv, 0.13, c, rounded=True, adj=0.5)
        text(s, M + 9.20 + 1.90 * v / mxv + 0.08, yy - 0.05, 0.80, 0.20, f"{num(v)} грн",
             size=8, bold=True, color=c)
        yy += 0.28
    foot(s, "Джерело: API 17 мереж zakaz.ua, 23.09.2026. Ці шість позицій — суміжна "
            f"категорія, у {N_SKU} позицій дослідження вони не входять: бляшанку не "
            "розігрівають у мікрохвильовці.")




# ══ ФІНАЛЬНА РЕДАКЦІЯ: нові/перероблені слайди ═══════════════════════════
# Джерело нових даних:
#  - митна база (codexmb), вивантаження 28.09.2026, УКТ ЗЕД 1904901000,
#    період 01.2025–05.2026 (без 12–31.03.2025)
#  - RTE_Rice_Suppliers.pptx — профілі BSCM Foods, CM Premium Rice, One's International
#  - RICE_PRICING_ALL_MARGINS_SELF_COST_NAMES.xlsx — фінмодель собівартість→полиця

# ── митні дані: код 1904901000, розбір асортименту ────────────────────────
CODE = "1904901000"
CODE_T = 250.275          # т за весь період вибірки
CODE_USD = 841445
FROZEN_T, FROZEN_SH = 187.056, 0.748     # заморожені сирі суміші рис+овочі
CHIPS_T, CHIPS_SH = 32.124, 0.128        # рисові чіпси-снеки
DOLMA_T, DOLMA_SH = 10.8, 0.043          # долма в банці
RISOTTO_T, RISOTTO_SH = 5.86, 0.023      # різото "довести до готовності" (не зварене)
OTHER_T, OTHER_SH = 8.711, 0.035         # неоднозначні картки (переважно ще різото/суміші)
RTE_T = 0.814              # Mars Austria + CJ CheilJedang + Clearspring
HENAN_T = 2.639            # залити окропом, суміжний сегмент
BENS_BATCH_KG = 675.84
BENS_IMPORTER = 'ТОВ "ФОЗЗІ КОММЕРЦ"'
BENS_DATE = "24.02.2025"

CODE_BARS = [
    ("Заморожені сирі суміші рис+овочі", FROZEN_T, FROZEN_SH, SLATE,
     "Oerlemans, «Рудь»: ризото й гавайська суміш, сирі, у морозилку. Не наш продукт."),
    ("Рисові чіпси й снеки", CHIPS_T, CHIPS_SH, GREY,
     "Ficosota «Livity Chips» — солоний снек, не страва."),
    ("Різото «довести до готовності»", RISOTTO_T, RISOTTO_SH, PLUM,
     "Riso Gallo, Casa Rinaldi, Cordero — рис неприготований, варити 15–18 хв."),
    ("Долма в банці", DOLMA_T, DOLMA_SH, ROSE,
     "Palirria: виноградне листя з рисом, інший формат і привід."),
    ("Інше / неоднозначний опис", OTHER_T, OTHER_SH, MIST_D,
     "Переважно той самий «варити» різото під іншою назвою."),
    ("Готовий рис для розігріву", RTE_T / CODE_T, RTE_T / CODE_T, ORANGE,
     f"Mars Austria (Ben's) {num(BENS_BATCH_KG/1000,3)} т, CJ CheilJedang, Clearspring."),
]


# ══ 01 · ОБКЛАДИНКА (перероблено) ═════════════════════════════════════════
def slide_cover():
    s = slide(dark=True)
    cover_pic(s, BG_TITLE, 0, 0, W, H)
    s.shapes.add_picture(LOGO, Inches(M), Inches(0.52), Inches(1.94), Inches(0.84))
    rect(s, M, 2.02, 0.54, 0.09, ORANGE)
    text(s, M, 2.34, 7.00, 0.34, "ФІНАЛЬНА ВЕРСІЯ · ЖОВТЕНЬ 2026", size=11.5,
         bold=True, color=AMBER)
    text(s, M, 2.78, 7.20, 1.70, "Готовий рис\nв Україні", size=46, bold=True,
         color=WHITE, line=1.04)
    text(s, M, 4.52, 7.00, 0.36, "Кімнатне зберігання · розігрів · без готування",
         size=14.5, color=RGBColor(0xC6, 0xCE, 0xE2))
    rect(s, M, 5.12, 5.90, 0.055, RGBColor(0x55, 0x60, 0x8C))
    text(s, M, 5.34, 6.90, 0.86,
         f"{pl(N_BRANDS, 'бренд', 'бренди', 'брендів')} · "
         f"{pl(N_SKU, 'позиція', 'позиції', 'позицій')} · ціна за одну упаковку\n"
         "Ціни — 23–28.09.2026, митна база — вивантаження 28.09.2026.",
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


# ══ 02 · ГОЛОВНЕ (перероблено) ════════════════════════════════════════════
def slide_summary():
    s = slide()
    header(s, "ГОЛОВНЕ", "Шість цифр про категорію",
           "Предмет: рис, що зберігається за кімнатної температури і лише розігрівається.")
    K = [("ФОРМАЛЬНИЙ ІМПОРТ READY-TO-EAT", f"{num(RTE_T,2)} т", ORANGE,
          "За 16 міс. (01.2025–05.2026) за\nмитною базою, код 1904901000."),
         ("ПОЗИЦІЙ У ПРОДАЖУ", f"{N_SKU}", PLUM,
          f"{N_BRANDS} брендів. {N_UA} позицій —\nукраїнського виробництва."),
         ("НА ПОЛИЦІ МЕРЕЖ", "12 SKU", TEAL,
          "11 — Ben's Original у «Сільпо».\nРитейл-аудит, 9 мереж, 28.09.2026."),
         ("МЕДІАНА ЗА УПАКОВКУ", f"{num(MED_UNIT)} грн", GREEN,
          f"Діапазон {num(min(x[4] for x in ALL8))}–{num(max(x[5] for x in ALL8))} грн.\n"
          f"Половина позицій — до {num(MED_UNIT)} грн."),
         ("НАЙДЕШЕВШИЙ СЕГМЕНТ", "101 грн", ROSE,
          "Український реторт-пауч 350 г,\n71–199 грн, 6 виробників."),
         ("МАСОВА ВАГА", f"{num(G_TOP[0][0])} г", SLATE,
          f"{pl(G_TOP[0][1], 'позиція', 'позиції', 'позицій')}. "
          f"Далі {num(G_TOP[1][0])} г — {G_TOP[1][1]},\n"
          f"{num(G_TOP[2][0])} г — {G_TOP[2][1]}.")]
    for i, (lb, v, c, note) in enumerate(K):
        kpi(s, M + i * 2.04, 1.88, 1.86, 1.92, lb, v, note, c)
    insight(s, M, 4.02, 5.86, 2.36, "Що показало дослідження",
            "1. Організованого імпорту категорії практично немає: за 16 місяців митниця "
            "бачить менше тонни готового рису для розігріву — одна партія Ben's (676 кг, "
            "24.02.2025, імпортер «Фоззі Коммерц») і сліди CJ та Clearspring.\n"
            "2. Українське виробництво вже є — і воно найдешевше: реторт-пауч 350 г за "
            "71–199 грн, медіана 101 грн.\n"
            "3. Корейська чаша — найдорожчий сегмент (медіана 252 грн), три азійські "
            "магазини.", ORANGE)
    insight(s, M + 6.12, 4.02, 5.86, 2.36, "Де порожньо",
            "1. Жодного постачальника, що возить контейнерами на постійній основі: "
            "усе, що є, — разові партії або локальне виробництво.\n"
            f"2. Немає українського гарніру — чистого рису в паучі. Усі {N_UA} українських "
            "позицій це рис із м'ясом.\n"
            "3. Немає жодної позиції категорії в АТБ, Novus, Metro, Varus, Ашан, Fozzy "
            "та інших мережах — перевірено 28.09.2026.", TEAL)
    foot(s, "Ціни: сайти продавців, 23.09.2026, за одну упаковку. Імпорт: митна база "
            "(codexmb), код 1904901000, вивантаження 28.09.2026. Полиця: ритейл-аудит "
            "9 мереж, 28.09.2026.")


# ══ 03 · ОБСЯГ РИНКУ (перероблено на митну базу) ═════════════════════════
def slide_volume():
    s = slide()
    header(s, "ОБСЯГ РИНКУ",
           f"Код 1904901000 — {num(CODE_T,1)} т, з них на наш продукт {num(RTE_T,2)} т",
           "Митний код «рис на основі зерна» ширший за нашу категорію в 300 разів за вагою.")
    rect(s, M, 1.88, 7.70, 4.00, MIST, rounded=True, adj=0.04)
    text(s, M + 0.26, 2.04, 6.80, 0.24, "З ЧОГО СКЛАДАЄТЬСЯ КОД 1904901000", size=9,
         bold=True, color=GREY)
    yy = 2.42
    mx = max(sh for _, _, sh, _, _ in CODE_BARS)
    for lb, t, sh, c, note in CODE_BARS:
        text(s, M + 0.26, yy, 3.60, 0.20, lb, size=9.5, bold=True, color=INK)
        text(s, M + 7.00, yy, 0.86, 0.20, f"{num(sh*100,1)} %", size=9.5, bold=True,
             color=c, align=PP_ALIGN.RIGHT)
        rect(s, M + 0.26, yy + 0.24, 7.18, 0.16, WHITE, rounded=True, adj=0.5)
        rect(s, M + 0.26, yy + 0.24, 7.18 * sh / mx, 0.16, c, rounded=True, adj=0.5)
        text(s, M + 0.26, yy + 0.43, 6.90, 0.24, note, size=7.2, color=GREY, line=1.1)
        yy += 0.565
    rect(s, M + 8.02, 1.88, 3.96, 4.00, NAVY, rounded=True, adj=0.05)
    rect(s, M + 8.02, 1.88, 0.09, 4.00, ORANGE, rounded=True, adj=0.5)
    text(s, M + 8.30, 2.06, 3.50, 0.52, "Єдина знайдена партія Ben's",
         size=13, bold=True, color=AMBER, line=1.10)
    text(s, M + 8.30, 2.64, 3.50, 0.30, f"{num(BENS_BATCH_KG,2)} кг", size=26, bold=True,
         color=WHITE)
    text(s, M + 8.30, 3.06, 3.50, 0.22, "усього за 16 місяців", size=9,
         color=RGBColor(0x9F, 0xA9, 0xC4))
    yy2 = 3.42
    for lb, v in [("Дата", BENS_DATE), ("Імпортер", "«Фоззі Коммерц» (Fozzy Group)"),
                  ("Виробник", "Mars Austria OG"), ("Упаковок", "≈ 2 900 шт. на 11 SKU")]:
        text(s, M + 8.30, yy2, 1.30, 0.34, lb, size=8, bold=True,
             color=RGBColor(0x9F, 0xA9, 0xC4))
        text(s, M + 9.54, yy2, 2.26, 0.34, v, size=9, color=WHITE, line=1.1)
        yy2 += 0.40
    text(s, M + 8.30, 5.10, 3.50, 0.70,
         "«Фоззі Коммерц» — торговий оператор Fozzy Group (власник «Сільпо»). "
         "Це пояснює, чому всі 11 SKU Ben's Original стоять саме в «Сільпо».",
         size=8.5, color=RGBColor(0xD5, 0xDB, 0xEA), line=1.26)
    insight(s, M, 6.02, 11.98, 0.84, None,
            "Висновок: організованого контейнерного імпорту категорії немає — є одна разова "
            "партія лідера та сліди ще двох брендів. Хто завезе постійний обсяг, матиме "
            "ринок практично без прямого імпортного конкурента.", ORANGE)
    foot(s, "Митна база (codexmb), вивантаження 28.09.2026, УКТ ЗЕД 1904901000, "
            "01.2025–05.2026 (без 12–31.03.2025). Фактурна вартість USD без ПДВ і мита. "
            "Частки — за вагою нетто рядків бази.")


# ══ РИТЕЙЛ-АУДИТ ═════════════════════════════════════════════════════════
RETAIL_ROWS = [
    ("Сільпо", "Ben's Original / Uncle Ben's Express", "пауч 220–250 г", 11, "90–179", "Mars · ЄС"),
    ("МегаМаркет · Ultramarket", "The Local Food по-тайськи з куркою", "лоток 350 г", 1, "207",
     "Україна"),
]
RETAIL_ADJ = [
    ("Ашан · Сільпо · NOVUS · МегаМаркет · Ultramarket", "Охолоджена кулінарія 2–5 діб",
     "лоток / кулінарія", 11, "75–439", "Україна"),
    ("Ашан · NOVUS · МегаМаркет · Сільпо", "Суміш під варіння (не готовий продукт)",
     "пакет 200–500 г", 3, "59–92", "Україна"),
]


def slide_retail_audit():
    s = slide()
    header(s, "РИТЕЙЛ-АУДИТ", "Що з готового рису реально стоїть у мережах",
           "Суцільна перевірка каталогів 9 мереж: Ашан, Сільпо, METRO, NOVUS, МегаМаркет, "
           "Ultramarket, ЕКО маркет, Epicentr, Космос — 28.09.2026.")
    rect(s, M, 1.86, 11.98, 0.50, NAVY, rounded=True, adj=0.20)
    text(s, M + 0.24, 1.98, 3.20, 0.26, "ЗБЕРІГАННЯ ПРИ КІМНАТНІЙ · НАШ СЕГМЕНТ", size=10,
         bold=True, color=AMBER)
    text(s, M + 10.40, 1.98, 1.40, 0.26, "12 SKU", size=12, bold=True, color=WHITE,
         align=PP_ALIGN.RIGHT)
    hdrs = ["МЕРЕЖА", "ЩО САМЕ СТОЇТЬ", "ФОРМАТ", "SKU", "ЦІНА, ГРН", "ПОХОДЖЕННЯ"]
    widths = [2.55, 3.85, 1.65, 0.65, 1.55, 1.73]
    yy = 2.50
    xx = M
    for h_, w in zip(hdrs, widths):
        text(s, xx, yy, w - 0.08, 0.20, h_, size=7.5, bold=True, color=GREY)
        xx += w
    yy += 0.26
    for row, c in [(RETAIL_ROWS[0], ORANGE), (RETAIL_ROWS[1], GREEN)]:
        net, what, fmt, n, price, origin = row
        rect(s, M, yy - 0.04, 11.98, 0.40, MIST, rounded=True, adj=0.3)
        xx = M
        for val, w in zip([net, what, fmt, str(n), price + " грн", origin], widths):
            text(s, xx + 0.10, yy + 0.02, w - 0.14, 0.30, val, size=8.5, bold=True,
                 color=NAVY if xx == M else INK, line=1.05)
            xx += w
        yy += 0.46
    yy += 0.10
    text(s, M, yy, 8, 0.20, "СУМІЖНЕ: ОХОЛОДЖЕНА КУЛІНАРІЯ Й СУМІШІ ПІД ВАРІННЯ — НЕ НАШ ПРОДУКТ",
         size=8, bold=True, color=GREY)
    yy += 0.28
    for net, what, fmt, n, price, origin in RETAIL_ADJ:
        xx = M
        for val, w in zip([net, what, fmt, str(n), price + " грн", origin], widths):
            text(s, xx + 0.10, yy, w - 0.14, 0.38, val, size=8, color=GREY, line=1.08)
            xx += w
        yy += 0.42
    yy += 0.06
    xx = M
    for val, w in zip(["METRO", "готового рису немає", "—", "0", "—", "лише крупа 1–10 кг"], widths):
        text(s, xx + 0.10, yy, w - 0.14, 0.30, val, size=8, color=GREY, italic=True)
        xx += w
    insight(s, M, 5.86, 11.98, 1.02, None,
            "11 із 12 SKU «нашого» сегмента — Ben's Original у «Сільпо». Єдина інша позиція: "
            "український лоток The Local Food по-тайськи 350 г за 207 грн у МегаМаркеті. "
            "METRO не тримає категорію взагалі — лише крупу під власними марками для HoReCa.",
            ORANGE)
    foot(s, "Суцільна перевірка каталогів мереж, 28.09.2026. Охолоджена кулінарія та суміші "
            "під варіння показані для контексту — у 32 перевірені позиції не входять.")


# ══ КАНАЛИ ПРОДАЖУ ═══════════════════════════════════════════════════════
def slide_channels():
    s = slide()
    header(s, "КАНАЛИ ПРОДАЖУ", "Де продається категорія і кому",
           "Дві осі: точка продажу і покупець, який у неї приходить. Вони не збігаються.")
    GROUPS = [("МЕРЕЖІ НАЦІОНАЛЬНОГО ПОКРИТТЯ", ORANGE, 3,
               "Сільпо · МегаМаркет · Ultramarket. Решта перевірених мереж категорії не мають.",
               "Масовий покупець · щоденне харчування"),
              ("ОНЛАЙН-МАРКЕТПЛЕЙСИ", TEAL, 3,
               "MAUDAU · Rozetka · Prom.",
               "Масовий покупець · плановане замовлення"),
              ("АЗІЙСЬКІ ТА ЕТНІЧНІ МАГАЗИНИ", PLUM, 16,
               "Тайякі Март · Апетітаріум · Gurmissimo · Edison Lee · Скарби Азії та ще 11.",
               "Покупець азійської кухні · знає категорію"),
              ("ФІРМОВІ МАГАЗИНИ ТА PROM", GREEN, 10,
               "portion.com.ua · СУХПАЙ · UP Shop · Козуб Маркет · Belorfoods та ще 5.",
               "Шукає конкретний бренд напряму")]
    yy = 1.86
    for title, c, n, chans, buyer in GROUPS:
        rh = 0.86
        rect(s, M, yy, 11.98, rh, MIST, rounded=True, adj=0.10)
        rect(s, M, yy, 0.09, rh, c, rounded=True, adj=0.5)
        text(s, M + 0.30, yy + 0.10, 4.00, 0.22, title, size=10, bold=True, color=c)
        text(s, M + 0.30, yy + 0.35, 9.40, 0.24, chans, size=8.5, color=INK, line=1.1)
        text(s, M + 0.30, yy + 0.59, 9.40, 0.20, "ПОКУПЕЦЬ: " + buyer, size=7.5, color=GREY)
        text(s, M + 11.00, yy + rh / 2 - 0.15, 0.84, 0.30, str(n), size=16,
             bold=True, color=c, align=PP_ALIGN.RIGHT)
        yy += rh + 0.10
    insight(s, M, yy + 0.04, 11.98, 1.06, "Що це означає",
            "Канал і покупець — різні осі. Азійські фудшопи дають найбільше точок продажу "
            "(16), і там уже стоять Ottogi, Bibigo, Henan — прямі конкуренти стакана. "
            "Мережевий роздріб — лише 3 мережі з одним Ben's. Український реторт-пауч не "
            "представлений у жодній мережі: живе винятково на Prom і у фірмових магазинах.",
            GREEN)
    foot(s, "Перелік мереж і магазинів — за картками продавців і каталогами, 23–28.09.2026.")


# ══ ПОСТАЧАЛЬНИКИ ════════════════════════════════════════════════════════
SUPPLIERS = [
    ("BSCM FOODS", "Таїланд · Бангкок", "sku/sup_bscm_pouch.jpg", ORANGE,
     [("Формати", "пауч 240 г · 150/200 г · стакан 150 г · подвійний стакан 2×125 г"),
      ("Термін придатності", "18 місяців за кімнатної температури"),
      ("Сертифікація", "BRC · HACCP · ISO 22000 · GMP · Organic"),
      ("FOB-діапазон", "$0,472 – $0,750 залежно від SKU і ваги"),
      ("Чому саме цей", "Єдиний постачальник зі стаканом — форматом, яким можна конкурувати "
                         "з Ottogi за ціною")]),
    ("ONE'S INTERNATIONAL", "Корея · Гурі-сі, бренд KBROS", "sku/sup_ones_kimchi.jpg", PLUM,
     [("Формати", "лоток 200 г (fried rice) · лоток 210 г (білий рис)"),
      ("Термін придатності", "12 місяців — найкоротший з трьох"),
      ("Сертифікація", "HACCP · FSSC 22000"),
      ("FOB-діапазон", "Не розкрито постачальником. Оцінка із собівартості: ~$0,571 і ~$1,238"),
      ("Чому саме цей", "Корейське походження — премія в азійському каналі, де вже є Bibigo")]),
    ("CM PREMIUM RICE", "Таїланд · Nakhon Ratchasima", "sku/sup_cmpremium_chefrey.jpg", TEAL,
     [("Формати", "пауч 250 г органік, бренд Chefrey · 24 шт у коробі"),
      ("Потужність", "120 000 т/рік · 30+ партнерських рисозаводів"),
      ("Сертифікація", "BRC · HACCP · ISO 9001 · Halal · Organic"),
      ("FOB-діапазон", "$0,667 — однаковий на всі 4 SKU"),
      ("Чому саме цей", "Сертифікована органіка — інша полиця й інший покупець, ніша")]),
]


def slide_suppliers():
    s = slide()
    header(s, "ПОСТАЧАЛЬНИКИ", "Три постачальники в обоймі",
           "BSCM і CM Premium — Таїланд, One's International — Корея. Оцінено за форматом, "
           "сертифікацією та ціновою позицією.")
    w = (11.98 - 0.10 * 2) / 3
    for i, (name, origin, img, c, rows) in enumerate(SUPPLIERS):
        x = M + i * (w + 0.10)
        rect(s, x, 1.88, w, 4.98, WHITE, line=MIST_D, lw=1.2, rounded=True, adj=0.04)
        rect(s, x, 1.88, w, 0.09, c, rounded=True, adj=0.5)
        rect(s, x + 0.14, 2.10, w - 0.28, 1.28, MIST, rounded=True, adj=0.08)
        pic(s, img, x + 0.22, 2.16, w - 0.44, 1.16)
        text(s, x + 0.18, 3.46, w - 0.36, 0.26, name, size=13, bold=True, color=NAVY)
        text(s, x + 0.18, 3.72, w - 0.36, 0.22, origin, size=8.5, color=GREY)
        yy = 3.98
        for lb, val in rows:
            text(s, x + 0.18, yy, w - 0.36, 0.16, lb.upper(), size=6.8, bold=True, color=c)
            text(s, x + 0.18, yy + 0.155, w - 0.36, 0.40, val, size=7.3, color=INK, line=1.12)
            yy += 0.555
        foot_src = {0: "каталог продукції BSCM Foods, комерційна пропозиція постачальника",
                    1: "Quotation One's International Co., Ltd. (KBROS), FOB Busan, "
                       "чинна до 31.12.2026",
                    2: "корпоративна презентація CM Premium Rice Co., Ltd., сайт cmp-rice.com"}
        text(s, x + 0.18, 6.68, w - 0.36, 0.20, foot_src[i], size=6.3, color=GREY, line=1.08)
    foot(s, "Профілі постачальників — за їхніми комерційними пропозиціями та "
            "корпоративними матеріалами, вересень 2026.")


# ══ ФІНМОДЕЛЬ ════════════════════════════════════════════════════════════
# Собівартість (self-cost, $/од., 20' навалом) з власної фінмоделі покупця,
# аркуш «СС_база». Полиця = собівартість(грн) × 3,5 за формулою аркуша
# «Финансовая модель — Ціни» (курс 45 грн/$, бонус мережі 25 %, цільова
# маржа 35 %). FOB = собівартість($) ÷ 1,7 — курс підтверджено точним
# збігом із власними оцінками постачальника (слайд «Постачальники»).
RATE, MARKUP, FOBFACT = 45, 3.5, 1.7
FIN_ROWS = [
    ("BSCM стакан 150 г", 0.8814, "Герой · азійський формат", "Ottogi 255 · Henan 140", ORANGE),
    ("ONE'S лоток 210 г", 0.9715, "Другий герой · Корея", "Bibigo 135–189", PLUM),
    ("BSCM пауч 200 г", 0.9526, "Об'єм · вхід у ціну", "пауч-медіана 109", ORANGE),
    ("BSCM пауч 240 г", 0.9954, "Мережа · друга хвиля", "Ben's медіана 90", ORANGE),
    ("CM Premium 250 г органік", 1.1330, "Преміум · ніша", "Clearspring 446", TEAL),
    ("ONE'S Fried Rice 200 г", 2.1049, "Переглянути або зняти", "Ottogi 225–356", PLUM),
    ("BSCM подвійний стакан 2×125 г", 1.3908, "Private label · пізніше", "верх вікна полиці",
     ORANGE),
]


def fin_calc(cc_usd):
    cc_h = cc_usd * RATE
    shelf = cc_h * MARKUP
    fob = cc_usd / FOBFACT
    return cc_h, shelf, fob


def slide_finmodel():
    s = slide()
    header(s, "ФІНМОДЕЛЬ ПРОТИ РИНКУ", "Скільки коштує кожен наш формат на полиці",
           "Полиця = собівартість × 3,5 за поточних умов (курс 45, бонус мережі 25 %, "
           "цільова маржа 35 %).")
    hdrs = ["ТОВАРНА ГРУПА", "РОЛЬ У ЛІНІЙЦІ", "СС, $/од.", "ПОЛИЦЯ ЗАРАЗ", "FOB (оцінка)",
            "ОРІЄНТИР НА ПОЛИЦІ"]
    widths = [2.75, 2.55, 1.35, 1.65, 1.65, 2.03]
    yy = 1.86
    xx = M
    for h_, w in zip(hdrs, widths):
        text(s, xx, yy, w - 0.08, 0.20, h_, size=7.5, bold=True, color=GREY)
        xx += w
    yy += 0.26
    shelf_values = []
    for name, cc, role, anchor, c in FIN_ROWS:
        cc_h, shelf, fob = fin_calc(cc)
        shelf_values.append((name, shelf, c))
        rect(s, M, yy - 0.03, 11.98, 0.42, MIST, rounded=True, adj=0.25)
        rect(s, M, yy - 0.03, 0.07, 0.42, c, rounded=True, adj=0.5)
        xx = M + 0.14
        vals = [name, role, f"${cc:.3f}", f"{num(shelf)} грн", f"${fob:.3f}", anchor]
        for val, w in zip(vals, widths):
            bold = val in (name,) or "грн" in val
            text(s, xx, yy + 0.03, w - 0.16, 0.34, val, size=8.3, bold=bold,
                 color=NAVY if val == name else (c if "грн" in val else INK), line=1.05)
            xx += w
        yy += 0.475
    yy += 0.08
    lo_s = min(v for _, v, _ in shelf_values)
    hi_s = max(v for _, v, _ in shelf_values)
    insight(s, M, yy, 11.98, 1.30, "Що це означає",
            f"Діапазон полиці за поточною собівартістю — від {num(lo_s)} до {num(hi_s)} грн. "
            "П'ять із семи товарних груп потрапляють у вікно мережевої полиці 45–200 грн. "
            "ONE'S Fried Rice 200 г випадає найсильніше: 332 грн, удвічі дорожче за "
            "власний лоток білого рису того самого постачальника — собівартість смаженої "
            "страви в лотку більш ніж удвічі вища за решту лінійки. Переглянути ціну або "
            "не заводити зараз. Подвійний стакан (219 грн) — кандидат на private label "
            "пізніше, не на перший запуск.", ORANGE)
    foot(s, "Власна фінмодель покупця (аркуші «СС_база» і «Финансовая модель — Ціни»), "
            "перевірено поштучно по кожній товарній групі, 05.10.2026.")


# ══ РЕКОМЕНДАЦІЯ ═════════════════════════════════════════════════════════
def slide_recommendation():
    s = slide()
    header(s, "РЕКОМЕНДАЦІЯ", "Що заводити: формат, ціна, постачальник",
           "Три кандидати, оцінені за конкурентною позицією, а не лише за собівартістю.")
    CARDS = [
        ("РЕКОМЕНДУЄМО", "BSCM · СТАКАН 150 Г", "129 грн", ORANGE,
         "Ottogi 225–356 грн · Henan 83–156 грн",
         "−49 % до медіани Ottogi (252 грн). Азійський формат, азійський канал — 16 фудшопів "
         "уже продають цю полицю.",
         "Потрібно від BSCM: FOB $0,480 замість поточних $0,500."),
        ("ДРУГИЙ ПРІОРИТЕТ", "ONE'S · ЛОТОК 210 Г", "153 грн", PLUM,
         "Bibigo 210 г · 135–189 грн",
         "Той самий грамаж, ціна в середині діапазону прямого конкурента. Найточніше "
         "зіставлення в лінійці.",
         "Ризик: MOQ 200 коробів на SKU, передоплата 100 % T/T, найкоротший термін "
         "придатності (12 міс)."),
        ("НЕ ЗАВОДИТИ ЗАРАЗ", "ONE'S · FRIED RICE 200 Г", "332 грн", GREY,
         "Ottogi 225–356 грн",
         "Собівартість лотка смаженого рису вдвічі вища за решту лінійки постачальника — "
         "полиця виходить за межі вікна 45–200 грн.",
         "Умова входу: FOB ≤ $0,740 (−40 %) або інший, дешевший SKU на заміну."),
    ]
    w = (11.98 - 0.10 * 2) / 3
    for i, (tag, name, price, c, comp, why, risk) in enumerate(CARDS):
        x = M + i * (w + 0.10)
        rect(s, x, 1.88, w, 3.50, WHITE, line=MIST_D, lw=1.2, rounded=True, adj=0.05)
        rect(s, x, 1.88, w, 0.09, c, rounded=True, adj=0.5)
        text(s, x + 0.20, 2.08, w - 0.40, 0.20, tag, size=8.5, bold=True, color=c)
        text(s, x + 0.20, 2.30, w - 0.40, 0.46, name, size=13, bold=True, color=NAVY, line=1.08)
        text(s, x + 0.20, 2.80, w - 0.40, 0.36, price, size=24, bold=True, color=c)
        text(s, x + 0.20, 3.22, w - 0.40, 0.18, "полиця за фінмоделлю", size=7.5, color=GREY)
        rect(s, x + 0.20, 3.46, w - 0.40, 0.02, MIST_D)
        text(s, x + 0.20, 3.56, w - 0.40, 0.18, "ПРЯМИЙ КОНКУРЕНТ", size=7, bold=True, color=GREY)
        text(s, x + 0.20, 3.74, w - 0.40, 0.20, comp, size=8.5, color=INK)
        text(s, x + 0.20, 3.98, w - 0.40, 0.70, why, size=8, color=INK, line=1.22)
        text(s, x + 0.20, 4.72, w - 0.40, 0.60, risk, size=7.5, color=GREY, line=1.2)
    insight(s, M, 5.54, 11.98, 1.34, "Послідовність запуску",
            "1. BSCM стакан 150 г — герой лінійки, запуск за цільовою ціною 129 грн, "
            "тисну на постачальника щодо FOB $0,480.\n"
            "2. ONE's лоток 210 г — другою хвилею, прямий аналог Bibigo, той самий канал.\n"
            "3. ONE's Fried Rice 200 г і BSCM подвійний стакан — відкласти: перший потребує "
            "перегляду ціни постачальника, другий — кандидат на private label пізніше.",
            GREEN)
    foot(s, "Конкурентні ціни — з карток продавців 23.09.2026. Полиця й FOB — власна "
            "фінмодель, 05.10.2026.")


# ══ ВИСНОВКИ (перероблено) ════════════════════════════════════════════════
def slide_conclusions():
    s = slide()
    header(s, "ВИСНОВКИ", "П'ять фактів і три обмеження",
           "Усе нижче спирається на ціни, перевірені поштучно 23–28 вересня, і митну базу, "
           "вивантажену 28.09.2026.")
    FACTS = [("Організованого імпорту немає", ORANGE,
              f"За 16 місяців митниця бачить {num(RTE_T,2)} т готового рису для розігріву — "
              "одна партія Ben's і сліди ще двох брендів. Нікого, хто возить постійно."),
             ("Гарнір тримає один бренд", PLUM,
              "Ben's Original — 11 позицій у «Сільпо» за 45–179 грн, завезені разовою "
              "партією через «Фоззі Коммерц». Другого бренду немає."),
             ("Українці виграють за ціною", GREEN,
              "Реторт-пауч 350 г за 71–199 грн, медіана 101 грн. Більша порція за ту саму "
              "ціну, що імпортний гарнір 220–250 г."),
             ("Мережі закриті для категорії", ROSE,
              "З 32 позицій у мережі продаються 12 — і 11 з них Ben's. Український пауч "
              "живе на Prom і у фірмових магазинах виробників."),
             ("Вхід без прямого конкурента", TEAL,
              "BSCM стакан 150 г за 129 грн заходить у нішу дешевше за Ottogi і Henan — "
              "і в категорії без жодного постійного постачальника-імпортера.")]
    for i, (t, c, body) in enumerate(FACTS):
        x = M + (i % 2) * 6.12
        y = 1.88 + (i // 2) * 1.14
        h = 1.04
        rect(s, x, y, 5.86, h, MIST, rounded=True, adj=0.08)
        rect(s, x, y, 0.09, h, c, rounded=True, adj=0.5)
        text(s, x + 0.28, y + 0.10, 5.30, 0.24, f"{i + 1}. {t}", size=11.5, bold=True,
             color=NAVY)
        text(s, x + 0.28, y + 0.38, 5.30, 0.60, body, size=8.5, color=INK, line=1.20)
    insight(s, M, 5.42, 11.98, 1.56, "Чого дані не показують",
            "1. Продажів. Ні мережі, ні маркетплейси не віддають обсяги.\n"
            "2. Обсягу українського виробництва. Верес, «МАКРО», «Маркел», «М'ясниця», "
            "Portion виробляють у країні — митниці вони не торкаються.\n"
            "3. Чому імпорт настільки малий. Митна база фіксує лише офіційні декларації: "
            "частина азійського асортименту може заходити дрібними партіями через інші "
            "коди чи посередників, яких наша вибірка не охоплює.", ORANGE)
    foot(s, f"{N_SKU} позицій, {N_BRANDS} брендів, 16 продавців. Митна база: код "
            "1904901000, 01.2025–05.2026. Зріз цін: 23–28.09.2026.")



# ══ ЗБІРКА (фінальна версія) ══════════════════════════════════════════════
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
slide_retail_audit()
slide_channels()
slide_suppliers()
slide_finmodel()
slide_recommendation()
slide_conclusions()

prs.save("Gotovyi_Rys_Final.pptx")
print("saved ·", len(prs.slides._sldIdLst), "slides ·", N_SKU, "SKU ·", N_BRANDS, "brands")
