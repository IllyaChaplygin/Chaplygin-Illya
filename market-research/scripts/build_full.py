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


# ══ ПОВНИЙ ПЕРЕЛІК SKU ════════════════════════════════════════════════════
# (файл пакшоту, назва, грамів, ціна_від, ціна_до, канал)
BENS = [
    ("md/bens_longgrain.jpg",   "Long Grain",        250, 45, 45,    "Сільпо"),
    ("md/bens_basmati250.jpg",  "Basmati",           250, 89.99, 89.99, "Сільпо"),
    ("md/bens_mediterran.jpg",  "Mediterran",        250, 89.99, 89.99, "Сільпо"),
    ("md/bens_risibisi.jpg",    "Risi Bisi",         250, 89.99, 89.99, "Сільпо"),
    ("md/bens_curry_indien.jpg","Indian Curry",      250, 89.99, 89.99, "Сільпо"),
    ("md/bens_curry_linsen.jpg","Curry з сочевицею", 220, 89.99, 89.99, "Сільпо"),
    ("md/bens_basmati220.jpg",  "Basmati",           220, 100, 144,   "MAUDAU · Сільпо"),
    ("md/bens_mexikanisch.jpg", "Mexikanisch",       220, 109, 109,   "Сільпо"),
    ("md/bens_curryreis.jpg",   "Curryreis Indien",  220, 144, 144,   "Сільпо"),
    ("md/bens_stickybowl.jpg",  "Для боулів",        220, 149, 149,   "Сільпо"),
    ("md/bens_sweetchili.jpg",  "Sweet Chili",       220, 179, 179,   "Сільпо"),
    ("md/bens_bio240.jpg",      "Bio Basmati",       240, 139, 139,   "Edison Lee"),
    ("md/bens_lang220.jpg",     "Long Grain",        220, 139, 139,   "Edison Lee"),
]
OTTOGI = [
    ("sku/ot_bulgogi.jpg",      "Бульгогі",            320, 252, 252, "Тайякі Март"),
    ("md/ottogi_hamburg.jpg",   "Гамбурзький стейк",   315, 255, 255, "Pulsar"),
    ("sku/ot_pork310.jpg",      "Свинина",             310, 252, 252, "Тайякі Март"),
    ("sku/ot_chicken_rib.jpg",  "Гострі курячі ребра", 310, 255, 260, "Pulsar · Gurmissimo"),
    ("sku/ot_octopus.jpg",      "Гострий восьминіг",   280, 255, 289, "Pulsar · Апетіт. · Gurm."),
    ("sku/ot_tuna.jpg",         "Тунець і майонез",    247, 225, 320, "Тайякі · Gurmissimo"),
    ("sku/ot_bibimbap.jpg",     "Пібімпаб",            269, 252, 334, "Тайякі · Апетітаріум"),
    ("sku/ot_pork_spicy.jpg",   "Гостра свинина",      269, 334, 334, "Апетітаріум"),
    ("sku/ot_beef320.jpg",      "Гостра яловичина",    320, 356, 356, "Апетітаріум"),
    ("md/ottogi_jjampong.jpg",  "Jjampong: рис у супі", 217.5, 255, 255, "Pulsar"),
]
# HENAN — так товар називають продавці; на упаковці марка Xiao Guo Zao (小锅造).
# Рис заливається окропом, 8 хвилин. Маса — сухого продукту.
HENAN = [
    ("sku/hn_braised174.jpg",       "Тушкована курка",         174, 129, 156, "СНЕКІС · DCM · Апетіт."),
    ("sku/hn_scallop174.jpg",       "Гребінці та гриби",       174, 140, 156, "MAUDAU · Gurm. · DCM"),
    ("sku/hn_sausage174.jpg",       "Кантонська ковбаска",     174, 118, 156, "OMG! Asia · Gurm. · Апетіт."),
    ("sku/hn_pepperbeef174.jpg",    "Яловичина з перцем",      174, 118, 156, "OMG! Asia · Gurm. · Апетіт."),
    ("sku/hn_pepperchicken174.jpg", "Курка з перцем чилі",     174, 112, 150, "OMG! · СНЕКІС · DCM"),
    ("sku/hn_spicychicken144.jpg",  "Гостра курка",            144, 83, 83,   "OMG! Asia"),
    ("sku/hn_sweetsourbeef144.jpg", "Кисло-солодка яловичина", 144, 135, 135, "DCM"),
    ("sku/hn_eggplant144.jpg",      "Курка з баклажанами",     144, 152, 152, "DCM"),
]
HAIDILAO = [
    ("sku/hd_stewed_chicken165.jpg", "Stewed Chicken*",       165, 380, 380, "Sweet Svitt"),
    ("sku/hd_beef_curry272.jpg",     "Яловичина з карі",      272, 770, 799, "Daruy"),
    ("sku/hd_beef_stew272.jpg",      "Тушкована яловичина",   272, 799, 799, "Daruy"),
    ("sku/hd_tomato_beef272.jpg",    "Томати й грудинка",     272, 799, 799, "Daruy"),
    ("sku/hd_canton187.jpg",         "Кантонський стиль",     187, 680, 680, "Daruy"),
    ("sku/hd_stew_chicken187.jpg",   "Тушкована курка",       187, 650, 680, "Daruy · Desna · Tactico"),
    ("sku/hd_pork_fish187.jpg",      "Свинина, рибний смак",  187, 650, 650, "Daruy"),
    ("sku/hd_teriyaki181.jpg",       "Курка теріякі",         181, 650, 660, "Daruy · Шериф"),
    ("sku/hd_spicy_chicken175.jpg",  "Гострий курячий соус",  175, 650, 650, "Daruy"),
    ("sku/hd_porkbelly360.jpg",      "Грудинка з гірчичною зеленню", 360, 1190, 1190, "Daruy"),
]
# (файл, бренд, назва, грамів, від, до, канал, колір)
READY_MORE = [
    ("md/bibigo_bowl.jpg", "Bibigo", "Білий рис", 210, 135, 189,
     "Смак Кореї · Rozetka · Prom", GREEN),
    ("md/clearspring.jpg", "Clearspring", "Brown & Wild Rice", 250, 252, 446,
     "MAUDAU · спец. магазини", TEAL),
    (None, "Portion", "Рис із куркою та овочами", 350, 109, 109, "Rozetka", ROSE),
    ("sku/mk_rice350.jpg", "Маркел", "Рис із соєвим м'ясом", 350, 94, 96,
     "СУХПАЙ · UPcompany", SLATE),
    ("sku/yt_beef84.jpg", "Gallina Blanca", "Yatekomo: яловичина", 84, 229, 229,
     "Товари з Іспанії", ORANGE),
    ("sku/yt_teriyaki84.jpg", "Gallina Blanca", "Yatekomo: теріякі", 84, 229, 229,
     "Товари з Іспанії", ORANGE),
]
SELFHEAT_MORE = [
    ("sku/mx_beef_bamboo.jpg", "Mo Xiao Xian", "Яловичина з бамбуком", 275, 931, 931,
     "Скарби Азії", PLUM),
    ("sku/mx_bacon_peas.jpg", "Mo Xiao Xian", "Бекон і горошок", 275, 931, 931,
     "Скарби Азії", PLUM),
    ("sku/qs_sh_pepperbeef.jpg", "Zihaiguo", "Яловичина з двома перцями", 440, 735, 735,
     "Скарби Азії", ROSE),
    ("sku/qs_sh_braisedpork.jpg", "Rongcheng Haoji", "Тушкована свинина", 440, 735, 735,
     "Скарби Азії", SLATE),
    ("sku/qs_spicybeef.jpg", "Qiaoshanmei", "Гостра яловичина", 146, 501, 501,
     "Скарби Азії", ORANGE),
    ("sku/qs_braisedpork.jpg", "Qiaoshanmei", "Свинина по-тайванськи", None, 501, 501,
     "Скарби Азії", ORANGE),
]
FD_IMPORT = [
    ("sku/tl_strog125.jpg", "Travellunch", "Бефстроганов", 125, 275, 515, "Гайдамака · ALANTUR", TEAL),
    ("sku/tl_beef_pepper125.jpg", "Travellunch", "Яловичина з перцем", 125, 295, 515, "Гайдамака · ALANTUR", TEAL),
    ("sku/tl_nasi125.jpg", "Travellunch", "Nasi Goreng", 125, 567, 567, "ALANTUR", TEAL),
    ("sku/tl_curry125.jpg", "Travellunch", "Курка карі з рисом", 125, 515, 515, "ALANTUR", TEAL),
    ("sku/tl_strog250.jpg", "Travellunch", "Бефстроганов", 250, 702, 811, "Highlander · ForCamp", TEAL),
    ("sku/tl_beef_pepper250.jpg", "Travellunch", "Яловичина з перцем", 250, 811, 811, "ALANTUR", TEAL),
    ("sku/tl_nasi250.jpg", "Travellunch", "Nasi Goreng", 250, 889, 889, "ALANTUR", TEAL),
    ("sku/te_fish.jpg", "Trek'n Eat", "Риба з рисом", None, 269, 269, "Freeride", PLUM),
    ("sku/te_strog.jpg", "Trek'n Eat", "Бефстроганов з рисом", None, 715, 715, "Kamanti", PLUM),
    ("sku/mh_curry.jpg", "Mountain House", "Жовте карі з куркою", 110, 699, 699, "110вольт", ROSE),
    ("sku/af_curry146.jpg", "Adventure Food", "Рис карі з фруктами", 146, 492, 492,
     "Клуб Мандрівник · ще 3", GREEN),
]
ADV_MENU = [  # 400 г — READY TO EAT (реторт), 110 г — сублімат LIGHTWEIGHT
    ("sku/am_korma400.jpg", "Adventure Menu", "Chicken Korma з рисом", 400, 376, 404,
     "ALANTUR · ForCamp · MK-Sport", GREEN),
    ("sku/am_wild.jpg", "Adventure Menu", "Курка в томатному соусі", 400, 315, 356,
     "Freeride · MK-Sport", GREEN),
    ("sku/am_meatballs.jpg", "Adventure Menu", "Фрикадельки з басматі", 400, 336, 336,
     "Лєєр", GREEN),
    ("sku/am_korma110.jpg", "Adventure Menu", "Korma з басматі · сублімат", 110, 378, 483,
     "Лєєр · OXO · MK · Freeride", TEAL),
    ("sku/am_tikka110.jpg", "Adventure Menu", "Tikka Masala · сублімат", 110, 390, 483,
     "Palmer · Лєєр · MK · Freeride", TEAL),
]
SUBLIMATE = [
    ("sku/sm_plov140.jpg", "SubliMate", "Плов узбецький", 140, 309, 310,
     "Highlander · ВсеОпт · Суренж", ROSE),
    ("sku/sm_chicken_fruit.jpg", "SubliMate", "Рис з куркою і фруктами", None, 289, 290,
     "Highlander · Суренж · ВсеОпт", ROSE),
    ("sku/sm_green_curry.jpg", "SubliMate", "Тайське зелене карі", None, 369, 370,
     "Highlander · ВсеОпт", ROSE),
]
FD_UA = [
    ("sku/jc_meatveg.jpg", "James Cook", "Рис з м'ясом та овочами", 90, 83, 138,
     "Highlander · ВсеОпт · Klever", GREEN),
    ("sku/jc_curry.jpg", "James Cook", "Карі з рисом та куркою", None, 133, 133,
     "Highlander · Activity", GREEN),
    ("sku/jc_mashkichiri.jpg", "James Cook", "Машкічірі: рис з бобами", None, 156, 156,
     "Highlander · Terra Incognita", GREEN),
    ("sku/jc_veg.jpg", "James Cook", "Рис з овочами", 90, 55, 81,
     "Highlander · ВсеОпт", GREEN),
    ("sku/yp_chicken85.jpg", "Їжа в Похід", "Рис з куркою та овочами", 85, 145, 145,
     "власний магазин", ORANGE),
    ("sku/yp_veg85.jpg", "Їжа в Похід", "Рис з овочами", 85, 130, 130,
     "власний магазин", ORANGE),
    ("sku/kh_pork.jpg", "Харчі", "Рис зі свининою", None, 120, 120,
     "Харчі ТМ · Висот-Нік", ROSE),
    ("sku/kh_ricemeat.jpg", "Харчі", "Рисова каша з м'ясом", None, 120, 120,
     "Highlander", ROSE),
    ("sku/kh_kichri.jpg", "Харчі", "Кічрі: боби з рисом", None, 83, 83,
     "Висот-Нік", ROSE),
    ("sku/kh_plovxl.jpg", "Харчі", "Плов XL", None, 321, 321,
     "8 продавців", ROSE),
    ("sku/fest_plov.jpg", "!FEST", "Плов", None, 133, 140,
     "SportStorm · Kalush-Craft", SLATE),
]

N_SKU = (len(BENS) + len(OTTOGI) + len(HENAN) + len(HAIDILAO) + len(READY_MORE) +
         len(SELFHEAT_MORE) + len(FD_IMPORT) + len(ADV_MENU) + len(SUBLIMATE) + len(FD_UA))
BRANDS = (["Ben's Original", "Ottogi", "Henan", "Haidilao"] +
          [b for _, b, *_ in READY_MORE + SELFHEAT_MORE + FD_IMPORT + ADV_MENU + SUBLIMATE + FD_UA])
N_BRANDS = len(dict.fromkeys(BRANDS))

# Готовий рис для розігріву — маса готової страви, ціни за 100 г порівнянні.
CORE = ([(g, lo) for _, _, g, lo, hi, _ in BENS] + [(g, lo) for _, _, g, lo, hi, _ in OTTOGI] +
        [(g, lo) for _, b, _, g, lo, hi, _, _ in READY_MORE if b != "Gallina Blanca"] +
        [(g, lo) for _, _, _, g, lo, hi, _, _ in ADV_MENU if g == 400])
CORE_MED = st.median(lo / g * 100 for g, lo in CORE)
N_CORE = len(CORE)

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



# ── допоміжне для сіток і ліг ──────────────────────────────────────────────
def as8(items, brand, c):
    """6-кортежі бренду → 8-кортежі (файл, бренд, назва, г, від, до, канал, колір)."""
    return [(im, brand, nm, g, lo, hi, ch, c) for im, nm, g, lo, hi, ch in items]


def grid(s, items, y, w, h, x0=M, gap=0.10, brand=True):
    for i, (im, b, nm, g, lo, hi, ch, c) in enumerate(items):
        sku_cell(s, x0 + i * (w + gap), y, w, h, im, nm, g, lo, hi, ch, c,
                 brand=b if brand else None)


def insight(s, x, y, w, h, title, body, c=ORANGE):
    rect(s, x, y, w, h, NAVY, rounded=True, adj=0.06)
    rect(s, x, y, 0.09, h, c, rounded=True, adj=0.5)
    text(s, x + 0.34, y + 0.22, w - 0.60, 0.30, title, size=13, bold=True, color=AMBER)
    text(s, x + 0.34, y + 0.62, w - 0.60, h - 0.80, body, size=10.5,
         color=RGBColor(0xD5, 0xDB, 0xEA), line=1.30)


def p100(items):
    v = [(lo / g * 100, hi / g * 100) for *_, g, lo, hi, _, _ in items if g]
    return min(a for a, _ in v), max(b for _, b in v)


ALL8 = (as8(BENS, "Ben's Original", ORANGE) + as8(OTTOGI, "Ottogi", PLUM) +
        as8(HENAN, "Henan", ROSE) + as8(HAIDILAO, "Haidilao", PLUM) + READY_MORE +
        SELFHEAT_MORE + FD_IMPORT + ADV_MENU + SUBLIMATE + FD_UA)
assert len(ALL8) == N_SKU
AM400 = [x for x in ADV_MENU if x[3] == 400]
AM110 = [x for x in ADV_MENU if x[3] == 110]
L_READY = (as8(BENS, "Ben's Original", ORANGE) + as8(OTTOGI, "Ottogi", PLUM) +
           [x for x in READY_MORE if x[1] != "Gallina Blanca"] + AM400)
L_INSTANT = (as8(HENAN, "Henan", ROSE) + [x for x in READY_MORE if x[1] == "Gallina Blanca"] +
             [x for x in SELFHEAT_MORE if x[1] == "Qiaoshanmei"])
L_HEAT = as8(HAIDILAO, "Haidilao", PLUM) + [x for x in SELFHEAT_MORE if x[1] != "Qiaoshanmei"]
L_FD = FD_IMPORT + AM110 + SUBLIMATE + FD_UA
assert len(L_READY) + len(L_INSTANT) + len(L_HEAT) + len(L_FD) == N_SKU
N_SILPO = sum(1 for x in ALL8 if "Сільпо" in x[6])


def brands_of(items):
    return list(dict.fromkeys(x[1] for x in items))


def pack_range(items):
    return min(x[4] for x in items), max(x[5] for x in items)


CH_CHAIN = ["Сільпо"]
CH_MARKET = ["MAUDAU", "Rozetka", "Prom"]
CH_FOOD = ["Смак Кореї", "Тайякі Март", "Pulsar", "Апетітаріум", "Gurmissimo", "Edison Lee",
           "OMG! Asia", "СНЕКІС", "DCM", "Скарби Азії", "Товари з Іспанії", "Sweet Svitt"]
CH_OUT = ["ALANTUR", "Highlander", "ForCamp", "Гайдамака", "Freeride", "MK-Sport", "Daruy",
          "Tactico", "Лєєр", "Activity", "OXO", "Palmer", "Kamanti", "Modern Shop", "110вольт",
          "Klever-Shop", "Terra Incognita", "Висот-Нік", "SportStorm", "Kalush-Craft", "M-LAW",
          "Суренж", "ВсеОпт", "СУХПАЙ", "UPcompany", "Desna", "Шериф", "Virnyy vybir",
          "Клуб Мандрівник", "Vkladovke", "Експрес Шоп", "Їжа в Похід", "Харчі ТМ", "Вояджер",
          "Непереможні", "Вартовий", "Драйв Сенс", "Ліхтар", "Military Style", "Націоналіст"]
N_CH = len(CH_CHAIN) + len(CH_MARKET) + len(CH_FOOD) + len(CH_OUT)

# ══ 01 · ОБКЛАДИНКА ═══════════════════════════════════════════════════════
s = slide(dark=True)
cover_pic(s, BG_TITLE, 0, 0, W, H)
s.shapes.add_picture(LOGO, Inches(M), Inches(0.52), Inches(1.94), Inches(0.84))
rect(s, M, 2.02, 0.54, 0.09, ORANGE)
text(s, M, 2.34, 7.00, 0.34, "ДОСЛІДЖЕННЯ РИНКУ · ВЕРЕСЕНЬ 2026", size=11.5,
     bold=True, color=AMBER)
text(s, M, 2.78, 7.20, 1.70, "Готовий рис\nв Україні", size=46, bold=True,
     color=WHITE, line=1.04)
text(s, M, 4.52, 6.80, 0.36, "Розбір кожного бренду: SKU, упаковка, ціна, канал",
     size=14.5, color=RGBColor(0xC6, 0xCE, 0xE2))
rect(s, M, 5.12, 5.90, 0.055, RGBColor(0x55, 0x60, 0x8C))
text(s, M, 5.34, 6.80, 0.80,
     f"{pl(N_BRANDS, 'бренд', 'бренди', 'брендів')} · {pl(N_SKU, 'позиція', 'позиції', 'позицій')} з пакшотом і ціною · {pl(N_CH, 'продавець', 'продавці', 'продавців')}\n"
     "Зріз українських онлайн-вітрин і мереж: 22–23 вересня 2026 р.",
     size=11.5, color=RGBColor(0x9F, 0xA9, 0xC4), line=1.40)
for im, x, y, w, h in [("md/bens_basmati250.jpg", 8.26, 0.96, 2.35, 2.40),
                       ("sku/hn_scallop174.jpg", 10.78, 0.96, 2.35, 2.40),
                       ("sku/hd_beef_stew272.jpg", 8.26, 3.50, 2.35, 2.10),
                       ("sku/ot_bibimbap.jpg",   10.78, 3.50, 2.35, 2.10)]:
    rect(s, x, y, w, h, WHITE, rounded=True, adj=0.06)
    pic(s, im, x + 0.16, y + 0.14, w - 0.32, h - 0.28)
rect(s, 8.26, 5.76, 4.87, 1.30, RGBColor(0x26, 0x2F, 0x4D), rounded=True, adj=0.12)
text(s, 8.50, 5.90, 4.44, 1.08, " · ".join(dict.fromkeys(BRANDS)),
     size=9, bold=True, color=WHITE, line=1.30)

# ══ 02 · КАРТА РИНКУ ══════════════════════════════════════════════════════
s = slide()
header(s, "КАРТА РИНКУ", f"{pl(N_BRANDS, 'бренд', 'бренди', 'брендів')}, чотири технології",
       "Групування за тим, як продукт готується. Кожен бренд далі розібрано окремо.")
LEAGUES = [("ГОТОВИЙ РИС · РОЗІГРІВ", L_READY, ORANGE,
            "Мікрохвильовка 60–90 с або реторт. Маса готової страви."),
           ("ЗАЛИТИ ОКРОПОМ", L_INSTANT, ROSE,
            "Сухий рис із соусом у чаші чи пакеті, 5–10 хв."),
           ("САМОРОЗІГРІВ", L_HEAT, PLUM,
            "Хімічний нагрівач у коробці, без техніки, 15 хв."),
           ("СУБЛІМАЦІЯ", L_FD, TEAL,
            "Сушіння виморожуванням, окріп. Туризм і армія.")]
for i, (kind, items, c, note) in enumerate(LEAGUES):
    x = M + i * 3.07
    lo, hi = pack_range(items)
    br = brands_of(items)
    rect(s, x, 1.94, 2.82, 2.86, MIST, rounded=True, adj=0.06)
    rect(s, x, 1.94, 2.82, 0.10, c, rounded=True, adj=0.5)
    text(s, x + 0.24, 2.16, 2.36, 0.24, kind, size=10, bold=True, color=c)
    text(s, x + 0.24, 2.44, 2.36, 0.40, f"{pl(len(items), 'позиція', 'позиції', 'позицій')} · {pl(len(br), 'бренд', 'бренди', 'брендів')}",
         size=13.5, bold=True, color=NAVY)
    text(s, x + 0.24, 2.84, 2.36, 0.74, " · ".join(br), size=8.5, color=INK, line=1.20)
    text(s, x + 0.24, 3.62, 2.36, 0.44, note, size=8.5, color=GREY, line=1.20)
    rect(s, x + 0.24, 4.14, 2.34, 0.02, MIST_D)
    text(s, x + 0.24, 4.24, 2.36, 0.22, "ЦІНА ЗА УПАКОВКУ", size=7.5, bold=True, color=GREY)
    text(s, x + 0.24, 4.44, 2.36, 0.30, f"{num(lo)} – {num(hi)} грн", size=14,
         bold=True, color=c)

rect(s, M, 5.00, 11.98, 1.94, NAVY, rounded=True, adj=0.07)
rect(s, M, 5.00, 0.10, 1.94, ORANGE, rounded=True, adj=0.5)
text(s, M + 0.42, 5.18, 5.40, 0.32, "Що показує розбір", size=15, bold=True, color=AMBER)
yy = 5.58
for t in [f"Ben's Original — {len(BENS)} позицій, єдиний бренд у мережі («Сільпо»).",
          f"Ottogi і Henan — {len(OTTOGI) + len(HENAN)} азійських рисових чаш; "
          "Henan удвічі дешевший за упаковку.",
          "Українське виробництво: Portion і Маркел (реторт 350 г) та 4 сублімати.",
          f"Саморозігрів — {num(pack_range(L_HEAT)[0])}–{num(pack_range(L_HEAT)[1])} грн "
          "за упаковку, найдорожча технологія."]:
    dot(s, M + 0.48, yy + 0.12, 0.11, AMBER)
    text(s, M + 0.70, yy, 5.60, 0.28, t, size=10.5, color=WHITE)
    yy += 0.31
rect(s, 6.90, 5.26, 0.04, 1.44, NAVY_L)
text(s, 7.24, 5.20, 5.10, 1.60,
     f"Медіана готового рису для розігріву ({N_CORE} позицій) — {num(CORE_MED)} грн "
     "за 100 г.\n"
     f"У мережі продається {N_SILPO} позицій із {N_SKU}. Решта {N_SKU - N_SILPO} — "
     "лише онлайн: маркетплейси, азійські фудшопи, туристичні й військові магазини.",
     size=10.5, color=RGBColor(0xB9, 0xC2, 0xDA), line=1.32)
foot(s, "Джерела: 17 мереж на zakaz.ua, «Сільпо», Prom, MAUDAU, Rozetka, спеціалізовані "
        "магазини — 22–23.09.2026.")

# ══ 03 · BEN'S ORIGINAL ══════════════════════════════════════════════════
lo100, hi100 = p100(as8(BENS, "", ORANGE))
s = slide()
header(s, "БРЕНД 1 · BEN'S ORIGINAL", f"{len(BENS)} позицій: пауч 220, 240 і 250 г",
       "Найглибша лінійка категорії і єдиний бренд, що дійшов до мережевої полиці.")
brand_head(s, M, 1.72, "Ben's Original", "Mars · ЄС", "пауч", len(BENS),
           lo100, hi100, 45, 179, ORANGE)
cw, ch = 1.62, 2.22
for i, (img, nm, g, lo, hi, chan) in enumerate(BENS[:7]):
    sku_cell(s, M + i * (cw + 0.10), 2.50, cw, ch, img, nm, g, lo, hi, chan, ORANGE)
for i, (img, nm, g, lo, hi, chan) in enumerate(BENS[7:]):
    sku_cell(s, M + i * (cw + 0.10), 4.78, cw, ch, img, nm, g, lo, hi, chan, ORANGE)
foot(s, "Джерела: каталог «Сільпо» (11 позицій), MAUDAU, Edison Lee — 22.09.2026. "
        "Basmati 220 г: 100 грн у MAUDAU і 144 грн у «Сільпо». На картках двох позицій "
        "Edison Lee — фото заводського ящика 6×; ціна й маса — за один пауч.")

# ══ 04 · OTTOGI ═══════════════════════════════════════════════════════════
lo100, hi100 = p100(as8(OTTOGI, "", PLUM))
s = slide()
header(s, "БРЕНД 2 · OTTOGI", f"{len(OTTOGI)} позицій: висока чаша 217–320 г",
       "Рис із наповнювачем для мікрохвильовки. Республіка Корея, чотири продавці.")
brand_head(s, M, 1.72, "Ottogi", "Республіка Корея", "чаша", len(OTTOGI),
           lo100, hi100, 225, 356, PLUM)
for i, (img, nm, g, lo, hi, chan) in enumerate(OTTOGI[:5]):
    sku_cell(s, M + i * 2.42, 2.50, 2.32, 2.22, img, nm, g, lo, hi, chan, PLUM)
for i, (img, nm, g, lo, hi, chan) in enumerate(OTTOGI[5:]):
    sku_cell(s, M + i * 2.42, 4.78, 2.32, 2.22, img, nm, g, lo, hi, chan, PLUM)
foot(s, "Джерела: Pulsar, Апетітаріум, Gurmissimo, Тайякі Март — 22.09.2026. "
        "Тунець 247 г: 225 грн у «Тайякі Март» і 320 грн у Gurmissimo — різниця 42 %.")

# ══ 05 · HENAN ════════════════════════════════════════════════════════════
lo100, hi100 = p100(as8(HENAN, "", ROSE))
s = slide()
header(s, "БРЕНД 3 · HENAN", f"{len(HENAN)} позицій: рисова чаша 144 і 174 г",
       "Китайський рис у чаші, заливається окропом на 8 хвилин. На упаковці — марка Xiao Guo Zao.")
brand_head(s, M, 1.72, "Henan", "Китай · Xiao Guo Zao", "чаша", len(HENAN),
           lo100, hi100, 83, 156, ROSE)
for i, (img, nm, g, lo, hi, chan) in enumerate(HENAN[:4]):
    sku_cell(s, M + i * 3.02, 2.50, 2.92, 2.22, img, nm, g, lo, hi, chan, ROSE)
for i, (img, nm, g, lo, hi, chan) in enumerate(HENAN[4:]):
    sku_cell(s, M + i * 3.02, 4.78, 2.92, 2.22, img, nm, g, lo, hi, chan, ROSE)
foot(s, "Джерела: MAUDAU, Апетітаріум, Gurmissimo, OMG! Asia, СНЕКІС, DCM — 22–23.09.2026. "
        "Маса — сухого продукту до заливання. Чаша Henan (83–156 грн) удвічі дешевша "
        "за Ottogi (225–356 грн).")

# ══ 06 · ЩЕ П'ЯТЬ БРЕНДІВ ГОТОВОГО РИСУ ═══════════════════════════════════
s = slide()
header(s, "БРЕНДИ 4–8", "Bibigo, Clearspring, Portion, Маркел, Gallina Blanca",
       "По одній-дві позиції в кожного: різні країни, різні формати, різні канали.")
grid(s, READY_MORE, 1.90, 1.90, 3.44, gap=0.116)
insight(s, M, 5.48, 11.98, 1.46, "Що це означає",
        "Portion і Маркел — українське реторт-виробництво 350 г за 94–109 грн: "
        "найдешевший грам категорії, 27–31 грн за 100 г. Bibigo — той самий товар у трьох "
        "каналах, розкид 40 %. Clearspring — органіка, найдорожчий грам серед паучів. "
        "Gallina Blanca Yatekomo — 84 г сухого рису в чаші, заливається окропом.")
foot(s, "Джерела: Смак Кореї, Rozetka, Prom, MAUDAU, СУХПАЙ, UPcompany, «Товари з Іспанії» — "
        "22–23.09.2026. Фото Portion: Rozetka не віддає зображення картки.")

# ══ 07 · HAIDILAO ═════════════════════════════════════════════════════════
lo100, hi100 = p100(as8(HAIDILAO, "", PLUM))
s = slide()
header(s, "БРЕНД 9 · HAIDILAO", f"{len(HAIDILAO)} позицій: саморозігрівальна коробка 165–360 г",
       "Хімічний нагрівач у коробці — готується без мікрохвильовки й окропу. Китай.")
brand_head(s, M, 1.72, "Haidilao", "Китай · саморозігрів", "коробка", len(HAIDILAO),
           lo100, hi100, 380, 1190, PLUM)
for i, (img, nm, g, lo, hi, chan) in enumerate(HAIDILAO[:5]):
    sku_cell(s, M + i * 2.42, 2.50, 2.32, 2.22, img, nm, g, lo, hi, chan, PLUM)
for i, (img, nm, g, lo, hi, chan) in enumerate(HAIDILAO[5:]):
    sku_cell(s, M + i * 2.42, 4.78, 2.32, 2.22, img, nm, g, lo, hi, chan, PLUM)
foot(s, "Джерела: Daruy, Tactico, Desna, Sweet Svitt, «Шериф» — 22–23.09.2026. "
        "Найдорожча упаковка категорії — 1 190 грн за 360 г. * Sweet Svitt: назва картки "
        "Stewed Chicken, на фото — рис з яловичиною та грибами.")

# ══ 08 · САМОРОЗІГРІВ І ІНСТАНТ-РИС: ЩЕ 4 БРЕНДИ ═════════════════════════
s = slide()
header(s, "БРЕНДИ 10–13", "Mo Xiao Xian, Zihaiguo, Rongcheng Haoji, Qiaoshanmei",
       "Китайський саморозігрів і рис швидкого приготування. Усе — в одному магазині.")
grid(s, SELFHEAT_MORE, 1.90, 1.90, 3.44, gap=0.116)
insight(s, M, 5.48, 11.98, 1.46, "Що це означає",
        "Zihaiguo і Rongcheng Haoji — коробка 440 г за 735 грн: 167 грн за 100 г, грам "
        "дешевший, ніж у Haidilao (230–371). Mo Xiao Xian — 931 грн за 275 г. "
        "Qiaoshanmei — не саморозігрів, а рис у пакеті під окріп, 501 грн. "
        "Усі чотири бренди в Україні продає один магазин — «Скарби Азії».", PLUM)
foot(s, "Джерело: «Скарби Азії» (Prom), 23.09.2026. Бренди Zihaiguo і Rongcheng Haoji визначено "
        "за написами на упаковці, Qiaoshanmei — за карткою продавця.")

# ══ 09 · СУБЛІМАЦІЯ · ІМПОРТ ══════════════════════════════════════════════
s = slide()
header(s, "БРЕНДИ 14–17 · СУБЛІМАЦІЯ, ІМПОРТ",
       "Travellunch, Trek'n Eat, Mountain House, Adventure Food",
       "Рис-страви у дойпаку: залити окропом на 8–10 хвилин. Німеччина, США, Нідерланди.")
grid(s, FD_IMPORT[:7], 1.90, 1.62, 2.50)
grid(s, FD_IMPORT[7:], 4.50, 1.62, 2.50)
insight(s, M + 4 * 1.72, 4.50, 11.98 - 4 * 1.72, 2.50, "Що це означає",
        "Travellunch — 7 позицій у двох вагах, 125 і 250 г, найширша імпортна лінійка. "
        "Mountain House — найдорожчий грам категорії: 635 грн за 100 г сухої маси. "
        "Канал — туристичні й військові магазини.", TEAL)
foot(s, "Джерела: ALANTUR, Highlander, ForCamp, Військторг Гайдамака, Freeride, Kamanti, "
        "110вольт, Клуб Мандрівник — 22–23.09.2026.")

# ══ 10 · ADVENTURE MENU + SUBLIMATE ═══════════════════════════════════════
s = slide()
header(s, "БРЕНДИ 18–19", "Adventure Menu і SubliMate",
       "Adventure Menu — єдиний бренд із двома технологіями: готова страва 400 г і сублімат 110 г.")
grid(s, ADV_MENU, 1.90, 2.30, 2.50, gap=0.12)
grid(s, SUBLIMATE, 4.50, 2.30, 2.50, gap=0.12)
insight(s, M + 3 * 2.42, 4.50, 11.98 - 3 * 2.42, 2.50, "Що це означає",
        "Лінія Adventure Menu 400 г READY TO EAT — реторт без води, 79–101 грн за 100 г: "
        "у тому ж коридорі, що Ben's Original і Ottogi, але продається лише в туристичних "
        "магазинах. SubliMate — 289–370 грн за упаковку.", GREEN)
foot(s, "Джерела: ALANTUR, ForCamp, MK-Sport, Freeride, Лєєр, OXO, Palmer, Highlander, "
        "ВсеОпт, Суренж — 22–23.09.2026. Маса двох SubliMate на картках не вказана.")

# ══ 11 · СУБЛІМАЦІЯ · УКРАЇНА ═════════════════════════════════════════════
fd_lo, fd_hi = pack_range(FD_UA)
im_lo, im_hi = pack_range(FD_IMPORT)
s = slide()
header(s, "БРЕНДИ 20–23 · СУБЛІМАЦІЯ, УКРАЇНА", "James Cook, Їжа в Похід, Харчі, !FEST",
       "Українські виробники рисових страв у дойпаку. Канал — туристичні й військові магазини.")
grid(s, FD_UA[:7], 1.90, 1.62, 2.50)
grid(s, FD_UA[7:], 4.50, 1.62, 2.50)
insight(s, M + 4 * 1.72, 4.50, 11.98 - 4 * 1.72, 2.50, "Що це означає",
        f"Українські сублімати — {num(fd_lo)}–{num(fd_hi)} грн за упаковку, імпортні — "
        f"{num(im_lo)}–{num(im_hi)} грн. James Cook — 4 рисові позиції, найширша українська "
        "лінійка. Маса на більшості карток не вказана.", ORANGE)
foot(s, "Джерела: Highlander, Activity, ВсеОпт, Klever-Shop, Terra Incognita, «Їжа в Похід», "
        "Харчі ТМ, Висот-Нік, SportStorm, Kalush-Craft — 22–23.09.2026.")

# ══ 12 · УПАКОВКА ═════════════════════════════════════════════════════════
P_POUCH = (as8(BENS, "Ben's Original", ORANGE) +
           [x for x in READY_MORE if x[1] in ("Clearspring", "Portion", "Маркел")] + AM400)
P_BOWL = (as8(OTTOGI, "Ottogi", PLUM) + as8(HENAN, "Henan", ROSE) +
          [x for x in READY_MORE if x[1] in ("Bibigo", "Gallina Blanca")])
P_BOX = L_HEAT
P_DOY = L_FD + [x for x in SELFHEAT_MORE if x[1] == "Qiaoshanmei"]
assert len(P_POUCH) + len(P_BOWL) + len(P_BOX) + len(P_DOY) == N_SKU


def mass_range(items):
    g = [x[3] for x in items if x[3]]
    return f"{num(min(g))} – {num(max(g))} г"


s = slide()
header(s, "ЯКА УПАКОВКА", "Чотири типи паковання",
       "Тип паковання визначає і спосіб приготування, і полицю, на яку товар потрапляє.")
TYPES = [("md/bens_basmati250.jpg", "ПАУЧ · РЕТОРТ", P_POUCH, ORANGE),
         ("sku/hn_scallop174.jpg", "ЧАША", P_BOWL, ROSE),
         ("sku/hd_beef_stew272.jpg", "КОРОБКА З НАГРІВАЧЕМ", P_BOX, PLUM),
         ("sku/tl_strog125.jpg", "ДОЙПАК", P_DOY, TEAL)]
for i, (img, kind, items, c) in enumerate(TYPES):
    x = M + i * 3.07
    br = brands_of(items)
    rect(s, x, 1.94, 2.82, 3.46, WHITE, line=MIST_D, lw=1.2, rounded=True, adj=0.06)
    rect(s, x, 1.94, 2.82, 0.10, c, rounded=True, adj=0.5)
    rect(s, x + 0.16, 2.16, 2.50, 1.20, MIST, rounded=True, adj=0.06)
    pic(s, img, x + 0.26, 2.22, 2.30, 1.08)
    text(s, x + 0.26, 3.46, 2.30, 0.26, kind, size=11.5, bold=True, color=c)
    text(s, x + 0.26, 3.74, 2.30, 0.26, mass_range(items), size=12, bold=True, color=NAVY)
    text(s, x + 0.26, 4.04, 2.30, 0.80,
         " · ".join(br[:5]) + (f" · ще {len(br) - 5}" if len(br) > 5 else ""),
         size=8.5, color=GREY, line=1.20)
    pill(s, x + 0.26, 4.96, f"{pl(len(items), 'позиція', 'позиції', 'позицій')} · {pl(len(br), 'бренд', 'бренди', 'брендів')}", c, size=8)

BUCKETS = [("до 100 г", 0, 100, TEAL), ("110–150 г", 101, 150, TEAL),
           ("165–187 г", 151, 200, PLUM), ("210–250 г", 201, 255, ORANGE),
           ("269–280 г", 256, 300, PLUM), ("310–360 г", 301, 370, PLUM),
           ("400–440 г", 371, 450, GREEN)]
MASSES = [x[3] for x in ALL8]
BARS = [(lb, sum(1 for g in MASSES if g and a <= g <= b), c) for lb, a, b, c in BUCKETS]
BARS.append(("маса н/д", sum(1 for g in MASSES if not g), GREY))
assert sum(n for _, n, _ in BARS) == N_SKU, BARS
text(s, M, 5.58, 4.20, 0.24, "ПОЗИЦІЙ У КОЖНІЙ ВАЗІ", size=9, bold=True, color=GREY)
bx = M
for lb, n, c in BARS:
    hgt = 0.06 + n * 0.030
    rect(s, bx, 6.66 - hgt, 1.06, hgt, c, rounded=True, adj=0.16)
    text(s, bx, 6.70, 1.06, 0.22, lb, size=7.5, bold=True, color=INK, align=PP_ALIGN.CENTER)
    text(s, bx, 6.66 - hgt - 0.24, 1.06, 0.22, str(n), size=9.5, bold=True, color=c,
         align=PP_ALIGN.CENTER)
    bx += 1.40
foot(s, f"Джерело: маса за картками товарів, 22–23.09.2026. Разом {N_SKU} позиції.")

# ══ 13 · ЦІНА ЗА 100 Г ════════════════════════════════════════════════════
def corridor(items, label, c):
    v = [(lo / g * 100, hi / g * 100) for *_, g, lo, hi, _, _ in items if g]
    return (label, min(a for a, _ in v), max(b for _, b in v), c)


def by_brand(items, b):
    return [x for x in items if x[1] == b]


LEFT = sorted([
    corridor(as8(BENS, "", ORANGE), "Ben's Original", ORANGE),
    corridor(as8(OTTOGI, "", PLUM), "Ottogi", PLUM),
    corridor(by_brand(READY_MORE, "Bibigo"), "Bibigo", GREEN),
    corridor(by_brand(READY_MORE, "Clearspring"), "Clearspring", TEAL),
    corridor(by_brand(READY_MORE, "Portion"), "Portion", ROSE),
    corridor(by_brand(READY_MORE, "Маркел"), "Маркел", SLATE),
    corridor(AM400, "Adventure Menu 400 г", GREEN),
    corridor(as8(HAIDILAO, "", PLUM), "Haidilao", PLUM),
    corridor(by_brand(SELFHEAT_MORE, "Mo Xiao Xian"), "Mo Xiao Xian", PLUM),
    corridor(by_brand(SELFHEAT_MORE, "Zihaiguo"), "Zihaiguo", ROSE),
    corridor(by_brand(SELFHEAT_MORE, "Rongcheng Haoji"), "Rongcheng Haoji", SLATE),
], key=lambda r: r[1])
RIGHT = sorted([
    corridor(as8(HENAN, "", ROSE), "Henan", ROSE),
    corridor(by_brand(READY_MORE, "Gallina Blanca"), "Gallina Blanca", ORANGE),
    corridor(by_brand(SELFHEAT_MORE, "Qiaoshanmei"), "Qiaoshanmei", ORANGE),
    corridor(by_brand(FD_UA, "James Cook"), "James Cook", GREEN),
    corridor(by_brand(FD_UA, "Їжа в Похід"), "Їжа в Похід", ORANGE),
    corridor(by_brand(SUBLIMATE, "SubliMate"), "SubliMate", ROSE),
    corridor(by_brand(FD_IMPORT, "Travellunch"), "Travellunch", TEAL),
    corridor(by_brand(FD_IMPORT, "Adventure Food"), "Adventure Food", GREEN),
    corridor(AM110, "Adventure Menu 110 г", TEAL),
    corridor(by_brand(FD_IMPORT, "Mountain House"), "Mountain House", ROSE),
], key=lambda r: r[1])


def panel(s, x, w, rows, vmax, step, title, sub, med=None):
    rect(s, x, 1.90, w, 4.98, MIST, rounded=True, adj=0.04)
    text(s, x + 0.24, 2.04, w - 0.48, 0.26, title, size=11.5, bold=True, color=NAVY)
    text(s, x + 0.24, 2.30, w - 0.48, 0.22, sub, size=8.5, color=GREY)
    ax0, axw = x + 1.86, w - 2.46

    def ax(v):
        return ax0 + axw * v / vmax
    top, bot = 2.66, 6.34
    for t in range(0, vmax + 1, step):
        rect(s, ax(t), top, 0.012, bot - top, RGBColor(0xDC, 0xE2, 0xEE))
        text(s, ax(t) - 0.30, bot + 0.06, 0.60, 0.22, str(t), size=8, color=GREY,
             align=PP_ALIGN.CENTER)
    if med:
        rect(s, ax(med), top, 0.025, bot - top, ORANGE)
        text(s, ax(med) + 0.06, top - 0.20, 1.60, 0.20, f"медіана {num(med)}",
             size=7.5, bold=True, color=ORANGE)
    gap = (bot - top - 0.20) / len(rows)
    yy = top + 0.20 + gap / 2
    for lb, lo, hi, c in rows:
        text(s, x + 0.24, yy - 0.12, 1.60, 0.24, lb, size=9, bold=True, color=NAVY)
        x0, x1 = ax(lo), ax(hi)
        rect(s, x0, yy - 0.05, max(x1 - x0, 0.02), 0.10, c, rounded=True, adj=0.5)
        dot(s, x0, yy, 0.15, c)
        text(s, x0 - 0.56, yy - 0.11, 0.46, 0.22, num(lo), size=8.5, bold=True, color=c,
             align=PP_ALIGN.RIGHT)
        if x1 - x0 >= 0.10:
            dot(s, x1, yy, 0.15, c)
            text(s, x1 + 0.10, yy - 0.11, 0.50, 0.22, num(hi), size=8.5, bold=True, color=c)
        yy += gap


s = slide()
header(s, "ЯКА ЦІНА", "Коридор кожного бренду за 100 грамів",
       "Смуга — від найдешевшої до найдорожчої позиції. Готову і суху масу порівнювати не можна.")
panel(s, M, 5.90, LEFT, 400, 50, "Маса готової страви",
      "Розігрів, реторт, саморозігрів", med=CORE_MED)
panel(s, M + 6.08, 5.90, RIGHT, 700, 100, "Суха маса — заливається окропом",
      "Інстант-рис і сублімати")
foot(s, f"Розрахунок за масою та ціною карток, 22–23.09.2026. Медіана готового рису для "
        f"розігріву ({N_CORE} позицій) — {num(CORE_MED)} грн за 100 г. Позиції без маси "
        "на картці до розрахунку не входять.")

# ══ 14 · ДЕ ПРЕДСТАВЛЕНО ══════════════════════════════════════════════════
s = slide()
header(s, "ДЕ ПРЕДСТАВЛЕНО", f"{N_CH} продавців, чотири типи каналів",
       f"Категорія живе в онлайні: у мережі — лише Ben's Original у «Сільпо», {N_SILPO} позицій.")
GROUPS = [("МЕРЕЖЕВИЙ РОЗДРІБ", ORANGE, CH_CHAIN, "Ben's Original"),
          ("ОНЛАЙН-СУПЕРМАРКЕТИ Й МАРКЕТПЛЕЙСИ", GREEN, CH_MARKET,
           "Ben's · Henan · Bibigo · Clearspring · Portion"),
          ("СПЕЦІАЛІЗОВАНІ ФУДШОПИ", PLUM, CH_FOOD,
           "Ottogi · Henan · Bibigo · саморозігрів · Yatekomo"),
          ("ТУРИСТИЧНІ ТА ВІЙСЬКОВІ", TEAL, CH_OUT,
           "Haidilao · Adventure Menu · усі сублімати · Маркел")]
yy = 1.92
for gi, (title, c, chans, what) in enumerate(GROUPS):
    rh = 0.82 if gi < 2 else 1.18
    rect(s, M, yy, 11.98, rh, MIST if gi % 2 == 0 else WHITE, rounded=True, adj=0.12)
    rect(s, M, yy, 0.09, rh, c, rounded=True, adj=0.5)
    text(s, M + 0.30, yy + 0.14, 3.60, 0.24, title, size=9.5, bold=True, color=c)
    text(s, M + 0.30, yy + 0.42, 3.60, 0.60, what, size=9, color=GREY, line=1.20)
    shown = chans if len(chans) <= 12 else chans[:10]
    cx, cy = M + 4.00, yy + 0.16
    for ch in shown + ([f"+ ще {len(chans) - 10}"] if len(chans) > 12 else []):
        w = 0.24 + 0.068 * len(ch)
        if cx + w > M + 11.00:
            cx, cy = M + 4.00, cy + 0.36
        pill(s, cx, cy, ch, c, w=w, size=8.5)
        cx += w + 0.12
    text(s, M + 11.20, yy + rh / 2 - 0.18, 0.62, 0.36, str(len(chans)), size=17, bold=True,
         color=c, align=PP_ALIGN.RIGHT)
    yy += rh + 0.08
text(s, M, 6.36, 11.98, 0.60,
     "Мережі без жодної порівнянної позиції: АТБ, Novus, Metro, Varus, Auchan, Fozzy, "
     "«Таврія В», МегаМаркет, ЕКО маркет, Ultramarket та ще 7 мереж на zakaz.ua. "
     "Перевірено онлайн 22–23.09.2026, без обходу полиці.",
     size=9, color=GREY, line=1.24)

prs.save("Gotovyi_Rys_Rozbir_Brendiv.pptx")
print("saved ·", len(prs.slides._sldIdLst), "slides ·", N_SKU, "SKU ·", N_BRANDS, "brands ·",
      N_CH, "channels · median", round(CORE_MED, 1))
