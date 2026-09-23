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
    ("sku/bens_bio240_single.jpg","Bio Basmati",      240, 139, 139,   "Edison Lee"),
    ("sku/bens_lang220_single.jpg","Long Grain",      220, 139, 139,   "Edison Lee"),
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
    ("sku/qs_braisedpork.jpg", "Qiaoshanmei", "Свинина по-тайванськи", 146, 501, 501,
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
    ("sku/mh_friedrice.jpg", "Mountain House", "Chicken Fried Rice", 133, 699, 699,
     "110вольт", ROSE),
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
    ("sku/sm_chicken_fruit.jpg", "SubliMate", "Рис з куркою і фруктами", 110, 290, 290,
     "Highlander · Суренж · ВсеОпт", ROSE),
    ("sku/sm_green_curry.jpg", "SubliMate", "Тайське зелене карі", 120, 370, 370,
     "Highlander · ВсеОпт", ROSE),
]
FD_UA = [
    ("sku/jc_meatveg.jpg", "James Cook", "Рис з м'ясом та овочами", 90, 83, 138,
     "Highlander · ВсеОпт · Klever", GREEN),
    ("sku/jc_curry.jpg", "James Cook", "Карі з рисом та куркою", 90, 133, 133,
     "Highlander · Activity", GREEN),
    ("sku/jc_mashkichiri.jpg", "James Cook", "Машкічірі: рис з бобами", 80, 156, 156,
     "Highlander · Terra Incognita", GREEN),
    ("sku/jc_veg.jpg", "James Cook", "Рис з овочами", 90, 55, 81,
     "Highlander · ВсеОпт", GREEN),
    ("sku/yp_chicken85.jpg", "Їжа в Похід", "Рис з куркою та овочами", 85, 145, 145,
     "власний магазин", ORANGE),
    ("sku/yp_veg85.jpg", "Їжа в Похід", "Рис з овочами", 85, 130, 130,
     "власний магазин", ORANGE),
    ("sku/kh_pork.jpg", "Харчі", "Рис зі свининою", 85, 120, 120,
     "Харчі ТМ · Висот-Нік", ROSE),
    ("sku/kh_ricemeat.jpg", "Харчі", "Рисова каша з м'ясом", 85, 120, 120,
     "Highlander", ROSE),
    ("sku/kh_kichri.jpg", "Харчі", "Кічрі: боби з рисом", 85, 83, 83,
     "Висот-Нік", ROSE),
    ("sku/kh_plovxl.jpg", "Харчі", "Плов XL", 100, 321, 321,
     "8 продавців", ROSE),
    ("sku/fest_plov.jpg", "!FEST", "Плов", 100, 133, 140,
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

# ── Тип паковання для кожної позиції (за картками продавців) ───────────────
#  pouch — плаский пауч/реторт-пакет, cup — жорстка чаша чи стакан,
#  box — картонна коробка з нагрівачем, doypack — дойпак під окріп, can — бляшанка
PACK = {}
for _, nm, g, *_ in BENS:              PACK[("Ben's Original", nm, g)] = "pouch"
for _, nm, g, *_ in OTTOGI:            PACK[("Ottogi", nm, g)] = "cup"
for _, nm, g, *_ in HENAN:             PACK[("Henan", nm, g)] = "cup"
for _, nm, g, *_ in HAIDILAO:          PACK[("Haidilao", nm, g)] = "box"
# Adventure Menu: 400 г — реторт-пауч, 110 г — дойпак; вирішується за масою.
PACK_BY_BRAND = {"Adventure Menu": None, "Bibigo": "cup", "Clearspring": "pouch", "Portion": "pouch", "Маркел": "pouch",
                 "Gallina Blanca": "cup", "Mo Xiao Xian": "cup", "Zihaiguo": "box",
                 "Rongcheng Haoji": "box", "Qiaoshanmei": "doypack", "Travellunch": "doypack",
                 "Trek'n Eat": "doypack", "Mountain House": "doypack", "Adventure Food": "doypack",
                 "SubliMate": "doypack", "James Cook": "doypack", "Їжа в Похід": "doypack",
                 "Харчі": "doypack", "!FEST": "doypack"}


def pack_of(brand, name, grams):
    """Тип паковання позиції."""
    if brand == "Adventure Menu":
        return "pouch" if grams == 400 else "doypack"
    return PACK.get((brand, name, grams)) or PACK_BY_BRAND[brand]


PACK_RU = {"pouch": "ПАУЧ · РЕТОРТ", "cup": "ЧАША · СТАКАН",
           "box": "КОРОБКА З НАГРІВАЧЕМ", "doypack": "ДОЙПАК"}

# ── Полиця мереж: суміжна категорія, яку ми НЕ рахуємо в 89 позицій ───────
# (файл, бренд, назва, г, ціна_від, ціна_до, мережі, колір)
CHAIN_SHELF = [
    ("sku/ch_hapay_rice.jpg", "hapay!", "Каша рисова зі свининою", 340, 62.9, 62.9, "Ашан", GREEN),
    ("sku/ch_lappetit_pork.jpg", "L'appetit", "Каша рисова зі свининою", 340, 119.9, 119.9, "Ашан", ORANGE),
    ("sku/ch_lappetit_beef.jpg", "L'appetit", "Каша рисова з яловичиною", 340, 125.9, 125.9, "Ашан", ORANGE),
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



# ══ РОЗРАХУНКИ ════════════════════════════════════════════════════════════
import collections


def as8(items, brand, c):
    return [(im, brand, nm, g, lo, hi, ch, c) for im, nm, g, lo, hi, ch in items]


ALL8 = (as8(BENS, "Ben's Original", ORANGE) + as8(OTTOGI, "Ottogi", PLUM) +
        as8(HENAN, "Henan", ROSE) + as8(HAIDILAO, "Haidilao", PLUM) + READY_MORE +
        SELFHEAT_MORE + FD_IMPORT + ADV_MENU + SUBLIMATE + FD_UA)
assert len(ALL8) == N_SKU

COUNTRY = {"Ben's Original": "ЄС", "Ottogi": "Корея", "Henan": "Китай", "Haidilao": "Китай",
           "Bibigo": "Корея", "Clearspring": "ЄС", "Portion": "Україна", "Маркел": "Україна",
           "Gallina Blanca": "ЄС", "Mo Xiao Xian": "Китай", "Zihaiguo": "Китай",
           "Rongcheng Haoji": "Китай", "Qiaoshanmei": "Китай", "Travellunch": "ЄС",
           "Trek'n Eat": "ЄС", "Mountain House": "США", "Adventure Food": "ЄС",
           "Adventure Menu": "ЄС", "SubliMate": "Україна", "James Cook": "Україна",
           "Їжа в Похід": "Україна", "Харчі": "Україна", "!FEST": "Україна"}

BY_PACK = collections.defaultdict(list)
for _x in ALL8:
    BY_PACK[pack_of(_x[1], _x[2], _x[3])].append(_x)
PACK_ORDER = ["doypack", "cup", "pouch", "box"]
BY_COUNTRY = collections.Counter(COUNTRY[x[1]] for x in ALL8)

AM400 = [x for x in ADV_MENU if x[3] == 400]
AM110 = [x for x in ADV_MENU if x[3] == 110]
L_READY = (as8(BENS, "Ben's Original", ORANGE) + as8(OTTOGI, "Ottogi", PLUM) +
           [x for x in READY_MORE if x[1] != "Gallina Blanca"] + AM400)
L_INSTANT = (as8(HENAN, "Henan", ROSE) + [x for x in READY_MORE if x[1] == "Gallina Blanca"] +
             [x for x in SELFHEAT_MORE if x[1] == "Qiaoshanmei"])
L_HEAT = as8(HAIDILAO, "Haidilao", PLUM) + [x for x in SELFHEAT_MORE if x[1] != "Qiaoshanmei"]
L_FD = [x for x in FD_IMPORT if x[1] != "Mountain House"] + \
       [x for x in FD_IMPORT if x[1] == "Mountain House"] + AM110 + SUBLIMATE + FD_UA
assert len(L_READY) + len(L_INSTANT) + len(L_HEAT) + len(L_FD) == N_SKU
N_SILPO = sum(1 for x in ALL8 if "Сільпо" in x[6])
N_UA = sum(1 for x in ALL8 if COUNTRY[x[1]] == "Україна")

# ── Імпорт готового рису (CN 1904 90 10 «Рис, приготовлений») ─────────────
# Дзеркальна статистика ЄС (Eurostat Comext, експорт ЄС→Україна) + UN Comtrade
# для Азії. Дзеркало точніше за українську митницю: воно не має прогалин.
IMPORT = [  # рік, тонн з ЄС, тис. € з ЄС
    (2019, 99.1, 294), (2020, 204.4, 553), (2021, 204.0, 567), (2022, 127.6, 434),
    (2023, 125.8, 428), (2024, 111.0, 475), (2025, 179.4, 823),
]
IMPORT_PARTNERS_2025 = [("Болгарія", 78.3, 527), ("Польща", 94.4, 239), ("Італія", 5.6, 46),
                        ("Іспанія", 0.4, 3), ("Франція", 0.4, 1)]
ASIA_2025 = [("Корея", 2.04, 10.2), ("Китай", 1.04, 3.6), ("Таїланд", 0.02, 0.06)]  # т, тис. $
EUR_UAH, USD_UAH = 47.15, 41.71          # НБУ, середній за 2025 р.
T_EU, V_EU = IMPORT[-1][1], IMPORT[-1][2]
T_ASIA = sum(t for _, t, _ in ASIA_2025)
T_TOTAL = T_EU + T_ASIA
CIF_UAH = (V_EU * EUR_UAH + sum(v for _, _, v in ASIA_2025) * USD_UAH) / 1000  # млн грн
AVG_PACK_G = 220                          # медіана маси паучів і чаш
PACKS_K = T_TOTAL * 1000_000 / AVG_PACK_G / 1000        # тис. упаковок на рік
RETAIL_LO, RETAIL_HI = CIF_UAH * 2.2, CIF_UAH * 2.8          # роздріб, млн грн
POP_M = 29.0                              # млн осіб, підконтрольна територія
EU_PER_CAP_USD = 1.43                     # ринок instant rice ЄС / населення ЄС
UA_PER_CAP_USD = RETAIL_LO * 1e6 / USD_UAH / (POP_M * 1e6)
GAP = EU_PER_CAP_USD / UA_PER_CAP_USD

# ── спільні примітиви аналітичних слайдів ─────────────────────────────────
def kpi(s, x, y, w, h, label, value, note, c):
    rect(s, x, y, w, h, MIST, rounded=True, adj=0.08)
    rect(s, x, y, w, 0.09, c, rounded=True, adj=0.5)
    text(s, x + 0.24, y + 0.26, w - 0.48, 0.22, label, size=8.5, bold=True, color=GREY)
    text(s, x + 0.24, y + 0.50, w - 0.48, 0.44, value, size=24, bold=True, color=c)
    text(s, x + 0.24, y + 1.00, w - 0.48, h - 1.10, note, size=9, color=INK, line=1.24)


def insight(s, x, y, w, h, title, body, c=ORANGE):
    rect(s, x, y, w, h, NAVY, rounded=True, adj=0.06)
    rect(s, x, y, 0.09, h, c, rounded=True, adj=0.5)
    text(s, x + 0.34, y + 0.20, w - 0.60, 0.28, title, size=12.5, bold=True, color=AMBER)
    text(s, x + 0.34, y + 0.56, w - 0.60, h - 0.72, body, size=10,
         color=RGBColor(0xD5, 0xDB, 0xEA), line=1.30)


def grid(s, items, y, w, h, x0=M, gap=0.10, brand=True):
    for i, (im, b, nm, g, lo, hi, ch, c) in enumerate(items):
        sku_cell(s, x0 + i * (w + gap), y, w, h, im, nm, g, lo, hi, ch, c,
                 brand=b if brand else None)


def brands_of(items):
    return list(dict.fromkeys(x[1] for x in items))


def pack_range(items):
    return min(x[4] for x in items), max(x[5] for x in items)


def p100(items):
    v = [(lo / g * 100, hi / g * 100) for *_, g, lo, hi, _, _ in items if g]
    return min(a for a, _ in v), max(b for _, b in v)


def unit_med(items):
    return st.median([x[4] for x in items])


# ══ ГОЛОВНЕ ═══════════════════════════════════════════════════════════════
def slide_summary():
    s = slide()
    header(s, "ГОЛОВНЕ", "Шість цифр, які описують ринок",
           "Усі ціни в колоді — за одну упаковку. Опт і ящики до розрахунку не входять.")
    K = [("ОБСЯГ ІМПОРТУ 2025", f"{num(T_TOTAL)} т", ORANGE,
          f"Готовий рис, CN 1904 90 10.\n+{num((T_TOTAL / (IMPORT[-2][1] + 3.5) - 1) * 100)} % до 2024 року."),
         ("ЄМНІСТЬ РОЗДРІБУ", f"{num(RETAIL_LO)}–{num(RETAIL_HI)}", GREEN,
          "млн грн на рік — оцінка\nза імпортом × 2,2–2,8."),
         ("ВІДСТАВАННЯ ВІД ЄС", f"×{num(GAP)}", ROSE,
          f"{num(UA_PER_CAP_USD, 2)} $ на особу проти\n1,43 $ у ЄС. Ринок на старті."),
         ("ПОЗИЦІЙ У ПРОДАЖУ", f"{N_SKU}", PLUM,
          f"{N_BRANDS} брендів, {N_CH} продавців.\nУ мережі — {N_SILPO} позицій."),
         ("НАЙПОШИРЕНІШИЙ ФОРМАТ", "дойпак", TEAL,
          f"{len(BY_PACK['doypack'])} позицій ({num(len(BY_PACK['doypack']) / N_SKU * 100)} %),\n"
          "але це туризм, не полиця."),
         ("МАСОВА ВАГА", "250 г", SLATE,
          "9 позицій. Далі 220 г — 7,\n174 г — 5, 85 г — 5.")]
    for i, (lb, v, c, note) in enumerate(K):
        kpi(s, M + i * 2.04, 1.88, 1.86, 1.94, lb, v, note, c)
    insight(s, M, 4.06, 5.86, 2.34, "Що це означає для полиці",
            "1. Категорія імпортується, але не потрапляє в мережі: 11 позицій із 84 — і всі "
            "це Ben's Original у «Сільпо».\n"
            "2. Зростання 2025 року дає Болгарія: 44,8 → 78,3 т за рік, +75 %.\n"
            "3. Формат, який реально працює на полиці, — пауч 220–250 г: медіана 109 грн за "
            "упаковку, найнижча в категорії.\n"
            "4. Готового рису в Україні не виробляють: 100 % обсягу — імпорт.", ORANGE)
    insight(s, M + 6.12, 4.06, 5.86, 2.34, "Чого на ринку немає",
            "1. Немає чаші за 100–150 грн у мережі: Ottogi коштує 225–356 грн, Henan 83–156, "
            "але обидва лише онлайн.\n"
            "2. Немає українського виробника готового рису в паучі для мікрохвильовки — "
            "українські 16 позицій це або консерва, або сублімат для туризму.\n"
            "3. Немає жодного бренду одночасно в мережі та на маркетплейсі, крім Ben's.\n"
            "4. Немає локалізації: жодної україномовної упаковки серед 84 позицій.", TEAL)
    foot(s, "Джерела: Eurostat Comext (CN 1904 90 10, експорт ЄС→Україна), UN Comtrade, "
            "17 мереж zakaz.ua, «Сільпо», Prom, Rozetka, MAUDAU — 22–23.09.2026. "
            "Ємність роздрібу — оцінка, метод на слайді 03.")


# ══ ОБСЯГ РИНКУ ═══════════════════════════════════════════════════════════
def slide_volume():
    s = slide()
    header(s, "ОБСЯГ РИНКУ", f"{num(T_TOTAL)} тонн імпорту, {num(RETAIL_LO)}–{num(RETAIL_HI)} млн грн роздрібу",
           "Готовий рис в Україні не виробляють у паучах — увесь обсяг категорії це імпорт.")
    # графік імпорту
    rect(s, M, 1.88, 7.30, 3.34, MIST, rounded=True, adj=0.05)
    text(s, M + 0.28, 2.06, 6.00, 0.24, "ІМПОРТ ГОТОВОГО РИСУ З ЄС, ТОНН НА РІК",
         size=9, bold=True, color=GREY)
    bx, by0, bw, bh = M + 0.34, 4.56, 0.86, 2.02
    mx = max(t for _, t, _ in IMPORT)
    for i, (yr, t, v) in enumerate(IMPORT):
        x = bx + i * 1.00
        h = bh * t / mx
        c = ORANGE if yr == 2025 else (NAVY_L if yr >= 2022 else RGBColor(0xC2, 0xC9, 0xDA))
        rect(s, x, by0 - h, bw, h, c, rounded=True, adj=0.10)
        text(s, x - 0.06, by0 - h - 0.24, 0.98, 0.22, num(t), size=8.5, bold=True,
             color=c if yr == 2025 else GREY, align=PP_ALIGN.CENTER)
        text(s, x - 0.06, by0 + 0.06, 0.98, 0.22, str(yr), size=8.5,
             bold=(yr == 2025), color=INK, align=PP_ALIGN.CENTER)
        text(s, x - 0.06, by0 + 0.26, 0.98, 0.20, f"€{num(v)}k", size=7.5, color=GREY,
             align=PP_ALIGN.CENTER)
    text(s, M + 0.34, 4.98, 7.00, 0.20,
         "2025 рік — рекорд за вартістю: 823 тис. € проти 475 тис. € у 2024-му.",
         size=8.5, color=GREY)
    # звідки їдe
    rect(s, M + 7.54, 1.88, 4.44, 3.34, WHITE, line=MIST_D, lw=1.2, rounded=True, adj=0.05)
    text(s, M + 7.80, 2.06, 3.92, 0.24, "ЗВІДКИ ЇДЕ, 2025 РІК", size=9, bold=True, color=GREY)
    yy = 2.30
    tot = sum(t for _, t, _ in IMPORT_PARTNERS_2025) + T_ASIA
    for nm, t, v in IMPORT_PARTNERS_2025 + [(n, t, 0) for n, t, _ in ASIA_2025 if t > 1]:
        w = 3.10 * t / tot
        text(s, M + 7.80, yy, 1.00, 0.22, nm, size=9, color=INK)
        rect(s, M + 8.84, yy + 0.05, max(w, 0.04), 0.13, ORANGE if t > 50 else TEAL,
             rounded=True, adj=0.5)
        text(s, M + 8.84 + max(w, 0.04) + 0.08, yy - 0.01, 0.90, 0.22, f"{num(t, 1)} т",
             size=8.5, bold=True, color=GREY)
        yy += 0.30
    text(s, M + 7.80, yy + 0.04, 3.92, 0.64,
         "Болгарія і Польща — 95 % тонажу. Це заводи ЄС, що пакують рис у пауч.\n"
         "Корея і Китай разом дають 3,1 т — це чаші Ottogi, Bibigo та Henan.",
         size=8.5, color=GREY, line=1.26)
    # розрахунок ємності
    rect(s, M, 5.26, 11.98, 1.44, NAVY, rounded=True, adj=0.07)
    rect(s, M, 5.26, 0.09, 1.44, GREEN, rounded=True, adj=0.5)
    text(s, M + 0.34, 5.42, 4.00, 0.26, "Як рахували ємність", size=12.5, bold=True, color=AMBER)
    STEPS = [(f"{num(T_TOTAL)} т", "імпорт 2025"), (f"{num(CIF_UAH)} млн грн", "CIF за курсом НБУ"),
             (f"× 2,2–2,8", "мито 0 %, ПДВ, маржа"),
             (f"{num(RETAIL_LO)}–{num(RETAIL_HI)} млн грн", "роздріб на рік"),
             (f"≈ {num(PACKS_K)} тис.", "упаковок на рік")]
    for i, (v, lb) in enumerate(STEPS):
        x = M + 4.30 + i * 1.54
        text(s, x, 5.42, 1.46, 0.28, v, size=12, bold=True, color=WHITE)
        text(s, x, 5.70, 1.46, 0.36, lb, size=8, color=RGBColor(0x9F, 0xA9, 0xC4), line=1.20)
        if i < len(STEPS) - 1:
            text(s, x + 1.30, 5.44, 0.20, 0.24, "→", size=12, bold=True, color=ORANGE)
    text(s, M + 0.34, 5.80, 3.80, 0.76,
         f"Перевірка з іншого боку: {num(PACKS_K)} тис. упаковок × 110 грн медіани "
         f"= {num(PACKS_K * 110 / 1000)} млн грн. Два методи сходяться.",
         size=8.5, color=RGBColor(0xB9, 0xC2, 0xDA), line=1.26)
    foot(s, "Джерела: Eurostat Comext DS-045409, CN 1904 90 10 «рис приготовлений», експорт "
            "ЄС→Україна, 2019–2025; UN Comtrade (Корея, Китай, Таїланд, 2025); курс НБУ "
            "(середній 2025: 47,15 грн/€, 41,71 грн/$). Ємність роздрібу — ОЦІНКА, не заміряні продажі.")


# ══ ФОРМАТИ УПАКОВКИ ══════════════════════════════════════════════════════
def slide_formats():
    s = slide()
    header(s, "ФОРМАТИ УПАКОВКИ", "Дойпак і чаша — 63 % позицій, але полицю тримає пауч",
           "Скільки позицій у кожному форматі, скільки коштує одна упаковка і скільки в ній грамів.")
    PICS = {"pouch": "md/bens_basmati250.jpg", "cup": "sku/hn_scallop174.jpg",
            "box": "sku/hd_beef_stew272.jpg", "doypack": "sku/tl_strog125.jpg"}
    NOTE = {"pouch": "Плаский реторт-пауч. Мікрохвильовка 90 с або окріп. Формат мережевої полиці.",
            "cup": "Жорстка чаша чи стакан. Готова страва або сухий рис під окріп. Формат для офісу.",
            "box": "Картон із хімічним нагрівачем. Без техніки. Дорогий і об'ємний.",
            "doypack": "Дойпак із сублімату. Легкий, але потребує окропу. Туризм і армія."}
    COL = {"pouch": ORANGE, "cup": ROSE, "box": PLUM, "doypack": TEAL}
    for i, k in enumerate(PACK_ORDER):
        it = BY_PACK[k]
        x = M + i * 3.07
        c = COL[k]
        pr = sorted(x2[4] for x2 in it)
        ms = sorted(x2[3] for x2 in it if x2[3])
        rect(s, x, 1.86, 2.82, 3.94, WHITE, line=MIST_D, lw=1.2, rounded=True, adj=0.05)
        rect(s, x, 1.86, 2.82, 0.10, c, rounded=True, adj=0.5)
        rect(s, x + 0.16, 2.06, 2.50, 1.10, MIST, rounded=True, adj=0.06)
        pic(s, PICS[k], x + 0.26, 2.12, 2.30, 0.98)
        text(s, x + 0.24, 3.26, 2.36, 0.24, PACK_RU[k], size=10.5, bold=True, color=c)
        text(s, x + 0.24, 3.52, 2.36, 0.44, NOTE[k], size=8, color=GREY, line=1.22)
        rect(s, x + 0.24, 4.04, 2.34, 0.18, MIST_D, rounded=True, adj=0.5)
        rect(s, x + 0.24, 4.04, 2.34 * len(it) / N_SKU, 0.18, c, rounded=True, adj=0.5)
        text(s, x + 0.24, 4.28, 2.36, 0.24,
             f"{pl(len(it), 'позиція', 'позиції', 'позицій')} · {num(len(it) / N_SKU * 100)} % · "
             f"{pl(len(set(x2[1] for x2 in it)), 'бренд', 'бренди', 'брендів')}",
             size=10.5, bold=True, color=NAVY)
        rect(s, x + 0.24, 4.60, 2.34, 0.02, MIST_D)
        for j, (lb, v) in enumerate([("ЦІНА ЗА 1 ШТ, МЕДІАНА", f"{num(st.median(pr))} грн"),
                                     ("ДІАПАЗОН ЦІНИ", f"{num(min(pr))} – {num(max(pr))} грн"),
                                     ("МАСА, МЕДІАНА", f"{num(st.median(ms))} г")]):
            yy = 4.70 + j * 0.38
            text(s, x + 0.24, yy, 1.80, 0.20, lb, size=7, bold=True, color=GREY)
            text(s, x + 0.24, yy + 0.16, 2.36, 0.22, v, size=10.5, bold=True,
                 color=c if j < 2 else INK)
    insight(s, M, 5.96, 11.98, 0.94, "",
            "Дойпак виграє за кількістю позицій лише тому, що туристичні магазини заводять "
            "десятки дрібних брендів. За доступністю для покупця виграє пауч: 109 грн медіана "
            "проти 315 грн у дойпака й 708 грн у коробки з нагрівачем.", ORANGE)
    foot(s, f"Розподіл {N_SKU} позицій за типом паковання з карток продавців, 22–23.09.2026.")


# ══ ГРАМАЖ ════════════════════════════════════════════════════════════════
def slide_grammage():
    s = slide()
    header(s, "ГРАМАЖ", "250 і 220 грамів — половина всієї полиці паучів",
           "Скільки позицій припадає на кожну вагу і як ваги розподілені між форматами.")
    masses = sorted([x[3] for x in ALL8 if x[3]])
    cnt = collections.Counter(masses)
    tops = sorted(cnt.items(), key=lambda kv: -kv[1])[:8]
    tops = sorted(tops, key=lambda kv: kv[0])
    text(s, M, 1.86, 6.00, 0.24, "НАЙЧАСТІШІ ВАГИ, ПОЗИЦІЙ", size=9, bold=True, color=GREY)
    bx, by0, bw = M, 4.34, 0.94
    mxn = max(n for _, n in tops)
    for i, (g_, n) in enumerate(tops):
        x = bx + i * 1.06
        h = 2.10 * n / mxn
        pk = collections.Counter(pack_of(y[1], y[2], y[3]) for y in ALL8 if y[3] == g_)
        c = {"pouch": ORANGE, "cup": ROSE, "box": PLUM, "doypack": TEAL}[pk.most_common(1)[0][0]]
        rect(s, x, by0 - h, bw, h, c, rounded=True, adj=0.12)
        text(s, x, by0 - h - 0.24, bw, 0.22, str(n), size=11, bold=True, color=c,
             align=PP_ALIGN.CENTER)
        text(s, x, by0 + 0.06, bw, 0.22, f"{num(g_)} г", size=8.5, bold=True, color=INK,
             align=PP_ALIGN.CENTER)
    text(s, M, 4.66, 8.50, 0.44,
         "Колір стовпця — формат, у якому ця вага трапляється найчастіше: "
         "помаранчевий — пауч, рожевий — чаша, бірюзовий — дойпак.",
         size=8.5, color=GREY, line=1.24)
    # ваги по формату
    rect(s, 9.10, 1.86, 3.56, 3.46, MIST, rounded=True, adj=0.05)
    text(s, 9.34, 2.04, 3.10, 0.24, "ВАГОВЕ ВІКНО ФОРМАТУ", size=9, bold=True, color=GREY)
    yy = 2.38
    for k in PACK_ORDER:
        ms = sorted(x[3] for x in BY_PACK[k] if x[3])
        c = {"pouch": ORANGE, "cup": ROSE, "box": PLUM, "doypack": TEAL}[k]
        text(s, 9.34, yy, 2.00, 0.22,
             {"pouch": "Пауч", "cup": "Чаша", "box": "Коробка з нагрівачем",
              "doypack": "Дойпак"}[k], size=9, bold=True, color=c)
        text(s, 9.34, yy + 0.20, 3.10, 0.22,
             f"{num(min(ms))}–{num(max(ms))} г · медіана {num(st.median(ms))} г",
             size=9, color=INK)
        yy += 0.62
    text(s, 9.34, yy + 0.02, 3.10, 0.40,
         "Чаша й пауч конкурують у вікні 210–250 г — це порція на одну людину.",
         size=8, color=GREY, line=1.22)
    insight(s, M, 5.48, 11.98, 1.44, "Що обирати під полицю",
            "Вагове ядро категорії — 220–250 г: 16 позицій із 84, усі пауч або чаша, "
            "усі в корені порції «обід для однієї людини». Це та вага, під яку вже налаштований "
            "покупець Ben's Original і Bibigo.\n"
            "Легкі ваги 80–146 г — це не менша порція, а сухий продукт: сублімат і інстант-рис, "
            "які треба залити окропом. Порівнювати їх із паучем за грам не можна.\n"
            "Важкі 310–440 г — корейські чаші Ottogi та китайські коробки з нагрівачем: "
            "порція «на голодного», але й ціна за упаковку 250–735 грн.", ORANGE)
    foot(s, f"Маса з карток продавців; {sum(1 for x in ALL8 if not x[3])} позицій без указаної "
            "маси до розрахунку не входять.")


# ══ ЦІНОВІ СХОДИ ЗА 1 ШТ ══════════════════════════════════════════════════
def slide_ladder():
    s = slide()
    header(s, "ЦІНА ЗА ОДНУ УПАКОВКУ", "Від 45 до 1 190 грн: сходи всіх 23 брендів",
           "Крапка — медіана бренду, смуга — від найдешевшої до найдорожчої його позиції.")
    rows = []
    for b in dict.fromkeys(x[1] for x in ALL8):
        it = [x for x in ALL8 if x[1] == b]
        rows.append((b, min(x[4] for x in it), max(x[5] for x in it),
                     st.median([x[4] for x in it]), len(it), COUNTRY[b]))
    rows.sort(key=lambda r: r[3])
    LO, HI = 0, 1250
    AX, AW = 4.16, 7.62

    def ax(v):
        return AX + AW * (v - LO) / (HI - LO)
    rect(s, M, 1.82, 11.98, 5.04, MIST, rounded=True, adj=0.04)
    for t in range(250, 1251, 250):
        rect(s, ax(t), 2.00, 0.012, 4.42, RGBColor(0xDC, 0xE2, 0xEE))
        text(s, ax(t) - 0.40, 6.46, 0.80, 0.22, num(t), size=8, color=GREY,
             align=PP_ALIGN.CENTER)
    text(s, M + 0.26, 6.46, 3.20, 0.22, "ГРН ЗА ОДНУ УПАКОВКУ", size=8, bold=True, color=GREY)
    text(s, M + 0.26, 2.00, 2.20, 0.18, "БРЕНД", size=6.5, bold=True, color=GREY)
    text(s, M + 2.42, 2.00, 0.70, 0.18, "КРАЇНА", size=6.5, bold=True, color=GREY)
    text(s, M + 3.16, 2.00, 0.36, 0.18, "SKU", size=6.5, bold=True, color=GREY,
         align=PP_ALIGN.RIGHT)
    yy = 2.26
    for b_, lo, hi, med, n, cty in rows:
        c = {"Україна": GREEN, "Китай": PLUM, "Корея": ROSE, "США": SLATE}.get(cty, ORANGE)
        text(s, M + 0.26, yy - 0.09, 2.14, 0.20, b_, size=8.5, bold=True, color=NAVY)
        text(s, M + 2.42, yy - 0.07, 0.70, 0.18, cty, size=7, color=GREY)
        text(s, M + 3.16, yy - 0.07, 0.36, 0.18, str(n), size=7, bold=True, color=GREY,
             align=PP_ALIGN.RIGHT)
        x0, x1 = ax(lo), ax(hi)
        rect(s, x0, yy - 0.035, max(x1 - x0, 0.02), 0.07, c, rounded=True, adj=0.5)
        dot(s, ax(med), yy, 0.15, c)
        lbl = num(med) if lo == hi else f"{num(lo)} – {num(hi)}"
        text(s, max(x1, ax(med)) + 0.10, yy - 0.10, 1.30, 0.20, lbl, size=7.5, bold=True,
             color=c)
        yy += 0.180
    rect(s, ax(45), 6.70, ax(180) - ax(45), 0.05, ORANGE)
    text(s, ax(45), 6.76, 3.60, 0.22, "вікно мережевої полиці: 45–180 грн", size=8.5,
         bold=True, color=ORANGE)
    text(s, M + 0.26, 6.76, 3.60, 0.22,
         "Колір — країна: Україна · Китай · Корея · ЄС · США", size=8, color=GREY)
    foot(s, "Ціна за одну упаковку в роздріб, без оптових порогів: у 6 продавців діє нижча ціна "
            "від 3–20 шт, вона в розрахунок не бралася. 22–23.09.2026.")


# ══ ПОЛИЦЯ МЕРЕЖ ══════════════════════════════════════════════════════════
def slide_chain_shelf():
    s = slide()
    header(s, "ЩО СТОЇТЬ У МЕРЕЖІ ЗАМІСТЬ НАС", "Рис із м'ясом є на полиці — але в бляшанці",
           "Мережі не мають пауча з готовим рисом, зате мають суміжну категорію за 63–160 грн.")
    for i, (im, b, nm, g, lo, hi, ch, c) in enumerate(CHAIN_SHELF):
        sku_cell(s, M + i * 2.02, 1.86, 1.92, 2.90, im, nm, g, lo, hi, ch, c, brand=b)
    insight(s, M, 4.96, 5.86, 1.92, "Чому це важливо",
            "Це прямий конкурент за той самий привід: гаряча страва з рисом і м'ясом без "
            "готування. Він уже стоїть на полиці, коштує 63–160 грн за 250–350 г і має "
            "українського виробника.\n"
            "Пауч із готовим рисом заходить у той самий гаманець, тож орієнтир ціни — "
            "не Ottogi за 255 грн, а ця полиця.", ROSE)
    rect(s, M + 6.12, 4.96, 5.86, 1.92, MIST, rounded=True, adj=0.06)
    text(s, M + 6.40, 5.14, 5.30, 0.26, "Порівняння за 100 г", size=12, bold=True, color=NAVY)
    CMP = [("Каша рисова hapay! 340 г", 18.5, GREEN),
           ("Плов Food Fabrika 250 г", 47.6, PLUM),
           ("Ben's Original, медіана", 41.0, ORANGE),
           ("Плов М'ясторія 350 г", 45.7, ROSE),
           ("Ottogi, медіана", 88.0, SLATE)]
    yy = 5.50
    mxv = max(v for _, v, _ in CMP)
    for lb, v, c in CMP:
        text(s, M + 6.40, yy - 0.04, 2.30, 0.20, lb, size=8.5, color=INK)
        rect(s, M + 8.80, yy + 0.01, 2.40 * v / mxv, 0.13, c, rounded=True, adj=0.5)
        text(s, M + 8.80 + 2.40 * v / mxv + 0.08, yy - 0.05, 0.80, 0.20, f"{num(v, 1)} грн",
             size=8, bold=True, color=c)
        yy += 0.26
    foot(s, "Джерело: API 17 мереж zakaz.ua, 23.09.2026. Ці 6 позицій — суміжна категорія, "
            f"у {N_SKU} позицій дослідження вони не входять. Охолоджені страви й кулінарію мереж "
            "не показано: у них термін 2–5 діб.")


# ══ ВИСНОВКИ ══════════════════════════════════════════════════════════════
def slide_conclusions():
    s = slide()
    header(s, "ВИСНОВКИ", "Що показало дослідження",
           "Чотири факти про ринок і три питання, на які дані відповіді не дають.")
    FACTS = [("Ринок є, полиці немає", ORANGE,
              f"{num(T_TOTAL)} т імпорту і {N_SKU} позицій у продажу — але {N_SKU - N_SILPO} із них "
              "живуть лише онлайн. Мережеву дистрибуцію має один бренд."),
             ("Вікно ціни — 45–180 грн", GREEN,
              "Стільки коштує упаковка, яку покупець бере в мережі. Усе, що дорожче 250 грн, "
              "продається поштучно в нішевих магазинах."),
             ("Формат — пауч 220–250 г", ROSE,
              "16 позицій у цьому вікні, медіана 109 грн. Чаша дорожча вдвічі, коробка з "
              "нагрівачем — усемеро."),
             ("Українського пауча немає", PLUM,
              f"{N_UA} українських позицій — це консерви й сублімати для туризму. "
              "Готового рису в паучі під мікрохвильовку не робить ніхто.")]
    for i, (t, c, body) in enumerate(FACTS):
        x = M + (i % 2) * 6.12
        y = 1.86 + (i // 2) * 1.38
        rect(s, x, y, 5.86, 1.28, MIST, rounded=True, adj=0.08)
        rect(s, x, y, 0.09, 1.28, c, rounded=True, adj=0.5)
        text(s, x + 0.30, y + 0.16, 5.30, 0.26, f"{i + 1}. {t}", size=12.5, bold=True, color=NAVY)
        text(s, x + 0.30, y + 0.48, 5.30, 0.70, body, size=9.5, color=INK, line=1.28)
    insight(s, M, 4.74, 11.98, 1.74, "Чого дані не показують",
            "1. Продажів. Ні мережі, ні маркетплейси не віддають обсяги. Ємність роздрібу на "
            "слайді 03 — оцінка за імпортом, а не заміряний продаж.\n"
            "2. Хто імпортер. У картках продавців імпортера не вказано; Prom показує лише "
            "продавця. Точний перелік дасть лише вивантаження Держмитслужби за УКТЗЕД 1904 90 10.\n"
            "3. Оборотності на полиці «Сільпо». Скільки Ben's Original реально продає за "
            "місяць — без даних мережі не порахувати.\n"
            "Наступний крок, який закриє 1 і 2: замовити митну базу за 1904 90 10 за 2024–2025 рр. "
            "і запитати в «Сільпо» дані про категорію.", ORANGE)
    foot(s, "Дослідження: 84 позиції, 23 бренди, 56 продавців, 17 мереж. Зріз 22–23 вересня 2026 р.")

import math

# ══ ОБКЛАДИНКА ════════════════════════════════════════════════════════════
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


def slide_cover():
    s = slide(dark=True)
    cover_pic(s, BG_TITLE, 0, 0, W, H)
    s.shapes.add_picture(LOGO, Inches(M), Inches(0.52), Inches(1.94), Inches(0.84))
    rect(s, M, 2.02, 0.54, 0.09, ORANGE)
    text(s, M, 2.34, 7.00, 0.34, "ДОСЛІДЖЕННЯ РИНКУ · ВЕРЕСЕНЬ 2026", size=11.5,
         bold=True, color=AMBER)
    text(s, M, 2.78, 7.20, 1.70, "Готовий рис\nв Україні", size=46, bold=True,
         color=WHITE, line=1.04)
    text(s, M, 4.52, 6.80, 0.36, "Обсяг ринку · формати · грамаж · ціна · канали",
         size=14.5, color=RGBColor(0xC6, 0xCE, 0xE2))
    rect(s, M, 5.12, 5.90, 0.055, RGBColor(0x55, 0x60, 0x8C))
    text(s, M, 5.34, 6.90, 0.80,
         f"{pl(N_BRANDS, 'бренд', 'бренди', 'брендів')} · "
         f"{pl(N_SKU, 'позиція', 'позиції', 'позицій')} з пакшотом і ціною за 1 шт · "
         f"{pl(N_CH, 'продавець', 'продавці', 'продавців')}\n"
         "Зріз українських онлайн-вітрин і 17 мереж: 22–23 вересня 2026 р.",
         size=11.5, color=RGBColor(0x9F, 0xA9, 0xC4), line=1.40)
    for im, x, y, w, h in [("md/bens_basmati250.jpg", 8.26, 0.96, 2.35, 2.40),
                           ("sku/hn_scallop174.jpg", 10.78, 0.96, 2.35, 2.40),
                           ("sku/hd_beef_stew272.jpg", 8.26, 3.50, 2.35, 2.10),
                           ("sku/ot_bibimbap.jpg", 10.78, 3.50, 2.35, 2.10)]:
        rect(s, x, y, w, h, WHITE, rounded=True, adj=0.06)
        pic(s, im, x + 0.16, y + 0.14, w - 0.32, h - 0.28)
    rect(s, 8.26, 5.76, 4.87, 1.30, RGBColor(0x26, 0x2F, 0x4D), rounded=True, adj=0.12)
    text(s, 8.50, 5.90, 4.44, 1.08, " · ".join(dict.fromkeys(x[1] for x in ALL8)),
         size=9, bold=True, color=WHITE, line=1.30)


# ══ КАРТА РИНКУ ═══════════════════════════════════════════════════════════
def slide_map():
    s = slide()
    header(s, "КАРТА РИНКУ", f"{pl(N_BRANDS, 'бренд', 'бренди', 'брендів')}, чотири технології",
           "Групування за тим, як продукт готується: від цього залежить і ціна, і канал.")
    LEAGUES = [("ГОТОВИЙ РИС · РОЗІГРІВ", L_READY, ORANGE,
                "Мікрохвильовка 60–90 с або реторт. Маса — готової страви."),
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
        rect(s, x, 1.90, 2.82, 2.94, MIST, rounded=True, adj=0.06)
        rect(s, x, 1.90, 2.82, 0.10, c, rounded=True, adj=0.5)
        text(s, x + 0.24, 2.10, 2.36, 0.24, kind, size=10, bold=True, color=c)
        text(s, x + 0.24, 2.38, 2.36, 0.40,
             f"{pl(len(items), 'позиція', 'позиції', 'позицій')} · "
             f"{pl(len(br), 'бренд', 'бренди', 'брендів')}", size=13.5, bold=True, color=NAVY)
        text(s, x + 0.24, 2.78, 2.36, 0.74, " · ".join(br), size=8.5, color=INK, line=1.20)
        text(s, x + 0.24, 3.56, 2.36, 0.44, note, size=8.5, color=GREY, line=1.20)
        rect(s, x + 0.24, 4.06, 2.34, 0.02, MIST_D)
        text(s, x + 0.24, 4.16, 2.36, 0.20, "ЗА 1 УПАКОВКУ · МЕДІАНА", size=7,
             bold=True, color=GREY)
        text(s, x + 0.24, 4.34, 2.36, 0.28, f"{num(unit_med(items))} грн", size=14,
             bold=True, color=c)
        text(s, x + 0.24, 4.60, 2.36, 0.20, f"діапазон {num(lo)} – {num(hi)} грн",
             size=8, color=GREY)
    rect(s, M, 5.00, 11.98, 1.88, NAVY, rounded=True, adj=0.07)
    rect(s, M, 5.00, 0.10, 1.88, ORANGE, rounded=True, adj=0.5)
    text(s, M + 0.42, 5.18, 5.40, 0.30, "Що показує розбір", size=14, bold=True, color=AMBER)
    yy = 5.56
    for t in [f"Ben's Original — {len(BENS)} позицій, єдиний бренд у мережі («Сільпо»).",
              f"Ottogi і Henan — {len(OTTOGI) + len(HENAN)} азійських чаш; Henan удвічі "
              "дешевший за упаковку.",
              f"Українських позицій — {N_UA}, але жодної в паучі для мікрохвильовки.",
              "Саморозігрів — найдорожча технологія: медіана 708 грн за упаковку."]:
        dot(s, M + 0.48, yy + 0.11, 0.10, AMBER)
        text(s, M + 0.68, yy, 5.62, 0.26, t, size=10, color=WHITE)
        yy += 0.30
    rect(s, 6.90, 5.24, 0.04, 1.40, NAVY_L)
    text(s, 7.24, 5.18, 5.10, 1.52,
         f"Медіана готового рису для розігріву ({N_CORE} позицій) — {num(CORE_MED)} грн за 100 г.\n"
         f"У мережі продається {N_SILPO} позицій із {N_SKU}. Решта {N_SKU - N_SILPO} — лише "
         "онлайн: маркетплейси, азійські фудшопи, туристичні й військові магазини.\n"
         "Країни: ЄС — 31 позиція, Китай — 24, Україна — 16, Корея — 11, США — 2.",
         size=10, color=RGBColor(0xB9, 0xC2, 0xDA), line=1.30)
    foot(s, "Джерела: 17 мереж на zakaz.ua, «Сільпо», Prom, MAUDAU, Rozetka, спеціалізовані "
            "магазини — 22–23.09.2026.")


def brand_slide(n, brand, origin, fmt, items, dek, src, cols, plo, phi):
    lo100, hi100 = p100(as8(items, "", ORANGE))
    s = slide()
    header(s, f"БРЕНД {n} · {brand.upper()}",
           f"{pl(len(items), 'позиція', 'позиції', 'позицій')}: {fmt}", dek)
    brand_head(s, M, 1.72, brand, origin, fmt.split(" ")[0], len(items),
               lo100, hi100, plo, phi, cols)
    half = (len(items) + 1) // 2
    w = (11.98 - 0.10 * (half - 1)) / half
    for i, (img, nm, g, lo, hi, chan) in enumerate(items[:half]):
        sku_cell(s, M + i * (w + 0.10), 2.50, w, 2.22, img, nm, g, lo, hi, chan, cols)
    for i, (img, nm, g, lo, hi, chan) in enumerate(items[half:]):
        sku_cell(s, M + i * (w + 0.10), 4.78, w, 2.22, img, nm, g, lo, hi, chan, cols)
    foot(s, src)


def multi_slide(n_from, n_to, title, dek, items, note_title, note, note_c, src, two_rows=False):
    """Сітка позицій; у два ряди врізка займає вільні клітинки другого ряду."""
    s = slide()
    header(s, f"БРЕНДИ {n_from}–{n_to}", title, dek)
    n = len(items)
    if not two_rows:
        w = (11.98 - 0.10 * (n - 1)) / n
        grid(s, items, 1.88, w, 3.44)
        insight(s, M, 5.48, 11.98, 1.42, note_title, note, note_c)
    else:
        cols = math.ceil((n + 2) / 2)          # ≥2 вільні клітинки під врізку
        w = (11.98 - 0.10 * (cols - 1)) / cols
        grid(s, items[:cols], 1.88, w, 2.34)
        rest = items[cols:]
        grid(s, rest, 4.32, w, 2.34)
        x0 = M + len(rest) * (w + 0.10)
        insight(s, x0, 4.32, M + 11.98 - x0, 2.34, note_title, note, note_c)
    foot(s, src)

# ══ КАНАЛИ ════════════════════════════════════════════════════════════════
def slide_channels():
    s = slide()
    header(s, "ДЕ ПРЕДСТАВЛЕНО", f"{pl(N_CH, 'продавець', 'продавці', 'продавців')}, чотири типи каналів",
           f"Категорія живе в онлайні: у мережі — лише Ben's Original у «Сільпо», {N_SILPO} позицій.")
    GROUPS = [("МЕРЕЖЕВИЙ РОЗДРІБ", ORANGE, CH_CHAIN, "Ben's Original"),
              ("ОНЛАЙН-СУПЕРМАРКЕТИ Й МАРКЕТПЛЕЙСИ", GREEN, CH_MARKET,
               "Ben's · Henan · Bibigo · Clearspring · Portion"),
              ("СПЕЦІАЛІЗОВАНІ ФУДШОПИ", PLUM, CH_FOOD,
               "Ottogi · Henan · Bibigo · саморозігрів · Gallina Blanca"),
              ("ТУРИСТИЧНІ ТА ВІЙСЬКОВІ", TEAL, CH_OUT,
               "Haidilao · Adventure Menu · усі сублімати · Маркел")]
    yy = 1.88
    for gi, (title, c, chans, what) in enumerate(GROUPS):
        rh = 0.76 if gi < 2 else 1.10
        rect(s, M, yy, 11.98, rh, MIST if gi % 2 == 0 else WHITE, rounded=True, adj=0.12)
        rect(s, M, yy, 0.09, rh, c, rounded=True, adj=0.5)
        text(s, M + 0.30, yy + 0.12, 3.60, 0.22, title, size=9, bold=True, color=c)
        text(s, M + 0.30, yy + 0.38, 3.60, 0.56, what, size=8.5, color=GREY, line=1.20)
        shown = chans if len(chans) <= 12 else chans[:10]
        cx, cy = M + 4.00, yy + 0.14
        for ch in shown + ([f"+ ще {len(chans) - 10}"] if len(chans) > 12 else []):
            w = 0.22 + 0.064 * len(ch)
            if cx + w > M + 11.00:
                cx, cy = M + 4.00, cy + 0.32
            pill(s, cx, cy, ch, c, w=w, size=8)
            cx += w + 0.10
        text(s, M + 11.20, yy + rh / 2 - 0.17, 0.62, 0.34, str(len(chans)), size=16, bold=True,
             color=c, align=PP_ALIGN.RIGHT)
        yy += rh + 0.08
    rect(s, M, 5.86, 5.86, 0.98, MIST, rounded=True, adj=0.10)
    text(s, M + 0.26, 6.00, 5.34, 0.24, "Що відомо про імпортерів", size=11, bold=True, color=NAVY)
    text(s, M + 0.26, 6.26, 5.34, 0.52,
         "Картки продавців імпортера не називають. OMG! Asia і СНЕКІС описують себе як "
         "прямих імпортерів з Азії; Ben's Original виробляє Mars (заводи Франція, Німеччина).",
         size=8.5, color=INK, line=1.24)
    rect(s, M + 6.12, 5.86, 5.86, 0.98, MIST, rounded=True, adj=0.10)
    text(s, M + 6.38, 6.00, 5.34, 0.24, "Мережі без жодної позиції категорії", size=11,
         bold=True, color=NAVY)
    text(s, M + 6.38, 6.26, 5.34, 0.52,
         "АТБ, Novus, Metro, Varus, Ашан, Fozzy, «Таврія В», МегаМаркет, ЕКО маркет, "
         "Ultramarket та ще 7 мереж. Суміжну категорію (слайд 17) частина з них має.",
         size=8.5, color=INK, line=1.24)


# ══ ЗБІРКА ════════════════════════════════════════════════════════════════
slide_cover()
slide_summary()
slide_volume()
slide_map()
slide_formats()
slide_grammage()
slide_ladder()

brand_slide(1, "Ben's Original", "Mars · ЄС", "пауч 220, 240 і 250 г", BENS,
            "Найглибша лінійка категорії і єдиний бренд, що дійшов до мережевої полиці.",
            "Джерела: каталог «Сільпо» (11 позицій), MAUDAU, Edison Lee — 22.09.2026. "
            "Basmati 220 г: 100 грн у MAUDAU і 144 грн у «Сільпо». Дві позиції Edison Lee "
            "продаються поштучно, ціна за один пауч.", ORANGE, 45, 179)
brand_slide(2, "Ottogi", "Республіка Корея", "чаша 217–320 г", OTTOGI,
            "Готовий рис із наповнювачем для мікрохвильовки. Продається у чотирьох продавців.",
            "Джерела: Pulsar, Апетітаріум, Gurmissimo, Тайякі Март — 22.09.2026. Ціни "
            "роздрібні: у «Тайякі Март» від 15–20 шт діє нижча ціна, вона не бралася.",
            PLUM, 225, 356)
brand_slide(3, "Henan", "Китай · Xiao Guo Zao", "чаша 144 і 174 г", HENAN,
            "Сухий рис у чаші, заливається окропом на 8 хвилин. Найдешевша чаша на ринку.",
            "Джерела: MAUDAU, Апетітаріум, Gurmissimo, OMG! Asia, СНЕКІС, DCM — 22–23.09.2026. "
            "Маса — сухого продукту до заливання.", ROSE, 83, 156)
multi_slide(4, 8, "Bibigo, Clearspring, Portion, Маркел, Gallina Blanca",
            "По одній-дві позиції в кожного: різні країни, різні формати, різні канали.",
            READY_MORE, "Що це означає",
            "Portion і Маркел — український реторт 350 г за 94–109 грн: найдешевший грам "
            "категорії, 27–31 грн за 100 г. Bibigo — той самий товар у трьох каналах, розкид "
            "40 %. Clearspring — органіка, найдорожчий грам серед паучів. Gallina Blanca "
            "Yatekomo — 84 г сухого рису в чаші, заливається окропом.", ORANGE,
            "Джерела: Смак Кореї, Rozetka, Prom, MAUDAU, СУХПАЙ, UPcompany, «Товари з Іспанії» "
            "— 22–23.09.2026. Фото Portion: Rozetka не віддає зображення картки.")
brand_slide(9, "Haidilao", "Китай · саморозігрів", "коробка 165–360 г", HAIDILAO,
            "Хімічний нагрівач у коробці — готується без мікрохвильовки й окропу.",
            "Джерела: Daruy, Tactico, Desna, Sweet Svitt, «Шериф» — 22–23.09.2026. "
            "Найдорожча упаковка категорії — 1 190 грн за 360 г. * Sweet Svitt: назва картки "
            "Stewed Chicken, на фото — рис з яловичиною та грибами.", PLUM, 380, 1190)
multi_slide(10, 13, "Mo Xiao Xian, Zihaiguo, Rongcheng Haoji, Qiaoshanmei",
            "Китайський саморозігрів і рис під окріп. Усі чотири бренди — в одному магазині.",
            SELFHEAT_MORE, "Що це означає",
            "Zihaiguo і Rongcheng Haoji — коробка 440 г за 735 грн: 167 грн за 100 г, дешевше "
            "за грам, ніж Haidilao (230–371). Mo Xiao Xian — 931 грн за 275 г, найдорожча "
            "чаша ринку. Qiaoshanmei — не саморозігрів, а рис у пакеті під окріп, 501 грн.",
            PLUM, "Джерело: «Скарби Азії» (Prom), 23.09.2026. Бренди Zihaiguo і Rongcheng Haoji "
            "визначено за написами на упаковці, Qiaoshanmei — за карткою продавця.")
multi_slide(14, 17, "Travellunch, Trek'n Eat, Mountain House, Adventure Food",
            "Рис-страви у дойпаку: залити окропом на 8–10 хвилин. Німеччина, США, Нідерланди.",
            FD_IMPORT, "Що це означає",
            "Travellunch — 7 позицій у двох вагах, 125 і 250 г, найширша імпортна лінійка. "
            "Mountain House — 699 грн за 110–133 г сухої маси, найдорожчий грам категорії. "
            "Канал — туристичні й військові магазини, у продуктовий роздріб ці бренди не йдуть.",
            TEAL, "Джерела: ALANTUR, Highlander, ForCamp, Військторг Гайдамака, Freeride, "
            "Kamanti, 110вольт, Клуб Мандрівник — 22–23.09.2026.", two_rows=True)
multi_slide(18, 19, "Adventure Menu і SubliMate",
            "Adventure Menu — єдиний бренд із двома технологіями: готова страва 400 г і сублімат 110 г.",
            ADV_MENU + SUBLIMATE, "Що це означає",
            "Лінія Adventure Menu 400 г READY TO EAT — реторт без води, 79–101 грн за 100 г: "
            "у тому ж коридорі, що Ben's Original і Ottogi, але продається лише в туристичних "
            "магазинах. SubliMate — український сублімат, 290–370 грн за дойпак.",
            GREEN, "Джерела: ALANTUR, ForCamp, MK-Sport, Freeride, Лєєр, OXO, Palmer, "
            "Highlander, ВсеОпт, Суренж — 22–23.09.2026.", two_rows=True)
multi_slide(20, 23, "James Cook, Їжа в Похід, Харчі, !FEST",
            "Українські виробники рисових страв у дойпаку. Канал — туристичні й військові магазини.",
            FD_UA, "Що це означає",
            "Українські сублімати — 55–321 грн за упаковку, імпортні — 269–889 грн. "
            "James Cook — 4 рисові позиції, найширша українська лінійка. Українське "
            "виробництво в категорії є, але воно цілком у туристичному каналі.",
            ORANGE, "Джерела: Highlander, Activity, ВсеОпт, Klever-Shop, Terra Incognita, "
            "«Їжа в Похід», Харчі ТМ, Висот-Нік, SportStorm, Kalush-Craft — 22–23.09.2026.",
            two_rows=True)
slide_chain_shelf()
slide_channels()
slide_conclusions()

prs.save("Gotovyi_Rys_Rozbir_Brendiv.pptx")
print("saved ·", len(prs.slides._sldIdLst), "slides ·", N_SKU, "SKU ·", N_BRANDS, "brands ·",
      N_CH, "sellers · median", round(CORE_MED, 1), "| retail",
      round(RETAIL_LO), "-", round(RETAIL_HI), "млн грн")
