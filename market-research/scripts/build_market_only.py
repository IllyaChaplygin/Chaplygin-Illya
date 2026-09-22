"""Готовий рис · Україна — дослідження ринку.

Чотири питання: хто представлений, яка упаковка, яка ціна, де представлено.
Без наших позицій і без статусів наявності — тільки склад ринку.

Фірмовий стиль Morskyi Dim: темно-синій #303A5D, помаранчевий #F9A50B,
хвильовий патерн і логотип із оглядової колоди Nissin.
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
CORAL  = RGBColor(0xE2, 0x56, 0x4B)
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

# ── дані: (грамів, грн) по кожній перевіреній картці ──────────────────────
DATA = {
    "Ben's Original": [(250, 45), (250, 89.99), (250, 89.99), (250, 89.99),
                       (250, 89.99), (220, 89.99), (220, 100), (220, 109),
                       (240, 139), (220, 139), (220, 144), (220, 144),
                       (220, 149), (220, 179)],
    "Bibigo":      [(210, 135), (210, 184), (210, 189)],
    "Clearspring": [(250, 252)],
    "Portion":     [(350, 109)],
    "Ottogi":      [(269, 334), (310, 260), (320, 356), (247, 320), (269, 334),
                    (280, 289), (280, 260), (310, 252), (269, 252), (247, 225),
                    (320, 252), (280, 255)],
    "Haidilao":    [(187, 680), (272, 799), (272, 799), (272, 770), (175, 650),
                    (181, 650), (187, 650)],
    "Travellunch": [(125, 275), (125, 295), (125, 515), (125, 567), (250, 702),
                    (250, 811), (250, 889)],
}
CORE = ("Ben's Original", "Bibigo", "Clearspring", "Portion", "Ottogi")


def rng(brand):
    per = sorted(p / g * 100 for g, p in DATA[brand])
    pk = sorted(p for g, p in DATA[brand])
    return per[0], per[-1], pk[0], pk[-1], len(DATA[brand])


CORE_PER = sorted(p / g * 100 for b in CORE for g, p in DATA[b])
CORE_MED = st.median(CORE_PER)
N_CARDS = sum(len(v) for v in DATA.values())


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
    text(s, 2.52, 0.34, 7.20, 0.24, eyebrow, size=9.5, bold=True, color=AMBER)
    text(s, 2.52, 0.60, 8.60, 0.40, title, size=19, bold=True, color=WHITE)
    text(s, W - M - 0.70, 0.44, 0.70, 0.36, f"{PG['n']:02d}", size=17, bold=True,
         color=NAVY_L, align=PP_ALIGN.RIGHT)
    if dek:
        text(s, M, HDR + 0.28, 11.9, 0.34, dek, size=12.5, color=GREY, line=1.20)


def foot(s, src):
    text(s, M, 7.06, 11.9, 0.26, src, size=8, color=GREY, line=1.14)


def pill(s, x, y, label, bg=GREEN, w=None, size=8.5, fg=WHITE):
    w = w or (0.24 + 0.072 * len(label))
    rect(s, x, y, w, 0.26, bg, rounded=True, adj=0.5)
    text(s, x, y, w, 0.26, label, size=size, bold=True, color=fg,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    return w


def kpi(s, x, y, w, value, unit, label, sub, accent=ORANGE):
    rect(s, x, y, w, 2.42, MIST, rounded=True, adj=0.05)
    rect(s, x, y, w, 0.10, accent, rounded=True, adj=0.5)
    text(s, x + 0.28, y + 0.32, w - 0.56, 0.72, value, size=38, bold=True, color=NAVY)
    text(s, x + 0.28, y + 1.06, w - 0.56, 0.24, unit, size=10, bold=True, color=accent)
    text(s, x + 0.28, y + 1.38, w - 0.56, 0.52, label, size=11.5, bold=True, color=INK,
         line=1.14)
    text(s, x + 0.28, y + 1.96, w - 0.56, 0.38, sub, size=9, color=GREY, line=1.16)


def qtag(s, x, y, letter, word):
    rect(s, x, y, 0.30, 0.30, ORANGE, rounded=True, adj=0.5)
    text(s, x, y, 0.30, 0.30, letter, size=12, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x + 0.40, y + 0.03, 3.20, 0.26, word, size=10, bold=True, color=ORANGE)


def num(v, d=0):
    return f"{v:.{d}f}".replace(".", ",")


# ══ 01 · ОБКЛАДИНКА ═══════════════════════════════════════════════════════
s = slide(dark=True)
cover_pic(s, BG_TITLE, 0, 0, W, H)
s.shapes.add_picture(LOGO, Inches(M), Inches(0.52), Inches(1.94), Inches(0.84))
rect(s, M, 2.02, 0.54, 0.09, ORANGE)
text(s, M, 2.34, 7.00, 0.34, "ДОСЛІДЖЕННЯ РИНКУ · ВЕРЕСЕНЬ 2026", size=11.5,
     bold=True, color=AMBER)
text(s, M, 2.78, 7.20, 1.70, "Готовий рис\nв Україні", size=46, bold=True,
     color=WHITE, line=1.04)
text(s, M, 4.52, 6.80, 0.36, "Хто представлений · яка упаковка · яка ціна · де представлено",
     size=14.5, color=RGBColor(0xC6, 0xCE, 0xE2))
rect(s, M, 5.12, 5.90, 0.055, RGBColor(0x55, 0x60, 0x8C))
text(s, M, 5.34, 6.80, 0.80,
     f"9 брендів · {N_CARDS} перевірених карток товару · 16 каналів продажу\n"
     "Зріз відкритих українських онлайн-вітрин: 22 вересня 2026 р.",
     size=11.5, color=RGBColor(0x9F, 0xA9, 0xC4), line=1.40)

for im, x, y, w, h in [("md/bens_basmati250.jpg", 8.26, 1.22, 2.35, 2.62),
                       ("md/bibigo_bowl.jpg",    10.78, 1.60, 2.35, 2.24),
                       ("md/clearspring.jpg",     8.26, 3.98, 2.35, 2.10),
                       ("md/ottogi_burger.jpg",  10.78, 3.98, 2.35, 2.10)]:
    rect(s, x, y, w, h, WHITE, rounded=True, adj=0.06)
    pic(s, im, x + 0.16, y + 0.14, w - 0.32, h - 0.28)
rect(s, 8.26, 6.22, 4.87, 0.78, RGBColor(0x26, 0x2F, 0x4D), rounded=True, adj=0.16)
text(s, 8.52, 6.34, 4.40, 0.56,
     "Ben's Original · Bibigo · Clearspring · Ottogi · Portion\n"
     "Haidilao · Mo Xiao Xian · Travellunch · SubliMate",
     size=10, bold=True, color=WHITE, line=1.30)

# ══ 02 · КАРТА РИНКУ ══════════════════════════════════════════════════════
s = slide()
header(s, "КАРТА РИНКУ", "Що показав зріз вітрин",
       "Чотири питання дослідження та коротка відповідь на кожне.")
for i, (v, u, lb, sub, c) in enumerate([
        ("9", "ХТО", "брендів у категорії та поруч",
         "5 — готовий рис, 4 — суміжні технології", ORANGE),
        ("210–350", "УПАКОВКА", "грамів",
         "пауч, лоток-чаша та висока чаша", TEAL),
        ("18–130", "ЦІНА", "грн за 100 г",
         f"медіана {num(CORE_MED)} грн; за упаковку 45–356 грн", PLUM),
        ("16", "ДЕ", "каналів продажу",
         "мережа, маркетплейси, спеціалізований онлайн", GREEN)]):
    kpi(s, M + i * 3.07, 2.04, 2.82, v, u, lb, sub, c)

rect(s, M, 4.72, 11.98, 1.92, MIST, rounded=True, adj=0.07)
rect(s, M, 4.72, 0.10, 1.92, ORANGE, rounded=True, adj=0.5)
text(s, M + 0.42, 4.98, 5.60, 0.34, "Структура пропозиції", size=16, bold=True, color=NAVY)
yy = 5.44
for t, c in [("Гарніри — чистий рис у пауці: Ben's Original, Clearspring.", ORANGE),
             ("Рис у чаші: Bibigo — білий рис, Ottogi — рис із наповнювачем.", GREEN),
             ("Суміжні технології: саморозігрів, сублімація — інша ціна й повід.", PLUM)]:
    dot(s, M + 0.48, yy + 0.13, 0.13, c)
    text(s, M + 0.70, yy, 5.40, 0.30, t, size=11, color=INK)
    yy += 0.38
rect(s, 7.60, 5.02, 0.04, 1.32, MIST_D)
text(s, 7.94, 4.98, 4.36, 1.36,
     "Основа — онлайн-асортимент відкритих вітрин. Частота карток не дорівнює "
     "частці продажів: один товар може бути виставлений у кількох продавців "
     "за різними цінами.",
     size=11, color=GREY, line=1.32)
foot(s, "Джерело: перевірені картки українських онлайн-вітрин, 22.09.2026.")

# ══ 03 · ХТО · СЕГМЕНТИ ═══════════════════════════════════════════════════
s = slide()
header(s, "ХТО ПРЕДСТАВЛЕНИЙ · 1/4", "Три способи продати готовий рис",
       "Сегменти відрізняються не смаком, а технологією та поводом споживання.")
qtag(s, M, 1.94, "1", "ХТО ПРЕДСТАВЛЕНИЙ")
SEG = [("md/bens_basmati250.jpg", "ГАРНІР У ПАУЧІ", "Ben's Original · Clearspring",
        "Чистий рис або легкий смак.\nПауч 220–250 г, розігрів 2 хв.",
        "15 карток · 18–101 грн/100 г", ORANGE),
       ("md/bibigo_bowl.jpg", "РИС У ЧАШІ", "Bibigo · Ottogi · Portion",
        "Біла основа або рис із наповнювачем.\nЛоток-чаша 210–350 г.",
        "16 карток · 31–130 грн/100 г", GREEN),
       (None, "СУМІЖНІ ТЕХНОЛОГІЇ", "Haidilao · Mo Xiao Xian\nTravellunch · SubliMate",
        "Саморозігрів і сублімація.\nІнший повід, інша цінова ліга.",
        "14 карток · 220–454 грн/100 г", PLUM)]
for i, (img, kind, brands, desc, n, c) in enumerate(SEG):
    x = M + i * 4.10
    rect(s, x, 2.40, 3.84, 4.28, MIST if i < 2 else WHITE,
         line=None if i < 2 else MIST_D, lw=1.2, rounded=True, adj=0.05)
    rect(s, x, 2.40, 3.84, 0.10, c, rounded=True, adj=0.5)
    rect(s, x + 0.20, 2.66, 3.44, 1.72, WHITE, rounded=True, adj=0.05)
    if img:
        pic(s, img, x + 0.34, 2.74, 3.16, 1.56)
    else:
        text(s, x + 0.20, 3.34, 3.44, 0.26, "ПАКШОТІВ НЕМАЄ", size=9.5, bold=True,
             color=GREY, align=PP_ALIGN.CENTER)
        text(s, x + 0.20, 3.60, 3.44, 0.24, "на картках цих продавців", size=9,
             color=GREY, align=PP_ALIGN.CENTER)
    text(s, x + 0.32, 4.52, 3.20, 0.24, kind, size=9.5, bold=True, color=c)
    text(s, x + 0.32, 4.80, 3.20, 0.54, brands, size=14, bold=True, color=NAVY, line=1.14)
    text(s, x + 0.32, 5.44, 3.20, 0.72, desc, size=11, color=GREY, line=1.30)
    rect(s, x + 0.32, 6.18, 3.20, 0.02, MIST_D)
    text(s, x + 0.32, 6.30, 3.20, 0.26, n, size=10.5, bold=True, color=c)
foot(s, "Джерело: картки товарів українських онлайн-вітрин, 22.09.2026. "
        "Ціни за 100 г — наш розрахунок за масою та ціною картки.")

# ══ 04 · ХТО · БРЕНДИ ГОТОВОГО РИСУ ═══════════════════════════════════════
s = slide()
header(s, "ХТО ПРЕДСТАВЛЕНИЙ · 2/4", "П'ять брендів готового рису",
       "Кожен займає власну нішу за форматом, походженням і ціною.")
qtag(s, M, 1.94, "1", "ХТО ПРЕДСТАВЛЕНИЙ")
BR = [("md/bens_lang220.jpg", "Ben's Original", "Mars · ЄС", "Пауч 220 / 240 / 250 г",
       "Найширша лінійка: 11 смаків", ORANGE),
      ("md/ottogi_kimchi.jpg", "Ottogi", "Корея", "Чаша 217–320 г",
       "Рис із наповнювачем, 10 страв", PLUM),
      ("md/bibigo_bowl2.jpg", "Bibigo", "CJ CheilJedang · Корея", "Лоток-чаша 210 г",
       "Білий рис без добавок", GREEN),
      ("md/clearspring.jpg", "Clearspring", "Велика Британія", "Пауч 250 г",
       "Органіка: суміш рисів із tamari", TEAL),
      (None, "Portion", "Пирятинський делікатес · Україна", "Лоток 350 г",
       "Єдиний український виробник", ROSE)]
for lx, lw, lb in [(M + 1.52, 2.90, "БРЕНД І ПОХОДЖЕННЯ"), (M + 4.60, 2.60, "ФОРМАТ"),
                   (M + 7.30, 3.10, "ЩО САМЕ"), (M + 10.50, 1.70, "ГРН ЗА УПАКОВКУ")]:
    text(s, lx, 2.34, lw, 0.24, lb, size=8.5, bold=True, color=GREY,
         align=PP_ALIGN.RIGHT if "ГРН" in lb else PP_ALIGN.LEFT)
for i, (img, brand, origin, fmt, what, c) in enumerate(BR):
    y = 2.64 + i * 0.86
    lo, hi, pk_lo, pk_hi, n = rng(brand)
    rect(s, M, y, 11.98, 0.80, MIST if i % 2 == 0 else WHITE, rounded=True, adj=0.16)
    rect(s, M, y, 0.09, 0.80, c, rounded=True, adj=0.5)
    rect(s, M + 0.22, y + 0.08, 1.02, 0.64, WHITE, rounded=True, adj=0.12)
    if img:
        pic(s, img, M + 0.27, y + 0.11, 0.92, 0.58)
    else:
        text(s, M + 0.22, y + 0.28, 1.02, 0.24, "фото\nнемає", size=7, color=GREY,
             align=PP_ALIGN.CENTER, line=1.10)
    text(s, M + 1.52, y + 0.13, 3.00, 0.28, brand, size=13.5, bold=True, color=NAVY)
    text(s, M + 1.52, y + 0.44, 3.00, 0.24, origin, size=9, color=GREY)
    text(s, M + 4.60, y + 0.26, 2.60, 0.28, fmt, size=11, color=INK)
    text(s, M + 7.30, y + 0.26, 3.10, 0.28, what, size=11, color=INK)
    rng_txt = (f"{pk_lo:g} – {pk_hi:g}" if pk_lo != pk_hi else f"{pk_lo:g}")
    text(s, M + 10.40, y + 0.16, 1.80, 0.34, rng_txt.replace(".", ","), size=14,
         bold=True, color=c, align=PP_ALIGN.RIGHT)
    text(s, M + 10.40, y + 0.50, 1.80, 0.22, f"{n} карток", size=8.5, color=GREY,
         align=PP_ALIGN.RIGHT)
foot(s, "Джерела: Сільпо, MAUDAU, Edison Lee, Смак Кореї, Rozetka, Prom, Pulsar, "
        "Апетітаріум, Gurmissimo, Тайякі Март — 22.09.2026.")

# ══ 05 · ХТО · СУМІЖНІ ТЕХНОЛОГІЇ ═════════════════════════════════════════
s = slide()
header(s, "ХТО ПРЕДСТАВЛЕНИЙ · 3/4", "Чотири бренди поруч із категорією",
       "Їх видає той самий пошук, але це інша технологія та інша цінова ліга.")
qtag(s, M, 1.94, "1", "ХТО ПРЕДСТАВЛЕНИЙ")
NEAR = [(None, "Haidilao", "Китай", "САМОРОЗІГРІВ",
         "Чаша 175–272 г із хімічним нагрівачем", "650 – 799 грн", "283–371", PLUM),
        (None, "Mo Xiao Xian", "Китай", "ШВИДКЕ ПРИГОТУВАННЯ",
         "Instant Rice 275 г, окріп", "ціну не вказано", "—", SLATE),
        (None, "Travellunch", "Німеччина", "СУБЛІМАЦІЯ",
         "Пакет 125 / 250 г, потребує окропу", "275 – 889 грн", "220–454", TEAL),
        (None, "SubliMate", "Україна", "СУБЛІМАЦІЯ",
         "Порційний пакет, туристичний канал", "ціна за запитом", "—", CORAL)]
for i, (img, brand, origin, tech, what, price, per, c) in enumerate(NEAR):
    x = M + (i % 2) * 6.12
    y = 2.42 + (i // 2) * 2.16
    rect(s, x, y, 5.86, 1.96, WHITE, line=MIST_D, lw=1.2, rounded=True, adj=0.07)
    rect(s, x, y, 5.86, 0.09, c, rounded=True, adj=0.5)
    rect(s, x + 0.22, y + 0.24, 1.50, 1.48, MIST, rounded=True, adj=0.08)
    if img:
        pic(s, img, x + 0.30, y + 0.30, 1.34, 1.36)
    else:
        text(s, x + 0.22, y + 0.86, 1.50, 0.24, "фото\nнемає", size=8, color=GREY,
             align=PP_ALIGN.CENTER, line=1.12)
    text(s, x + 1.92, y + 0.28, 2.60, 0.30, brand, size=15, bold=True, color=NAVY)
    text(s, x + 1.92, y + 0.60, 2.60, 0.24, origin, size=9.5, color=GREY)
    pill(s, x + 1.92, y + 0.90, tech, c, size=8)
    text(s, x + 1.92, y + 1.28, 3.70, 0.44, what, size=10, color=GREY, line=1.22)
    text(s, x + 4.00, y + 0.26, 1.66, 0.30, price, size=13, bold=True, color=c,
         align=PP_ALIGN.RIGHT)
    text(s, x + 4.00, y + 0.58, 1.66, 0.24,
         (per + " грн/100 г") if per != "—" else "", size=8.5, color=GREY,
         align=PP_ALIGN.RIGHT)
rect(s, M, 6.70, 11.98, 0.02, MIST_D)
text(s, M, 6.80, 11.98, 0.26,
     "Канали: Daruy, Tactico, ALANTUR, ForCamp, Highlander, Військторг Гайдамака — "
     "переважно туристичні та військові магазини, а не продуктова роздріб.",
     size=9, color=GREY, line=1.16)

# ══ 06 · ХТО · BEN'S ORIGINAL ═════════════════════════════════════════════
s = slide()
header(s, "ХТО ПРЕДСТАВЛЕНИЙ · 4/4", "Ben's Original: 11 смаків — найглибша лінійка",
       "Для порівняння: у решти брендів гарнірів по одному-трьох варіантах.")
WALL = [("md/bens_basmati250.jpg", "Basmati", "250 г", "89,99"),
        ("md/bens_mediterran.jpg", "Mediterran", "250 г", "89,99"),
        ("md/bens_longgrain.jpg", "Long Grain", "250 г", "45"),
        ("md/bens_risibisi.jpg", "Risi Bisi", "250 г", "89,99"),
        ("md/bens_curry_indien.jpg", "Indian Curry", "250 г", "89,99"),
        ("md/bens_curry_linsen.jpg", "Curry · сочевиця", "220 г", "89,99"),
        ("md/bens_sweetchili.jpg", "Sweet Chili", "220 г", "179"),
        ("md/bens_mexikanisch.jpg", "Mexikanisch", "220 г", "109"),
        ("md/bens_curryreis.jpg", "Curryreis Indien", "220 г", "144"),
        ("md/bens_basmati220.jpg", "Basmati", "220 г", "144"),
        ("md/bens_stickybowl.jpg", "Для боулів", "220 г", "149")]
cw, gap = 1.60, 0.14
for i, (img, nm, g, p) in enumerate(WALL):
    col = i if i < 7 else i - 7
    x = M + col * (cw + gap)
    y = 2.04 if i < 7 else 4.52
    rect(s, x, y, cw, 2.30, WHITE, line=MIST_D, lw=1.0, rounded=True, adj=0.08)
    rect(s, x + 0.10, y + 0.10, cw - 0.20, 1.26, MIST, rounded=True, adj=0.08)
    pic(s, img, x + 0.16, y + 0.14, cw - 0.32, 1.18)
    text(s, x + 0.12, y + 1.46, cw - 0.24, 0.32, nm, size=8.5, bold=True, color=NAVY,
         line=1.10)
    text(s, x + 0.12, y + 1.82, cw - 0.24, 0.22, g, size=8, color=GREY)
    text(s, x + 0.12, y + 2.02, cw - 0.24, 0.22, p + " грн", size=10.5, bold=True,
         color=ORANGE)
rect(s, 7.66, 4.52, 4.98, 2.30, NAVY, rounded=True, adj=0.07)
rect(s, 7.66, 4.52, 0.09, 2.30, ORANGE, rounded=True, adj=0.5)
text(s, 8.02, 4.76, 4.36, 0.30, "Дві вагові платформи", size=15, bold=True, color=AMBER)
for i, (lb, vl, sub) in enumerate([("250 г", "45 – 89,99 грн", "5 смаків"),
                                   ("220 г", "89,99 – 179 грн", "6 смаків")]):
    yy = 5.20 + i * 0.62
    text(s, 8.02, yy, 0.90, 0.30, lb, size=13, bold=True, color=WHITE)
    text(s, 8.96, yy + 0.02, 2.10, 0.28, vl, size=12, bold=True, color=WHITE)
    text(s, 11.12, yy + 0.04, 1.30, 0.26, sub, size=9.5, color=RGBColor(0x9F, 0xA9, 0xC4))
text(s, 8.02, 6.32, 4.36, 0.40,
     "Плюс Bio Basmati 240 г і Long Grain 220 г у спеціалізованому онлайні.",
     size=10, color=RGBColor(0x9F, 0xA9, 0xC4), line=1.22)
foot(s, "Джерела: каталог «Сільпо» (11 карток), MAUDAU, Edison Lee — 22.09.2026.")

# ══ 07 · ХТО · OTTOGI ═════════════════════════════════════════════════════
s = slide()
header(s, "OTTOGI · ГЛИБИНА ЛІНІЙКИ", "Ottogi: 10 страв у чаші, ціна залежить від продавця",
       "Другий за глибиною бренд категорії. Один товар трапляється у чотирьох продавців.")
OT = [("md/ottogi_chicken.jpg", "Гострі курячі ребра", "310 г", "255 – 260"),
      ("md/ottogi_hamburg.jpg", "Гамбурзький стейк", "315 г", "255"),
      ("md/ottogi_octopus.jpg", "Гострий восьминіг", "280 г", "255 – 289"),
      ("md/ottogi_jjampong.jpg", "Jin Jjampong", "217,5 г", "255"),
      ("md/ottogi_tuna.jpg", "Тунець і майонез", "247 г", "225 – 320"),
      ("md/ottogi_kimchi.jpg", "Кімчі й тунець", "310 г", "255")]
for i, (img, nm, g, p) in enumerate(OT):
    x = M + (i % 3) * 3.00
    y = 2.04 + (i // 3) * 2.36
    rect(s, x, y, 2.72, 2.20, WHITE, line=MIST_D, lw=1.0, rounded=True, adj=0.07)
    rect(s, x + 0.12, y + 0.12, 2.48, 1.16, MIST, rounded=True, adj=0.08)
    pic(s, img, x + 0.20, y + 0.16, 2.32, 1.08)
    text(s, x + 0.18, y + 1.38, 2.36, 0.32, nm, size=9.5, bold=True, color=NAVY, line=1.10)
    text(s, x + 0.18, y + 1.72, 2.36, 0.22, g, size=8.5, color=GREY)
    text(s, x + 0.18, y + 1.92, 2.36, 0.24, p + " грн", size=11, bold=True, color=PLUM)

rect(s, 9.68, 2.04, 2.96, 4.52, MIST, rounded=True, adj=0.06)
rect(s, 9.68, 2.04, 2.96, 0.10, PLUM, rounded=True, adj=0.5)
text(s, 9.96, 2.32, 2.44, 0.30, "Ще 4 смаки", size=14, bold=True, color=NAVY)
text(s, 9.96, 2.66, 2.44, 0.30, "без фото на картках", size=9.5, color=GREY)
yy = 3.10
for nm, g, p in [("Пібімпаб", "269 г", "252 – 334"), ("Гостра свинина", "269 г", "334"),
                 ("Свинина", "310 г", "252"), ("Бульгогі / яловичина", "320 г", "252 – 356")]:
    text(s, 9.96, yy, 2.44, 0.24, nm, size=10.5, bold=True, color=INK)
    text(s, 9.96, yy + 0.22, 1.10, 0.22, g, size=8.5, color=GREY)
    text(s, 11.10, yy + 0.20, 1.30, 0.24, p + " грн", size=9.5, bold=True, color=PLUM,
         align=PP_ALIGN.RIGHT)
    yy += 0.62
rect(s, 9.96, 5.66, 2.44, 0.02, MIST_D)
text(s, 9.96, 5.80, 2.44, 0.62,
     "Розкид цін на один товар — до 42 % між продавцями.",
     size=10, bold=True, color=PLUM, line=1.24)
foot(s, "Джерела: Pulsar, Апетітаріум, Gurmissimo, Тайякі Март, Rozetka — 22.09.2026.")

# ══ 08 · УПАКОВКА · ГРАМАЖ ════════════════════════════════════════════════
s = slide()
header(s, "ЯКА УПАКОВКА · 1/2", "Два вагові кластери: 210–250 г і 247–350 г",
       "Гарніри тримаються 220–250 г, страви з наповнювачем важчі за визначенням.")
qtag(s, M, 1.94, "2", "ЯКА УПАКОВКА")
BARS = [("220 г", 6, "Ben's Original", ORANGE), ("250 г", 6, "Ben's · Clearspring", ORANGE),
        ("210 г", 1, "Bibigo", GREEN), ("240 г", 1, "Ben's Bio", ORANGE),
        ("247–320 г", 10, "Ottogi", PLUM), ("350 г", 1, "Portion", ROSE)]
text(s, M, 2.42, 5.60, 0.26, "ВАРІАНТІВ У КОЖНІЙ ВАЗІ", size=9.5, bold=True, color=GREY)
yy = 2.78
for lb, n, who, c in BARS:
    text(s, M, yy, 1.20, 0.34, lb, size=12.5, bold=True, color=NAVY,
         anchor=MSO_ANCHOR.MIDDLE)
    rect(s, M + 1.30, yy + 0.07, 3.00, 0.22, MIST, rounded=True, adj=0.5)
    rect(s, M + 1.30, yy + 0.07, max(0.18, 0.29 * n), 0.22, c, rounded=True, adj=0.5)
    text(s, M + 4.42, yy, 0.40, 0.34, str(n), size=12.5, bold=True, color=c,
         anchor=MSO_ANCHOR.MIDDLE)
    text(s, M + 4.92, yy, 2.00, 0.34, who, size=9.5, color=GREY, anchor=MSO_ANCHOR.MIDDLE)
    yy += 0.58
rect(s, M, 6.26, 6.70, 0.62, MIST, rounded=True, adj=0.22)
text(s, M + 0.28, 6.38, 6.20, 0.40,
     "Одного лідера за грамажем немає: 220 і 250 г представлені порівну.",
     size=11, bold=True, color=NAVY)

GAL = [("md/bens_basmati220.jpg", "220 г", "пауч"), ("md/bens_basmati250.jpg", "250 г", "пауч"),
       ("md/bibigo_bowl.jpg", "210 г", "лоток-чаша"), ("md/ottogi_octopus.jpg", "280 г", "чаша")]
for i, (img, cap, n) in enumerate(GAL):
    x = 7.66 + (i % 2) * 2.52
    y = 2.42 + (i // 2) * 2.26
    rect(s, x, y, 2.30, 2.06, MIST, rounded=True, adj=0.07)
    pic(s, img, x + 0.16, y + 0.12, 1.98, 1.42)
    text(s, x, y + 1.58, 2.30, 0.26, cap, size=12, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER)
    text(s, x, y + 1.80, 2.30, 0.22, n, size=9, color=GREY, align=PP_ALIGN.CENTER)
foot(s, "Джерело: маса за картками товарів, 22.09.2026. Суміжні технології "
        "(саморозігрів, сублімація) у цей розподіл не входять.")

# ══ 09 · УПАКОВКА · ТИПИ ══════════════════════════════════════════════════
s = slide()
header(s, "ЯКА УПАКОВКА · 2/2", "Три типи паковання на вітрині",
       "Тип паковання визначає і спосіб споживання, і полицю, на яку товар потрапляє.")
qtag(s, M, 1.94, "2", "ЯКА УПАКОВКА")
TYPES = [("md/bens_basmati250.jpg", "ПАУЧ", "220 · 240 · 250 г",
          "Ben's Original · Clearspring", "Плаский реторт-пауч.",
          "Рис висипається у тарілку.", "15 карток", ORANGE),
         ("md/bibigo_bowl.jpg", "ЛОТОК-ЧАША", "210 г",
          "Bibigo", "Пласка чаша з плівкою.",
          "Білий рис, їдять з упаковки.", "3 картки", GREEN),
         ("md/ottogi_octopus.jpg", "ВИСОКА ЧАША", "217 – 350 г",
          "Ottogi · Portion", "Глибока чаша з наповнювачем.",
          "Самостійний обід, не гарнір.", "13 карток", PLUM)]
for i, (img, kind, mass, brands, pack, use, share, c) in enumerate(TYPES):
    x = M + i * 4.10
    rect(s, x, 2.40, 3.84, 4.28, WHITE, line=MIST_D, lw=1.2, rounded=True, adj=0.05)
    rect(s, x, 2.40, 3.84, 0.10, c, rounded=True, adj=0.5)
    rect(s, x + 0.20, 2.66, 3.44, 1.58, MIST, rounded=True, adj=0.05)
    pic(s, img, x + 0.34, 2.74, 3.16, 1.42)
    text(s, x + 0.32, 4.38, 3.20, 0.30, kind, size=14, bold=True, color=c)
    text(s, x + 0.32, 4.72, 3.20, 0.28, mass, size=13, bold=True, color=NAVY)
    rect(s, x + 0.32, 5.08, 3.20, 0.02, MIST_D)
    text(s, x + 0.32, 5.20, 3.20, 0.26, brands, size=10.5, bold=True, color=INK)
    text(s, x + 0.32, 5.50, 3.20, 0.28, pack, size=10, color=GREY, line=1.22)
    text(s, x + 0.32, 5.82, 3.20, 0.28, use, size=10, color=GREY, line=1.22)
    pill(s, x + 0.32, 6.22, share, c, size=8)
foot(s, "Джерело: тип паковання за фото та описом карток, 22.09.2026.")

# ══ 10 · ЦІНА · ЗА УПАКОВКУ ═══════════════════════════════════════════════
s = slide()
header(s, "ЯКА ЦІНА · 1/2", "За упаковку: від 45 до 356 грн",
       "Крайні точки кожного бренду. Один товар у різних продавців має різну ціну.")
qtag(s, M, 1.94, "3", "ЯКА ЦІНА")
ROWS = [("md/bens_longgrain.jpg", "Ben's Long Grain", "250 г · Сільпо", 45, ORANGE),
        ("md/bens_basmati250.jpg", "Ben's Basmati", "250 г · Сільпо", 90, ORANGE),
        (None, "Portion рис із куркою", "350 г · Rozetka", 109, ROSE),
        ("md/bibigo_bowl.jpg", "Bibigo білий рис", "210 г · Смак Кореї", 135, GREEN),
        ("md/bens_sweetchili.jpg", "Ben's Sweet Chili", "220 г · Сільпо", 179, ORANGE),
        ("md/bibigo_bowl2.jpg", "Bibigo білий рис", "210 г · Prom", 189, GREEN),
        ("md/ottogi_tuna.jpg", "Ottogi тунець", "247 г · Тайякі Март", 225, PLUM),
        ("md/clearspring.jpg", "Clearspring Brown & Wild", "250 г · MAUDAU", 252, TEAL),
        ("md/ottogi_hamburg.jpg", "Ottogi яловичина", "320 г · Апетітаріум", 356, PLUM)]
MAXP, BX, BW = 370.0, 5.20, 4.60
yy = 2.36
for img, nm, chan, price, c in ROWS:
    rect(s, M, yy, 11.98, 0.48, MIST if yy == 2.36 else WHITE,
         line=None if yy == 2.36 else MIST_D, lw=1.0, rounded=True, adj=0.28)
    rect(s, M + 0.08, yy + 0.05, 0.50, 0.38, WHITE, rounded=True, adj=0.16)
    if img:
        pic(s, img, M + 0.11, yy + 0.07, 0.44, 0.34)
    text(s, M + 0.72, yy + 0.02, 2.90, 0.24, nm, size=11, bold=True, color=NAVY)
    text(s, M + 0.72, yy + 0.25, 2.90, 0.22, chan, size=8.5, color=GREY)
    rect(s, BX, yy + 0.17, BW, 0.14, MIST_D, rounded=True, adj=0.5)
    rect(s, BX, yy + 0.17, max(0.12, BW * price / MAXP), 0.14, c, rounded=True, adj=0.5)
    text(s, 10.20, yy + 0.06, 1.24, 0.34, f"{price} грн", size=13, bold=True, color=NAVY,
         align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    per = price / dict(zip([r[1] for r in ROWS], [250, 250, 350, 210, 220, 210, 247, 250, 320]))[nm] * 100
    yy += 0.505
for t in (0, 100, 200, 300):
    x = BX + BW * t / MAXP
    rect(s, x, 6.92, 0.02, 0.12, MIST_D)
    text(s, x - 0.34, 7.06, 0.68, 0.22, str(t), size=8.5, color=GREY, align=PP_ALIGN.CENTER)
text(s, BX + BW + 0.10, 7.04, 1.40, 0.24, "грн за упаковку", size=8.5, bold=True, color=GREY)

# ══ 11 · ЦІНА · ЗА 100 Г ══════════════════════════════════════════════════
s = slide()
header(s, "ЯКА ЦІНА · 2/2", "За 100 грамів: кожен бренд у своєму коридорі",
       "Смуга показує розкид цін бренду; цифри — крайні точки в грн за 100 г.")
qtag(s, M, 1.94, "3", "ЯКА ЦІНА")

LO, HI = 0.0, 140.0
AX_X, AX_W = 4.10, 7.10
BRND = [("md/bens_longgrain.jpg", "Ben's Original", "пауч 220–250 г", ORANGE),
        ("md/bibigo_bowl.jpg", "Bibigo", "чаша 210 г", GREEN),
        ("md/clearspring.jpg", "Clearspring", "пауч 250 г", TEAL),
        (None, "Portion", "лоток 350 г", ROSE),
        ("md/ottogi_octopus.jpg", "Ottogi", "чаша 217–320 г", PLUM)]


def axx(v):
    return AX_X + AX_W * (v - LO) / (HI - LO)


rect(s, M, 2.34, 11.98, 3.52, MIST, rounded=True, adj=0.05)
for t in range(0, 141, 20):
    rect(s, axx(t), 2.52, 0.015, 3.06, RGBColor(0xDC, 0xE2, 0xEE))
    text(s, axx(t) - 0.36, 5.60, 0.72, 0.24, str(t), size=9.5, color=GREY,
         align=PP_ALIGN.CENTER)
text(s, M + 0.84, 5.56, 3.20, 0.24, "ГРН ЗА 100 ГРАМІВ", size=9, bold=True, color=GREY)

yy = 2.96
for img, brand, fmt, c in BRND:
    lo, hi, *_ = rng(brand)
    rect(s, M + 0.18, yy - 0.19, 0.52, 0.40, WHITE, rounded=True, adj=0.16)
    if img:
        pic(s, img, M + 0.21, yy - 0.17, 0.46, 0.36)
    text(s, M + 0.84, yy - 0.20, 2.10, 0.26, brand, size=12, bold=True, color=NAVY)
    text(s, M + 0.84, yy + 0.04, 2.10, 0.22, fmt, size=8.5, color=GREY)
    x0, x1 = axx(lo), axx(hi)
    if x1 - x0 < 0.16:                       # один рівень ціни
        dot(s, x0, yy, 0.22, c)
        text(s, x0 - 0.60, yy - 0.44, 1.20, 0.24, num(lo), size=11, bold=True,
             color=c, align=PP_ALIGN.CENTER)
    else:
        rect(s, x0, yy - 0.075, x1 - x0, 0.15, c, rounded=True, adj=0.5)
        dot(s, x0, yy, 0.22, c); dot(s, x1, yy, 0.22, c)
        text(s, x0 - 0.62, yy - 0.44, 1.20, 0.24, num(lo), size=11, bold=True,
             color=c, align=PP_ALIGN.CENTER)
        text(s, x1 - 0.58, yy - 0.44, 1.20, 0.24, num(hi), size=11, bold=True,
             color=c, align=PP_ALIGN.CENTER)
    yy += 0.58

rect(s, M, 6.06, 5.80, 0.80, NAVY, rounded=True, adj=0.16)
text(s, M + 0.28, 6.18, 5.30, 0.56,
     [[("Медіана категорії — ", {"color": RGBColor(0xC6, 0xCE, 0xE2)}),
       (f"{num(CORE_MED)} грн за 100 г", {"bold": True, "color": AMBER}),
       (f" · 31 картка п'яти брендів", {"color": RGBColor(0xC6, 0xCE, 0xE2)})]],
     size=12)
rect(s, 6.86, 6.06, 5.78, 0.80, MIST, rounded=True, adj=0.16)
rect(s, 6.86, 6.06, 0.09, 0.80, PLUM, rounded=True, adj=0.5)
text(s, 7.16, 6.16, 5.30, 0.60,
     "Поза шкалою: саморозігрівальний Haidilao 283–371 і сублімований "
     "Travellunch 220–454 грн за 100 г.", size=10.5, color=INK, line=1.24)
foot(s, "Розрахунок за масою та ціною карток, 22.09.2026. Смуга — від мінімальної "
        "до максимальної ціни бренду за 100 г.")

# ══ 12 · ДЕ ПРЕДСТАВЛЕНО ══════════════════════════════════════════════════
s = slide()
header(s, "ДЕ ПРЕДСТАВЛЕНО", "Шістнадцять каналів: від мережі до туристичних магазинів",
       "Категорія майже повністю живе в онлайні — мережева роздріб представлена одним гравцем.")
qtag(s, M, 1.94, "4", "ДЕ ПРЕДСТАВЛЕНО")
GROUPS = [("МЕРЕЖЕВА РОЗДРІБ", ORANGE, ["Сільпо"],
           "Ben's Original — 11 смаків у каталозі"),
          ("ОНЛАЙН-СУПЕРМАРКЕТИ", GREEN, ["MAUDAU", "Rozetka", "Prom"],
           "Ben's Original · Clearspring · Bibigo · Ottogi · Portion"),
          ("СПЕЦІАЛІЗОВАНИЙ АЗІЙСЬКИЙ", PLUM,
           ["Смак Кореї", "Тайякі Март", "Pulsar", "Апетітаріум", "Gurmissimo", "Edison Lee"],
           "Bibigo · Ottogi · Ben's Original"),
          ("ТУРИСТИЧНІ ТА ВІЙСЬКОВІ", TEAL,
           ["Daruy", "Tactico", "ALANTUR", "ForCamp", "Highlander", "Військторг Гайдамака"],
           "Haidilao · Travellunch · SubliMate")]
yy = 2.40
for title, c, chans, what in GROUPS:
    rect(s, M, yy, 11.98, 1.02, MIST if yy in (2.40, 4.58) else WHITE, rounded=True, adj=0.14)
    rect(s, M, yy, 0.09, 1.02, c, rounded=True, adj=0.5)
    text(s, M + 0.30, yy + 0.16, 3.00, 0.26, title, size=10, bold=True, color=c)
    text(s, M + 0.30, yy + 0.46, 3.30, 0.40, what, size=10, color=GREY, line=1.20)
    cx = M + 3.90
    for ch in chans:
        w = pill(s, cx, yy + 0.36, ch, c, size=9)
        cx += w + 0.16
    text(s, M + 11.20, yy + 0.34, 0.62, 0.34, str(len(chans)), size=17, bold=True,
         color=c, align=PP_ALIGN.RIGHT)
    yy += 1.09
rect(s, M, 6.82, 11.98, 0.02, MIST_D)
text(s, M, 6.94, 11.98, 0.30,
     "У пошуку не знайдено порівнянних позицій в АТБ, Novus, Metro, Varus, Auchan, "
     "Fozzy та «Таврія В». Перевірка велася онлайн, без обходу полиці.",
     size=9, color=GREY, line=1.16)

prs.save("Gotovyi_Rys_Doslidzhennia_Rynku.pptx")
print("saved ·", len(prs.slides._sldIdLst), "slides")
