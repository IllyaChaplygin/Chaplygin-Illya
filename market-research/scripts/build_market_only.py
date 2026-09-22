"""Готовий рис · Україна — дослідження ринку.

Чотири питання і нічого понад них: хто представлений, яка упаковка, яка ціна,
де представлено. Жодних наших позицій, собівартості чи цінових сценаріїв —
полична ціна рахується окремо.

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


# ── дані дослідження ──────────────────────────────────────────────────────
# (бренд, варіант, канал, грамів, грн, у наявності)
CARDS = [
    ("Ben's Original", "Long Grain",        "Сільпо",     250, 45.00, False),
    ("Ben's Original", "Basmati",           "Сільпо",     250, 89.99, False),
    ("Ben's Original", "Mediterran",        "Сільпо",     250, 89.99, False),
    ("Ben's Original", "Risi Bisi",         "Сільпо",     250, 89.99, False),
    ("Ben's Original", "Indian Curry",      "Сільпо",     250, 89.99, False),
    ("Ben's Original", "Curry з сочевицею", "Сільпо",     220, 89.99, False),
    ("Ben's Original", "Basmati",           "MAUDAU",     220, 100.00, False),
    ("Ben's Original", "Mexikanisch",       "Сільпо",     220, 109.00, False),
    ("Ben's Original", "Bio Basmati",       "Edison Lee", 240, 139.00, None),
    ("Ben's Original", "Long Grain",        "Edison Lee", 220, 139.00, None),
    ("Bibigo",         "Білий рис",         "Смак Кореї", 210, 135.00, True),
    ("Ben's Original", "Curryreis Indien",  "Сільпо",     220, 144.00, False),
    ("Ben's Original", "Basmati",           "Сільпо",     220, 144.00, False),
    ("Ben's Original", "Для боулів",        "Сільпо",     220, 149.00, False),
    ("Ben's Original", "Sweet Chili",       "Сільпо",     220, 179.00, False),
    ("Bibigo",         "Білий рис",         "Rozetka",    210, 184.00, False),
    ("Bibigo",         "Білий рис",         "Prom",       210, 189.00, False),
    ("Clearspring",    "Brown & Wild",      "MAUDAU",     250, 252.00, False),
]
PER = sorted(p / g * 100 for *_, g, p, _ in CARDS)
PER_MIN, PER_MED, PER_MAX = PER[0], st.median(PER), PER[-1]


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
        text(s, M, HDR + 0.30, 11.9, 0.34, dek, size=12.5, color=GREY, line=1.20)


def foot(s, src):
    text(s, M, 7.04, 11.9, 0.26, src, size=8, color=GREY, line=1.14)


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
    """Позначка питання дослідження — ХТО / УПАКОВКА / ЦІНА / ДЕ."""
    rect(s, x, y, 0.30, 0.30, ORANGE, rounded=True, adj=0.5)
    text(s, x, y, 0.30, 0.30, letter, size=12, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x + 0.40, y + 0.03, 2.20, 0.26, word, size=10, bold=True, color=ORANGE)


# ══ 01 · ОБКЛАДИНКА ═══════════════════════════════════════════════════════
s = slide(dark=True)
cover_pic(s, BG_TITLE, 0, 0, W, H)
s.shapes.add_picture(LOGO, Inches(M), Inches(0.52), Inches(1.94), Inches(0.84))
rect(s, M, 2.02, 0.54, 0.09, ORANGE)
text(s, M, 2.34, 7.00, 0.34, "ДОСЛІДЖЕННЯ РИНКУ · ВЕРЕСЕНЬ 2026", size=11.5,
     bold=True, color=AMBER)
text(s, M, 2.78, 7.20, 1.70, "Готовий рис\nв Україні", size=46, bold=True,
     color=WHITE, line=1.04)
text(s, M, 4.52, 6.60, 0.36, "Хто представлений · яка упаковка · яка ціна · де представлено",
     size=14.5, color=RGBColor(0xC6, 0xCE, 0xE2))
rect(s, M, 5.10, 5.90, 0.055, RGBColor(0x55, 0x60, 0x8C))
text(s, M, 5.32, 6.60, 0.80,
     "Зріз відкритих українських онлайн-вітрин: 22 вересня 2026 р.\n"
     "18 перевірених карток товару · 5 компаній · 6 каналів",
     size=11.5, color=RGBColor(0x9F, 0xA9, 0xC4), line=1.40)
rect(s, M, 6.40, 5.90, 0.60, RGBColor(0x26, 0x2F, 0x4D), rounded=True, adj=0.20)
text(s, M + 0.26, 6.52, 5.40, 0.36,
     "Полична ціна власного продукту в це дослідження не входить.",
     size=10, italic=True, color=RGBColor(0xB9, 0xC2, 0xDA))

for im, x, y, w, h in [("md/bens_basmati250.jpg", 8.26, 1.22, 2.35, 2.62),
                       ("md/bibigo_bowl.jpg",    10.78, 1.60, 2.35, 2.24),
                       ("md/clearspring.jpg",     8.26, 3.98, 2.35, 2.10),
                       ("md/ottogi_burger.jpg",  10.78, 3.98, 2.35, 2.10)]:
    rect(s, x, y, w, h, WHITE, rounded=True, adj=0.06)
    pic(s, im, x + 0.16, y + 0.14, w - 0.32, h - 0.28)
rect(s, 8.26, 6.22, 4.87, 0.78, RGBColor(0x26, 0x2F, 0x4D), rounded=True, adj=0.16)
text(s, 8.58, 6.34, 4.30, 0.56,
     [[("Ben's Original · Bibigo · Clearspring · Ottogi · Portion",
        {"bold": True, "color": WHITE})]], size=11, line=1.24)

# ══ 02 · КАРТА РИНКУ ══════════════════════════════════════════════════════
s = slide()
header(s, "КАРТА РИНКУ", "Що показав зріз вітрин",
       "Чотири питання дослідження та короткі відповіді на кожне.")
for i, (v, u, lb, sub, c) in enumerate([
        ("5", "ХТО", "компаній у категорії",
         "Ben's Original, Bibigo, Clearspring, Ottogi, Portion", ORANGE),
        ("220–250", "УПАКОВКА", "грамів — основний пауч",
         "12 із 14 порівнянних варіантів", TEAL),
        ("18–101", "ЦІНА", "грн за 100 г",
         f"медіана {PER_MED:.1f} грн; за упаковку 45–252 грн".replace(".", ","), PLUM),
        ("6", "ДЕ", "перевірених каналів",
         "мережа, маркетплейси та спеціалізований онлайн", GREEN)]):
    kpi(s, M + i * 3.07, 2.06, 2.82, v, u, lb, sub, c)

rect(s, M, 4.76, 11.98, 1.92, MIST, rounded=True, adj=0.07)
rect(s, M, 4.76, 0.10, 1.92, ORANGE, rounded=True, adj=0.5)
text(s, M + 0.42, 5.02, 5.60, 0.34, "Що ще показав зріз", size=16, bold=True, color=NAVY)
yy = 5.48
for t in ["Ben's Original формує 13 із 18 перевірених карток.",
          "Лише 1 позиція з 18 мала підтверджений залишок.",
          "Жодної позиції не знайдено в АТБ, Novus, Metro, Varus, Auchan."]:
    rect(s, M + 0.42, yy + 0.09, 0.09, 0.09, ORANGE, rounded=True, adj=0.5)
    text(s, M + 0.66, yy, 5.40, 0.30, t, size=11, color=INK)
    yy += 0.38
rect(s, 7.60, 5.06, 0.04, 1.32, MIST_D)
text(s, 7.94, 5.02, 4.36, 1.36,
     "Основа — онлайн-асортимент. Частота SKU не дорівнює частці продажів, "
     "а ціна картки без залишку є орієнтиром каталогу, а не доступною "
     "пропозицією. Офлайн-полиця не перевірялася.",
     size=11, color=GREY, line=1.32)
foot(s, "Джерела: S1–S11 · перевірені картки українських онлайн-вітрин, 22.09.2026.")

# ══ 03 · ХТО · СЕГМЕНТИ ═══════════════════════════════════════════════════
s = slide()
header(s, "ХТО ПРЕДСТАВЛЕНИЙ · 1/3", "Три сегменти категорії",
       "Готовий рис на українських вітринах розпадається на три різні пропозиції.")
qtag(s, M, 1.98, "1", "ХТО ПРЕДСТАВЛЕНИЙ")
SEG = [("md/bens_basmati250.jpg", "ПОВСЯКДЕННИЙ ГАРНІР", "Ben's Original · Bibigo",
        "Чистий рис або легкий смак.\nПауч 220–250 г, лоток-чаша 210 г.",
        "14 карток", ORANGE),
       ("md/clearspring.jpg", "ОРГАНІЧНИЙ ПРЕМІУМ", "Clearspring",
        "Суміш рисів із соусом tamari.\nПауч 250 г.", "1 картка", TEAL),
       ("md/ottogi_burger.jpg", "ПОВНОЦІННА СТРАВА", "Ottogi · Portion",
        "Рис із наповнювачем.\nЧаша 217–315 г, лоток 350 г.", "7 карток", PLUM)]
for i, (img, kind, brands, desc, n, c) in enumerate(SEG):
    x = M + i * 4.10
    rect(s, x, 2.44, 3.84, 4.24, MIST if i == 0 else WHITE,
         line=None if i == 0 else MIST_D, lw=1.2, rounded=True, adj=0.05)
    rect(s, x, 2.44, 3.84, 0.10, c, rounded=True, adj=0.5)
    rect(s, x + 0.20, 2.70, 3.44, 1.78, WHITE, rounded=True, adj=0.05)
    pic(s, img, x + 0.34, 2.78, 3.16, 1.62)
    text(s, x + 0.32, 4.62, 3.20, 0.24, kind, size=9, bold=True, color=c)
    text(s, x + 0.32, 4.90, 3.20, 0.34, brands, size=15, bold=True, color=NAVY)
    text(s, x + 0.32, 5.32, 3.20, 0.72, desc, size=11, color=GREY, line=1.30)
    pill(s, x + 0.32, 6.14, n, c)
foot(s, "Джерела: S1, S2, S3, S9, S11 · фото — картки товарів відповідних магазинів.")

# ══ 04 · ХТО · КОМПАНІЇ ═══════════════════════════════════════════════════
s = slide()
header(s, "ХТО ПРЕДСТАВЛЕНИЙ · 2/3", "П'ять компаній формують усю пропозицію",
       "Найширша лінійка — у Ben's Original. Частки продажів не встановлені.")
qtag(s, M, 1.98, "1", "ХТО ПРЕДСТАВЛЕНИЙ")
PL = [("md/bens_lang220.jpg", "Ben's Original", "Mars", "13 карток",
       "Пауч 220 / 240 / 250 г", "Сільпо · MAUDAU · Edison Lee", ORANGE),
      ("md/ottogi_kimchi.jpg", "Ottogi", "Республіка Корея", "6 карток",
       "Чаша 217–315 г", "Pulsar · продавці Rozetka", PLUM),
      ("md/bibigo_bowl2.jpg", "Bibigo", "CJ CheilJedang", "3 картки",
       "Лоток-чаша 210 г", "Смак Кореї · Rozetka · Prom", GREEN),
      ("md/clearspring.jpg", "Clearspring", "Велика Британія", "1 картка",
       "Пауч 250 г", "MAUDAU", TEAL),
      (None, "Portion", "«Пирятинський делікатес»", "1 картка",
       "Лоток 350 г · 109 грн", "Rozetka", NAVY_L)]
for lx, lw, lb in [(M + 1.52, 2.50, "БРЕНД І ВЛАСНИК"), (M + 4.20, 2.40, "ФОРМАТ"),
                   (M + 6.70, 3.30, "КАНАЛИ"), (M + 10.30, 1.00, "КАРТОК")]:
    text(s, lx, 2.30, lw, 0.24, lb, size=8.5, bold=True, color=GREY)
for i, (img, brand, owner, n, fmt, chan, c) in enumerate(PL):
    y = 2.60 + i * 0.86
    rect(s, M, y, 11.98, 0.80, MIST if i % 2 == 0 else WHITE, rounded=True, adj=0.16)
    rect(s, M, y, 0.09, 0.80, c, rounded=True, adj=0.5)
    rect(s, M + 0.22, y + 0.08, 1.02, 0.64, WHITE, rounded=True, adj=0.12)
    if img:
        pic(s, img, M + 0.27, y + 0.11, 0.92, 0.58)
    else:
        text(s, M + 0.22, y + 0.28, 1.02, 0.24, "фото\nнемає", size=7, color=GREY,
             align=PP_ALIGN.CENTER, line=1.10)
    text(s, M + 1.52, y + 0.14, 2.60, 0.28, brand, size=13.5, bold=True, color=NAVY)
    text(s, M + 1.52, y + 0.44, 2.60, 0.24, owner, size=9, color=GREY)
    text(s, M + 4.20, y + 0.26, 2.40, 0.28, fmt, size=11, color=INK)
    text(s, M + 6.70, y + 0.26, 3.40, 0.28, chan, size=11, color=INK)
    pill(s, M + 10.30, y + 0.27, n, c, w=1.00, size=9)
rect(s, M, 6.98, 11.98, 0.02, MIST_D)
text(s, M, 7.08, 11.98, 0.26,
     "Сумарно 18 перевірених карток гарнірів (Ben's Original, Bibigo, Clearspring) "
     "і 7 карток страв із наповнювачем (Ottogi, Portion).",
     size=8.5, color=GREY, line=1.16)

# ══ 05 · ХТО · ГЛИБИНА ASSORTMENT ═════════════════════════════════════════
s = slide()
header(s, "ХТО ПРЕДСТАВЛЕНИЙ · 3/3", "Ben's Original: 11 варіантів у каталозі «Сільпо»",
       "Найглибша лінійка категорії. Усі одинадцять позначені як відсутні.")
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
    y = 2.06 if i < 7 else 4.54
    rect(s, x, y, cw, 2.30, WHITE, line=MIST_D, lw=1.0, rounded=True, adj=0.08)
    rect(s, x + 0.10, y + 0.10, cw - 0.20, 1.26, MIST, rounded=True, adj=0.08)
    pic(s, img, x + 0.16, y + 0.14, cw - 0.32, 1.18)
    text(s, x + 0.12, y + 1.46, cw - 0.24, 0.32, nm, size=8.5, bold=True, color=NAVY,
         line=1.10)
    text(s, x + 0.12, y + 1.82, cw - 0.24, 0.22, g, size=8, color=GREY)
    text(s, x + 0.12, y + 2.02, cw - 0.24, 0.22, p + " грн", size=10.5, bold=True,
         color=ORANGE)

rect(s, 7.66, 4.54, 4.98, 2.30, NAVY, rounded=True, adj=0.07)
rect(s, 7.66, 4.54, 0.09, 2.30, ORANGE, rounded=True, adj=0.5)
text(s, 8.02, 4.78, 4.36, 0.30, "Розподіл цін лінійки", size=15, bold=True, color=AMBER)
for i, (lb, vl, sub) in enumerate([("250 г", "45 – 89,99 грн", "5 варіантів"),
                                   ("220 г", "89,99 – 179 грн", "6 варіантів")]):
    yy = 5.22 + i * 0.66
    text(s, 8.02, yy, 0.90, 0.30, lb, size=13, bold=True, color=WHITE)
    text(s, 8.96, yy + 0.02, 2.10, 0.28, vl, size=12, bold=True, color=WHITE)
    text(s, 11.12, yy + 0.04, 1.30, 0.26, sub, size=9.5, color=RGBColor(0x9F, 0xA9, 0xC4))
text(s, 8.02, 6.42, 4.36, 0.26, "Ціни — картки «Сільпо» на 22.09.2026",
     size=9, color=RGBColor(0x9F, 0xA9, 0xC4))
foot(s, "Джерело: S1 · каталог «Сільпо», 11 карток Ben's Original.")

# ══ 06 · УПАКОВКА · ГРАМАЖ ════════════════════════════════════════════════
s = slide()
header(s, "ЯКА УПАКОВКА · 1/2", "Пауч 220–250 г — основний формат",
       "14 порівнянних варіантів гарнірів; страви з наповнювачем рахуються окремо.")
qtag(s, M, 1.98, "2", "ЯКА УПАКОВКА")
BARS = [("220 г", 6, ORANGE), ("250 г", 6, ORANGE), ("210 г", 1, GREEN), ("240 г", 1, TEAL)]
text(s, M, 2.48, 4.60, 0.26, "РОЗПОДІЛ ЗА МАСОЮ, ВАРІАНТІВ", size=9.5, bold=True,
     color=GREY)
yy = 2.90
for lb, n, c in BARS:
    text(s, M, yy, 0.86, 0.34, lb, size=13, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
    rect(s, M + 0.96, yy + 0.05, 3.60, 0.24, MIST, rounded=True, adj=0.5)
    rect(s, M + 0.96, yy + 0.05, 0.60 * n, 0.24, c, rounded=True, adj=0.5)
    text(s, M + 4.70, yy, 0.50, 0.34, str(n), size=13, bold=True, color=c,
         anchor=MSO_ANCHOR.MIDDLE)
    yy += 0.60
rect(s, M, 5.42, 5.34, 1.26, NAVY, rounded=True, adj=0.08)
text(s, M + 0.34, 5.60, 4.70, 0.32, "13 паучів / 1 лоток-чаша", size=15.5, bold=True,
     color=AMBER)
text(s, M + 0.34, 5.98, 4.70, 0.56,
     "Одного лідера за грамажем немає: 220 і 250 г представлені порівну.",
     size=11, color=RGBColor(0xC6, 0xCE, 0xE2), line=1.28)

GAL = [("md/bens_basmati220.jpg", "220 г", "6 варіантів"),
       ("md/bens_basmati250.jpg", "250 г", "6 варіантів"),
       ("md/bibigo_bowl.jpg", "210 г", "1 варіант"),
       ("md/bens_bio240.jpg", "240 г", "1 варіант")]
for i, (img, cap, n) in enumerate(GAL):
    x = 6.50 + (i % 2) * 3.28
    y = 2.42 + (i // 2) * 2.26
    rect(s, x, y, 3.06, 2.06, MIST, rounded=True, adj=0.07)
    pic(s, img, x + 0.18, y + 0.12, 2.70, 1.44)
    text(s, x, y + 1.58, 3.06, 0.26, cap, size=12, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER)
    text(s, x, y + 1.80, 3.06, 0.22, n, size=9, color=GREY, align=PP_ALIGN.CENTER)
foot(s, "Джерела: S1, S2, S3, S5 · маса за картками товарів, 22.09.2026. "
        "Long Grain 220 / 250 г виключено зі статистики маси через суперечність картки.")

# ══ 07 · УПАКОВКА · ТИПИ ══════════════════════════════════════════════════
s = slide()
header(s, "ЯКА УПАКОВКА · 2/2", "Три типи паковання на вітрині",
       "Тип упаковки визначає і спосіб споживання, і полицю, на яку товар потрапляє.")
qtag(s, M, 1.98, "2", "ЯКА УПАКОВКА")
TYPES = [("md/bens_basmati250.jpg", "ПАУЧ", "220 · 240 · 250 г", "14 із 18 карток",
          ["Ben's Original", "Clearspring"], "Плаский реторт-пауч.",
          "Гарнір, що висипається у тарілку.", ORANGE),
         ("md/bibigo_bowl.jpg", "ЛОТОК-ЧАША", "210 г", "4 із 18 карток",
          ["Bibigo"], "Жорстка чаша з плівкою. Один товар\nу трьох каналах.",
          "Їдять просто з упаковки.", GREEN),
         ("md/ottogi_octopus.jpg", "ЧАША ЗІ СТРАВОЮ", "217 – 315 г", "6 карток окремо",
          ["Ottogi"], "Висока чаша, рис із наповнювачем.",
          "Самостійний обід, не гарнір.", PLUM)]
for i, (img, kind, mass, share, brands, pack, use, c) in enumerate(TYPES):
    x = M + i * 4.10
    rect(s, x, 2.44, 3.84, 4.24, WHITE, line=MIST_D, lw=1.2, rounded=True, adj=0.05)
    rect(s, x, 2.44, 3.84, 0.10, c, rounded=True, adj=0.5)
    rect(s, x + 0.20, 2.70, 3.44, 1.60, MIST, rounded=True, adj=0.05)
    pic(s, img, x + 0.34, 2.78, 3.16, 1.44)
    text(s, x + 0.32, 4.44, 3.20, 0.30, kind, size=14, bold=True, color=c)
    text(s, x + 0.32, 4.78, 3.20, 0.28, mass, size=13, bold=True, color=NAVY)
    rect(s, x + 0.32, 5.14, 3.20, 0.02, MIST_D)
    text(s, x + 0.32, 5.26, 3.20, 0.26, " · ".join(brands), size=10.5, bold=True,
         color=INK)
    text(s, x + 0.32, 5.56, 3.20, 0.30, pack, size=10, color=GREY, line=1.22)
    text(s, x + 0.32, 5.88, 3.20, 0.30, use, size=10, color=GREY, line=1.22)
    pill(s, x + 0.32, 6.26, share, c, size=8)
foot(s, "Джерела: S1, S2, S3, S5, S9, S11 · тип паковання за фото та описом карток, 22.09.2026.")

# ══ 08 · ЦІНА · ЗА УПАКОВКУ ═══════════════════════════════════════════════
s = slide()
header(s, "ЯКА ЦІНА · 1/2", "За упаковку: від 45 до 252 грн",
       "Ціна картки без залишку є орієнтиром каталогу, а не доступною пропозицією.")
qtag(s, M, 1.98, "3", "ЯКА ЦІНА")
ROWS = [("md/bens_longgrain.jpg", "Ben's Long Grain", "Сільпо · 250 г", 45, 18.00, False),
        ("md/bens_basmati250.jpg", "Ben's Basmati", "Сільпо · 250 г", 89.99, 36.00, False),
        ("md/bens_basmati220.jpg", "Ben's Basmati", "MAUDAU · 220 г", 100, 45.45, False),
        ("md/bibigo_bowl.jpg", "Bibigo білий рис", "Смак Кореї · 210 г", 135, 64.29, True),
        ("md/bens_bio240.jpg", "Ben's Bio Basmati", "Edison Lee · 240 г", 139, 57.92, None),
        ("md/bens_sweetchili.jpg", "Ben's Sweet Chili", "Сільпо · 220 г", 179, 81.36, False),
        ("md/bibigo_bowl2.jpg", "Bibigo білий рис", "Prom · 210 г", 189, 90.00, False),
        ("md/clearspring.jpg", "Clearspring Brown & Wild", "MAUDAU · 250 г", 252, 100.80, False)]
MAXP, BX, BW = 260.0, 5.40, 4.30
yy = 2.38
for img, nm, chan, price, per, ok in ROWS:
    rect(s, M, yy, 11.98, 0.54, MIST if ok else WHITE,
         line=None if ok else MIST_D, lw=1.0, rounded=True, adj=0.26)
    rect(s, M + 0.08, yy + 0.05, 0.56, 0.44, WHITE, rounded=True, adj=0.16)
    pic(s, img, M + 0.11, yy + 0.07, 0.50, 0.40)
    text(s, M + 0.80, yy + 0.04, 2.60, 0.24, nm, size=11, bold=True, color=NAVY)
    text(s, M + 0.80, yy + 0.28, 2.60, 0.22, chan, size=8.5, color=GREY)
    c = GREEN if ok else (GREY if ok is None else CORAL)
    rect(s, BX, yy + 0.20, BW, 0.14, MIST_D, rounded=True, adj=0.5)
    rect(s, BX, yy + 0.20, max(0.12, BW * price / MAXP), 0.14,
         ORANGE if ok else (NAVY_L if ok is None else RGBColor(0xC3, 0xCB, 0xDC)),
         rounded=True, adj=0.5)
    ptxt = f"{price:g}".replace(".", ",") + " грн"
    text(s, 10.00, yy + 0.08, 1.10, 0.36, ptxt, size=13, bold=True, color=NAVY,
         align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    text(s, 11.18, yy + 0.10, 1.04, 0.32, f"{per:.2f}".replace(".", ",") + " /100 г",
         size=9.5, bold=True, color=ORANGE, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    pill(s, 12.32, yy + 0.14, "Є" if ok else ("?" if ok is None else "НЕМАЄ"), c,
         w=0.66, size=7.5)
    yy += 0.565
text(s, M, 6.96, 11.98, 0.26,
     "Показано 8 із 18 карток — крайні точки та ключові канали; повний перелік — слайд 05. "
     "Онлайн-ціни без доставки.", size=8.5, color=GREY, line=1.16)

# ══ 09 · ЦІНА · ЗА 100 Г ══════════════════════════════════════════════════
s = slide()
header(s, "ЯКА ЦІНА · 2/2", f"За 100 грамів: від 18 до 101 грн",
       f"Медіана {PER_MED:.1f} грн за 100 г. Розкид у 5,6 раза — це різні сегменти, "
       f"а не різні ціни на один товар.".replace(".", ",", 1))
qtag(s, M, 1.98, "3", "ЯКА ЦІНА")

BR_COLOR = {"Ben's Original": ORANGE, "Bibigo": GREEN,
            "Clearspring": TEAL, "Ottogi": PLUM}
pts = sorted(((p / g * 100, b, v, ch, g, p) for b, v, ch, g, p, _ in CARDS))
AX_X, AX_W, AX_Y = M + 0.40, 11.10, 5.86
LO, HI = 0.0, 110.0


def ax(v):
    return AX_X + AX_W * (v - LO) / (HI - LO)


rect(s, M, 3.02, 11.98, 3.30, MIST, rounded=True, adj=0.05)
rect(s, AX_X, AX_Y, AX_W, 0.03, RGBColor(0xB0, 0xBA, 0xD2))
for t in range(0, 111, 10):
    rect(s, ax(t), AX_Y, 0.02, 0.13, RGBColor(0xB0, 0xBA, 0xD2))
    text(s, ax(t) - 0.32, AX_Y + 0.20, 0.64, 0.24, str(t), size=9.5, color=GREY,
         align=PP_ALIGN.CENTER)
text(s, AX_X, 3.20, 3.40, 0.24, "ГРН ЗА 100 ГРАМІВ", size=9, bold=True, color=GREY)

# бджолиний рій: крапки піднімаються від осі, підпис праворуч від крапки
lanes = [-9.0] * 10
for per, b, v, ch, g, p in pts:
    x = ax(per)
    lane = 0
    while lane < 9 and lanes[lane] > x - 0.62:
        lane += 1
    lanes[lane] = x
    y = AX_Y - 0.36 - lane * 0.34
    c = BR_COLOR.get(b, GREY)
    rect(s, x - 0.009, y + 0.16, 0.018, AX_Y - y - 0.16, RGBColor(0xCF, 0xD7, 0xE8))
    rect(s, x - 0.085, y, 0.17, 0.17, c, rounded=True, adj=0.5)
    text(s, x + 0.14, y - 0.02, 0.66, 0.22, f"{per:.0f}", size=9, bold=True, color=c)

# три цінові смуги — підказка, як читати рій
BANDS = [("до 45 грн", "великі паучі та акційні залишки", 5, ORANGE),
         ("45 – 70 грн", "основна маса лінійок Ben's і Bibigo", 9, TEAL),
         ("понад 80 грн", "преміум і націнка маркетплейсів", 4, PLUM)]
bx = M
for lb, sub, n, c in BANDS:
    rect(s, bx, 2.38, 3.90, 0.52, WHITE, line=MIST_D, lw=1.0, rounded=True, adj=0.22)
    rect(s, bx, 2.38, 0.08, 0.52, c, rounded=True, adj=0.5)
    text(s, bx + 0.24, 2.46, 1.30, 0.24, lb, size=11.5, bold=True, color=NAVY)
    text(s, bx + 0.24, 2.68, 3.40, 0.20, sub, size=8.5, color=GREY)
    text(s, bx + 3.00, 2.48, 0.72, 0.28, str(n), size=15, bold=True, color=c,
         align=PP_ALIGN.RIGHT)
    bx += 4.04

# легенда
lx = M + 0.40
for b, c in BR_COLOR.items():
    if b == "Ottogi":
        continue
    rect(s, lx, 6.52, 0.15, 0.15, c, rounded=True, adj=0.5)
    text(s, lx + 0.24, 6.46, 1.80, 0.26, b, size=10.5, bold=True, color=INK)
    lx += 2.10
text(s, 7.60, 6.44, 4.72, 0.30,
     f"n = {len(pts)} карток · мін {PER_MIN:.0f} · медіана {PER_MED:.1f} · макс {PER_MAX:.1f}"
     .replace(".", ","), size=10.5, bold=True, color=NAVY, align=PP_ALIGN.RIGHT)
foot(s, "Джерела: S1–S8 · розрахунок за масою та ціною карток, 22.09.2026. "
        "Страви Ottogi (81–117 грн/100 г) не входять до вибірки гарнірів — див. слайд 12.")

# ══ 10 · ДЕ · КАНАЛИ ══════════════════════════════════════════════════════
s = slide()
header(s, "ДЕ ПРЕДСТАВЛЕНО · 1/2", "Шість каналів: мережа, маркетплейси, спеціалізований онлайн",
       "Наявність на дату перевірки 22.09.2026; «Сільпо» перевірено без вибору адреси.")
qtag(s, M, 1.98, "4", "ДЕ ПРЕДСТАВЛЕНО")
CH = [("Сільпо", "Мережа супермаркетів", "11 SKU Ben's Original",
       "УСІ БЕЗ ЗАЛИШКУ", CORAL),
      ("MAUDAU", "Онлайн-супермаркет", "Ben's Basmati 220 г · Clearspring",
       "БЕЗ ЗАЛИШКУ", CORAL),
      ("Смак Кореї", "Спеціалізований онлайн", "Bibigo 210 г",
       "У НАЯВНОСТІ · 8 УП.", GREEN),
      ("Edison Lee", "Спеціалізований онлайн", "Ben's Bio 240 г · Long Grain 220 г",
       "СКЛАД НЕ ПІДТВЕРДЖЕНО", GREY),
      ("Rozetka · Prom", "Маркетплейси", "Bibigo · страви Ottogi · Portion",
       "КАРТКИ НЕДОСТУПНІ", CORAL),
      ("Pulsar", "Спеціалізований онлайн", "6 рисових страв Ottogi",
       "УСІ БЕЗ ЗАЛИШКУ", CORAL)]
for i, (nm, kind, what, st_, c) in enumerate(CH):
    x = M + (i % 3) * 4.05
    y = 2.44 + (i // 3) * 1.76
    rect(s, x, y, 3.78, 1.54, WHITE, line=MIST_D, lw=1.2, rounded=True, adj=0.10)
    rect(s, x, y, 3.78, 0.09, c, rounded=True, adj=0.5)
    text(s, x + 0.28, y + 0.24, 3.20, 0.32, nm, size=15, bold=True, color=NAVY)
    text(s, x + 0.28, y + 0.58, 3.20, 0.24, kind.upper(), size=8, bold=True, color=c)
    text(s, x + 0.28, y + 0.84, 3.20, 0.34, what, size=10, color=GREY, line=1.18)
    pill(s, x + 0.28, y + 1.18, st_, c, size=8)
rect(s, M, 6.12, 11.98, 0.72, MIST, rounded=True, adj=0.18)
rect(s, M, 6.12, 0.09, 0.72, ORANGE, rounded=True, adj=0.5)
text(s, M + 0.34, 6.26, 11.30, 0.46,
     [[("Не знайдено пошуком: ", {"bold": True, "color": NAVY}),
       ("АТБ · Novus · Metro · Varus · Auchan · Fozzy · Таврія В. Це не доказ "
        "відсутності в мережі — перевірка велася онлайн, без обходу полиці.",
        {"color": GREY})]], size=11)
foot(s, "Джерела: S1–S10 · зріз 22.09.2026.")

# ══ 11 · ДЕ · НАЯВНІСТЬ ═══════════════════════════════════════════════════
s = slide()
header(s, "ДЕ ПРЕДСТАВЛЕНО · 2/2", "Асортимент є на вітрині, товару на складі — майже немає",
       "Із 18 перевірених карток гарнірів залишок підтверджено лише в однієї.")
qtag(s, M, 1.98, "4", "ДЕ ПРЕДСТАВЛЕНО")
for i, (v, u, lb, sub, c) in enumerate([
        ("1", "У НАЯВНОСТІ", "Bibigo 210 г · «Смак Кореї»", "8 упаковок на 22.09.2026", GREEN),
        ("2", "НЕ ПІДТВЕРДЖЕНО", "Edison Lee", "склад продавця не підтверджено", GREY),
        ("15", "БЕЗ ЗАЛИШКУ", "решта карток гарнірів", "ціна лишається в каталозі", CORAL),
        ("6", "БЕЗ ЗАЛИШКУ", "страви Ottogi у Pulsar", "усі шість карток", PLUM)]):
    kpi(s, M + i * 3.07, 2.44, 2.82, v, u, lb, sub, c)

rect(s, M, 5.14, 11.98, 1.54, NAVY, rounded=True, adj=0.07)
rect(s, M, 5.14, 0.10, 1.54, ORANGE, rounded=True, adj=0.5)
text(s, M + 0.42, 5.36, 5.40, 0.32, "Як читати цю картину", size=15, bold=True, color=AMBER)
text(s, M + 0.42, 5.76, 5.40, 0.76,
     "Каталоги мереж і маркетплейсів категорію знають і тримають. "
     "Товарного покриття за цими картками немає.",
     size=11.5, color=WHITE, line=1.30)
rect(s, 7.30, 5.40, 0.04, 1.02, NAVY_L)
text(s, 7.64, 5.36, 4.70, 1.16,
     "Причина відсутності залишку — постачання чи обіговість — із відкритих "
     "даних не встановлюється. Це питання до самих мереж і дистриб'юторів.",
     size=11, color=RGBColor(0xB9, 0xC2, 0xDA), line=1.30)
foot(s, "Джерела: S1–S11 · статус наявності за картками на 22.09.2026.")

# ══ 12 · ДОДАТОК · OTTOGI ═════════════════════════════════════════════════
s = slide()
header(s, "ДОДАТОК · СУМІЖНИЙ СЕГМЕНТ", "Страви Ottogi: єдина ціна 255 грн на шість позицій",
       "Рис із наповнювачем — окремий сегмент; у статистику гарнірів не входить.")
OT = [("md/ottogi_chicken.jpg", "Гостра курка", "310/315* г", 255, 82.3),
      ("md/ottogi_hamburg.jpg", "Гамбурзький стейк", "315 г", 255, 81.0),
      ("md/ottogi_octopus.jpg", "Гострий восьминіг", "280 г", 255, 91.1),
      ("md/ottogi_jjampong.jpg", "Jin Jjampong", "217,5 г", 255, 117.2),
      ("md/ottogi_tuna.jpg", "Тунець і майонез", "247 г", 255, 103.2),
      ("md/ottogi_kimchi.jpg", "Кімчі й тунець", "310 г", 255, 82.3)]
for i, (img, nm, g, p, per) in enumerate(OT):
    x = M + (i % 3) * 4.10
    y = 2.20 + (i // 3) * 2.32
    rect(s, x, y, 3.84, 2.10, WHITE, line=MIST_D, lw=1.2, rounded=True, adj=0.07)
    rect(s, x + 0.14, y + 0.12, 1.80, 1.86, MIST, rounded=True, adj=0.07)
    pic(s, img, x + 0.20, y + 0.18, 1.68, 1.74)
    text(s, x + 2.08, y + 0.30, 1.62, 0.40, nm, size=11.5, bold=True, color=NAVY,
         line=1.12)
    text(s, x + 2.08, y + 0.78, 1.62, 0.24, g, size=9.5, color=GREY)
    text(s, x + 2.08, y + 1.04, 1.62, 0.32, "255 грн", size=15, bold=True, color=PLUM)
    text(s, x + 2.08, y + 1.38, 1.62, 0.24,
         f"{per:.1f}".replace(".", ",") + " /100 г", size=9.5, bold=True, color=ORANGE)
    pill(s, x + 2.08, y + 1.66, "НЕМАЄ", CORAL, w=0.86, size=8)
text(s, M, 6.82, 11.98, 0.30,
     "* Яловичина / бульгогі та свинина 269 / 310 г на Rozetka: ідентичність SKU потребує "
     "звірки EAN та етикетки. Повторні пропозиції продавців не рахуються окремими SKU.",
     size=8.5, color=GREY, line=1.16)
foot(s, "Джерела: O01–O06, S10, R01–R12 · картки Pulsar і продавців Rozetka, 22.09.2026.")

# ══ 13 · МЕТОДОЛОГІЯ ══════════════════════════════════════════════════════
s = slide()
header(s, "МЕТОДОЛОГІЯ", "Основа дослідження та межі висновків",
       "Відкриті українські онлайн-вітрини, зріз на 22 вересня 2026 року.")
COLS = [("ЩО ПЕРЕВІРЕНО", ORANGE,
         ["Сільпо · MAUDAU · Смак Кореї", "Edison Lee · Rozetka · Prom · Pulsar",
          "18 карток гарнірів", "6 карток страв Ottogi",
          "Маса, ціна, канал, наявність"]),
        ("ЩО НЕ ПЕРЕВІРЕНО", CORAL,
         ["АТБ · Novus · Metro · Varus", "Auchan · Fozzy · Таврія В",
          "Офлайн-полиця в жодній мережі", "Фактичні продажі та обіговість",
          "Причина відсутності залишку"]),
        ("МЕЖІ ВИСНОВКІВ", TEAL,
         ["Частота SKU ≠ частка продажів", "Ціна без залишку — орієнтир каталогу",
          "Відсутність у пошуку ≠ відсутність у мережі",
          "Long Grain виключено зі статистики маси",
          "Доля ринку жодного бренду не встановлена"])]
for i, (title, c, items) in enumerate(COLS):
    x = M + i * 4.10
    rect(s, x, 2.06, 3.84, 4.10, WHITE, line=MIST_D, lw=1.2, rounded=True, adj=0.06)
    rect(s, x, 2.06, 3.84, 0.10, c, rounded=True, adj=0.5)
    text(s, x + 0.30, 2.34, 3.24, 0.30, title, size=11, bold=True, color=c)
    yy = 2.78
    for it in items:
        rect(s, x + 0.30, yy + 0.09, 0.09, 0.09, c, rounded=True, adj=0.5)
        text(s, x + 0.54, yy, 3.00, 0.52, it, size=10.5, color=INK, line=1.24)
        yy += 0.62
rect(s, M, 6.30, 11.98, 0.54, MIST, rounded=True, adj=0.22)
text(s, M + 0.34, 6.42, 11.30, 0.32,
     "Точні посилання — у примітках відповідних слайдів. Полична ціна власного "
     "продукту в це дослідження не входить і рахується окремо.",
     size=10.5, color=GREY)
foot(s, "Morskyi Dim · дослідження ринку · вересень 2026.")

prs.save("Gotovyi_Rys_Doslidzhennia_Rynku.pptx")
print("saved Gotovyi_Rys_Doslidzhennia_Rynku.pptx ·", len(prs.slides._sldIdLst), "slides")
