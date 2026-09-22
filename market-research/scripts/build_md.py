"""Готовий рис · Україна — колода в фірмовому стилі Morskyi Dim.

Ідентичність узято з оглядової колоди Nissin: темно-синій #303A5D, помаранчевий
#F9A50B, хвильовий патерн і логотип. Далі — яскравіше: жодної щільної таблиці,
кожне число має пакшот або графіку.
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from PIL import Image

W, H = 13.3333, 7.5
M = 0.66
HDR = 1.16                                  # висота фірмової шапки

NAVY   = RGBColor(0x30, 0x3A, 0x5D)
NAVY_D = RGBColor(0x23, 0x2B, 0x47)
NAVY_L = RGBColor(0x46, 0x53, 0x82)
ORANGE = RGBColor(0xF9, 0xA5, 0x0B)
AMBER  = RGBColor(0xFF, 0xC9, 0x5C)
TEAL   = RGBColor(0x1E, 0x9E, 0xA6)
CORAL  = RGBColor(0xE2, 0x56, 0x4B)
GREEN  = RGBColor(0x37, 0xA1, 0x69)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
PAPER  = RGBColor(0xFF, 0xFF, 0xFF)
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
    """Продукт цілком, без кадрування, по центру рамки."""
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
    rect(s, 0, 0, W, H, NAVY if dark else PAPER)
    PG["n"] += 1
    return s


def header(s, eyebrow, title, dek=None):
    """Фірмова шапка: синя смуга, логотип, помаранчева лінія."""
    rect(s, 0, 0, W, HDR, NAVY)
    rect(s, 0, HDR, W, 0.055, ORANGE)
    s.shapes.add_picture(LOGO, Inches(M), Inches(0.26), Inches(1.52), Inches(0.66))
    text(s, 2.52, 0.34, 6.60, 0.24, eyebrow, size=9.5, bold=True, color=AMBER)
    text(s, 2.52, 0.60, 8.40, 0.40, title, size=19, bold=True, color=WHITE)
    text(s, W - M - 0.70, 0.44, 0.70, 0.36, f"{PG['n']:02d}", size=17, bold=True,
         color=NAVY_L, align=PP_ALIGN.RIGHT)
    if dek:
        text(s, M, HDR + 0.30, 11.9, 0.34, dek, size=12.5, color=GREY, line=1.20)


def foot(s, src):
    text(s, M, 7.04, 11.9, 0.26, src, size=8, color=GREY, line=1.14)


def pill(s, x, y, label, fg=WHITE, bg=GREEN, w=None, size=8.5):
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


def prodcard(s, x, y, w, h, img, title, sub, price, note, tone=NAVY, ok=True):
    rect(s, x, y, w, h, WHITE, line=MIST_D, lw=1.2, rounded=True, adj=0.05)
    rect(s, x, y, w, 0.09, tone, rounded=True, adj=0.5)
    ph = h * 0.46
    rect(s, x + 0.12, y + 0.22, w - 0.24, ph, MIST, rounded=True, adj=0.05)
    pic(s, img, x + 0.22, y + 0.30, w - 0.44, ph - 0.16)
    yy = y + ph + 0.40
    text(s, x + 0.24, yy, w - 0.48, 0.24, title, size=8.5, bold=True, color=ORANGE)
    text(s, x + 0.24, yy + 0.24, w - 0.48, 0.46, sub, size=12, bold=True, color=INK, line=1.10)
    text(s, x + 0.24, yy + 0.76, w - 0.48, 0.30, price, size=14, bold=True, color=NAVY)
    pill(s, x + 0.24, yy + 1.12, note, WHITE, GREEN if ok else CORAL)


# ══ 01 · ОБКЛАДИНКА ═══════════════════════════════════════════════════════
s = slide(dark=True)
cover_pic(s, BG_TITLE, 0, 0, W, H)
s.shapes.add_picture(LOGO, Inches(M), Inches(0.52), Inches(1.94), Inches(0.84))
rect(s, M, 2.18, 0.54, 0.09, ORANGE)
text(s, M, 2.52, 7.00, 0.34, "МАРКЕТИНГОВЕ ДОСЛІДЖЕННЯ · ВЕРЕСЕНЬ 2026",
     size=11.5, bold=True, color=AMBER)
text(s, M, 2.98, 7.20, 1.70, "Ринок готового рису\nв Україні",
     size=44, bold=True, color=WHITE, line=1.04)
text(s, M, 4.66, 6.60, 0.40, "Конкурентне середовище, асортимент, ціни та стратегія виходу",
     size=15, color=RGBColor(0xC6, 0xCE, 0xE2))
rect(s, M, 5.36, 5.90, 0.055, RGBColor(0x55, 0x60, 0x8C))
text(s, M, 5.58, 6.60, 0.60,
     "Для власника компанії та керівників розвитку і комерції\n"
     "Зріз онлайн-вітрин: 22 вересня 2026 р.",
     size=11, color=RGBColor(0x9F, 0xA9, 0xC4), line=1.36)

for i, (im, x, y, w, h) in enumerate([
        ("md/bens_basmati250.jpg", 8.26, 1.22, 2.35, 2.62),
        ("md/bibigo_bowl.jpg",    10.78, 1.60, 2.35, 2.24),
        ("md/clearspring.jpg",     8.26, 3.98, 2.35, 2.10),
        ("md/ottogi_burger.jpg",  10.78, 3.98, 2.35, 2.10)]):
    rect(s, x, y, w, h, WHITE, rounded=True, adj=0.06)
    pic(s, im, x + 0.16, y + 0.14, w - 0.32, h - 0.28)
rect(s, 8.26, 6.22, 4.87, 0.86, RGBColor(0x26, 0x2F, 0x4D), rounded=True, adj=0.14)
text(s, 8.58, 6.34, 4.30, 0.64,
     [[("3 бренди гарнірів · 14 порівнянних варіантів · ", {"bold": True, "color": WHITE}),
       ("ціни 100–252 грн. Наявність — головна проблема категорії.",
        {"color": RGBColor(0xB9, 0xC2, 0xDA)})]], size=11, line=1.28)

# ══ 02 · РЕЗЮМЕ ═══════════════════════════════════════════════════════════
s = slide()
header(s, "РЕЗЮМЕ ДЛЯ ВЛАСНИКА", "Ринок існує. Масовий попит ще не доведено",
       "Категорія знайома рітейлу, але майже кожна перевірена картка — без залишку.")
for i, (v, u, lb, sub, c) in enumerate([
        ("3", "БРЕНДИ ГАРНІРІВ", "Ben's Original · Bibigo\nClearspring",
         "плюс Ottogi та Portion у стравах", ORANGE),
        ("220–250", "ГРАМІВ", "основний діапазон",
         "12 із 14 порівнянних варіантів", TEAL),
        ("135", "ГРИВЕНЬ", "діюча пропозиція",
         "Bibigo 210 г · «Смак Кореї»", GREEN),
        ("1 з 14", "У НАЯВНОСТІ", "решта — без залишку",
         "наявність, а не ціна, стримує категорію", CORAL)]):
    kpi(s, M + i * 3.07, 2.06, 2.82, v, u, lb, sub, c)

rect(s, M, 4.76, 11.98, 1.92, NAVY, rounded=True, adj=0.07)
rect(s, M, 4.76, 0.10, 1.92, ORANGE, rounded=True, adj=0.5)
text(s, M + 0.42, 5.02, 6.30, 0.40, "Комерційний висновок", size=16, bold=True, color=AMBER)
text(s, M + 0.42, 5.52, 6.30, 0.92,
     "Вхід через базовий гарнір у пауче з обмеженою стартовою матрицею: "
     "2 SKU, 220–250 г, ціновий тест 109–129 грн.",
     size=12.5, color=WHITE, line=1.30)
rect(s, 7.60, 5.06, 0.04, 1.32, NAVY_L)
text(s, 7.94, 5.02, 4.36, 1.36,
     "Основа аналізу — онлайн-асортимент. Частота SKU не дорівнює частці "
     "продажів; позиції без залишку позначені окремо на кожному слайді.",
     size=11, color=RGBColor(0xB9, 0xC2, 0xDA), line=1.32)
foot(s, "Джерела: S1, S2, S3, S5 · перевірені картки українських онлайн-вітрин, 22.09.2026.")

# ══ 03 · ТРИ СЕГМЕНТИ ═════════════════════════════════════════════════════
s = slide()
header(s, "КОНКУРЕНТНЕ СЕРЕДОВИЩЕ", "Три сегменти. Наш — перший",
       "Вологий готовий гарнір без холодильника — цільовий сегмент.")
SEG = [("md/bens_basmati250.jpg", "ПОВСЯКДЕННИЙ ГАРНІР", "Ben's Original · Bibigo",
        "Чистий рис або легкий смак.\nПауч 220–250 г, лоток 210 г.",
        "ЦІЛЬОВИЙ СЕГМЕНТ", ORANGE, True),
       ("md/clearspring.jpg", "ОРГАНІЧНИЙ ПРЕМІУМ", "Clearspring",
        "Суміш рисів із tamari.\nПауч 250 г; картка 252 грн.",
        "ВУЗЬКА НІША", TEAL, False),
       ("md/ottogi_burger.jpg", "ПОВНОЦІННА СТРАВА", "Ottogi · Portion",
        "Рис із наповнювачем.\nЧаша або реторт-пакет.",
        "ІНШИЙ ПОВІД", NAVY_L, False)]
for i, (img, kind, brands, desc, tag, c, hero) in enumerate(SEG):
    x = M + i * 4.10
    rect(s, x, 2.06, 3.84, 4.62, MIST if hero else WHITE,
         line=None if hero else MIST_D, lw=1.2, rounded=True, adj=0.05)
    rect(s, x, 2.06, 3.84, 0.10, c, rounded=True, adj=0.5)
    rect(s, x + 0.20, 2.34, 3.44, 1.94, WHITE, rounded=True, adj=0.05)
    pic(s, img, x + 0.34, 2.44, 3.16, 1.74)
    text(s, x + 0.32, 4.44, 3.20, 0.24, kind, size=9, bold=True, color=c)
    text(s, x + 0.32, 4.72, 3.20, 0.34, brands, size=15, bold=True, color=NAVY)
    text(s, x + 0.32, 5.16, 3.20, 0.72, desc, size=11, color=GREY, line=1.30)
    pill(s, x + 0.32, 6.12, tag, WHITE, c)
foot(s, "Джерела: S1, S2, S3, S9, S11 · фото — картки товарів відповідних магазинів.")

# ══ 04 · ГРАВЦІ ═══════════════════════════════════════════════════════════
s = slide()
header(s, "ГРАВЦІ КАТЕГОРІЇ", "П'ять компаній формують усю пропозицію",
       "Найширша лінійка гарнірів — у Ben's Original. Частки продажів не встановлені.")
PL = [("md/bens_lang220.jpg", "Mars · Ben's Original", "11 SKU у «Сільпо» + 2 позиції",
       "Сільпо · MAUDAU · Edison Lee", 11, ORANGE),
      ("md/bibigo_bowl2.jpg", "CJ · Bibigo", "Білий рис 210 г",
       "Смак Кореї · Prom · Rozetka", 1, GREEN),
      ("md/clearspring.jpg", "Clearspring", "Brown & Wild Rice з tamari",
       "MAUDAU", 1, TEAL),
      ("md/ottogi_burger.jpg", "Ottogi", "Рисові страви: тунець, курка, кімчі",
       "Pulsar · продавці Rozetka", 6, NAVY_L)]
for i, (img, brand, assort, chan, n, c) in enumerate(PL):
    y = 2.04 + i * 1.18
    rect(s, M, y, 11.98, 1.04, MIST if i % 2 == 0 else WHITE, rounded=True, adj=0.14)
    rect(s, M, y, 0.09, 1.04, c, rounded=True, adj=0.5)
    rect(s, M + 0.24, y + 0.10, 1.20, 0.84, WHITE, rounded=True, adj=0.10)
    pic(s, img, M + 0.30, y + 0.14, 1.08, 0.76)
    text(s, M + 1.62, y + 0.20, 3.10, 0.30, brand, size=14, bold=True, color=NAVY)
    text(s, M + 1.62, y + 0.56, 3.10, 0.28, assort, size=10, color=GREY)
    text(s, M + 5.00, y + 0.36, 3.30, 0.30, chan, size=11, color=INK)
    rect(s, M + 8.90, y + 0.22, 0.86, 0.60, c, rounded=True, adj=0.16)
    text(s, M + 8.90, y + 0.22, 0.86, 0.60, str(n), size=20, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, M + 9.86, y + 0.36, 1.90, 0.32, "SKU у\nвибірці", size=8.5, color=GREY, line=1.14)
text(s, M, 6.86, 11.98, 0.26,
     "Portion / «Пирятинський делікатес» — рис із куркою та овочами, 350 г, Rozetka. "
     "Edison Lee: наявність не підтверджено; у Long Grain суперечність маси 220 / 250 г.",
     size=9, color=GREY, line=1.16)
foot(s, "Джерела: S1, S2, S3, S9, S11, S18 · зріз 22.09.2026.")

# ══ 05 · ДИСТРИБУЦІЯ ══════════════════════════════════════════════════════
s = slide()
header(s, "ДИСТРИБУЦІЯ", "Асортимент є. Залишку майже немає",
       "Наявність на дату перевірки 22.09.2026; «Сільпо» перевірено без вибору адреси.")
CH = [("Сільпо", "11 SKU Ben's Original", "УСІ БЕЗ ЗАЛИШКУ", CORAL),
      ("MAUDAU", "Ben's Basmati 220 г · Clearspring", "БЕЗ ЗАЛИШКУ", CORAL),
      ("Смак Кореї", "Bibigo 210 г", "У НАЯВНОСТІ · 8 УП.", GREEN),
      ("Edison Lee", "Ben's Bio Basmati 240 г · Long Grain", "СКЛАД НЕ ПІДТВЕРДЖЕНО", GREY),
      ("Prom / Rozetka", "Bibigo · страви Ottogi", "КАРТКИ НЕДОСТУПНІ", CORAL),
      ("Pulsar", "6 рисових страв Ottogi", "УСІ БЕЗ ЗАЛИШКУ", CORAL)]
for i, (nm, what, st, c) in enumerate(CH):
    x = M + (i % 3) * 4.05
    y = 2.10 + (i // 3) * 1.66
    rect(s, x, y, 3.78, 1.44, WHITE, line=MIST_D, lw=1.2, rounded=True, adj=0.10)
    rect(s, x, y, 3.78, 0.09, c, rounded=True, adj=0.5)
    text(s, x + 0.28, y + 0.26, 3.20, 0.32, nm, size=15, bold=True, color=NAVY)
    text(s, x + 0.28, y + 0.62, 3.20, 0.34, what, size=10, color=GREY, line=1.18)
    pill(s, x + 0.28, y + 1.02, st, WHITE, c, size=8)

rect(s, M, 5.52, 11.98, 1.16, MIST, rounded=True, adj=0.10)
rect(s, M, 5.52, 0.09, 1.16, ORANGE, rounded=True, adj=0.5)
text(s, M + 0.34, 5.72, 11.30, 0.80,
     [[("АТБ, Novus, Metro, Varus, Auchan, Fozzy, Таврія В — ", {"bold": True, "color": NAVY}),
       ("порівнянний SKU не підтверджено пошуком. Це не доказ відсутності в мережі: "
        "перевірка велася онлайн, без обходу полиці.", {"color": GREY})]],
     size=11.5, line=1.30)
foot(s, "Джерела: S1, S2, S3, S4, S5, S7, S8, S9, S10 · зріз 22.09.2026.")

# ══ 06 · ФОРМАТ ═══════════════════════════════════════════════════════════
s = slide()
header(s, "ФОРМАТ І ГРАМАЖ", "Пауч 220–250 г — основний формат",
       "14 порівнянних варіантів гарнірів; страви з м'ясом та імпорт під замовлення виключені.")
BARS = [("220 г", 6, ORANGE), ("250 г", 6, ORANGE), ("210 г", 1, GREEN), ("240 г", 1, TEAL)]
text(s, M, 2.16, 4.60, 0.30, "РОЗПОДІЛ ЗА МАСОЮ, ВАРІАНТІВ", size=9.5, bold=True, color=GREY)
yy = 2.60
for lb, n, c in BARS:
    text(s, M, yy, 0.86, 0.34, lb, size=13, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
    rect(s, M + 0.96, yy + 0.05, 3.60, 0.24, MIST, rounded=True, adj=0.5)
    rect(s, M + 0.96, yy + 0.05, 0.60 * n, 0.24, c, rounded=True, adj=0.5)
    text(s, M + 4.70, yy, 0.50, 0.34, str(n), size=13, bold=True, color=c,
         anchor=MSO_ANCHOR.MIDDLE)
    yy += 0.60
rect(s, M, 5.16, 5.34, 1.52, NAVY, rounded=True, adj=0.08)
text(s, M + 0.34, 5.36, 4.70, 0.34, "13 паучів / 1 лоток", size=17, bold=True, color=AMBER)
text(s, M + 0.34, 5.78, 4.70, 0.76,
     "Одного лідера за грамажем немає. Частота асортименту не визначає "
     "найпродаванішу упаковку.", size=11, color=RGBColor(0xC6, 0xCE, 0xE2), line=1.28)

GAL = [("md/bens_basmati220.jpg", "220 г"), ("md/bens_basmati250.jpg", "250 г"),
       ("md/bibigo_bowl.jpg", "210 г"), ("md/bens_bio240.jpg", "240 г")]
for i, (img, cap) in enumerate(GAL):
    x = 6.50 + (i % 2) * 3.28
    y = 2.10 + (i // 2) * 2.34
    rect(s, x, y, 3.06, 2.12, MIST, rounded=True, adj=0.07)
    pic(s, img, x + 0.18, y + 0.14, 2.70, 1.60)
    text(s, x, y + 1.76, 3.06, 0.26, cap, size=11.5, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER)
foot(s, "Джерела: S1, S2, S3, S5 · маса за картками товарів, 22.09.2026.")

# ══ 07 · ЦІНИ ═════════════════════════════════════════════════════════════
s = slide()
header(s, "ЦІНОВИЙ ЛАНДШАФТ", "Від 100 до 252 грн у ключових картках",
       "Ціна відсутнього товару — орієнтир каталогу, а не доступна пропозиція.")
ROWS = [("md/bens_basmati220.jpg", "Ben's Basmati", "MAUDAU · 220 г", 100, 45.45, False),
        ("md/bibigo_bowl.jpg", "Bibigo білий рис", "Смак Кореї · 210 г", 135, 64.29, True),
        ("md/bens_bio240.jpg", "Ben's Bio Basmati", "Edison Lee · 240 г", 139, 57.92, None),
        ("md/bens_basmati250.jpg", "Ben's Basmati", "Сільпо · 220 г", 144, 65.45, False),
        ("md/bibigo_bowl2.jpg", "Bibigo білий рис", "Rozetka · 210 г", 184, 87.62, False),
        ("md/clearspring.jpg", "Clearspring Brown & Wild", "MAUDAU · 250 г", 252, 100.80, False)]
MAXP, BX, BW = 260.0, 5.40, 4.40
yy = 2.02
for img, nm, chan, price, per, ok in ROWS:
    rect(s, M, yy, 11.98, 0.72, MIST if ok else WHITE,
         line=None if ok else MIST_D, lw=1.0, rounded=True, adj=0.20)
    rect(s, M + 0.10, yy + 0.06, 0.72, 0.60, WHITE, rounded=True, adj=0.16)
    pic(s, img, M + 0.14, yy + 0.09, 0.64, 0.54)
    text(s, M + 0.98, yy + 0.10, 2.40, 0.28, nm, size=11.5, bold=True, color=NAVY)
    text(s, M + 0.98, yy + 0.38, 2.40, 0.24, chan, size=9, color=GREY)
    c = GREEN if ok else (GREY if ok is None else CORAL)
    rect(s, BX, yy + 0.29, BW, 0.16, MIST_D, rounded=True, adj=0.5)
    rect(s, BX, yy + 0.29, max(0.14, BW * price / MAXP), 0.16,
         ORANGE if ok else (NAVY_L if ok is None else RGBColor(0xC3, 0xCB, 0xDC)),
         rounded=True, adj=0.5)
    text(s, 10.12, yy + 0.16, 1.02, 0.40, f"{price} грн", size=14, bold=True,
         color=NAVY, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    text(s, 11.26, yy + 0.18, 1.00, 0.36, f"{per:.2f}".replace(".", ",") + " /100 г",
         size=10, bold=True, color=ORANGE, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    lbl = "Є" if ok else ("?" if ok is None else "НЕМАЄ")
    pill(s, 12.30, yy + 0.22, lbl, WHITE, c, w=0.70, size=8)
    yy += 0.76
text(s, M, 6.62, 11.98, 0.26,
     "Онлайн-ціни без доставки. У повному асортименті «Сільпо» трапляється 45 грн "
     "за відсутній Long Grain 250 г — це залишок каталогу, а не пропозиція.",
     size=9, color=GREY, line=1.16)
foot(s, "Джерела: S1, S2, S3, S4, S5, S7, S8 · зріз 22.09.2026.")

# ══ 08 · BEN'S ORIGINAL ═══════════════════════════════════════════════════
s = slide()
header(s, "ГЛИБИНА АСОРТИМЕНТУ", "Ben's Original: 11 варіантів — і жодного залишку",
       "Категорія знайома рітейлу. Причина відсутності — постачання чи обіговість — не встановлена.")
WALL = [("md/bens_basmati250.jpg", "Basmati", "250 г", "89,99"),
        ("md/bens_mediterran.jpg", "Mediterran", "250 г", "89,99"),
        ("md/bens_longgrain.jpg", "Long Grain", "250 г", "45"),
        ("md/bens_risibisi.jpg", "Risi Bisi", "250 г", "89,99"),
        ("md/bens_curry_indien.jpg", "Indian Curry", "250 г", "89,99"),
        ("md/bens_curry_linsen.jpg", "Curry · Lentils", "220 г", "89,99"),
        ("md/bens_sweetchili.jpg", "Sweet Chili", "220 г", "179"),
        ("md/bens_mexikanisch.jpg", "Mexikanisch", "220 г", "109"),
        ("md/bens_curryreis.jpg", "Curryreis Indien", "220 г", "144"),
        ("md/bens_basmati220.jpg", "Basmati", "220 г", "144"),
        ("md/bens_stickybowl.jpg", "Для боулів", "220 г", "149")]
cw, gap = 1.60, 0.14
for i, (img, nm, g, p) in enumerate(WALL):
    x = M + i * (cw + gap)
    if i >= 7:
        x = M + (i - 7) * (cw + gap)
    y = 2.06 if i < 7 else 4.54
    rect(s, x, y, cw, 2.30, WHITE, line=MIST_D, lw=1.0, rounded=True, adj=0.08)
    rect(s, x + 0.10, y + 0.10, cw - 0.20, 1.26, MIST, rounded=True, adj=0.08)
    pic(s, img, x + 0.16, y + 0.14, cw - 0.32, 1.18)
    text(s, x + 0.12, y + 1.46, cw - 0.24, 0.32, nm, size=8.5, bold=True, color=NAVY, line=1.10)
    text(s, x + 0.12, y + 1.82, cw - 0.24, 0.22, g, size=8, color=GREY)
    text(s, x + 0.12, y + 2.02, cw - 0.24, 0.22, p + " грн", size=10.5, bold=True, color=ORANGE)

rect(s, 7.66, 4.54, 4.98, 2.30, NAVY, rounded=True, adj=0.07)
rect(s, 7.66, 4.54, 0.09, 2.30, ORANGE, rounded=True, adj=0.5)
text(s, 8.02, 4.78, 4.36, 0.34, "Що це означає", size=16, bold=True, color=AMBER)
text(s, 8.02, 5.22, 4.36, 1.36,
     "Одинадцять варіантів у каталозі — це визнання категорії мережею. "
     "Нуль залишку — це або збій постачання, або низька обіговість. "
     "Перше відкриває нам двері, друге закриває: перевірити обидві версії "
     "треба до закупівлі.",
     size=11.5, color=WHITE, line=1.34)
text(s, 8.02, 6.44, 4.36, 0.26, "Ціни — картки «Сільпо» на 22.09.2026", size=9,
     color=RGBColor(0x9F, 0xA9, 0xC4))
foot(s, "Джерело: S1 · каталог «Сільпо», 11 карток Ben's Original, усі позначені як відсутні.")

# ══ 09 · ЦІНА ПОРЦІЇ ══════════════════════════════════════════════════════
s = slide()
header(s, "ЦІННІСТЬ ДЛЯ ПОКУПЦЯ", "Чистий гарнір конкурує з готовою стравою",
       "Для покупця важлива вартість обіду, а не тільки ціна 100 г рису.")
VAL = [("md/bibigo_bowl.jpg", "Bibigo", "Білий рис 210 г", "135 грн", "є",
        "Швидкий гарнір. Білок і соус\nпокупець додає сам.", GREEN),
       (None, "Portion", "Рис із куркою 350 г", "109 грн", "немає",
        "Готова страва більшої маси\nза меншу ціну.", CORAL),
       ("md/ottogi_octopus.jpg", "Ottogi", "Рис із наповнювачами", "255 грн", "немає",
        "Корейська страва в чаші.\nІнший повід споживання.", CORAL),
       ("md/clearspring.jpg", "Clearspring", "Суміш рисів 250 г", "252 грн", "немає",
        "Органічне преміальне\nпозиціювання.", CORAL)]
for i, (img, brand, nm, price, st, why, c) in enumerate(VAL):
    x = M + i * 3.07
    rect(s, x, 2.06, 2.82, 3.46, WHITE, line=MIST_D, lw=1.2, rounded=True, adj=0.06)
    rect(s, x + 0.14, 2.20, 2.54, 1.34, MIST, rounded=True, adj=0.07)
    if img:
        pic(s, img, x + 0.22, 2.26, 2.38, 1.22)
    else:
        text(s, x + 0.14, 2.68, 2.54, 0.24, "ФОТО КАРТКИ", size=9, bold=True,
             color=GREY, align=PP_ALIGN.CENTER)
        text(s, x + 0.14, 2.92, 2.54, 0.24, "недоступне", size=9, color=GREY,
             align=PP_ALIGN.CENTER)
    text(s, x + 0.26, 3.68, 2.30, 0.24, brand.upper(), size=8.5, bold=True, color=ORANGE)
    text(s, x + 0.26, 3.92, 2.30, 0.38, nm, size=11.5, bold=True, color=INK, line=1.12)
    text(s, x + 0.26, 4.36, 1.50, 0.34, price, size=16, bold=True, color=NAVY)
    pill(s, x + 1.86, 4.42, st, WHITE, c, w=0.70, size=8)
    text(s, x + 0.26, 4.84, 2.30, 0.56, why, size=9.5, color=GREY, line=1.24)

rect(s, M, 5.76, 11.98, 0.94, MIST, rounded=True, adj=0.14)
rect(s, M, 5.76, 0.09, 0.94, TEAL, rounded=True, adj=0.5)
text(s, M + 0.34, 5.94, 11.30, 0.60,
     [[("Висновок для продукту: ", {"bold": True, "color": NAVY}),
       ("продавати зручність і якість текстури. Порівняння ціни лише з корейським "
        "імпортом завищує привабливість запуску — поруч стоїть готова страва за 109 грн.",
        {"color": INK})]], size=11.5, line=1.28)
foot(s, "Джерела: S2, S3, S9, S11 · ціни карток на 22.09.2026.")

# ══ 10 · СТАРТОВА МАТРИЦЯ ═════════════════════════════════════════════════
s = slide()
header(s, "РЕКОМЕНДАЦІЯ", "Стартова матриця: два SKU на одній платформі",
       "Рекомендація для тесту, а не підтверджений прогноз продажів.")
MX = [("md/own_pouch_jasmine.jpg", "SKU 1 · БАЗА", "Базовий рис",
       "Пауч 220–250 г. Один сорт — jasmine або basmati за результатами дегустації. "
       "Основне завдання — швидкий гарнір.", ORANGE),
      ("md/own_pouch_mexican.jpg", "SKU 2 · РОЗШИРЕННЯ", "Рис із м'яким смаком",
       "Та сама маса й формат. Овочі або помірний curry. Розширює привід покупки "
       "без ускладнення лінійки.", TEAL)]
for i, (img, tag, nm, desc, c) in enumerate(MX):
    x = M + i * 4.32
    rect(s, x, 2.06, 4.06, 4.10, WHITE, line=MIST_D, lw=1.2, rounded=True, adj=0.06)
    rect(s, x, 2.06, 4.06, 0.10, c, rounded=True, adj=0.5)
    rect(s, x + 0.20, 2.32, 3.66, 1.56, MIST, rounded=True, adj=0.06)
    pic(s, img, x + 0.34, 2.40, 3.38, 1.40)
    pill(s, x + 0.30, 4.04, tag, WHITE, c)
    text(s, x + 0.30, 4.46, 3.46, 0.36, nm, size=16, bold=True, color=NAVY)
    text(s, x + 0.30, 4.92, 3.46, 1.06, desc, size=11, color=GREY, line=1.30)

rect(s, 9.34, 2.06, 3.30, 4.10, NAVY, rounded=True, adj=0.07)
text(s, 9.64, 2.30, 2.70, 0.30, "Параметри запуску", size=14, bold=True, color=AMBER)
yy = 2.80
for lb, vl in [("Ціна", "тест 109 / 119 / 129 грн"),
               ("Канали", "один мережевий партнер\nі спеціалізований онлайн"),
               ("Повідомлення", "«Готовий рис», час розігріву\nна лицьовій стороні")]:
    text(s, 9.64, yy, 2.70, 0.24, lb.upper(), size=8.5, bold=True, color=ORANGE)
    text(s, 9.64, yy + 0.26, 2.70, 0.62, vl, size=11, color=WHITE, line=1.28)
    yy += 1.06
rect(s, 9.64, 5.62, 2.70, 0.04, NAVY_L)
text(s, 9.64, 5.76, 2.70, 0.32, "Фото — формати наших постачальників",
     size=8.5, italic=True, color=RGBColor(0x9F, 0xA9, 0xC4), line=1.18)
foot(s, "Комерційна рекомендація / розрахунковий сценарій.")

# ══ 11 · ЕКОНОМІКА ════════════════════════════════════════════════════════
s = slide()
header(s, "ЕКОНОМІКА ОДНІЄЇ УПАКОВКИ", "119 грн: 12,86 грн на покриття витрат проєкту",
       "Розрахунковий сценарій для пауча. Результат не дорівнює чистому прибутку.")
SC = [(109, 53.14, 45.15, 7.99, RGBColor(0xC3, 0xCB, 0xDC)),
      (119, 58.01, 45.15, 12.86, ORANGE),
      (129, 62.89, 45.15, 17.74, RGBColor(0xC3, 0xCB, 0xDC))]
text(s, M, 2.10, 5.20, 0.26, "ПІСЛЯ ПДВ, МАРЖІ МЕРЕЖІ ТА УТРИМАНЬ, ГРН / УПАКОВКУ",
     size=9, bold=True, color=GREY)
BASE, SCALE = 6.30, 0.058
for i, (rrp, after, cost, contrib, c) in enumerate(SC):
    y = 2.52 + i * 1.42
    rect(s, M, y, 7.40, 1.22, MIST if c == ORANGE else WHITE,
         line=None if c == ORANGE else MIST_D, lw=1.0, rounded=True, adj=0.10)
    text(s, M + 0.26, y + 0.16, 1.20, 0.38, f"{rrp} грн", size=17, bold=True, color=NAVY)
    text(s, M + 0.26, y + 0.60, 1.20, 0.24, "роздріб", size=9, color=GREY)
    bx = M + 1.62
    rect(s, bx, y + 0.34, cost * SCALE, 0.44, NAVY_L, rounded=True, adj=0.28)
    rect(s, bx + cost * SCALE + 0.05, y + 0.34, contrib * SCALE + 0.22, 0.44, c,
         rounded=True, adj=0.28)
    text(s, bx + 0.14, y + 0.34, cost * SCALE, 0.44, "продукт + витрати 45,15",
         size=9, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    text(s, bx + cost * SCALE + 0.05, y + 0.34, contrib * SCALE + 0.22, 0.44,
         f"{contrib:.2f}".replace(".", ","), size=11, bold=True,
         color=WHITE if c == ORANGE else NAVY, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE)
    text(s, M + 6.10, y + 0.32, 1.10, 0.28, f"{after:.2f}".replace(".", ","),
         size=12, bold=True, color=NAVY, align=PP_ALIGN.RIGHT)
    text(s, M + 6.10, y + 0.62, 1.10, 0.24, "після утримань", size=8, color=GREY,
         align=PP_ALIGN.RIGHT)

rect(s, 8.44, 2.10, 4.20, 2.54, NAVY, rounded=True, adj=0.07)
text(s, 8.76, 2.34, 3.60, 0.30, "Ризик промо", size=14, bold=True, color=AMBER)
text(s, 8.76, 2.78, 3.60, 0.60, "Промо −20 % при базі 119 грн", size=11,
     color=RGBColor(0xC6, 0xCE, 0xE2), line=1.24)
text(s, 8.76, 3.40, 3.60, 0.60, "1,26 грн", size=30, bold=True, color=CORAL)
text(s, 8.76, 4.04, 3.60, 0.30, "вклад на покриття замість 12,86", size=10,
     color=RGBColor(0x9F, 0xA9, 0xC4))

rect(s, 8.44, 4.82, 4.20, 1.86, MIST, rounded=True, adj=0.08)
text(s, 8.76, 5.02, 3.60, 0.26, "ДОПУЩЕННЯ", size=9, bold=True, color=ORANGE)
text(s, 8.76, 5.32, 3.60, 1.22,
     "ПДВ 20 % · маржа мережі 35 % · утримання 10 % · місцеві витрати 5 грн.\n"
     "Вартість 40,15 грн з початкової моделі потребує оновлення та звірки ПДВ.",
     size=10, color=GREY, line=1.26)
foot(s, "Джерело: S0 · розрахунковий сценарій, не комерційна пропозиція.")

# ══ 12 · ОБСЯГ ТЕСТУ ══════════════════════════════════════════════════════
s = slide()
header(s, "ОБСЯГ ПЕРШОЇ ПАРТІЇ", "20 магазинів × 2 SKU × 8 тижнів",
       "Першу партію рахувати від числа магазинів і допустимого запасу, а не від прайсу.")
VOL = [("2", "шт. / SKU / магазин / тиждень", "640", "8 232 грн", RGBColor(0xC3, 0xCB, 0xDC)),
       ("5", "шт. / SKU / магазин / тиждень", "1 600", "20 580 грн", ORANGE),
       ("10", "шт. / SKU / магазин / тиждень", "3 200", "41 160 грн", TEAL)]
for i, (rate, unit, vol, contrib, c) in enumerate(VOL):
    x = M + i * 4.10
    hero = c == ORANGE
    rect(s, x, 2.10, 3.84, 3.60, MIST if hero else WHITE,
         line=None if hero else MIST_D, lw=1.2, rounded=True, adj=0.07)
    rect(s, x, 2.10, 3.84, 0.10, c, rounded=True, adj=0.5)
    text(s, x + 0.34, 2.40, 3.16, 0.80, rate, size=44, bold=True, color=NAVY)
    text(s, x + 0.34, 3.24, 3.16, 0.30, unit, size=9.5, color=GREY, line=1.18)
    rect(s, x + 0.34, 3.74, 3.16, 0.03, MIST_D)
    text(s, x + 0.34, 3.90, 1.60, 0.26, "ЗА 8 ТИЖНІВ", size=8.5, bold=True, color=GREY)
    text(s, x + 0.34, 4.16, 1.60, 0.38, vol, size=20, bold=True, color=NAVY)
    text(s, x + 0.34, 4.58, 1.60, 0.24, "упаковок", size=9, color=GREY)
    text(s, x + 2.00, 3.90, 1.50, 0.26, "ВКЛАД", size=8.5, bold=True, color=c,
         align=PP_ALIGN.RIGHT)
    text(s, x + 1.70, 4.16, 1.80, 0.38, contrib, size=16, bold=True, color=c,
         align=PP_ALIGN.RIGHT)
    text(s, x + 1.70, 4.58, 1.80, 0.24, "при 12,86 грн / шт.", size=8.5, color=GREY,
         align=PP_ALIGN.RIGHT)
    text(s, x + 0.34, 5.00, 3.16, 0.52,
         "консервативний сценарій" if i == 0 else
         ("робоча гіпотеза для розрахунку партії" if i == 1 else "оптимістичний сценарій"),
         size=9.5, bold=hero, color=NAVY if hero else GREY, line=1.20)

rect(s, M, 5.92, 11.98, 0.80, NAVY, rounded=True, adj=0.16)
text(s, M + 0.34, 6.08, 11.30, 0.48,
     "Масштабувати — лише після фактичного промо, підтверджених повторних покупок "
     "і стабільної обіговості.", size=12, bold=True, color=WHITE)
foot(s, "Комерційна рекомендація / розрахунковий сценарій.")

# ══ 13 · РІШЕННЯ ══════════════════════════════════════════════════════════
s = slide(dark=True)
cover_pic(s, BG_TITLE, 0, 0, W, H)
rect(s, 0, 0, 7.70, H, NAVY)
s.shapes.add_picture(LOGO, Inches(M), Inches(0.52), Inches(1.52), Inches(0.66))
rect(s, M, 1.70, 0.54, 0.09, ORANGE)
text(s, M, 2.02, 6.20, 0.30, "РІШЕННЯ", size=11, bold=True, color=AMBER)
text(s, M, 2.40, 6.40, 1.10, "Запускати тест,\nобмеживши обсяг закупівлі",
     size=32, bold=True, color=WHITE, line=1.10)
rect(s, M, 3.96, 6.20, 0.04, RGBColor(0x55, 0x60, 0x8C))
yy = 4.22
for n, t in [("1", "Два SKU на одній упаковочній платформі, пауч 220–250 г."),
             ("2", "Ціновий тест 109 / 119 / 129 грн; базовий сценарій — 119 грн."),
             ("3", "Один мережевий партнер плюс спеціалізований онлайн."),
             ("4", "До закупівлі з'ясувати, чому Ben's Original стоїть без залишку.")]:
    rect(s, M, yy, 0.34, 0.34, ORANGE, rounded=True, adj=0.5)
    text(s, M, yy, 0.34, 0.34, n, size=12, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, M + 0.52, yy + 0.02, 5.70, 0.32, t, size=11.5, color=WHITE, line=1.24)
    yy += 0.62
rect(s, 8.16, 1.90, 4.50, 4.34, RGBColor(0xFF, 0xFF, 0xFF), rounded=True, adj=0.05)
pic(s, "md/bibigo_bowl.jpg", 8.42, 2.14, 3.98, 2.90)
rect(s, 8.42, 5.16, 3.98, 0.03, MIST_D)
text(s, 8.42, 5.32, 3.98, 0.28, "Єдина позиція з підтвердженим залишком",
     size=10.5, bold=True, color=NAVY)
text(s, 8.42, 5.62, 3.98, 0.46, "Bibigo, білий рис 210 г · 135 грн\n«Смак Кореї», 8 упаковок",
     size=10, color=GREY, line=1.26)
text(s, M, 6.94, 6.40, 0.26, "Зріз онлайн-вітрин 22.09.2026", size=9,
     color=RGBColor(0x8A, 0x95, 0xB4))

# ══ 14 · ДОДАТОК · OTTOGI ═════════════════════════════════════════════════
s = slide()
header(s, "ДОДАТОК", "Страви Ottogi: 255 грн і жодного залишку",
       "Pulsar: шість карток, усі без залишку. Інший сегмент — рис із наповнювачем.")
OT = [("md/ottogi_chicken.jpg", "Гостра курка", "310/315* г"),
      ("md/ottogi_hamburg.jpg", "Гамбурзький стейк", "315 г"),
      ("md/ottogi_octopus.jpg", "Гострий восьминіг", "280 г"),
      ("md/ottogi_jjampong.jpg", "Jin Jjampong", "217,5 г"),
      ("md/ottogi_tuna.jpg", "Тунець і майонез", "247 г"),
      ("md/ottogi_kimchi.jpg", "Кімчі й тунець", "310 г")]
for i, (img, nm, g) in enumerate(OT):
    x = M + (i % 3) * 4.10
    y = 2.06 + (i // 3) * 2.38
    rect(s, x, y, 3.84, 2.16, WHITE, line=MIST_D, lw=1.2, rounded=True, adj=0.07)
    rect(s, x + 0.14, y + 0.12, 1.86, 1.92, MIST, rounded=True, adj=0.07)
    pic(s, img, x + 0.20, y + 0.18, 1.74, 1.80)
    text(s, x + 2.14, y + 0.36, 1.56, 0.40, nm, size=12, bold=True, color=NAVY, line=1.12)
    text(s, x + 2.14, y + 0.86, 1.56, 0.24, g, size=10, color=GREY)
    text(s, x + 2.14, y + 1.16, 1.56, 0.34, "255 грн", size=15, bold=True, color=ORANGE)
    pill(s, x + 2.14, y + 1.62, "НЕМАЄ", WHITE, CORAL, w=0.86, size=8)
text(s, M, 6.76, 11.98, 0.30,
     "* Яловичина / бульгогі та свинина 269 / 310 г на Rozetka: ідентичність SKU "
     "потребує звірки EAN та етикетки. Повторні пропозиції продавців не рахуються окремими SKU.",
     size=9, color=GREY, line=1.16)
foot(s, "Джерела: O01–O06, S10, R01–R12 · картки Pulsar і продавців Rozetka, 22.09.2026.")

# ══ 15 · ОСНОВА ДОСЛІДЖЕННЯ ═══════════════════════════════════════════════
s = slide()
header(s, "МЕТОДОЛОГІЯ", "Основа дослідження та межі висновків",
       "Відкриті українські онлайн-вітрини, зріз на 22 вересня 2026 року.")
COLS = [("ЩО ПЕРЕВІРЕНО", ORANGE,
         ["Сільпо · MAUDAU · Смак Кореї", "Edison Lee · Prom · Rozetka",
          "Pulsar · спеціалізовані магазини", "14 порівнянних варіантів гарнірів",
          "31 перевірена картка товару"]),
        ("ЩО НЕ ПЕРЕВІРЕНО", CORAL,
         ["АТБ · Novus · Metro · Varus", "Auchan · Fozzy · Таврія В",
          "Офлайн-полиця в жодній мережі", "Фактичні продажі та обіговість",
          "Причина відсутності залишку"]),
        ("МЕЖІ ВИСНОВКІВ", TEAL,
         ["Частота SKU ≠ частка продажів", "Ціна відсутнього товару — орієнтир",
          "Відсутність у пошуку ≠ відсутність у мережі",
          "Економіка — розрахунок, не прогноз",
          "Собівартість потребує оновлення"])]
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
     "Точні посилання — у примітках відповідних слайдів. Розрахункові параметри "
     "запуску та економіки позначені як допущення.", size=10.5, color=GREY)
foot(s, "Morskyi Dim · маркетингове дослідження · вересень 2026.")

prs.save("Gotovyi_Rys_Ukraina_MD.pptx")
print("saved Gotovyi_Rys_Ukraina_MD.pptx ·", len(prs.slides._sldIdLst), "slides")
