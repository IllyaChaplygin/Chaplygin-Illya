"""RTE Rice · Ukraine — SHELF deck.

A different design system from both earlier decks: light warm paper instead of
the suppliers' dark green, Segoe UI + Georgia instead of Cambria + Calibri, and
photography carrying the page instead of filled panels. Every product claim on
a slide has the pack shot next to it.
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from PIL import Image

W, H = 13.3333, 7.5
M = 0.70                                   # page margin

PAPER  = RGBColor(0xF5, 0xF1, 0xE9)
CARD   = RGBColor(0xFF, 0xFE, 0xFB)
SAND   = RGBColor(0xE9, 0xE1, 0xD2)
LINE   = RGBColor(0xD9, 0xCF, 0xBE)
INK    = RGBColor(0x16, 0x13, 0x0F)
MUTED  = RGBColor(0x8A, 0x80, 0x73)
SIGNAL = RGBColor(0xD6, 0x45, 0x2B)        # terracotta — the category accent
PINE   = RGBColor(0x1D, 0x3A, 0x31)        # deep green for full-bleed panels
MOSS   = RGBColor(0x3E, 0x6B, 0x55)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
CREAM  = RGBColor(0xF2, 0xEE, 0xE4)

UI, NUM = "Segoe UI", "Georgia"

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(W), Inches(H)
BLANK = prs.slide_layouts[6]
PAGE = {"n": 0}


# ── primitives ────────────────────────────────────────────────────────────
def rect(s, x, y, w, h, fill, line=None, lw=1.0, rounded=False, adj=0.08):
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
         font=UI, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, line=None, caps=False):
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


def rule(s, x, y, w, color=LINE, h=0.011):
    return rect(s, x, y, w, h, color)


def picture_contain(s, path, x, y, w, h):
    """Whole product, never cropped, centred in the box."""
    iw, ih = Image.open(path).size
    sc = min(w / iw, h / ih)
    pw, ph = iw * sc, ih * sc
    return s.shapes.add_picture(path, Inches(x + (w - pw) / 2), Inches(y + (h - ph) / 2),
                                Inches(pw), Inches(ph))


def picture_cover(s, path, x, y, w, h):
    iw, ih = Image.open(path).size
    sc = max(w / iw, h / ih)
    pw, ph = iw * sc, ih * sc
    pic = s.shapes.add_picture(path, Inches(x - (pw - w) / 2), Inches(y - (ph - h) / 2),
                               Inches(pw), Inches(ph))
    pic.crop_left = pic.crop_right = max(0.0, (pw - w) / pw / 2)
    pic.crop_top = pic.crop_bottom = max(0.0, (ph - h) / ph / 2)
    pic.left, pic.top = Inches(x), Inches(y)
    pic.width, pic.height = Inches(w), Inches(h)
    return pic


def slide(ground=PAPER):
    s = prs.slides.add_slide(BLANK)
    rect(s, 0, 0, W, H, ground)
    PAGE["n"] += 1
    return s


def head(s, eyebrow, title, dek=None, tone=INK):
    rect(s, M, 0.52, 0.30, 0.055, SIGNAL)
    text(s, M + 0.44, 0.40, 8.6, 0.28, eyebrow, size=10.5, bold=True, color=SIGNAL)
    text(s, M, 0.78, 11.2, 0.72, title, size=29, bold=True, color=tone, font=UI, line=1.02)
    if dek:
        text(s, M, 1.56, 11.0, 0.34, dek, size=12.5, color=MUTED, line=1.20)


def foot(s, src):
    rule(s, M, 6.92, W - 2 * M, LINE)
    text(s, M, 7.04, 11.0, 0.28, src, size=8.5, color=MUTED, line=1.16)
    text(s, W - M - 0.60, 7.04, 0.60, 0.28, f"{PAGE['n']:02d}", size=10, bold=True,
         color=SIGNAL, font=NUM, align=PP_ALIGN.RIGHT)


def chip(s, x, y, label, fg=SIGNAL, bg=None, w=None):
    w = w or (0.26 + 0.078 * len(label))
    rect(s, x, y, w, 0.28, bg, line=None if bg else fg, lw=1.0, rounded=True, adj=0.5)
    text(s, x, y, w, 0.28, label, size=8.5, bold=True, color=WHITE if bg else fg,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    return w


def bignum(s, x, y, w, value, unit, label, sub=None, color=SIGNAL):
    size = 48 if len(value) <= 2 else 38
    text(s, x, y, w, 0.72, value, size=size, bold=True, font=NUM, color=color)
    text(s, x, y + 0.76, w, 0.26, unit, size=11, bold=True, color=MUTED)
    rule(s, x, y + 1.08, 1.10, SIGNAL)
    text(s, x, y + 1.24, w, 0.30, label, size=12, bold=True, color=INK)
    if sub:
        text(s, x, y + 1.56, w, 0.52, sub, size=10, color=MUTED, line=1.18)


# ══ 01 · COVER ════════════════════════════════════════════════════════════
s = slide(PINE)
picture_cover(s, "assets/bg_field.jpg", 6.30, 0, 7.04, H)
rect(s, 6.30, 0, 7.04, H, PINE)          # scrim
rect(s, 6.30, 0, 7.04, H, None)
rect(s, 7.34, 1.18, 5.10, 5.14, CARD)
picture_contain(s, "assets/ua_bibigo.jpg", 7.58, 1.42, 4.62, 3.90)
rule(s, 7.72, 5.48, 4.34, LINE)
text(s, 7.72, 5.64, 4.34, 0.28, "bibigo · CJ Foods — Cooked White Rice 210 г",
     size=11, bold=True, color=INK)
text(s, 7.72, 5.92, 3.00, 0.26, "smak-korea.com.ua · у наявності", size=9, color=MUTED)
text(s, 10.72, 5.86, 1.34, 0.30, "135 ₴", size=15, bold=True, color=SIGNAL,
     font=NUM, align=PP_ALIGN.RIGHT)
text(s, 7.34, 6.44, 5.10, 0.26, "єдина позиція категорії, що має залишок на 22.09.2026",
     size=8.5, italic=True, color=RGBColor(0x8F, 0xAB, 0x9C), align=PP_ALIGN.CENTER)
rect(s, 0, 0, 6.40, H, PINE)
rect(s, M, 1.30, 0.34, 0.06, SIGNAL)
text(s, M, 1.62, 5.30, 0.30, "РИНОК УКРАЇНИ · ПЕРЕВІРКА ПОЛИЦІ", size=11, bold=True,
     color=RGBColor(0xE9, 0x9C, 0x74))
text(s, M, 2.02, 5.30, 1.80, "Ready-to-eat рис", size=44, bold=True, color=WHITE,
     font=UI, line=1.00)
text(s, M, 3.10, 5.30, 0.90, "Хто вже стоїть на полиці — і за скільки",
     size=16, color=RGBColor(0xC9, 0xDA, 0xCF), font=NUM, line=1.22)
rule(s, M, 4.16, 5.30, RGBColor(0x3E, 0x6B, 0x55))
text(s, M, 4.36, 5.30, 1.30,
     "Ретортний і асептичний пауч / cup / лоток · розігрів 60–90 сек · "
     "зберігання за кімнатної температури 12–18 міс.",
     size=11.5, color=RGBColor(0xA9, 0xC2, 0xB4), line=1.28)
rect(s, M, 5.86, 5.30, 0.90, RGBColor(0x17, 0x2F, 0x28))
text(s, M + 0.26, 5.98, 4.80, 0.66,
     [[("Друга редакція. ", {"bold": True, "color": RGBColor(0xE9, 0x9C, 0x74)}),
       ("Перша версія стверджувала, що категорії в Україні немає. "
        "Це було неточно — див. слайд 02.", {"color": RGBColor(0xC9, 0xDA, 0xCF)})]],
     size=10.5, line=1.24)
text(s, M, 6.96, 3.40, 0.28, "Зріз полиці · 22 вересня 2026 р.", size=10,
     color=RGBColor(0x8F, 0xAB, 0x9C))
text(s, 4.10, 6.96, 1.90, 0.28, "Morskyi Dim Group", size=10, bold=True,
     color=RGBColor(0xE9, 0x9C, 0x74), align=PP_ALIGN.RIGHT)

# ══ 02 · CORRECTION ═══════════════════════════════════════════════════════
s = slide()
head(s, "ПОПРАВКА ДО ПЕРШОЇ РЕДАКЦІЇ", "Категорія в Україні є. Дистрибуції — немає",
     "Ви маєте рацію: у пошуку товар видно. Ось що змінилося після повторної перевірки.")

rect(s, M, 2.06, 5.86, 4.10, CARD, line=LINE)
chip(s, M + 0.34, 2.32, "БУЛО СКАЗАНО", MUTED)
text(s, M + 0.34, 2.82, 5.20, 1.00,
     "«Жодної позиції RTE-рису в Україні. 0 SKU у 18 мережах, 2 SKU на весь "
     "спеціалізований канал».", size=13, color=INK, line=1.26, italic=True)
rule(s, M + 0.34, 3.98, 5.20)
text(s, M + 0.34, 4.14, 5.20, 1.70,
     "Помилка була не в цифрах по мережах — вони підтвердились. Помилка була в "
     "тому, що маркетплейси Rozetka, Prom і Епіцентр я позначив як «недоступні "
     "для перевірки» і не став шукати далі. Саме там товар і продається.",
     size=11, color=MUTED, line=1.30)

rect(s, 6.78, 2.06, 5.86, 4.10, CARD, line=SIGNAL, lw=1.6)
chip(s, 7.12, 2.32, "ПЕРЕВІРЕНО ПОВТОРНО", None, SIGNAL)
text(s, 7.12, 2.82, 5.20, 1.00,
     "Чотири позиції справді присутні в українському ритейлі. Дві з них — "
     "у наявності просто зараз.", size=13, bold=True, color=INK, line=1.26)
rule(s, 7.12, 3.98, 5.20)
rows = [("bibigo Cooked White Rice 210 г", "smak-korea.com.ua", "135 ₴", "у наявності", SIGNAL),
        ("«Готовий рис з Кореї»", "asia-goods.com.ua", "190 ₴", "у наявності", SIGNAL),
        ("Рис швидк. пригот. Bibigo 210 г", "Prom.ua · ТАЙЯКІ МАРТ", "189 ₴", "недоступний", MUTED),
        ("Рис готовий Ottogi бібімбап 269 г", "Епіцентр · Gurmissimo", "н/д", "немає в наявності", MUTED)]
yy = 4.16
for nm, shop, pr, st, c in rows:
    text(s, 7.12, yy, 3.00, 0.26, nm, size=10, bold=True, color=INK)
    text(s, 7.12, yy + 0.22, 3.00, 0.24, shop, size=8.5, color=MUTED)
    text(s, 10.28, yy + 0.02, 0.86, 0.26, pr, size=11, bold=True, color=c,
         font=NUM, align=PP_ALIGN.RIGHT)
    text(s, 11.20, yy + 0.04, 1.12, 0.24, st, size=8.5, bold=True, color=c,
         align=PP_ALIGN.RIGHT)
    yy += 0.50
foot(s, "Джерело: smak-korea.com.ua (Store API), asia-goods.com.ua, prom.ua, epicentrk.ua — "
        "перевірено 22.09.2026. Мережі — stores-api.zakaz.ua (17 мереж) і sf-ecom-api.silpo.ua.")

# ══ 03 · HEADLINE ═════════════════════════════════════════════════════════
s = slide()
head(s, "ГОЛОВНЕ", "Лістинг є. Полиці немає",
     "Категорія живе в довгому хвості маркетплейсів, а не в дистрибуції мереж.")
for i, (v, u, lb, sub) in enumerate([
        ("4", "лістинги", "знайдено в Україні", "маркетплейси та спеціалізовані магазини"),
        ("2", "з них у наявності", "решта — out of stock", "лістинг без товару — це ще не дистрибуція"),
        ("0", "SKU", "у 18 мережах", "жодної позиції у роздрібних мережах"),
        ("64,3", "₴ за 100 г", "нижня полична ціна", "bibigo 210 г за 135 ₴")]):
    x = M + i * 3.03
    bignum(s, x, 2.18, 2.80, v, u, lb, sub)

rect(s, M, 4.66, W - 2 * M, 1.96, CARD, line=LINE)
picture_contain(s, "assets/ua_bibigo.jpg", M + 0.22, 4.84, 1.60, 1.60)
text(s, M + 2.06, 4.92, 4.40, 0.30, "Що це означає для закупівлі", size=13,
     bold=True, color=INK)
text(s, M + 2.06, 5.28, 4.40, 1.14,
     "Конкурента, чию частку треба відбирати, немає. Але немає й виробленої "
     "звички: жодна мережа ще не вчила покупця, що рис буває готовим.",
     size=10.5, color=MUTED, line=1.28)
rect(s, 7.30, 4.84, 0.05, 1.60, SIGNAL)
text(s, 7.54, 4.92, 5.00, 0.30, "Чому це вигідніше за зайнятий ринок", size=13,
     bold=True, color=INK)
text(s, 7.54, 5.28, 5.00, 1.14,
     "Наш пауч 240 г обходиться в 16,7 ₴/100 г — учетверо дешевше за єдину "
     "полічну ціну в країні. Такий запас дозволяє одночасно тримати маржу "
     "й заходити нижче корейського імпорту.", size=10.5, color=MUTED, line=1.28)
foot(s, "Джерело: власна перевірка полиць і лістингів, 22.09.2026; собівартість — колода "
        "«RTE Rice · Постачальники», курс 45,00 ₴/$.")

# ══ 04 · THE FOUR LISTINGS, WITH PHOTOS ═══════════════════════════════════
s = slide()
head(s, "КАРТА КАТЕГОРІЇ · 1/2", "Усе, що продається в Україні як готовий рис",
     "Чотири лістинги. Усі — прямий корейський імпорт, без локалізації смаку та упаковки.")

CARDS = [("assets/ua_bibigo.jpg", "bibigo · CJ Foods", "Cooked White Rice 210 г",
          "135 ₴", "64,3 ₴/100 г", "smak-korea.com.ua", "У НАЯВНОСТІ", SIGNAL,
          "Product of S. Korea · імпортер H Mart Europe (UK) · мікрохвильовка 90 сек"),
         ("assets/ua_bibigo.jpg", "Bibigo · той самий SKU", "Рис швидк. пригот. 210 г",
          "189 ₴", "90,0 ₴/100 г", "Prom.ua · ТАЙЯКІ МАРТ", "НЕДОСТУПНИЙ", MUTED,
          "Той самий товар на маркетплейсі — на 40 % дорожче й без залишку"),
         (None, "Ottogi", "Бібімбап з овочами та м'ясом 269 г",
          "н/д", "ціну не показано", "Епіцентр · Gurmissimo", "НЕМАЄ В НАЯВНОСТІ", MUTED,
          "Лістинг у мережевому маркетплейсі, продавець — сторонній"),
         (None, "без бренду", "«Готовий рис з Кореї», арт. 10296",
          "190 ₴", "вага не вказана", "asia-goods.com.ua", "У НАЯВНОСТІ", SIGNAL,
          "Ані ваги, ані виробника на картці — потрібна тестова закупка")]
for i, (img, brand, name, price, per, shop, st, c, note) in enumerate(CARDS):
    x = M + i * 3.03
    rect(s, x, 2.12, 2.80, 4.54, CARD, line=LINE)
    rect(s, x, 2.12, 2.80, 0.05, c)
    if img:
        picture_contain(s, img, x + 0.18, 2.26, 2.44, 1.62)
    else:
        rect(s, x + 0.24, 2.30, 2.32, 1.50, SAND)
        text(s, x + 0.24, 2.92, 2.32, 0.30, "ФОТО ВІДСУТНЄ", size=9, bold=True,
             color=MUTED, align=PP_ALIGN.CENTER)
        text(s, x + 0.24, 3.20, 2.32, 0.24, "на картці продавця", size=8,
             color=MUTED, align=PP_ALIGN.CENTER)
    chip(s, x + 0.24, 3.94, st, c if c == SIGNAL else MUTED)
    text(s, x + 0.24, 4.36, 2.32, 0.24, brand.upper(), size=8.5, bold=True, color=MUTED)
    text(s, x + 0.24, 4.62, 2.32, 0.56, name, size=12, bold=True, color=INK, line=1.12)
    rule(s, x + 0.24, 5.26, 2.32)
    text(s, x + 0.24, 5.38, 2.32, 0.42, price, size=23, bold=True, color=c, font=NUM)
    text(s, x + 0.24, 5.84, 2.32, 0.22, per, size=9.5, color=SIGNAL if c == SIGNAL else MUTED)
    text(s, x + 0.24, 6.06, 2.32, 0.22, shop, size=8.5, bold=True, color=INK)
    text(s, x + 0.24, 6.28, 2.32, 0.36, note, size=7.5, color=MUTED, line=1.14)
foot(s, "Джерело: smak-korea.com.ua (WooCommerce Store API, фото — картка товару), "
        "prom.ua, epicentrk.ua, asia-goods.com.ua — зчитано 22.09.2026.")

# ══ 05 · WHAT GOOGLE SHOWS BUT IS NOT OUR CATEGORY ════════════════════════
s = slide()
head(s, "КАРТА КАТЕГОРІЇ · 2/2", "Що видає пошук — і що з цього не наша категорія",
     "За запитом «готовий рис» Google показує здебільшого сусідні технології. Ось вони.")

NEAR = [("assets/ua_selfheat_187.jpg", "Саморозігрівальний рис", "Dacia · 187 г",
         "680 ₴", "363,6 ₴/100 г", "хімічний нагрівач у коробці"),
        ("assets/ua_selfheat_272.jpg", "Саморозігрівальний рис", "Dacia · 272 г",
         "799 ₴", "293,8 ₴/100 г", "з тушкованою яловичиною"),
        ("assets/sub_noodle_cup.jpg", "Локшина б/п у стакані", "Nongshim · 67 г",
         "103 ₴", "153,7 ₴/100 г", "заливається окропом"),
        ("assets/sub_rice_bags.jpg", "Рис у варильних пакетах", "Novus · 5×80 г",
         "38,59 ₴", "9,6 ₴/100 г", "варіння 15–20 хв")]
for i, (img, kind, sku, price, per, how) in enumerate(NEAR):
    x = M + i * 3.03
    rect(s, x, 2.12, 2.80, 2.96, CARD, line=LINE)
    picture_contain(s, img, x + 0.26, 2.26, 2.28, 1.34)
    text(s, x + 0.24, 3.72, 2.32, 0.24, kind, size=10.5, bold=True, color=INK)
    text(s, x + 0.24, 3.98, 2.32, 0.22, sku, size=9, color=MUTED)
    text(s, x + 0.24, 4.26, 2.32, 0.34, price, size=15, bold=True, color=INK, font=NUM)
    text(s, x + 0.24, 4.64, 2.32, 0.24, per, size=9, color=SIGNAL)
    text(s, x + 0.24, 4.86, 2.32, 0.22, how, size=8, color=MUTED)

rect(s, M, 5.32, W - 2 * M, 1.34, PINE)
text(s, M + 0.34, 5.50, 3.40, 0.30, "Чому це не наш продукт", size=13, bold=True,
     color=RGBColor(0xE9, 0x9C, 0x74))
text(s, M + 0.34, 5.86, 5.60, 0.66,
     "Саморозігрівальний рис має хімічний нагрівач і коштує вп'ятеро дорожче. "
     "Локшина та сублімований SubliMate потребують окропу. Garde Manger — "
     "заморожений, страви Сільпо — охолоджені на 2–5 діб.",
     size=10.5, color=RGBColor(0xC9, 0xDA, 0xCF), line=1.26)
rect(s, 7.10, 5.50, 0.05, 0.98, SIGNAL)
text(s, 7.34, 5.50, 5.24, 0.98,
     "Наш продукт — єдиний, що поєднує кімнатне зберігання 12–18 місяців, "
     "90 секунд до готовності й ціну гарніра, а не делікатесу. "
     "У цій точці на полиці зараз немає нікого.",
     size=11, bold=True, color=WHITE, line=1.26)
foot(s, "Джерело: prom.ua (картки товарів із цінами), stores-api.zakaz.ua, "
        "sf-ecom-api.silpo.ua — 22.09.2026. ₴/100 г — наш перерахунок.")

# ══ 06 · CHAINS ═══════════════════════════════════════════════════════════
s = slide()
head(s, "ДИСТРИБУЦІЯ", "У мережах — жодної позиції",
     "21 пошуковий запит у кожній мережі, перевірено переглядом назв товарів.")

rect(s, M, 2.06, 7.30, 4.60, CARD, line=LINE)
GROUPS = [("METRO · Ашан · NOVUS", "0"), ("МегаМаркет · Ultramarket", "0"),
          ("ЕКО маркет · Таврія В", "0"), ("Епіцентр FOOD · Космос", "0"),
          ("Восторг · Харків · Чудо Маркет", "0"),
          ("Гроно · Ідеал · ОнDe · Za Raz · Torba", "0"), ("Сільпо", "0")]
yy = 2.30
for i, (nm, v) in enumerate(GROUPS):
    if i:
        rule(s, M + 0.30, yy - 0.06, 6.70)
    text(s, M + 0.30, yy, 5.40, 0.30, nm, size=11, color=INK, anchor=MSO_ANCHOR.MIDDLE)
    text(s, M + 5.90, yy, 1.10, 0.30, v, size=15, bold=True, color=MUTED, font=NUM,
         align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    yy += 0.60

rect(s, 8.44, 2.06, 4.20, 4.60, PINE)
text(s, 8.74, 2.32, 3.60, 0.30, "Що стоїть натомість", size=13, bold=True,
     color=RGBColor(0xE9, 0x9C, 0x74))
picture_contain(s, "assets/sub_noodle_cup.jpg", 8.74, 2.76, 1.60, 1.30)
picture_contain(s, "assets/sub_tteok.jpg", 10.50, 2.76, 1.60, 1.30)
text(s, 8.74, 4.20, 3.60, 0.96,
     "Локшина швидкого приготування, рамьон, токпоккі, рис у варильних пакетах "
     "і охолоджені страви власної кухні мережі.",
     size=10.5, color=RGBColor(0xC9, 0xDA, 0xCF), line=1.26)
rule(s, 8.74, 5.32, 3.60, RGBColor(0x3E, 0x6B, 0x55))
text(s, 8.74, 5.48, 3.60, 0.96,
     "Пошук у мережах нечіткий: запит «hetbahn» у Сільпо повертає 15 позицій — "
     "усі горілка «Гетьман». Тому нулі перевірено очима.",
     size=9.5, italic=True, color=RGBColor(0x9E, 0xB8, 0xA9), line=1.26)
foot(s, "Джерело: stores-api.zakaz.ua (17 торгових мереж) і sf-ecom-api.silpo.ua, "
        "запити 22.09.2026. АТБ, Varus і Фора машинно не зчитуються — потрібна ручна перевірка.")

# ══ 07 · OUR FORMATS ══════════════════════════════════════════════════════
s = slide()
head(s, "НАША ПРОПОЗИЦІЯ", "Чим ми можемо зайняти цю полицю",
     "Формат не обмежений звичкою покупця — категорії ще немає, тож ми задаємо її самі.")

FMT = [("assets/own_pouch_jasmine.jpg", "Пауч 240 г", "BSCM · Jasmine", "40,15 ₴", "16,7"),
       ("assets/own_chefrey_turm.jpg", "Пауч 250 г", "Chefrey · organic", "45,78 ₴", "18,3"),
       ("assets/own_cup.jpg", "Cup 150 г", "BSCM · single cup", "34,95 ₴", "23,3"),
       ("assets/own_spar_cups.jpg", "Cup 2×125 г", "BSCM · private label", "64,51 ₴", "25,8"),
       ("assets/own_kbros_kimchi.jpg", "Лоток 200 г", "KBROS · Korea", "87,26 ₴", "43,6")]
cw = (W - 2 * M - 4 * 0.24) / 5
for i, (img, fmt, who, cost, per) in enumerate(FMT):
    x = M + i * (cw + 0.24)
    rect(s, x, 2.12, cw, 3.30, CARD, line=LINE)
    picture_contain(s, img, x + 0.16, 2.28, cw - 0.32, 1.42)
    text(s, x + 0.20, 3.86, cw - 0.40, 0.28, fmt, size=12, bold=True, color=INK)
    text(s, x + 0.20, 4.14, cw - 0.40, 0.24, who, size=8.5, color=MUTED)
    rule(s, x + 0.20, 4.46, cw - 0.40)
    text(s, x + 0.20, 4.58, cw - 0.40, 0.34, cost, size=16, bold=True, color=INK, font=NUM)
    text(s, x + 0.20, 4.98, cw - 0.40, 0.24, per + " ₴/100 г", size=9.5, bold=True,
         color=SIGNAL)

rect(s, M, 5.62, W - 2 * M, 1.04, SAND)
text(s, M + 0.30, 5.78, 11.4, 0.74,
     [[("Собівартість — завантаження 40′ FCL, курс 45,00 ₴/$. ", {"bold": True}),
       ("Пауч 240 г — найдешевший грам у лінійці й найближчий формат до того, "
        "що вже приймає полиця. Cup 2×125 г існує лише як private label — "
        "у виробника власної марки в цьому форматі немає.", {"color": MUTED})]],
     size=10.5, line=1.26)
foot(s, "Джерело: колода «RTE Rice · Постачальники» (Self Cost Rice Brand.xlsx), 40′ FCL, "
        "курс 45,00 ₴/$; фото — офіційні знімки упаковки постачальників. ₴/100 г — наш перерахунок.")

# ══ 08 · PRICE LADDER OF THE SHELF ════════════════════════════════════════
s = slide()
head(s, "ЦІНОВА ШКАЛА", "Скільки коштує та сама порція на сусідніх полицях",
     "Єдина спільна одиниця виміру — гривня за 100 грамів.")

BANDS = [("Рис сухий, варильні пакети", 9.6, "Novus 5×80 г · 38,59 ₴", MUTED),
         ("Рис сухий, медіана мереж", 16.5, "410 SKU у вибірці", MUTED),
         ("Наш пауч 240 г · собівартість", 16.7, "40′ FCL, до націнки", SIGNAL),
         ("Каші швидкого приготування", 48.9, "медіана, 102 SKU", MUTED),
         ("Локшина швидкого приготування", 48.9, "медіана, 478 SKU", MUTED),
         ("Охолоджені страви з рисом, Сільпо", 52.0, "медіана, 11 SKU", MUTED),
         ("bibigo 210 г · полиця України", 64.3, "smak-korea, 135 ₴", SIGNAL),
         ("Bibigo 210 г · маркетплейс", 90.0, "Prom.ua, 189 ₴", MUTED),
         ("Токпоккі", 143.7, "медіана, 36 SKU", MUTED),
         ("Саморозігрівальний рис", 293.8, "Dacia 272 г · 799 ₴", MUTED)]
MAXV, BAR_X, BAR_W = 300.0, 5.70, 5.30
yy = 2.08
for nm, v, note, c in BANDS:
    text(s, M, yy, 4.70, 0.30, nm, size=10.5, bold=(c == SIGNAL), color=INK,
         anchor=MSO_ANCHOR.MIDDLE)
    rect(s, BAR_X, yy + 0.09, BAR_W, 0.14, RGBColor(0xEA, 0xE3, 0xD6))
    rect(s, BAR_X, yy + 0.09, max(0.06, BAR_W * v / MAXV), 0.14,
         SIGNAL if c == SIGNAL else RGBColor(0xB9, 0xAF, 0x9E))
    text(s, 11.10, yy, 0.86, 0.30, f"{v:.1f}".replace(".", ","), size=11.5, bold=True,
         color=c if c == SIGNAL else INK, font=NUM, align=PP_ALIGN.RIGHT,
         anchor=MSO_ANCHOR.MIDDLE)
    text(s, 12.02, yy, 0.60, 0.30, note[:0] or "", size=8, color=MUTED)
    yy += 0.39

text(s, 11.10, yy - 0.02, 1.50, 0.26, "₴ за 100 г", size=9, bold=True, color=MUTED,
     align=PP_ALIGN.RIGHT)
rect(s, M, 6.20, W - 2 * M, 0.54, SAND)
text(s, M + 0.30, 6.31, 11.4, 0.34,
     [[("Вікно для РРЦ — між 37 і 54 ₴/100 г. ", {"bold": True, "color": SIGNAL}),
       ("Вище сухого рису, бо це готова страва; нижче корейського імпорту "
        "й охолодженої кулінарії, бо це наша перевага.", {"color": INK})]], size=10.5)
foot(s, "Джерело: розрахунок за каталогами 17 мереж (stores-api.zakaz.ua), Сільпо, "
        "prom.ua та smak-korea.com.ua — 22.09.2026. Питома ціна рахується лише для позицій із вагою.")

# ══ 09 · PRICE SCENARIOS ══════════════════════════════════════════════════
s = slide()
head(s, "ЦІНОУТВОРЕННЯ · ОЦІНКА", "Де поставити полицю нашому паучу 240 г",
     "Модель: РРЦ без ПДВ, націнка мережі 35 %, собівартість 40,15 ₴ (40′ FCL).")

rect(s, M, 2.06, 4.50, 4.60, PINE)
picture_contain(s, "assets/own_pouch_jasmine.jpg", M + 0.30, 2.26, 3.90, 1.70)
text(s, M + 0.30, 4.14, 3.90, 0.30, "BSCM · Jasmine Rice, пауч 240 г", size=12,
     bold=True, color=WHITE)
for lb, vl in [("Собівартість 40′ FCL", "40,15 ₴"), ("Те саме за 100 г", "16,7 ₴"),
               ("Полиця bibigo за 100 г", "64,3 ₴"), ("Запас", "×3,8")]:
    pass
yy = 4.56
for lb, vl in [("Собівартість, 40′ FCL", "40,15 ₴"), ("за 100 г", "16,7 ₴"),
               ("Полиця bibigo, за 100 г", "64,3 ₴"), ("Запас по грамах", "×3,8")]:
    text(s, M + 0.30, yy, 2.50, 0.26, lb, size=10, color=RGBColor(0xA9, 0xC2, 0xB4))
    text(s, M + 2.80, yy, 1.40, 0.26, vl, size=11.5, bold=True, color=WHITE, font=NUM,
         align=PP_ALIGN.RIGHT)
    yy += 0.44

rect(s, 5.52, 2.06, 7.12, 4.60, CARD, line=LINE)
cols = [(5.86, 1.20, "РРЦ", PP_ALIGN.LEFT), (7.30, 1.10, "₴/100 Г", PP_ALIGN.RIGHT),
        (8.62, 1.30, "vs BIBIGO", PP_ALIGN.RIGHT), (10.14, 1.30, "МАРЖА", PP_ALIGN.RIGHT),
        (11.62, 0.76, "%", PP_ALIGN.RIGHT)]
for x, w, lb, al in cols:
    text(s, x, 2.32, w, 0.24, lb, size=8.5, bold=True, color=MUTED, align=al)
rule(s, 5.86, 2.62, 6.52)
SCEN = [("89 ₴", "37,1", "−42 %", "14,79 ₴", "27 %", False),
        ("99 ₴", "41,2", "−36 %", "20,96 ₴", "34 %", False),
        ("109 ₴", "45,4", "−29 %", "27,13 ₴", "40 %", True),
        ("119 ₴", "49,6", "−23 %", "33,31 ₴", "45 %", True),
        ("129 ₴", "53,8", "−16 %", "39,48 ₴", "50 %", False)]
yy = 2.78
for a, b, c, d, e, hl in SCEN:
    if hl:
        rect(s, 5.72, yy - 0.04, 6.76, 0.50, RGBColor(0xF7, 0xE7, 0xDF))
    vals = [a, b, c, d, e]
    for (x, w, _, al), v in zip(cols, vals):
        text(s, x, yy, w, 0.42, v,
             size=15 if v is a else 11.5, bold=True if (hl or v is a) else False,
             color=SIGNAL if hl else INK, font=NUM if v is a else UI,
             align=al, anchor=MSO_ANCHOR.MIDDLE)
    yy += 0.52
rule(s, 5.86, yy + 0.02, 6.52)
text(s, 5.86, yy + 0.18, 6.52, 1.00,
     "Рекомендація: стартувати з 109–119 ₴. Це на 23–29 % дешевше за корейський "
     "імпорт, лишає мережі повну націнку, а нам — 40–45 % маржі. Умови входу "
     "(listing fee, промо, відстрочка) в модель не закладені.",
     size=10.5, color=MUTED, line=1.28)
foot(s, "Джерело: собівартість — колода «RTE Rice · Постачальники» (BSCM пауч 240 г, 40′ FCL, "
        "курс 45,00 ₴/$); полиця bibigo — smak-korea.com.ua, 22.09.2026. Сценарії — наш розрахунок.")

# ══ 10 · ASSORTMENT ARCHITECTURE ══════════════════════════════════════════
s = slide()
head(s, "АРХІТЕКТУРА АСОРТИМЕНТУ", "Що саме ставити на полицю першим",
     "Погляд категорійного менеджера: три ролі, а не список SKU.")

ROLES = [("assets/own_pouch_jasmine.jpg", "ТРАФІК", "Jasmine 240 г",
          "Найдешевший грам, найзрозуміліший смак. Тримає цінове сприйняття "
          "категорії й перший пробний купівельний акт.", "109 ₴", SIGNAL),
         ("assets/own_pouch_mexican.jpg", "МАРЖА", "Mexican · Hot & Spicy 240 г",
          "Смакові версії за тією ж собівартістю +5–10 %, але з вищою РРЦ. "
          "Саме вони роблять категорію прибутковою.", "129 ₴", INK),
         ("assets/own_spar_cups.jpg", "ТЕСТ", "Cup 2×125 г, private label",
          "Формат для тесту в одній мережі. Виробник дає його лише під марку "
          "замовника — це вже готовий private label кейс (SPAR).", "159 ₴", MUTED)]
for i, (img, role, name, why, rrp, c) in enumerate(ROLES):
    x = M + i * 4.05
    rect(s, x, 2.12, 3.80, 4.54, CARD, line=LINE)
    rect(s, x, 2.12, 3.80, 0.05, c)
    picture_contain(s, img, x + 0.26, 2.34, 3.28, 1.56)
    chip(s, x + 0.26, 4.04, role, c)
    text(s, x + 0.26, 4.48, 3.28, 0.34, name, size=13.5, bold=True, color=INK)
    text(s, x + 0.26, 4.90, 3.28, 1.00, why, size=10, color=MUTED, line=1.28)
    rule(s, x + 0.26, 5.96, 3.28)
    text(s, x + 0.26, 6.10, 1.80, 0.34, rrp, size=17, bold=True, color=c, font=NUM)
    text(s, x + 2.10, 6.18, 1.44, 0.26, "орієнтовна РРЦ", size=8.5, color=MUTED,
         align=PP_ALIGN.RIGHT)
foot(s, "Ролі SKU — наша рекомендація на підставі цінової шкали (слайд 08) і собівартості "
        "(слайд 07). РРЦ — оцінка, не комерційна пропозиція.")

# ══ 11 · IMPORT STATS ═════════════════════════════════════════════════════
s = slide()
head(s, "ІМПОРТ", "HS 1904 90 — найближчий публічний код",
     "Окремого коду для RTE-рису не існує; це кошик готових виробів із зерна.")

rect(s, M, 2.06, 7.10, 4.60, CARD, line=LINE)
for x, w, lb, al in [(M + 0.30, 1.00, "РІК", PP_ALIGN.LEFT),
                     (M + 2.00, 1.60, "ТОНН", PP_ALIGN.RIGHT),
                     (M + 3.90, 1.60, "ТИС. $ CIF", PP_ALIGN.RIGHT),
                     (M + 5.70, 1.10, "$/КГ", PP_ALIGN.RIGHT)]:
    text(s, x, 2.32, w, 0.24, lb, size=8.5, bold=True, color=MUTED, align=al)
rule(s, M + 0.30, 2.62, 6.50)
DATA = [("2018", "956,6", "839,7", "0,88", False), ("2019", "17,6", "74,6", "4,24", True),
        ("2020", "1,2", "12,8", "10,67", True), ("2021", "0,4", "2,3", "5,75", True),
        ("2022", "362,2", "534,4", "1,48", False), ("2023", "287,5", "706,7", "2,46", False),
        ("2024", "207,2", "608,7", "2,94", False)]
yy = 2.76
for yr, t, v, pk, gap in DATA:
    col = MUTED if gap else INK
    text(s, M + 0.30, yy, 1.00, 0.30, yr, size=11, bold=True, color=col)
    text(s, M + 2.00, yy, 1.60, 0.30, t, size=11, color=col, align=PP_ALIGN.RIGHT, font=NUM)
    text(s, M + 3.90, yy, 1.60, 0.30, v, size=11, color=col, align=PP_ALIGN.RIGHT, font=NUM)
    text(s, M + 5.70, yy, 1.10, 0.30, pk, size=11, bold=True,
         color=MUTED if gap else SIGNAL, align=PP_ALIGN.RIGHT, font=NUM)
    yy += 0.46
text(s, M + 0.30, yy + 0.10, 6.50, 0.50,
     "2019–2021 — майже нулі. Це схоже на прогалину у звітності, а не на зникнення імпорту.",
     size=9.5, italic=True, color=SIGNAL, line=1.22)

rect(s, 8.24, 2.06, 4.40, 4.60, SAND)
text(s, 8.54, 2.32, 3.80, 0.30, "Межі цих даних", size=13, bold=True, color=INK)
for t in ["Код кошиковий: готовий рис — лише його частина.",
          "Публічний UN Comtrade віддає тільки світовий агрегат, розбивки за "
          "країнами немає — частку Таїланду та Кореї не видно.",
          "Даних за 2025 рік на дату зрізу немає."]:
    pass
yy = 2.76
for t in ["Код кошиковий: готовий рис — лише його частина.",
          "Публічний UN Comtrade віддає тільки світовий агрегат — частку Таїланду "
          "й Кореї з нього не видно.",
          "Даних за 2025 рік на дату зрізу немає."]:
    rect(s, 8.54, yy + 0.06, 0.10, 0.10, SIGNAL)
    text(s, 8.78, yy, 3.56, 0.90, t, size=10.5, color=INK, line=1.26)
    yy += 1.00
rect(s, 8.54, 5.78, 3.80, 0.72, CARD)
text(s, 8.74, 5.92, 3.44, 0.48,
     "Для точної картини потрібне вивантаження Держмитслужби за УКТЗЕД 1904 90 10.",
     size=9.5, bold=True, color=SIGNAL, line=1.24)
foot(s, "Джерело: UN Comtrade (comtradeapi.un.org), репортер Україна (804), потік «імпорт», "
        "код HS 190490, звернення 22.09.2026. $/кг — наш розрахунок.")

# ══ 12 · REGULATORY ═══════════════════════════════════════════════════════
s = slide()
head(s, "РЕГУЛЯТОРИКА", "Що потрібно, щоб завести продукт",
     "Оглядово, без глибокого юридичного аналізу — із позначенням зон ризику.")

rect(s, M, 2.06, 7.10, 4.60, CARD, line=LINE)
REG = [("Маркування", "Українською мовою — ЗУ № 2639-VIII «Про інформацію для "
        "споживачів щодо харчових продуктів»"),
       ("Обов'язкові дані", "Склад, поживна цінність, алергени, термін придатності, "
        "партія, країна походження, виробник та імпортер"),
       ("Оператор ринку", "Реєстрація потужностей імпортера в Держпродспоживслужбі, "
        "система НАССР"),
       ("Походження", "Рецептура рослинного походження — ветеринарний сертифікат "
        "не потрібен"),
       ("Платежі", "Мито 10 % і ПДВ 20 % уже враховані в собівартості")]
yy = 2.34
for lb, vl in REG:
    text(s, M + 0.30, yy, 1.80, 0.26, lb, size=10, bold=True, color=SIGNAL)
    text(s, M + 2.16, yy, 4.66, 0.72, vl, size=10, color=INK, line=1.24)
    yy += 0.88

rect(s, 8.24, 2.06, 4.40, 4.60, PINE)
text(s, 8.54, 2.32, 3.80, 0.30, "Зони ризику", size=13, bold=True,
     color=RGBColor(0xE9, 0x9C, 0x74))
yy = 2.78
for t1, t2 in [("Смаки з м'ясом і яйцем",
                "Курка, яйце, морепродукти переводять товар у режим продукції "
                "тваринного походження — інший пакет документів і затверджені "
                "потужності експортера."),
               ("Воєнні спрощення маркування",
                "Дозвіл на супровідну документацію українською замість друку на "
                "упаковці (КМУ, 21.03.2022). Чинність у 2026 році потребує "
                "висновку юриста."),
               ("Заяви на упаковці",
                "«Organic», «без глютену», «без консервантів» потребують "
                "підтвердження; organic — визнаної сертифікації.")]:
    rect(s, 8.54, yy + 0.04, 0.05, 0.94, SIGNAL)
    text(s, 8.76, yy, 3.60, 0.26, t1, size=10.5, bold=True, color=WHITE)
    text(s, 8.76, yy + 0.28, 3.60, 0.76, t2, size=9.5,
         color=RGBColor(0xC9, 0xDA, 0xCF), line=1.24)
    yy += 1.30
foot(s, "Джерело: ЗУ № 2639-VIII (zakon.rada.gov.ua/go/2639-19); повідомлення Кабінету "
        "Міністрів України від 21.03.2022. Чинність спрощень станом на 22.09.2026 не перевірялася.")

# ══ 13 · VERDICT ══════════════════════════════════════════════════════════
s = slide(PINE)
rect(s, 0, 0, W, H, PINE)
picture_cover(s, "assets/bg_plant.jpg", 7.60, 0, 5.74, H)
rect(s, 7.60, 0, 5.74, H, None)
rect(s, 0, 0, 7.90, H, PINE)
rect(s, M, 1.10, 0.30, 0.055, SIGNAL)
text(s, M + 0.44, 0.98, 6.60, 0.28, "ВИСНОВОК", size=10.5, bold=True,
     color=RGBColor(0xE9, 0x9C, 0x74))
text(s, M, 1.42, 6.60, 1.20, "Ніша вільна, але покупця доведеться вчити",
     size=30, bold=True, color=WHITE, line=1.06)
rule(s, M, 2.84, 6.60, RGBColor(0x3E, 0x6B, 0x55))
text(s, M, 3.04, 6.60, 1.30,
     "Категорія в Україні існує у вигляді чотирьох лістингів, з яких два мають "
     "залишок. У жодній мережі її немає. Витісняти нікого не треба — але й "
     "готового попиту, вихованого чужим маркетингом, теж немає.",
     size=12.5, color=RGBColor(0xC9, 0xDA, 0xCF), line=1.32)
yy = 4.56
for n, t in [("1", "Перший SKU — Jasmine 240 г за 109 ₴: найдешевший грам і найнижчий бар'єр проби."),
             ("2", "Комунікація продає не рис, а 90 секунд: конкурент — не крупа, а доставка."),
             ("3", "Вхід через одну мережу з обмеженим SKU-сетом, а не широкий лістинг одразу.")]:
    text(s, M, yy, 0.40, 0.30, n, size=17, bold=True, color=SIGNAL, font=NUM)
    text(s, M + 0.46, yy + 0.02, 6.14, 0.56, t, size=11, color=WHITE, line=1.26)
    yy += 0.72
rect(s, M, 6.82, 6.60, 0.42, RGBColor(0x17, 0x2F, 0x28))
text(s, M + 0.20, 6.88, 6.20, 0.30,
     "Зріз полиці 22.09.2026 · дані з позначкою «потребує перевірки» у висновок не включені",
     size=8.5, color=RGBColor(0x8F, 0xAB, 0x9C))

# ══ 14 · NEXT STEPS ═══════════════════════════════════════════════════════
s = slide()
head(s, "ПЛАН ДІЙ", "Що зробити до рішення про закупівлю",
     "Чотири кроки, кожен закриває конкретну прогалину цього дослідження.")

STEPS = [("01", "Ручний моніторинг полиці", "АТБ, Сільпо, Novus, Ашан, METRO — 2 міста, "
          "фотофіксація цінників. Закриває єдиний канал, який не зчитується машинно.",
          "1 тиждень"),
         ("02", "Тестова закупка чотирьох лістингів", "Зафіксувати вагу, склад, маркування "
          "й якість — і зрозуміти, чому дві позиції з чотирьох стоять без залишку.",
          "2 тижні"),
         ("03", "Вивантаження Держмитслужби", "УКТЗЕД 1904 90 10 з розбивкою за країнами: "
          "покаже, хто вже возить і в яких обсягах.", "за запитом"),
         ("04", "Перемовини з однією мережею", "Новий продукт без аналога на полиці — "
          "питання не ціни, а готовності мережі відкрити позицію.", "паралельно")]
for i, (n, t1, t2, when) in enumerate(STEPS):
    y = 2.18 + i * 1.14
    rect(s, M, y, W - 2 * M, 1.00, CARD, line=LINE)
    rect(s, M, y, 0.05, 1.00, SIGNAL)
    text(s, M + 0.34, y + 0.24, 0.80, 0.52, n, size=24, bold=True, color=SAND, font=NUM)
    text(s, M + 1.30, y + 0.18, 4.20, 0.30, t1, size=13, bold=True, color=INK)
    text(s, M + 1.30, y + 0.52, 4.20, 0.30, "", size=8)
    text(s, M + 5.70, y + 0.18, 5.10, 0.70, t2, size=10, color=MUTED, line=1.26)
    text(s, W - M - 1.40, y + 0.34, 1.10, 0.30, when, size=10, bold=True, color=SIGNAL,
         align=PP_ALIGN.RIGHT)
foot(s, "Підсумок за матеріалами слайдів 02–13. Усі ціни та залишки зафіксовано 22.09.2026.")

prs.save("RTE_Rice_Ukraine_Shelf.pptx")
print("saved RTE_Rice_Ukraine_Shelf.pptx ·", len(prs.slides._sldIdLst), "slides")
