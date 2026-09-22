"""RTE Rice · Ukraine market study — deck built in the RTE_Rice_Suppliers design system.

Same page size, same two backgrounds (reused from the supplier deck), same grid:
eyebrow 0.62/0.30, title 0.62/0.60 Cambria 31, dek 0.64/1.22, panels from 1.68,
source footer at 0.62/6.95. Colours and type are lifted from that deck so the two
read as one document.
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

W, H = 13.3333, 7.5
MARGIN = 0.62

INK      = RGBColor(0x12, 0x45, 0x2B)   # dark panel
PANEL    = RGBColor(0x12, 0x45, 0x2B)
GROUND   = RGBColor(0x1A, 0x54, 0x33)
GOLD     = RGBColor(0xB0, 0x76, 0x2A)
GOLD_M   = RGBColor(0xC8, 0x9C, 0x40)
GOLD_L   = RGBColor(0xF5, 0xCE, 0x76)
AMBER    = RGBColor(0xF0, 0xAC, 0x58)
CREAM    = RGBColor(0xFC, 0xF9, 0xF0)
PALE     = RGBColor(0xD2, 0xE4, 0xCE)
FOOT     = RGBColor(0xC3, 0xD7, 0xBF)
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
GREEN    = RGBColor(0x35, 0x8A, 0x48)
RED      = RGBColor(0xB3, 0x31, 0x2D)

SERIF, SANS = "Cambria", "Calibri"
BG_TITLE = "un/ppt/media/image1.jpeg"
BG_PAGE  = "un/ppt/media/image2.jpeg"

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(W), Inches(H)
BLANK = prs.slide_layouts[6]


# ── primitives ────────────────────────────────────────────────────────────
def slide(bg=BG_PAGE):
    s = prs.slides.add_slide(BLANK)
    s.shapes.add_picture(bg, 0, 0, Inches(W), Inches(H))
    return s


def rect(s, x, y, w, h, fill, line=None, lw=1.0, rounded=False, adj=0.12):
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


def text(s, x, y, w, h, body, size=11, bold=False, italic=False, color=CREAM,
         font=SANS, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, line=None, space=None):
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
        if space and i:
            p.space_before = Pt(space)
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


def head(s, eyebrow, title, dek=None):
    text(s, MARGIN, 0.30, 9.0, 0.26, eyebrow, size=11, bold=True, color=AMBER)
    text(s, MARGIN, 0.60, 12.10, 0.62, title, size=31, bold=True, color=WHITE, font=SERIF)
    if dek:
        text(s, 0.64, 1.22, 12.10, 0.30, dek, size=12.5, color=PALE)


def source(s, txt):
    text(s, MARGIN, 6.95, 12.10, 0.30, txt, size=9, color=FOOT)


def panel(s, x, y, w, h, title=None):
    rect(s, x, y, w, h, PANEL)
    if title:
        text(s, x + 0.30, y + 0.22, w - 0.60, 0.34, title, size=17, bold=True,
             color=GOLD_L, font=SERIF)
    return y + (0.68 if title else 0.22)


def stat(s, x, y, w, big, unit, label, bar=True):
    if bar:
        rect(s, x, y, w, 0.53, PANEL)
    text(s, x + 0.20, y + 0.06, w * 0.42, 0.40,
         [[(big, {"size": 19, "bold": True}), (" " + unit, {"size": 11})]],
         color=GOLD_L, font=SERIF, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x + w * 0.44, y + 0.06, w * 0.54, 0.40, label, size=10.5, color=CREAM,
         anchor=MSO_ANCHOR.MIDDLE)


def kv(s, x, y, wl, wv, label, value, vcolor=CREAM, vbold=True):
    text(s, x, y, wl, 0.28, label, size=10.5, color=PALE)
    text(s, x + wl + 0.03, y, wv, 0.28, value, size=11.5, bold=vbold, color=vcolor)


def tag(s, x, y, txt, color=GOLD_L, w=None):
    w = w or (0.30 + 0.085 * len(txt))
    rect(s, x, y, w, 0.30, None, line=color, lw=1.0, rounded=True, adj=0.5)
    text(s, x, y, w, 0.30, txt, size=9.5, bold=True, color=color,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    return w


def num(v, d=1):
    return f"{v:,.{d}f}".replace(",", "\u00a0").replace(".", ",")


def rule(s, x, y, w, color=GOLD):
    rect(s, x, y, w, 0.013, color)


# ══ 1 · TITLE ═════════════════════════════════════════════════════════════
s = slide(BG_TITLE)
rect(s, 0.70, 1.75, 6.95, 4.05, PANEL)
text(s, 1.20, 2.20, 5.95, 0.30, "РИНОК УКРАЇНИ · ЗУСТРІЧНИЙ АНАЛІЗ", size=11,
     bold=True, color=AMBER)
text(s, 1.20, 2.62, 5.95, 0.80, "READY-TO-EAT РИС", size=38, bold=True,
     color=WHITE, font=SERIF)
rule(s, 1.20, 3.60, 2.00, GOLD_M)
text(s, 1.20, 3.80, 5.95, 0.44,
     "Хто вже продає його кінцевому споживачу — і за скільки",
     size=15, color=GOLD_L, font=SERIF)
text(s, 1.20, 4.40, 5.95, 0.72,
     "Ретортні та асептичні пауч / cup / лоток · розігрів 60–90 сек · "
     "зберігання за кімнатної температури 12–18 міс.",
     size=11.5, color=PALE, line=1.25)
rule(s, 1.20, 5.22, 5.95, RGBColor(0x2C, 0x6A, 0x45))
text(s, 1.20, 5.38, 3.70, 0.32, "Кабінетне дослідження · 22 вересня 2026 р.",
     size=11, color=CREAM)
text(s, 4.95, 5.38, 2.20, 0.32, "Morskyi Dim Group", size=11, bold=True,
     color=GOLD_L, align=PP_ALIGN.RIGHT)

# ══ 2 · METHOD ════════════════════════════════════════════════════════════
s = slide()
head(s, "МЕТОДОЛОГІЯ", "Що саме перевірено — і чого перевірити не вдалося",
     "Кожна цифра в колоді має джерело й дату. Де джерела немає — так і написано.")

y = panel(s, MARGIN, 1.68, 6.30, 3.30, "Охоплення перевірки")
for lb, vl in [("Дата зрізу цін", "22 вересня 2026 р."),
               ("Мережі", "17 мереж через API платформи zakaz.ua"),
               ("Окремо", "Сільпо — власний API мережі, київська філія"),
               ("Маркетплейси", "Prom.ua · Rozetka (лише індекс пошуку)"),
               ("Спец. канал", "6 азійських е-магазинів, ~5 500 позицій"),
               ("Статистика", "UN Comtrade, репортер Україна (804)")]:
    kv(s, MARGIN + 0.30, y, 1.75, 4.10, lb, vl)
    y += 0.365

y = panel(s, MARGIN, 5.14, 6.30, 1.60, "Три рівні даних")
for lb, vl, c in [("ПІДТВЕРДЖЕНО", "ціна й наявність зчитані з каталогу", GOLD_L),
                  ("ОЦІНКА", "наш розрахунок на підтверджених даних", AMBER),
                  ("ПРИПУЩЕННЯ", "позначено окремо, у висновки не входить", PALE)]:
    text(s, MARGIN + 0.30, y, 1.75, 0.26, lb, size=10, bold=True, color=c)
    text(s, MARGIN + 2.08, y, 4.10, 0.26, vl, size=10.5, color=CREAM)
    y += 0.30

y = panel(s, 7.20, 1.68, 5.52, 5.06, "Межі дослідження")
text(s, 7.50, y, 4.92, 1.30,
     "Rozetka, АТБ, Varus і Фора закриті від машинного зчитування "
     "(Cloudflare). Їхній індекс пошуку показує збіги за брендами, але "
     "перевірити назву, вагу й ціну неможливо.", size=11.5, color=CREAM, line=1.25)
text(s, 7.50, y + 1.42, 4.92, 0.90,
     "Лічильник збігів сам собою нічого не доводить: пошук у мережах "
     "нечіткий. У Сільпо запит «hetbahn» повертає 15 позицій — усі "
     "горілка «Гетьман».", size=11.5, color=CREAM, line=1.25)
rule(s, 7.50, y + 2.46, 4.92, RGBColor(0x2C, 0x6A, 0x45))
text(s, 7.50, y + 2.62, 4.92, 0.62,
     "Тому нуль у цій колоді — це нуль, отриманий переглядом назв товарів, "
     "а не відсутністю збігів.", size=11.5, bold=True, color=GOLD_L, line=1.25)
rect(s, 7.50, y + 3.36, 4.92, 0.90, GROUND)
text(s, 7.68, y + 3.48, 4.56, 0.66,
     ["ПОТРЕБУЄ РУЧНОЇ ПЕРЕВІРКИ В ТОЧКАХ ПРОДАЖУ:",
      "Rozetka · АТБ · Varus · Фора · Ашан офлайн"],
     size=10.5, bold=True, color=AMBER, line=1.30)
source(s, "Джерело: власні запити до відкритих API stores-api.zakaz.ua, "
          "sf-ecom-api.silpo.ua, search.rozetka.com.ua, comtradeapi.un.org та "
          "до карт сайтів спеціалізованих магазинів, 22.09.2026.")

# ══ 3 · HEADLINE ══════════════════════════════════════════════════════════
s = slide()
head(s, "ГОЛОВНИЙ ВИСНОВОК", "Категорії в українському роздробі не існує",
     "Це не «вільна ніша всередині ринку». Ринку немає — є порожня полиця.")

for i, (big, unit, lab, sub) in enumerate([
        ("0", "SKU", "у 18 мережах", "жодної позиції RTE-рису"),
        ("2", "SKU", "на ~5 500 позицій", "увесь спеціалізований азійський канал"),
        ("64,3", "₴", "за 100 г", "єдина зафіксована полична ціна")]):
    x = MARGIN + i * 4.13
    rect(s, x, 1.80, 3.90, 2.30, PANEL)
    text(s, x + 0.34, 2.06, 3.22, 0.90,
         [[(big, {"size": 54, "bold": True}), ("  " + unit, {"size": 17})]],
         color=GOLD_L, font=SERIF)
    rule(s, x + 0.34, 3.04, 1.40, GOLD_M)
    text(s, x + 0.34, 3.22, 3.22, 0.30, lab, size=13, bold=True, color=WHITE, font=SERIF)
    text(s, x + 0.34, 3.56, 3.22, 0.30, sub, size=10.5, color=PALE)

y = panel(s, MARGIN, 4.34, 12.10, 2.40, "Що це означає комерційно")
col = [(["Вхід — це створення категорії,", "а не відвоювання частки"],
        "Немає лідера, чию частку треба відбирати. Немає й попиту, "
        "вихованого чужим маркетингом: його доведеться створювати самим."),
       ("Формат і грамаж ми задаємо самі",
        "Жодний формат не є «звичним для полиці». Пауч 240 г і cup 150 г "
        "стартують з однакових позицій — обмежень з боку звички немає."),
       ("Єдиний орієнтир ціни — 64,3 ₴/100 г",
        "Корейський Hetbahn у нішевому онлайні. Наша собівартість пауча "
        "240 г — 16,7–18,8 ₴/100 г, тобто вчетверо нижча.")]
for i, (t1, t2) in enumerate(col):
    x = MARGIN + 0.30 + i * 3.90
    text(s, x, y, 3.55, 0.60, t1, size=12.5, bold=True, color=GOLD_L, font=SERIF, line=1.10)
    text(s, x, y + 0.70, 3.55, 1.10, t2, size=11, color=CREAM, line=1.22)
source(s, "Джерело: власні запити до каталогів мереж і спеціалізованих магазинів, "
          "22.09.2026; собівартість — колода «RTE Rice · Постачальники», курс 45,00 ₴/$.")

# ══ 4 · RETAIL MAP ════════════════════════════════════════════════════════
s = slide()
head(s, "КОНКУРЕНТНА КАРТА · 1/2", "Мультибрендовий роздріб: жодного SKU",
     "21 пошуковий запит у кожній мережі — українською, російською та англійською.")

rect(s, MARGIN, 1.68, 12.10, 0.40, PANEL)
for lx, lw, lt in [(0.30, 4.30, "МЕРЕЖА"), (4.70, 2.20, "RTE-РИС"),
                   (7.00, 5.00, "ЩО СТОЇТЬ НА ЦІЙ ПОЛИЦІ НАТОМІСТЬ")]:
    text(s, MARGIN + lx, 1.68, lw, 0.40, lt, size=9.5, bold=True, color=GOLD_L,
         anchor=MSO_ANCHOR.MIDDLE)

ROWS = [("METRO · Ашан · NOVUS", "0", "рис сухий, локшина б/п, рамьон, токпоккі"),
        ("МегаМаркет · Ultramarket", "0", "рис сухий, локшина б/п"),
        ("ЕКО маркет · Таврія В", "0", "рис сухий, локшина б/п, рамьон Ottogi"),
        ("Епіцентр FOOD · Космос", "0", "рис сухий, локшина б/п, токпоккі"),
        ("Восторг · Харків · Чудо Маркет", "0", "рис сухий, локшина б/п"),
        ("Гроно · Ідеал · ОнDe · Za Raz · Torba", "0", "рис сухий, локшина б/п, токпоккі"),
        ("Сільпо", "0", "рис сухий; охолоджені страви власної кухні")]
y = 2.14
for i, (a, b, c) in enumerate(ROWS):
    if i % 2 == 0:
        rect(s, MARGIN, y - 0.04, 12.10, 0.44, RGBColor(0x17, 0x4E, 0x31))
    text(s, MARGIN + 0.30, y, 4.30, 0.36, a, size=11.5, bold=True, color=CREAM,
         anchor=MSO_ANCHOR.MIDDLE)
    text(s, MARGIN + 4.70, y, 2.20, 0.36, b, size=15, bold=True, color=GOLD_L,
         font=SERIF, anchor=MSO_ANCHOR.MIDDLE)
    text(s, MARGIN + 7.00, y, 5.00, 0.36, c, size=11, color=PALE, anchor=MSO_ANCHOR.MIDDLE)
    y += 0.46

y = panel(s, MARGIN, 5.52, 12.10, 1.22)
text(s, MARGIN + 0.30, y, 11.50, 0.80,
     [[("Єдині збіги на запит «готовий рис» у всіх 17 мережах — корм для собак "
        "«Леопольд» 500 г. ", {"bold": True, "color": GOLD_L}),
       ("Позицій RTE-рису для людей не знайдено в жодній мережі. "
        "АТБ, Varus і Фора машинно не зчитуються — потребують ручної перевірки в залі.",
        {"color": CREAM})]], size=11.5, line=1.25)
source(s, "Джерело: stores-api.zakaz.ua (17 торгових мереж) і sf-ecom-api.silpo.ua, "
          "запити 22.09.2026; перевірено переглядом назв товарів, не лише лічильником збігів.")

# ══ 5 · SPECIALIST CHANNEL ════════════════════════════════════════════════
s = slide()
head(s, "КОНКУРЕНТНА КАРТА · 2/2", "Спеціалізований азійський канал: 2 позиції",
     "Саме тут категорія існує — у вигляді двох SKU на понад п'ять тисяч позицій.")

rect(s, MARGIN, 1.68, 12.10, 0.40, PANEL)
for lx, lw, lt in [(0.30, 3.60, "МАГАЗИН"), (4.00, 2.60, "ПОЗИЦІЙ У КАТАЛОЗІ"),
                   (6.80, 2.20, "RTE-РИС"), (9.10, 2.90, "ЩО САМЕ")]:
    text(s, MARGIN + lx, 1.68, lw, 0.40, lt, size=9.5, bold=True, color=GOLD_L,
         anchor=MSO_ANCHOR.MIDDLE)

SHOPS = [("asiafoods.com.ua", "1 984", "0", "—"),
         ("yaponskiy-kvartal.com", "1 624", "0", "—"),
         ("sushipovar.ua", "1 179", "0", "—"),
         ("samguk.com.ua", "285", "0", "—"),
         ("asia-goods.com.ua", "250", "1", "«Готовий рис з Кореї»"),
         ("smak-korea.com.ua", "141", "1", "CJ Hetbahn 210 г")]
y = 2.14
for i, (a, b, c, d) in enumerate(SHOPS):
    hit = c != "0"
    if i % 2 == 0:
        rect(s, MARGIN, y - 0.04, 12.10, 0.44, RGBColor(0x17, 0x4E, 0x31))
    if hit:
        rect(s, MARGIN, y - 0.04, 0.05, 0.44, GOLD_M)
    text(s, MARGIN + 0.30, y, 3.60, 0.36, a, size=11.5, bold=True,
         color=GOLD_L if hit else CREAM, anchor=MSO_ANCHOR.MIDDLE)
    text(s, MARGIN + 4.00, y, 2.60, 0.36, b, size=11.5, color=PALE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, MARGIN + 6.80, y, 2.20, 0.36, c, size=15, bold=True,
         color=GOLD_L if hit else PALE, font=SERIF, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE)
    text(s, MARGIN + 9.10, y, 2.90, 0.36, d, size=11, color=CREAM, anchor=MSO_ANCHOR.MIDDLE)
    y += 0.46

y = panel(s, MARGIN, 5.02, 12.10, 1.72, "Що це говорить про повід споживання")
text(s, MARGIN + 0.30, y, 5.60, 1.00,
     "Обидві позиції стоять у категорії «ЇЖА» поруч із рамьоном, токпоккі та "
     "кімчі — тобто продаються як корейська кухня вдома, а не як "
     "«швидкий обід» чи «їжа в дорогу».", size=11.5, color=CREAM, line=1.25)
text(s, MARGIN + 6.20, y, 5.60, 1.00,
     "Полиці «готовий гарнір за 90 секунд» в Україні не існує в жодному каналі. "
     "Це і ризик (повід треба пояснювати), і можливість — конкурента за цей "
     "повід немає.", size=11.5, color=CREAM, line=1.25)
source(s, "Джерело: карти сайтів і каталоги магазинів asiafoods.com.ua, "
          "yaponskiy-kvartal.com, sushipovar.ua, samguk.com.ua, asia-goods.com.ua, "
          "smak-korea.com.ua (WooCommerce Store API), 22.09.2026.")

# ══ 6 · THE TWO SKUs ══════════════════════════════════════════════════════
s = slide()
head(s, "ЗНАЙДЕНІ ПОЗИЦІЇ", "Усе, що фізично продається в Україні",
     "Дві позиції. Обидві — прямий імпорт з Кореї, без локалізації смаку та упаковки.")

CARDS = [
    ("ПІДТВЕРДЖЕНО", GOLD_L, ["CJ Hetbahn", "«Варений рис» 210 г"], "135,00 ₴",
     "64,3 ₴ за 100 г",
     [("Магазин", "smak-korea.com.ua"), ("Наявність", "у наявності"),
      ("Країна", "Республіка Корея"), ("Формат", "лоток-чаша, ambient"),
      ("Бренд", "CJ CheilJedang · Hetbahn"), ("Локалізація", "немає — прямий імпорт")]),
    ("ПІДТВЕРДЖЕНО", GOLD_L, ["«Готовий рис", "з Кореї», арт. 10296"], "190,00 ₴",
     "вага не вказана",
     [("Магазин", "asia-goods.com.ua"), ("Наявність", "у наявності"),
      ("Країна", "Республіка Корея"), ("Формат", "не вказано"),
      ("Бренд", "не вказано"), ("Локалізація", "немає — прямий імпорт")]),
    ("ПОТРЕБУЄ ПЕРЕВІРКИ", AMBER, ["Rozetka · АТБ", "Varus · Фора"], "н/д",
     "ціну зчитати неможливо",
     [("Причина", "Cloudflare блокує зчитування"), ("Індекс Rozetka", "збіги є, назви — ні"),
      ("Ризик", "лічильник збігів нечіткий"), ("Метод", "ручний моніторинг у залі"),
      ("Пріоритет", "АТБ і Сільпо офлайн"), ("Термін", "до фіксації умов із заводом")]),
]
for i, (badge, bc, name, price, per, rows) in enumerate(CARDS):
    x = MARGIN + i * 4.13
    rect(s, x, 1.68, 3.90, 5.06, PANEL)
    tag(s, x + 0.28, 1.92, badge, bc)
    text(s, x + 0.28, 2.42, 3.34, 0.70, name, size=15.5, bold=True, color=WHITE,
         font=SERIF, line=1.10)
    rule(s, x + 0.28, 3.26, 3.34, RGBColor(0x2C, 0x6A, 0x45))
    text(s, x + 0.28, 3.42, 3.34, 0.50, price, size=27, bold=True,
         color=GOLD_L if i < 2 else PALE, font=SERIF)
    text(s, x + 0.28, 3.96, 3.34, 0.26, per, size=11, color=AMBER)
    rule(s, x + 0.28, 4.32, 3.34, RGBColor(0x2C, 0x6A, 0x45))
    yy = 4.48
    for lb, vl in rows:
        text(s, x + 0.28, yy, 1.24, 0.26, lb, size=9.5, color=PALE)
        text(s, x + 1.56, yy, 2.06, 0.26, vl, size=9.5, bold=True, color=CREAM)
        yy += 0.345
source(s, "Джерело: smak-korea.com.ua (WooCommerce Store API) та asia-goods.com.ua "
          "(картка товару), зчитано 22.09.2026. Вага другої позиції на сайті не "
          "вказана — питома ціна не рахується.")

# ══ 7 · FORMATS & GRAMMAGE ════════════════════════════════════════════════
s = slide()
head(s, "ФОРМАТИ ТА ГРАМАЖ", "Домінуючого формату немає, бо немає категорії",
     "Порівняння того, що є на полиці України, з тим, що пропонують наші заводи.")

y = panel(s, MARGIN, 1.68, 5.86, 5.06, "На полиці України")
text(s, MARGIN + 0.30, y, 5.26, 0.58,
     "Один підтверджений формат на всю країну — корейський лоток-чаша 210 г.",
     size=12, color=CREAM, line=1.25)
rect(s, MARGIN + 0.30, y + 0.78, 5.26, 1.06, GROUND)
text(s, MARGIN + 0.54, y + 0.92, 1.30, 0.44, "210 г", size=24, bold=True,
     color=GOLD_L, font=SERIF)
text(s, MARGIN + 1.94, y + 0.98, 3.50, 0.34, "лоток-чаша · CJ Hetbahn",
     size=11.5, color=CREAM)
text(s, MARGIN + 1.94, y + 1.30, 3.50, 0.28, "єдиний формат, зафіксований в Україні",
     size=9.5, italic=True, color=PALE)
text(s, MARGIN + 0.30, y + 2.02, 5.26, 1.30,
     "Сегментації за грамажем — «порція на одного» проти «сімейного формату» — "
     "не існує. Для сегментації потрібні щонайменше дві лінійки одного бренду; "
     "в Україні немає жодної.", size=11.5, color=PALE, line=1.25)
rule(s, MARGIN + 0.30, y + 3.40, 5.26, RGBColor(0x2C, 0x6A, 0x45))
text(s, MARGIN + 0.30, y + 3.56, 5.26, 0.56,
     "Висновок: формат обираємо ми, а не звичка полиці.",
     size=12, bold=True, color=GOLD_L, font=SERIF, line=1.20)

y = panel(s, 6.86, 1.68, 5.86, 5.06, "У пропозиції наших постачальників")
FMT = [("Пауч 150 г", "BSCM · спецвиробництво", "33,25–38,24 ₴", "22,2–25,5"),
       ("Пауч 200 г", "BSCM · спецвиробництво", "38,07–43,18 ₴", "19,0–21,6"),
       ("Пауч 240 г", "BSCM · стандарт", "40,15–45,03 ₴", "16,7–18,8"),
       ("Пауч 250 г", "Chefrey · organic", "45,78–60,27 ₴", "18,3–24,1"),
       ("Cup 150 г", "BSCM · single cup", "34,95–39,99 ₴", "23,3–26,7"),
       ("Cup 2×125 г", "BSCM · private label", "64,51–73,51 ₴", "25,8–29,4"),
       ("Лоток 210 г", "KBROS · Корея", "40,27–43,81 ₴", "19,2–20,9")]
text(s, 7.16, y, 1.60, 0.24, "ФОРМАТ", size=9, bold=True, color=GOLD_L)
text(s, 8.80, y, 1.90, 0.24, "ПОСТАЧАЛЬНИК", size=9, bold=True, color=GOLD_L)
text(s, 10.74, y, 1.02, 0.24, "₴/ОД.", size=9, bold=True, color=GOLD_L,
     align=PP_ALIGN.RIGHT)
text(s, 11.80, y, 0.72, 0.24, "₴/100Г", size=9, bold=True, color=GOLD_L,
     align=PP_ALIGN.RIGHT)
yy = y + 0.34
for a, b, c, dd in FMT:
    text(s, 7.16, yy, 1.60, 0.28, a, size=11, bold=True, color=CREAM)
    text(s, 8.80, yy, 1.90, 0.28, b, size=10, color=PALE)
    text(s, 10.60, yy, 1.16, 0.28, c, size=10.5, color=CREAM, align=PP_ALIGN.RIGHT)
    text(s, 11.70, yy, 0.82, 0.28, dd, size=10.5, bold=True, color=GOLD_L,
         align=PP_ALIGN.RIGHT)
    yy += 0.40
rule(s, 7.16, yy + 0.04, 5.26, RGBColor(0x2C, 0x6A, 0x45))
text(s, 7.16, yy + 0.22, 5.26, 0.56,
     "Діапазон собівартості — від завантаження 40′ FCL до збірного 17 м³.",
     size=10.5, color=PALE, line=1.20)
source(s, "Джерело: полиця — smak-korea.com.ua, 22.09.2026; собівартість — колода "
          "«RTE Rice · Постачальники» (Self Cost Rice Brand.xlsx), курс 45,00 ₴/$. "
          "₴/100 г — наш перерахунок.")

# ══ 8 · WHY NO MEDIAN ═════════════════════════════════════════════════════
s = slide()
head(s, "РОЗДРІБНІ ЦІНИ", "Чому тут немає «середньої ціни по ринку»",
     "Вибірка з двох позицій не дає ні медіани, ні розкиду. Робити вигляд, що дає — обман.")

y = panel(s, MARGIN, 1.68, 6.30, 2.74, "Що ми маємо насправді")
for lb, vl, c in [("Спостережень", "n = 2", GOLD_L),
                  ("З відомою вагою", "n = 1", GOLD_L),
                  ("Медіана", "не рахується", AMBER),
                  ("Мін / макс", "не рахується", AMBER),
                  ("Єдина питома ціна", "64,3 ₴ за 100 г", GOLD_L)]:
    kv(s, MARGIN + 0.30, y, 2.60, 3.30, lb, vl, vcolor=c)
    y += 0.32
rule(s, MARGIN + 0.30, y + 0.02, 5.70, RGBColor(0x2C, 0x6A, 0x45))
text(s, MARGIN + 0.30, y + 0.14, 5.70, 0.28,
     "Одна точка — це орієнтир, а не ринкова ціна.", size=11.5, bold=True, color=CREAM)

y = panel(s, MARGIN, 4.62, 6.30, 2.12, "Як закрити цю прогалину")
for n, t in [("1", "Ручний моніторинг полиці: АТБ, Сільпо, Novus, Ашан, METRO — "
                   "фотофіксація цінників, 2 міста, 1 тиждень."),
             ("2", "Запит у Rozetka Marketplace і Prom.ua на вивантаження "
                   "категорії з цінами продавців."),
             ("3", "Тестова закупка обох корейських позицій — щоб зафіксувати "
                   "вагу, склад і реальне маркування.")]:
    text(s, MARGIN + 0.30, y, 0.30, 0.26, n, size=13, bold=True, color=GOLD_L, font=SERIF)
    text(s, MARGIN + 0.66, y, 5.34, 0.46, t, size=10.5, color=CREAM, line=1.18)
    y += 0.46

y = panel(s, 7.20, 1.68, 5.52, 5.06, "Тому шкалу задають субститути")
text(s, 7.50, y, 4.92, 1.10,
     "Ціну RTE-рису в Україні поки що можна оцінити лише через сусідні полиці — "
     "те, чим споживач фактично закриває той самий прийом їжі.",
     size=11.5, color=CREAM, line=1.25)
BANDS = [("Рис сухий", "410", "3,9", "16,5", "206"),
         ("Каші б/п", "102", "11,8", "48,9", "398"),
         ("Локшина б/п", "478", "14,5", "48,9", "244"),
         ("Готова їжа з рисом, охол.", "11", "35,6", "52", "66"),
         ("RTE-рис", "1", "—", "64,3", "—"),
         ("Токпоккі", "36", "90,3", "144", "235")]
yy = y + 1.32
text(s, 7.50, yy, 2.40, 0.24, "КАТЕГОРІЯ", size=9, bold=True, color=GOLD_L)
for lx, lt in [(9.96, "N"), (10.72, "МІН"), (11.48, "МЕД"), (12.14, "МАКС")]:
    text(s, lx, yy, 0.62, 0.24, lt, size=9, bold=True, color=GOLD_L, align=PP_ALIGN.RIGHT)
yy += 0.32
for i, (a, n, lo, me, hi) in enumerate(BANDS):
    hl = a == "RTE-рис"
    if hl:
        rect(s, 7.42, yy - 0.05, 5.08, 0.40, RGBColor(0x2C, 0x6A, 0x45))
    text(s, 7.50, yy, 2.46, 0.30, a, size=10.5, bold=hl, color=GOLD_L if hl else CREAM)
    for lx, v in [(9.96, n), (10.72, lo), (11.48, me), (12.14, hi)]:
        text(s, lx, yy, 0.62, 0.30, v, size=10.5, bold=(hl or lx == 11.48),
             color=GOLD_L if hl else (CREAM if lx == 11.48 else PALE),
             align=PP_ALIGN.RIGHT)
    yy += 0.39
text(s, 7.50, yy + 0.06, 4.92, 0.28, "усі значення — ₴ за 100 г, 22.09.2026",
     size=9.5, italic=True, color=FOOT)
source(s, "Джерело: розрахунок за каталогами 17 мереж (stores-api.zakaz.ua) і Сільпо, "
          "22.09.2026. N — кількість SKU у вибірці; питома ціна рахується лише для "
          "позицій із вагою в назві.")

# ══ 9 · SUBSTITUTES ═══════════════════════════════════════════════════════
s = slide()
head(s, "СУБСТИТУТИ", "Чим українець фактично закриває цей прийом їжі",
     "Охолоджена готова їжа власної кухні мережі — головний конкурент, а не локшина.")

y = panel(s, MARGIN, 1.68, 6.72, 3.70, "Сільпо · власна кухня, охолоджені страви з рисом")
ROWS9 = [("Рис з яйцем та овочами «Чіл Міл»", "180 г", "63,99 ₴", "35,6"),
         ("Боул: рис зі смаженим курячим філе", "330 г", "139,00 ₴", "42,1"),
         ("Смажений рис з куркою в теріякі", "350 г", "169,00 ₴", "48,3"),
         ("Рис з овочами та курячим битком", "300 г", "149,00 ₴", "49,7"),
         ("Рис з овочами та яйцем у кунжуті", "160 г", "88,99 ₴", "55,6"),
         ("Карі з рисом та куркою", "350 г", "219,00 ₴", "62,6")]
text(s, MARGIN + 0.30, y, 3.20, 0.24, "СТРАВА", size=9, bold=True, color=GOLD_L)
for lx, lw, lt in [(3.66, 0.72, "ВАГА"), (4.52, 1.00, "ЦІНА"), (5.66, 0.76, "₴/100 Г")]:
    text(s, MARGIN + lx, y, lw, 0.24, lt, size=9, bold=True, color=GOLD_L,
         align=PP_ALIGN.RIGHT)
yy = y + 0.34
for a, b, c, dd in ROWS9:
    text(s, MARGIN + 0.30, yy, 3.30, 0.30, a, size=10.5, color=CREAM)
    text(s, MARGIN + 3.66, yy, 0.72, 0.30, b, size=10.5, color=PALE, align=PP_ALIGN.RIGHT)
    text(s, MARGIN + 4.52, yy, 1.00, 0.30, c, size=10.5, color=CREAM, align=PP_ALIGN.RIGHT)
    text(s, MARGIN + 5.66, yy, 0.76, 0.30, dd, size=11, bold=True, color=GOLD_L,
         align=PP_ALIGN.RIGHT)
    yy += 0.42

y = panel(s, MARGIN, 5.54, 6.72, 1.20)
text(s, MARGIN + 0.30, y, 6.12, 0.80,
     [[("Діапазон 35,6 – 66,4 ₴/100 г. ", {"bold": True, "color": GOLD_L}),
       ("Hetbahn за 64,3 ₴/100 г коштує як повноцінна страва з куркою — "
        "хоча це просто рис. Ось де наш ціновий важіль.", {"color": CREAM})]],
     size=11.5, line=1.25)

y = panel(s, 7.62, 1.68, 5.10, 5.06, "Чи відбирають вони попит")
BLOCKS = [("Локшина б/п · рамьон", "48,9 ₴/100 г",
           "Окрема ніша. Це самостійна страва з бульйоном, а не гарнір. "
           "Попит перетинається лише в поводі «швидко й гаряче»."),
          ("Каші б/п", "48,9 ₴/100 г",
           "Сніданковий повід, солодкий профіль. З RTE-рисом майже не конкурує."),
          ("Рис сухий, зокрема 5×80 г", "16,5 ₴/100 г",
           "Головний ціновий якір у голові споживача: «навіщо платити вчетверо». "
           "Саме з цим доведеться сперечатися комунікацією."),
          ("Охолоджена готова їжа", "35,6–66,4 ₴/100 г",
           "Прямий конкурент за повід і за ціну. Але потребує холодильника "
           "й живе 2–5 діб — тут наша перевага.")]
yy = y
for t1, t2, t3 in BLOCKS:
    text(s, 7.92, yy, 3.10, 0.26, t1, size=11, bold=True, color=CREAM)
    text(s, 11.02, yy, 1.42, 0.26, t2, size=10, bold=True, color=GOLD_L,
         align=PP_ALIGN.RIGHT)
    text(s, 7.92, yy + 0.28, 4.52, 0.70, t3, size=10, color=PALE, line=1.20)
    yy += 1.16
source(s, "Джерело: sf-ecom-api.silpo.ua (київська філія) та stores-api.zakaz.ua, "
          "22.09.2026. ₴/100 г — наш перерахунок; вагові позиції з прилавка виключено.")

# ══ 10 · CHANNELS ═════════════════════════════════════════════════════════
s = slide()
head(s, "ТОЧКИ ПРОДАЖУ", "Де категорія живе сьогодні — і де її немає",
     "Розподіл по каналах прямо показує, який повід споживання зараз обслуговується.")

CH = [("Продуктові мережі", "0", "METRO · Ашан · NOVUS · МегаМаркет · Ultramarket · "
       "ЕКО маркет · Таврія В · Епіцентр FOOD · Сільпо та ще 9 регіональних мереж", PALE),
      ("Маркетплейси", "?", "Prom.ua на запит «готовий рис пауч» дає лише корм для "
       "тварин. Rozetka: індекс показує збіги, картки закриті — ручна перевірка.", AMBER),
      ("Спеціалізований азійський онлайн", "2", "smak-korea.com.ua (CJ Hetbahn 210 г, "
       "135 ₴) · asia-goods.com.ua («Готовий рис з Кореї», 190 ₴)", GOLD_L),
      ("Мережі поза зчитуванням", "?", "АТБ · Varus · Фора — Cloudflare; потребують "
       "ручного моніторингу полиці", AMBER)]
y = 1.72
for a, b, c, col in CH:
    rect(s, MARGIN, y, 12.10, 1.06, PANEL)
    rect(s, MARGIN, y, 0.06, 1.06, col)
    text(s, MARGIN + 0.34, y + 0.16, 4.10, 0.32, a, size=14, bold=True, color=WHITE,
         font=SERIF)
    text(s, MARGIN + 0.34, y + 0.56, 4.10, 0.30, "знайдено позицій RTE-рису", size=9.5,
         color=PALE)
    text(s, MARGIN + 4.50, y + 0.20, 0.90, 0.58, b, size=30, bold=True, color=col,
         font=SERIF, align=PP_ALIGN.CENTER)
    text(s, MARGIN + 5.60, y + 0.18, 6.20, 0.74, c, size=10.5, color=CREAM, line=1.22)
    y += 1.14

y = panel(s, MARGIN, 6.30, 12.10, 0.44)
text(s, MARGIN + 0.30, y - 0.08, 11.50, 0.32,
     [[("Сезонність: ", {"bold": True, "color": GOLD_L}),
       ("оцінити неможливо — для цього потрібен ряд спостережень за рік, "
        "а категорії на полиці немає.", {"color": CREAM})]], size=11)
source(s, "Джерело: stores-api.zakaz.ua, sf-ecom-api.silpo.ua, prom.ua, "
          "search.rozetka.com.ua, каталоги спеціалізованих магазинів — 22.09.2026.")

# ══ 11 · IMPORT STATS ═════════════════════════════════════════════════════
s = slide()
head(s, "ІМПОРТНА СТАТИСТИКА", "HS 1904 90 — найближчий публічний код",
     "Прямого коду саме для RTE-рису не існує; це кошиковий код готових круп'яних виробів.")

y = panel(s, MARGIN, 1.68, 7.10, 5.06, "Імпорт України, HS 1904 90")
text(s, MARGIN + 0.30, y, 1.30, 0.24, "РІК", size=9, bold=True, color=GOLD_L)
for lx, lt in [(2.40, "ТОНН"), (4.20, "ТИС. $ CIF"), (6.00, "$/КГ")]:
    text(s, MARGIN + lx, y, 1.30, 0.24, lt, size=9, bold=True, color=GOLD_L,
         align=PP_ALIGN.RIGHT)
DATA = [("2018", 956.6, 839.7), ("2019", 17.6, 74.6), ("2020", 1.2, 12.8),
        ("2021", 0.4, 2.3), ("2022", 362.2, 534.4), ("2023", 287.5, 706.7),
        ("2024", 207.2, 608.7)]
yy = y + 0.34
for yr, t, v in DATA:
    gap = yr in ("2019", "2020", "2021")
    text(s, MARGIN + 0.30, yy, 1.30, 0.30, yr, size=11.5, bold=True,
         color=PALE if gap else CREAM)
    text(s, MARGIN + 1.80, yy, 1.90, 0.30, num(t), size=11.5,
         color=PALE if gap else CREAM, align=PP_ALIGN.RIGHT)
    text(s, MARGIN + 3.60, yy, 1.90, 0.30, num(v), size=11.5,
         color=PALE if gap else CREAM, align=PP_ALIGN.RIGHT)
    text(s, MARGIN + 5.40, yy, 1.90, 0.30, num(v / t, 2), size=11.5,
         bold=True, color=PALE if gap else GOLD_L, align=PP_ALIGN.RIGHT)
    yy += 0.40
rule(s, MARGIN + 0.30, yy + 0.06, 6.50, RGBColor(0x2C, 0x6A, 0x45))
text(s, MARGIN + 0.30, yy + 0.24, 6.50, 0.60,
     "2019–2021 — майже нульові значення. Це схоже на прогалину у звітності, "
     "а не на реальне зникнення імпорту. Ряд читати обережно.",
     size=10.5, italic=True, color=AMBER, line=1.22)

y = panel(s, 7.92, 1.68, 4.80, 5.06, "Межі цих даних")
for t in ["HS 1904 90 — «готові харчові вироби з зерна, інші». Готовий рис — "
          "лише частина коду; сюди ж потрапляють інші крупи.",
          "Публічний доступ UN Comtrade віддає лише світовий агрегат. Розбивки "
          "за країнами-партнерами немає — тож частку Таїланду та Кореї "
          "з цього джерела не видно.",
          "Дані за 2025 рік на дату зрізу відсутні."]:
    text(s, 8.22, y, 4.20, 1.04, "— " + t, size=11, color=CREAM, line=1.22)
    y += 1.10
rect(s, 8.22, y + 0.10, 4.20, 0.92, GROUND)
text(s, 8.42, y + 0.22, 3.80, 0.72,
     "Для точної картини потрібне вивантаження Держмитслужби за УКТЗЕД "
     "1904 90 10 з розбивкою за країнами — це платний / запитовий ресурс.",
     size=10.5, bold=True, color=AMBER, line=1.22)
source(s, "Джерело: UN Comtrade (comtradeapi.un.org), репортер Україна (804), "
          "потік «імпорт», код HS 190490, звернення 22.09.2026. $/кг — наш розрахунок.")

# ══ 12 · REGULATORY ═══════════════════════════════════════════════════════
s = slide()
head(s, "РЕГУЛЯТОРИКА", "Що потрібно для введення продукту на ринок",
     "Оглядово, без глибокого юридичного аналізу — з позначенням зон ризику.")

y = panel(s, MARGIN, 1.68, 7.10, 5.06, "Базові вимоги")
REG = [("Маркування", "Українською мовою — ЗУ «Про інформацію для споживачів "
        "щодо харчових продуктів» № 2639-VIII"),
       ("Обов'язкові дані", "Склад, поживна цінність, алергени, термін "
        "придатності, партія, країна походження, виробник та імпортер"),
       ("Оператор ринку", "Реєстрація потужностей імпортера в "
        "Держпродспоживслужбі; система НАССР"),
       ("Походження", "Рецептура рослинного походження — ветеринарний "
        "сертифікат не потрібен. Він з'являється, щойно в складі є "
        "компонент тваринного походження"),
       ("Митні платежі", "Мито 10 % і ПДВ 20 % — уже враховані в розрахунку "
        "собівартості")]
for lb, vl in REG:
    text(s, MARGIN + 0.30, y, 1.90, 0.26, lb, size=10.5, bold=True, color=GOLD_L)
    text(s, MARGIN + 2.30, y, 4.50, 0.80, vl, size=10.5, color=CREAM, line=1.22)
    y += 0.94

y = panel(s, 7.92, 1.68, 4.80, 5.06, "Зони ризику")
RISK = [("Рецептура з м'ясом чи яйцем",
         "Смаки з куркою, яйцем або морепродуктами переводять товар у режим "
         "продукції тваринного походження — інший пакет документів і "
         "затверджені потужності експортера."),
        ("Воєнні спрощення маркування",
         "У березні 2022 року Кабмін дозволив супровідну документацію "
         "українською замість друку на упаковці. Чи діє це у 2026 році — "
         "потребує юридичної перевірки."),
        ("Заявки на упаковці",
         "«Без консервантів», «organic», «без глютену» вимагають "
         "підтвердження; organic — ще й визнаної сертифікації.")]
for t1, t2 in RISK:
    rect(s, 8.22, y - 0.04, 0.05, 1.28, RED)
    text(s, 8.42, y, 3.90, 0.26, t1, size=11, bold=True, color=AMBER)
    text(s, 8.42, y + 0.28, 3.90, 0.96, t2, size=10, color=CREAM, line=1.22)
    y += 1.44
source(s, "Джерело: ЗУ № 2639-VIII (zakon.rada.gov.ua/go/2639-19); повідомлення "
          "Кабінету Міністрів України від 21.03.2022 про спрощення маркування "
          "імпортних продуктів. Станом на 22.09.2026 чинність спрощень не "
          "перевірялася — потрібен висновок юриста.")

# ══ 13 · COST VS SHELF ════════════════════════════════════════════════════
s = slide()
head(s, "СОБІВАРТІСТЬ ПРОТИ ПОЛИЦІ", "Де саме відкривається вікно",
     "Єдиний орієнтир ціни — 64,3 ₴/100 г. Наша собівартість — учетверо нижча.")

y = panel(s, MARGIN, 1.68, 5.60, 2.60, "Відправна точка")
for lb, vl, c in [("Hetbahn 210 г, полиця", "135,00 ₴", GOLD_L),
                  ("те саме за 100 г", "64,3 ₴", GOLD_L),
                  ("наш пауч 240 г, 40′ FCL", "40,15 ₴", CREAM),
                  ("те саме за 100 г", "16,7 ₴", GOLD_L),
                  ("розрив", "×3,8", AMBER)]:
    kv(s, MARGIN + 0.30, y, 2.90, 2.30, lb, vl, vcolor=c)
    y += 0.38

y = panel(s, MARGIN, 4.44, 5.60, 2.30, "Наш пауч 240 г коштує менше, ніж")
for a, b in [("Охолоджений боул Сільпо 330 г", "139,00 ₴"),
             ("Локшина Samyang Buldak 70 г", "137,00 ₴"),
             ("Токпоккі Yopokki 140 г", "≈160 ₴"),
             ("Рис сухий 5×80 г, Novus", "38,59 ₴")]:
    text(s, MARGIN + 0.30, y, 3.60, 0.28, a, size=10.5, color=CREAM)
    text(s, MARGIN + 3.90, y, 1.40, 0.28, b, size=10.5, bold=True, color=GOLD_L,
         align=PP_ALIGN.RIGHT)
    y += 0.36
text(s, MARGIN + 0.30, y - 0.02, 5.00, 0.28,
     "Сухий рис лишається якорем «дешево» — і це головне заперечення.",
     size=10, italic=True, color=PALE)

y = panel(s, 6.52, 1.68, 6.20, 5.06,
          "Цінові сценарії · пауч 240 г, 40′ FCL   ·   ОЦІНКА")
text(s, 6.82, y, 1.30, 0.24, "РЕКОМ. ЦІНА", size=9, bold=True, color=GOLD_L)
for lx, lw, lt in [(8.30, 1.05, "₴/100 Г"), (9.55, 1.25, "vs HETBAHN"),
                   (10.95, 1.00, "МАРЖА"), (12.00, 0.55, "%")]:
    text(s, lx, y, lw, 0.24, lt, size=9, bold=True, color=GOLD_L, align=PP_ALIGN.RIGHT)
SCEN = [("89 ₴", "37,1", "−42 %", "14,79 ₴", "27 %", False),
        ("99 ₴", "41,2", "−36 %", "20,96 ₴", "34 %", False),
        ("109 ₴", "45,4", "−29 %", "27,13 ₴", "40 %", True),
        ("119 ₴", "49,6", "−23 %", "33,31 ₴", "45 %", True),
        ("129 ₴", "53,8", "−16 %", "39,48 ₴", "50 %", False)]
yy = y + 0.36
for a, b, c, dd, e, hl in SCEN:
    if hl:
        rect(s, 6.72, yy - 0.05, 5.80, 0.42, RGBColor(0x2C, 0x6A, 0x45))
    text(s, 6.82, yy, 1.30, 0.32, a, size=13, bold=True, color=GOLD_L if hl else CREAM,
         font=SERIF)
    for lx, lw, v in [(8.30, 1.05, b), (9.55, 1.25, c), (10.95, 1.00, dd), (12.00, 0.55, e)]:
        text(s, lx, yy, lw, 0.32, v, size=11, bold=hl,
             color=GOLD_L if hl else PALE, align=PP_ALIGN.RIGHT)
    yy += 0.46
rule(s, 6.82, yy + 0.06, 5.70, RGBColor(0x2C, 0x6A, 0x45))
text(s, 6.82, yy + 0.24, 5.70, 1.30,
     "Розрахунок: роздрібна ціна без ПДВ, націнка мережі 35 %, маржа — різниця "
     "між ціною відвантаження та собівартістю 40,15 ₴. Це модель для обговорення, "
     "а не комерційна пропозиція: реальна націнка мереж і умови входу "
     "(listing fee, промо-бюджет, відстрочка) у розрахунку не враховані.",
     size=10.5, italic=True, color=AMBER, line=1.24)
source(s, "Джерело: собівартість — колода «RTE Rice · Постачальники» (BSCM пауч 240 г, "
          "40′ FCL, курс 45,00 ₴/$); ціна Hetbahn — smak-korea.com.ua, 22.09.2026. "
          "Сценарії — наш розрахунок.")

# ══ 14 · VERDICT ══════════════════════════════════════════════════════════
s = slide()
head(s, "ВИСНОВОК", "Ніша вільна — але це не готовий попит",
     "Пряма відповідь на питання, поставлене на початку дослідження.")

y = panel(s, MARGIN, 1.68, 6.30, 3.00, "Так, ніша вільна")
text(s, MARGIN + 0.30, y, 5.70, 1.30,
     "У жодній із 18 перевірених мереж немає RTE-рису. У чотирьох із шести "
     "перевірених азійських е-магазинів — теж. Вхід не потребує витіснення "
     "конкурента: полиця порожня, і власна марка може стати першою.",
     size=12, color=CREAM, line=1.28)
rule(s, MARGIN + 0.30, y + 1.56, 5.70, GOLD_M)
text(s, MARGIN + 0.30, y + 1.72, 5.70, 0.32,
     "Диференціація за ціною не потрібна — немає від кого.",
     size=11.5, bold=True, color=GOLD_L)

y = panel(s, MARGIN, 4.86, 6.30, 1.88, "Але порожня полиця — не те саме, що попит")
text(s, MARGIN + 0.30, y, 5.70, 1.10,
     "Категорії немає тому, що ніхто її не пробував, — або тому, що спроби не "
     "окупилися. Відрізнити одне від одного кабінетно неможливо. Тому головна "
     "інвестиція на старті — не ціна, а пояснення поводу споживання.",
     size=11.5, color=CREAM, line=1.26)

y = panel(s, 7.20, 1.68, 5.52, 5.06, "Що зробити до рішення про закупівлю")
STEPS = [("1", "Ручний моніторинг полиці",
          "АТБ, Сільпо, Novus, Ашан, METRO — 2 міста, фотофіксація цінників. "
          "Закриває єдину прогалину цього дослідження."),
         ("2", "Тестова закупка двох корейських SKU",
          "Зафіксувати вагу, склад, маркування й реальну якість — щоб мати "
          "з чим порівнювати зразки заводів."),
         ("3", "Вивантаження Держмитслужби за УКТЗЕД 1904 90 10",
          "Розбивка за країнами покаже, чи хтось уже возить і в яких обсягах."),
         ("4", "Перемовини з мережею про вхід у категорію",
          "Новий продукт без аналога на полиці — це питання не ціни, "
          "а готовності мережі відкрити позицію.")]
yy = y
for n, t1, t2 in STEPS:
    text(s, 7.50, yy, 0.34, 0.30, n, size=17, bold=True, color=GOLD_L, font=SERIF)
    text(s, 7.92, yy, 4.50, 0.28, t1, size=11.5, bold=True, color=WHITE)
    text(s, 7.92, yy + 0.30, 4.50, 0.76, t2, size=10, color=PALE, line=1.22)
    yy += 1.14
source(s, "Підсумок за матеріалами слайдів 2–13. Усі полиці й ціни зафіксовано "
          "22.09.2026; дані з позначкою «потребує ручної перевірки» у висновок не включені.")

prs.save("RTE_Rice_Ukraine_Market.pptx")
print("saved RTE_Rice_Ukraine_Market.pptx ·", len(prs.slides._sldIdLst), "slides")
