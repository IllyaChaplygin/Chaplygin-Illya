# -*- coding: utf-8 -*-
"""Build the suppliers deck: supplier -> product -> short description -> self-cost."""
import os
import sys

from pptx import Presentation
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches

sys.path.insert(0, os.path.dirname(__file__))
from catalog import BASE, DATA, KORIKO, SCENARIOS, SUPPLIERS, photo  # noqa: E402
from theme import (BAR_H, BODY, BRIGHT, DARK, GREEN, HEAD, M, MINT, PAGE, SH, SOFT,  # noqa: E402
                   SW, TEXT, WHITE, blank, fmt_uah, fmt_usd, header, picture, pill,
                   rect, text)

OUT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..',
                                   'Постачальники_снеків_собівартість.pptx'))

# scenarios in ascending cost order, with the short labels used inside cards
SC_ORDER = [(BASE, '40′ контейнер', '40′ конт.'),
            ("20'", '20′ контейнер', '20′ конт.'),
            ('LCL 34', 'Збірний 34 м³', 'Збірн. 34 м³'),
            ('LCL 17', 'Збірний 17 м³', 'Збірн. 17 м³')]

CARD_TOP = 1.42
CARD_BOTTOM = 5.34


def costs(sheet, key):
    for row in DATA[sheet]:
        if row['name'] == key:
            return row['cost']
    raise KeyError('%s / %s' % (sheet, key))


# =================================================================== cost block ===
def cost_rows(s, x, y, w, h, cost, size=6.5, label_size=6.5):
    """Vertical read-out of every logistics scenario; the cheapest one is highlighted."""
    rect(s, x, y, w, h, fill=MINT)
    text(s, x + 0.10, y + 0.05, w - 0.20, 0.16, 'СОБІВАРТІСТЬ ЗА ОДИНИЦЮ',
         size=label_size, color=SOFT, bold=True)
    rh = (h - 0.26) / len(SC_ORDER)
    yy = y + 0.24
    for key, _, short in SC_ORDER:
        v = cost[key]
        hot = key == BASE
        if hot:
            rect(s, x + 0.06, yy, w - 0.12, rh, fill=GREEN)
        col = WHITE if hot else SOFT
        text(s, x + 0.14, yy, w * 0.36 - 0.14, rh, short, size=size, color=col,
             bold=hot, anchor=MSO_ANCHOR.MIDDLE)
        text(s, x + w * 0.36, yy, w * 0.30, rh, fmt_usd(v['usd']), size=size,
             color=col, bold=True, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
        text(s, x + w * 0.66, yy, w * 0.34 - 0.14, rh, fmt_uah(v['uah']), size=size,
             color=col, bold=True, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
        yy += rh


def cost_cells(s, x, y, w, h, cost):
    """Same read-out laid out horizontally, for the roomier two-up cards."""
    rect(s, x, y, w, h, fill=MINT)
    text(s, x + 0.14, y + 0.06, w - 0.28, 0.18, 'СОБІВАРТІСТЬ ЗА ОДИНИЦЮ',
         size=7.5, color=SOFT, bold=True)
    cw = (w - 0.28) / len(SC_ORDER)
    cy = y + 0.26
    ch = h - 0.34
    for i, (key, label, _) in enumerate(SC_ORDER):
        v = cost[key]
        cx = x + 0.14 + i * cw
        hot = key == BASE
        if hot:
            rect(s, cx, cy, cw, ch, fill=GREEN)
        col = WHITE if hot else SOFT
        text(s, cx, cy + 0.07, cw, 0.16, label, size=7, color=col, bold=hot,
             align=PP_ALIGN.CENTER)
        text(s, cx, cy + 0.25, cw, 0.26, fmt_usd(v['usd']), size=12, font=HEAD,
             bold=True, color=WHITE if hot else DARK, align=PP_ALIGN.CENTER)
        text(s, cx, cy + 0.52, cw, 0.20, fmt_uah(v['uah']), size=9, bold=True,
             color=col, align=PP_ALIGN.CENTER)


# ======================================================================= cards ===
def card_grid(s, x, w, p, cost):
    y, h = CARD_TOP, CARD_BOTTOM - CARD_TOP
    rect(s, x, y, w, h, fill=WHITE, shadow=True)
    picture(s, photo(p['photo']), x + 0.14, y + 0.12, w - 0.28, 1.16)

    bw = min(w - 0.24, 0.058 * len(p['badge']) + 0.30)
    pill(s, x + (w - bw) / 2, y + 1.38, bw, 0.26, p['badge'], size=7)
    text(s, x + 0.16, y + 1.72, w - 0.32, 0.44, p['title'], size=11.5, font=HEAD,
         bold=True, color=DARK, line_spacing=0.98)
    text(s, x + 0.16, y + 2.20, w - 0.32, 0.62, p['desc'], size=8, color=TEXT,
         line_spacing=1.18)
    cost_rows(s, x + 0.12, y + h - 1.10, w - 0.24, 0.98, cost)


def card_duo(s, x, w, p, cost):
    y, h = CARD_TOP, CARD_BOTTOM - CARD_TOP
    rect(s, x, y, w, h, fill=WHITE, shadow=True)
    picture(s, photo(p['photo']), x + 0.24, y + 0.14, w - 0.48, 1.24)

    bw = 0.075 * len(p['badge']) + 0.40
    pill(s, x + (w - bw) / 2, y + 1.48, bw, 0.30, p['badge'], size=9)
    text(s, x + 0.30, y + 1.86, w - 0.60, 0.42, p['title'].replace('\n', ' '),
         size=15, font=HEAD, bold=True, color=DARK, align=PP_ALIGN.CENTER,
         line_spacing=1.0)
    text(s, x + 0.42, y + 2.32, w - 0.84, 0.42, p['desc'], size=9, color=TEXT,
         align=PP_ALIGN.CENTER, line_spacing=1.2)
    cost_cells(s, x + 0.24, y + 2.82, w - 0.48, 1.00, cost)


# ====================================================================== slides ===
def slide_cover(prs):
    s = blank(prs)
    rect(s, 0, 0, SW, SH, fill=DARK)
    rect(s, 0, 0, 3.55, SH, fill=GREEN)
    for i, name in enumerate(['zek_tempura_corn30', 'sk_original',
                              'tmk_roll_orig', 'zek_topping_veg35']):
        col, row = i % 2, i // 2
        cx, cy = 0.30 + col * 1.55, 0.84 + row * 2.05
        rect(s, cx, cy, 1.40, 1.90, fill=WHITE)
        picture(s, photo(name), cx + 0.08, cy + 0.08, 1.24, 1.74)

    text(s, 4.10, 1.24, 5.5, 0.26, 'КАТАЛОГ ПОСТАЧАЛЬНИКІВ', size=10, color=BRIGHT,
         bold=True)
    text(s, 4.10, 1.62, 5.5, 1.30, 'Азійські снеки\nз водоростей і рису', size=31,
         font=HEAD, color=WHITE, bold=True, line_spacing=1.12)
    rect(s, 4.10, 3.18, 1.30, 0.05, fill=BRIGHT)
    text(s, 4.10, 3.46, 5.4, 1.00,
         'Продукти постачальників, короткі описи та собівартість\n'
         'одиниці товару в доларах і гривні — за всіма сценаріями\n'
         'доставки з файлу Self Cost.',
         size=11, color=WHITE, line_spacing=1.4)


def slide_supplier(prs, sup, index):
    s = blank(prs)
    rect(s, 0, 0, SW, SH, fill=PAGE)
    rect(s, 5.85, 0, SW - 5.85, SH, fill=DARK)
    rect(s, 6.35, 1.05, 3.20, 3.20, fill=GREEN)
    text(s, 6.45, 1.35, 3.00, 1.60, sup['brand_mark'], size=25, font=HEAD,
         color=WHITE, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         line_spacing=1.08)
    rect(s, 7.60, 3.02, 0.70, 0.04, fill=BRIGHT)
    text(s, 6.45, 3.24, 3.00, 0.60, sup['brand'], size=11, color=BRIGHT,
         align=PP_ALIGN.CENTER)
    text(s, 6.35, 4.52, 3.20, 0.40, '%d SKU із собівартістю' % len(sup['products']),
         size=10, color=WHITE, bold=True, align=PP_ALIGN.CENTER)

    rect(s, M, 0.62, 0.26, 0.26, fill=BRIGHT)
    text(s, M + 0.40, 0.60, 4.6, 0.30, 'ПОСТАЧАЛЬНИК %d' % index, size=10,
         color=GREEN, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    text(s, M, 1.16, 5.15, 1.50, sup['name'], size=25, font=HEAD, color=DARK,
         bold=True, line_spacing=1.05)
    rect(s, M, 2.78, 1.30, 0.05, fill=BRIGHT)
    text(s, M, 3.04, 5.10, 1.10, sup['summary'], size=10.5, color=TEXT,
         line_spacing=1.35)

    x = M
    for label in (sup['country'], sup['port']):
        w = 0.085 * len(label) + 0.44
        pill(s, x, 4.44, w, 0.34, label, fill=MINT, color=GREEN, size=9)
        x += w + 0.16


def slides_products(prs, sup):
    items = sup['products']
    pages = []
    rest = list(items)
    while rest:                      # never leave a lone card on the last page
        take = 4 if len(rest) != 6 else 3
        pages.append(rest[:take])
        rest = rest[take:]

    for idx, page in enumerate(pages):
        s = blank(prs)
        suffix = '' if len(pages) == 1 else ' · %d/%d' % (idx + 1, len(pages))
        header(s, '%s%s' % (sup['short'], suffix), tag=sup['brand'])
        n = len(page)
        if n <= 2:
            gap = 0.30
            w = (SW - 2 * M - gap * (n - 1)) / n
            draw = card_duo
        else:
            gap = 0.18
            w = (SW - 2 * M - gap * (n - 1)) / n
            draw = card_grid
        x0 = (SW - (w * n + gap * (n - 1))) / 2
        for i, p in enumerate(page):
            draw(s, x0 + i * (w + gap), w, p, costs(sup['sheet'], p['key']))
        if sup.get('note'):
            text(s, M, SH - 0.24, SW - 2 * M, 0.18, sup['note'], size=7,
                 color=SOFT, italic=True)


def slide_koriko(prs):
    s = blank(prs)
    header(s, KORIKO['short'], tag=KORIKO['brand'])
    rect(s, M, 0.94, SW - 2 * M, 0.46, fill=MINT)
    text(s, M + 0.20, 0.94, SW - 2 * M - 0.40, 0.46, KORIKO['summary'], size=9,
         color=SOFT, italic=True, anchor=MSO_ANCHOR.MIDDLE)

    w, gap = 2.98, 0.18
    for i, (ph, name, pack, price) in enumerate(KORIKO['items']):
        col, row = i % 3, i // 3
        x = M + col * (w + gap)
        y = 1.58 + row * 1.92
        rect(s, x, y, w, 1.76, fill=WHITE, shadow=True)
        picture(s, photo(ph), x + 0.14, y + 0.14, 1.16, 1.48)
        text(s, x + 1.44, y + 0.22, w - 1.62, 0.44, name, size=10, font=HEAD,
             bold=True, color=DARK, line_spacing=1.05)
        text(s, x + 1.44, y + 0.72, w - 1.62, 0.44, pack, size=8, color=TEXT,
             line_spacing=1.2)
        pill(s, x + 1.44, y + 1.22, min(w - 1.62, 0.075 * len(price) + 0.40), 0.30,
             price, size=9)


def main():
    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(SH)

    slide_cover(prs)
    for i, sup in enumerate(SUPPLIERS, start=1):
        slide_supplier(prs, sup, i)
        slides_products(prs, sup)
    slide_koriko(prs)

    prs.save(OUT)
    print('saved %s — %d slides' % (OUT, len(prs.slides._sldIdLst)))


if __name__ == '__main__':
    main()
