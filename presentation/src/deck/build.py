# -*- coding: utf-8 -*-
"""Build the suppliers deck: supplier -> product -> short description -> self-cost."""
import os
import sys

from pptx import Presentation
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches

sys.path.insert(0, os.path.dirname(__file__))
from catalog import BASE, DATA, KORIKO, SUPPLIERS, photo  # noqa: E402
from theme import (ACCENT, ACCENT_TX, BODY_TX, BRIGHT, CONTENT_BOTTOM, CONTENT_TOP,  # noqa: E402
                   HEAD, INK, M, MUTED, PANEL, RULE, SH, SW, TINT, WHITE, blank,
                   fmt_uah, fmt_usd, footer, header, hline, label, picture, rect, text)

OUT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..',
                                   'Постачальники_снеків_собівартість.pptx'))

# scenarios in ascending cost order; the cheapest is the highlighted row
SC_ORDER = [(BASE, '40′ контейнер'), ("20'", '20′ контейнер'),
            ('LCL 34', 'Збірний 34 м³'), ('LCL 17', 'Збірний 17 м³')]

_page = [0]


def close(slide, note=None):
    _page[0] += 1
    footer(slide, _page[0], note)


def costs(sheet, key):
    for row in DATA[sheet]:
        if row['name'] == key:
            return row['cost']
    raise KeyError('%s / %s' % (sheet, key))


# ================================================================ cost blocks ===
def cost_stack(s, x, y, w, cost, rh=0.18):
    """Full-bleed rows across the card; the cheapest scenario is an accent band."""
    for i, (key, name) in enumerate(SC_ORDER):
        v = cost[key]
        yy = y + i * rh
        hot = key == BASE
        if hot:
            rect(s, x, yy, w, rh, fill=ACCENT)
        elif i:
            hline(s, x + 0.16, yy, w - 0.32)
        col = WHITE if hot else BODY_TX
        text(s, x + 0.16, yy, w * 0.40, rh, name, size=6.5,
             color=WHITE if hot else MUTED, bold=hot, anchor=MSO_ANCHOR.MIDDLE)
        text(s, x + w * 0.40, yy, w * 0.27, rh, fmt_usd(v['usd']), size=7.5,
             color=col, bold=True, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
        text(s, x + w * 0.67, yy, w * 0.33 - 0.16, rh, fmt_uah(v['uah']), size=7.5,
             color=col, bold=True, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)


def cost_row(s, x, y, w, h, cost):
    """Four side-by-side cells, for the roomier two-up cards."""
    cw = w / 4
    for i, (key, name) in enumerate(SC_ORDER):
        v = cost[key]
        cx = x + i * cw
        hot = key == BASE
        rect(s, cx, y, cw, h, fill=ACCENT if hot else PANEL)
        if not hot and i:
            rect(s, cx, y + 0.12, 0.01, h - 0.24, fill=RULE)
        text(s, cx, y + 0.10, cw, 0.16, name, size=7,
             color=WHITE if hot else MUTED, bold=True, align=PP_ALIGN.CENTER)
        text(s, cx, y + 0.29, cw, 0.28, fmt_usd(v['usd']), size=13, font=HEAD,
             bold=True, color=WHITE if hot else INK, align=PP_ALIGN.CENTER)
        text(s, cx, y + 0.58, cw, 0.20, fmt_uah(v['uah']), size=8.5, bold=True,
             color=WHITE if hot else BODY_TX, align=PP_ALIGN.CENTER)


# ====================================================================== cards ===
def card_grid(s, x, w, p, cost):
    y = CONTENT_TOP
    h = CONTENT_BOTTOM - CONTENT_TOP
    rect(s, x, y, w, h, fill=WHITE, line=RULE)
    rect(s, x, y, w, 1.40, fill=PANEL)
    picture(s, photo(p['photo']), x + 0.16, y + 0.10, w - 0.32, 1.20)

    label(s, x + 0.16, y + 1.54, w - 0.32, p['badge'], color=ACCENT_TX, size=6.5)
    text(s, x + 0.16, y + 1.74, w - 0.32, 0.44, p['title'], size=12, font=HEAD,
         bold=True, color=INK, line_spacing=1.0)
    text(s, x + 0.16, y + 2.24, w - 0.32, 0.64, p['desc'], size=8, color=BODY_TX,
         line_spacing=1.22)

    label(s, x + 0.16, y + 2.96, w - 0.32, 'Собівартість за одиницю', size=6)
    cost_stack(s, x, y + 3.16, w, cost)


def card_duo(s, x, w, p, cost):
    y = CONTENT_TOP
    h = CONTENT_BOTTOM - CONTENT_TOP
    rect(s, x, y, w, h, fill=WHITE, line=RULE)
    rect(s, x, y, w, 1.55, fill=PANEL)
    picture(s, photo(p['photo']), x + 0.30, y + 0.12, w - 0.60, 1.31)

    label(s, x + 0.30, y + 1.70, w - 0.60, p['badge'], color=ACCENT_TX, size=7.5,
          align=PP_ALIGN.CENTER)
    text(s, x + 0.30, y + 1.90, w - 0.60, 0.42, p['title'].replace('\n', ' '),
         size=16, font=HEAD, bold=True, color=INK, align=PP_ALIGN.CENTER)
    text(s, x + 0.55, y + 2.38, w - 1.10, 0.46, p['desc'], size=9, color=BODY_TX,
         align=PP_ALIGN.CENTER, line_spacing=1.24)

    label(s, x + 0.30, y + 2.94, w - 0.60, 'Собівартість за одиницю', size=6.5,
          align=PP_ALIGN.CENTER)
    cost_row(s, x + 0.01, y + 3.12, w - 0.02, 0.80, cost)


# ===================================================================== slides ===
def slide_cover(prs):
    s = blank(prs)
    label(s, M, 0.80, 6.0, 'Каталог постачальників', color=ACCENT_TX, size=9)
    text(s, M, 1.10, 8.0, 1.40, 'Азійські снеки\nз водоростей і рису', size=36,
         font=HEAD, color=INK, bold=True, line_spacing=1.12)
    rect(s, M, 2.66, 1.10, 0.045, fill=ACCENT)
    text(s, M, 2.92, 6.4, 0.70,
         'Продукти п’яти постачальників, короткі описи та собівартість одиниці\n'
         'товару в доларах і гривні — за всіма сценаріями доставки.',
         size=11.5, color=BODY_TX, line_spacing=1.36)

    strip = ['zek_tempura_corn30', 'sk_original', 'tn_original', 'tmk_roll_orig',
             'zek_topping_veg35']
    tw, gap = 1.68, 0.20
    x = M
    for name in strip:
        rect(s, x, 3.86, tw, 1.30, fill=PANEL)
        picture(s, photo(name), x + 0.12, 3.96, tw - 0.24, 1.10)
        x += tw + gap

    text(s, SW - M - 3.0, 0.80, 3.0, 0.20,
         '5 постачальників · 28 SKU · 4 сценарії доставки', size=8, color=MUTED,
         align=PP_ALIGN.RIGHT)
    _page[0] += 1


def slide_supplier(prs, sup, index):
    s = blank(prs)
    label(s, M, 0.80, 5.3, 'Постачальник %02d · %s' % (index, sup['brand']),
          color=ACCENT_TX, size=9)
    text(s, M, 1.08, 5.3, 1.10, sup['name'], size=27, font=HEAD, color=INK,
         bold=True, line_spacing=1.1)
    rect(s, M, 2.44, 1.10, 0.045, fill=ACCENT)
    text(s, M, 2.70, 5.15, 1.10, sup['summary'], size=10.5, color=BODY_TX,
         line_spacing=1.34)

    facts = [('Країна', sup['country']), ('Умови', sup['port']),
             ('Позицій із собівартістю', str(len(sup['products'])))]
    y = 4.06
    hline(s, M, y, 5.15)
    for i, (k, v) in enumerate(facts):
        yy = y + 0.06 + i * 0.34
        text(s, M, yy, 2.4, 0.24, k, size=8.5, color=MUTED,
             anchor=MSO_ANCHOR.MIDDLE)
        text(s, M + 2.4, yy, 2.7, 0.24, v, size=9.5, color=INK, bold=True,
             anchor=MSO_ANCHOR.MIDDLE)
        hline(s, M, yy + 0.28, 5.15)

    # bright product tiles on the right so the divider is not a wall of type
    shots = [p['photo'] for p in sup['products']][:4]
    bx, by, gap = 6.00, 0.80, 0.16
    bw, bh = SW - M - bx, 4.36
    cols = 1 if len(shots) == 1 else 2
    rows = -(-len(shots) // cols)
    tw = (bw - gap * (cols - 1)) / cols
    # packs are portrait; cap the tile so a short grid does not leave a tall
    # band of empty grey, then centre the block in the available height
    th = min((bh - gap * (rows - 1)) / rows, tw * 1.45)
    top = by + (bh - (th * rows + gap * (rows - 1))) / 2
    for i, name in enumerate(shots):
        x = bx + (i % cols) * (tw + gap)
        yy = top + (i // cols) * (th + gap)
        rect(s, x, yy, tw, th, fill=PANEL)
        picture(s, photo(name), x + 0.16, yy + 0.14, tw - 0.32, th - 0.28)
    close(s)


def slides_products(prs, sup):
    items = sup['products']
    pages, rest = [], list(items)
    while rest:                      # never leave a lone card on the last page
        take = 4 if len(rest) != 6 else 3
        pages.append(rest[:take])
        rest = rest[take:]

    for idx, page in enumerate(pages):
        s = blank(prs)
        eyebrow = sup['short'] if len(pages) == 1 else '%s · %d/%d' % (
            sup['short'], idx + 1, len(pages))
        header(s, sup['headline'], eyebrow=eyebrow, tag=sup['brand'])
        n = len(page)
        gap = 0.30 if n <= 2 else 0.20
        w = (SW - 2 * M - gap * (n - 1)) / n
        draw = card_duo if n <= 2 else card_grid
        x0 = (SW - (w * n + gap * (n - 1))) / 2
        for i, p in enumerate(page):
            draw(s, x0 + i * (w + gap), w, p, costs(sup['sheet'], p['key']))
        close(s, sup.get('note') or
              'Собівартість за одиницю товару. Джерело: SelfCost.xlsx, вкладка «%s».'
              % sup['sheet'])


def slide_koriko(prs):
    s = blank(prs)
    header(s, KORIKO['headline'], eyebrow=KORIKO['short'], tag=KORIKO['brand'])
    rect(s, M, CONTENT_TOP, SW - 2 * M, 0.44, fill=TINT)
    rect(s, M, CONTENT_TOP, 0.045, 0.44, fill=ACCENT)
    text(s, M + 0.22, CONTENT_TOP, SW - 2 * M - 0.44, 0.44, KORIKO['summary'],
         size=8.5, color=ACCENT_TX, anchor=MSO_ANCHOR.MIDDLE)

    w, gap = 2.98, 0.20
    for i, (ph, name, pack, price) in enumerate(KORIKO['items']):
        col, row = i % 3, i // 3
        x = M + col * (w + gap)
        y = 1.86 + row * 1.72
        rect(s, x, y, w, 1.56, fill=WHITE, line=RULE)
        rect(s, x, y, 1.16, 1.56, fill=PANEL)
        picture(s, photo(ph), x + 0.10, y + 0.12, 0.96, 1.32)
        text(s, x + 1.34, y + 0.20, w - 1.52, 0.44, name, size=10.5, font=HEAD,
             bold=True, color=INK, line_spacing=1.05)
        text(s, x + 1.34, y + 0.72, w - 1.52, 0.42, pack, size=8, color=BODY_TX,
             line_spacing=1.2)
        hline(s, x + 1.34, y + 1.16, w - 1.52)
        text(s, x + 1.34, y + 1.22, w - 1.52, 0.24, price, size=10, font=HEAD,
             bold=True, color=ACCENT_TX, anchor=MSO_ANCHOR.MIDDLE)
    close(s, 'Джерело: Snacks. Price Comparison.xlsx, вкладка «Nature Best Food Co., Ltd».')


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
