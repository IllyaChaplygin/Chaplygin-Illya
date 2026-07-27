# -*- coding: utf-8 -*-
"""Build the suppliers deck: supplier -> product -> short description -> self-cost."""
import os
import sys

from pptx import Presentation
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches

sys.path.insert(0, os.path.dirname(__file__))
from catalog import BASE, DATA, SUPPLIERS, photo  # noqa: E402
from theme import (BODY_TX, CARD_R, CONTENT_BOTTOM, CONTENT_TOP, GOLD, HEAD, INK,  # noqa: E402
                   M, MUTED, PAPER, RULE, SH, SUPPLIER_COLORS, SW, WHITE, blank,
                   deepen, fit_size, fmt_uah, fmt_usd, footer, header, hline, label,
                   mix, picture, rect, text, tint)

OUT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..',
                                   'Постачальники_снеків_собівартість.pptx'))

# scenarios in ascending cost order; the cheapest is the highlighted row
SC_ORDER = [(BASE, '40′ контейнер'), ("20'", '20′ контейнер'),
            ('LCL 34', 'Збірний 34 м³'), ('LCL 17', 'Збірний 17 м³')]

_page = [0]


def close(slide, note=None):
    _page[0] += 1
    footer(slide, _page[0], note)


def accent_of(sup):
    return SUPPLIER_COLORS[sup['id']]


def costs(sheet, key):
    for row in DATA[sheet]:
        if row['name'] == key:
            return row['cost']
    raise KeyError('%s / %s' % (sheet, key))


# ================================================================ cost blocks ===
def cost_stack(s, x, y, w, cost, accent, rh=0.18):
    """Full-bleed rows across the card; the cheapest scenario is a filled band."""
    for i, (key, name) in enumerate(SC_ORDER):
        v = cost[key]
        yy = y + i * rh
        hot = key == BASE
        if hot:
            rect(s, x, yy, w, rh, fill=accent)
        elif i:
            hline(s, x + 0.16, yy, w - 0.32)
        col = WHITE if hot else BODY_TX
        text(s, x + 0.14, yy, w * 0.36 - 0.14, rh, name, size=6.5,
             color=WHITE if hot else MUTED, bold=hot, anchor=MSO_ANCHOR.MIDDLE)
        text(s, x + w * 0.36, yy, w * 0.30, rh, fmt_usd(v['usd']), size=7.5,
             color=col, bold=True, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
        text(s, x + w * 0.66, yy, w * 0.34 - 0.14, rh, fmt_uah(v['uah']), size=7.5,
             color=col, bold=True, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)


def cost_row(s, x, y, w, h, cost, accent):
    """Four side-by-side cells, for the roomier two-up cards."""
    cw = w / 4
    pale = tint(accent, 0.07)
    for i, (key, name) in enumerate(SC_ORDER):
        v = cost[key]
        cx = x + i * cw
        hot = key == BASE
        rect(s, cx, y, cw, h, fill=accent if hot else pale)
        if not hot and i:
            rect(s, cx, y + 0.12, 0.01, h - 0.24, fill=WHITE)
        text(s, cx, y + 0.10, cw, 0.16, name, size=7,
             color=WHITE if hot else MUTED, bold=True, align=PP_ALIGN.CENTER)
        text(s, cx, y + 0.29, cw, 0.28, fmt_usd(v['usd']), size=13, font=HEAD,
             bold=True, color=WHITE if hot else INK, align=PP_ALIGN.CENTER)
        text(s, cx, y + 0.58, cw, 0.20, fmt_uah(v['uah']), size=8.5, bold=True,
             color=WHITE if hot else deepen(accent, 0.9), align=PP_ALIGN.CENTER)


# ====================================================================== cards ===
def photo_band(page, bw, bh=1.22):
    """Height the photo strip needs on this page.

    Whole suppliers ship only thumbnail-sized shots. Holding a fixed tall strip
    would leave them marooned, so the strip shrinks to the tallest picture on the
    page — the cost block below stays pinned, so cards still line up.
    """
    tallest = max(fit_size(photo(p['photo']), bw, bh)[1] for p in page)
    return min(bh, max(0.80, tallest + 0.10))


def card_grid(s, x, w, p, cost, accent, band=1.22):
    y = CONTENT_TOP
    h = CONTENT_BOTTOM - CONTENT_TOP
    rect(s, x, y, w, h, fill=WHITE, radius=CARD_R, line=RULE)
    rect(s, x + 0.01, y + 0.01, w - 0.02, band + 0.22, fill=tint(accent, 0.07),
         radius=CARD_R)
    rect(s, x + 0.01, y + band + 0.10, w - 0.02, 0.13, fill=tint(accent, 0.07))
    picture(s, photo(p['photo']), x + 0.16, y + 0.12, w - 0.32, band)

    top = y + 0.12 + band
    rect(s, x + 0.16, top + 0.24, 0.13, 0.13, fill=accent, radius=0.5)
    label(s, x + 0.36, top + 0.24, w - 0.50, p['badge'],
          color=deepen(accent, 0.9), size=6.5)
    text(s, x + 0.16, top + 0.46, w - 0.32, 0.44, p['title'], size=12, font=HEAD,
         bold=True, color=INK, line_spacing=1.0)
    text(s, x + 0.16, top + 0.96, w - 0.32, 0.64, p['desc'], size=8, color=BODY_TX,
         line_spacing=1.22)

    label(s, x + 0.16, y + 2.96, w - 0.32, 'Собівартість за одиницю', size=6)
    cost_stack(s, x, y + 3.16, w, cost, accent)


def card_duo(s, x, w, p, cost, accent, band=1.34):
    y = CONTENT_TOP
    h = CONTENT_BOTTOM - CONTENT_TOP
    rect(s, x, y, w, h, fill=WHITE, radius=CARD_R, line=RULE)
    wash = tint(accent, 0.07)
    rect(s, x + 0.01, y + 0.01, w - 0.02, band + 0.36, fill=wash, radius=CARD_R)
    rect(s, x + 0.01, y + band + 0.13, w - 0.02, 0.24, fill=wash)
    picture(s, photo(p['photo']), x + 0.30, y + 0.14, w - 0.60, band)

    top = y + 0.14 + band
    label(s, x + 0.30, top + 0.38, w - 0.60, p['badge'], color=deepen(accent, 0.9),
          size=7.5, align=PP_ALIGN.CENTER)
    text(s, x + 0.30, top + 0.58, w - 0.60, 0.42, p['title'].replace('\n', ' '),
         size=16, font=HEAD, bold=True, color=INK, align=PP_ALIGN.CENTER)
    text(s, x + 0.55, top + 1.04, w - 1.10, 0.46, p['desc'], size=9, color=BODY_TX,
         align=PP_ALIGN.CENTER, line_spacing=1.24)

    label(s, x + 0.30, y + 3.06, w - 0.60, 'Собівартість за одиницю', size=6.5,
          align=PP_ALIGN.CENTER)
    cost_row(s, x + 0.01, y + 3.24, w - 0.02, 0.80, cost, accent)


# ===================================================================== slides ===
def slide_cover(prs):
    s = blank(prs, accent=SUPPLIER_COLORS['singha'])
    # the top rule carries every supplier's colour — the deck's range at a glance
    seg = SW / len(SUPPLIERS)
    for i, sup in enumerate(SUPPLIERS):
        rect(s, i * seg, 0, seg, 0.07, fill=accent_of(sup))

    label(s, M, 0.86, 6.0, 'Каталог постачальників · 2026', color=GOLD, size=9)
    text(s, M, 1.18, 8.4, 1.50, 'Азійські снеки\nз водоростей і рису', size=38,
         font=HEAD, color=INK, bold=True, line_spacing=1.10)
    rect(s, M, 2.88, 1.10, 0.05, fill=SUPPLIER_COLORS['zek'])
    text(s, M, 3.12, 6.6, 0.70,
         'Продукти чотирьох постачальників, короткі описи та собівартість\n'
         'одиниці товару в доларах і гривні — за всіма сценаріями доставки.',
         size=11.5, color=BODY_TX, line_spacing=1.36)

    text(s, SW - M - 3.0, 0.86, 3.0, 0.20,
         '4 постачальники · 28 SKU · 4 сценарії доставки', size=8, color=MUTED,
         align=PP_ALIGN.RIGHT)

    # staggered tiles, each washed in its supplier's colour
    strip = [('zek_tempura_corn30', 'zek'), ('sk_original', 'singha'),
             ('tn_original', 'thainichi'), ('tmk_roll_orig', 'tmk'),
             ('zek_topping_veg35', 'zek')]
    tw, gap = 1.68, 0.20
    x = M
    for i, (name, sid) in enumerate(strip):
        lift = 0.16 if i % 2 else 0.0
        rect(s, x, 3.98 - lift, tw, 1.34, fill=tint(SUPPLIER_COLORS[sid], 0.10),
             radius=0.06)
        picture(s, photo(name), x + 0.12, 4.08 - lift, tw - 0.24, 1.14)
        x += tw + gap
    _page[0] += 1


def route_strip(s, x, y, w, origin, accent):
    """Origin port to Ukraine — a small designed cue instead of another text line."""
    label(s, x, y, w, 'Маршрут постачання', size=6.5)
    cy = y + 0.34
    rect(s, x + 0.05, cy, 0.14, 0.14, fill=accent, radius=0.5)
    for i in range(11):
        rect(s, x + 0.34 + i * 0.20, cy + 0.055, 0.10, 0.03, fill=mix(accent, PAPER, 0.4))
    rect(s, x + 2.52, cy, 0.14, 0.14, fill=INK, radius=0.5)
    text(s, x, cy + 0.24, 1.5, 0.20, origin, size=8, color=INK, bold=True)
    text(s, x + 1.7, cy + 0.24, 1.0, 0.20, 'Україна', size=8, color=INK, bold=True,
         align=PP_ALIGN.RIGHT)


def slide_supplier(prs, sup, index):
    accent = accent_of(sup)
    s = blank(prs, accent=accent)

    # chapter marker: a filled square carrying the section number
    rect(s, M, 0.74, 0.44, 0.44, fill=accent, radius=0.10)
    text(s, M, 0.74, 0.44, 0.44, '%02d' % index, size=14, font=HEAD, bold=True,
         color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    label(s, M + 0.60, 0.74, 4.7, 'Постачальник', size=7, h=0.20)
    label(s, M + 0.60, 0.94, 4.7, sup['brand'], color=deepen(accent, 0.9), size=9,
          h=0.22)

    text(s, M, 1.42, 5.3, 1.10, sup['name'], size=27, font=HEAD, color=INK,
         bold=True, line_spacing=1.1)
    rect(s, M, 2.70, 1.10, 0.05, fill=accent)
    text(s, M, 2.94, 5.15, 1.10, sup['summary'], size=10.5, color=BODY_TX,
         line_spacing=1.34)

    route_strip(s, M, 3.92, 3.0, sup['port'], accent)

    # facts as tiles rather than a table row
    fx = M + 3.30
    for value, cap in [(sup['country'], 'Країна'),
                       (str(len(sup['products'])), 'Позицій із с/в')]:
        rect(s, fx, 3.92, 0.90, 0.92, fill=WHITE, radius=0.08, line=RULE)
        text(s, fx, 4.06, 0.90, 0.32, value, size=13, font=HEAD, bold=True,
             color=deepen(accent, 0.9), align=PP_ALIGN.CENTER)
        text(s, fx, 4.46, 0.90, 0.22, cap, size=6.5, color=MUTED,
             align=PP_ALIGN.CENTER)
        fx += 1.02

    supplier_collage(s, sup, accent)
    close(s)


def supplier_collage(s, sup, accent):
    """Two product tiles over a company panel — the panel takes a factory shot."""
    bx, bw = 5.95, SW - M - 5.95
    gap = 0.14
    tw = (bw - gap) / 2
    shots = [p['photo'] for p in sup['products']][:2]

    room = 2.30
    tallest = max(fit_size(photo(n), tw - 0.28, room - 0.28)[1] for n in shots)
    th = min(room, tallest + 0.34)
    for i, name in enumerate(shots):
        x = bx + i * (tw + gap)
        rect(s, x, 0.84, tw, th, fill=WHITE, radius=0.06, line=RULE)
        picture(s, photo(name), x + 0.14, 0.98, tw - 0.28, th - 0.28)

    cy = 0.84 + th + gap
    ch = CONTENT_BOTTOM - cy
    rect(s, bx, cy, bw, ch, fill=accent, radius=0.06)
    if sup.get('company_photo'):
        picture(s, photo(sup['company_photo']), bx + 0.02, cy + 0.02, bw - 0.04,
                ch - 0.04)
    else:
        # no verified factory photography yet — a brand plate holds the slot
        text(s, bx + 0.28, cy + 0.24, bw - 0.56, 0.60, sup['brand_mark'], size=22,
             font=HEAD, bold=True, color=WHITE, line_spacing=1.06)
        rect(s, bx + 0.28, cy + ch - 0.58, 0.60, 0.035, fill=WHITE)
        text(s, bx + 0.28, cy + ch - 0.44, bw - 0.56, 0.24,
             'Виробництво · %s' % sup['country'], size=8.5,
             color=mix(WHITE, accent, 0.85))


def slides_products(prs, sup):
    accent = accent_of(sup)
    items = sup['products']
    pages, rest = [], list(items)
    while rest:                      # never leave a lone card on the last page
        take = 4 if len(rest) != 6 else 3
        pages.append(rest[:take])
        rest = rest[take:]

    for idx, page in enumerate(pages):
        s = blank(prs, accent=accent)
        eyebrow = sup['short'] if len(pages) == 1 else '%s · %d/%d' % (
            sup['short'], idx + 1, len(pages))
        header(s, sup['headline'], eyebrow=eyebrow, tag=sup['brand'], accent=accent)
        n = len(page)
        gap = 0.30 if n <= 2 else 0.20
        w = (SW - 2 * M - gap * (n - 1)) / n
        x0 = (SW - (w * n + gap * (n - 1))) / 2
        band = (photo_band(page, w - 0.60, 1.34) if n <= 2
                else photo_band(page, w - 0.32))
        for i, p in enumerate(page):
            cost = costs(sup['sheet'], p['key'])
            cx = x0 + i * (w + gap)
            if n <= 2:
                card_duo(s, cx, w, p, cost, accent, band)
            else:
                card_grid(s, cx, w, p, cost, accent, band)
        close(s, sup.get('note') or
              'Собівартість за одиницю товару. Джерело: SelfCost.xlsx, вкладка «%s».'
              % sup['sheet'])


def main():
    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(SH)

    slide_cover(prs)
    for i, sup in enumerate(SUPPLIERS, start=1):
        slide_supplier(prs, sup, i)
        slides_products(prs, sup)

    prs.save(OUT)
    print('saved %s — %d slides' % (OUT, len(prs.slides._sldIdLst)))


if __name__ == '__main__':
    main()
