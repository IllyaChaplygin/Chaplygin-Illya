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
                   M, MUTED, RULE, SH, SUPPLIER_COLORS, SW, WHITE, blank, deepen,
                   fit_size, fmt_uah, fmt_usd, footer, header, hline, label, mix,
                   picture, rect, text, tint)

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


def brand_panel(s, sup, accent):
    """Full-height colour column with the brand held on a white plate."""
    px, pw = 5.85, SW - 5.85
    rect(s, px, 0, pw, SH, fill=accent)

    cx, cw = px + 0.50, pw - 1.00
    rect(s, cx, 1.10, cw, 2.55, fill=WHITE, radius=0.05)
    if sup.get('logo'):
        picture(s, photo(sup['logo']), cx + 0.30, 1.42, cw - 0.60, 1.00)
    else:
        # no logo file for this supplier — a typographic mark, as the avocado deck
        # sets "SYROS" rather than an image
        text(s, cx + 0.16, 1.42, cw - 0.32, 1.00, sup['brand_mark'], size=25,
             font=HEAD, bold=True, color=deepen(accent, 0.95),
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.08)
    rect(s, cx + (cw - 0.66) / 2, 2.68, 0.66, 0.035, fill=accent)
    text(s, cx + 0.16, 2.88, cw - 0.32, 0.30, sup.get('tagline', sup['brand']),
         size=10.5, color=MUTED, align=PP_ALIGN.CENTER)

    label(s, px, 4.04, pw, 'Виробництво', color=mix(WHITE, accent, 0.72), size=7,
          align=PP_ALIGN.CENTER, h=0.20)
    text(s, px, 4.26, pw, 0.30, sup['country'], size=13, font=HEAD, bold=True,
         color=WHITE, align=PP_ALIGN.CENTER)


def slide_supplier(prs, sup, index):
    """Company profile: brand column on the right, credentials on the left."""
    accent = accent_of(sup)
    s = blank(prs, accent=accent, rule=False)
    brand_panel(s, sup, accent)

    rect(s, M, 0.52, 0.26, 0.26, fill=accent)
    label(s, M + 0.42, 0.52, 4.4, 'Профіль постачальника · %02d' % index,
          color=deepen(accent, 0.9), size=8, h=0.26)

    text(s, M, 0.94, 5.05, 1.00, sup['name'], size=26, font=HEAD, color=INK,
         bold=True, line_spacing=1.08)
    text(s, M, 2.10, 5.05, 0.30, sup['category'], size=11, color=deepen(accent, 0.9))

    cw = 0.062 * len(sup['port']) + 0.44
    rect(s, M, 2.50, cw, 0.30, fill=tint(accent, 0.12), radius=0.24)
    text(s, M, 2.50, cw, 0.30, sup['port'], size=8.5, color=deepen(accent, 0.95),
         bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    label(s, M, 3.04, 5.05, 'Про компанію', color=deepen(accent, 0.9), size=7)
    text(s, M, 3.24, 5.05, 0.80, sup['summary'], size=10, color=BODY_TX,
         line_spacing=1.32)

    tw, gap = 1.62, 0.145
    for i, (value, cap) in enumerate(sup['stats']):
        x = M + i * (tw + gap)
        rect(s, x, 4.16, tw, 0.94, fill=WHITE, radius=0.06, line=RULE)
        rect(s, x, 4.16, tw, 0.05, fill=accent)
        text(s, x, 4.32, tw, 0.32, value, size=17, font=HEAD, bold=True,
             color=deepen(accent, 0.95), align=PP_ALIGN.CENTER)
        text(s, x + 0.10, 4.68, tw - 0.20, 0.36, cap, size=7, color=MUTED,
             align=PP_ALIGN.CENTER, line_spacing=1.14)

    # the brand column runs to the slide edge, so the page number rides on it
    _page[0] += 1
    text(s, SW - M - 0.6, 5.34, 0.6, 0.16, str(_page[0]), size=7,
         color=mix(WHITE, accent, 0.70), bold=True, align=PP_ALIGN.RIGHT)


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
