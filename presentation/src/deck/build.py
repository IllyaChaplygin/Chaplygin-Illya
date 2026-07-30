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
                   picture, picture_cover, rect, text, tint)

OUT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..',
                                   'Постачальники_снеків_собівартість.pptx'))

SC_NAME = {"40'": '40′ контейнер', "20'": '20′ контейнер',
           'LCL 34': 'Збірний 34 м³', 'LCL 17': 'Збірний 17 м³'}


def sc_order(cost):
    """Scenarios cheapest-first. The workbook's ordering changes between
    revisions, so it is read off the figures rather than hard-coded."""
    return sorted(SC_NAME, key=lambda k: cost[k]['usd'])

_page = [0]


def close(slide, note=None):
    _page[0] += 1
    footer(slide, _page[0], note)


def accent_of(sup):
    return SUPPLIER_COLORS[sup['id']]


def sections(sup):
    """Every container scenario priced for this supplier, in slide order."""
    out = [dict(sheet=sup['sheet'], title=sup['headline'],
                scenario=sup['scenario'], products=sup['products'])]
    if sup.get('extra_section'):
        out.append(sup['extra_section'])
    return out


def costs(sheet, key):
    for row in DATA[sheet]:
        if row['name'] == key:
            return row['cost']
    raise KeyError('%s / %s' % (sheet, key))


# ================================================================ cost blocks ===
BLOCK_H = 1.13          # label + headline band + three reference rows


def cost_block(s, x, y, w, cost, accent, unit, big=False):
    """Price block with a hard hierarchy.

    The cheapest scenario is the headline figure; the other three sit under it as
    a reference row. The caption always names the unit the price is for — that is
    what stops a 2,5 g sachet being read against a 3 kg bag.
    """
    # geometry is fixed (BLOCK_H) so cards line up; `big` only scales the type
    label(s, x + 0.16, y, w - 0.32, 'Собівартість · %s' % unit,
          color=deepen(accent, 0.9), size=7 if big else 6)

    head = y + 0.20
    hh = 0.48
    order = sc_order(cost)
    v = cost[order[0]]
    rect(s, x, head, w, hh, fill=accent)
    text(s, x + 0.16, head + 0.06, w - 0.32, 0.16, SC_NAME[order[0]].upper(),
         size=6.5 if big else 6, color=WHITE, bold=True, spc=0.9)
    text(s, x + 0.16, head + 0.21, (w - 0.32) * 0.52, 0.30, fmt_usd(v['usd']),
         size=15 if big else 13.5, font=HEAD, bold=True, color=WHITE)
    text(s, x + 0.16 + (w - 0.32) * 0.45, head + 0.25, (w - 0.32) * 0.55, 0.26,
         fmt_uah(v['uah']), size=11 if big else 10, bold=True, color=WHITE,
         align=PP_ALIGN.RIGHT)

    rh = 0.15
    ry = head + hh
    for i, key in enumerate(order[1:]):
        v = cost[key]
        yy = ry + i * rh
        if i:
            hline(s, x + 0.16, yy, w - 0.32)
        text(s, x + 0.16, yy, (w - 0.32) * 0.42, rh, SC_NAME[key], size=6.5,
             color=MUTED, anchor=MSO_ANCHOR.MIDDLE)
        text(s, x + 0.16 + (w - 0.32) * 0.40, yy, (w - 0.32) * 0.28, rh,
             fmt_usd(v['usd']), size=7, color=BODY_TX, bold=True,
             align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
        text(s, x + 0.16 + (w - 0.32) * 0.68, yy, (w - 0.32) * 0.32, rh,
             fmt_uah(v['uah']), size=7, color=BODY_TX, bold=True,
             align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)


# ====================================================================== cards ===
def photo_band(page, bw, bh=1.10):
    """Height the photo strip needs on this page.

    Whole suppliers ship only thumbnail-sized shots. Holding a fixed tall strip
    would leave them marooned, so the strip shrinks to the tallest picture on the
    page — the cost block below stays pinned, so cards still line up.
    """
    tallest = max(fit_size(photo(p['photo']), bw, bh)[1] for p in page)
    return min(bh, max(0.80, tallest + 0.10))


def card_grid(s, x, w, p, cost, accent, band=1.10):
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
    text(s, x + 0.16, top + 0.96, w - 0.32, 0.54, p['desc'], size=8, color=BODY_TX,
         line_spacing=1.22)

    cost_block(s, x, CONTENT_BOTTOM - BLOCK_H - 0.03, w, cost, accent, p['unit'])


def card_duo(s, x, w, p, cost, accent, band=1.25):
    y = CONTENT_TOP
    h = CONTENT_BOTTOM - CONTENT_TOP
    rect(s, x, y, w, h, fill=WHITE, radius=CARD_R, line=RULE)
    wash = tint(accent, 0.07)
    rect(s, x + 0.01, y + 0.01, w - 0.02, band + 0.36, fill=wash, radius=CARD_R)
    rect(s, x + 0.01, y + band + 0.13, w - 0.02, 0.24, fill=wash)
    picture(s, photo(p['photo']), x + 0.30, y + 0.14, w - 0.60, band)

    top = y + 0.14 + band
    label(s, x + 0.30, top + 0.26, w - 0.60, p['badge'], color=deepen(accent, 0.9),
          size=7.5, align=PP_ALIGN.CENTER)
    text(s, x + 0.30, top + 0.48, w - 0.60, 0.42, p['title'].replace('\n', ' '),
         size=16, font=HEAD, bold=True, color=INK, align=PP_ALIGN.CENTER)
    text(s, x + 0.55, top + 0.94, w - 1.10, 0.36, p['desc'], size=9, color=BODY_TX,
         align=PP_ALIGN.CENTER, line_spacing=1.24)

    cost_block(s, x + 0.55, CONTENT_BOTTOM - BLOCK_H - 0.05, w - 1.10, cost, accent,
               p['unit'], big=True)


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

    skus = sum(len({p['title'] for sec in sections(sup) for p in sec['products']})
               for sup in SUPPLIERS)
    text(s, SW - M - 3.2, 0.86, 3.2, 0.20,
         '%d постачальники · %d SKU · 4 сценарії доставки' % (len(SUPPLIERS), skus),
         size=8, color=MUTED, align=PP_ALIGN.RIGHT)

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


def logo_plate(s, x, y, w, h, sup, accent):
    """White plate holding the real logo, or a typographic mark when none exists."""
    rect(s, x, y, w, h, fill=WHITE, radius=0.05)
    if sup.get('logo'):
        picture(s, photo(sup['logo']), x + 0.26, y + 0.16, w - 0.52, h - 0.32)
    else:
        # no logo file for this supplier — a typographic mark, the way the avocado
        # deck sets "SYROS" rather than an image
        text(s, x + 0.14, y + 0.10, w - 0.28, h - 0.20, sup['brand_mark'], size=24,
             font=HEAD, bold=True, color=deepen(accent, 0.95),
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.08)


def brand_panel(s, sup, accent):
    """Full-height colour column: company imagery, brand plate, origin."""
    px, pw = 5.85, SW - 5.85
    rect(s, px, 0, pw, SH, fill=accent)
    hero = sup.get('hero')

    if hero == 'bleed':
        # photography big enough to crop edge to edge; the plate laps its lower edge
        picture_cover(s, photo(sup['company_photo']), px, 0, pw, 3.36, focus=0.5)
        rect(s, px, 3.36, pw, 0.06, fill=mix(WHITE, accent, 0.35))
        logo_plate(s, px + 0.50, 2.78, pw - 1.00, 1.24, sup, accent)
        text(s, px + 0.50, 4.10, pw - 1.00, 0.20, sup['photo_caption'], size=6.5,
             color=mix(WHITE, accent, 0.72), italic=True, align=PP_ALIGN.CENTER)
        cap_y = 4.36
    elif hero == 'card':
        # a small supplier shot stays on its own card rather than being stretched
        logo_plate(s, px + 0.50, 0.62, pw - 1.00, 1.20, sup, accent)
        cw = pw - 1.00
        rect(s, px + 0.50, 2.02, cw, 2.00, fill=WHITE, radius=0.05)
        picture(s, photo(sup['company_photo']), px + 0.58, 2.10, cw - 0.16, 1.60)
        text(s, px + 0.58, 3.74, cw - 0.16, 0.20, sup['photo_caption'], size=6.5,
             color=MUTED, italic=True, align=PP_ALIGN.CENTER)
        cap_y = 4.28
    else:
        logo_plate(s, px + 0.50, 1.34, pw - 1.00, 1.60, sup, accent)
        text(s, px + 0.50, 3.10, pw - 1.00, 0.30, sup.get('tagline', sup['brand']),
             size=10.5, color=mix(WHITE, accent, 0.80), align=PP_ALIGN.CENTER)
        cap_y = 4.10

    label(s, px, cap_y, pw, 'Виробництво', color=mix(WHITE, accent, 0.70), size=7,
          align=PP_ALIGN.CENTER, h=0.20)
    text(s, px, cap_y + 0.22, pw, 0.30, sup['country'], size=13, font=HEAD, bold=True,
         color=WHITE, align=PP_ALIGN.CENTER)


def cert_strip(s, x, y, w, names, accent):
    """Certification marks, sized to a common height and laid out on one line."""
    h = 0.26
    label(s, x, y, w, 'Сертифікати виробництва', color=deepen(accent, 0.9), size=7)
    widths = [fit_size(photo(n), w, h)[0] for n in names]
    gap = min(0.20, max(0.10, (w - sum(widths)) / max(1, len(names) - 1)))
    cx = x
    for name, cw in zip(names, widths):
        picture(s, photo(name), cx, y + 0.20, cw, h)
        cx += cw + gap


def slide_supplier(prs, sup, index):
    """Company profile: brand column on the right, credentials on the left."""
    accent = accent_of(sup)
    s = blank(prs, accent=accent, rule=False)
    brand_panel(s, sup, accent)

    rect(s, M, 0.52, 0.26, 0.26, fill=accent)
    label(s, M + 0.42, 0.52, 4.4, 'Профіль постачальника · %02d' % index,
          color=deepen(accent, 0.9), size=8, h=0.26)

    text(s, M, 0.92, 5.05, 1.00, sup['name'], size=26, font=HEAD, color=INK,
         bold=True, line_spacing=1.08)
    text(s, M, 2.08, 5.05, 0.30, sup['category'], size=11, color=deepen(accent, 0.9))

    cw = 0.062 * len(sup['port']) + 0.44
    rect(s, M, 2.46, cw, 0.30, fill=tint(accent, 0.12), radius=0.24)
    text(s, M, 2.46, cw, 0.30, sup['port'], size=8.5, color=deepen(accent, 0.95),
         bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    label(s, M, 2.96, 5.05, 'Про компанію', color=deepen(accent, 0.9), size=7)
    text(s, M, 3.16, 5.05, 0.72, sup['summary'], size=10, color=BODY_TX,
         line_spacing=1.30)

    certs = sup.get('certs')
    tw, gap = 1.62, 0.145
    sy = 3.94 if certs else 4.16
    for i, (value, cap) in enumerate(sup['stats']):
        x = M + i * (tw + gap)
        th = 0.78 if certs else 0.94
        rect(s, x, sy, tw, th, fill=WHITE, radius=0.06, line=RULE)
        rect(s, x, sy, tw, 0.05, fill=accent)
        text(s, x, sy + 0.14, tw, 0.30, value, size=16 if certs else 17, font=HEAD,
             bold=True, color=deepen(accent, 0.95), align=PP_ALIGN.CENTER)
        text(s, x + 0.10, sy + 0.44, tw - 0.20, 0.34, cap, size=6.5 if certs else 7,
             color=MUTED, align=PP_ALIGN.CENTER, line_spacing=1.12)
    if certs:
        cert_strip(s, M, 4.86, 5.05, certs, accent)

    # the brand column runs to the slide edge, so the page number rides on it
    _page[0] += 1
    text(s, SW - M - 0.6, 5.34, 0.6, 0.16, str(_page[0]), size=7,
         color=mix(WHITE, accent, 0.70), bold=True, align=PP_ALIGN.RIGHT)


VAT_NOTE = ('Собівартість повна: ціна постачальника + логістика + мито + ПДВ 20%. '
            'Це не ціна продажу.')


def slides_products(prs, sup, section):
    """One section = one container scenario, so every figure on a slide shares a basis."""
    accent = accent_of(sup)
    pages, rest = [], list(section['products'])
    while rest:                      # never leave a lone card on the last page
        take = 4 if len(rest) != 6 else 3
        pages.append(rest[:take])
        rest = rest[take:]

    for idx, page in enumerate(pages):
        s = blank(prs, accent=accent)
        eyebrow = sup['short'] if len(pages) == 1 else '%s · %d/%d' % (
            sup['short'], idx + 1, len(pages))
        header(s, section['title'], eyebrow=eyebrow, accent=accent)

        # the container the prices assume, stated on the slide itself
        cap = section['scenario']
        pw = min(4.4, 0.058 * len(cap) + 0.44)
        rect(s, SW - M - pw, 0.58, pw, 0.32, fill=tint(accent, 0.12), radius=0.24)
        text(s, SW - M - pw, 0.58, pw, 0.32, cap, size=7.5,
             color=deepen(accent, 0.95), bold=True, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE)

        n = len(page)
        gap = 0.30 if n <= 2 else 0.20
        w = (SW - 2 * M - gap * (n - 1)) / n
        x0 = (SW - (w * n + gap * (n - 1))) / 2
        band = (photo_band(page, w - 0.60, 1.25) if n <= 2
                else photo_band(page, w - 0.32))
        for i, p in enumerate(page):
            cost = costs(section['sheet'], p['key'])
            cx = x0 + i * (w + gap)
            if n <= 2:
                card_duo(s, cx, w, p, cost, accent, band)
            else:
                card_grid(s, cx, w, p, cost, accent, band)

        note = VAT_NOTE + (' ' + sup['note'] if sup.get('note') else '')
        close(s, note + ' Джерело: SelfCost.xlsx, «%s».' % section['sheet'])


def main():
    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(SH)

    slide_cover(prs)
    for i, sup in enumerate(SUPPLIERS, start=1):
        slide_supplier(prs, sup, i)
        for section in sections(sup):
            slides_products(prs, sup, section)

    prs.save(OUT)
    print('saved %s — %d slides' % (OUT, len(prs.slides._sldIdLst)))


if __name__ == '__main__':
    main()
