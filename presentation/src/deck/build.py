# -*- coding: utf-8 -*-
"""Build the suppliers & self-cost deck."""
import os
import sys

from pptx import Presentation
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

sys.path.insert(0, os.path.dirname(__file__))
from catalog import BASE, DATA, KORIKO, SCENARIO_LABEL, SCENARIOS, SUPPLIERS, photo  # noqa: E402
from theme import (ACCENT, ACCENT_L, CREAM, GOLD, HEAD, INK, INK_SOFT, LINE, M,  # noqa: E402
                   MUTED, PRIMARY, PRIMARY_L, RGBColor, SH, SW, TEXT, WHITE, blank,
                   chip, fmt_num, fmt_uah, fmt_usd, footer, header, kv_table,
                   page_bg, picture, rect, stat, text)

SOFT = RGBColor(0xA9, 0xB8, 0xBD)   # muted body text on the dark closing slide

OUT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..',
                                   'Постачальники_снеків_собівартість.pptx'))

_page = [0]


def close(slide, note=None):
    _page[0] += 1
    footer(slide, _page[0], note)


def costs(sheet, key):
    for row in DATA[sheet]:
        if row['name'] == key:
            return row['cost']
    raise KeyError('%s / %s' % (sheet, key))


def span_uah(lo, hi):
    """Range label that collapses to a single figure when the SKUs cost the same."""
    return fmt_uah(lo) if abs(hi - lo) < 0.005 else '%s – %s' % (fmt_num(lo), fmt_uah(hi))


def all_rows():
    for s in SUPPLIERS:
        for p in s['products']:
            yield s, p, costs(s['sheet'], p['key'])


# =========================================================== cover / intro ===
def slide_cover(prs):
    s = blank(prs)
    page_bg(s, INK)
    rect(s, 0, 0, 3.15, SH, fill=INK_SOFT)
    rect(s, 0, 0, 0.10, SH, fill=ACCENT)

    for i, name in enumerate(['zek_tempura_corn30', 'tmk_roll_orig', 'sk_original',
                              'zek_topping_veg35']):
        picture(s, photo(name), 0.30, 0.42 + i * 1.20, 2.55, 1.05)

    text(s, 3.75, 1.02, 5.9, 0.24, 'КАТАЛОГ ПОСТАЧАЛЬНИКІВ · 2026', size=9,
         color=GOLD, bold=True)
    text(s, 3.75, 1.40, 5.9, 1.55,
         [('Азійські снеки', {'size': 34}), ('з водоростей і рису', {'size': 34})],
         font=HEAD, color=WHITE, bold=True, line_spacing=0.95)
    rect(s, 3.75, 3.02, 1.5, 0.035, fill=ACCENT)
    text(s, 3.75, 3.24, 5.85, 0.80,
         'Продуктові лінійки, короткі описи та собівартість одиниці\n'
         'у доларах і гривні — за даними файлів Self Cost і Price Comparison.',
         size=11, color=WHITE, line_spacing=1.35)

    x = 3.75
    for value, label in [('5', 'постачальників'), ('28', 'SKU із собівартістю'),
                         ('4', 'сценарії логістики')]:
        rect(s, x, 4.32, 1.72, 0.72, fill=INK, radius=0.12)
        text(s, x, 4.42, 1.72, 0.30, value, size=17, font=HEAD, bold=True,
             color=GOLD, align=PP_ALIGN.CENTER)
        text(s, x, 4.74, 1.72, 0.22, label, size=7.5, color=WHITE,
             align=PP_ALIGN.CENTER)
        x += 1.86
    _page[0] += 1


def slide_overview(prs):
    s = blank(prs)
    page_bg(s)
    header(s, 'Постачальники у деці', eyebrow='огляд')

    cards = [(sup['short'], sup['name'], sup['country'],
              '%d SKU із собівартістю' % len(sup['products']),
              min(costs(sup['sheet'], p['key'])[BASE]['uah'] for p in sup['products']),
              max(costs(sup['sheet'], p['key'])[BASE]['uah'] for p in sup['products']))
             for sup in SUPPLIERS]

    w, gap = 2.28, 0.16
    x = M
    for short, full, country, skus, lo, hi in cards:
        rect(s, x, 1.05, w, 2.05, fill=WHITE, radius=0.09, line=LINE)
        rect(s, x, 1.05, w, 0.05, fill=PRIMARY)
        text(s, x + 0.16, 1.24, w - 0.32, 0.46, short, size=12.5, font=HEAD,
             bold=True, color=TEXT, line_spacing=0.95)
        text(s, x + 0.16, 1.74, w - 0.32, 0.44, full, size=7, color=MUTED,
             line_spacing=1.15)
        chip(s, x + 0.16, 2.22, country, fill=PRIMARY_L, color=PRIMARY)
        text(s, x + 0.16, 2.54, w - 0.32, 0.20, skus, size=7.5, color=TEXT, bold=True)
        text(s, x + 0.16, 2.76, w - 0.32, 0.22, 'с/в %s' % span_uah(lo, hi), size=8,
             color=ACCENT, bold=True)
        x += w + gap

    rect(s, M, 3.30, w, 1.62, fill=CREAM, radius=0.09, line=LINE)
    text(s, M + 0.16, 3.46, w - 0.32, 0.42, 'Nature Best\nKoriKo', size=12.5,
         font=HEAD, bold=True, color=TEXT, line_spacing=0.95)
    chip(s, M + 0.16, 3.96, 'Таїланд', fill=PRIMARY_L, color=PRIMARY)
    text(s, M + 0.16, 4.26, w - 0.32, 0.55,
         '24 позиції у прайсі.\nСобівартість ще не\nрозрахована.', size=7.5,
         color=MUTED, line_spacing=1.2)

    bx = M + w + gap
    rect(s, bx, 3.30, SW - 2 * M - w - gap, 1.62, fill=WHITE, radius=0.09, line=LINE)
    text(s, bx + 0.22, 3.46, 6.6, 0.24, 'ЯК ЧИТАТИ ЦЮ ПРЕЗЕНТАЦІЮ', size=8,
         color=ACCENT, bold=True)
    rows = [
        'Кожен постачальник — це профіль компанії, картки продуктів і таблиця собівартості.',
        'На картці показано собівартість за сценарієм 40′ контейнер — він найвигідніший.',
        'Повна таблиця за чотирма сценаріями логістики йде одразу після карток.',
        'Усі суми — за одиницю товару (пакет або пакетик), а не за картон.',
    ]
    y = 3.78
    for i, r in enumerate(rows):
        rect(s, bx + 0.24, y + 0.045, 0.115, 0.115, fill=ACCENT, radius=0.5)
        text(s, bx + 0.46, y, 6.3, 0.24, r, size=8.5, color=TEXT)
        y += 0.28
    close(s)


def slide_method(prs):
    s = blank(prs)
    page_bg(s)
    header(s, 'Як формується собівартість', eyebrow='методологія')

    text(s, M, 0.95, 5.9, 0.24,
         'Приклад: ZEK Tempura Seaweed Corn 30 г, 20′ контейнер, 1 728 пакетів',
         size=8.5, color=MUTED, italic=True)

    parts = [('Товар FOB', 743.04, PRIMARY), ('Транспорт і брокер', 409.92, INK),
             ('Мито 10%', 101.86, GOLD), ('ПДВ 20%', 224.12, ACCENT),
             ('Термінал, сертифікація', 2.36, MUTED)]
    total = sum(p[1] for p in parts)

    bar_x, bar_y, bar_w, bar_h = M, 1.34, SW - 2 * M, 0.46
    x = bar_x
    for i, (label, val, color) in enumerate(parts):
        w = bar_w * val / total
        rect(s, x, bar_y, w, bar_h, fill=color)
        if w > 0.55:
            text(s, x, bar_y, w, bar_h, '%.1f%%' % (100 * val / total),
                 size=9.5, color=WHITE, bold=True, align=PP_ALIGN.CENTER,
                 anchor=MSO_ANCHOR.MIDDLE)
        x += w

    # labels: wide segments get one underneath, narrow ones a right-hand legend
    x = bar_x
    legend = []
    for label, val, color in parts:
        w = bar_w * val / total
        if w > 0.95:
            text(s, x, bar_y + bar_h + 0.09, w, 0.34,
                 [(label, {'size': 7.5, 'bold': True, 'color': TEXT}),
                  ('$%s' % fmt_num(val), {'size': 7, 'color': MUTED})],
                 line_spacing=1.15)
        else:
            legend.append((label, val, color, 100 * val / total))
        x += w

    ly = bar_y + bar_h + 0.44          # own row, clear of the under-bar labels
    lx = bar_x + bar_w
    for label, val, color, pct in reversed(legend):
        w = 0.050 * len(label) + 0.86
        lx -= w + 0.20
        rect(s, lx, ly + 0.035, 0.09, 0.09, fill=color, radius=0.5)
        text(s, lx + 0.16, ly, w, 0.20,
             '%s · %.1f%% · $%s' % (label, pct, fmt_num(val)), size=7, color=MUTED,
             wrap=False)

    y = 2.58
    text(s, M, y, 4.5, 0.24, 'СКЛАДОВІ РОЗРАХУНКУ', size=8, color=ACCENT, bold=True)
    kv_table(s, M, y + 0.30, 4.5, [
        ('Ціна постачальника', 'FOB, USD за одиницю'),
        ('Логістика ЗЕД', 'фрахт, розподіл на вагу/об’єм'),
        ('Мито', '10% від митної вартості'),
        ('ПДВ', '20% від (митна вартість + мито)'),
        ('Транспорт по Україні', 'авто + брокерські послуги'),
        ('Термінал і сертифікація', 'карантин, декларації, аутсорс'),
        ('Кредитні гроші', '+2,5% на 120 днів'),
    ], rh=0.245, key_w=0.46, size=8)

    bx = M + 4.9
    bw = SW - M - bx
    rect(s, bx, y, bw, 2.44, fill=CREAM, radius=0.09, line=LINE)
    text(s, bx + 0.22, y + 0.16, bw - 0.44, 0.24, 'ЩО ВАЖЛИВО ЗНАТИ', size=8,
         color=ACCENT, bold=True)
    notes = [
        ('Валюта', 'Підсумкова колонка у Self Cost підписана «€», але вхідні ціни — '
                   'FOB у доларах і перерахунку в євро в розрахунку немає. Тому в деці '
                   'показник наведено як $.'),
        ('ПДВ у собівартості', 'ПДВ 20% включено в підсумкову суму — у прикладі це '
                               '15,1%. Якщо ПДВ відшкодовується, фактична база для '
                               'націнки нижча.'),
        ('Курс', 'Гривневі суми взято з файлу без перерахунку: застосовано курс від '
                 '45 до 49 ₴ залежно від позиції.'),
    ]
    yy = y + 0.46
    for title, body in notes:
        rect(s, bx + 0.22, yy + 0.05, 0.10, 0.10, fill=ACCENT, radius=0.5)
        text(s, bx + 0.42, yy - 0.02, bw - 0.66, 0.20, title, size=8, bold=True,
             color=TEXT)
        text(s, bx + 0.42, yy + 0.19, bw - 0.66, 0.46, body, size=7.5, color=MUTED,
             line_spacing=1.18)
        yy += 0.64
    close(s, 'Джерело: SelfCost.xlsx, вкладка «ZEK -Retail», рядок 6 (20′ контейнер).')


def slide_logistics(prs):
    s = blank(prs)
    page_bg(s)
    header(s, 'Чотири сценарії логістики', eyebrow='як доставка впливає на ціну')

    ref = costs('ZEK -Retail', 'ZEK TEMPURA SEAWEED - Corn 30g')
    base = ref[BASE]['usd']
    order = ["40'", "20'", 'LCL 34', 'LCL 17']
    blurbs = {
        "40'": 'Повний 40-футовий контейнер. Максимальний обсяг — мінімальна логістична '
               'частка в одиниці товару. Базовий сценарій цієї деки.',
        "20'": 'Повний 20-футовий контейнер. Менший обсяг, ті самі фіксовані витрати — '
               'собівартість помітно вища.',
        'LCL 34': 'Збірний вантаж 34 м³. Компроміс для середніх обсягів, коли контейнер '
                  'ще не набирається.',
        'LCL 17': 'Збірний вантаж 17 м³. Найдорожчий варіант — підходить для тесту '
                  'товару чи першої партії.',
    }
    w, gap = 2.28, 0.16
    x = M
    for i, sc in enumerate(order):
        c = ref[sc]
        delta = 100 * (c['usd'] / base - 1)
        top = PRIMARY if i == 0 else INK
        rect(s, x, 1.05, w, 2.46, fill=WHITE, radius=0.09, line=LINE)
        rect(s, x, 1.05, w, 0.44, fill=top, radius=0.09)
        rect(s, x, 1.30, w, 0.19, fill=top)
        text(s, x, 1.05, w, 0.44, SCENARIO_LABEL[sc], size=10, font=HEAD, bold=True,
             color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(s, x + 0.14, 1.60, w - 0.28, 0.30, fmt_usd(c['usd']), size=19, font=HEAD,
             bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
        text(s, x + 0.14, 1.94, w - 0.28, 0.22, fmt_uah(c['uah']), size=10.5,
             bold=True, color=TEXT, align=PP_ALIGN.CENTER)
        label = 'базовий сценарій' if i == 0 else '+%.0f%% до базового' % delta
        chip(s, x + (w - 1.30) / 2, 2.22, label,
             fill=PRIMARY_L if i == 0 else ACCENT_L,
             color=PRIMARY if i == 0 else ACCENT, w=1.30, size=6.5)
        text(s, x + 0.16, 2.56, w - 0.32, 1.00, blurbs[sc], size=7.5, color=MUTED,
             line_spacing=1.2)
        x += w + gap

    rect(s, M, 3.86, SW - 2 * M, 1.02, fill=CREAM, radius=0.09, line=LINE)
    rect(s, M, 3.86, 0.05, 1.02, fill=ACCENT)
    text(s, M + 0.24, 4.02, SW - 2 * M - 0.5, 0.24,
         'ГОЛОВНИЙ ВИСНОВОК ДЛЯ ЦІНОУТВОРЕННЯ', size=8, color=ACCENT, bold=True)
    text(s, M + 0.24, 4.28, SW - 2 * M - 0.5, 0.50,
         'Різниця між найдешевшим і найдорожчим сценарієм — понад дві собівартості: '
         '%s проти %s за той самий пакет 30 г. Перш ніж рахувати націнку, треба '
         'зафіксувати сценарій доставки — інакше маржа «попливе» на 100%%+.'
         % (fmt_uah(ref["40'"]['uah']), fmt_uah(ref['LCL 17']['uah'])),
         size=9.5, color=TEXT, line_spacing=1.25)
    close(s, 'Приклад розрахований на ZEK Tempura Seaweed Corn 30 г.')


# ============================================================== suppliers ===
def slide_profile(prs, sup):
    s = blank(prs)
    page_bg(s)
    header(s, sup['name'], eyebrow='профіль постачальника', tag=sup['brand'])

    rect(s, M, 1.02, 5.5, 1.30, fill=CREAM, radius=0.09, line=LINE)
    rect(s, M, 1.02, 0.05, 1.30, fill=PRIMARY)
    text(s, M + 0.24, 1.18, 5.0, 0.22, 'ПРО КОМПАНІЮ', size=8, color=ACCENT, bold=True)
    text(s, M + 0.24, 1.44, 5.05, 0.80, sup['summary'], size=9, color=TEXT,
         line_spacing=1.28)

    x = M
    for value, label in sup['stats']:
        stat(s, x, 2.46, 1.72, 0.94, value, label)
        x += 1.89

    text(s, M, 3.56, 5.5, 0.22, 'КЛЮЧОВІ ПАРАМЕТРИ', size=8, color=ACCENT, bold=True)
    kv_table(s, M, 3.84, 5.5, sup['facts'], rh=0.235, key_w=0.42, size=8)

    bx = M + 5.75
    bw = SW - M - bx
    rect(s, bx, 1.02, bw, 4.10, fill=WHITE, radius=0.09, line=LINE)
    picture(s, photo(sup['extra_photo']), bx + 0.18, 1.18, bw - 0.36, 1.28)
    text(s, bx + 0.20, 2.60, bw - 0.40, 0.42, sup['extra_title'], size=8,
         color=ACCENT, bold=True, line_spacing=1.15)
    y = 3.06
    for item in sup['extra']:
        rect(s, bx + 0.22, y + 0.05, 0.08, 0.08, fill=PRIMARY, radius=0.5)
        text(s, bx + 0.40, y - 0.02, bw - 0.62, 0.32, item, size=7.5, color=MUTED,
             line_spacing=1.15)
        y += 0.33
    close(s)


def product_card(s, x, y, w, h, p, cost):
    rect(s, x, y, w, h, fill=WHITE, radius=0.075, line=LINE)
    rect(s, x + 0.09, y + 0.09, w - 0.18, 1.38, fill=CREAM, radius=0.06)
    picture(s, photo(p['photo']), x + 0.17, y + 0.15, w - 0.34, 1.26)

    chip(s, x + 0.14, y + 1.56, p['badge'], w=min(w - 0.28, 0.062 * len(p['badge']) + 0.30))
    text(s, x + 0.14, y + 1.86, w - 0.28, 0.44, p['title'], size=11, font=HEAD,
         bold=True, color=TEXT, line_spacing=0.98)
    text(s, x + 0.14, y + 2.34, w - 0.28, 0.72, p['desc'], size=7.5, color=MUTED,
         line_spacing=1.2)
    rect(s, x + 0.14, y + 3.10, w - 0.28, 0.01, fill=LINE)
    text(s, x + 0.14, y + 3.17, w - 0.28, 0.20, p['specs'], size=7, color=TEXT)

    py = y + h - 0.62
    rect(s, x + 0.14, py, w - 0.28, 0.52, fill=ACCENT, radius=0.12)
    text(s, x + 0.14, py + 0.06, w - 0.28, 0.16, 'СОБІВАРТІСТЬ · 40′ КОНТЕЙНЕР',
         size=6, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    text(s, x + 0.14, py + 0.23, w - 0.28, 0.26,
         '%s  ·  %s' % (fmt_usd(cost[BASE]['usd']), fmt_uah(cost[BASE]['uah'])),
         size=12.5, font=HEAD, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


def wide_card(s, x, y, w, h, p, cost):
    """Horizontal card (photo left, copy right) used when a supplier has few SKUs."""
    rect(s, x, y, w, h, fill=WHITE, radius=0.06, line=LINE)
    pw = w * 0.40
    rect(s, x + 0.12, y + 0.12, pw, h - 0.24, fill=CREAM, radius=0.05)
    picture(s, photo(p['photo']), x + 0.20, y + 0.20, pw - 0.16, h - 0.40)

    tx = x + pw + 0.28
    tw = w - pw - 0.44
    chip(s, tx, y + 0.22, p['badge'], w=min(tw, 0.062 * len(p['badge']) + 0.30))
    text(s, tx, y + 0.56, tw, 0.52, p['title'], size=14, font=HEAD, bold=True,
         color=TEXT, line_spacing=1.0)
    text(s, tx, y + 1.14, tw, 0.86, p['desc'], size=8.5, color=MUTED, line_spacing=1.25)
    rect(s, tx, y + 2.04, tw, 0.01, fill=LINE)
    text(s, tx, y + 2.12, tw, 0.20, p['specs'], size=7.5, color=TEXT)

    py = y + h - 0.74
    rect(s, tx, py, tw, 0.58, fill=ACCENT, radius=0.10)
    text(s, tx, py + 0.08, tw, 0.16, 'СОБІВАРТІСТЬ · 40′ КОНТЕЙНЕР', size=6.5,
         color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    text(s, tx, py + 0.26, tw, 0.28,
         '%s  ·  %s' % (fmt_usd(cost[BASE]['usd']), fmt_uah(cost[BASE]['uah'])),
         size=13.5, font=HEAD, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


def lineup_panel(s, x, y, w, h, sup):
    """Fills the empty grid slots on a supplier's last product page."""
    vals = [costs(sup['sheet'], p['key']) for p in sup['products']]
    lo = min(v[BASE]['uah'] for v in vals)
    hi = max(v[BASE]['uah'] for v in vals)
    cheapest = min(zip(sup['products'], vals), key=lambda t: t[1][BASE]['uah'])

    rect(s, x, y, w, h, fill=CREAM, radius=0.075, line=LINE)
    rect(s, x, y, 0.05, h, fill=PRIMARY)
    text(s, x + 0.28, y + 0.26, w - 0.56, 0.22, 'ПІДСУМОК ЛІНІЙКИ', size=8,
         color=ACCENT, bold=True)
    text(s, x + 0.28, y + 0.52, w - 0.56, 0.42, sup['short'], size=17, font=HEAD,
         bold=True, color=TEXT)

    rows = [('%d SKU' % len(sup['products']), 'позицій із розрахованою собівартістю'),
            (span_uah(lo, hi), 'розкид собівартості за сценарієм 40′ контейнер'),
            (fmt_uah(cheapest[1][BASE]['uah']),
             'найдешевша позиція — %s' % cheapest[0]['title'].replace('\n', ' '))]
    yy = y + 1.12
    for value, label in rows:
        rect(s, x + 0.28, yy + 0.06, 0.09, 0.09, fill=ACCENT, radius=0.5)
        text(s, x + 0.46, yy - 0.02, w - 0.76, 0.26, value, size=12, font=HEAD,
             bold=True, color=ACCENT)
        text(s, x + 0.46, yy + 0.24, w - 0.76, 0.36, label, size=7.5, color=MUTED,
             line_spacing=1.18)
        yy += 0.74

    text(s, x + 0.28, y + h - 0.72, w - 0.56, 0.52,
         'Повна таблиця собівартості за чотирма сценаріями логістики — '
         'на наступному слайді.', size=8, color=TEXT, line_spacing=1.22)


def slides_products(prs, sup):
    items = sup['products']
    wide = len(items) <= 2
    per = 2 if wide else 4
    pages = [items[i:i + per] for i in range(0, len(items), per)]
    for idx, page in enumerate(pages):
        s = blank(prs)
        page_bg(s)
        suffix = '' if len(pages) == 1 else ' · %d/%d' % (idx + 1, len(pages))
        header(s, '%s — продукти%s' % (sup['short'], suffix),
               eyebrow='картки товарів', tag=sup['brand'])
        gap = 0.20 if wide else 0.16
        w = (SW - 2 * M - gap * (per - 1)) / per
        h = 3.30 if wide else 4.00
        y = 1.10 if wide else 1.02
        for i, p in enumerate(page):
            cost = costs(sup['sheet'], p['key'])
            draw = wide_card if wide else product_card
            draw(s, M + i * (w + gap), y, w, h, p, cost)
        if not wide and len(page) < per:
            px = M + len(page) * (w + gap)
            lineup_panel(s, px, y, SW - M - px, h, sup)
        if wide:
            rect(s, M, 4.62, SW - 2 * M, 0.56, fill=CREAM, radius=0.08, line=LINE)
            rect(s, M, 4.62, 0.05, 0.56, fill=PRIMARY)
            text(s, M + 0.22, 4.74, SW - 2 * M - 0.5, 0.34,
                 'Решта прайсу цього постачальника — на попередньому слайді профілю. '
                 'Повна таблиця собівартості за чотирма сценаріями логістики — на наступному.',
                 size=8.5, color=TEXT, line_spacing=1.2)
        close(s, sup.get('note') or
              'Собівартість — за одиницю товару, сценарій 40′ контейнер. Джерело: SelfCost.xlsx.')


def cost_insight(s, sup, y, avail):
    """Portfolio read-out under a cost table; collapses when the table is long."""
    if avail < 0.72:
        return
    vals = [costs(sup['sheet'], p['key']) for p in sup['products']]
    lo = min(v[BASE]['uah'] for v in vals)
    hi = max(v[BASE]['uah'] for v in vals)
    avg = sum(v[BASE]['uah'] for v in vals) / len(vals)
    up17 = 100 * (sum(v['LCL 17']['usd'] / v[BASE]['usd'] for v in vals) / len(vals) - 1)
    up20 = 100 * (sum(v["20'"]['usd'] / v[BASE]['usd'] for v in vals) / len(vals) - 1)

    tiles = [(span_uah(lo, hi), 'розкид собівартості\nу портфелі'),
             (fmt_uah(avg), 'середня собівартість\nодиниці'),
             ('+%.0f%%' % up17, 'надбавка збірного\nвантажу 17 м³')]
    h = min(1.34, avail)
    rect(s, M, y, SW - 2 * M, h, fill=CREAM, radius=0.09, line=LINE)
    rect(s, M, y, 0.05, h, fill=PRIMARY)

    if h >= 1.10:
        text(s, M + 0.24, y + 0.14, 4.0, 0.22, 'ЧИТАННЯ ПОРТФЕЛЯ', size=8,
             color=ACCENT, bold=True)
        x = M + 0.24
        for value, label in tiles:
            text(s, x, y + 0.42, 1.95, 0.28, value, size=13, font=HEAD, bold=True,
                 color=ACCENT)
            text(s, x, y + 0.74, 1.95, 0.42, label, size=7, color=MUTED,
                 line_spacing=1.15)
            x += 2.06
        text(s, x + 0.10, y + 0.40, SW - M - x - 0.30, 0.78,
             'Перехід із 40′ контейнера на 20′ додає в середньому +%.0f%% до собівартості, '
             'а збірний вантаж 17 м³ — +%.0f%%. Для цього постачальника сценарій доставки '
             'важить більше, ніж різниця між його найдешевшою і найдорожчою позицією.'
             % (up20, up17), size=8, color=TEXT, line_spacing=1.22)
    else:
        x = M + 0.24
        for value, label in tiles:
            text(s, x, y, 3.0, h,
                 [('%s — %s' % (value, label.replace('\n', ' ')),
                   {'size': 8, 'bold': True, 'color': TEXT})],
                 anchor=MSO_ANCHOR.MIDDLE)
            x += 3.05


def slides_cost_table(prs, sup, chunk=8):
    items = sup['products']
    pages = [items[i:i + chunk] for i in range(0, len(items), chunk)]
    for idx, page in enumerate(pages):
        s = blank(prs)
        page_bg(s)
        suffix = '' if len(pages) == 1 else ' · %d/%d' % (idx + 1, len(pages))
        header(s, '%s — собівартість за сценаріями%s' % (sup['short'], suffix),
               eyebrow='повна таблиця', tag=sup['brand'])

        cols = [3.30, 0.78] + [1.24] * 4
        x0 = M
        y0 = 1.06
        rh = 0.52 if len(page) <= 3 else (0.36 if len(page) <= 6 else 0.30)

        rect(s, x0, y0, sum(cols), 0.52, fill=INK)
        heads = ['Продукт', 'FOB $'] + [SCENARIO_LABEL[sc] for sc in SCENARIOS]
        x = x0
        for i, (head, cw) in enumerate(zip(heads, cols)):
            al = PP_ALIGN.LEFT if i == 0 else PP_ALIGN.CENTER
            pad = 0.12 if i == 0 else 0
            text(s, x + pad, y0, cw - pad, 0.52, head, size=8, color=WHITE, bold=True,
                 align=al, anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.05)
            if i == 3:
                rect(s, x, y0, cw, 0.045, fill=GOLD)
            x += cw

        y = y0 + 0.52
        for r, p in enumerate(page):
            c = costs(sup['sheet'], p['key'])
            if r % 2 == 0:
                rect(s, x0, y, sum(cols), rh, fill=CREAM)
            rect(s, x0 + cols[0] + cols[1] + cols[2], y, cols[3], rh, fill=PRIMARY_L)
            title = p['title'].replace('\n', ' ')
            text(s, x0 + 0.12, y, cols[0] - 0.20, rh, title, size=8, color=TEXT,
                 bold=True, anchor=MSO_ANCHOR.MIDDLE)
            text(s, x0 + cols[0], y, cols[1], rh, p['specs'].split('FOB ')[-1],
                 size=7.5, color=MUTED, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            x = x0 + cols[0] + cols[1]
            for i, sc in enumerate(SCENARIOS):
                v = c[sc]
                strong = sc == BASE
                text(s, x, y, cols[2 + i], rh,
                     [(fmt_usd(v['usd']), {'size': 8.5, 'bold': True,
                                           'color': ACCENT if strong else TEXT}),
                      (fmt_uah(v['uah']), {'size': 7, 'color': MUTED})],
                     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.0)
                x += cols[2 + i]
            y += rh
        rect(s, x0, y, sum(cols), 0.012, fill=LINE)

        text(s, M, y + 0.14, SW - 2 * M, 0.24,
             'Виділена колонка — 40′ контейнер: саме ці цифри показані на картках товарів. '
             'Курс у розрахунку — 45–49 ₴ за $ залежно від позиції.',
             size=7.5, color=MUTED, line_spacing=1.2)
        cost_insight(s, sup, y + 0.50, SH - 0.50 - (y + 0.50))
        close(s, 'Джерело: SelfCost.xlsx, вкладка «%s».' % sup['sheet'])


def slide_koriko(prs):
    s = blank(prs)
    page_bg(s)
    header(s, KORIKO['name'], eyebrow='постачальник без розрахунку',
           tag=KORIKO['brand'])

    rect(s, M, 1.02, SW - 2 * M, 0.86, fill=ACCENT_L, radius=0.09)
    rect(s, M, 1.02, 0.05, 0.86, fill=ACCENT)
    text(s, M + 0.24, 1.16, SW - 2 * M - 0.5, 0.22,
         'СОБІВАРТІСТЬ ЩЕ НЕ РОЗРАХОВАНА', size=8, color=ACCENT, bold=True)
    text(s, M + 0.24, 1.40, SW - 2 * M - 0.5, 0.42, KORIKO['summary'], size=8.5,
         color=TEXT, line_spacing=1.22)

    w, gap = 2.98, 0.18
    for i, (ph, name, pack, ctn, price) in enumerate(KORIKO['items']):
        col, row = i % 3, i // 3
        x = M + col * (w + gap)
        y = 2.06 + row * 1.62
        rect(s, x, y, w, 1.46, fill=WHITE, radius=0.075, line=LINE)
        rect(s, x + 0.10, y + 0.10, 1.10, 1.26, fill=CREAM, radius=0.06)
        picture(s, photo(ph), x + 0.16, y + 0.16, 0.98, 1.14)
        text(s, x + 1.32, y + 0.16, w - 1.46, 0.40, name, size=9, font=HEAD,
             bold=True, color=TEXT, line_spacing=1.05)
        text(s, x + 1.32, y + 0.60, w - 1.46, 0.36,
             '%s\n%s' % (pack, ctn), size=7, color=MUTED, line_spacing=1.2)
        chip(s, x + 1.32, y + 1.04, price, w=min(w - 1.46, 0.068 * len(price) + 0.32),
             size=7.5, h=0.26)
    close(s, 'Джерело: Snacks. Price Comparison.xlsx, вкладка «Nature Best Food Co., Ltd».')


# ============================================================== summary ======
def slide_ranking(prs):
    s = blank(prs)
    page_bg(s)
    header(s, 'Усі 28 SKU за собівартістю', eyebrow='зведений рейтинг',
           tag='40′ контейнер')

    rows = sorted(all_rows(), key=lambda t: t[2][BASE]['uah'])
    hi = max(r[2][BASE]['uah'] for r in rows)

    ncol = 2
    per = (len(rows) + ncol - 1) // ncol
    colw = (SW - 2 * M - 0.30) / ncol
    rh = 0.265
    # column offsets within one half: rank | name | supplier | bar | ₴ | $
    RANK, NAME, SUPP, BAR, UAH, USD = 0.08, 0.34, 1.94, 2.58, 3.32, 3.98
    BARW = 0.68

    for c in range(ncol):
        x = M + c * (colw + 0.30)
        rect(s, x, 1.02, colw, 0.30, fill=INK)
        for label, ox, ow, al in [('#', RANK, 0.24, PP_ALIGN.LEFT),
                                  ('Продукт', NAME, 1.55, PP_ALIGN.LEFT),
                                  ('Постачальник', SUPP, 0.62, PP_ALIGN.LEFT),
                                  ('₴ / од.', UAH, 0.62, PP_ALIGN.RIGHT),
                                  ('$ / од.', USD, 0.44, PP_ALIGN.RIGHT)]:
            text(s, x + ox, 1.02, ow, 0.30, label, size=7, color=WHITE, bold=True,
                 align=al, anchor=MSO_ANCHOR.MIDDLE)
        y = 1.32
        for i, (sup, p, cost) in enumerate(rows[c * per:(c + 1) * per], start=c * per + 1):
            v = cost[BASE]
            if i % 2:
                rect(s, x, y, colw, rh, fill=CREAM)
            text(s, x + RANK, y, 0.24, rh, str(i), size=7, color=MUTED, bold=True,
                 anchor=MSO_ANCHOR.MIDDLE)
            text(s, x + NAME, y, 1.55, rh, p['title'].replace('\n', ' '), size=7,
                 color=TEXT, bold=True, anchor=MSO_ANCHOR.MIDDLE)
            text(s, x + SUPP, y, 0.62, rh, sup['short'], size=6, color=MUTED,
                 anchor=MSO_ANCHOR.MIDDLE)
            rect(s, x + BAR, y + 0.095, BARW, 0.075, fill=ACCENT_L)
            rect(s, x + BAR, y + 0.095, max(BARW * v['uah'] / hi, 0.025), 0.075,
                 fill=ACCENT)
            text(s, x + UAH, y, 0.62, rh, fmt_uah(v['uah']), size=7, color=TEXT,
                 bold=True, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
            text(s, x + USD, y, 0.44, rh, fmt_usd(v['usd']), size=7, color=ACCENT,
                 bold=True, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
            y += rh
    close(s, 'Сценарій 40′ контейнер. Довжина смужки — відносно найдорожчої позиції. '
             'Джерело: SelfCost.xlsx.')


def slide_conclusions(prs):
    s = blank(prs)
    page_bg(s)
    header(s, 'Висновки для маркетингу та продажів', eyebrow='підсумок')

    rows = sorted(all_rows(), key=lambda t: t[2][BASE]['uah'])
    cheap = rows[0]
    dear = rows[-1]
    ref = costs('ZEK -Retail', 'ZEK TEMPURA SEAWEED - Corn 30g')
    spread = 100 * (ref['LCL 17']['usd'] / ref[BASE]['usd'] - 1)

    cards = [
        ('Вхід у категорію', fmt_uah(cheap[2][BASE]['uah']),
         '%s (%s) — найдешевша одиниця в портфелі. Формат 2,5 г тримає '
         'найнижчу полицеву ціну і добре працює в прикасовій зоні.'
         % (cheap[1]['title'].replace('\n', ' '), cheap[0]['short'])),
        ('Верхній сегмент', fmt_uah(dear[2][BASE]['uah']),
         '%s (%s) — найдорожча позиція. Велика фасовка 70 г, ціль — родинне '
         'споживання та HoReCa, а не імпульс.'
         % (dear[1]['title'].replace('\n', ' '), dear[0]['short'])),
        ('Ціна логістики', '+%.0f%%' % spread,
         'Настільки зростає собівартість при переході з 40′ контейнера на збірний '
         'вантаж 17 м³. Сценарій доставки треба фіксувати до розрахунку націнки.'),
    ]
    w, gap = 3.05, 0.20
    x = M
    for title, value, body in cards:
        rect(s, x, 1.02, w, 1.72, fill=WHITE, radius=0.09, line=LINE)
        rect(s, x, 1.02, w, 0.05, fill=ACCENT)
        text(s, x + 0.20, 1.20, w - 0.40, 0.22, title.upper(), size=7.5, color=MUTED,
             bold=True)
        text(s, x + 0.20, 1.44, w - 0.40, 0.34, value, size=20, font=HEAD, bold=True,
             color=ACCENT)
        text(s, x + 0.20, 1.86, w - 0.40, 0.78, body, size=8, color=TEXT,
             line_spacing=1.25)
        x += w + gap

    text(s, M, 2.96, SW - 2 * M, 0.24, 'ЩО РОБИТИ ДАЛІ', size=8, color=ACCENT,
         bold=True)
    steps = [
        ('Зафіксувати сценарій логістики',
         'Погодити з відділом ЗЕД, який сценарій береться за базу для прайсу — '
         'зараз у розрахунку їх чотири, і різниця між ними більша за саму націнку.'),
        ('Уточнити роль ПДВ',
         'У підсумкову собівартість включено ПДВ 20%. Якщо він відшкодовується, '
         'база для націнки нижча — це треба підтвердити з фінансовим відділом.'),
        ('Дорахувати решту прайсу',
         'Собівартість є для 28 SKU із приблизно 105 позицій прайсів. Пріоритет — '
         'Thai-Nichi (28 позицій) і Nature Best / KoriKo (розрахунку немає взагалі).'),
        ('Звести курс до одного значення',
         'У файлі застосовано курси від 45 до 49 ₴ за $. Для порівнянного прайсу '
         'потрібен єдиний плановий курс.'),
    ]
    y = 3.26
    for i, (title, body) in enumerate(steps):
        col, row = i % 2, i // 2
        x = M + col * (4.68 + 0.24)
        yy = y + row * 0.92
        rect(s, x, yy, 4.68, 0.82, fill=CREAM, radius=0.08, line=LINE)
        rect(s, x, yy, 0.30, 0.82, fill=PRIMARY, radius=0.08)
        rect(s, x + 0.15, yy, 0.15, 0.82, fill=PRIMARY)
        text(s, x, yy, 0.30, 0.82, str(i + 1), size=12, font=HEAD, bold=True,
             color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(s, x + 0.44, yy + 0.10, 4.10, 0.20, title, size=9, font=HEAD, bold=True,
             color=TEXT)
        text(s, x + 0.44, yy + 0.32, 4.10, 0.44, body, size=7.5, color=MUTED,
             line_spacing=1.18)
    close(s)


def slide_sources(prs):
    s = blank(prs)
    page_bg(s, INK)
    rect(s, 0, 0, 0.10, SH, fill=ACCENT)
    text(s, M + 0.3, 1.10, 6.0, 0.24, 'ДЖЕРЕЛА ТА ПРИПУЩЕННЯ', size=9, color=GOLD,
         bold=True)
    text(s, M + 0.3, 1.42, 8.4, 0.60, 'Звідки взяті цифри', size=26, font=HEAD,
         bold=True, color=WHITE)
    rect(s, M + 0.3, 2.16, 1.5, 0.035, fill=ACCENT)

    items = [
        ('SelfCost.xlsx', 'Собівартість одиниці товару в $ і ₴ — колонки «ИТОГО С/С, € ед. '
                          'Товара» та «ИТОГО С/С, UAH ед. Товара», вкладки Singha Kameda, '
                          'Thai-Nichi, TMK, ZEK. Вкладка «Koriko» порожня.'),
        ('Snacks. Price Comparison.xlsx', 'Фото продуктів, ваги, фасовки, терміни '
                                          'придатності, ціни FOB і місткість контейнерів.'),
        ('Припущення деки', 'Підсумкова колонка собівартості підписана «€», але вхідні '
                            'ціни — FOB у доларах без перерахунку в євро, тому показник '
                            'наведено як $. Смаки на фото KOKIRI Wow розібрані за '
                            'кольоровим кодуванням бренду: зелений — Original, '
                            'червоний — Spicy, фіолетовий — Squid.'),
    ]
    y = 2.44
    for title, body in items:
        rect(s, M + 0.3, y + 0.05, 0.10, 0.10, fill=GOLD, radius=0.5)
        text(s, M + 0.52, y - 0.02, 8.2, 0.22, title, size=10, font=HEAD, bold=True,
             color=WHITE)
        text(s, M + 0.52, y + 0.22, 8.2, 0.56, body, size=8, color=SOFT,
             line_spacing=1.25)
        y += 0.94
    _page[0] += 1
    text(s, SW - M - 0.6, SH - 0.32, 0.6, 0.20, str(_page[0]), size=7.5, color=GOLD,
         bold=True, align=PP_ALIGN.RIGHT)


def main():
    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(SH)

    slide_cover(prs)
    slide_overview(prs)
    slide_method(prs)
    slide_logistics(prs)
    for sup in SUPPLIERS:
        slide_profile(prs, sup)
        slides_products(prs, sup)
        slides_cost_table(prs, sup)
    slide_koriko(prs)
    slide_ranking(prs)
    slide_conclusions(prs)
    slide_sources(prs)

    prs.save(OUT)
    print('saved %s — %d slides' % (OUT, len(prs.slides._sldIdLst)))


if __name__ == '__main__':
    main()
