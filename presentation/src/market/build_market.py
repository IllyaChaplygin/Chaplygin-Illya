# -*- coding: utf-8 -*-
"""Market research deck: seasoned seaweed snacks & rice-cracker-norimaki in
Ukraine, October 2026. A separate artifact from the supplier catalogue deck —
never writes to it. Reuses the same design system (theme.py) so it reads as
the same company's work, in the structure of the earlier SIAS/rice studies.
"""
import json
import os
import re
import statistics
import sys
from collections import defaultdict

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches

HERE = os.path.dirname(__file__)
sys.path.insert(0, os.path.join(HERE, '..', 'deck'))
sys.path.insert(0, HERE)

from theme import (BODY, BODY_TX, CARD_R, GOLD, HEAD, INK, M, MUTED, PAPER,  # noqa: E402
                   RULE, SH, SUPPLIER_COLORS, SW, WHITE, blank, deepen, footer,
                   header, hline, label, mix, rect, text, tint)
from charts import hbar_chart, price_ladder  # noqa: E402
from price_model import SENSITIVITY_MARGINS as MARGINS  # noqa: E402
import market_data as MD  # noqa: E402

OUT = os.path.join(HERE, '..', '..', 'Дослідження_ринку_снеків_з_водоростей.pptx')

MODEL = json.load(open(os.path.join(HERE, 'price_model.json'), encoding='utf-8'))
ROWS = [r for r in MODEL['rows'] if r['is_cheapest']]
ALL_ROWS = MODEL['rows']
PARAMS = MODEL['params']
SENS = MODEL['sensitivity']

SUP_ORDER = ['singha', 'thainichi', 'tmk', 'zek']
SUP_NAME = {'singha': 'Singha Kameda', 'thainichi': 'Thai-Nichi',
           'tmk': 'TMK · KOKIRI', 'zek': 'HanJin · ZEK'}
COMPETITOR = RGBColor(0x5C, 0x5F, 0x66)

_page = [0]


def close(slide, note=None):
    _page[0] += 1
    footer(slide, _page[0], note)


def grams(unit):
    m = re.search(r'([\d,]+)\s*(г|кг)', unit)
    if not m:
        return None
    v = float(m.group(1).replace(',', '.'))
    return v * 1000 if m.group(2) == 'кг' else v


def per_gram_by_supplier():
    by_sup = defaultdict(list)
    for r in ROWS:
        g = grams(r['unit'])
        if g:
            by_sup[r['supplier_id']].append(r['shelf_uah'] / g)
    return by_sup


def stat_tile(s, x, y, w, h, value, cap, accent):
    rect(s, x, y, w, h, fill=WHITE, radius=0.06, line=RULE)
    rect(s, x, y, w, 0.05, fill=accent)
    text(s, x, y + 0.16, w, 0.34, value, size=18, font=HEAD, bold=True, color=INK,
         align=PP_ALIGN.CENTER)
    text(s, x + 0.08, y + 0.52, w - 0.16, 0.38, cap, size=6.8, color=MUTED,
         align=PP_ALIGN.CENTER, line_spacing=1.12)


# ===================================================================== slides ===
def slide_cover(prs):
    s = blank(prs, accent=SUPPLIER_COLORS['singha'])
    seg = SW / len(SUP_ORDER)
    for i, sid in enumerate(SUP_ORDER):
        rect(s, i * seg, 0, seg, 0.07, fill=SUPPLIER_COLORS[sid])

    label(s, M, 0.60, 8.6, 'ДОСЛІДЖЕННЯ РИНКУ ТА ПОЗИЦІОНУВАННЯ · ЖОВТЕНЬ 2026',
         color=GOLD, size=9)
    text(s, M, 0.90, 9.2, 1.50, 'Снеки з водоростей і рисові крекери\nв Україні',
         size=30, font=HEAD, color=INK, bold=True, line_spacing=1.10)
    rect(s, M, 2.26, 1.10, 0.05, fill=SUPPLIER_COLORS['zek'])
    text(s, M, 2.48, 8.6, 0.6,
        'Хто вже стоїть на полиці, по чому продають і де в цій картині опиняється '
        'наш асортимент — 4 постачальники, 28 SKU.',
        size=11.5, color=BODY_TX, line_spacing=1.34)

    tiles = [('4', 'реальні позиції\nконкурентів знайдено'), ('6', 'мереж\nперевірено'),
            ('0', 'мереж продають\nкатегорію'), ('28', 'SKU в нашому\nасортименті')]
    tw, gap = (SW - 2 * M - 0.16 * 3) / 4, 0.16
    ty = 3.45
    for i, (v, cap) in enumerate(tiles):
        x = M + i * (tw + gap)
        stat_tile(s, x, ty, tw, 0.92, v, cap, SUPPLIER_COLORS[SUP_ORDER[i]])
    close(s)


def slide_methodology(prs):
    accent = SUPPLIER_COLORS['tmk']
    s = blank(prs, accent=accent)
    header(s, 'Методологія', eyebrow='ОБСЯГ ДОСЛІДЖЕННЯ', accent=accent)
    text(s, M, 1.26, SW - 2 * M, 0.55, MD.METHOD_NOTE, size=10.5, color=BODY_TX,
         line_spacing=1.38)

    colw = (SW - 2 * M - 0.34) / 2
    label(s, M, 2.18, colw, 'Що зроблено', color=deepen(accent, 0.9), size=8)
    done = ['Перевірено онлайн-каталоги 6 мереж: АТБ, Сільпо, Novus, Varus, '
           'Ашан, Metro',
           'Перевірено маркетплейси: Prom.ua, Rozetka, Bigl.ua, Allo.ua',
           'Кожна знайдена позиція — з ціною, вагою, продавцем і каналом',
           'Власна собівартість → ціна партнеру → полиця — за формулою '
           'проєкту «Готовий рис» (той самий розподіл: бонус 25%, маржа 35%, '
           'націнка мережі +40%)']
    y = 2.46
    for d in done:
        text(s, M, y, colw, 0.62, '•  ' + d, size=9, color=BODY_TX, line_spacing=1.26)
        y += 0.62

    x2 = M + colw + 0.34
    label(s, x2, 2.18, colw, 'Чого немає в цьому дослідженні', color=deepen(accent, 0.9),
         size=8)
    missing = ['Митної бази (обсяги імпорту категорії) — на відміну від '
              'досліджень SIAS і рису, доступу до цих даних не було',
              'Офлайн-перевірки полиць наживо — лише онлайн-каталоги мереж',
              'Повного асортименту конкурентів — тільки те, що показує '
              'відкритий пошук']
    y = 2.46
    for d in missing:
        text(s, x2, y, colw, 0.70, '•  ' + d, size=9, color=BODY_TX, line_spacing=1.26)
        y += 0.74
    rect(s, x2, y + 0.05, colw, 0.46, fill=tint(accent, 0.10), radius=0.06)
    text(s, x2 + 0.12, y + 0.13, colw - 0.24, 0.32,
        'Де написано «не знайдено» — йдеться про видимість в онлайн-пошуку, '
        'а не доведену відсутність на ринку.', size=7.3, italic=True, color=deepen(accent, 0.92),
        line_spacing=1.2)
    close(s)


def slide_segmentation(prs):
    accent = SUPPLIER_COLORS['tmk']
    s = blank(prs, accent=accent)
    header(s, 'Сегментація категорії', eyebrow='ЩО ШУКАЛИ', accent=accent)

    segs = [
        dict(name='Сегмент A', sub='Приправлені снеки з водоростей',
             match='Відповідає TMK · KOKIRI та HanJin · ZEK', found='3', unit='бренди',
             found2='4', unit2='позиції з ціною', accent=SUPPLIER_COLORS['tmk'],
             note='Tao Kae Noi, AKURA (2 позиції), гуртова «Seaweed Traditional».'),
        dict(name='Сегмент B', sub='Рисові крекери в норі (арарe, сенбей)',
             match='Відповідає Singha Kameda та Thai-Nichi', found='0', unit='брендів',
             found2='0', unit2='позицій з ціною', accent=SUPPLIER_COLORS['singha'],
             note=MD.SEGMENT_B_NOTE),
    ]
    colw = (SW - 2 * M - 0.30) / 2
    for i, seg in enumerate(segs):
        x = M + i * (colw + 0.30)
        rect(s, x, 1.30, colw, 3.70, fill=WHITE, radius=CARD_R, line=RULE)
        rect(s, x, 1.30, colw, 0.06, fill=seg['accent'])
        label(s, x + 0.26, 1.54, colw - 0.5, seg['name'], color=deepen(seg['accent'], 0.9),
             size=8)
        text(s, x + 0.26, 1.74, colw - 0.5, 0.56, seg['sub'], size=15, font=HEAD,
             bold=True, color=INK, line_spacing=1.1)
        text(s, x + 0.26, 2.30, colw - 0.5, 0.26, seg['match'], size=8.5, color=MUTED,
             italic=True)
        fy = 2.68
        fw = (colw - 0.52 - 0.14) / 2
        for j, (val, cap) in enumerate([(seg['found'], seg['unit']), (seg['found2'], seg['unit2'])]):
            fx = x + 0.26 + j * (fw + 0.14)
            rect(s, fx, fy, fw, 0.82, fill=tint(seg['accent'], 0.08), radius=0.06)
            text(s, fx, fy + 0.12, fw, 0.36, val, size=22, font=HEAD, bold=True,
                 color=deepen(seg['accent'], 0.95), align=PP_ALIGN.CENTER)
            text(s, fx + 0.06, fy + 0.50, fw - 0.12, 0.28, cap, size=7, color=MUTED,
                 align=PP_ALIGN.CENTER, line_spacing=1.1)
        text(s, x + 0.26, 3.70, colw - 0.52, 1.20, seg['note'], size=8.7, color=BODY_TX,
             line_spacing=1.32)
    close(s, note='Джерело: відкритий моніторинг маркетплейсів і мереж, жовтень 2026.')


def slide_competitors(prs):
    accent = SUPPLIER_COLORS['tmk']
    s = blank(prs, accent=accent)
    header(s, 'Конкуренти · каталог', eyebrow='СЕГМЕНТ A · ЗНАЙДЕНО В ВІДКРИТОМУ ПОШУКУ',
          accent=accent, tag='4 позиції · 3 бренди')

    cards = MD.COMPETITORS
    n = len(cards)
    gap = 0.20
    w = (SW - 2 * M - gap * (n - 1)) / n
    y, h = 1.30, 3.56
    for i, c in enumerate(cards):
        x = M + i * (w + gap)
        rect(s, x, y, w, h, fill=WHITE, radius=CARD_R, line=RULE)
        rect(s, x, y, w, 0.90, fill=tint(COMPETITOR, 0.08))
        rect(s, x, y, w, 0.90, fill=None, radius=CARD_R)
        text(s, x + 0.16, y + 0.14, w - 0.32, 0.30, c['brand'], size=13, font=HEAD,
             bold=True, color=INK, line_spacing=1.0)
        text(s, x + 0.16, y + 0.48, w - 0.32, 0.36, c['origin'], size=7.3, color=MUTED,
             line_spacing=1.1)
        text(s, x + 0.16, y + 1.04, w - 0.32, 0.46, c['product'], size=10.5, font=HEAD,
             bold=True, color=deepen(COMPETITOR, 0.85), line_spacing=1.12)
        label(s, x + 0.16, y + 1.54, w - 0.32, c['pack'], color=MUTED, size=6.6)

        py = y + 1.84
        rect(s, x + 0.16, py, w - 0.32, 0.56, fill=tint(COMPETITOR, 0.06), radius=0.05)
        text(s, x + 0.16, py + 0.06, w - 0.32, 0.30,
            '%.0f ₴' % c['price_uah'], size=17, font=HEAD, bold=True, color=INK,
            align=PP_ALIGN.CENTER)
        if c.get('price_was'):
            text(s, x + 0.16, py + 0.36, w - 0.32, 0.18,
                'база %.0f ₴ · %.1f ₴/г' % (c['price_was'], c['per_g']), size=6.3,
                color=MUTED, align=PP_ALIGN.CENTER)
        ny = py + 0.68
        text(s, x + 0.16, ny, w - 0.32, 0.22, c['channel'], size=7.2, bold=True,
            color=deepen(COMPETITOR, 0.9), line_spacing=1.1)
        text(s, x + 0.16, ny + 0.24, w - 0.32, 0.18, 'продавець: ' + c['seller'],
            size=6.6, color=MUTED, italic=True)
        text(s, x + 0.16, ny + 0.48, w - 0.32, h - (ny + 0.48 - y) - 0.14, c['note'],
            size=7.4, color=BODY_TX, line_spacing=1.22)
    close(s, note='Ціни й продавці — як показані на момент перевірки, жовтень 2026.')


def slide_channels(prs):
    accent = SUPPLIER_COLORS['zek']
    s = blank(prs, accent=accent)
    header(s, 'Де продається', eyebrow='КАНАЛИ', accent=accent)
    text(s, M, 1.26, SW - 2 * M, 0.5, MD.CHANNEL_SUMMARY, size=10.5, color=BODY_TX,
         line_spacing=1.36)

    label(s, M, 2.10, 5.0, 'Перевірені мережі · позицій категорії знайдено',
         color=deepen(accent, 0.9), size=8)
    rows = [dict(name=c['name'], value=0.02, sub=None, color=mix(WHITE, MUTED, 0.5))
           for c in MD.CHAINS_CHECKED]
    hbar_chart(s, M, 2.34, 5.25, 2.30, rows, max_val=1, value_fmt=lambda v: '0')

    x2 = M + 5.25 + 0.35
    label(s, x2, 2.10, SW - M - x2, 'Де категорія реально є', color=deepen(accent, 0.9),
         size=8)
    markets = [('Prom.ua', 'маркетплейс · найбільше позицій знайдено тут'),
              ('Rozetka', 'маркетплейс · продавець — сама Rozetka (AKURA)'),
              ('Bigl.ua / Allo.ua', 'маркетплейси · перевірено, окремих позицій не '
                                     'зафіксовано')]
    y = 2.40
    for name, note in markets:
        rect(s, x2, y, SW - M - x2, 0.74, fill=WHITE, radius=0.06, line=RULE)
        rect(s, x2, y, 0.05, 0.74, fill=accent)
        text(s, x2 + 0.18, y + 0.10, SW - M - x2 - 0.3, 0.26, name, size=10.5,
             font=HEAD, bold=True, color=INK)
        text(s, x2 + 0.18, y + 0.38, SW - M - x2 - 0.3, 0.32, note, size=7.6,
             color=MUTED, line_spacing=1.18)
        y += 0.88
    close(s)


def slide_price_per_gram(prs):
    accent = SUPPLIER_COLORS['singha']
    s = blank(prs, accent=accent)
    header(s, 'Ціна за грам — конкуренти проти нашого асортименту',
          eyebrow='ПОЗИЦІОНУВАННЯ', accent=accent)

    by_sup = per_gram_by_supplier()
    rows = []
    for c in MD.COMPETITORS:
        if 'Опт' in c['product']:
            continue
        rows.append(dict(name=c['brand'], sub=c['product'][:26], value=c['per_g'],
                         color=COMPETITOR))
    for sid in SUP_ORDER:
        vals = by_sup[sid]
        rows.append(dict(name=SUP_NAME[sid],
                         sub='медіана полиці · %d SKU' % len(vals),
                         value=statistics.median(vals), color=SUPPLIER_COLORS[sid]))
    rows.sort(key=lambda r: -r['value'])
    max_val = max(r['value'] for r in rows) * 1.12

    hbar_chart(s, M, 1.34, SW - 2 * M, 3.55, rows, max_val,
              value_fmt=lambda v: '%.1f ₴/г' % v, name_w=2.75, val_w=0.95)
    close(s, note='Наша ціна — розрахункова полиця за формулою нижче, не чинний прайс. '
                  'Конкуренти — реальні ціни з відкритого пошуку.')


def slide_pricing_formula(prs):
    accent = SUPPLIER_COLORS['tmk']
    s = blank(prs, accent=accent)
    header(s, 'Ціноутворення', eyebrow='З ЧОГО СКЛАДАЄТЬСЯ ЦІНА НА ПОЛИЦІ', accent=accent)

    text(s, M, 1.22, SW - 2 * M, 0.26,
        'ЦІНА ПАРТНЕРУ — ЗА НЕЮ МИ ПРОДАЄМО МЕРЕЖІ (100 %)', size=9, bold=True,
        color=deepen(accent, 0.9))
    parts = [('Собівартість', 1 - PARAMS['bonus'] - PARAMS['margin'], mix(WHITE, accent, 0.18)),
            ('Бонус мережі', PARAMS['bonus'], mix(WHITE, accent, 0.45)),
            ('Наша маржа', PARAMS['margin'], accent)]
    bx, bw, by, bh = M, SW - 2 * M, 1.50, 0.46
    cx = bx
    for name, share, color in parts:
        pw = bw * share
        rect(s, cx, by, pw, bh, fill=color, radius=0.03)
        text(s, cx, by, pw, bh, '%s %.0f%%' % (name, share * 100), size=9, bold=True,
             color=WHITE if share > 0.12 else deepen(accent, 0.9),
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        cx += pw
    text(s, bx, by + bh + 0.10, bw, 0.24,
        'ПОЛИЦЯ = ціна партнеру × %.2f. Бонус %d%% і маржа %d%% — частки ціни '
        'партнеру; націнка мережі %d%% — понад ціну партнеру.'
        % (PARAMS['markup'], PARAMS['bonus'] * 100, PARAMS['margin'] * 100,
           round((PARAMS['markup'] - 1) * 100)),
        size=8, italic=True, color=MUTED)

    ty = by + bh + 0.46
    cols = [('ТОВАР', 1.95), ('СС, ₴', 0.72), ('ПАРТНЕР, ₴', 0.95), ('БОНУС, ₴', 0.80),
           ('ПРИБУТОК, ₴', 0.95), ('ПОЛИЦЯ, ₴', 0.95)]
    cx = M
    for name, w in cols:
        label(s, cx, ty, w, name, color=MUTED, size=6.6)
        cx += w
    hline(s, M, ty + 0.20, sum(w for _, w in cols))

    reps = []
    seen = set()
    for r in ROWS:
        if r['supplier_id'] in seen:
            continue
        seen.add(r['supplier_id'])
        reps.append(r)
    ry = ty + 0.28
    for r in reps:
        accent_r = SUPPLIER_COLORS[r['supplier_id']]
        cx = M
        vals = [r['title'][:26], '%.1f' % r['cost_uah'], '%.1f' % r['partner_uah'],
               '%.1f' % r['bonus_uah'], '%.1f' % r['profit_uah'], '%.1f' % r['shelf_uah']]
        for i, (v, (_, w)) in enumerate(zip(vals, cols)):
            text(s, cx, ry, w - 0.06, 0.30, v, size=8.3,
                 font=HEAD if i == 0 or i == 5 else BODY,
                 bold=(i == 0 or i == 5), color=INK if i != 5 else deepen(accent_r, 0.92))
            cx += w
        ry += 0.34
    close(s, note='СС — Self_Cost (SelfCost_Snacks.xlsx, найдешевший контейнерний '
                  'сценарій). Формула — як у проєкті «Готовий рис».')


def competitor_median_per_g():
    vals = [c['per_g'] for c in MD.COMPETITORS if 'Опт' not in c['product']]
    return statistics.median(vals)


def slide_shelf(prs, sup_id):
    accent = SUPPLIER_COLORS[sup_id]
    s = blank(prs, accent=accent)
    rows = [r for r in ROWS if r['supplier_id'] == sup_id]
    header(s, 'Полиця — %s' % SUP_NAME[sup_id], eyebrow='РОЗРАХУНКОВА ЦІНА НА ПОЛИЦІ',
          accent=accent, tag='%d SKU' % len(rows))

    cols = [('ТОВАР', 3.00), ('ФАСОВКА', 1.15), ('СС, ₴', 0.80), ('ПАРТНЕР, ₴', 1.00),
           ('ПОЛИЦЯ, ₴', 1.00), ('ПОЛИЦЯ, ₴/г', 1.15)]
    ty = 1.30
    cx = M
    for name, w in cols:
        label(s, cx, ty, w, name, color=deepen(accent, 0.9), size=6.8)
        cx += w
    hline(s, M, ty + 0.22, sum(w for _, w in cols), color=accent)

    ry = ty + 0.32
    row_h = min(0.295, (5.00 - ry) / max(len(rows), 1))
    for i, r in enumerate(rows):
        if i % 2 == 1:
            rect(s, M, ry - 0.02, sum(w for _, w in cols), row_h, fill=tint(accent, 0.05))
        g = grams(r['unit'])
        per_g = r['shelf_uah'] / g if g else None
        vals = [r['title'], r['unit'], '%.2f' % r['cost_uah'], '%.1f' % r['partner_uah'],
               '%.1f' % r['shelf_uah'], ('%.2f' % per_g) if per_g else '—']
        cx = M
        for j, (v, (_, w)) in enumerate(zip(vals, cols)):
            text(s, cx, ry, w - 0.06, row_h, v, size=8.1,
                 bold=(j == 0 or j == 4), color=INK if j != 4 else deepen(accent, 0.92),
                 anchor=MSO_ANCHOR.MIDDLE)
            cx += w
        ry += row_h

    if len(rows) <= 4:
        by_sup = per_gram_by_supplier()
        med = statistics.median(by_sup[sup_id])
        comp_med = competitor_median_per_g()
        diff = (med / comp_med - 1) * 100
        verdict = 'дешевше за' if diff < 0 else 'дорожче за'
        rect(s, M, ry + 0.30, sum(w for _, w in cols), 0.56, fill=tint(accent, 0.09),
             radius=0.06)
        text(s, M + 0.18, ry + 0.40, sum(w for _, w in cols) - 0.36, 0.40,
            'Медіана полиці — %.1f ₴/г, це на %.0f%% %s медіану знайдених '
            'конкурентів (%.1f ₴/г).' % (med, abs(diff), verdict, comp_med),
            size=8.6, color=deepen(accent, 0.92), line_spacing=1.2)

    close(s, note='Партнер і полиця — за формулою «Собівартість → партнер → полиця» '
                  '(бонус %d%%, маржа %d%%, націнка +%d%%). Собівартість — найдешевший '
                  'контейнерний сценарій із SelfCost_Snacks.xlsx.'
                  % (PARAMS['bonus'] * 100, PARAMS['margin'] * 100,
                     round((PARAMS['markup'] - 1) * 100)))


def slide_positioning(prs):
    accent = SUPPLIER_COLORS['zek']
    s = blank(prs, accent=accent)
    header(s, 'Позиціонування', eyebrow='ХТО З КИМ СТОЇТЬ', accent=accent)
    text(s, M, 1.24, SW - 2 * M, 0.42,
        'Медіанна полиця кожного постачальника поруч із реальними конкурентами, '
        '₴ за грам.', size=9.5, color=MUTED, italic=True)

    by_sup = per_gram_by_supplier()
    points = []
    for c in MD.COMPETITORS:
        if 'Опт' in c['product']:
            continue
        points.append(dict(label=c['brand'], value=c['per_g'], color=COMPETITOR))
    for sid in SUP_ORDER:
        points.append(dict(label=SUP_NAME[sid], value=statistics.median(by_sup[sid]),
                           color=SUPPLIER_COLORS[sid]))
    lo = 0
    hi = max(p['value'] for p in points) * 1.15
    price_ladder(s, M + 0.2, 1.80, SW - 2 * M - 0.4, 2.90, points, lo, hi)

    insight = ('ZEK і рисові крекери (Singha, Thai-Nichi) виходять на 2–2,5× дешевше '
              'за знайдених конкурентів. Малі пакетики TMK · KOKIRI (2,5–12 г) — '
              'навпаки, дорожчі за грам, ніж AKURA і Tao Kae Noi: фіксована логістика '
              'на маленький пакетик з’їдає перевагу в собівартості.')
    rect(s, M, 4.78, SW - 2 * M, 0.42, fill=tint(accent, 0.10), radius=0.05)
    text(s, M + 0.16, 4.85, SW - 2 * M - 0.32, 0.30, insight, size=8.2,
        color=deepen(accent, 0.92), line_spacing=1.2)
    close(s)


def slide_sensitivity(prs):
    accent = SUPPLIER_COLORS['thainichi']
    s = blank(prs, accent=accent)
    header(s, 'Чутливість маржі', eyebrow='ЯКА ПОЛИЦЯ ПРИ ІНШІЙ МАРЖІ', accent=accent,
          tag='бонус мережі незмінний — 25%')

    cols = [('ТОВАР', 2.55)] + [('%d%%' % int(m * 100), 1.33) for m in MARGINS]
    ty = 1.32
    cx = M
    for name, w in cols:
        label(s, cx, ty, w, name, color=deepen(accent, 0.9), size=7)
        cx += w
    hline(s, M, ty + 0.22, sum(w for _, w in cols), color=accent)

    ry = ty + 0.32
    row_h = 0.44
    reps = []
    seen = set()
    for r in SENS:
        if r['supplier_id'] in seen:
            continue
        seen.add(r['supplier_id'])
        reps.append(r)
    for i, r in enumerate(reps):
        if i % 2 == 1:
            rect(s, M, ry - 0.03, sum(w for _, w in cols), row_h, fill=tint(accent, 0.05))
        cx = M
        text(s, cx, ry, cols[0][1] - 0.06, row_h, r['title'], size=8.4, bold=True,
             color=INK, anchor=MSO_ANCHOR.MIDDLE)
        cx += cols[0][1]
        for j, m in enumerate(MARGINS):
            shelf = r['shelf_by_margin'][str(m)]
            text(s, cx, ry, cols[j + 1][1] - 0.06, row_h, '%.0f ₴' % shelf, size=8.6,
                font=HEAD, bold=(j == 0), color=deepen(accent, 0.9) if j == 0 else INK,
                anchor=MSO_ANCHOR.MIDDLE)
            cx += cols[j + 1][1]
        ry += row_h
    close(s, note='Кожен товар — у своєму найдешевшому контейнерному сценарії. '
                  '35% — базова маржа цього дослідження (виділена колонка).')


def slide_conclusions(prs):
    accent = SUPPLIER_COLORS['zek']
    s = blank(prs, accent=accent)
    header(s, 'Висновки', eyebrow='ЩО З ЦИМ РОБИТИ', accent=accent)

    points = [
        ('Категорія майже порожня', 'Приправлені снеки з водоростей — лише 3 бренди, 4 '
         'позиції на весь відкритий ринок; рисові крекери в норі — 0. Жодна з 6 '
         'перевірених мереж не продає категорію взагалі.'),
        ('ZEK і рисові крекери — дешевше за ринок', 'Розрахункова полиця ZEK, Singha '
         'Kameda і Thai-Nichi виходить у 2–2,5 рази нижче знайдених конкурентів за грам '
         '— є запас і на маржу, і на вхід у мережу.'),
        ('Малі пакетики TMK · KOKIRI потребують перегляду', 'Формат 2,5–12 г програє '
         'конкурентам за ціною на грам через фіксовану логістику на одиницю — варто '
         'рахувати більші фасовки або інший контейнерний сценарій.'),
        ('Де шукати конкурентів офлайн', 'AKURA (продавець — сама Rozetka) і Tao Kae Noi '
         '— єдині реальні орієнтири. Обидва живуть на маркетплейсах, не в мережах: вхід '
         'через мережу був би першим для категорії.'),
    ]
    y = 1.32
    for i, (t, d) in enumerate(points):
        rect(s, M, y, 0.30, 0.30, fill=tint(accent, 0.12), radius=0.5)
        text(s, M, y, 0.30, 0.30, str(i + 1), size=12, font=HEAD, bold=True,
             color=deepen(accent, 0.92), align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(s, M + 0.46, y - 0.02, SW - 2 * M - 0.46, 0.28, t, size=12, font=HEAD,
             bold=True, color=INK)
        text(s, M + 0.46, y + 0.30, SW - 2 * M - 0.46, 0.60, d, size=9, color=BODY_TX,
             line_spacing=1.3)
        y += 0.98
    close(s)


def main():
    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(SH)

    slide_cover(prs)
    slide_methodology(prs)
    slide_segmentation(prs)
    slide_competitors(prs)
    slide_channels(prs)
    slide_price_per_gram(prs)
    slide_pricing_formula(prs)
    for sid in SUP_ORDER:
        slide_shelf(prs, sid)
    slide_positioning(prs)
    slide_sensitivity(prs)
    slide_conclusions(prs)

    prs.save(OUT)
    print('saved %s — %d slides' % (OUT, len(prs.slides)))


if __name__ == '__main__':
    main()
