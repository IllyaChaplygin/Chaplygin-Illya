# -*- coding: utf-8 -*-
"""Cross-segment market maps and our own price slides (formula from the rice study)."""
import collections, os
from md import *  # noqa
from md import K, rect, text, caps, picture, kpi, section_tag, hbars, footnote, rgb
import mdata as D
from mdata import SEGS, rows_of, lines, med, gfmt, uah, rng, CHAIN, NATIONAL, ROWS, OUR, PM
from slides_market import CHAINS_ORDER, src_note, HERE

BON, MAR, MUP = PM['params']['bonus'], PM['params']['margin'], PM['params']['markup']
NORI_SEGS = ('mini', 'chips', 'tempura', 'ricecr')
SHORT = {'megamarket': 'Мега-\nМаркет', 'ultramarket': 'Ultra-\nmarket', 'chudomarket': 'Чудо-\nМаркет', 'epicentr': 'Епі-\nцентр', 'vostorg': 'Вос-\nторг', 'tavriav': 'Таврія\nВ'}
SUP_NAME = {'singha': 'Singha Kameda', 'thainichi': 'Thai-Nichi', 'tmk': 'TMK · KOKIRI', 'zek': 'HanJin · ZEK'}
SEG_OF = {o['key']: o['seg'] for o in OUR}


def pct(a, b):
    return '%+d %%' % round((a / b - 1) * 100)


def fnum(v, d=2):
    return ('%.' + str(d) + 'f') % v if False else (('%.' + str(d) + 'f') % v).replace('.', ',')


# ================================================================= market maps
def brand_network(deck):
    s = deck.slide('Зведення · бренди × мережі', 'Усі бренди в усіх мережах')
    rs = [r for r in ROWS if r['seg'] in NORI_SEGS]
    brands = collections.Counter(r['brand'] for r in rs).most_common(13)
    chains = [c for c in CHAINS_ORDER]
    x0, y0, nw, cw, rh = 0.62, 1.72, 1.5, 0.585, 0.37
    text(s, x0, 1.36, 12, 0.2, 'Верхнє число — кількість позицій бренду в мережі; нижнє — медіана ціни упаковки, ₴. Чим темніше, тим більше SKU.', size=8.6, color=MUTED)
    for j, c in enumerate(chains):
        x = x0 + nw + j * cw
        rect(s, x + 0.02, y0, cw - 0.04, 0.42, fill=NAVY if c in NATIONAL else TEAL)
        text(s, x, y0 + 0.02, cw, 0.38, SHORT.get(c, CHAIN[c]), size=6.2, color=WHITE, bold=True, align='c', anchor='m')
    xs = x0 + nw + len(chains) * cw + 0.06
    for k, h in enumerate(['SKU', 'МЕРЕЖ']):
        rect(s, xs + k * 0.52, y0, 0.5, 0.42, fill=GOLD)
        text(s, xs + k * 0.52, y0 + 0.02, 0.5, 0.38, h, size=6.4, color=WHITE, bold=True, align='c', anchor='m')
    mx = 6
    for i, (b, n) in enumerate(brands):
        y = y0 + 0.48 + i * rh
        text(s, x0, y, nw - 0.1, rh, b, size=8.2, color=NAVY, bold=True, anchor='m', align='r')
        brs = [r for r in rs if r['brand'] == b]
        for j, c in enumerate(chains):
            here = [r for r in brs if c in r['chains']]
            x = x0 + nw + j * cw
            if not here:
                rect(s, x + 0.02, y + 0.02, cw - 0.04, rh - 0.04, fill=rgb('F6F7FB'))
                text(s, x, y, cw, rh, '–', size=7.2, color=GREY, align='c', anchor='m')
                continue
            t = min(1, len(here) / mx)
            rect(s, x + 0.02, y + 0.02, cw - 0.04, rh - 0.04, fill=mixc(NAVY, rgb('E4E9F4'), 0.25 + 0.65 * t))
            fg = WHITE if t > 0.45 else NAVY
            text(s, x, y + 0.01, cw, rh * 0.55, str(len(here)), size=8.4, color=fg, bold=True, align='c')
            text(s, x, y + rh * 0.52, cw, rh * 0.4, '%d' % round(med(r['pmed'] for r in here)), size=6.2, color=fg, align='c')
        text(s, xs, y, 0.5, rh, str(n), size=8.2, color=NAVY, bold=True, align='c', anchor='m')
        text(s, xs + 0.52, y, 0.5, rh, str(len({c for r in brs for c in r['chains']})), size=8.2, color=NAVY, bold=True, align='c', anchor='m')
    src_note(s)
    return s


def weight_map(deck):
    s = deck.slide('Формат × вага', 'Де насправді щільна полиця: вага упаковки')
    rs = [r for r in ROWS if r['seg'] in NORI_SEGS]
    bins = [(0, 5.5, 'до 5 г'), (5.5, 16, '8–15 г'), (16, 31, '20–30 г'), (31, 61, '35–60 г'), (61, 999, '70 г і більше')]
    total = len(rs)
    text(s, 0.62, 1.36, 12.1, 0.22, 'Кількість позицій і медіана ціни упаковки за ваговими смугами. Золотою позначено смуги, де є наші SKU.', size=9.2, color=MUTED)
    caps(s, 0.62, 1.78, 4, 'Позицій у ваговій смузі', size=7)
    caps(s, 7.3, 1.78, 5, 'Медіана ціни упаковки, ₴ · медіана ₴/г', size=7)
    maxn = max(sum(1 for r in rs if a <= r['grams'] < b) for a, b, _ in bins)
    mine_in = {}
    for a, b, lab in bins:
        sel = [r for r in rs if a <= r['grams'] < b]
        mine = [o for o in OUR if a <= o['grams'] < b]
        mine_in[lab] = mine
        i = [x[2] for x in bins].index(lab)
        y = 2.1 + i * 0.78
        if mine:
            rect(s, 0.55, y - 0.06, 12.25, 0.72, fill=rgb('FFF3DC'))
        text(s, 0.62, y, 1.3, 0.55, lab, size=9.4, color=NAVY, bold=True, anchor='m')
        w = 4.3 * len(sel) / maxn
        rect(s, 2.0, y + 0.12, max(0.06, w), 0.32, fill=NAVY)
        text(s, 2.0 + w + 0.1, y, 1.8, 0.55, '%d · %d %%' % (len(sel), round(100 * len(sel) / total)), size=8.6, color=MUTED, bold=True, anchor='m')
        mp = med(r['pmed'] for r in sel)
        w2 = 3.2 * min(mp, 200) / 200
        rect(s, 7.3, y + 0.12, w2, 0.32, fill=AMBER)
        text(s, 7.3 + w2 + 0.1, y, 2.3, 0.55, '%d ₴ · %s ₴/г' % (round(mp), fnum(med(r['per_g'] for r in sel), 1)), size=8.6, color=MUTED, bold=True, anchor='m')
        if mine:
            text(s, 11.0, y, 1.7, 0.55, 'наші SKU: %d' % len(mine), size=8, color=AMBER, bold=True, anchor='m', align='r')
    # insight boxes — numbers come from the data above
    stats = {}
    for a_, b_, lab in bins:
        sel = [r for r in rs if a_ <= r['grams'] < b_]
        stats[lab] = (len(sel), med(r['pmed'] for r in sel), med(r['per_g'] for r in sel))
    n5, n8, n20, n35 = stats['до 5 г'], stats['8–15 г'], stats['20–30 г'], stats['35–60 г']
    tm = [o for o in OUR if 8 <= o['grams'] <= 15]
    tmk_g = med(o['per_g'] for o in tm)
    y = 6.1
    boxes = [('ДЕ СТОЇТЬ НАШ АСОРТИМЕНТ',
              'Наші SKU є в усіх п’яти вагових смугах. Найщільніша полиця — до 5 г (%d SKU, %d %%); у смузі 35–60 г — %d позицій ринку й 10 наших SKU (Singha, Thai-Nichi, ZEK).' % (n5[0], round(100 * n5[0] / total), n35[0]), NAVY),
             ('ДЕ ВАГА ВЕДЕ ДО ЦІНИ',
              'Чим менша упаковка, тим дорожчий грам: до 5 г — %s ₴/г, 8–15 г — %s, 20–30 г — %s, 35–60 г — %s. Наші TMK 10–12 г — %s ₴/г, тобто у %s рази вище за медіану своєї смуги.' % (fnum(n5[2], 1), fnum(n8[2], 1), fnum(n20[2], 1), fnum(n35[2], 1), fnum(tmk_g, 1), fnum(tmk_g / n8[2], 1)), AMBER)]
    for k, (ttl, body, col) in enumerate(boxes):
        x = 0.62 + k * 6.15
        rect(s, x, y, 5.95, 0.92, fill=PANEL); rect(s, x, y, 0.04, 0.92, fill=col)
        text(s, x + 0.2, y + 0.1, 5.6, 0.16, ttl, size=7, color=col, bold=True, spc=0.2)
        text(s, x + 0.2, y + 0.3, 5.6, 0.6, body, size=7.8, color=INK)
    src_note(s)
    return s


def range_chart(s, x, y, w, rows, axis_max, step, unit_fmt, per='pack', row_h=0.86):
    """rows: dict(seg, lines, ours). Market range as a bar, line medians as dots, ours as supplier-coloured markers."""
    lab_w = 2.3
    tx, tw = x + lab_w, w - lab_w - 0.2
    def px(v):
        return tx + tw * min(v, axis_max) / axis_max
    # grid
    v = 0
    while v <= axis_max + 1e-9:
        rect(s, px(v), y, 0.01, row_h * len(rows) + 0.05, fill=RULE)
        text(s, px(v) - 0.3, y + row_h * len(rows) + 0.08, 0.6, 0.18, unit_fmt(v), size=7, color=MUTED, align='c')
        v += step
    for i, r in enumerate(rows):
        yy = y + i * row_h
        col = SEG_COLOR[r['seg']]
        if i % 2 == 0:
            rect(s, x, yy, w, row_h, fill=rgb('F8F9FC'))
        text(s, x + 0.1, yy + 0.12, lab_w - 0.2, 0.22, SEGS[r['seg']]['name'], size=8.4, color=col, bold=True)
        text(s, x + 0.1, yy + 0.36, lab_w - 0.2, 0.4, '%d лінійок · %d SKU' % (len(r['lines']), len(r['rows'])) + ('\nмедіана %s' % unit_fmt(r['median'])), size=7, color=MUTED)
        vals = [l['pmed'] if per == 'pack' else l['per_g'] for l in r['lines']]
        lo, hi = min(vals), max(vals)
        cy = yy + row_h * 0.36
        rect(s, px(lo), cy - 0.05, max(0.04, px(hi) - px(lo)), 0.1, fill=tint(col, 0.35))
        for vv in vals:
            rect(s, px(vv) - 0.045, cy - 0.045, 0.09, 0.09, fill=col)
        rect(s, px(r['median']) - 0.012, cy - 0.12, 0.024, 0.24, fill=INK)
        if hi > axis_max:
            text(s, px(axis_max) - 0.55, cy - 0.3, 0.7, 0.16, '%s →' % unit_fmt(hi), size=6.6, color=col, bold=True, align='r')
        # ours
        ours = collections.OrderedDict()
        for o in r['ours']:
            ours.setdefault((o['sup'], o['grams']), []).append(o)
        used = []
        for (sid, g), os_ in ours.items():
            v = (os_[0]['shelf'] if per == 'pack' else os_[0]['per_g'])
            slot = 0
            while any(abs(px(v) - u[0]) < 0.55 and u[1] == slot for u in used): slot += 1
            used.append((px(v), slot))
            oy = yy + row_h * 0.62 + (slot % 2) * 0.0
            rect(s, px(v) - 0.085, yy + row_h * 0.52, 0.17, 0.17, fill=SUP_COLOR[sid], line=WHITE, lw=1, radius=0.5)
            text(s, px(v) - 0.5, yy + row_h * 0.52 + 0.19 + slot * 0.17, 1.0, 0.16, '%s · %s' % (gfmt(g).replace(' г', ''), unit_fmt(v)) if False else unit_fmt(v), size=6.8, color=SUP_COLOR[sid], bold=True, align='c')
    return y + row_h * len(rows)


def seg_rows(per):
    out = []
    for sg in NORI_SEGS:
        rs = rows_of(sg); ls = lines(rs)
        out.append(dict(seg=sg, rows=rs, lines=ls, ours=[o for o in OUR if o['seg'] == sg],
                        median=med((l['pmed'] if per == 'pack' else l['per_g']) for l in ls)))
    return out


def legend_ours(s, x, y):
    for k, sid in enumerate(['tmk', 'zek', 'singha', 'thainichi']):
        rect(s, x + k * 1.9, y + 0.03, 0.14, 0.14, fill=SUP_COLOR[sid], radius=0.5)
        text(s, x + k * 1.9 + 0.2, y, 1.6, 0.2, SUP_NAME[sid], size=7.6, color=INK, bold=True)
    rect(s, x + 7.8, y + 0.07, 0.4, 0.06, fill=tint(NAVY, 0.35)); text(s, x + 8.28, y, 1.4, 0.2, 'діапазон ринку', size=7.4, color=MUTED)
    rect(s, x + 9.8, y + 0.05, 0.09, 0.09, fill=NAVY); text(s, x + 9.95, y, 1.2, 0.2, 'лінійка', size=7.4, color=MUTED)
    rect(s, x + 10.9, y + 0.01, 0.024, 0.18, fill=INK); text(s, x + 10.98, y, 1.2, 0.2, 'медіана', size=7.4, color=MUTED)


def pack_price_map(deck):
    s = deck.slide('Ціна за одну упаковку · ринок і наші SKU', 'Ціна упаковки: наша полиця проти ринку')
    legend_ours(s, 0.62, 1.4)
    rows = seg_rows('pack')
    end = range_chart(s, 0.62, 1.8, 12.1, rows, 260, 40, lambda v: '%d' % round(v), per='pack', row_h=1.12)
    text(s, 0.62, end + 0.32, 12.1, 0.4, 'Наші ціни — полиця за формулою рису (партнер ÷ 0,40 × 1,40) для найдешевшого контейнера. Темпура: Tao Kae Noi 59 г стоїть 528 ₴ в Onde — виходить за шкалу. Рисові чипси не показано: це не наш формат.', size=7.8, color=MUTED)
    src_note(s)
    return s


def per_g_map(deck):
    s = deck.slide('Ціна за грам · ринок і наші SKU', 'Ціна за грам: де ми дорожчі за ринок')
    legend_ours(s, 0.62, 1.4)
    rows = seg_rows('g')
    end = range_chart(s, 0.62, 1.8, 12.1, rows, 24, 4, lambda v: ('%.0f' % v if v == int(v) else ('%.1f' % v).replace('.', ',')), per='g', row_h=1.12)
    text(s, 0.62, end + 0.32, 12.1, 0.4, 'Ціна за грам — нормалізація до одного знаменника. Найдорожчі в категорії міні-пакети 4–5 г (до 22 ₴/г у органіки Clearspring); у більших форматах ціна за грам падає у 3–4 рази.', size=7.8, color=MUTED)
    src_note(s)
    return s


def audit_networks(deck):
    s = deck.slide('Ритейл-аудит · мережі', 'Що з норі-снеків стоїть у мережах')
    rs = [r for r in ROWS if r['seg'] in NORI_SEGS]
    text(s, 0.62, 1.36, 12.1, 0.22, 'Суцільна перевірка онлайн-каталогів 18 мереж. SKU — унікальні позиції норі-снеків і рисових крекерів. Праворуч: мін.–макс. (лінія) і медіана (крапка) ціни упаковки.', size=8.8, color=MUTED)
    x0, y0 = 0.62, 1.8
    heads = [('МЕРЕЖА', 0, 1.6), ('ТИП', 1.7, 0.8), ('SKU', 2.6, 3.6), ('ЦІНА УПАКОВКИ · мін.–макс. · медіана', 6.5, 5.0), ('МЕДІАНА', 11.5, 1.1)]
    for t, dx, w in heads:
        caps(s, x0 + dx, y0, w, t, size=6.4, align='l' if dx < 11 else 'r')
    chains = sorted([c for c in CHAIN if any(c in r['chains'] for r in rs)], key=lambda c: (-(c in NATIONAL), -sum(1 for r in rs if c in r['chains'])))
    rh = 0.268
    mx = max(sum(1 for r in rs if c in r['chains']) for c in chains)
    axis = 200.0
    tx, tw = x0 + 6.5, 4.9
    for v in (0, 50, 100, 150, 200):
        rect(s, tx + tw * v / axis, y0 + 0.28, 0.01, rh * (len(chains) + 0.2), fill=RULE)
        text(s, tx + tw * v / axis - 0.3, y0 + 0.28 + rh * (len(chains) + 0.2), 0.6, 0.16, str(v), size=6.6, color=MUTED, align='c')
    for i, c in enumerate(chains):
        y = y0 + 0.3 + i * rh
        here = [r for r in rs if c in r['chains']]
        if i % 2 == 0:
            rect(s, x0, y, 12.1, rh, fill=rgb('F8F9FC'))
        text(s, x0 + 0.05, y, 1.6, rh, CHAIN[c], size=8.2, color=NAVY, bold=True, anchor='m')
        nat = c in NATIONAL
        rect(s, x0 + 1.7, y + 0.05, 0.78, rh - 0.1, fill=NAVY if nat else TEAL)
        text(s, x0 + 1.7, y + 0.05, 0.78, rh - 0.1, 'НАЦ.' if nat else 'РЕГІОН.', size=6.2, color=WHITE, bold=True, align='c', anchor='m')
        # stacked by segment
        cx = x0 + 2.6; full = 3.0
        for sg in NORI_SEGS:
            n = sum(1 for r in here if r['seg'] == sg)
            if n:
                w = full * n / mx
                rect(s, cx, y + 0.07, w, rh - 0.14, fill=SEG_COLOR[sg])
                if w > 0.2: text(s, cx, y + 0.05, w, rh - 0.1, str(n), size=6.8, color=WHITE, bold=True, align='c', anchor='m')
                cx += w
        text(s, cx + 0.08, y, 0.6, rh, str(len(here)), size=8.2, color=MUTED, bold=True, anchor='m')
        pr = sorted(r['pmed'] for r in here)
        lo, hi, m = pr[0], pr[-1], med(pr)
        rect(s, tx + tw * min(lo, axis) / axis, y + rh / 2 - 0.02, max(0.03, tw * (min(hi, axis) - min(lo, axis)) / axis), 0.04, fill=GREY)
        rect(s, tx + tw * min(m, axis) / axis - 0.06, y + rh / 2 - 0.06, 0.12, 0.12, fill=AMBER, radius=0.5)
        text(s, x0 + 11.5, y, 1.1, rh, '%d ₴' % round(m), size=8.2, color=NAVY, bold=True, anchor='m', align='r')
    ly = y0 + 0.3 + rh * len(chains) + 0.32
    for k, sg in enumerate(NORI_SEGS):
        rect(s, x0 + 0.1 + k * 2.4, ly + 0.03, 0.14, 0.14, fill=SEG_COLOR[sg])
        text(s, x0 + 0.32 + k * 2.4, ly, 2.2, 0.2, SEGS[sg]['name'], size=7.4, color=INK)
    src_note(s)
    return s


def channels(deck):
    s = deck.slide('Канали продажу', 'Де продається категорія і кому')
    rs = [r for r in ROWS if r['seg'] in NORI_SEGS]
    ncn = len({c for r in rs for c in r['chains'] if c in NATIONAL}); nre = len({c for r in rs for c in r['chains'] if c not in NATIONAL})
    tk = sorted({CHAIN[c] for r in ROWS if r['brand'] == 'Tao Kae Noi' for c in r['chains']})
    ww = sorted({CHAIN[c] for r in ROWS if r['brand'] == 'Want Want' for c in r['chains']})
    mini = [r['pmed'] for r in rows_of('mini')]
    text(s, 0.62, 1.38, 12.1, 0.22, 'Мережі закривають мініпакети 4–5 г і чипси 25 г; темпура (Tao Kae Noi) є лише в %s, а рисові крекери — майже тільки Want Want.' % ', '.join(tk), size=8.8, color=MUTED)
    blocks = [
        ('Супермаркети · %d національних і %d регіональних мереж' % (ncn, nre), 'ЦІЛЬОВИЙ КАНАЛ', RED,
         ['Мініпакет 4–5 г: від %d ₴ до %d ₴, медіана %d ₴ (Ock Dong Ja, Akura, Norris, Haelove — найширша присутність)' % (round(min(mini)), round(max(mini)), round(med(mini))),
          'Чипси 25 г — Hokkaido Club і Norris: 47–88 ₴ залежно від мережі',
          'Рисові крекери: Want Want mini 60 г — 140–149 ₴ (%s)' % ', '.join(ww),
          'Окремі розділи «Норі» / «Азія» є в каталогах METRO, NOVUS, Ашан, Зараз, Onde, Епіцентр']),
        ('Азійські інтернет-магазини · Asia Foods, Японський Квартал, До Смаку, Суші Повар', 'ЦІЛЬОВИЙ КАНАЛ', AMBER,
         ['Tao Kae Noi Hot & Spicy 15 г — 107 ₴ (Asia Foods); Wasabi 15 г — 78 ₴ (Суші Повар)',
          'Tao Kae Noi Roll 6 × 3 г — 89 ₴ (Суші Повар)',
          'Kimnori 5 г — 55 ₴ (Японський Квартал); Akura 4,5 г — 35 ₴ (До Смаку)',
          'Рисові крекери з водоростями Want Want 160 г — 357 ₴ (Asia Foods)']),
        ('Маркетплейси · Rozetka, Prom.ua', 'ВХІД', TEAL,
         ['Akura Nori Chips 3 × 4,5 г — 129 ₴ (Original) і 89 ₴ (Kimchi) на Rozetka',
          'Tao Kae Noi Big Roll Classic 6 × 3 г — 154,85 ₴ на Prom.ua',
          'Edward & Sons рисові крекери з нори 100 г — 351 ₴ (Prom.ua, органіка)',
          'Опт: Akura 72 × 4,5 г — 1 820 ₴ (Exotic Food, Prom.ua)']),
        ('HoReCa · суші-бари, азійські ресторани', 'НЕ ОХОПЛЕНО', GREY,
         ['У каталогах мереж є листи норі для суші (Katana, JS, Hokkaido Club, Akura) — це інша категорія, у зріз не включено',
          'Окремий аудит HoReCa не проводили']),
    ]
    heights = [1.38, 1.38, 1.38, 0.88]
    y = 1.72
    for (ttl, tag, col, lines_), h in zip(blocks, heights):
        rect(s, 0.62, y, 12.1, h, fill=PANEL); rect(s, 0.62, y, 0.05, h, fill=col)
        text(s, 0.9, y + 0.1, 9.0, 0.24, ttl, size=9.4, color=NAVY, bold=True)
        rect(s, 10.6, y + 0.1, 1.9, 0.24, fill=col)
        text(s, 10.6, y + 0.1, 1.9, 0.24, tag, size=6.4, color=WHITE, bold=True, align='c', anchor='m')
        for k, t in enumerate(lines_):
            text(s, 0.9, y + 0.42 + k * 0.235, 11.5, 0.22, '•  ' + t, size=8.0, color=INK)
        y += h + 0.09
    footnote(s, 'Мережі — каталоги zakaz.ua, 07.10.2026. Онлайн-магазини й маркетплейси — оголошення, знайдені пошуком; повного аудиту цих каналів не проводили, ціни можуть змінитись.')
    return s


# ================================================================ our pricing
def clean_title(o):
    t = o['title'].replace('Seaweed Topping', 'Topping').replace('Tempura Seaweed', 'Tempura').replace('Sandwich Seaweed', 'Sandwich').replace('Arare Norimaki', 'Norimaki')
    import re
    return re.sub(r'\s*·\s*\d+\s*г$', '', t)


def pricing_formula(deck):
    s = deck.slide('Ціноутворення', 'З чого складається ціна на полиці')
    rect(s, 0.62, 1.45, 12.1, 1.38, fill=PANEL)
    caps(s, 0.85, 1.56, 8, 'Ціна партнеру — за неї ми продаємо мережі (100 %)', size=7)
    parts = [('Собівартість 40 %', 40, NAVY), ('Бонус мережі 25 %', 25, AMBER), ('Наша маржа 35 %', 35, GOLD), ('Націнка магазину +40 %', 40, GREY)]
    tot = 140; w = 11.3; x = 0.85
    for t, v, c in parts:
        ww = w * v / tot
        rect(s, x, 1.84, ww - 0.03, 0.38, fill=c)
        text(s, x, 1.84, ww, 0.38, t, size=8.6, color=WHITE if c != GOLD else INK, bold=True, align='c', anchor='m')
        x += ww
    rect(s, 0.85, 2.28, w * 100 / tot, 0.015, fill=INK)
    text(s, 0.85, 2.3, w * 100 / tot, 0.2, '= ціна партнеру', size=7.8, color=INK, bold=True, align='c')
    text(s, 0.85, 2.55, 11.3, 0.22, 'ПОЛИЦЯ = ціна партнеру × 1,40 · ціна партнеру = собівартість ÷ 0,40 · бонус 25 % і маржа 35 % — частки ціни партнеру', size=8.2, color=NAVY, bold=True, align='c')
    groups = []
    seen = set()
    for o in OUR:
        key = (o['sup'], o['grams'], round(o['shelf']))
        if key in seen: continue
        seen.add(key); groups.append(o)
    hy = 2.98
    caps(s, 0.62, hy, 4, 'Товар', size=6.4)
    caps(s, 2.8, hy, 5, 'Грн за упаковку · СС · бонус · маржа · націнка', size=6.4)
    for edge, t in [(9.2, 'ЦІНА ПАРТНЕРУ'), (10.4, 'FOB, $'), (11.5, 'СС, $'), (12.7, 'СС, ГРН')]:
        caps(s, edge - 1.2, hy, 1.2, t, size=6.2, align='r')
    ry0 = 3.25
    rh = min(0.30, (6.78 - ry0) / len(groups))
    mx = max(o['shelf'] for o in groups)
    bw = 4.9
    for i, o in enumerate(groups):
        y = ry0 + i * rh
        rect(s, 0.62, y, 12.1, rh - 0.03, fill=PANEL)
        rect(s, 0.62, y, 0.05, rh - 0.03, fill=SUP_COLOR[o['sup']])
        text(s, 0.78, y, 2.0, rh - 0.03, '%s %s' % (clean_title(o), gfmt(o['grams'])), size=7.6, color=NAVY, bold=True, anchor='m')
        x = 2.8
        sc = bw / mx
        for v, c, fg in [(o['cost'], NAVY, WHITE), (o['bonus'], AMBER, WHITE), (o['profit'], GOLD, INK), (o['shelf'] - o['partner'], GREY, WHITE)]:
            ww = v * sc
            rect(s, x, y + 0.04, ww - 0.02, rh - 0.11, fill=c)
            if ww > 0.34: text(s, x, y + 0.02, ww, rh - 0.07, '%d' % round(v), size=6.6, color=fg, bold=True, align='c', anchor='m')
            x += ww
        text(s, x + 0.08, y, 0.8, rh - 0.03, '%d' % round(o['shelf']), size=9, color=INK, bold=True, anchor='m')
        text(s, 8.2, y, 1.0, rh - 0.03, '%d' % round(o['partner']), size=8.8, color=INK, bold=True, anchor='m', align='r')
        text(s, 9.4, y, 1.0, rh - 0.03, '$%s' % fnum(o['fob'], 3), size=7.8, color=MUTED, anchor='m', align='r')
        text(s, 10.5, y, 1.0, rh - 0.03, '$%s' % fnum(o['usd'], 3), size=7.8, color=MUTED, anchor='m', align='r')
        text(s, 11.7, y, 1.0, rh - 0.03, fnum(o['cost'], 1), size=7.8, color=MUTED, anchor='m', align='r')
    ly = ry0 + rh * len(groups) + 0.05
    for k, (t, c) in enumerate([('Собівартість (СС)', NAVY), ('Бонус мережі', AMBER), ('Наша маржа', GOLD), ('Націнка магазину', GREY)]):
        rect(s, 0.7 + k * 2.2, ly + 0.03, 0.13, 0.13, fill=c)
        text(s, 0.92 + k * 2.2, ly, 2.0, 0.2, t, size=7.4, color=INK)
    footnote(s, 'Фінмодель: курс 45 ₴/$, найдешевший контейнер для кожного SKU (переважно 40′), СС — Self-Cost_Snacks, ціна вже з ПДВ. Повна модель по всіх сценаріях і маржах — файл SNACKS_FINMODEL_ALL_MARGINS.xlsx.')
    return s


def shelf_table(deck, sids, market_cards=None):
    items = [o for o in OUR if o['sup'] in sids]
    col = SUP_COLOR[sids[0]]
    names = ' · '.join(SUP_NAME[x] for x in sids) if len(sids) == 1 else ' і '.join(SUP_NAME[x] for x in sids)
    s = deck.slide('Полиця · %s' % names, 'Наша ціна на полиці: %s' % names)
    # right edges (relative to x=0.62) and widths of the numeric columns
    cols = [('ВАГА', 3.6, 0.65), ('FOB, $', 4.4, 0.7), ('СС, $', 5.15, 0.7), ('СС, ГРН', 5.95, 0.7), ('КОНТЕЙНЕР', 6.85, 0.9),
            ('ПАРТНЕР', 7.7, 0.7), ('БОНУС 25 %', 8.7, 0.9), ('МАРЖА 35 %', 9.75, 0.95), ('ПОЛИЦЯ', 10.65, 0.8), ('₴/Г', 11.3, 0.5), ('VS РИНОК', 12.1, 0.7)]
    n = len(items)
    top = 1.5
    rh = min(0.42, 5.0 / (n + 0.5)) if market_cards is None else 0.42
    rect(s, 0.62, top - 0.05, 12.1, 0.3, fill=col)
    text(s, 0.62 + 0.75, top, 2.5, 0.22, 'ПОЗИЦІЯ', size=6.2, color=WHITE, bold=True, anchor='m')
    for t, edge, w in cols:
        text(s, 0.62 + edge - w, top, w, 0.22, t, size=6.2, color=WHITE, bold=True, align='r', anchor='m')
    for i, o in enumerate(items):
        y = top + 0.35 + i * rh
        if i % 2 == 0:
            rect(s, 0.62, y, 12.1, rh, fill=rgb('F8F9FC'))
        rect(s, 0.62, y, 0.05, rh, fill=SUP_COLOR[o['sup']])
        if os.path.exists(o['photo']):
            picture(s, o['photo'], 0.72, y + 0.02, 0.5, rh - 0.04)
        text(s, 0.62 + 0.75, y, 2.6, rh, clean_title(o), size=8, color=NAVY, bold=True, anchor='m')
        ps = D.peer_stats(o)
        vals = [gfmt(o['grams']), '$' + fnum(o['fob'], 3), '$' + fnum(o['usd'], 3), fnum(o['cost'], 1), o['scenario'].replace(' контейнер', ''), fnum(o['partner'], 0), fnum(o['bonus'], 0),
                fnum(o['profit'], 0), '%d ₴' % round(o['shelf']), fnum(o['per_g'], 1), pct(o['per_g'], ps['per_g'])]
        for (t, edge, w), v in zip(cols, vals):
            bold = t in ('ПОЛИЦЯ', 'VS РИНОК')
            c = INK if t == 'ПОЛИЦЯ' else (MUTED if t in ('FOB, $', 'СС, $', 'СС, ГРН', 'КОНТЕЙНЕР') else NAVY)
            if t == 'VS РИНОК':
                d = round((o['per_g'] / ps['per_g'] - 1) * 100)
                c = RED if d > 25 else (GREEN if d <= 0 else AMBER)
            text(s, 0.62 + edge - w, y, w, rh, v, size=8.6 if t == 'ПОЛИЦЯ' else 7.8, color=c, bold=bold, anchor='m', align='r')
    if market_cards:
        y0 = top + 0.35 + n * rh + 0.2
        section_tag(s, 0.62, y0, 8, 'Що стоїть поруч на полиці мереж · рисові крекери', color=RED)
        for k, l in enumerate(market_cards):
            x = 0.62 + k * 2.4
            rect(s, x, y0 + 0.32, 2.3, 1.45, fill=PANEL); rect(s, x, y0 + 0.32, 2.3, 0.04, fill=RED)
            rect(s, x + 0.08, y0 + 0.44, 0.95, 0.95, fill=WHITE)
            picture(s, os.path.join(HERE, l['img']), x + 0.1, y0 + 0.46, 0.91, 0.91)
            text(s, x + 1.12, y0 + 0.46, 1.1, 0.4, '%s · %s' % (l['brand'], gfmt(l['grams'])), size=8, color=INK, bold=True)
            text(s, x + 1.12, y0 + 0.88, 1.1, 0.26, '%s ₴' % rng(l['pmin'], l['pmax']), size=12, color=INK, bold=True)
            text(s, x + 1.12, y0 + 1.16, 1.1, 0.2, '%s ₴/г' % fnum(l['per_g'], 1), size=7.4, color=MUTED)
            text(s, x + 0.1, y0 + 1.46, 2.1, 0.2, ' · '.join(CHAIN[c] for c in l['chains'][:3]), size=6.8, color=MUTED)
    note = {'singha': '', 'thainichi': '',
            'tmk': 'TMK продається у двох сегментах ринку: 2,5–5 г — міні-пакет; 10–12 г — чипси. ',
            'zek': 'ZEK: темпура 30–50 г і топінг 35–70 г — сегмент «Темпура й топінг»; Sandwich 25 г — сегмент «Чипси». '}[sids[0]]
    footnote(s, note + '«vs ринок» — наш ₴/г проти медіани ₴/г близьких за вагою (×0,5–×2) позицій того ж сегмента. Фінмодель: курс 45 ₴/$, бонус 25 %, маржа 35 %, полиця = партнер × 1,40.')
    return s


# ======================================================== who stands with whom
def stand(deck, seg_list, eyebrow, title, lo, hi, step, ours_filter=None, sub=None, takeaways=None):
    s = deck.slide(eyebrow, title)
    mk = []
    for sg in seg_list:
        for l in lines(rows_of(sg)):
            mk.append(dict(name='%s · %s' % (l['brand'], gfmt(l['grams'])), lo=l['pmin'], hi=l['pmax'], med=l['pmed'], n=l['n'], seg=sg, ours=False, chains=l['chains']))
    og = collections.OrderedDict()
    for o in OUR:
        if o['seg'] in seg_list and (ours_filter is None or ours_filter(o)):
            og.setdefault((o['sup'], o['grams'], round(o['shelf'])), []).append(o)
    for (sid, g, sh), os_ in og.items():
        o = os_[0]
        base = clean_title(o)
        for flv in (' Original', ' Corn', ' Chicken Floss', ' Vegetables', ' Sesame', ' Meat Floss', ' Hot Spicy', ' Spicy Squid', ' Squid', ' Spicy', ' Wasabi', ' Mala Crawfish', ' Mala', ' Chicken'):
            base = base.replace(flv, '')
        mk.append(dict(name='НАШ · %s · %s' % ({'tmk': 'TMK', 'zek': 'ZEK', 'singha': 'Singha', 'thainichi': 'Thai-Nichi'}[sid], base.replace('Wow ', '').replace('Norimaki', 'Norimaki')),
                       lo=sh, hi=sh, med=sh, n=len(os_), seg=seg_list[0], ours=True, sid=sid, grams=g))
    mk.sort(key=lambda m: m['med'])
    ours_prices = [m['med'] for m in mk if m['ours']]
    clo, chi = min(ours_prices) * 0.8, max(ours_prices) * 1.2
    n = len(mk)
    compact = bool(takeaways) and n > 14
    if compact:
        sub = (sub or '') + '  ' + '  '.join('%s: %s' % (t[0].capitalize(), t[1]) for t in takeaways)
        takeaways = None
    text(s, 0.62, 1.34, 12.1, 0.55, sub or '', size=8.2 if compact else 9.0, color=MUTED)
    top = 1.98 if compact else 1.74
    rh = min(0.285, (6.62 - top - 0.34) / n) if not takeaways else min(0.285, (6.95 - top - 0.34 - 1.65) / n)
    tx, tw = 6.0, 6.2
    px = lambda v: tx + tw * (min(max(v, lo), hi) - lo) / (hi - lo)
    rect(s, 0.62, top - 0.04, 12.1, 0.3, fill=SEG_COLOR[seg_list[0]])
    for t, dx, w, al in [('ЛІНІЙКА · ВІД ДЕШЕВИХ ДО ДОРОГИХ', 0.1, 4.0, 'l'), ('ВАГА', 3.5, 0.8, 'l'), ('SKU', 4.6, 0.5, 'l')]:
        text(s, 0.62 + dx, top, w, 0.22, t, size=6.4, color=WHITE, bold=True, align=al, anchor='m')
    rect(s, px(clo), top + 0.26, px(chi) - px(clo), rh * n + 0.05, fill=rgb('FFF3DC'))
    v = lo
    while v <= hi + 1e-9:
        rect(s, px(v), top + 0.26, 0.008, rh * n + 0.05, fill=RULE)
        text(s, px(v) - 0.3, top + 0.3 + rh * n, 0.6, 0.16, '%d' % v, size=6.6, color=MUTED, align='c')
        v += step
    for i, m in enumerate(mk):
        y = top + 0.3 + i * rh
        if m['ours']:
            rect(s, 0.62, y, 12.1, rh, fill=rgb('FFEBC2'))
        c = SUP_COLOR[m['sid']] if m['ours'] else SEG_COLOR[m['seg']]
        text(s, 0.7, y, 3.9, rh, m['name'], size=7.8, color=INK, bold=True, anchor='m')
        text(s, 4.2, y, 0.8, rh, gfmt(m['grams']) if m['ours'] else m['name'].split(' · ')[-1], size=7.2, color=MUTED, anchor='m')
        text(s, 5.2, y, 0.5, rh, str(m['n']), size=7.4, color=MUTED, anchor='m')
        cy = y + rh / 2
        if m['hi'] > m['lo']:
            rect(s, px(m['lo']), cy - 0.025, px(m['hi']) - px(m['lo']), 0.05, fill=tint(c, 0.4))
            rect(s, px(m['lo']) - 0.04, cy - 0.04, 0.08, 0.08, fill=WHITE, line=c, lw=0.75, radius=0.5)
            rect(s, px(m['hi']) - 0.04, cy - 0.04, 0.08, 0.08, fill=WHITE, line=c, lw=0.75, radius=0.5)
        d = 0.19 if m['ours'] else 0.14
        rect(s, px(m['med']) - d / 2, cy - d / 2, d, d, fill=c, line=(INK if m['ours'] else None), lw=1.0, radius=0.5)
        lab = ('%d' % round(m['med'])) if m['hi'] <= m['lo'] else '%d–%d' % (round(m['lo']), round(m['hi']))
        out = (' →' if m['hi'] > hi else '')
        text(s, px(m['hi']) + 0.14 if px(m['hi']) < tx + tw - 1.0 else px(m['lo']) - 1.15, y, 1.0, rh, lab + out, size=7.6, color=INK, bold=True, anchor='m', align='l' if px(m['hi']) < tx + tw - 1.0 else 'r')
    ey = top + 0.34 + rh * n + 0.22
    text(s, 0.62, ey, 12.1, 0.2, 'Грн за упаковку · велика точка — медіана, малі — межі діапазону між мережами; золота зона — наша полиця ±20 %', size=7, color=MUTED)
    if takeaways:
        ty = ey + 0.35
        w = 12.1 / len(takeaways)
        for k, (ttl, body, col) in enumerate(takeaways):
            x = 0.62 + k * w
            rect(s, x, ty, w - 0.12, 1.0, fill=PANEL); rect(s, x, ty, 0.04, 1.0, fill=col)
            text(s, x + 0.2, ty + 0.1, w - 0.5, 0.16, ttl, size=7, color=col, bold=True, spc=0.2)
            text(s, x + 0.2, ty + 0.32, w - 0.5, 0.8, body, size=8.0, color=INK)
    src_note(s)
    return s


def neighbours(deck):
    s = deck.slide('Хто з ким стоїть · сусіди', 'Найближчі за ціною до кожного нашого товару')
    mk = []
    for sg in NORI_SEGS:
        for l in lines(rows_of(sg)):
            mk.append(dict(l, seg=sg))
    text(s, 0.62, 1.36, 12.1, 0.22, 'Для кожної нашої позиції — три найближчі за ціною упаковки лінійки ринку з того самого сегмента (±20 %, якщо такі є).', size=9.0, color=MUTED)
    groups = []
    seen = set()
    for o in OUR:
        key = (o['sup'], o['grams'], round(o['shelf']))
        if key not in seen:
            seen.add(key); groups.append(o)
    top = 1.76
    rh = min(0.37, 4.85 / (len(groups) + 0.5))
    rect(s, 0.62, top, 12.1, 0.28, fill=NAVY)
    for t, dx, w in [('НАШ ТОВАР', 0.1, 2.5), ('ПОЛИЦЯ', 2.7, 0.8), ('СУСІД 1', 3.75, 2.8), ('СУСІД 2', 6.65, 2.8), ('СУСІД 3', 9.55, 2.8)]:
        text(s, 0.62 + dx, top, w, 0.28, t, size=6.4, color=WHITE, bold=True, anchor='m')
    for i, o in enumerate(groups):
        y = top + 0.32 + i * rh
        if i % 2 == 0: rect(s, 0.62, y, 12.1, rh, fill=rgb('F8F9FC'))
        rect(s, 0.62, y, 0.05, rh, fill=SUP_COLOR[o['sup']])
        text(s, 0.78, y, 2.6, rh, '%s %s' % (clean_title(o), gfmt(o['grams'])), size=7.4, color=NAVY, bold=True, anchor='m')
        text(s, 3.32, y, 0.8, rh, '%d ₴' % round(o['shelf']), size=9, color=INK, bold=True, anchor='m')
        pool = [l for l in mk if l['seg'] == o['seg'] and 0.4 * o['grams'] <= l['grams'] <= 2.5 * o['grams']]
        pool = pool if len(pool) >= 3 else [l for l in mk if l['seg'] == o['seg']]
        cand = sorted(pool, key=lambda l: abs(l['pmed'] - o['shelf']))[:3]
        for k, l in enumerate(cand):
            dpc = (l['pmed'] / o['shelf'] - 1) * 100
            near = abs(dpc) <= 20
            text(s, 4.37 + k * 2.9, y, 2.85, rh, [[('%s %s  ' % (l['brand'], gfmt(l['grams'])), {'bold': True, 'color': NAVY}), ('%d ₴' % round(l['pmed']), {'bold': True, 'color': INK}),
                                                  ('  %+d %%' % round(dpc), {'color': GREEN if near else RED, 'bold': True})]], size=7.2, anchor='m')
    footnote(s, 'Зелений — сусід у коридорі ±20 % від нашої полиці; червоний — далі. Порівняння за ціною упаковки (не за грам): різна вага упаковки — окремий ризик, див. слайд «Ціна за грам».')
    return s


def sensitivity(deck):
    s = deck.slide('Чутливість', 'Як змінюється полиця при іншій марже')
    ms = ['0.35', '0.3', '0.25', '0.2', '0.15']
    sens = {r['title']: r for r in PM['sensitivity']}
    groups = []
    seen = set()
    for o in OUR:
        key = (o['sup'], o['grams'], round(o['shelf']))
        if key not in seen:
            seen.add(key); groups.append(o)
    text(s, 0.62, 1.38, 12.1, 0.22, 'Полиця при цільовій маржі 35 → 15 % (бонус 25 %, ×1,40). Зелена — не вище за медіану упаковки близьких за вагою позицій ринку; червона — вище за весь їхній діапазон.', size=8.6, color=MUTED)
    top = 1.78; rh = min(0.34, 4.9 / (len(groups) + 0.5))
    rect(s, 0.62, top, 12.1, 0.3, fill=NAVY)
    text(s, 0.78, top, 3.5, 0.3, 'ТОВАР', size=6.4, color=WHITE, bold=True, anchor='m')
    text(s, 3.7, top, 2.0, 0.3, 'РИНОК ПОРУЧ · МЕДІАНА', size=6.2, color=WHITE, bold=True, anchor='m', align='r')
    X0, STEP, W = 6.1, 1.3, 1.2
    for k, m in enumerate(ms):
        text(s, X0 + k * STEP, top, W, 0.3, 'МАРЖА %d %%' % round(float(m) * 100), size=6.4, color=WHITE, bold=True, anchor='m', align='r')
    for i, o in enumerate(groups):
        y = top + 0.34 + i * rh
        if i % 2 == 0: rect(s, 0.62, y, 12.1, rh, fill=rgb('F8F9FC'))
        rect(s, 0.62, y, 0.05, rh, fill=SUP_COLOR[o['sup']])
        text(s, 0.78, y, 3.0, rh, '%s %s' % (clean_title(o), gfmt(o['grams'])), size=7.6, color=NAVY, bold=True, anchor='m')
        ps = D.peer_stats(o)
        text(s, 3.7, y, 2.0, rh, '%d ₴' % round(ps['pack']), size=7.6, color=MUTED, anchor='m', align='r')
        sv = sens[o['title']]['shelf_by_margin']
        for k, m in enumerate(ms):
            v = sv[m]
            bg = tint(GREEN, 0.22) if v <= ps['pack'] else (tint(RED, 0.25) if v > ps['hi'] else None)
            if bg is not None:
                rect(s, X0 + k * STEP + 0.02, y + 0.03, W, rh - 0.06, fill=bg)
            text(s, X0 + k * STEP, y, W - 0.1, rh, '%d ₴' % round(v), size=8.2, color=INK, bold=(m == '0.35'), anchor='m', align='r')
    footnote(s, 'Близькі за вагою позиції — той самий сегмент, вага ×0,5–×2 від нашої. Повна фінмодель з усіма сценаріями — SNACKS_FINMODEL_ALL_MARGINS.xlsx.')
    return s
