# -*- coding: utf-8 -*-
"""Cross-segment market maps and our own price slides (formula from the rice study)."""
import collections, os
from md import *  # noqa
from md import K, rect, text, caps, picture, kpi, section_tag, hbars, footnote, rgb
import mdata as D
from mdata import SEGS, rows_of, lines, med, gfmt, uah, rng, CHAIN, NATIONAL, ROWS, OUR, PM
from slides_market import CHAINS_ORDER, src_note, HERE

BON, MAR, MUP = PM['params']['bonus'], PM['params']['margin'], PM['params']['markup']
MP = round(MAR * 100)                   # our margin, % of partner price
SSP = round((1 - BON - MAR) * 100)      # self-cost share of partner price, %
NORI_SEGS = ('mini', 'chips', 'tempura', 'ricecr')
SHORT = {'megamarket': 'Мега-\nМаркет', 'ultramarket': 'Ultra-\nmarket', 'chudomarket': 'Чудо-\nМаркет', 'epicentr': 'Епі-\nцентр', 'vostorg': 'Вос-\nторг', 'tavriav': 'Таврія\nВ'}
SUP_NAME = {'singha': 'Singha Kameda', 'thainichi': 'Thai-Nichi', 'tmk': 'TMK · KOKIRI', 'zek': 'HanJin · ZEK'}
SEG_OF = {o['key']: o['seg'] for o in OUR}


def sub_line(s, t_, y=1.34, size=8.4):
    text(s, 0.62, y, 12.1, 0.45, t_, size=size, color=MUTED)


def pct(a, b):
    return '%+d %%' % round((a / b - 1) * 100)


def fnum(v, d=2):
    return ('%.' + str(d) + 'f') % v if False else (('%.' + str(d) + 'f') % v).replace('.', ',')


# ================================================================= market maps
def brand_network(deck):
    s = deck.slide('Зведення · бренди × мережі', 'Який бренд у якій мережі і за скільки')
    rs = [r for r in ROWS if r['seg'] in NORI_SEGS]
    brands = collections.Counter(r['brand'] for r in rs).most_common(12)
    chains = [c for c in CHAINS_ORDER]
    x0, y0, nw, cw, rh = 0.62, 1.95, 2.05, 0.555, 0.37
    text(s, 0.62, 1.36, 12.1, 0.5, 'У клітинці — типова ціна упаковки цього бренду в цій мережі, ₴. Тире — бренду там немає. Чим темніше, тим більше позицій бренду. Під назвою бренду — що саме він продає (формат і вага); хто бренд і хто його завозить — на наступному слайді.', size=8.6, color=MUTED)
    for j, c in enumerate(chains):
        x = x0 + nw + j * cw
        rect(s, x + 0.02, y0, cw - 0.04, 0.42, fill=NAVY if c in NATIONAL else TEAL)
        text(s, x, y0 + 0.02, cw, 0.38, SHORT.get(c, CHAIN[c]), size=6.2, color=WHITE, bold=True, align='c', anchor='m')
    xs = x0 + nw + len(chains) * cw + 0.06
    for k, h in enumerate(['SKU', 'МЕРЕЖ']):
        rect(s, xs + k * 0.56, y0, 0.54, 0.42, fill=GOLD)
        text(s, xs + k * 0.56, y0 + 0.02, 0.54, 0.38, h, size=6.0, color=WHITE, bold=True, align='c', anchor='m')
    mx = 6
    for i, (b, n) in enumerate(brands):
        y = y0 + 0.48 + i * rh
        pf = D.brand_profile(b)
        text(s, x0, y + 0.02, nw - 0.1, rh * 0.5, b, size=8.4, color=NAVY, bold=True, align='r')
        text(s, x0, y + rh * 0.5, nw - 0.1, rh * 0.5, pf['fmt'][:34], size=6.0, color=MUTED, align='r')
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
            text(s, x, y, cw, rh, '%d' % round(med(r['pmed'] for r in here)), size=9, color=WHITE if t > 0.45 else NAVY, bold=True, align='c', anchor='m')
        text(s, xs, y, 0.54, rh, str(n), size=8.2, color=NAVY, bold=True, align='c', anchor='m')
        text(s, xs + 0.56, y, 0.54, rh, str(len({c for r in brs for c in r['chains']})), size=8.2, color=NAVY, bold=True, align='c', anchor='m')
    src_note(s)
    return s


def brand_directory(deck):
    s = deck.slide('Бренди · що це і хто стоїть за ними', 'Бренди на полиці: що продають і хто завозить')
    text(s, 0.62, 1.36, 12.1, 0.45, 'Формат і ціна — з каталогів мереж. Імпортер і країна — з митної бази (за описами декларацій і назвою імпортера); якщо в базі бренд не знайдено, так і написано — власника бренду не вгадуємо.', size=8.6, color=MUTED)
    brands = [b for b, _ in collections.Counter(r['brand'] for r in ROWS if r['seg'] in NORI_SEGS or r['seg'] == 'ricecr').most_common(15)]
    top = 1.98
    heads = [('БРЕНД', 0.1, 1.6, 'l'), ('ЩО ПРОДАЄ · ФОРМАТ І ВАГА', 1.8, 3.9, 'l'), ('SKU', 5.75, 0.5, 'r'), ('МЕРЕЖ', 6.3, 0.6, 'r'), ('ЦІНА УПАКОВКИ, ₴', 7.0, 1.4, 'l'), ('ІМПОРТЕР В УКРАЇНІ (МИТНА БАЗА)', 8.5, 2.4, 'l'), ('ВИРОБНИК / КРАЇНА', 10.95, 1.2, 'l')]
    rect(s, 0.62, top, 12.1, 0.3, fill=NAVY)
    for t, dx, w, al in heads:
        text(s, 0.62 + dx, top, w, 0.3, t, size=6.2, color=WHITE, bold=True, anchor='m', align=al)
    rh = min(0.33, (7.0 - top - 0.4) / len(brands))
    for i, b in enumerate(brands):
        pf = D.brand_profile(b); imp, ctry = pf['imp']
        y = top + 0.34 + i * rh
        if i % 2 == 0: rect(s, 0.62, y, 12.1, rh, fill=rgb('F8F9FC'))
        text(s, 0.72, y, 1.7, rh, b, size=8.4, color=NAVY, bold=True, anchor='m')
        text(s, 2.42, y, 3.95, rh, pf['fmt'], size=7.4, color=INK, anchor='m')
        text(s, 6.37, y, 0.5, rh, str(pf['n']), size=8, color=INK, anchor='m', align='r')
        text(s, 6.92, y, 0.6, rh, str(pf['chains']), size=8, color=INK, anchor='m', align='r')
        text(s, 7.62, y, 1.4, rh, '%s ₴' % rng(pf['pmin'], pf['pmax']), size=7.6, color=INK, anchor='m')
        text(s, 9.12, y, 2.4, rh, imp or 'у митній базі не знайдено', size=7.6, color=(INK if imp else MUTED), bold=bool(imp), anchor='m')
        text(s, 11.57, y, 1.15, rh, ctry or '—', size=7, color=(INK if ctry else MUTED), anchor='m')
    src_note(s)
    return s


def weight_map(deck):
    s = deck.slide('Формат × вага', 'Де насправді щільна полиця: вага упаковки')
    rs = [r for r in ROWS if r['seg'] in NORI_SEGS]
    bins = [(0, 5.5, 'до 5 г'), (5.5, 16, '8–15 г'), (16, 31, '20–30 г'), (31, 61, '35–60 г'), (61, 999, '70 г і більше')]
    total = len(rs)
    sub_line(s, 'Скільки позицій ринку стоїть у кожній ваговій смузі, яка там типова ціна упаковки, і які саме наші SKU (із цінами нашої полиці при марже %d %%) потрапляють у цю смугу.' % MP, y=1.34, size=8.8)
    caps(s, 0.62, 1.85, 4, 'Позицій ринку у смузі', size=7)
    caps(s, 5.55, 1.85, 3, 'Типова ціна упаковки, ₴', size=7)
    caps(s, 8.0, 1.85, 5, 'Наші SKU у смузі · наша полиця, ₴', size=7, color=AMBER)
    maxn = max(sum(1 for r in rs if a <= r['grams'] < b) for a, b, _ in bins)
    groups = D.our_groups()
    stats = {}
    pitch = 0.86
    for i, (a, b, lab) in enumerate(bins):
        sel = [r for r in rs if a <= r['grams'] < b]
        mine = [o for o in groups if a <= o['grams'] < b]
        y = 2.12 + i * pitch
        rect(s, 0.55, y - 0.04, 12.25, pitch - 0.08, fill=rgb('FFF3DC') if mine else rgb('F8F9FC'))
        text(s, 0.62, y, 1.2, pitch - 0.16, lab, size=9.4, color=NAVY, bold=True, anchor='m')
        w = 2.1 * len(sel) / maxn
        rect(s, 1.85, y + 0.17, max(0.06, w), 0.32, fill=NAVY)
        text(s, 1.85 + w + 0.08, y, 1.5, pitch - 0.16, '%d · %d %%' % (len(sel), round(100 * len(sel) / total)), size=8.6, color=MUTED, bold=True, anchor='m')
        mp = med(r['pmed'] for r in sel)
        stats[lab] = (len(sel), mp)
        w2 = 1.4 * min(mp, 250) / 250
        rect(s, 5.55, y + 0.17, w2, 0.32, fill=AMBER)
        text(s, 5.55 + w2 + 0.08, y, 1.0, pitch - 0.16, '%d ₴' % round(mp), size=8.6, color=MUTED, bold=True, anchor='m')
        xx, yy = 8.0, y + 0.06
        merged = collections.OrderedDict()
        for o in mine:
            base = clean_title(o)
            for flv in (' Original', ' Corn', ' Chicken Floss', ' Vegetables', ' Sesame', ' Meat Floss', ' Hot Spicy', ' Spicy Squid', ' Squid', ' Spicy', ' Wasabi', ' Mala Crawfish', ' Mala', ' Chicken'):
                base = base.replace(flv, '')
            base = base.replace('Wow ', '').replace('Norimaki', 'Norimaki')
            merged.setdefault((o['sup'], base, o['grams']), []).append(o)
        for (sid, base, g), os_ in merged.items():
            pr = sorted(round(o['shelf']) for o in os_)
            ptxt = ('%d' % pr[0]) if pr[0] == pr[-1] else '%d–%d' % (pr[0], pr[-1])
            lab_t = '%s %s · %s ₴' % (base, gfmt(g), ptxt)
            w_ = 0.25 + 0.07 * len(lab_t)
            if xx + w_ > 12.75: xx, yy = 8.0, yy + 0.27
            rect(s, xx, yy, w_, 0.24, fill=WHITE, line=SUP_COLOR[sid], lw=1.0, radius=0.3)
            text(s, xx, yy, w_, 0.24, lab_t, size=6.4, color=INK, bold=True, align='c', anchor='m')
            xx += w_ + 0.06
        if not mine:
            text(s, 8.0, y, 4.6, pitch - 0.16, 'наших SKU у цій смузі немає', size=7.6, color=MUTED, italic=True, anchor='m')
    n5, n8, n20, n35 = stats['до 5 г'], stats['8–15 г'], stats['20–30 г'], stats['35–60 г']
    tm = [o for o in OUR if 8 <= o['grams'] <= 15]
    y = 6.5
    boxes = [('ДЕ СТОЇТЬ НАШ АСОРТИМЕНТ',
              'Наші SKU є у всіх вагових смугах. Найщільніша полиця — до 5 г (%d позицій, %d %%); у смузі 35–60 г — %d позицій ринку й %d наших SKU (Singha, Thai-Nichi, ZEK).' % (n5[0], round(100 * n5[0] / total), n35[0], len([o for o in OUR if 31 <= o['grams'] < 61])), NAVY),
             ('ЦІНА ЗРОСТАЄ З ВАГОЮ',
              'Типова упаковка: до 5 г — %d ₴, 8–15 г — %d ₴, 20–30 г — %d ₴, 35–60 г — %d ₴. Наші TMK 10–12 г — %d ₴, у %s рази вище за типову упаковку своєї смуги.' % (round(n5[1]), round(n8[1]), round(n20[1]), round(n35[1]), round(med(o['shelf'] for o in tm)), ('%.1f' % (med(o['shelf'] for o in tm) / n8[1])).replace('.', ',')), AMBER)]
    for k, (ttl, body, col) in enumerate(boxes):
        x = 0.62 + k * 6.15
        rect(s, x, y, 5.95, 0.52, fill=PANEL); rect(s, x, y, 0.04, 0.52, fill=col)
        text(s, x + 0.2, y + 0.04, 5.6, 0.14, ttl, size=6.4, color=col, bold=True, spc=0.2)
        text(s, x + 0.2, y + 0.19, 5.65, 0.34, body, size=7.0, color=INK)
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
            text(s, px(v) - 0.5, yy + row_h * 0.52 + 0.19 + slot * 0.17, 1.0, 0.16, unit_fmt(v), size=6.8, color=SUP_COLOR[sid], bold=True, align='c')
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
    text(s, 0.62, end + 0.32, 12.1, 0.4, 'Кольорові кружки — наша полиця (₴ за упаковку), смуга — діапазон цін ринку в сегменті, квадрати — лінійки брендів, риска — типова ціна (медіана). Tao Kae Noi 59 г коштує 528 ₴ в Onde — за шкалою. Рисові чипси — не наш формат, не показано.', size=7.8, color=MUTED)
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
    cnt = collections.Counter(c for r in rs for c in r['chains'])
    ncn = len({c for r in rs for c in r['chains'] if c in NATIONAL}); nre = len({c for r in rs for c in r['chains'] if c not in NATIONAL})
    tk = sorted({CHAIN[c] for r in ROWS if r['brand'] == 'Tao Kae Noi' for c in r['chains']})
    ww = sorted({CHAIN[c] for r in ROWS if r['brand'] == 'Want Want' for c in r['chains']})
    mini = [r['pmed'] for r in rows_of('mini')]
    text(s, 0.62, 1.34, 12.1, 0.45, 'Ліворуч — скільки позицій нори-снеків і рисових крекерів стоїть у кожній мережі; праворуч — чотири канали: що в них знайшли, за які гроші і що це означає для нашого асортименту.', size=8.6, color=MUTED)
    caps(s, 0.62, 1.95, 5, 'Позицій у мережі', size=6.8)
    order = sorted([c for c in cnt], key=lambda c: -cnt[c])
    hbars(s, 0.62, 2.25, 5.2, [(CHAIN[c], cnt[c], (NAVY if c in NATIONAL else TEAL)) for c in order], color=NAVY, row_h=min(0.3, 4.3 / len(order)), name_w=1.3, val_w=0.5, maxv=max(cnt.values()), size=8.4)
    rect(s, 0.62, 6.78, 0.13, 0.13, fill=NAVY); text(s, 0.82, 6.75, 1.6, 0.2, 'національні (%d)' % ncn, size=7, color=MUTED)
    rect(s, 2.4, 6.78, 0.13, 0.13, fill=TEAL); text(s, 2.6, 6.75, 1.8, 0.2, 'регіональні (%d)' % nre, size=7, color=MUTED)
    cards = [
        ('Супермаркети · %d мереж' % len(cnt), 'ЦІЛЬОВИЙ КАНАЛ', RED,
         ['Знайшли: мініпакети 4–5 г (%d ₴ типова), чипси 25 г (47–88 ₴), темпура Tao Kae Noi — лише %s, рисові крекери Want Want mini 60 г (140–149 ₴) — %s.' % (round(med(mini)), ', '.join(tk), ', '.join(ww)),
          'Постачальники за митною базою: Metro Cash & Carry (Metro Chef), Норріс Груп, ТД ІТС (Haelove), Оріенталь Плюс (Akura).',
          'Для нас: через мережі йдуть мініпакети — це TMK ROLL / DOUBLE; ZEK й рисові крекери в мережах майже без конкурентів.']),
        ('Азійські інтернет-магазини · Asia Foods, Японський Квартал, До Смаку, Суші Повар', 'ЦІЛЬОВИЙ КАНАЛ', AMBER,
         ['Знайшли: Tao Kae Noi Hot & Spicy 15 г — 107 ₴, Wasabi 15 г — 78 ₴, Roll 6 × 3 г — 89 ₴; Kimnori 5 г — 55 ₴; Akura 4,5 г — 35 ₴; Want Want 160 г — 357 ₴.',
          'Для нас: тут стоять формати 15–160 г, яких немає в мережах, — місце для ZEK Tempura / Topping і рисових крекерів.']),
        ('Маркетплейси · Rozetka, Prom.ua', 'ВХІД', TEAL,
         ['Знайшли: Akura 3 × 4,5 г — 129 ₴ і 89 ₴ (Rozetka); Tao Kae Noi Big Roll 6 × 3 г — 154,85 ₴; Edward & Sons 100 г — 351 ₴; опт Akura 72 × 4,5 г — 1 820 ₴ (Prom.ua).',
          'Для нас: ціни тут вищі за мережі, є опт і преміум — канал для невеликих партій.']),
        ('HoReCa · суші-бари, азійські ресторани', 'НЕ ОХОПЛЕНО', GREY,
         ['У каталогах мереж є листи норі для суші (Katana, JS, Hokkaido Club, Akura) — інша категорія, у зріз не входить. Окремий аудит HoReCa не проводили.']),
    ]
    x0, y = 6.1, 1.95
    heights = [1.62, 1.28, 1.2, 0.76]
    for (ttl, tag, col, lns), h in zip(cards, heights):
        rect(s, x0, y, 6.62, h, fill=PANEL); rect(s, x0, y, 0.05, h, fill=col)
        text(s, x0 + 0.2, y + 0.08, 4.5, 0.4, ttl, size=8.8, color=NAVY, bold=True)
        rect(s, x0 + 4.85, y + 0.1, 1.65, 0.22, fill=col)
        text(s, x0 + 4.85, y + 0.1, 1.65, 0.22, tag, size=6.2, color=WHITE, bold=True, align='c', anchor='m')
        yy = y + (0.48 if len(ttl) > 40 else 0.4)
        for t in lns:
            text(s, x0 + 0.2, yy, 6.25, 0.5, '•  ' + t, size=7.2, color=INK)
            yy += 0.27 + 0.15 * (len(t) // 100)
        y += h + 0.08
    footnote(s, 'Мережі — каталоги zakaz.ua, 07.10.2026; імпортери — митна база. Онлайн-магазини й маркетплейси — оголошення, знайдені пошуком; повного аудиту цих каналів не проводили, ціни можуть змінитись.')
    return s


# ================================================================ our pricing
def clean_title(o):
    t = o['title'].replace('Seaweed Topping', 'Topping').replace('Tempura Seaweed', 'Tempura').replace('Sandwich Seaweed', 'Sandwich').replace('Arare Norimaki', 'Norimaki')
    import re
    return re.sub(r'\s*·\s*\d+\s*г$', '', t)


def pricing_formula(deck, sids, part):
    s = deck.slide('Ціноутворення · %s · %d з 2' % (', '.join(SUP_NAME[x] for x in sids), part), 'З чого складається ціна на полиці')
    rect(s, 0.62, 1.45, 12.1, 1.38, fill=PANEL)
    caps(s, 0.85, 1.56, 8, 'Ціна партнеру — за неї ми продаємо мережі (100 %)', size=7)
    parts = [('Собівартість %d %%' % SSP, SSP, NAVY), ('Бонус мережі 25 %', 25, AMBER), ('Наша маржа %d %%' % MP, MP, GOLD), ('Націнка магазину +40 %', 40, GREY)]
    tot = 100 + 40; w = 11.3; x = 0.85
    for t, v, c in parts:
        ww = w * v / tot
        rect(s, x, 1.84, ww - 0.03, 0.38, fill=c)
        text(s, x, 1.84, ww, 0.38, t, size=8.6, color=WHITE if c != GOLD else INK, bold=True, align='c', anchor='m')
        x += ww
    rect(s, 0.85, 2.28, w * 100 / tot, 0.015, fill=INK)
    text(s, 0.85, 2.3, w * 100 / tot, 0.2, '= ціна партнеру (ціна, за якою купує мережа)', size=7.8, color=INK, bold=True, align='c')
    text(s, 0.85, 2.55, 11.3, 0.22, 'ПОЛИЦЯ (ціна покупцю) = ціна партнеру × 1,40 · ціна партнеру = собівартість ÷ %s · бонус 25 %% і маржа %d %% — частки ціни партнеру' % (('%.2f' % (SSP / 100)).replace('.', ','), MP), size=8.0, color=NAVY, bold=True, align='c')
    items = [o for o in OUR if o['sup'] in sids]
    hy = 2.98
    caps(s, 0.62, hy, 4, 'Позиція (кожен SKU)', size=6.4)
    caps(s, 2.9, hy, 5, 'Із чого складається полиця, ₴ за упаковку', size=6.4)
    for edge, t in [(8.7, 'ЦІНА ПАРТНЕРУ'), (9.8, 'FOB, $'), (10.9, 'СС, $'), (12.7, 'СС, ГРН')]:
        caps(s, edge - 1.2, hy, 1.2, t, size=6.2, align='r')
    ry0 = 3.25
    rh = min(0.30, (6.8 - ry0) / len(items))
    mx = max(o['shelf'] for o in OUR)
    bw = 3.3
    for i, o in enumerate(items):
        y = ry0 + i * rh
        rect(s, 0.62, y, 12.1, rh - 0.03, fill=PANEL)
        rect(s, 0.62, y, 0.05, rh - 0.03, fill=SUP_COLOR[o['sup']])
        text(s, 0.78, y, 2.15, rh - 0.03, '%s %s%s' % (clean_title(o), gfmt(o['grams']), (' · %d смаки' % o.get('n_sku', 1)) if o.get('n_sku', 1) > 1 else ''), size=7.4, color=NAVY, bold=True, anchor='m')
        x = 2.95
        sc = bw / mx
        for v, c, fg in [(o['cost'], NAVY, WHITE), (o['bonus'], AMBER, WHITE), (o['profit'], GOLD, INK), (o['shelf'] - o['partner'], GREY, WHITE)]:
            ww = v * sc
            rect(s, x, y + 0.04, ww - 0.02, rh - 0.11, fill=c)
            if ww > 0.34: text(s, x, y + 0.02, ww, rh - 0.07, '%d' % round(v), size=6.6, color=fg, bold=True, align='c', anchor='m')
            x += ww
        text(s, x + 0.08, y, 0.9, rh - 0.03, 'полиця %d' % round(o['shelf']), size=8.6, color=INK, bold=True, anchor='m')
        text(s, 7.5, y, 1.2, rh - 0.03, '%d' % round(o['partner']), size=8.8, color=INK, bold=True, anchor='m', align='r')
        text(s, 8.6, y, 1.2, rh - 0.03, '$%s' % fnum(o['fob'], 3), size=7.8, color=MUTED, anchor='m', align='r')
        text(s, 9.7, y, 1.2, rh - 0.03, '$%s' % fnum(o['usd'], 3), size=7.8, color=MUTED, anchor='m', align='r')
        text(s, 11.5, y, 1.2, rh - 0.03, fnum(o['cost'], 1), size=7.8, color=MUTED, anchor='m', align='r')
    ly = ry0 + rh * len(items) + 0.05
    for k, (t, c) in enumerate([('Собівартість (СС)', NAVY), ('Бонус мережі', AMBER), ('Наша маржа', GOLD), ('Націнка магазину', GREY)]):
        rect(s, 0.7 + k * 2.2, ly + 0.03, 0.13, 0.13, fill=c)
        text(s, 0.92 + k * 2.2, ly, 2.0, 0.2, t, size=7.4, color=INK)
    footnote(s, 'СС — «ИТОГО С/С» з Self-Cost_Snacks (FOB + мито + логістика + ПДВ), курс 45 ₴/$, найдешевший контейнер для SKU (переважно 40′). Повна модель по всіх контейнерах і маржах — SNACKS_FINMODEL_ALL_MARGINS.xlsx.')
    return s


def shelf_table(deck, sids, market_cards=None):
    items = [o for o in OUR if o['sup'] in sids]
    col = SUP_COLOR[sids[0]]
    names = ' і '.join(SUP_NAME[x] for x in sids)
    s = deck.slide('Полиця · %s' % names, 'Наша ціна на полиці: %s' % names)
    # (title, right edge from x=0.62, width, align)
    cols = [('ВАГА', 3.45, 0.6), ('FOB, $', 4.2, 0.7), ('СС, ГРН', 5.0, 0.75), ('КОНТЕЙНЕР', 5.95, 0.85), ('ЦІНА МЕРЕЖІ', 6.9, 0.9), ('ПОЛИЦЯ', 7.85, 0.85)]
    n = len(items)
    text(s, 0.62, 1.34, 12.1, 0.5, 'Як рахується: ціна мережі = собівартість ÷ %s (собівартість — %d %% ціни мережі, бонус мережі — 25 %%, наша маржа — %d %%); полиця = ціна мережі × 1,40. Праворуч — із чим порівнюється наша полиця на ринку.' % (('%.2f' % (SSP / 100)).replace('.', ','), SSP, MP), size=8.4, color=MUTED)
    top = 1.98
    rh = min(0.42, (6.75 - top - 0.35) / n) if market_cards is None else 0.42
    rect(s, 0.62, top - 0.05, 12.1, 0.3, fill=col)
    text(s, 0.62 + 0.75, top, 2.5, 0.22, 'ПОЗИЦІЯ', size=6.2, color=WHITE, bold=True, anchor='m')
    for t, edge, w in cols:
        text(s, 0.62 + edge - w, top, w, 0.22, t, size=6.2, color=WHITE, bold=True, align='r', anchor='m')
    text(s, 0.62 + 8.1, top, 2.6, 0.22, 'РИНОК ПОРУЧ · ТИПОВА ЦІНА ПАКЕТА', size=6.2, color=WHITE, bold=True, anchor='m')
    text(s, 0.62 + 10.5, top, 1.6, 0.22, 'РІЗНИЦЯ ДО РИНКУ', size=6.2, color=WHITE, bold=True, align='r', anchor='m')
    for i, o in enumerate(items):
        y = top + 0.35 + i * rh
        if i % 2 == 0:
            rect(s, 0.62, y, 12.1, rh, fill=rgb('F8F9FC'))
        rect(s, 0.62, y, 0.05, rh, fill=SUP_COLOR[o['sup']])
        if os.path.exists(o['photo']):
            picture(s, o['photo'], 0.72, y + 0.02, 0.5, rh - 0.04)
        text(s, 0.62 + 0.75, y, 2.0, rh, clean_title(o), size=8, color=NAVY, bold=True, anchor='m')
        vals = [gfmt(o['grams']), '$' + fnum(o['fob'], 3), fnum(o['cost'], 1), o['scenario'].replace(' контейнер', ''), '%d ₴' % round(o['partner']), '%d ₴' % round(o['shelf'])]
        for (t, edge, w), v in zip(cols, vals):
            bold = t in ('ПОЛИЦЯ', 'ЦІНА МЕРЕЖІ')
            c = INK if bold else MUTED
            text(s, 0.62 + edge - w, y, w, rh, v, size=9.4 if t == 'ПОЛИЦЯ' else 7.8, color=c, bold=bold, anchor='m', align='r')
        bk = D.basket(o)
        gw = ('%s–%s' % (('%g' % bk['gmin']).replace('.', ','), gfmt(bk['gmax']))) if bk['gmin'] != bk['gmax'] else gfmt(bk['gmax'])
        text(s, 0.62 + 8.1, y + 0.02, 2.5, rh * 0.55, [[('%d ₴' % round(bk['pmed']), {'bold': True, 'color': INK, 'size': 9.4}), ('   типова', {'color': MUTED, 'size': 7})]], anchor='m')
        text(s, 0.62 + 8.1, y + rh * 0.52, 2.5, rh * 0.45, '%d поз. · %s · %s ₴' % (bk['n'], gw, rng(bk['pmin'], bk['pmax'])), size=6.8, color=MUTED, anchor='t')
        diff = o['shelf'] - bk['pmed']; pc = round(diff / bk['pmed'] * 100)
        c = RED if pc > 25 else (GREEN if pc < -10 else AMBER)
        text(s, 0.62 + 10.5, y, 1.6, rh, [[('%s%d ₴' % ('+' if diff > 0 else '−', abs(round(diff))), {'bold': True, 'color': c, 'size': 9.4}), ('  %+d %%' % pc, {'color': c, 'size': 7.6})]], anchor='m', align='r')
    if market_cards:
        y0 = top + 0.35 + n * rh + 0.2
        section_tag(s, 0.62, y0, 8, 'Що стоїть поруч на полиці мереж · рисові крекери', color=RED)
        for k, l in enumerate(market_cards):
            x = 0.62 + k * 4.1
            rect(s, x, y0 + 0.32, 3.95, 2.0, fill=PANEL); rect(s, x, y0 + 0.32, 3.95, 0.05, fill=RED)
            rect(s, x + 0.14, y0 + 0.48, 1.5, 1.5, fill=WHITE)
            picture(s, os.path.join(HERE, l['img']), x + 0.18, y0 + 0.52, 1.42, 1.42)
            text(s, x + 1.8, y0 + 0.5, 2.05, 0.5, '%s · %s' % (l['brand'], gfmt(l['grams'])), size=10, color=INK, bold=True)
            text(s, x + 1.8, y0 + 1.0, 2.05, 0.36, '%s ₴' % rng(l['pmin'], l['pmax']), size=16, color=INK, bold=True)
            text(s, x + 1.8, y0 + 1.5, 2.05, 0.4, 'Мережі: ' + ' · '.join(CHAIN[c] for c in l['chains'][:3]), size=7.6, color=MUTED)
    footnote(s, '«Типова ціна» — медіана цін позицій того ж сегмента з близькою вагою (±45 %%); «Різниця до ринку» = наша полиця − типова ціна. Розрахунок: курс 45 ₴/$, бонус мережі 25 %%, наша маржа %d %%. СС з Self-Cost включає імпортний ПДВ (≈16 %% СС).' % MP)
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
        takeaways = None
    text(s, 0.62, 1.34, 12.1, 0.55, sub or '', size=8.2 if compact else 9.0, color=MUTED)
    top = 1.98 if compact else 1.74
    cap = 0.285 if n > 10 else 0.42
    rh = min(cap, (6.62 - top - 0.34) / n) if not takeaways else min(cap, (6.95 - top - 0.34 - 1.65) / n)
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
            text(s, x + 0.2, ty + 0.32, w - 0.5, 0.8, body, size=8.6, color=INK)
    src_note(s)
    return s


def neighbours(deck):
    s = deck.slide('Хто з ким стоїть · сусіди', 'Найближчі за ціною до кожного нашого товару')
    mk = []
    for sg in NORI_SEGS:
        for l in lines(rows_of(sg)):
            mk.append(dict(l, seg=sg))
    text(s, 0.62, 1.36, 12.1, 0.22, 'Для кожної нашої позиції (смаки з однаковою ціною — одним рядком) — три найближчі за ціною упаковки ринку з того самого сегмента.', size=9.0, color=MUTED)
    groups = D.our_groups()
    top = 1.76
    rh = min(0.37, 4.85 / (len(groups) + 0.5))
    rect(s, 0.62, top, 12.1, 0.28, fill=NAVY)
    for t, dx, w in [('НАШ ТОВАР', 0.1, 2.5), ('ПОЛИЦЯ', 2.7, 0.8), ('СУСІД 1', 3.75, 2.8), ('СУСІД 2', 6.65, 2.8), ('СУСІД 3', 9.55, 2.8)]:
        text(s, 0.62 + dx, top, w, 0.28, t, size=6.4, color=WHITE, bold=True, anchor='m')
    for i, o in enumerate(groups):
        y = top + 0.32 + i * rh
        if i % 2 == 0: rect(s, 0.62, y, 12.1, rh, fill=rgb('F8F9FC'))
        rect(s, 0.62, y, 0.05, rh, fill=SUP_COLOR[o['sup']])
        text(s, 0.78, y, 2.6, rh, '%s %s%s' % (clean_title(o), gfmt(o['grams']), (' · %d смаки' % o.get('n_sku', 1)) if o.get('n_sku', 1) > 1 else ''), size=7.4, color=NAVY, bold=True, anchor='m')
        text(s, 3.32, y, 0.8, rh, '%d ₴' % round(o['shelf']), size=9, color=INK, bold=True, anchor='m')
        pool = [l for l in mk if l['seg'] == o['seg'] and 0.4 * o['grams'] <= l['grams'] <= 2.5 * o['grams']]
        pool = pool if len(pool) >= 3 else [l for l in mk if l['seg'] == o['seg']]
        cand = sorted(pool, key=lambda l: abs(l['pmed'] - o['shelf']))[:3]
        for k, l in enumerate(cand):
            dpc = (l['pmed'] / o['shelf'] - 1) * 100
            near = abs(dpc) <= 20
            text(s, 4.37 + k * 2.9, y, 2.85, rh, [[('%s %s  ' % (l['brand'], gfmt(l['grams'])), {'bold': True, 'color': NAVY}), ('%d ₴' % round(l['pmed']), {'bold': True, 'color': INK}),
                                                  ('  %+d %%' % round(dpc), {'color': GREEN if near else RED, 'bold': True})]], size=7.2, anchor='m')
    footnote(s, 'Відсоток біля сусіда — на скільки його упаковка дорожча (+) або дешевша (−) за нашу полицю. Зелений — у межах ±20 % від нашої ціни, червоний — далі. Вага упаковок може відрізнятися.')
    return s


def verification(deck):
    from sc_breakdown import breakdown
    s = deck.slide('Перевірка розрахунку', 'Перевірка: як з FOB виходить полиця')
    text(s, 0.62, 1.34, 12.1, 0.45, 'Усі числа взято з вашого Self-Cost_Snacks (контейнер 40′, рядок вказано під назвою) і перераховано: собівартість = FOB + мито + імпортний ПДВ + транспорт + кредитні гроші; ціна мережі = собівартість ÷ %s; полиця = ціна мережі × 1,40.' % ('%.2f' % (SSP / 100)).replace('.', ','), size=8.4, color=MUTED)
    picks = [('singha', 'SINGHA KAMEDA - Retail', 0, 'Singha · Norimaki Original 42 г'), ('thainichi', 'Thai-Nichi - Retail', 0, 'Thai-Nichi · Norimaki Original 50 г'),
             ('tmk', 'TMK Thailand Co., Ltd -Retail', 0, 'TMK · WOW ROLL Original 2,5 г'), ('zek', 'ZEK -Retail', 0, 'ZEK · Tempura Corn 30 г')]
    for i, (sid, sheet, idx, ttl) in enumerate(picks):
        b = breakdown(sheet, idx)
        x = 0.62 + i * 3.06; w = 2.95; col = SUP_COLOR[sid]
        rect(s, x, 1.95, w, 0.36, fill=col)
        text(s, x + 0.12, 2.03, w - 0.2, 0.22, ttl.upper(), size=7.2, color=WHITE, bold=True)
        rect(s, x, 2.31, w, 4.72, fill=PANEL)
        rate = 45.0
        pr = b['parts']
        rows = [('FOB', pr['fob']), ('Мито', pr['duty']), ('Імпортний ПДВ', pr['vat']), ('Транспорт', pr['transport'] + pr['other']), ('Кредитні гроші', pr['credit'])]
        mx = b['ss_usd']
        text(s, x + 0.14, 2.4, w - 0.28, 0.16, 'СОБІВАРТІСТЬ УПАКОВКИ · рядок %d' % b['row'], size=5.8, color=col, bold=True, spc=0.2)
        yy = 2.62
        for lab, v in rows:
            text(s, x + 0.14, yy, 1.1, 0.2, lab, size=7.2, color=INK)
            bwid = 0.5 * v / mx
            rect(s, x + 1.25, yy + 0.04, max(0.03, bwid), 0.13, fill=col if lab == 'FOB' else tint(col, 0.45))
            text(s, x + 1.25 + bwid + 0.05, yy, 1.4, 0.2, '$%s · %s ₴' % (('%.3f' % v).replace('.', ','), ('%.2f' % (v * rate)).replace('.', ',')), size=6.4, color=MUTED)
            yy += 0.24
        rect(s, x + 0.14, yy + 0.03, w - 0.28, 0.012, fill=INK)
        text(s, x + 0.14, yy + 0.08, w - 0.28, 0.24, 'СС = $%s × 45 = %s ₴' % (('%.4f' % b['ss_usd']).replace('.', ','), ('%.2f' % b['ss_uah']).replace('.', ',')), size=8.8, color=INK, bold=True)
        ss = b['ss_usd'] * 45
        partner = ss / (SSP / 100)
        bon = partner * BON; mar = partner * MAR; shelf = partner * MUP
        yy += 0.5
        lines_ = [('Ціна мережі', '%s ÷ %s = %s ₴' % (('%.2f' % ss).replace('.', ','), ('%.2f' % (SSP / 100)).replace('.', ','), ('%.2f' % partner).replace('.', ',')), False),
                  ('Бонус 25 %', '%s ₴' % ('%.2f' % bon).replace('.', ','), False),
                  ('Маржа %d %%' % MP, '%s ₴' % ('%.2f' % mar).replace('.', ','), False),
                  ('Собівартість', '%s ₴' % ('%.2f' % ss).replace('.', ','), False),
                  ('Сума (перевірка)', '%s ₴' % ('%.2f' % (bon + mar + ss)).replace('.', ','), False)]
        for lab, v, bd in lines_:
            text(s, x + 0.14, yy, 1.0, 0.2, lab, size=7.4, color=INK)
            text(s, x + 1.0, yy, w - 1.14, 0.2, v, size=7.2, color=INK, bold=(lab == 'Ціна мережі'), align='r')
            yy += 0.25
        yy += 0.1
        rect(s, x + 0.14, yy, w - 0.28, 0.6, fill=WHITE)
        text(s, x + 0.2, yy + 0.04, w - 0.4, 0.2, 'ПОЛИЦЯ = %s × 1,40' % ('%.2f' % partner).replace('.', ','), size=7.4, color=MUTED, bold=True)
        text(s, x + 0.2, yy + 0.24, w - 0.4, 0.34, '%s ₴' % ('%.2f' % shelf).replace('.', ','), size=16, color=col, bold=True)
        yy += 0.72
        text(s, x + 0.14, yy, w - 0.28, 0.6, 'ПДВ у собівартості: %d %% СС. Полиця / FOB = ×%s.' % (round(100 * pr['vat'] / b['ss_usd']), ('%.1f' % (shelf / (pr['fob'] * 45))).replace('.', ',')), size=7, color=MUTED)
    footnote(s, 'Імпортний ПДВ входить у собівартість за методикою Self-Cost (так само, як у фінмоделі рису); ціна мережі і полиця — вже з ПДВ.')
    return s


def scenarios(deck):
    s = deck.slide('Перевірка розрахунку · контейнери', 'Полиця за сценаріями поставки')
    text(s, 0.62, 1.34, 12.1, 0.45, 'Полиця при марже %d %% для кожного сценарію поставки з Self-Cost. У презентації для кожного SKU береться найдешевший (виділено); «Retail + Bulk» — той самий товар у змішаному контейнері з навалом, де собівартість нижча.' % MP, size=8.4, color=MUTED)
    scen = [("20'", '20′'), ("40'", '40′'), ('LCL 17', 'LCL 17'), ('LCL 34', 'LCL 34')]
    byk = collections.OrderedDict()
    for r in PM['rows']:
        byk.setdefault((r['supplier_id'], r['title']), {})[r['scenario']] = r
    seen = set(); rows = []
    for o in OUR:
        key = (o['sup'], o['title'])
        if key in seen: continue
        seen.add(key); rows.append((o, byk[key]))
    # collapse flavours with identical numbers
    uniq = collections.OrderedDict()
    for o, sc in rows:
        sig = (o['sup'], o['grams'], tuple(round(sc[k]['shelf_uah']) for k, _ in scen))
        uniq.setdefault(sig, (o, sc, []))[2].append(o)
    top = 1.95
    heads = [('ПОЗИЦІЯ', 0.1, 3.0, 'l'), ('FOB, $', 3.1, 0.7, 'r')] + [('СС %s' % lab, 3.95 + k * 0.85, 0.8, 'r') for k, (_, lab) in enumerate(scen)] + [('ПОЛИЦЯ %s' % lab, 7.5 + k * 1.15, 1.05, 'r') for k, (_, lab) in enumerate(scen)]
    rect(s, 0.62, top, 12.1, 0.3, fill=NAVY)
    for t, dx, w, al in heads:
        text(s, 0.62 + dx, top, w, 0.3, t, size=5.8, color=WHITE, bold=True, anchor='m', align=al)
    rh = min(0.32, (6.85 - top - 0.36) / len(uniq))
    for i, (sig, (o, sc, grp)) in enumerate(uniq.items()):
        y = top + 0.34 + i * rh
        if i % 2 == 0: rect(s, 0.62, y, 12.1, rh, fill=rgb('F8F9FC'))
        rect(s, 0.62, y, 0.05, rh, fill=SUP_COLOR[o['sup']])
        text(s, 0.78, y, 3.0, rh, '%s %s%s' % (clean_title(o), gfmt(o['grams']), (' · %d смаки' % len(grp)) if len(grp) > 1 else ''), size=7.4, color=NAVY, bold=True, anchor='m')
        text(s, 0.62 + 3.1, y, 0.7, rh, '$%s' % fnum(o['fob'], 3), size=7.4, color=MUTED, anchor='m', align='r')
        cheapest = min(sc.values(), key=lambda r: r['cost_usd'])['scenario']
        for k, (key, lab) in enumerate(scen):
            r = sc[key]
            text(s, 0.62 + 3.95 + k * 0.85, y, 0.8, rh, fnum(r['cost_uah'], 1), size=7.4, color=MUTED, anchor='m', align='r')
            hit = key == cheapest
            if hit: rect(s, 0.62 + 7.5 + k * 1.15 + 0.1, y + 0.03, 0.95, rh - 0.06, fill=tint(GOLD, 0.35))
            text(s, 0.62 + 7.5 + k * 1.15, y, 1.05, rh, '%d ₴' % round(r['shelf_uah']), size=8.2, color=INK, bold=hit, anchor='m', align='r')
    ly = top + 0.34 + rh * len(uniq) + 0.1
    rb = lambda sh, k: D.DATA[sh][0]['cost'][k]['uah']
    sk_b, sk_r = rb('SINGHA KAMEDA - Retail+Bulk', "40'"), rb('SINGHA KAMEDA - Retail', "40'")
    tn_b, tn_r = rb('Thai-Nichi - Retail+Bulk', "40'"), rb('Thai-Nichi - Retail', "40'")
    text(s, 0.62, ly, 12.1, 0.5, 'Для порівняння: у файлі Retail + Bulk собівартість того ж товару в 40′ нижча — Singha Original 42 г: %s ₴ → полиця %d ₴; Thai-Nichi Original 50 г: %s ₴ → %d ₴ (у таблиці Retail: %s ₴ і %s ₴). Для презентації взято таблицю Retail, як у каталозі постачальників.' % (fnum(sk_b, 1), round(sk_b / (SSP / 100) * MUP), fnum(tn_b, 1), round(tn_b / (SSP / 100) * MUP), fnum(sk_r, 1), fnum(tn_r, 1)), size=7.6, color=MUTED)
    footnote(s, 'Self-Cost_Snacks: «ИТОГО С/С, UAH ед. товара»; собівартість у таблиці = $ × 45. Ціна мережі = СС ÷ %s; полиця = ціна мережі × 1,40.' % ('%.2f' % (SSP / 100)).replace('.', ','))
    return s


def sensitivity(deck):
    s = deck.slide('Чутливість', 'Полиця при різній марже: усі маржі на одному слайді')
    ms = ['0.35', '0.3', '0.25', '0.2', '0.15']
    sens = {r['title']: r for r in PM['sensitivity']}
    groups = D.our_groups()
    text(s, 0.62, 1.34, 12.1, 0.45, 'Наша маржа 35 → 15 %% ціни партнеру (бонус мережі 25 %%, націнка ×1,40). Основна колонка — маржа %d %% (жирним). Зелена клітинка — полиця не вище за типову ціну близьких за вагою позицій ринку; червона — вище за весь їхній діапазон.' % MP, size=8.4, color=MUTED)
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
        text(s, 0.78, y, 3.0, rh, '%s %s%s' % (clean_title(o), gfmt(o['grams']), (' · %d смаки' % o.get('n_sku', 1)) if o.get('n_sku', 1) > 1 else ''), size=7.6, color=NAVY, bold=True, anchor='m')
        ps = D.peer_stats(o)
        text(s, 3.7, y, 2.0, rh, '%d ₴' % round(ps['pack']), size=7.6, color=MUTED, anchor='m', align='r')
        sv = sens[o['title']]['shelf_by_margin']
        for k, m in enumerate(ms):
            v = sv[m]
            bg = tint(GREEN, 0.22) if v <= ps['pack'] else (tint(RED, 0.25) if v > ps['hi'] else None)
            if bg is not None:
                rect(s, X0 + k * STEP + 0.02, y + 0.03, W, rh - 0.06, fill=bg)
            text(s, X0 + k * STEP, y, W - 0.1, rh, '%d ₴' % round(v), size=8.2, color=INK, bold=(m == str(MAR)), anchor='m', align='r')
    footnote(s, 'Близькі за вагою позиції — той самий сегмент, вага ×0,5–×2 від нашої. Excel SNACKS_FINMODEL_ALL_MARGINS.xlsx: базова маржа в клітинці D5 — 35 %%, значення при 30 %% — блок «МАРЖА 30 %%» (колонка U).')
    return s


def price_by_chain(deck, sgs, title, eyebrow, must=(), max_cols=8):
    """Competitor lines (columns) x chains (rows): the price of one pack in each network."""
    seg_col = SEG_COLOR[sgs[0]]
    ls = [l for sg in sgs for l in lines(rows_of(sg))]
    ls.sort(key=lambda l: -(l['n'] * len(l['chains'])))
    cols = ls[:max_cols]
    for mname in must:
        if not any(c['brand'] == mname for c in cols):
            extra = [l for l in ls if l['brand'] == mname]
            if extra: cols = cols[:max_cols - 1] + extra[:1]
    cols.sort(key=lambda l: (l['grams'], l['brand']))
    chains = [c for c in CHAINS_ORDER if any(c in l['chain_price'] for l in cols)]
    s = deck.slide(eyebrow, title)
    sub_line(s, 'Ціна однієї упаковки (₴) у кожній мережі — найпопулярніші лінійки сегмента. Порожньо — лінійки в мережі немає; зелений — найдешевше для лінійки, червоний — найдорожче. Нижні рядки: діапазон між мережами й наша полиця поруч (при марже %d %%).' % MP, y=1.34, size=8.6)
    x0, y0 = 0.62, 1.95
    namew = 1.7
    cw = min(1.45, (12.1 - namew) / len(cols))
    head_h = 1.25
    for j, l in enumerate(cols):
        x = x0 + namew + j * cw
        hit = l['brand'] in must
        rect(s, x + 0.03, y0, cw - 0.06, head_h, fill=tint(seg_col, 0.22) if hit else PANEL)
        picture(s, os.path.join(HERE, l['img']), x + 0.1, y0 + 0.05, cw - 0.2, 0.7)
        text(s, x + 0.05, y0 + 0.78, cw - 0.1, 0.2, l['brand'], size=8, color=INK, bold=True, align='c')
        text(s, x + 0.05, y0 + 0.98, cw - 0.1, 0.2, '%s · %d смаки' % (gfmt(l['grams']), l['n']) if l['n'] > 1 else '%s · 1 смак' % gfmt(l['grams']), size=6.6, color=MUTED, align='c')
    rh = min(0.62, (6.3 - y0 - head_h - 0.1) / (len(chains) + 2))
    fs_ = 10.5 if rh > 0.45 else (9 if rh > 0.3 else 8)
    for i, c in enumerate(chains):
        y = y0 + head_h + 0.08 + i * rh
        if i % 2 == 0: rect(s, x0, y, namew + cw * len(cols), rh, fill=rgb('F8F9FC'))
        text(s, x0 + 0.08, y, namew - 0.1, rh, CHAIN[c], size=fs_ - 0.5, color=NAVY, bold=True, anchor='m')
        for j, l in enumerate(cols):
            v = l['chain_price'].get(c)
            x = x0 + namew + j * cw
            if v is None:
                text(s, x, y, cw, rh, '–', size=7.5, color=GREY, align='c', anchor='m')
            else:
                lo, hi = min(l['chain_price'].values()), max(l['chain_price'].values())
                col = GREEN if (hi > lo and v <= lo + 0.001) else (RED if (hi > lo and v >= hi - 0.001) else INK)
                text(s, x, y, cw, rh, '%d' % round(v), size=fs_, color=col, bold=True, align='c', anchor='m')
    yb = y0 + head_h + 0.08 + len(chains) * rh
    rect(s, x0, yb, namew + cw * len(cols), 0.012, fill=INK)
    text(s, x0 + 0.08, yb + 0.03, namew, rh, 'Діапазон, ₴', size=7.6, color=MUTED, bold=True, anchor='m')
    for j, l in enumerate(cols):
        text(s, x0 + namew + j * cw, yb + 0.03, cw, rh, rng(l['pmin'], l['pmax']), size=8, color=INK, bold=True, align='c', anchor='m')
    mine = [o for o in D.our_groups() if o['seg'] in sgs or (sgs[0] == 'tempura' and o['seg'] == 'tempura')]
    yo = yb + rh + 0.12
    rect(s, x0, yo, 12.1, 0.5, fill=rgb('FFF3DC')); rect(s, x0, yo, 0.04, 0.5, fill=GOLD)
    text(s, x0 + 0.2, yo + 0.04, 2.2, 0.42, 'НАША ПОЛИЦЯ ПОРУЧ', size=7, color=AMBER, bold=True, anchor='m')
    xx = x0 + 2.2
    for o in mine:
        t_ = '%s %s · %d ₴' % (clean_title(o), gfmt(o['grams']), round(o['shelf']))
        w_ = 0.25 + 0.064 * len(t_)
        if xx + w_ > x0 + 12.0: break
        rect(s, xx, yo + 0.12, w_, 0.26, fill=WHITE, line=SUP_COLOR[o['sup']], lw=1.0, radius=0.3)
        text(s, xx, yo + 0.12, w_, 0.26, t_, size=6.6, color=INK, bold=True, align='c', anchor='m')
        xx += w_ + 0.08
    src_note(s)
    return s
