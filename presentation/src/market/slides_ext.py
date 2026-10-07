# -*- coding: utf-8 -*-
"""Customs-base slides and product-explainer slides."""
import collections, json, os
from md import *  # noqa
from md import K, rect, text, caps, picture, kpi, section_tag, hbars, footnote, rgb, mixc
import mdata as D
from mdata import SEGS, rows_of, lines, med, gfmt, uah, rng, CHAIN, OUR
from slides_market import HERE

C = json.load(open(os.path.join(HERE, 'customs.json'), encoding='utf-8'))
MON = ['січ', 'лют', 'бер', 'кві', 'тра', 'чер', 'лип', 'сер', 'вер', 'жов', 'лис', 'гру']
CNOTE = ('Митна база: імпорт в Україну за УКТ ЗЕД, 2025 рік і січень–травень 2026 (17 місяців), фактурна вартість USD без імпортного ПДВ. '
         'Декларації не розбито на SKU, тому нори-снеки виділено за описом товару.')


def n1(v, d=1):
    return ('%.' + str(d) + 'f') % v if False else (('%.' + str(d) + 'f') % v).replace('.', ',')


def n0(v):
    return ('{:,.0f}'.format(v)).replace(',', ' ')


def cname(c):
    return {'КОРЕЯ РЕСПУБЛІКА': 'Корея', 'КИТАЙ': 'Китай', 'ТАЇЛАНД': 'Таїланд', 'ТУНІС': 'Туніс', 'КРАЇНИ ЄС': 'Країни ЄС', "В'ЄТНАМ": "В'єтнам", 'ПОЛЬЩА': 'Польща', 'ЯПОНІЯ': 'Японія'}.get(c, c.title())


def cnote(s):
    footnote(s, CNOTE)


# ---------------------------------------------------------------- customs 1
def customs_overview(deck):
    s = deck.slide('Обсяг ринку · митна база', 'Що показує митна база за нашими кодами')
    text(s, 0.62, 1.36, 12.1, 0.4, 'Чотири точні коди УКТ ЗЕД за 2025 рік і січень–травень 2026 (5 місяців). Кольорова смужка зліва — коди, під якими оформлюються наші продукти; сірі — локшина й готовий рис із попередніх досліджень, для масштабу.', size=8.8, color=MUTED)
    colr = {'2008999990': SUP_COLOR['tmk'], '1905905500': SUP_COLOR['singha']}
    ov = {o['code']: o for o in C['overview']}
    order = ['2008999990', '1905905500', '1902301000', '1904901000']
    mx = max(max(o['tons_2025'] / 12, o['tons_2026'] / 5) for o in C['overview'])
    x0, y0, rh = 0.62, 1.95, 1.2
    for i, c in enumerate(order):
        o = ov[c]; y = y0 + i * (rh + 0.08)
        rect(s, x0, y, 12.1, rh, fill=PANEL); rect(s, x0, y, 0.06, rh, fill=colr.get(c, GREY))
        text(s, x0 + 0.25, y + 0.1, 2.9, 0.22, '%s %s %s %s' % (c[:4], c[4:6], c[6:8], c[8:]), size=12, color=NAVY, bold=True)
        text(s, x0 + 0.25, y + 0.4, 2.9, 0.5, o['name'], size=8, color=MUTED)
        text(s, x0 + 0.25, y + 0.88, 2.9, 0.25, '%d імпортерів · %d декларацій' % (o['importers'], o['decl']), size=7.4, color=MUTED)
        text(s, x0 + 3.3, y + 0.1, 3.2, 0.3, '%s т' % n0(o['tons']), size=15, color=NAVY, bold=True)
        text(s, x0 + 3.3, y + 0.46, 3.2, 0.22, '$%s млн · $%s/кг' % (n1(o['usd'] / 1e6), n1(o['usd_kg'])), size=8.4, color=MUTED)
        text(s, x0 + 3.3, y + 0.72, 3.4, 0.4, 'найбільший імпортер %s — %d %%; головна країна %s — %d %%' % (o['top_importer'].replace('ТОВ ', '').replace('ПРАТ', 'ПрАТ '), round(100 * o['top_share']), cname(o['top_origin']), round(100 * o['origin_share'])), size=7.2, color=MUTED)
        bw = 3.6
        for k, (lab, tons, mths) in enumerate([('2025 · 12 міс.', o['tons_2025'], 12), ('2026 · січ–трав.', o['tons_2026'], 5)]):
            yy = y + 0.2 + k * 0.5
            text(s, x0 + 7.0, yy, 1.1, 0.3, lab, size=7.4, color=MUTED, bold=True, anchor='m')
            w = bw * (tons / mths) / mx * 0.6
            rect(s, x0 + 8.15, yy + 0.04, max(0.04, w), 0.22, fill=(NAVY if k == 0 else AMBER))
            text(s, x0 + 8.15 + w + 0.08, yy, 1.6, 0.3, '%s т · %s т/міс.' % (n0(tons), n1(tons / mths, 0 if tons / mths > 20 else 1)), size=7.8, color=NAVY, bold=True, anchor='m')
    cnote(s)
    return s


# ---------------------------------------------------------------- customs 2
def customs_seaweed(deck):
    sw = C['seaweed']; al = sw['all']; y25, y26 = sw['y2025'], sw['y2026']
    s = deck.slide('Митна база · нори-снеки', 'Нори-снеки: скільки і звідки ввозять')
    text(s, 0.62, 1.36, 12.1, 0.4, 'Нори-снеки виділено серед усього коду 2008 99 99 90 за описом товару (смажені / хрусткі водорості, снек, чипси, топінг, рол). Решта коду — фрукти, соуси, водорості для салатів — не рахується.', size=8.8, color=MUTED)
    kpi(s, 0.62, 1.9, 2.9, 1.2, 'Усього · 17 місяців', '%s т' % n1(al['tons']), '$%s млн · $%s/кг' % (n1(al['usd'] / 1e6, 2), n1(al['usd_kg'])))
    kpi(s, 3.68, 1.9, 2.9, 1.2, '2025 рік · 12 місяців', '%s т' % n1(y25['tons']), '%s т на місяць · $%s млн' % (n1(y25['tons'] / 12), n1(y25['usd'] / 1e6, 2)), color=NAVY)
    kpi(s, 6.74, 1.9, 2.9, 1.2, '2026 · січень–травень', '%s т' % n1(y26['tons']), '%s т на місяць · $%s млн' % (n1(y26['tons'] / 5), n1(y26['usd'] / 1e6, 2)), color=AMBER)
    kpi(s, 9.8, 1.9, 2.9, 1.2, 'Зміна темпу імпорту', '%+d %%' % round((y26['tons'] / 5) / (y25['tons'] / 12) * 100 - 100), 'т/міс.: 2026 проти 2025', color=GREEN)
    section_tag(s, 0.62, 3.3, 8, 'Імпорт помісячно, тонн · темно — чисті снеки, світло — змішані декларації (снеки + листи для суші)')
    ma, mb = sw['monthly_A'], sw['monthly_B']
    mx = max(a + b for a, b in zip(ma, mb))
    cx, cy, cw, ch = 0.62, 3.65, 8.0, 3.2
    n = len(ma); gap = 0.07; bw = (cw - gap * (n - 1)) / n
    y_base = cy + ch - 0.5
    for i, (a, b) in enumerate(zip(ma, mb)):
        x = cx + i * (bw + gap)
        ha = (ch - 0.85) * a / mx; hb = (ch - 0.85) * b / mx
        yr = int(C['months'][i][:4])
        rect(s, x, y_base - ha, bw, ha, fill=(NAVY if yr == 2025 else AMBER))
        if hb > 0: rect(s, x, y_base - ha - hb, bw, hb, fill=tint(NAVY if yr == 2025 else AMBER, 0.4))
        text(s, x - 0.1, y_base - ha - hb - 0.2, bw + 0.2, 0.18, n1(a + b, 1), size=6.8, color=NAVY, bold=True, align='c')
        text(s, x - 0.1, y_base + 0.04, bw + 0.2, 0.16, MON[int(C['months'][i][5:]) - 1], size=6.4, color=MUTED, align='c')
    # year brackets under the months
    i26 = next(i for i, m in enumerate(C['months']) if m.startswith('2026'))
    x_split = cx + i26 * (bw + gap) - gap / 2
    rect(s, cx, y_base + 0.26, x_split - cx - 0.04, 0.26, fill=NAVY)
    text(s, cx, y_base + 0.26, x_split - cx - 0.04, 0.26, '2025 · %s т' % n1(y25['tons']), size=7.8, color=WHITE, bold=True, align='c', anchor='m')
    rect(s, x_split + 0.04, y_base + 0.26, cx + cw - x_split - 0.04, 0.26, fill=AMBER)
    text(s, x_split + 0.04, y_base + 0.26, cx + cw - x_split - 0.04, 0.26, '2026 · %s т' % n1(y26['tons']), size=7.8, color=WHITE, bold=True, align='c', anchor='m')
    section_tag(s, 8.95, 3.3, 4, 'Країна походження', color=AMBER)
    orig = [o for o in sw['origins'] if o['share'] > 0.002]
    omx = max(o['tons'] for o in orig)
    for i, o in enumerate(orig):
        yy = 3.7 + i * 0.78
        text(s, 8.95, yy, 1.0, 0.3, cname(o['country']), size=9.4, color=NAVY, bold=True, anchor='m')
        bwid = 2.2 * o['tons'] / omx
        rect(s, 9.95, yy + 0.04, max(0.04, bwid), 0.26, fill=AMBER)
        text(s, 9.95 + bwid + 0.08, yy, 1.0, 0.34, '%s т' % n1(o['tons']), size=9.4, color=NAVY, bold=True, anchor='m')
        text(s, 8.95, yy + 0.36, 3.8, 0.34, '%d %% · 2025: %s т · 2026: %s т · $%s/кг' % (round(100 * o['share']), n1(o['t25']), n1(o['t26']), n1(o['usd_kg'])), size=7.6, color=MUTED)
    cnote(s)
    return s


# ---------------------------------------------------------------- customs 3
def customs_importers(deck):
    sw = C['seaweed']
    s = deck.slide('Митна база · хто завозить', 'Хто завозить нори-снеки в Україну')
    text(s, 0.62, 1.36, 12.1, 0.4, 'Імпортери нори-снеків (код 2008 99 99 90) за 17 місяців; бренди — ті, що названо в описах декларацій. Це ті самі бренди, що ми бачили на полицях мереж.', size=8.8, color=MUTED)
    caps(s, 0.62, 1.85, 6, 'Імпортер', size=6.6); caps(s, 3.9, 1.85, 3, 'Тонн за 17 місяців і частка', size=6.6)
    caps(s, 7.0, 1.85, 1.0, '2025, т', size=6.6, align='r'); caps(s, 8.1, 1.85, 1.3, '2026 січ–трав., т', size=6.6, align='r')
    caps(s, 9.75, 1.85, 1.0, 'Митна ціна', size=6.6); caps(s, 10.85, 1.85, 1.8, 'Бренди', size=6.6)
    imps = [i for i in sw['importers'] if i['tons'] > 0.3][:8]
    mx = max(i['tons'] for i in imps)
    rh = 0.42
    for k, i in enumerate(imps):
        y = 2.12 + k * rh
        if k % 2 == 0: rect(s, 0.62, y, 12.1, rh, fill=rgb('F8F9FC'))
        text(s, 0.75, y, 3.1, rh, i['name'].replace('ТОВ ', ''), size=8.2, color=NAVY, bold=True, anchor='m')
        bw = 1.9 * i['tons'] / mx
        rect(s, 3.9, y + 0.11, max(0.04, bw), rh - 0.22, fill=NAVY)
        text(s, 3.9 + bw + 0.08, y, 1.7, rh, '%s т · %d %%' % (n1(i['tons']), round(100 * i['share'])), size=7.8, color=MUTED, bold=True, anchor='m')
        text(s, 7.0, y, 1.0, rh, n1(i['t25']), size=8.4, color=INK, bold=True, anchor='m', align='r')
        text(s, 8.1, y, 1.3, rh, n1(i['t26']), size=8.4, color=AMBER, bold=True, anchor='m', align='r')
        text(s, 9.75, y, 1.0, rh, '$%s/кг' % n1(i['usd_kg']), size=8.6, color=INK, bold=True, anchor='m')
        br = ', '.join(i['brands']) if i['brands'] else ('Norris' if 'НОРРІС' in i['name'] else '—')
        text(s, 10.85, y, 1.9, rh, [[(br, {'bold': True, 'color': INK, 'size': 7.4})]], anchor='m')
    y = 2.12 + len(imps) * rh + 0.2
    section_tag(s, 0.62, y, 6, 'Виробники за описами декларацій (Корея, Китай, Таїланд)', color=AMBER)
    mf = [m for m in sw['manufacturers'] if not m['name'].startswith('(')][:6]
    x = 0.62
    for m in mf:
        w = 1.95
        rect(s, x, y + 0.3, w - 0.08, 0.62, fill=PANEL)
        text(s, x + 0.1, y + 0.34, w - 0.25, 0.3, m['name'].title().replace('Co Ltd', '').replace(' Kr', ''), size=7.4, color=INK, bold=True)
        text(s, x + 0.1, y + 0.64, w - 0.25, 0.25, '%s т · $%s/кг · %s' % (n1(m['tons']), n1(m['usd_kg']), cname(m['country'])), size=6.8, color=MUTED)
        x += w
    cnote(s)
    return s


# ---------------------------------------------------------------- customs 4
def customs_rice(deck):
    r = C['rice']
    s = deck.slide('Митна база · рисові снеки', 'Рисові крекери з норі: імпорту майже немає')
    text(s, 0.62, 1.36, 12.1, 0.4, 'Код 1905 90 55 00 — екструдовані й експандовані снеки; під ним оформлюють і наші Singha Kameda, Thai-Nichi та ZEK Tempura. Дивимось, хто і що завозить за цим кодом.', size=8.8, color=MUTED)
    kpi(s, 0.62, 1.9, 2.9, 1.2, 'Увесь код', '%s т' % n0(r['code_total_tons']), '2025: %s т · 2026 (5 міс.): %s т' % (n0(r['y25']), n0(r['y26'])))
    kpi(s, 3.68, 1.9, 2.9, 1.2, 'Рисові снеки', '%s т' % n1(r['rice_all']['tons']), '2025: %s т · 2026: %s т · $%s/кг' % (n1(r['rice25']), n1(r['rice26']), n1(r['rice_all']['usd_kg'])), color=AMBER)
    kpi(s, 6.74, 1.9, 2.9, 1.2, 'Рисові крекери в описах', '%s т' % n1(r['rice_cracker']['tons'], 2), '%d декларацій з 2 509 у коді' % r['rice_cracker']['decl'], color=RED)
    kpi(s, 9.8, 1.9, 2.9, 1.2, 'З водоростями (норі)', '%s т' % n1(r['nori_anything']['tons']), 'будь-які снеки з водоростями в коді', color=PURPLE)
    section_tag(s, 0.62, 3.3, 6, 'Хто завозить рисові снеки · тонн', color=AMBER)
    hbars(s, 0.62, 3.65, 5.9, [(i['name'].replace('ТОВ ', ''), i['tons']) for i in r['importers'][:6]], color=AMBER, row_h=0.4, name_w=2.2, val_w=0.6, fmt=lambda v: '%s т' % n1(v), size=8.4)
    section_tag(s, 6.9, 3.3, 6, 'Виробники · країна', color=NAVY)
    y = 3.65
    for m in r['manufacturers'][:6]:
        if m['name'].upper().startswith('MAFIN SRL ТОРГ'): continue
        nm = m['name'] if not m['name'].startswith('(') else 'Виробника в описі не вказано'
        rect(s, 6.9, y, 5.8, 0.36, fill=PANEL)
        text(s, 7.0, y, 3.5, 0.36, nm.title() if nm.isupper() else nm, size=8, color=INK, bold=True, anchor='m')
        text(s, 10.3, y, 2.3, 0.36, '%s т · %s' % (n1(m['tons']), cname(m['country'])), size=7.8, color=MUTED, anchor='m', align='r')
        y += 0.42
    rect(s, 0.62, 6.35, 12.1, 0.62, fill=rgb('FFF3DC')); rect(s, 0.62, 6.35, 0.04, 0.62, fill=GOLD)
    text(s, 0.85, 6.4, 11.7, 0.55, 'Висновок: рисових крекерів у стилі арарe/норімакі в цьому коді практично немає — це вільна ніша. Застереження: в мережах є Want Want mini, а в базі його не видно — його, ймовірно, ввозять за іншим кодом УКТ ЗЕД, якого немає у вашій вибірці.', size=8.4, color=INK)
    cnote(s)
    return s


# ---------------------------------------------------------------- customs 5
def customs_price(deck):
    sw = C['seaweed']
    s = deck.slide('Митна база · ціна закупівлі', 'Митна вартість проти нашого FOB, $ за кг')
    text(s, 0.62, 1.36, 12.1, 0.4, 'Скільки імпортери заявляють за кілограм нори-снеків (фактурна вартість) і скільки коштує кілограм у наших постачальників за FOB. Чим нижче наш FOB — тим більше запасу на полиці.', size=8.8, color=MUTED)
    mk = [('Корея · усі імпортери', [o for o in sw['origins'] if o['country'].startswith('КОРЕЯ')][0]['usd_kg'], NAVY),
          ('Таїланд · Tao Kae Noi', [o for o in sw['origins'] if o['country'].startswith('ТАЇ')][0]['usd_kg'], NAVY),
          ('Китай · Norris Group', [o for o in sw['origins'] if o['country'].startswith('КИТАЙ')][0]['usd_kg'], NAVY)]
    for m in sw['manufacturers']:
        if m['name'].startswith(('GODB', 'HAESA', 'SOLMOI F C CO LTD KR', 'HYOSUNG')):
            mk.append(('%s' % m['name'].title().replace('Co Ltd', '').replace(' Kr', ''), m['usd_kg'], tint(NAVY, 0.6)))
    seen = {}
    for o in OUR:
        key = (o['sup'], o['grams'], round(o['fob'], 3))
        if key in seen: continue
        seen[key] = o
    ours = []
    for (sid, g, f), o in seen.items():
        nm = O_name(o)
        ours.append(('%s %s' % (nm, gfmt(g)), f / g * 1000, SUP_COLOR[sid]))
    ours.sort(key=lambda r: -r[1])
    mxv = 55
    section_tag(s, 0.62, 1.85, 6, 'Ринок: митна вартість $/кг', color=NAVY)
    hbars(s, 0.62, 2.2, 5.9, mk, color=NAVY, row_h=0.46, name_w=2.3, val_w=0.7, maxv=mxv, fmt=lambda v: '$%s' % n1(v), size=8.4)
    section_tag(s, 6.9, 1.85, 6, 'Наші постачальники: FOB $/кг', color=AMBER)
    hbars(s, 6.9, 2.2, 5.8, ours, color=AMBER, row_h=0.3, name_w=2.5, val_w=0.7, maxv=mxv, fmt=lambda v: '$%s' % n1(v), size=7.6)
    ry = 2.2 + max(len(mk) * 0.46, len(ours) * 0.3) + 0.12
    rect(s, 0.62, ry, 12.1, 0.72, fill=rgb('FFF3DC')); rect(s, 0.62, ry, 0.04, 0.72, fill=GOLD)
    text(s, 0.85, ry + 0.06, 11.7, 0.64, [[('Як читати. ', {'bold': True, 'color': NAVY}), ('Корейські нори-снеки ввозять по $25–50/кг (в середньому $%s/кг), китайські — по $%s/кг. TMK KOKIRI (норі з Кореї) — $37–49/кг, на рівні корейського ринку. ZEK, Singha Kameda і Thai-Nichi — $11–18/кг: це рисові крекери, темпура й норі з начинкою — інші продукти, дешевші за кілограм.' % (n1(mk[0][1], 1), n1(mk[2][1], 1)), {})]], size=8.2, color=INK)
    cnote(s)
    return s


def O_name(o):
    import re
    t = o['title'].replace('Seaweed Topping', 'Topping').replace('Tempura Seaweed', 'Tempura').replace('Sandwich Seaweed', 'Sandwich').replace('Arare Norimaki', 'Singha Norimaki').replace('Norimaki', 'Thai-Nichi Norimaki') if o['sup'] == 'thainichi' else o['title'].replace('Seaweed Topping', 'Topping').replace('Tempura Seaweed', 'Tempura').replace('Sandwich Seaweed', 'Sandwich')
    t = re.sub(r'\s*·\s*\d+\s*г$', '', t)
    for flv in (' Original', ' Corn', ' Chicken Floss', ' Vegetables', ' Sesame', ' Meat Floss', ' Hot Spicy', ' Spicy Squid', ' Squid', ' Spicy', ' Wasabi', ' Mala Crawfish', ' Mala', ' Chicken'):
        t = t.replace(flv, '')
    return t.replace('Wow ', 'WOW ')


# ================================================================ products
def glossary(deck):
    s = deck.slide('Що це за продукти', 'Формати категорії: що саме лежить на полиці')
    text(s, 0.62, 1.36, 12.1, 0.4, 'Опис — за митними описами товарів (графа 31 декларації) та написами на упаковках. Фото — продукти, які ми знайшли в мережах.', size=8.8, color=MUTED)
    def pic(sg, brand, g):
        l = next(x for x in lines(rows_of(sg)) if x['brand'] == brand and x['grams'] == g)
        return os.path.join(HERE, l['img']), l
    cards = [
        ('mini', 'Норі-снек у міні-пакеті', 'Тонкі листи морської водорості норі, підсмажені з олією, сіллю й спеціями; у пакетику 1–2 аркуші по 4–5 г. Їдять як хрусткі чипси або загортають у них рис.', 'Haelove', 4.5, 'Тут: TMK ROLL і DOUBLE ROLL (рулончик 2,5 і 5 г)'),
        ('chips', 'Норі-чипси та сендвічі', 'Хрусткі квадрати норі в пакеті 25 г (Hokkaido Club, Norris) і «сендвіч» — два листи норі, склеєні начинкою з рису, кунжуту, насіння чи м’ясної стружки.', 'Hokkaido Club', 25.0, 'Тут: ZEK Sandwich 25 г і TMK SEAWEED / MINI 10–12 г'),
        ('tempura', 'Темпура з норі', 'Норі в хрусткому клярі темпура — як чипси, але товстіші, з гострими й васабі-смаками. Одного бренду в мережах — Tao Kae Noi. У ZEK — рисовий снек із сушеною водорістю.', 'Tao Kae Noi', 25.0, 'Тут: ZEK Tempura 30 і 50 г'),
        ('tempura', 'Топінг із норі', 'Подрібнена водорість із олією, цукром і начинкою (кунжут, овочі, курячa стружка). Не снек, а посипка для рису, локшини, супів і салатів.', 'Ock Dong Ja', 70.0, 'Тут: ZEK Topping 35 і 70 г'),
        ('ricecr', 'Рисовий крекер із норі (арарe / норімакі)', 'Обсмажений рисовий крекер із соєвим соусом, обгорнутий або присипаний смаженою норі. Основа — клейкий рис (близько 75–80 %).', 'Want Want', 60.0, 'Тут: Singha Kameda Norimaki 42 г, Thai-Nichi Mizuho Norimaki 50–55 г'),
        ('ricechip', 'Рисові чипси', 'Тонкі хрусткі рисові чипси зі смаком сиру, паприки чи сметани. Суміжний формат, не наш: ціна 40–70 ₴ за пакет.', 'Rice Up!', 60.0, 'У нашому асортименті аналогів немає'),
    ]
    for i, (sg, name, desc, br, g, ours) in enumerate(cards):
        x = 0.62 + (i % 3) * 4.1; y = 1.95 + (i // 3) * 2.62
        col = SEG_COLOR[sg]
        p, l = pic(sg, br, g)
        rect(s, x, y, 3.95, 2.5, fill=PANEL); rect(s, x, y, 3.95, 0.05, fill=col)
        rect(s, x + 0.12, y + 0.2, 1.3, 1.5, fill=WHITE)
        picture(s, p, x + 0.16, y + 0.24, 1.22, 1.42)
        text(s, x + 1.55, y + 0.18, 2.3, 0.5, name, size=9.4, color=mixc(col, INK, 0.75), bold=True)
        text(s, x + 1.55, y + 0.72, 2.3, 1.0, desc, size=7.2, color=INK)
        text(s, x + 0.14, y + 1.78, 3.7, 0.2, '%s · %s · типова ціна %d ₴' % (l['brand'], gfmt(l['grams']), round(l['pmed'])), size=7, color=MUTED, bold=True)
        rect(s, x + 0.14, y + 2.0, 3.67, 0.4, fill=WHITE)
        text(s, x + 0.22, y + 2.0, 3.55, 0.4, ours, size=7.6, color=NAVY, bold=True, anchor='m')
    footnote(s, 'Джерела: митні описи товарів (графа 31 МД) постачальників; каталоги мереж zakaz.ua, 07.10.2026.')
    return s


def passports(deck):
    s = deck.slide('Що це за продукти · наш асортимент', 'Паспорт продукту: що саме ми пропонуємо')
    text(s, 0.62, 1.36, 12.1, 0.4, 'За вашою таблицею УКТ ЗЕД: класифікація, мито та склад продуктів кожного постачальника. Мито — пільгове / повне.', size=8.8, color=MUTED)
    sups = [
        ('singha', 'Singha Kameda · Arare Norimaki', 'sk_original', '1905 90 55 00', '10 % / 10 %',
         'Рисовий крекер, обсмажений / екструдований: клейкий рис 75–80 %, соєвий соус (у Original — соєво-рибний з екстрактом боніто), смажена норі 3,5–4 %. Роздріб — пакет-стойка 42 г; також навал 3 кг.',
         'Таїланд', 'Original · Wasabi (у таблиці ще Squid Chili, Mala)'),
        ('thainichi', 'Thai-Nichi · Mizuho Norimaki', 'tn_original', '1905 90 55 00', '10 % / 10 %',
         'Рисовий крекер, обсмажений / екструдований, із соєвим соусом і сушеною смаженою норі. Алюмінієвий пакет з азотним наповненням; 50 і 55 г.',
         'Таїланд', 'Original · Wasabi'),
        ('tmk', 'TMK · KOKIRI WOW', 'tmk_roll_orig', '2008 99 99 90', '10 % / 15 %',
         'Смажена норі (корейська, 75–95 %) зі смаковою приправою 5 %. ROLL — рулончик 2,5 г; DOUBLE ROLL — подвійний 2 × 2,5 г; MINI — дрібні рулончики 10 г; SEAWEED — хрустка норі, смажена в пальмовій олії, 12 г.',
         'Таїланд (сировина з Кореї)', 'Original · Hot Spicy · Squid'),
        ('zek', 'HanJin · ZEK', 'zek_tempura_corn30', '1905 90 55 00 · 2008 99 99 90', '10 % / 10 % · 10 % / 15 %',
         'Tempura 30 / 50 г — рисовий снек із сушеною норі, крохмалем і приправою (1905). Sandwich 25 г — листи норі з начинкою (мальтоза, м’ясна стружка / кунжут). Topping 35 / 70 г — подрібнена норі з олією, цукром і начинкою (2008).',
         'Китай', 'Corn · Wasabi · Mala · Meat Floss · Sesame · Vegetables · Chicken'),
    ]
    for i, (sid, name, ph, code, duty, desc, ctry, flv) in enumerate(sups):
        x = 0.62 + i * 3.06; w = 2.95
        col = SUP_COLOR[sid]
        rect(s, x, 1.9, w, 0.4, fill=col)
        text(s, x + 0.12, 1.99, w - 0.2, 0.24, name.upper(), size=7.4, color=WHITE, bold=True, spc=0.15)
        rect(s, x, 2.3, w, 4.75, fill=PANEL)
        pth = D.sup_photo(ph)
        rect(s, x + 0.12, 2.42, w - 0.24, 1.35, fill=WHITE)
        if os.path.exists(pth): picture(s, pth, x + 0.2, 2.46, w - 0.4, 1.27)
        yy = 3.9
        for lab, val in [('ЩО ЦЕ', desc), ('КРАЇНА', ctry), ('СМАКИ', flv), ('УКТ ЗЕД', code), ('МИТО', duty)]:
            text(s, x + 0.14, yy, w - 0.28, 0.16, lab, size=6.2, color=col, bold=True, spc=0.3)
            hh = 1.0 if lab == 'ЩО ЦЕ' else 0.34
            text(s, x + 0.14, yy + 0.17, w - 0.28, hh, val, size=7.0 if lab == 'ЩО ЦЕ' else 7.6, color=INK, bold=(lab in ('УКТ ЗЕД', 'МИТО')))
            yy += 1.15 if lab == 'ЩО ЦЕ' else 0.46
    footnote(s, 'Джерело: таблиця УКТ ЗЕД постачальників (опис до графи 31 МД). Класифікацію підтверджує митний брокер; мито — за вашою таблицею.')
    return s
