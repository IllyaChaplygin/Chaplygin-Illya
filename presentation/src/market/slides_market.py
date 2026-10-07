# -*- coding: utf-8 -*-
"""Market-study slides (segmentation, catalogues, where sold, price maps)."""
import collections, os
from md import *  # noqa
from pptx.util import Inches, Pt
from md import K, rect, text, caps, picture, kpi, section_tag, hbars, columns, footnote, mixc
import mdata as D
from mdata import SEGS, rows_of, lines, med, gfmt, uah, rng, CHAIN, NATIONAL, ROWS, OUR

HERE = os.path.dirname(os.path.abspath(__file__))
ALLN = [r for r in ROWS]                  # every position in the study
NORI = [r for r in ROWS if r['seg'] in ('mini', 'chips', 'tempura')]
CHAINS_ORDER = sorted({c for r in ROWS for c in r['chains']},
                      key=lambda c: (-(c in NATIONAL), -sum(1 for r in ROWS if c in r['chains']), CHAIN[c]))
SRC = ('Каталоги мереж на zakaz.ua (METRO, NOVUS, Ашан, Зараз, МегаМаркет, Ultramarket, Епіцентр, Космос, Onde, Ідеал, Восторг, Клас, '
       'Таврія В, ЧудоМаркет, Торба, Grono) · %s · ціна — онлайн-полиця одного магазину мережі, діапазон — між мережами.')


def img(r):
    return os.path.join(HERE, r['img'])


def sub(s, t, y=1.36, size=8.8):
    text(s, 0.62, y, 12.1, 0.4, t, size=size, color=MUTED)


def src_note(s):
    d = D.DATE.split('-'); ds = '%s.%s.%s' % (d[2], d[1], d[0])
    footnote(s, SRC % ds)


# ------------------------------------------------------------------- cover
def cover(deck):
    s = deck.prs.slides.add_slide(deck.prs.slide_layouts[6])
    deck.page += 1
    rect(s, 0, 0, 13.3333, 7.5, fill=rgb('FFF3DC'))
    rect(s, 8.6, 0, 4.7333, 7.5, fill=GOLD)
    rect(s, 7.6, 3.9, 3.6, 3.6, fill=rgb('FFB82E'), radius=0.5)
    s.shapes.add_shape  # keep linter quiet
    rect(s, 0.8, 0.8, 2.2, 0.98, fill=NAVY, radius=0.12)
    s.shapes.add_picture(os.path.join(HERE, 'assets', 'logo.png'), Inches(1.0 * K), Inches(0.9 * K), Inches(1.8 * K), Inches(0.78 * K))
    text(s, 0.8, 3.0, 7.5, 0.25, 'ПОСТАЧАЛЬНИКИ · ЗРІЗ РИНКУ · ЦІНА НА ПОЛИЦІ · ЖОВТЕНЬ 2026', size=9.5, color=AMBER, bold=True, spc=0.3)
    text(s, 0.8, 3.4, 7.6, 2.4, 'Норі-снеки\nта рисові крекери\nв Україні', size=38, color=NAVY, bold=True, ls=1.0)
    rect(s, 0.85, 6.0, 1.4, 0.07, fill=GOLD)
    pics = [('mini', 0), ('chips', 0), ('tempura', 1), ('ricecr', 0)]
    picks = []
    for sg, _ in pics:
        ls = lines(rows_of(sg))
        picks.append(ls[len(ls) // 2]['img'] if sg != 'ricecr' else [l for l in ls if l['brand'] == 'Want Want'][0]['img'])
    pos = [(9.2, 0.45), (11.0, 1.6), (9.2, 3.1), (11.0, 4.5)]
    for (x, y), p in zip(pos, picks):
        rect(s, x + 0.06, y + 0.08, 1.9, 2.3, fill=rgb('E08F00'), radius=0.08)
        rect(s, x, y, 1.9, 2.3, fill=WHITE, radius=0.08)
        picture(s, os.path.join(HERE, p), x + 0.12, y + 0.12, 1.66, 2.06)
    return s


# ------------------------------------------------------------- methodology
def method(deck):
    s = deck.slide('Методологія · джерела', 'Що перевірили і як рахували')
    n_chain = len(CHAINS_ORDER)
    kpi(s, 0.62, 1.55, 2.9, 1.2, 'Позицій у дослідженні', '%d SKU' % len(ROWS), 'унікальні штрихкоди (EAN)')
    kpi(s, 3.68, 1.55, 2.9, 1.2, 'Мереж із позиціями', '%d з 18' % n_chain, '7 національних · %d регіональних' % (n_chain - len([c for c in CHAINS_ORDER if c in NATIONAL])))
    kpi(s, 6.74, 1.55, 2.9, 1.2, 'Брендів', '%d' % len({r['brand'] for r in ROWS}), 'без обліку приватних марок мереж', color=AMBER)
    kpi(s, 9.8, 1.55, 2.9, 1.2, 'Дата зрізу', '07.10.2026', 'онлайн-каталоги мереж', color=PURPLE)
    section_tag(s, 0.62, 3.0, 5, 'Як рахували')
    pts = ['Ціна — поточна ціна на онлайн-полиці мережі (zakaz.ua), грн за упаковку; діапазон показує різницю між мережами та акціями.',
           'Позиція = унікальний штрихкод. Лінійка = бренд + вага; смаки однієї лінійки згруповано в одну картку.',
           'Сегменти — за форматом упаковки й вагою (те, що бачить покупець на полиці), так само, як у дослідженнях рису і рамену.',
           'Наші SKU порівнюємо з сегментом за ціною упаковки та ціною за грам; коридор «поруч» — ±20 % від нашої полиці.']
    for i, t in enumerate(pts):
        rect(s, 0.62, 3.4 + i * 0.78, 0.34, 0.34, fill=NAVY, radius=0.5)
        text(s, 0.62, 3.4 + i * 0.78 + 0.05, 0.34, 0.25, str(i + 1), size=10, color=WHITE, bold=True, align='c')
        text(s, 1.1, 3.4 + i * 0.78, 5.2, 0.7, t, size=8.6, color=INK)
    section_tag(s, 6.9, 3.0, 5, 'Мережі в зрізі', color=AMBER)
    nat = [c for c in CHAINS_ORDER if c in NATIONAL]; reg = [c for c in CHAINS_ORDER if c not in NATIONAL]
    def chips(y, label, cs, col):
        caps(s, 6.9, y, 5, label, size=7)
        x, yy = 6.9, y + 0.28
        for c in cs:
            w = 0.2 + 0.085 * len(CHAIN[c])
            if x + w > 12.7: x, yy = 6.9, yy + 0.34
            rect(s, x, yy, w, 0.27, fill=tint(col, 0.14), radius=0.3)
            text(s, x, yy + 0.02, w, 0.24, CHAIN[c], size=7.8, color=INK, bold=True, align='c', anchor='m')
            x += w + 0.1
        return yy + 0.4
    y = chips(3.4, 'Національне покриття', nat, AMBER)
    y = chips(y + 0.05, 'Регіональні', reg, NAVY)
    rect(s, 6.9, y + 0.15, 5.8, 1.5, fill=rgb('FFF3DC'))
    rect(s, 6.9, y + 0.15, 0.04, 1.5, fill=GOLD)
    text(s, 7.15, y + 0.27, 5.4, 1.3, [[('Що ще в зрізі. ', {'bold': True, 'color': NAVY}),
         ('Митна база (УКТ ЗЕД, 2025 + січень–травень 2026) — обсяги, імпортери, країни й митна ціна. ', {}),
         ('Чого немає: ', {'bold': True, 'color': NAVY}),
         ('офлайн-акцій «в залі»; товарів під іншими кодами УКТ ЗЕД. Арарe/норімакі в мережах не знайдено — «не знайдено в каталогах» не означає «не продається».', {})]], size=8.2, color=INK)
    src_note(s)
    return s


# ------------------------------------------------------------ segmentation
def top_pack(sg):
    t = dict(pack_table([sg]))[sg][0]
    return '%s · %d %% присутності' % (gfmt(t['grams']), round(100 * t['share']))


def segmentation(deck):
    s = deck.slide('Сегментація', 'Сегментація за форматом упаковки')
    sub(s, 'Усі знайдені в мережах позиції розкладено за форматом упаковки й вагою — так, як їх бачить покупець на полиці. У кожній картці: скільки позицій, яка упаковка найходовіша, типова ціна і діапазон.')
    total = len(ROWS)
    order = list(SEGS)
    cw, ch = 3.9, 2.5
    for i, sg in enumerate(order):
        rs = rows_of(sg); ls = lines(rs)
        col = SEG_COLOR[sg]
        x = 0.62 + (i % 3) * 4.1; y = 1.85 + (i // 3) * 2.62
        rect(s, x, y, cw, 0.36, fill=col)
        text(s, x + 0.14, y + 0.07, cw - 0.2, 0.22, SEGS[sg]['name'].upper(), size=8.0, color=WHITE, bold=True, spc=0.2)
        rect(s, x + 0.04, y + 0.4, cw - 0.08, ch - 0.4, fill=PANEL)
        pic = max(ls, key=lambda l: l['n'])['img']
        rect(s, x + 0.14, y + 0.5, 1.1, 1.1, fill=WHITE)
        picture(s, os.path.join(HERE, pic), x + 0.19, y + 0.54, 1.0, 1.02)
        text(s, x + 1.4, y + 0.44, 1.3, 0.5, str(len(rs)), size=26, color=col, bold=True)
        text(s, x + 1.4, y + 0.98, 2.4, 0.2, 'позицій · %d %% зрізу' % round(100 * len(rs) / total), size=8.2, color=MUTED)
        text(s, x + 1.4, y + 1.2, 2.45, 0.4, SEGS[sg]['desc'], size=7.4, color=MUTED)
        facts = [('ВАГА', '%s–%s' % (('%g' % min(r['grams'] for r in rs)).replace('.', ','), gfmt(max(r['grams'] for r in rs)))),
                 ('МЕДІАНА ЦІНИ ЗА УПАКОВКУ', uah(med(r['pmed'] for r in rs))),
                 ('НАЙХОДОВІША УПАКОВКА', top_pack(sg)),
                 ('ДІАПАЗОН · БРЕНДІВ', '%s ₴ · %d' % (rng(min(r['pmin'] for r in rs), max(r['pmax'] for r in rs)), len({r['brand'] for r in rs})))]
        for k, (a, b) in enumerate(facts):
            yy = y + 1.68 + k * 0.215
            text(s, x + 0.14, yy + 0.02, 2.3, 0.16, a, size=6.4, color=MUTED, bold=True)
            text(s, x + 2.0, yy, 1.8, 0.2, b, size=8.6, color=NAVY, bold=True, align='r')
    # last tile: reading guide + heat strip of ours
    x, y = 0.62 + 2 * 4.1, 1.85 + 2.62
    rect(s, x, y, cw, ch, fill=rgb('FFF3DC'))
    rect(s, x, y, 0.04, ch, fill=GOLD)
    caps(s, x + 0.25, y + 0.18, 3.4, 'Де наші постачальники', size=7.4, color=AMBER)
    mapping = [('TMK · KOKIRI', 'tmk', 'міні-пакет · чипси'), ('HanJin · ZEK', 'zek', 'темпура · топінг · сендвіч'),
               ('Singha Kameda', 'singha', 'рисові крекери'), ('Thai-Nichi', 'thainichi', 'рисові крекери')]
    for k, (nm, sid, where) in enumerate(mapping):
        yy = y + 0.55 + k * 0.48
        rect(s, x + 0.25, yy, 0.2, 0.34, fill=SUP_COLOR[sid])
        text(s, x + 0.58, yy - 0.01, 3.0, 0.2, nm, size=9, color=INK, bold=True)
        text(s, x + 0.58, yy + 0.19, 3.0, 0.18, where, size=7.6, color=MUTED)
    src_note(s)
    return s


# --------------------------------------------------------- segment overview
def seg_overview(deck, sg, headline, note=None, bins=None, rep_n=6):
    rs = rows_of(sg); ls = lines(rs); col = SEG_COLOR[sg]
    s = deck.slide('Сегмент · %s · %d позицій' % (SEGS[sg]['name'], len(rs)), headline)
    if note:
        text(s, 0.62, 1.38, 12.1, 0.22, note, size=8.8, color=MUTED)
    mp = med(r['pmed'] for r in rs); mg = med(r['per_g'] for r in rs)
    kpi(s, 0.62, 1.74, 3.1, 1.18, 'Позицій у сегменті', '%d SKU' % len(rs), '%d лінійок · %d брендів' % (len(ls), len({r['brand'] for r in rs})), color=col)
    kpi(s, 3.88, 1.74, 3.1, 1.18, 'Медіана упаковки', uah(mp), 'діапазон %s ₴' % rng(min(r['pmin'] for r in rs), max(r['pmax'] for r in rs)), color=col)
    tp = dict(pack_table([sg]))[sg][0]
    kpi(s, 7.14, 1.74, 3.1, 1.18, 'Найходовіша упаковка', gfmt(tp['grams']), '%d %% присутності на полицях · медіана %d ₴' % (round(100 * tp['share']), round(tp['pmed'])), color=col)
    kpi(s, 10.4, 1.74, 2.32, 1.18, 'Вага', '%s–%s' % (('%g' % min(r['grams'] for r in rs)).replace('.', ','), gfmt(max(r['grams'] for r in rs))), None, color=col, vsize=15)
    # who holds the segment
    cnt = collections.Counter(r['brand'] for r in rs).most_common(8)
    section_tag(s, 0.62, 3.1, 5, 'Хто тримає сегмент · SKU на бренд', color=col)
    hbars(s, 0.62, 3.42, 5.6, [(b, n) for b, n in cnt], color=col, row_h=0.235, name_w=1.5)
    # price distribution
    section_tag(s, 6.9, 3.1, 5, 'Позиції за ціною упаковки · ₴', color=col)
    vals = [r['pmed'] for r in rs]
    bins = bins or [(0, 40), (40, 55), (55, 80), (80, 110), (110, 150), (150, 1000)]
    data = []
    for a, b in bins:
        lab = ('до %d' % b) if a == 0 else (('%d+' % a) if b >= 1000 else '%d–%d' % (a, b))
        data.append((lab, sum(1 for v in vals if a <= v < b)))
    columns(s, 6.9, 3.38, 5.8, 1.85, data, color=col)
    # representatives
    section_tag(s, 0.62, 5.35, 8, 'Представники сегмента · від найдешевшої до найдорожчої', color=col)
    step = max(1, len(ls) // rep_n)
    reps = [ls[min(len(ls) - 1, i * step)] for i in range(rep_n)] if len(ls) >= rep_n else ls
    w = 1.92
    for i, l in enumerate(reps):
        x = 0.62 + i * (w + 0.12)
        rect(s, x, 5.68, w, 1.62, fill=PANEL); rect(s, x, 5.68, w, 0.04, fill=col)
        rect(s, x + 0.07, 5.78, 0.8, 0.8, fill=WHITE)
        picture(s, os.path.join(HERE, l['img']), x + 0.09, 5.8, 0.76, 0.76)
        text(s, x + 0.94, 5.82, 0.95, 0.5, l['brand'], size=7.8, color=INK, bold=True)
        text(s, x + 0.94, 6.3, 0.95, 0.3, uah(l['pmed']), size=11, color=INK, bold=True)
        text(s, x + 0.1, 6.68, w - 0.15, 0.2, '%s · %d %s' % (gfmt(l['grams']), l['n'], 'смак' if l['n'] == 1 else ('смаки' if l['n'] < 5 else 'смаків')), size=7, color=MUTED)
        text(s, x + 0.1, 6.9, w - 0.15, 0.3, ' · '.join(CHAIN[c] for c in l['chains'][:3]) + ('' if len(l['chains']) <= 3 else ' +%d' % (len(l['chains']) - 3)), size=6.8, color=MUTED)
    return s


# ------------------------------------------------------------ catalogue
def line_card(s, x, y, w, h, l, col):
    rect(s, x, y, w, h, fill=PANEL); rect(s, x, y, w, 0.04, fill=col)
    rect(s, x + 0.08, y + 0.12, w - 0.16, 1.32, fill=WHITE)
    picture(s, os.path.join(HERE, l['img']), x + 0.14, y + 0.15, w - 0.28, 1.26)
    text(s, x + 0.1, y + 1.52, w - 0.2, 0.2, l['brand'], size=8.8, color=INK, bold=True)
    fl = ', '.join(f.lower() if i else f for i, f in enumerate(l['flavors'][:3]))
    if len(fl) > 52: fl = fl[:50].rsplit(',', 1)[0]
    if l['n'] > len(fl.split(',')): fl += ' +%d' % (l['n'] - len(fl.split(',')))
    text(s, x + 0.1, y + 1.72, w - 0.2, 0.34, fl, size=7, color=MUTED)
    text(s, x + 0.1, y + 2.07, w - 0.2, 0.28, uah(l['pmin']).replace(' ₴', '') + ('' if round(l['pmin']) == round(l['pmax']) else '–%d' % round(l['pmax'])) + ' ₴', size=13, color=INK, bold=True)
    text(s, x + 0.1, y + 2.36, w - 0.2, 0.18, '%s · %d %s' % (gfmt(l['grams']), l['n'], 'смак' if l['n'] == 1 else ('смаки' if l['n'] < 5 else 'смаків')), size=7.2, color=MUTED)
    text(s, x + 0.1, y + 2.54, w - 0.2, 0.2, ' · '.join(CHAIN[c] for c in l['chains'][:2]) + ('' if len(l['chains']) <= 2 else ' +%d' % (len(l['chains']) - 2)), size=6.8, color=MUTED)


def catalog(deck, items, eyebrow, title):
    """items: list of (line, colour) — 12 per slide, in 2 rows of 6."""
    s = deck.slide(eyebrow, title)
    sub(s, 'Картка = лінійка бренда (бренд + вага пакета). Ціна — діапазон по мережах, де лінійка знайдена; смаки з однаковою упаковкою згруповано, кількість смаків — під ціною.', y=1.34, size=8.4)
    w, h = 1.92, 2.78
    for i, (l, col) in enumerate(items):
        x = 0.62 + (i % 6) * (w + 0.12); y = 1.68 + (i // 6) * (h + 0.06)
        line_card(s, x, y, w, h, l, col)
    return s


# ------------------------------------------------------------- where sold
def where_sold(deck, sgs, headline, eyebrow):
    rs = [r for sg in sgs for r in rows_of(sg)]
    col = SEG_COLOR[sgs[0]]
    s = deck.slide(eyebrow, headline)
    cnt = collections.Counter(c for r in rs for c in r['chains'])
    sub(s, 'Скільки позицій цього сегмента стоїть у кожній мережі (ліворуч) і розклад мереж за типом: національні — головний цільовий канал, регіональні — вхід, без позицій — вільна полиця.', y=1.34, size=8.4)
    caps(s, 0.62, 1.95, 5, 'Позицій сегмента в мережі', size=7)
    order = sorted(CHAINS_ORDER, key=lambda c: -cnt.get(c, 0))
    rows = [(CHAIN[c], cnt[c]) for c in order if cnt.get(c)]
    hbars(s, 0.62, 2.25, 6.4, rows, color=col, row_h=min(0.4, 4.7 / len(rows)), name_w=1.6, val_w=0.5, maxv=max(cnt.values()), size=9.2)
    # right: chips by channel group
    x0 = 7.3
    def block(y, title, cs, tag, tagcol, fillc):
        rect(s, x0, y, 5.4, 0.3, fill=WHITE, line=None)
        text(s, x0, y + 0.02, 3.5, 0.26, title, size=9.4, color=NAVY, bold=True)
        rect(s, x0 + 3.5, y + 0.02, 1.9, 0.24, fill=tagcol)
        text(s, x0 + 3.5, y + 0.02, 1.9, 0.24, tag, size=6.4, color=WHITE, bold=True, align='c', anchor='m')
        xx, yy = x0, y + 0.42
        for c in cs:
            n = cnt.get(c, 0)
            lab = CHAIN[c]; w = 0.5 + 0.082 * len(lab) + (0.3 if n else 0)
            if xx + w > x0 + 5.4: xx, yy = x0, yy + 0.36
            rect(s, xx, yy, w, 0.28, fill=WHITE, line=RULE)
            text(s, xx + 0.08, yy + 0.02, w - 0.4, 0.24, lab, size=7.6, color=INK, bold=True, anchor='m')
            if n:
                rect(s, xx + w - 0.32, yy + 0.04, 0.26, 0.2, fill=fillc)
                text(s, xx + w - 0.32, yy + 0.04, 0.26, 0.2, str(n), size=7, color=WHITE, bold=True, align='c', anchor='m')
            xx += w + 0.1
        return yy + 0.5
    allc = [c for c in CHAINS_ORDER]
    nat = [c for c in D.CHAIN if c in NATIONAL and cnt.get(c)]; reg = [c for c in CHAINS_ORDER if c not in NATIONAL and cnt.get(c)]
    none = [c for c in CHAIN if not cnt.get(c)]
    y = 1.95
    y = block(y, 'Національні мережі · %d' % len(nat), sorted(nat, key=lambda c: -cnt[c]), 'ЦІЛЬОВИЙ КАНАЛ', RED, RED)
    y = block(y + 0.1, 'Регіональні мережі · %d' % len(reg), sorted(reg, key=lambda c: -cnt[c]), 'РЕГІОНАЛЬНИЙ ВХІД', TEAL, TEAL)
    if none:
        block(y + 0.1, 'Без позицій сегмента · %d' % len(none), none, 'ВІЛЬНА ПОЛИЦЯ', GREY, GREY)
    src_note(s)
    return s


# ------------------------------------------------------- popular pack sizes
def pack_table(sgs):
    """Per weight: how widely the pack size is stocked (SKU x chain placements — the
    only popularity signal available without sales data), and what it costs."""
    out = []
    for sg in sgs:
        rs = rows_of(sg)
        tot = sum(len(r['chains']) for r in rs)
        by = collections.OrderedDict()
        for r in rs:
            by.setdefault(round(r['grams'], 1), []).append(r)
        part = []
        for g, v in by.items():
            lst = sum(len(r['chains']) for r in v)
            part.append(dict(seg=sg, grams=g, listings=lst, share=lst / tot, n=len(v), brands=len({r['brand'] for r in v}),
                             chains=len({c for r in v for c in r['chains']}), pmin=min(r['pmin'] for r in v),
                             pmax=max(r['pmax'] for r in v), pmed=med(r['pmed'] for r in v), per_g=med(r['per_g'] for r in v)))
        part.sort(key=lambda p: -p['listings'])
        for i, p in enumerate(part):
            p['rank'] = i + 1
        out.append((sg, part))
    return out


def popular_packs(deck, sgs, headline, eyebrow, note=None):
    tabs = pack_table(sgs)
    s = deck.slide(eyebrow, headline)
    text(s, 0.62, 1.38, 12.1, 0.22, note or 'Популярність — це присутність упаковки на полицях: кількість позицій × мереж, де вона стоїть (даних про продажі немає). Справа — наша полиця проти типової ціни пакетів близької ваги.', size=8.6, color=MUTED)
    x0 = 0.62
    heads = [('№', 0, 0.35), ('ВАГА', 0.4, 0.9), ('ПРИСУТНІСТЬ · ПОЗИЦІЙ × МЕРЕЖ', 1.4, 3.1), ('SKU · БРЕНДІВ', 4.6, 1.0), ('МЕРЕЖ', 5.7, 0.6), ('ЦІНА, ₴ · МІН.–МАКС.', 6.4, 1.6), ('МЕДІАНА ЦІНИ', 8.0, 1.2)]
    ny = sum(len(p) for _, p in tabs)
    rh = min(0.62, 4.0 / (ny + 1.3 * len(tabs)))
    y = 1.85
    pos = {}
    for sg, part in tabs:
        col = SEG_COLOR[sg]
        rect(s, x0, y, 9.4, 0.3, fill=col)
        text(s, x0 + 0.12, y + 0.03, 4, 0.24, SEGS[sg]['name'].upper(), size=7.4, color=WHITE, bold=True, spc=0.2)
        for t, dx, w in heads:
            text(s, x0 + dx, y + 0.33, w + (0.2 if dx > 6 else 0.5), 0.16, t, size=5.6, color=MUTED, bold=True, align='r' if dx > 8 else 'l')
        y += 0.52
        mx = max(p['listings'] for p in part)
        for p in part:
            if p['rank'] % 2 == 1:
                rect(s, x0, y, 9.4, rh, fill=rgb('F8F9FC'))
            pos[(sg, p['grams'])] = (p, y)
            rect(s, x0 + 0.04, y + rh / 2 - 0.13, 0.26, 0.26, fill=col if p['rank'] == 1 else tint(col, 0.35), radius=0.5)
            text(s, x0 + 0.04, y + rh / 2 - 0.13, 0.26, 0.26, str(p['rank']), size=7.6, color=WHITE, bold=True, align='c', anchor='m')
            text(s, x0 + 0.4, y, 0.95, rh, gfmt(p['grams']), size=10.5, color=NAVY, bold=True, anchor='m')
            bw = 2.3 * p['listings'] / mx
            rect(s, x0 + 1.4, y + rh * 0.27, max(0.04, bw), rh * 0.46, fill=col if p['rank'] == 1 else tint(col, 0.55))
            text(s, x0 + 1.4 + bw + 0.08, y, 1.2, rh, '%d · %d %%' % (p['listings'], round(100 * p['share'])), size=7.8, color=MUTED, bold=True, anchor='m')
            text(s, x0 + 4.6, y, 1.0, rh, '%d · %d' % (p['n'], p['brands']), size=8, color=NAVY, anchor='m')
            text(s, x0 + 5.7, y, 0.6, rh, str(p['chains']), size=8, color=NAVY, anchor='m')
            text(s, x0 + 6.4, y, 1.6, rh, '%s ₴' % rng(p['pmin'], p['pmax']), size=8, color=MUTED, anchor='m')
            text(s, x0 + 8.0, y, 1.2, rh, '%d ₴' % round(p['pmed']), size=11, color=INK, bold=True, anchor='m', align='r')
            y += rh
        y += 0.12
    # column captions under the first band are folded into one caption line at the bottom of the table
    # what it means for our pack sizes
    if y < 6.3:
        notes = []
        seen_w = set()
        for o in OUR:
            if o['seg'] not in sgs or (o['seg'], o['grams']) in seen_w:
                continue
            seen_w.add((o['seg'], o['grams']))
            part = dict(tabs)[o['seg']]
            exact = [p for p in part if abs(p['grams'] - o['grams']) < 0.05]
            sup = {'tmk': 'TMK', 'zek': 'ZEK', 'singha': 'Singha', 'thainichi': 'Thai-Nichi'}[o['sup']]
            if exact:
                p = exact[0]
                notes.append('%s %s — такий розмір є на ринку: №%d, %d %% присутності, типова ціна %d ₴.' % (sup, gfmt(o['grams']), p['rank'], round(100 * p['share']), round(p['pmed'])))
            else:
                nr = sorted(part, key=lambda p: abs(p['grams'] - o['grams']))[:2]
                notes.append('%s %s — такого розміру на ринку немає; найближчі: %s.' % (sup, gfmt(o['grams']), ', '.join('%s (№%d, %d %%)' % (gfmt(p['grams']), p['rank'], round(100 * p['share'])) for p in nr)))
        notes = notes[:6]
        by = y + 0.1
        bh = 0.3 + 0.24 * len(notes)
        rect(s, x0, by, 9.4, bh, fill=rgb('FFF3DC')); rect(s, x0, by, 0.04, bh, fill=GOLD)
        text(s, x0 + 0.2, by + 0.08, 8, 0.18, 'ЩО ЦЕ ЗНАЧИТЬ ДЛЯ НАШИХ РОЗМІРІВ', size=7, color=AMBER, bold=True, spc=0.2)
        for k, t_ in enumerate(notes):
            text(s, x0 + 0.2, by + 0.3 + k * 0.24, 9.0, 0.22, t_, size=8, color=INK)
    # right: ours
    ours = collections.OrderedDict()
    for o in OUR:
        if o['seg'] in sgs:
            ours.setdefault((o['sup'], o['grams'], round(o['shelf'])), []).append(o)
    px = 10.2
    section_tag(s, px, 1.85, 2.5, 'Наша полиця проти ринку', color=AMBER)
    ch = min(0.98, 5.0 / max(1, len(ours)))
    yy = 2.2
    for (sid, g, sh), os_ in ours.items():
        o = os_[0]
        sg = o['seg']
        part = dict(tabs)[sg]
        bk = D.basket(o)
        diff = o['shelf'] - bk['pmed']
        rect(s, px, yy, 2.52, ch - 0.08, fill=PANEL); rect(s, px, yy, 0.05, ch - 0.08, fill=SUP_COLOR[sid])
        text(s, px + 0.14, yy + 0.04, 2.3, 0.2, '%s · %s' % (O_clean(o), gfmt(g)) + ('  ×%d' % len(os_) if len(os_) > 1 else ''), size=7.4, color=INK, bold=True)
        text(s, px + 0.14, yy + 0.25, 2.3, 0.2, 'наша полиця %d ₴' % round(o['shelf']), size=8.6, color=SUP_COLOR[sid], bold=True)
        text(s, px + 0.14, yy + 0.47, 2.3, ch - 0.55, 'ринок поруч: %d поз., типова %d ₴ → наша %s на %d ₴' % (bk['n'], round(bk['pmed']), 'вища' if diff > 0 else 'нижча', abs(round(diff))), size=7, color=MUTED)
        yy += ch
    src_note(s)
    return s


def O_clean(o):
    import re
    t = o['title'].replace('Seaweed Topping', 'Topping').replace('Tempura Seaweed', 'Tempura').replace('Sandwich Seaweed', 'Sandwich').replace('Arare Norimaki', 'Norimaki')
    return re.sub(r'\s*·\s*\d+\s*г$', '', t)
