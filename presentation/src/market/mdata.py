# -*- coding: utf-8 -*-
"""Aggregations over the market catalogue and our own price model."""
import collections, json, os, re, statistics, sys
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, '..', 'deck'))
from catalog import SUPPLIERS, DATA, photo as sup_photo  # noqa: E402

CAT = json.load(open(os.path.join(HERE, 'market_catalog.json'), encoding='utf-8'))
ROWS = CAT['rows']
DATE = CAT['date']
CHAIN = CAT['chain_name']
NATIONAL = set(CAT['national'])
PM = json.load(open(os.path.join(HERE, 'price_model.json'), encoding='utf-8'))

SEGS = collections.OrderedDict([
    ('mini', dict(name='Норі-снек у міні-пакеті', short='Міні-пакет', desc='3–5 г · один–два аркуші в пакетику')),
    ('chips', dict(name='Норі-чипси та сендвічі', short='Чипси', desc='8–40 г · мультипак, чипси, сендвіч')),
    ('tempura', dict(name='Темпура та топінг', short='Темпура / топінг', desc='15–70 г · у клярі, подрібнені')),
    ('ricecr', dict(name='Рисові крекери', short='Рис. крекери', desc='50–60 г · арарe, сенбей, міні-крекер')),
    ('ricechip', dict(name='Рисові чипси', short='Рис. чипси', desc='50–150 г · суміжний, не наш формат')),
])
# topping merged into tempura for the story (specialty formats from HanJin/ZEK's range)
for r in ROWS:
    r['sub'] = r['seg']
    if r['seg'] == 'topping':
        r['seg'] = 'tempura'


def rows_of(seg):
    return [r for r in ROWS if r['seg'] == seg]


def med(xs):
    xs = list(xs)
    return statistics.median(xs) if xs else None


def lines(seg_rows):
    g = collections.OrderedDict()
    for r in sorted(seg_rows, key=lambda r: (r['brand'], r['grams'], r['pmed'])):
        g.setdefault((r['brand'], r['grams']), []).append(r)
    out = []
    for (brand, grams), rs in g.items():
        chains = sorted({c for r in rs for c in r['chains']}, key=lambda c: (-(c in NATIONAL), CHAIN[c]))
        prices = [r['pmed'] for r in rs]
        out.append(dict(brand=brand, grams=grams, n=len(rs), flavors=[r['name'] for r in rs],
                        pmin=min(r['pmin'] for r in rs), pmax=max(r['pmax'] for r in rs), pmed=med(prices),
                        per_g=med(prices) / grams, chains=chains, img=rs[0]['img'], seg=rs[0]['seg']))
    out.sort(key=lambda l: l['pmed'])
    return out


def gfmt(g):
    return ('%g' % g).replace('.', ',') + ' г'


def uah(v, d=0):
    return (('%.' + str(d) + 'f') % v).replace('.', ',') + ' ₴'


def rng(a, b):
    a, b = round(a), round(b)
    return '%d' % a if a == b else '%d–%d' % (a, b)


# ------------------------------------------------------------------- our SKUs
def _weight(badge):
    m = re.match(r'([\d,]+)\s*г', badge)
    return float(m.group(1).replace(',', '.'))


def our_skus():
    cheap = {r['sku']: r for r in PM['rows'] if r['is_cheapest']}
    out = []
    for sup in SUPPLIERS:
        fob_of = {r['name']: r['fob'] for r in DATA[sup['sheet']]}
        for p in sup['products']:
            r = cheap[p['key']]
            g = _weight(p['badge'])
            if sup['id'] in ('singha', 'thainichi'):
                seg = 'ricecr'
            elif sup['id'] == 'tmk':
                seg = 'mini' if g <= 5 else 'chips'
            else:
                seg = 'chips' if 'SANDWICH' in p['key'] else 'tempura'
            out.append(dict(sup=sup['id'], sup_short=sup['short'], brand=sup['brand'], key=p['key'],
                            title=p['title'].replace('\n', ' '), grams=g, seg=seg,
                            photo=sup_photo(p['photo']), fob=fob_of[p['key']], scenario=r['scenario_name'],
                            usd=r['cost_usd'], cost=r['cost_uah'], partner=r['partner_uah'], bonus=r['bonus_uah'],
                            profit=r['profit_uah'], shelf=r['shelf_uah'], per_g=r['shelf_uah'] / g))
    return out


OUR = our_skus()


def peers(o, lo=0.5, hi=2.0):
    """Market positions comparable to one of our SKUs: same segment, weight within [lo, hi] x ours."""
    rs = [r for r in rows_of(o['seg']) if lo * o['grams'] <= r['grams'] <= hi * o['grams']]
    return rs


def peer_stats(o):
    rs = peers(o)
    if len(rs) < 3:
        rs = peers(o, 0.34, 3.0)
    if not rs:
        rs = rows_of(o['seg'])
    return dict(n=len(rs), per_g=med(r['per_g'] for r in rs), pack=med(r['pmed'] for r in rs),
                lo=min(r['pmin'] for r in rs), hi=max(r['pmax'] for r in rs))


if __name__ == '__main__':
    for sg in SEGS:
        rs = rows_of(sg); ls = lines(rs)
        print(sg, len(rs), 'SKU', len(ls), 'lines', 'median pack %.0f' % med(r['pmed'] for r in rs), '₴/г %.1f' % med(r['per_g'] for r in rs),
              'range %.0f-%.0f' % (min(r['pmin'] for r in rs), max(r['pmax'] for r in rs)), 'brands', len({r['brand'] for r in rs}))
        mine = [o for o in OUR if o['seg'] == sg]
        for o in mine: print('   ours', o['sup'], o['title'][:40], o['grams'], round(o['shelf']), round(o['per_g'], 1))


def analog(o):
    """The market pack our SKU is actually compared with: the most widely stocked line
    of the same segment whose weight is within x0.6..x1.7 of ours (fallback: nearest weight)."""
    ls = lines(rows_of(o['seg']))
    near = [l for l in ls if 0.6 * o['grams'] <= l['grams'] <= 1.7 * o['grams']]
    if not near:
        return min(ls, key=lambda l: abs(l['grams'] - o['grams']))
    return max(near, key=lambda l: (l['n'] * len(l['chains']), -abs(l['grams'] - o['grams'])))


def our_groups():
    """Our SKUs collapsed by (supplier, weight, shelf price) — flavours with identical numbers."""
    g = collections.OrderedDict()
    for o in OUR:
        g.setdefault((o['sup'], o['grams'], round(o['shelf'])), []).append(o)
    out = []
    for k, v in g.items():
        o = dict(v[0]); o['n_sku'] = len(v); o['all'] = v
        out.append(o)
    return out


def basket(o):
    """Market positions of the same segment with a pack weight close to ours — the
    like-for-like shelf we are priced against. Weight window x0.55..x1.45, widened
    to x0.4..x1.7 when fewer than three positions fall inside."""
    pool = rows_of(o['seg'])
    if o['seg'] == 'tempura':      # topping and tempura are different products: compare like with like
        want_top = 'TOPPING' in o['key'].upper()
        pool = [r for r in pool if (r['sub'] == 'topping') == want_top]
    rs = [r for r in pool if 0.55 * o['grams'] <= r['grams'] <= 1.45 * o['grams']]
    if len(rs) < 3:
        rs = [r for r in pool if 0.4 * o['grams'] <= r['grams'] <= 1.7 * o['grams']]
    if len(rs) < 3:
        rs = sorted(pool, key=lambda r: abs(r['grams'] - o['grams']))[:3]
    return dict(n=len(rs), gmin=min(r['grams'] for r in rs), gmax=max(r['grams'] for r in rs), pmed=med(r['pmed'] for r in rs),
                pmin=min(r['pmed'] for r in rs), pmax=max(r['pmed'] for r in rs), brands=sorted({r['brand'] for r in rs}))


# ---------------------------------------------------------------- brand directory
# Importer / origin come ONLY from the customs base (importer names and brand tokens in the
# declaration descriptions). Where the base shows nothing, the cell says so — no guessing.
_CUST = json.load(open(os.path.join(HERE, 'customs.json'), encoding='utf-8'))
_TOKEN = {'Haelove': 'Haelove / Delisse', 'Tao Kae Noi': 'Tao Kae Noi', 'Akura': 'Akura', 'Royal Tiger': 'Royal Tiger',
          'Ock Dong Ja': 'Ock Dong Ja', 'Metro Chef': 'Metro Chef'}
BRAND_SRC = {'Norris': ('Норріс Груп', 'Китай'), 'Clearspring': ('Бюро Вин', 'Корея (Sahm Yook Sea Foods)')}


def _short_imp(n):
    return n.replace('ТОВ ', '').replace('ОРІЄНТАЛЬ ПЛЮС', 'Оріенталь Плюс').replace('ФОЗЗІ КОММЕРЦ', 'Фоззі Коммерс').replace('ОРІЄНТАЛ МЕРЧАНТ', 'Оріентал Мерчант').replace('Метро Кеш Енд Кері Україна', 'Metro Cash & Carry').replace('Торговий дім ІТС', 'ТД ІТС').replace('АЗІАТІК', 'Азіатік').replace('НОРРІС ГРУП', 'Норріс Груп')


def brand_origin(brand):
    if brand in BRAND_SRC:
        return BRAND_SRC[brand]
    tok = _TOKEN.get(brand)
    if tok:
        imps = [i for i in _CUST['seaweed']['importers'] if tok in i['brands']]
        if imps:
            ctry = {'КОРЕЯ РЕСПУБЛІКА': 'Корея', 'КИТАЙ': 'Китай', 'ТАЇЛАНД': 'Таїланд'}.get(imps[0]['origin'], imps[0]['origin'].title())
            if brand == 'Tao Kae Noi':
                ctry = 'Таїланд (Taokaenoi Food)'
            return (', '.join(_short_imp(i['name']) for i in imps[:2]), ctry)
    return (None, None)


def brand_profile(brand):
    rs = [r for r in ROWS if r['brand'] == brand]
    segs = collections.OrderedDict()
    for r in rs:
        segs.setdefault(SEGS[r['seg']]['short'], set()).add(r['grams'])
    fmt = '; '.join('%s %s' % (k, ' / '.join(('%g' % g).replace('.', ',') for g in sorted(v)) + ' г') for k, v in segs.items())
    pr = [r['pmed'] for r in rs]
    return dict(brand=brand, fmt=fmt, n=len(rs), chains=len({c for r in rs for c in r['chains']}), pmin=min(pr), pmax=max(pr), pmed=med(pr),
                imp=brand_origin(brand))
