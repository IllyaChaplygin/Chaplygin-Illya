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
