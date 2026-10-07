# -*- coding: utf-8 -*-
"""Turn raw zakaz.ua hits into the market catalogue: one row per EAN, with the
segment, shelf-price range across the chains that carry it, and a local photo.
Rules are explicit so the segmentation can be reviewed and argued with."""
import collections, hashlib, json, os, re, statistics, urllib.request
HERE = os.path.dirname(os.path.abspath(__file__))
RAW = os.path.join(HERE, 'raw', 'zakaz_raw.json')
IMG = os.path.join(HERE, 'raw', 'img')

CHAIN_NAME = {'auchan': 'Ашан', 'novus': 'NOVUS', 'metro': 'METRO', 'zaraz': 'Зараз', 'megamarket': 'МегаМаркет',
              'ultramarket': 'Ultramarket', 'cosmos': 'Космос', 'epicentr': 'Епіцентр', 'onde': 'Onde',
              'ideal': 'Ідеал', 'vostorg': 'Восторг', 'kharkiv': 'Клас', 'tavriav': 'Таврія В',
              'chudomarket': 'ЧудоМаркет', 'torba': 'Торба', 'grono': 'Grono', 'ekomarket': 'ЕКО', 'biotus': 'Біотус'}
# national = federal/multi-region networks; the rest regional (same split as in the SIAS deck)
NATIONAL = {'auchan', 'novus', 'metro', 'zaraz', 'megamarket', 'ultramarket', 'epicentr'}

SEG = collections.OrderedDict([
    ('mini', 'Норі-снек у міні-пакеті'),
    ('chips', 'Норі-чипси та рол'),
    ('tempura', 'Норі в темпурі'),
    ('topping', 'Топінг і подрібнені норі'),
    ('ricecr', 'Рисові крекери'),
    ('ricechip', 'Рисові чипси'),
])

EXCL = ('для суші', 'для приготування', 'листів', 'аркуш', 'ковбас', 'маринован', 'ламінар', 'кукурудз', 'печиво',
        'масло', 'соус', 'маска', 'крем', 'гель', 'бальзам', 'корм', 'боул', 'молоко', 'вакаме', 'сіль', 'капсул',
        'таблеток', 'омега', 'загущувач', 'рідина', 'папір', 'локшина', 'lay', 'картопл', 'кокосов', 'креветоч',
        'креветков', 'салат', 'листи', 'листа', 'gold 10', 'морські сушені', 'сушені', 'wakame', 'з курки', 'креветки', '10шт', 'lorenz', 'gerber', 'cookia', 'пшенич', 'приправ', 'nongshim', 'крекер nongshim', 'кім nicos')

BRAND_FIX = {'Taokaenoi': 'Tao Kae Noi', 'Taokaenoi ': 'Tao Kae Noi', 'Ock Dong Ja': 'Ock Dong Ja', 'Kimnori Snack': 'Kimnori',
             'B.yond': 'B.yond', 'Rice UP!': 'Rice Up!'}


def short_title(t, brand):
    s = t
    s = re.sub(r'\s*\d+[.,]?\d*\s*(г|мл)\b', '', s)
    s = re.sub(r'^(Чипси|Чіпси|Норі|Снеки?|Водорості|Хрусткі|Крекери?|Рисовий|Темпура)\s+(норі\s+)?', '', s, flags=re.I)
    for b in {brand, brand.replace(' ', ''), 'Tao Kae Noi', 'Taokaenoi', 'Tао Kае Nоі'}:
        s = re.sub(re.escape(b), '', s, flags=re.I)
    s = re.sub(r'\b(снек|снеки|норі|чипси|рисові|рисовий)\b', '', s, flags=re.I)
    s = re.sub(r'\bз з\b', 'з', s)
    s = re.sub(r'\s+', ' ', s).strip(' ,-·')
    s = re.sub(r'\bз з\b', 'з', s)
    s = re.sub(r'^(Ock-?Dong-?Ja)\s+', '', s, flags=re.I)
    s = re.sub(r'\s+', ' ', s).strip(' ,-·.')
    return s[:1].upper() + s[1:] if s else t


def segment(t, g, brand):
    tl = t.lower()
    if any(e in tl for e in EXCL) and not ('тамарі' in tl):
        return None
    if g is None:
        m = re.search(r'(\d+[.,]?\d*)\s*г', tl)
        g = float(m.group(1).replace(',', '.')) if m else None
    if 'рисов' in tl or 'крекер' in tl:
        return 'rice'
    if 'темпур' in tl or 'tempura' in tl or brand.startswith('Taokaenoi') or 'хрусткі норі' in tl or 'шрірача' in tl:
        return 'tempura'
    if 'подрібнен' in tl or 'нарізан' in tl or 'топінг' in tl:
        return 'topping'
    if g is None:
        return None
    if g <= 5.5 and ('рол' not in tl):
        return 'mini'
    if 'норі' in tl or 'кім' in tl or 'seaweed' in tl or 'водорост' in tl:
        if g <= 5.5: return 'mini'
        return 'chips'
    return None


def main():
    raw = json.load(open(RAW, encoding='utf-8'))
    by = collections.defaultdict(list)
    for x in raw['items']:
        by[x['ean']].append(x)
    os.makedirs(IMG, exist_ok=True)
    rows = []
    for ean, v in by.items():
        x = v[0]
        brand = BRAND_FIX.get((x['producer'] or '').strip(), (x['producer'] or '').strip())
        t = x['title']
        if brand in ('без тм', '') and 'ock-dong-ja' in t.lower():
            brand = 'Ock Dong Ja'
        g = x['weight'] if x['weight'] and x['unit'] == 'pcs' else None
        seg = segment(t, g, (x['producer'] or '').strip())
        if seg == 'rice':
            seg = 'ricecr' if brand in ('Want Want', 'Clearspring', 'Ultra Pop') else 'ricechip'
        if not seg:
            continue
        if g is None:
            m = re.search(r'(\d+[.,]?\d*)\s*г', t)
            g = float(m.group(1).replace(',', '.')) if m else None
        # a Dong Won / Ock Dong Ja pack stated as 4,0 vs 4,5 etc. is kept as listed
        if not g:
            continue
        prices = sorted(p['price'] / 100 for p in v if p.get('price'))
        chains = sorted({p['chain'] for p in v}, key=lambda c: (-int(c in NATIONAL), CHAIN_NAME[c]))
        img = next((p['img'] for p in v if p.get('img')), None)
        path = None
        if img:
            path = os.path.join(IMG, '%s.jpg' % ean)
            if not os.path.exists(path):
                try:
                    r = urllib.request.Request(img, headers={'User-Agent': 'Mozilla/5.0'})
                    open(path, 'wb').write(urllib.request.urlopen(r, timeout=30).read())
                except Exception as e:
                    path = None
        rows.append(dict(ean=ean, brand=brand, title=t, name=short_title(t, brand), grams=g, seg=seg,
                         pmin=prices[0], pmax=prices[-1], pmed=statistics.median(prices), chains=chains,
                         n_chains=len(chains), per_g=statistics.median(prices) / g,
                         img=os.path.relpath(path, HERE) if path else None))
    rows.sort(key=lambda r: (list(SEG).index(r['seg']), r['pmed']))
    json.dump(dict(date=raw['date'], rows=rows, chain_name=CHAIN_NAME, national=sorted(NATIONAL)),
              open(os.path.join(HERE, 'market_catalog.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    c = collections.Counter(r['seg'] for r in rows)
    print(len(rows), dict(c), 'no photo:', sum(1 for r in rows if not r['img']))


if __name__ == '__main__':
    main()
