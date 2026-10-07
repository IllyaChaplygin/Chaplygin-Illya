# -*- coding: utf-8 -*-
"""Retail audit via the public zakaz.ua store API (the aggregator behind the
online catalogues of Metro, Novus, Auchan, Varus-style regional chains).
One representative store per chain; search queries cover seaweed snacks and
rice crackers. Raw hits are written unfiltered — classification is a separate,
reviewable step (classify.py)."""
import json, os, sys, time, urllib.parse, urllib.request, datetime
HERE = os.path.dirname(os.path.abspath(__file__))
QUERIES = ['нори', 'норі', 'водорості', 'морська капуста', 'чипси норі', 'снек норі', 'рисові крекери',
           'рисові чіпси', 'крекер рисовий', 'сенбей', 'аpare', 'арарe', 'tempura seaweed', 'seaweed',
           'tao kae noi', 'рисовий снек', 'кімнорі', 'снеки азійські', 'норі снек', 'рисові хлібці з водоростями']
EXTRA = ['гім','кім','сушені водорості','хрусткі водорості','водорості снек','норі темпура','крекери','крекер азійський','чипси рисові','рисові кульки','мочі','снеки з морепродуктів','креветочні чипси','азійські снеки','корейські снеки','японські снеки','тайські снеки','снек з рису','рисовий крекер з водоростями','печиво рисове','хлібці рисові','тайський снек']
SKIP = {'alcohub', 'winetime', 'okwine', 'masterzoo'}

def get(url, tries=3):
    for i in range(tries):
        try:
            r = urllib.request.Request(url, headers={'Accept-Language': 'uk', 'User-Agent': 'Mozilla/5.0'})
            return json.load(urllib.request.urlopen(r, timeout=30))
        except Exception as e:
            time.sleep(1.5 * (i + 1)); err = e
    return {'error': str(err)}

def main():
    stores = json.load(open(os.path.join(HERE, 'raw', 'stores.json')))
    chosen = {}
    for s in stores:
        if s['retail_chain'] in SKIP: continue
        # prefer Kyiv, else first
        cur = chosen.get(s['retail_chain'])
        if cur is None or (s['city'] == 'kiev' and cur['city'] != 'kiev'): chosen[s['retail_chain']] = s
    hits = {}
    for chain, s in chosen.items():
        for q in (EXTRA if os.environ.get('EXTRA') else QUERIES):
            d = get('https://stores-api.zakaz.ua/stores/%s/products/search/?q=%s' % (s['id'], urllib.parse.quote(q)))
            for p in d.get('results', []):
                key = (chain, p['ean'])
                if key in hits: continue
                hits[key] = dict(chain=chain, store=s['name'], city=s['city'], ean=p['ean'], title=p.get('title'),
                                 producer=(p.get('producer') or {}).get('trademark'), weight=p.get('weight'),
                                 unit=p.get('unit'), price=p.get('price'), old=p.get('old_price') if 'old_price' in p else None,
                                 in_stock=p.get('in_stock'), img=(p.get('img') or {}).get('s350x350') or (p.get('img') or {}).get('s150x150'),
                                 web_url=p.get('web_url'), q=q, cat=p.get('category_id') if 'category_id' in p else None)
        print(chain, s['name'], len([k for k in hits if k[0] == chain]), flush=True)
    out = dict(date=datetime.date.today().isoformat(), items=list(hits.values()))
    json.dump(out, open(os.path.join(HERE, 'raw', ('zakaz_raw_extra.json' if os.environ.get('EXTRA') else 'zakaz_raw.json')), 'w'), ensure_ascii=False, indent=1)
if __name__ == '__main__': main()
