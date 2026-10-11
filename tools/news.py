#!/usr/bin/env python3
"""Collect real-estate and land headlines from public RSS feeds into docs/news.json.

Stored per item: headline, source, date, link and, when the feed itself supplies one,
the publisher's thumbnail URL (hotlinked, not copied). No article text. Run by .github/workflows/news.yml every 30 min.
"""
import json, re, sys, time, urllib.request, urllib.parse, html
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
from pathlib import Path

OUT = Path(__file__).resolve().parent.parent / 'docs' / 'news.json'
UA = 'Mozilla/5.0 (compatible; AllianzKapitalNewsBot/1.0; +https://github.com/illyachaplygin)'
MAX_AGE_DAYS = 14
MAX_ITEMS = 150

def gnews(q, hl, gl, ceid):
    return 'https://news.google.com/rss/search?' + urllib.parse.urlencode(
        {'q': q + ' when:7d', 'hl': hl, 'gl': gl, 'ceid': ceid})

# (url, lang, topic, source override or None, keyword filter or None)
KW_ME = re.compile(r'nekretnin|zemljišt|zemljist|stan[ao]?v|gradnj|građevin|gradjevin|urbanist|katastar|vila|apartman|plac', re.I)
FEEDS = [
    (gnews('real estate market OR housing market', 'en-US', 'US', 'US:en'), 'en', 'world', None, None),
    (gnews('land sales OR farmland OR "land prices" OR "vacant land"', 'en-US', 'US', 'US:en'), 'en', 'land', None, None),
    (gnews('European property market OR Spain property OR Portugal property OR Croatia property', 'en', 'GB', 'GB:en'), 'en', 'world', None, None),
    (gnews('Montenegro real estate OR Montenegro property', 'en', 'GB', 'GB:en'), 'en', 'me', None, None),
    (gnews('недвижимость Черногория', 'ru', 'RU', 'RU:ru'), 'ru', 'me', None, None),
    (gnews('рынок недвижимости зарубежная недвижимость', 'ru', 'RU', 'RU:ru'), 'ru', 'world', None, None),
    (gnews('земельные участки рынок земли', 'ru', 'RU', 'RU:ru'), 'ru', 'land', None, None),
    (gnews('nekretnine Crna Gora', 'sr', 'ME', 'ME:sr'), 'me', 'me', None, None),
    (gnews('zemljište cijene Crna Gora', 'sr', 'ME', 'ME:sr'), 'me', 'land', None, None),
    ('https://www.theguardian.com/money/property/rss', 'en', 'world', 'The Guardian', None),
    ('https://www.housingwire.com/feed/', 'en', 'world', 'HousingWire', None),
    ('https://www.bisnow.com/rss', 'en', 'world', 'Bisnow', None),
    ('https://www.propertywire.com/feed/', 'en', 'world', 'Property Wire', None),
    ('https://www.vijesti.me/rss', 'me', 'me', 'Vijesti', KW_ME),
    ('https://www.monitor.co.me/feed/', 'me', 'me', 'Monitor', KW_ME),
    ('https://balkaninsight.com/feed/', 'en', 'me', 'Balkan Insight', re.compile(r'montenegro|real estate|property|land', re.I)),
]
LAND = re.compile(r'(?<!e-)\bland\b|farmland|\bacres?\b|\bparcels?\b|участ|земельн|земли\b|zemljišt|zemljist', re.I)
# a headline must be about property at all (Google News queries are loose)
RELEVANT = re.compile(r'real estate|property|properties|housing|home ?(?:sales|prices|buyers)|mortgage|rent|landlord|tenant|developer|construction|land\b|acre|apartment|villa|'
                      r'недвижим|жиль[ея]|квартир|ипотек|аренд|застройщик|участ|земельн|земл[иея]|'
                      r'nekretnin|zemljišt|zemljist|stan[ao]|stanov|gradnj|hipotek|zakup|kvadrat', re.I)
# local transaction lists and unrelated brands
NOISE = re.compile(r'(?:real estate|property) (?:sales|transactions|transfers)\b|property transfers|home sales (?:for|from)|e-land|sports? betting|\bnfl\b|\bmlb\b', re.I)

def fetch(url):
    req = urllib.request.Request(url, headers={'User-Agent': UA, 'Accept': 'application/rss+xml, application/xml, text/xml, */*'})
    with urllib.request.urlopen(req, timeout=25) as r:
        return r.read()

def txt(el, tag):
    n = el.find(tag)
    return (n.text or '').strip() if n is not None and n.text else ''

MEDIA = '{http://search.yahoo.com/mrss/}'

def image_of(it):
    for tag in (MEDIA + 'thumbnail', MEDIA + 'content'):
        n = it.find(tag)
        if n is not None and n.get('url', '').startswith('https://'):
            u = n.get('url')
            if 'i.guim.co.uk' in u:          # Guardian thumbnails come at width=140
                u = re.sub(r'width=\d+', 'width=640', u)
            return html.unescape(u)
    n = it.find('enclosure')
    if n is not None and n.get('url', '').startswith('https://') and (n.get('type', 'image').startswith('image')):
        return n.get('url')
    return ''

def parse(raw):
    root = ET.fromstring(raw)
    for it in root.iter('item'):
        src = it.find('source')
        yield {
            'title': html.unescape(txt(it, 'title')),
            'link': txt(it, 'link'),
            'date': txt(it, 'pubDate'),
            'src': (src.text or '').strip() if src is not None and src.text else '',
            'img': image_of(it),
        }
    for it in root.iter('{http://www.w3.org/2005/Atom}entry'):
        ns = '{http://www.w3.org/2005/Atom}'
        l = it.find(ns + 'link')
        yield {'title': html.unescape(txt(it, ns + 'title')), 'link': l.get('href') if l is not None else '',
               'date': txt(it, ns + 'updated') or txt(it, ns + 'published'), 'src': '', 'img': ''}

def to_dt(s):
    try:
        d = parsedate_to_datetime(s)
    except Exception:
        try:
            d = datetime.fromisoformat(s.replace('Z', '+00:00'))
        except Exception:
            return None
    if d.tzinfo is None:
        d = d.replace(tzinfo=timezone.utc)
    return d.astimezone(timezone.utc)

def norm(t):
    return re.sub(r'\W+', ' ', t.lower()).strip()[:70]

def main():
    now = datetime.now(timezone.utc)
    seen, items, ok, fail = set(), [], 0, []
    for url, lang, topic, src_over, kw in FEEDS:
        try:
            entries = list(parse(fetch(url)))
            ok += 1
        except Exception as e:
            fail.append(f'{url[:60]}: {e}')
            continue
        for e in entries:
            title, link = e['title'], e['link']
            if not title or not link.startswith('http'):
                continue
            src = src_over or e['src']
            if not src_over and ' - ' in title:          # Google News appends " - Publisher"
                head, _, tail = title.rpartition(' - ')
                if head and (not src or tail == src):
                    title, src = head.strip(), src or tail.strip()
            if kw and not kw.search(title):
                continue
            if NOISE.search(title) or (not src_over and not RELEVANT.search(title)):
                continue
            d = to_dt(e['date'])
            if not d or (now - d).days > MAX_AGE_DAYS or d > now:
                continue
            k = norm(title)
            if k in seen:
                continue
            seen.add(k)
            tp = 'land' if (topic != 'me' and LAND.search(title)) else topic
            items.append({'t': title, 'u': link, 's': src or '', 'd': d.strftime('%Y-%m-%dT%H:%M:%SZ'), 'l': lang, 'c': tp, **({'i': e['img']} if e.get('img') else {})})
    items.sort(key=lambda x: x['d'], reverse=True)
    # keep balance: cap per (lang, topic) so one busy feed cannot push others out
    # per (lang, topic) cap; items with a publisher photo are kept first so the page is not all placeholders
    groups = {}
    for it in items:
        groups.setdefault((it['l'], it['c']), []).append(it)
    out = []
    for g in groups.values():
        pic = [x for x in g if x.get('i')][:18]
        rest = [x for x in g if x not in pic][:max(0, 25 - len(pic))]
        out += pic + rest
    out.sort(key=lambda x: x['d'], reverse=True)
    out = out[:MAX_ITEMS]
    if not out:
        print('no items collected, keeping previous file', fail, file=sys.stderr)
        sys.exit(1)
    OUT.write_text(json.dumps({'updated': now.strftime('%Y-%m-%dT%H:%M:%SZ'), 'items': out}, ensure_ascii=False, separators=(',', ':')), encoding='utf-8')
    print(f'{len(out)} items, {ok}/{len(FEEDS)} feeds ok; failed: {fail}')

if __name__ == '__main__':
    main()
