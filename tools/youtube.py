#!/usr/bin/env python3
"""Collect the channel's videos (public RSS) into docs/videos.json.
The RSS feed returns the 15 latest uploads. Lot codes are read from titles ("... М14" -> M14).
Usage: python3 tools/youtube.py [CHANNEL_ID]"""
import json, re, sys, urllib.request, html
import xml.etree.ElementTree as ET
from pathlib import Path

CHANNEL = sys.argv[1] if len(sys.argv) > 1 else 'UCPHnm9xxywCQo2dQbSZXsYA'
OUT = Path(__file__).resolve().parent.parent / 'docs' / 'videos.json'
NS = {'a': 'http://www.w3.org/2005/Atom', 'yt': 'http://www.youtube.com/xml/schemas/2015', 'm': 'http://search.yahoo.com/mrss/'}

raw = urllib.request.urlopen(urllib.request.Request(f'https://www.youtube.com/feeds/videos.xml?channel_id={CHANNEL}', headers={'User-Agent': 'Mozilla/5.0'}), timeout=30).read()
out = []
for e in ET.fromstring(raw).findall('a:entry', NS):
    title = html.unescape(e.findtext('a:title', '', NS)).strip()
    m = re.search(r'(?<![\w])[МM]\s?(\d{1,2})\b', title)        # Cyrillic and Latin M
    desc = (e.findtext('m:group/m:description', '', NS) or '').strip().replace('\n', ' ')
    out.append({'id': e.findtext('yt:videoId', '', NS), 'title': title, 'date': e.findtext('a:published', '', NS)[:10],
                'lot': f'M{int(m.group(1)):02d}' if m else None, 'desc': desc[:220]})
out.sort(key=lambda v: v['date'], reverse=True)
OUT.write_text(json.dumps({'channel': CHANNEL, 'videos': out}, ensure_ascii=False, separators=(',', ':')), encoding='utf-8')
print(len(out), 'videos;', sum(1 for v in out if v['lot']), 'matched to lots')
