#!/usr/bin/env python3
"""Pull purpose (sale / joint construction), pins and parcel outlines from the company's Google My Maps.

Main map: https://www.google.com/maps/d/viewer?mid=MAIN  (folders = purpose, one pin per lot, the pin
description links to a per-lot detail map with the cadastral polygons).
Writes tools/mymaps.json, which build.py merges into docs/data.js.
Usage: python3 tools/mymaps.py [MAIN_MID]
"""
import json, math, re, sys, urllib.request
from pathlib import Path

MAIN = sys.argv[1] if len(sys.argv) > 1 else '1DhXd4XgfnVFDvkOLD34y6rwQNvU_fZE'
OUT = Path(__file__).resolve().parent / 'mymaps.json'
FOLDER_PURPOSE = {'joint construction': 'joint', 'sale': 'sale'}

def kml(mid):
    req = urllib.request.Request(f'https://www.google.com/maps/d/kml?mid={mid}&forcekml=1', headers={'User-Agent': 'Mozilla/5.0'})
    return urllib.request.urlopen(req, timeout=40).read().decode('utf-8', 'replace')

def ring_area(ring):
    """Area in m2 of a lng/lat ring (local equirectangular projection)."""
    lat0 = sum(p[1] for p in ring) / len(ring)
    kx, ky = 111320 * math.cos(math.radians(lat0)), 110574
    pts = [(p[0] * kx, p[1] * ky) for p in ring]
    return abs(sum(pts[i][0] * pts[(i + 1) % len(pts)][1] - pts[(i + 1) % len(pts)][0] * pts[i][1] for i in range(len(pts)))) / 2

def polygons(doc):
    out = []
    for pm in re.findall(r'<Placemark>(.*?)</Placemark>', doc, re.S):
        if '<Polygon' not in pm:
            continue
        name = (re.search(r'<name>(.*?)</name>', pm, re.S) or [None, ''])[1].strip()
        c = re.search(r'<outerBoundaryIs>.*?<coordinates>(.*?)</coordinates>', pm, re.S)
        if not c:
            continue
        ring = [tuple(map(float, p.split(',')[:2])) for p in c.group(1).split()]
        if len(ring) > 1 and ring[0] == ring[-1]:
            ring = ring[:-1]
        if len(ring) < 3:
            continue
        out.append({'cad': name.split('_')[0], 'm2': round(ring_area(ring)), 'ring': [[round(y, 6), round(x, 6)] for x, y in ring]})
    return out

def main():
    main_kml = kml(MAIN)
    res = {}
    for folder in re.findall(r'<Folder>(.*?)</Folder>', main_kml, re.S):
        purpose = FOLDER_PURPOSE.get(re.search(r'<name>(.*?)</name>', folder).group(1).strip().lower())
        for pm in re.findall(r'<Placemark>(.*?)</Placemark>', folder, re.S):
            name = re.search(r'<name>(.*?)</name>', pm, re.S).group(1).strip()
            m = re.match(r'(M\d+)\s*-\s*([\d\s]+)\s*m2', name)
            if not m:
                continue
            code, stated = m.group(1), int(re.sub(r'\s', '', m.group(2)))
            lng, lat = map(float, re.search(r'<coordinates>\s*([^<]*?)\s*</coordinates>', pm).group(1).split(',')[:2])
            det = re.search(r'maps/d/edit\?mid=([\w-]+)', pm)
            polys = polygons(kml(det.group(1))) if det else []
            res[code] = {'purpose': purpose, 'pin': [lat, lng], 'stated_m2': stated, 'polys': polys}
            tot = sum(p['m2'] for p in polys)
            print(f'{code}: {purpose}, pin {lat:.5f},{lng:.5f}, {len(polys)} polygons, outline {tot} m2 vs stated {stated}')
    OUT.write_text(json.dumps(res, ensure_ascii=False, separators=(',', ':')), encoding='utf-8')

if __name__ == '__main__':
    main()
