"""Собирает docs/data.js из CSV-выгрузки таблицы: один лот на каждый "Объединенный участок" (колонка B).
Использование: python3 tools/build.py sheet.csv
Нужны tools/files.json (фото из Drive), tools/tourimg.json (обложки 3D-туров),
tools/arrays.json (данные из презентации) и tools/coords.csv (координаты лотов)."""
import csv, json, os, re, sys
import warnings
from PIL import Image
warnings.filterwarnings('ignore')

HERE = os.path.dirname(os.path.abspath(__file__))
PUBLIC = {'наша', 'частично'}          # остальные статусы на сайт не попадают
import collections, urllib.request
LAT = str.maketrans('ПпКкСс', 'PpKkSs')

def num(s):
    try:
        return float(s.replace('\xa0', '').replace(' ', '').replace(',', '.'))
    except ValueError:
        return None

def floors(v):
    v = v.strip().replace('цокольный этаж', 'S')
    return v.translate(LAT).replace(' ', '')

def drive_id(u):
    m = re.search(r'/(?:d|folders)/([\w-]+)', u or '')
    return m.group(1) if m else ''

files = json.load(open(os.path.join(HERE, 'files.json'), encoding='utf-8'))
covers = json.load(open(os.path.join(HERE, 'tourimg.json'), encoding='utf-8'))
dead = {'LhWSH', 'hR7pr'}              # Kuula-ссылки, которые отдают 404

def tour(u):
    return '' if not u or any(d in u for d in dead) else u

def dhash(i):
    px = list(Image.open(os.path.join(HERE, '..', 'docs', 'photos', i + '_s.jpg')).convert('L').resize((17, 16)).getdata())
    return ''.join('1' if px[r * 17 + c] > px[r * 17 + c + 1] else '0' for r in range(16) for c in range(16))

def dist(a, b):
    return sum(x != y for x, y in zip(a, b))

def photos_of(folder):
    ids = [i for i, _ in sorted(files.get(folder, []), key=lambda x: x[1])] if folder else []
    return ids[1:] + ids[:1]    # первый кадр в папках это спутниковый план, уводим в конец


def read_coords():
    out = {}
    path = os.path.join(HERE, 'coords.csv')
    if not os.path.exists(path):
        return out
    for r in csv.DictReader(open(path, encoding='utf-8')):
        lat, lng = (r.get('lat') or '').strip(), (r.get('lng') or '').strip()
        if (not lat or not lng) and r.get('url'):
            try:                      # ссылка Google Maps (в том числе короткая): достаём координаты
                u = urllib.request.urlopen(urllib.request.Request(r['url'], headers={'User-Agent': 'Mozilla/5.0'}), timeout=30).geturl()
                m = re.search(r'!3d(-?[\d.]+)!4d(-?[\d.]+)', u) or re.search(r'@(-?[\d.]+),(-?[\d.]+)', u) or re.search(r'[?&]q=(-?[\d.]+),(-?[\d.]+)', u)
                if m:
                    lat, lng = m.group(1), m.group(2)
            except Exception as e:
                print('не удалось разобрать ссылку', r['url'], e)
        if lat and lng:
            out[r['group'].strip()] = {'lat': float(lat), 'lng': float(lng), 'r': float(r.get('radius') or 60)}
    return out

arrays = json.load(open(os.path.join(HERE, 'arrays.json'), encoding='utf-8'))
coords = read_coords()

rows = [r + [''] * (30 - len(r)) for r in list(csv.reader(open(sys.argv[1], encoding='utf-8')))[2:] if any(r)]
groups = collections.OrderedDict()
for r in rows:
    if r[2] in PUBLIC:
        groups.setdefault(r[1], []).append(r)

def uniq(seq):
    seen, res = set(), []
    for x in seq:
        if x and x not in seen:
            seen.add(x); res.append(x)
    return res

def top(seq):
    return collections.Counter(seq).most_common(1)[0][0]

out = []
for code, ms in groups.items():
    areas = [num(m[11]) for m in ms]
    total_area = sum(a for a in areas if a)
    totals = [num(m[12]) if num(m[11]) else None for m in ms]
    total_price = sum(t for t in totals if t)
    photos, hashes = [], []
    for m in ms:
        for i in photos_of(m[22]):
            h = dhash(i)
            if all(dist(h, x) > 6 for x in hashes):
                hashes.append(h); photos.append(i)
    tours = uniq([tour(m[24]) for m in ms] + [tour(m[25]) for m in ms])
    cover = next((covers[t] for t in tours if t in covers), '')
    lot = {
        'id': code, 'view': top([m[3] for m in ms]), 'region': top([m[6] for m in ms]), 'municipality': top([m[7] for m in ms]),
        'area': total_area or None, 'total': total_price or None,
        'ppm': round(total_price / total_area) if total_area and total_price else None,
        'n': len(ms), 'restricted': sum(1 for m in ms if m[4] == 'Да'), 'partial': sum(1 for m in ms if m[2] == 'частично'),
        'parcels': [{'cadastre': m[8], 'area': num(m[11]), 'total': num(m[12]) if num(m[11]) else None, 'partial': m[2] == 'частично', 'restricted': m[4] == 'Да'} for m in ms],
        'plan': ' / '.join(uniq([('' if m[16] in ('No', 'Yes', '') else m[16]) for m in ms])),
        'occupancy': ' / '.join(uniq([m[17] for m in ms])), 'far': ' / '.join(uniq([m[18] for m in ms])),
        'floors': ' / '.join(uniq([floors(m[19]) for m in ms])),
        'photos': photos,
        'videos': uniq([drive_id(m[23]) for m in ms if '/file/d/' in m[23]]),
        'tours': tours, 'tourCover': cover,
        'docs': {k: next((m[c] for m in ms if m[c]), '') for k, c in (('idea', 21), ('utu', 20), ('deck', 26))},
        'geo': coords.get(code), 'info': arrays.get(code),
    }
    lot['docs'] = {k: v for k, v in lot['docs'].items() if v}
    out.append(lot)

open(os.path.join(HERE, '..', 'docs', 'data.js'), 'w', encoding='utf-8').write(
    'const LOTS = ' + json.dumps(out, ensure_ascii=False, separators=(',', ':')) + ';\n')
print(len(out), 'лотов из', sum(len(v) for v in groups.values()), 'кадастровых участков;',
      sum(1 for p in out if p['photos']), 'с фото;', sum(1 for p in out if p['videos']), 'с видео;',
      sum(1 for p in out if p['tours']), 'с 3D;', sum(1 for p in out if p['geo']), 'с координатами')
