"""Собирает docs/data.js из CSV-выгрузки таблицы.
Использование: python3 tools/build.py sheet.csv
Нужны tools/files.json (фото из Drive) и tools/tourimg.json (обложки 3D-туров)."""
import csv, json, os, re, sys
import warnings
from PIL import Image
warnings.filterwarnings('ignore')

HERE = os.path.dirname(os.path.abspath(__file__))
PUBLIC = {'наша', 'частично'}          # остальные статусы на сайт не попадают
LAT = str.maketrans('ПпКкСс', 'PpKkSs')

def num(s):
    try:
        return float(s.replace('\xa0', '').replace(' ', '').replace(',', '.'))
    except ValueError:
        return None

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

rows = [r + [''] * (30 - len(r)) for r in list(csv.reader(open(sys.argv[1], encoding='utf-8')))[2:] if any(r)]
def photos_of(folder):
    ids = [i for i, _ in sorted(files.get(folder, []), key=lambda x: x[1])] if folder else []
    return ids[1:] + ids[:1]    # первый кадр в папках это спутниковый план, уводим в конец

out = []
for r in rows:
    if r[2] not in PUBLIC:
        continue
    area = num(r[11])
    t, tg = tour(r[24]), tour(r[25])
    out.append({
        'id': int(r[0]), 'group': r[1], 'status': r[2], 'view': r[3],
        'restricted': r[4] == 'Да', 'region': r[6], 'municipality': r[7],
        'cadastre': r[8], 'area': area, 'total': num(r[12]) if area else None,
        'ppm': num(r[13]) if area else None,
        'plan': '' if r[16] in ('No', '') else r[16],
        'occupancy': r[17], 'far': r[18], 'floors': r[19].translate(LAT),
        'photos': photos_of(r[22]),
        'video': drive_id(r[23]) if '/file/d/' in r[23] else '',
        'tour': t, 'tourGroup': tg,
        'tourCover': covers.get(t) or covers.get(tg) or '',
        'docs': {k: v for k, v in (('idea', r[21]), ('utu', r[20]), ('deck', r[26])) if v},
    })
# Во многих папках один и тот же набор снимков территории на несколько участков: считаем, сколько участков делят набор.
sigs = []
for p in out:
    if not p['photos']:
        continue
    h = dhash(p['photos'][0])
    for s_ in sigs:
        if dist(h, s_[0]) <= 6:
            s_[1].append(p); break
    else:
        sigs.append((h, [p]))
for _, ps in sigs:
    for p in ps:
        p['shared'] = len(ps)

open(os.path.join(HERE, '..', 'docs', 'data.js'), 'w', encoding='utf-8').write(
    'const PARCELS = ' + json.dumps(out, ensure_ascii=False, separators=(',', ':')) + ';\n')
print(len(out), 'участков;', sum(1 for p in out if p['photos']), 'с фото;',
      sum(1 for p in out if p['video']), 'с видео;', sum(1 for p in out if p['tour'] or p['tourGroup']), 'с 3D;',
      sum(1 for p in out if not p['photos'] and not p['tourCover']), 'без медиа')
