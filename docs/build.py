"""Генерирует data.js из CSV-экспорта таблицы земель. Использование: python3 build.py sheet.csv"""
import csv, json, sys

def num(s):
    try:
        return float(s.replace('\xa0', '').replace(' ', '').replace(',', '.'))
    except ValueError:
        return None

rows = [r for r in list(csv.reader(open(sys.argv[1], encoding='utf-8')))[2:] if any(r)]
out = []
for r in rows:
    r += [''] * (30 - len(r))
    area = num(r[11])
    out.append({
        'id': int(r[0]), 'group': r[1], 'status': r[2], 'view': r[3],
        'restricted': r[4] == 'Да', 'region': r[6], 'municipality': r[7],
        'cadastre': r[8], 'area': area, 'total': num(r[12]) if area else None,
        'ppm': num(r[13]), 'plan': r[16], 'occupancy': r[17], 'far': r[18],
        'floors': r[19], 'tour': r[24], 'tourGroup': r[25],
        'hasPhoto': bool(r[22]), 'hasVideo': bool(r[23]),
    })
open('data.js', 'w', encoding='utf-8').write('const PARCELS = ' + json.dumps(out, ensure_ascii=False) + ';\n')
print(len(out), 'участков')
