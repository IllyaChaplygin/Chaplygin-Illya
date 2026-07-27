"""Extract per-SKU self-cost by logistics scenario from SelfCost.xlsx into data.json."""
import json
import openpyxl

SRC = '/root/.claude/uploads/c53f34b4-ae43-52db-aeb2-01bdcda48cc3/9c9fe902-SelfCost.xlsx'

# sheet -> (name_col, unit_price_col, usd_col, uah_col, rate_col, {block_start_row: scenario})
LAYOUT = {
    'SINGHA KAMEDA - Retail': dict(
        name='C', price='I', usd='AI', uah='AJ', rate='AE',
        blocks={6: "20'", 15: "40'", 24: 'LCL 17', 33: 'LCL 34'}, span=2),
    'Thai-Nichi - Retail': dict(
        name='C', price='I', usd='AI', uah='AJ', rate='AE',
        blocks={6: "20'", 15: "40'", 24: 'LCL 17', 33: 'LCL 34'}, span=2),
    'TMK Thailand Co., Ltd -Retail': dict(
        name='C', price='M', usd='AM', uah='AN', rate='AI',
        blocks={6: "20'", 22: "40'", 39: 'LCL 17', 56: 'LCL 34'}, span=10),
    'ZEK -Retail': dict(
        name='C', price='M', usd='AM', uah='AN', rate='AI',
        blocks={6: "20'", 26: "40'", 46: 'LCL 17', 66: 'LCL 34'}, span=14),
}


def clean(s):
    return ' '.join(str(s).replace('•', '').replace('\t', ' ').split())


def main():
    wb = openpyxl.load_workbook(SRC, data_only=True)
    out = {}
    for sheet, cfg in LAYOUT.items():
        ws = wb[sheet]
        skus = {}
        order = []
        for start, scenario in cfg['blocks'].items():
            for i in range(cfg['span']):
                r = start + i
                name = ws[cfg['name'] + str(r)].value
                assert name, (sheet, r)
                name = clean(name)
                if name not in skus:
                    skus[name] = {'name': name, 'fob': ws[cfg['price'] + str(r)].value,
                                  'cost': {}}
                    order.append(name)
                skus[name]['cost'][scenario] = {
                    'usd': round(ws[cfg['usd'] + str(r)].value, 4),
                    'uah': round(ws[cfg['uah'] + str(r)].value, 2),
                    'rate': ws[cfg['rate'] + str(r)].value,
                }
        out[sheet] = [skus[n] for n in order]

    with open('/home/user/work/data.json', 'w') as f:
        json.dump(out, f, ensure_ascii=False, indent=1)

    for sheet, rows in out.items():
        print('=' * 70)
        print(sheet, len(rows), 'SKU')
        for r in rows:
            c = r['cost']
            print('  %-42s FOB $%-7s | ' % (r['name'][:42], r['fob']) +
                  ' | '.join('%s $%.3f/%.2f₴' % (k, v['usd'], v['uah'])
                             for k, v in c.items()))


if __name__ == '__main__':
    main()
