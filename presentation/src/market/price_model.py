# -*- coding: utf-8 -*-
"""Self-cost -> partner price -> shelf price, the same formula used for rice.

From RICE_RICE_FINMODEL_ALL_MARGINS_Final.xlsx ("Финансовая модель - Ціни"):

    partner_price = self_cost / (1 - bonus% - margin%)      # SS is the leftover share
    network_bonus = partner_price * bonus%
    our_profit    = partner_price - self_cost - network_bonus   (== partner_price * margin%)
    shelf_price   = partner_price * retail_markup            # store's own markup, on top

Default scenario in that file: bonus 25%, margin 35% (so self-cost = 40% of
partner price), shelf markup x1.4 (+40%). We reuse the identical defaults
here, plus the same margin-sensitivity ladder (35/30/25/20/15%) the rice
workbook carried in its side columns.
"""
import json
import os
import sys

HERE = os.path.dirname(__file__)
DECK = os.path.join(HERE, '..', 'deck')
sys.path.insert(0, DECK)
from catalog import SUPPLIERS, DATA  # noqa: E402

RATE = 45          # грн/$ — one rate for the whole model, as in the rice workbook (D3)
BONUS = 0.25
MARGIN = 0.30
RETAIL_MARKUP = 1.40
SENSITIVITY_MARGINS = [0.35, 0.30, 0.25, 0.20, 0.15]

SC_NAME = {"40'": '40′ контейнер', "20'": '20′ контейнер',
           'LCL 34': 'Збірний 34 м³', 'LCL 17': 'Збірний 17 м³'}


def sc_order(cost):
    return sorted(SC_NAME, key=lambda k: cost[k]['usd'])


def costs(sheet, key):
    for row in DATA[sheet]:
        if row['name'] == key:
            return row['cost']
    raise KeyError('%s / %s' % (sheet, key))


def price_chain(self_cost_uah, bonus=BONUS, margin=MARGIN, markup=RETAIL_MARKUP):
    partner = self_cost_uah / (1 - bonus - margin)
    net_bonus = partner * bonus
    profit = partner - self_cost_uah - net_bonus
    shelf = partner * markup
    return dict(partner=partner, bonus=net_bonus, profit=profit, shelf=shelf,
               margin_check=profit / partner)


def build():
    """One row per retail consumer SKU (BULK/catering packs excluded — they
    aren't a shelf item), at every container scenario, cheapest flagged."""
    rows = []
    for sup in SUPPLIERS:
        for p in sup['products']:
            cost = costs(sup['sheet'], p['key'])
            order = sc_order(cost)
            cheapest = order[0]
            for scen in order:
                c = cost[scen]
                cost_uah = c['usd'] * RATE      # same as the workbook: $ x rate
                chain = price_chain(cost_uah)
                rows.append(dict(
                    supplier_id=sup['id'], supplier=sup['short'], brand=sup['brand'],
                    sku=p['key'], title=p['title'].replace('\n', ' '), unit=p['unit'],
                    badge=p['badge'], scenario=scen, scenario_name=SC_NAME[scen],
                    is_cheapest=(scen == cheapest),
                    cost_usd=c['usd'], cost_uah=cost_uah,
                    partner_uah=chain['partner'], bonus_uah=chain['bonus'],
                    profit_uah=chain['profit'], shelf_uah=chain['shelf'],
                    margin_check=chain['margin_check'],
                ))
    return rows


def build_sensitivity():
    """Shelf price at each SKU's cheapest scenario, across margin targets —
    mirrors the rice workbook's Q/V/AA/AF sensitivity columns."""
    rows = []
    for sup in SUPPLIERS:
        for p in sup['products']:
            cost = costs(sup['sheet'], p['key'])
            scen = sc_order(cost)[0]
            c = cost[scen]
            by_margin = {m: price_chain(c['usd'] * RATE, margin=m) for m in SENSITIVITY_MARGINS}
            rows.append(dict(
                supplier_id=sup['id'], supplier=sup['short'], title=p['title'].replace('\n', ' '),
                unit=p['unit'], scenario=scen, scenario_name=SC_NAME[scen],
                cost_uah=c['usd'] * RATE,
                shelf_by_margin={m: by_margin[m]['shelf'] for m in SENSITIVITY_MARGINS},
                partner_by_margin={m: by_margin[m]['partner'] for m in SENSITIVITY_MARGINS},
            ))
    return rows


if __name__ == '__main__':
    rows = build()
    out = os.path.join(HERE, 'price_model.json')
    json.dump({'params': {'bonus': BONUS, 'margin': MARGIN, 'markup': RETAIL_MARKUP},
              'rows': rows, 'sensitivity': build_sensitivity()},
             open(out, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print('wrote', out, '-', len(rows), 'rows')
    for r in rows:
        if r['is_cheapest']:
            print('%-14s %-28s %-10s СС=%6.2f₴  партнер=%6.2f₴  полиця=%6.2f₴  (маржа %.0f%%)' % (
                r['supplier'], r['title'][:28], r['scenario_name'][:10],
                r['cost_uah'], r['partner_uah'], r['shelf_uah'], r['margin_check'] * 100))
