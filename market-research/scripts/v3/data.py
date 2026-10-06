"""Єдина база 84 позицій із фінальної колоди користувача (картки брендів) + 6 карток зі
старого слайду «поодинокі бренди», які користувач прибрав: їх треба розподілити по сегментах."""
import json, re, statistics as st
from PIL import Image
V3 = '/tmp/claude-0/-home-user-Chaplygin-Illya/c5886f59-fb29-5de4-ac0f-e7da8a3536ab/scratchpad/v3/'
RTE = '/tmp/claude-0/-home-user-Chaplygin-Illya/c5886f59-fb29-5de4-ac0f-e7da8a3536ab/scratchpad/rte/'
cards = json.load(open(V3 + 'cards.json'))
BRAND_NORM = {"MO XIAO XIAN": "Mo Xiao Xian", "ZIHAIGUO": "Zihaiguo", "RONGCHENG HAOJI": "Rongcheng Haoji",
              "QIAOSHANMEI": "Qiaoshanmei", "TRAVELLUNCH": "Travellunch", "TREK'N EAT": "Trek'n Eat",
              "MOUNTAIN HOUSE": "Mountain House", "ADVENTURE FOOD": "Adventure Food",
              "ADVENTURE MENU": "Adventure Menu", "SUBLIMATE": "SubliMate", "JAMES COOK": "James Cook",
              "ЇЖА В ПОХІД": "Їжа в Похід", "ХАРЧІ": "Харчі", "!FEST": "!FEST"}
SLIDE_BRAND = {12: "Ben's Original", 13: "Ottogi", 14: "Henan", 15: "Haidilao"}
SLIDE_FMT = {12: 'pouch', 13: 'cup', 14: 'cup', 15: 'box', 16: None, 17: 'doypack', 18: None, 19: 'doypack'}
SLIDE_SEG = {12: 1, 13: 1, 14: 2, 15: 3, 16: None, 17: 4, 18: None, 19: 4}


def money(t):
    t = t.replace('\xa0', '').replace(' ', '').replace('грн', '')
    a = [float(x.replace(',', '.')) for x in re.split('[–-]', t)]
    return a[0], a[-1]


def grams(t):
    m = re.search(r'([\d]+(?:[.,]\d+)?)\s*Г', t.upper())
    return float(m.group(1).replace(',', '.')) if m else None


SKU = []
for sl in cards:
    n = sl['slide']
    if n not in SLIDE_FMT:
        continue
    for c in sl['cards']:
        tx = c['texts']
        name = tx[0]
        if n in SLIDE_BRAND:
            brand = SLIDE_BRAND[n]; wtxt = tx[1]; ptxt = tx[2]; seller = tx[3]
        else:
            bline = tx[1]; brand = BRAND_NORM[bline.split('·')[0].strip().upper()]
            wtxt = bline; ptxt = tx[2]; seller = tx[3]
        g = grams(wtxt)
        lo, hi = money(ptxt)
        fmt, seg = SLIDE_FMT[n], SLIDE_SEG[n]
        if n == 16:
            if brand == 'Mo Xiao Xian': fmt, seg = 'cup', 3
            elif brand in ('Zihaiguo', 'Rongcheng Haoji'): fmt, seg = 'box', 3
            else: fmt, seg = 'doypack', 2           # Qiaoshanmei — плоский пакет, рис під окріп
        if n == 18:
            if brand == 'Adventure Menu' and g == 400: fmt, seg = 'pouch', 1
            else: fmt, seg = 'doypack', 4
        SKU.append(dict(brand=brand, name=name, g=g, lo=lo, hi=hi, seller=seller, fmt=fmt, seg=seg,
                        img=V3 + c['img']))
# 6 карток зі старого слайду 15 (оновлені фото там, де в колоді «фото немає»)
EXTRA = [
    dict(brand='Bibigo', name='Білий рис', g=210, lo=135, hi=189, seller='Смак Кореї · Rozetka · Prom',
         fmt='cup', seg=1, img=RTE + 'md/bibigo_bowl.jpg'),
    dict(brand='Clearspring', name='Brown & Wild Rice', g=250, lo=252, hi=446,
         seller='MAUDAU · спец. магазини', fmt='pouch', seg=1, img=RTE + 'md/clearspring.jpg'),
    dict(brand='Portion', name='Рис із куркою та овочами', g=350, lo=109, hi=109, seller='Rozetka',
         fmt='pouch', seg=1, img=RTE + 'sku/ua_portion_pack.jpg'),
    dict(brand='Маркел', name="Рис із соєвим м'ясом", g=350, lo=94, hi=96, seller='СУХПАЙ · UPcompany',
         fmt='pouch', seg=1, img=V3 + 'img/0a3d4c5c54.jpg'),
    dict(brand='Gallina Blanca', name='Yatekomo: яловичина', g=84, lo=229, hi=229, seller='Товари з Іспанії',
         fmt='cup', seg=2, img=V3 + 'img/51b547b7d3.jpg'),
    dict(brand='Gallina Blanca', name='Yatekomo: теріякі', g=84, lo=229, hi=229, seller='Товари з Іспанії',
         fmt='cup', seg=2, img=V3 + 'img/47975023c3.jpg'),
]
SKU += EXTRA

FMT_ORDER = ['pouch', 'cup', 'doypack', 'box']
FMT_NAME = {'pouch': 'Пауч', 'cup': 'Чаша', 'doypack': 'Дойпак', 'box': 'Коробка'}
FMT_FULL = {'pouch': 'ПАУЧ · РЕТОРТ', 'cup': 'ЧАША · СТАКАН', 'doypack': 'ДОЙПАК · ПЛОСКИЙ ПАКЕТ',
            'box': 'КОРОБКА З НАГРІВАЧЕМ'}
SEG_NAME = {1: 'Готовий рис · розігрів', 2: 'Сухий рис · під окріп', 3: 'Саморозігрів',
            4: 'Сублімація'}


def by_fmt(f): return [s for s in SKU if s['fmt'] == f]


if __name__ == '__main__':
    print(len(SKU))
    for f in FMT_ORDER:
        L = by_fmt(f)
        brands = list(dict.fromkeys(s['brand'] for s in L))
        gs = [s['g'] for s in L if s['g']]
        print(f, len(L), len(brands), min(gs), max(gs), st.median([s['lo'] for s in L]), brands)
