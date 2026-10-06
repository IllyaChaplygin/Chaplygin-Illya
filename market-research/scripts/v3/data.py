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
         fmt='pouch', seg=1, img=RTE + 'sku/ua_markel_soy.jpg'),
    dict(brand='Gallina Blanca', name='Yatekomo: яловичина', g=84, lo=229, hi=229, seller='Товари з Іспанії',
         fmt='cup', seg=2, img=V3 + 'img/51b547b7d3.jpg'),
    dict(brand='Gallina Blanca', name='Yatekomo: теріякі', g=84, lo=229, hi=229, seller='Товари з Іспанії',
         fmt='cup', seg=2, img=V3 + 'img/47975023c3.jpg'),
]
SKU += EXTRA

# ── уточнення діапазонів за новим обходом Prom (06.10.2026) ──
def _upd(brand, nm, **kw):
    for x in SKU:
        if x['brand'] == brand and x['name'] == nm and (kw.get('g_match') in (None, x['g'])):
            kw.pop('g_match', None); x.update(kw); return
    raise KeyError((brand, nm))
_upd('Portion', 'Рис із куркою та овочами', name='Каша рисова з куркою', lo=71, hi=109, seller='portion.com.ua · Rozetka')
_upd('Маркел', "Рис із соєвим м'ясом", lo=80, hi=96, seller='СУХПАЙ · UPcompany · Мартел-shop')
_upd('James Cook', 'Карі з рисом та куркою', lo=73, seller='ВсеОпт · Highlander · Activity')
_upd('!FEST', 'Плов', hi=238, seller='SportStorm · Kalush-Craft · Вояджер · Військторг')
_upd('Travellunch', 'Бефстроганов', g_match=250.0, lo=415, seller='Freeride · Highlander · ForCamp')
_upd('Ben\'s Original', 'Basmati', g_match=220.0, hi=210, seller='MAUDAU · Сільпо · Товари з Іспанії')

NEW = V3 + 'new/'
EXTRA2 = [
    # ── пауч: український реторт (роздріб 71–199 грн, перевірено на сайтах продавців)
    dict(brand="М'ясниця", name='Каша зі свининою, горошком, кукурудзою', g=350, lo=88, hi=89, seller='Belorfoods', fmt='pouch', seg=1, img=NEW + 'ua_myasnytsia.jpg'),
    dict(brand='МАКРО', name='Каша зі свининою та овочами', g=350, lo=96, hi=135, seller='UPcompany · СУХПАЙ · Ліхтар', fmt='pouch', seg=1, img=NEW + 'ua_pak_pork.jpg'),
    dict(brand='МАКРО', name='Каша з курятиною та овочами', g=350, lo=101, hi=101, seller='СУХПАЙ', fmt='pouch', seg=1, img=NEW + 'ua_makro_chicken.jpg'),
    dict(brand='МАКРО', name='Рис з яловичиною та солодким перцем', g=350, lo=104, hi=107, seller='UPcompany · СУХПАЙ', fmt='pouch', seg=1, img=NEW + 'ua_pak_beef.jpg'),
    dict(brand='Верес', name='Каша зі свининою, горошком, кукурудзою', g=350, lo=100, hi=100, seller='МореПродуктів', fmt='pouch', seg=1, img=NEW + 'ua_veres_pork.jpg'),
    dict(brand='Верес', name="Каша з м'ясом курки", g=350, lo=108, hi=135, seller='Козуб · Смачна адреса · Продукт-Shop', fmt='pouch', seg=1, img=NEW + 'ua_veres_chicken.jpg'),
    dict(brand='Ходорівський', name='Каша зі свининою та овочами', g=350, lo=105, hi=105, seller='Молочний склад', fmt='pouch', seg=1, img=NEW + 'ua_khodoriv_350.jpg'),
    dict(brand='Ходорівський', name='Каша зі свининою, горошком, кукурудзою', g=350, lo=140, hi=199, seller='Foodi Shop · Food Shop · МореПродуктів', fmt='pouch', seg=1, img=NEW + 'ua_khodoriv_pork.jpg'),
    dict(brand='Маркел', name='Рис із рослинним фаршем', g=350, lo=127, hi=144, seller='СУХПАЙ · UPcompany', fmt='pouch', seg=1, img=NEW + 'ua_markel_eat.jpg'),
    # ── пауч: Seeds of Change (США, органіка), перепродаж через Prom
    dict(brand='Seeds of Change', name='Organic Jasmine Rice', g=240, lo=225, hi=449, seller='FRESH · VITADOBI', fmt='pouch', seg=1, img=NEW + 'soc_jasmine.jpg'),
    dict(brand='Seeds of Change', name='Brown & Wild Rice, томат і часник', g=240, lo=300, hi=300, seller='FRESH', fmt='pouch', seg=1, img=NEW + 'soc_brownwild.jpg'),
    dict(brand='Seeds of Change', name='Brown Jasmine, кінза й лайм', g=240, lo=259, hi=259, seller='FRESH', fmt='pouch', seg=1, img=NEW + 'soc_cilantro.jpg'),
    dict(brand='Seeds of Change', name='Brown Basmati Rice', g=240, lo=208, hi=287, seller='Карман · BIOLIFE · ТАБЛЕТКА', fmt='pouch', seg=1, img=NEW + 'soc_brownbasmati.jpg'),
    dict(brand='Seeds of Change', name='Quinoa & Brown Rice з часником', g=240, lo=210, hi=400, seller='Zdorovo · FAIR', fmt='pouch', seg=1, img=NEW + 'soc_quinoagarlic.jpg'),
    dict(brand='Seeds of Change', name='Spanish Style з кіноа та перцем', g=240, lo=407, hi=407, seller='BIOLIFE', fmt='pouch', seg=1, img=NEW + 'soc_spanish.jpg'),
    # ── дойпак: нові сублімати й рис під окріп
    dict(brand='Харчі', name='Плов з басматі та телятиною', g=85, lo=195, hi=195, seller='Харчі ТМ', fmt='doypack', seg=4, img=NEW + 'kh_plov_veal.jpg'),
    dict(brand='Харчі', name='Плов: багато рису, курка й морква', g=85, lo=217, hi=217, seller='Харчі ТМ', fmt='doypack', seg=4, img=NEW + 'kh_plov_rice.jpg'),
    dict(brand='Adventure Food', name='Nasi Cashew: індонезійський рис', g=140, lo=492, hi=492, seller='UA-market', fmt='doypack', seg=4, img=NEW + 'af_cashew.jpg'),
    dict(brand='Qiaoshanmei', name='Курка із зеленим перцем', g=146, lo=501, hi=501, seller='Скарби Азії', fmt='doypack', seg=2, img=NEW + 'qs_chicken.jpg'),
    # ── коробка з нагрівачем
    dict(brand='Haidilao', name='Гострий м\'ясний рис', g=360, lo=950, hi=950, seller='Daruy', fmt='box', seg=3, img=NEW + 'hd_spicy360.jpg'),
    dict(brand='Forestia', name='Бефстроганов з рисом', g=None, lo=935, hi=935, seller='Kamanti', fmt='box', seg=3, img=NEW + 'forestia.jpg'),
]
SKU += EXTRA2

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

# ── похідні величини для всієї колоди ──
TOURIST = {'Daruy', 'Highlander', 'ALANTUR', 'ВсеОпт', 'Freeride', 'Лєєр', 'Гайдамака', 'ForCamp', 'Kamanti', '110вольт',
           'MK-Sport', 'MK', 'Суренж', 'Висот-Нік', 'Desna', 'Tactico', 'Шериф', 'Клуб Мандрівник', 'OXO', 'Palmer', 'Klever',
           'Activity', 'Terra Incognita', 'SportStorm', 'Kalush-Craft', 'Вояджер', 'Військторг', 'Ліхтар', 'UA-market',
           'Харчі ТМ', 'власний магазин', 'ще 3', '8 продавців'}
SELL_NORM = {'Gurm.': 'Gurmissimo', 'Апетіт.': 'Апетітаріум', 'OMG!': 'OMG! Asia', 'MK': 'MK-Sport', 'Тайякі': 'Тайякі Март',
             'ще 3': None, '8 продавців': None, 'спец. магазини': None, 'власний магазин': None}


def _toks(x):
    return [t.strip() for t in re.split(r'\s*·\s*', x['seller']) if t.strip()]


for x in SKU:
    x['mass'] = any(t not in TOURIST for t in _toks(x))
N = len(SKU)
N_MASS = sum(1 for x in SKU if x['mass'])
N_BR = len({x['brand'] for x in SKU})
_S = set()
for x in SKU:
    for t in _toks(x):
        t = SELL_NORM.get(t, t)
        if t: _S.add(t)
N_SELL = len(_S)
MASS_BY_FMT = {f: sum(1 for x in by_fmt(f) if x['mass']) for f in FMT_ORDER}
COUNTRY = {"Ben's Original": 'ЄС', 'Маркел': 'Україна', 'Portion': 'Україна', "М'ясниця": 'Україна', 'МАКРО': 'Україна',
           'Верес': 'Україна', 'Ходорівський': 'Україна', 'Bibigo': 'Корея', 'Ottogi': 'Корея', 'Clearspring': 'ЄС',
           'Adventure Menu': 'ЄС', 'Seeds of Change': 'США', 'Henan': 'Китай', 'Gallina Blanca': 'ЄС', 'Qiaoshanmei': 'Китай',
           'Haidilao': 'Китай', 'Zihaiguo': 'Китай', 'Rongcheng Haoji': 'Китай', 'Forestia': 'ЄС', 'Mo Xiao Xian': 'Китай',
           'James Cook': 'Україна', '!FEST': 'Україна', 'Їжа в Похід': 'Україна', 'Харчі': 'Україна', 'SubliMate': 'Україна',
           "Trek'n Eat": 'ЄС', 'Adventure Food': 'ЄС', 'Travellunch': 'ЄС', 'Mountain House': 'США'}


def brand_row(brand, fmt=None, seg=None, label=None):
    L = [x for x in SKU if x['brand'] == brand and (fmt is None or x['fmt'] == fmt) and (seg is None or x['seg'] == seg)]
    lo = min(x['lo'] for x in L); hi = max(x['hi'] for x in L)
    return (label or brand, round(lo), round(hi), round(st.median([x['lo'] for x in L])), len(L), COUNTRY[brand], False)
