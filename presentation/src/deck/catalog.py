# -*- coding: utf-8 -*-
"""Supplier / product catalogue.

Self-cost figures come from SelfCost.xlsx via data.json (never typed by hand).
`key` matches the "Наименование" column there; `photo` matches ../photo/<name>.png.
"""
import json
import os

_HERE = os.path.dirname(__file__)
DATA = json.load(open(os.path.join(_HERE, '..', 'data.json')))
PHOTO = os.path.join(_HERE, '..', 'photo')

SCENARIOS = ["20'", "40'", 'LCL 17', 'LCL 34']
BASE = "40'"          # cheapest scenario — highlighted in every cost block


def photo(name):
    return os.path.join(PHOTO, name + '.png')


SUPPLIERS = [
    dict(
        id='singha',
        extra_section=dict(
            sheet='SINGHA KAMEDA - Retail+Bulk',
            title='Arare Norimaki — роздріб і BULK',
            scenario='Контейнер везе роздрібну фасовку разом із BULK 3 кг',
            products=[
                dict(key='Norimaki Original 42 g — RETAIL', unit='1 пакет 42 г',
                     title='Arare Norimaki\nOriginal', photo='sk_original',
                     badge='42 г · 24 шт/карт.',
                     desc='Хрусткі рисові палички арарe в обгортці з водорості норі. '
                     'Смак соєвого соусу, без гостроти.'),
                dict(key='Norimaki Wasabi 42 g — RETAIL', unit='1 пакет 42 г',
                     title='Arare Norimaki\nWasabi', photo='sk_wasabi',
                     badge='42 г · 24 шт/карт.',
                     desc='Ті самі рисові палички в норі, але з відчутною гостротою '
                     'японського васабі.'),
                dict(key='Norimaki Original 3 kg x2 — BULK', unit='1 пакет 3 кг',
                     title='Norimaki Original\nBULK 3 кг', photo='sk_original',
                     badge='3 кг · 2 пак/карт.',
                     desc='Той самий продукт у форматі 3 кг — для HoReCa та фасування '
                     'на місці. На фото — роздрібна фасовка.'),
                dict(key='Norimaki Wasabi 3 kg x2 — BULK', unit='1 пакет 3 кг',
                     title='Norimaki Wasabi\nBULK 3 кг', photo='sk_wasabi',
                     badge='3 кг · 2 пак/карт.',
                     desc='Васабі у форматі 3 кг — для HoReCa та фасування на місці. '
                     'На фото — роздрібна фасовка.'),
            ]),
        scenario='Контейнер везе тільки роздрібну фасовку',
        category='Рисові крекери арарe в норі · Retail',
        stats=[('2', 'SKU із розрахованою\nсобівартістю'),
               ('12', 'місяців термін\nпридатності'),
               ('3 626', 'картонів\nу 40′HC')],
        specs=[('Фасовка', '24 шт / картон'),
               ('Термін придатності', '12 місяців'),
               ('Місткість 20′', '1 485 картонів'),
               ('Місткість 40′HC', '3 626 картонів')],
        sheet='SINGHA KAMEDA - Retail',
        name='Singha Kameda\n(Thailand) Co., Ltd.',
        short='Singha Kameda',
        headline='Arare Norimaki — рисові крекери в норі',
        brand='Arare Norimaki',
        brand_mark='SINGHA\nKAMEDA',
        country='Таїланд',
        port='FOB Bangkok',
        summary='Спільне підприємство японської Kameda Seika — світового лідера ринку '
                'рисових крекерів — та тайської групи Singha. Виробництво в Таїланді '
                'за японськими стандартами.',
        products=[
            dict(key='Original Flavour', unit='1 пакет 42 г', title='Arare Norimaki\nOriginal',
                 photo='sk_original', badge='42 г · 24 шт/карт.',
                 desc='Хрусткі рисові палички арарe в обгортці з водорості норі. Смак '
                 'соєвого соусу, без гостроти.'),
            dict(key='Wasabi Flavour', unit='1 пакет 42 г', title='Arare Norimaki\nWasabi',
                 photo='sk_wasabi', badge='42 г · 24 шт/карт.',
                 desc='Ті самі рисові палички в норі, але з відчутною гостротою японського '
                 'васабі.'),
        ],
    ),
    dict(
        id='thainichi',
        extra_section=dict(
            sheet='Thai-Nichi - Retail+Bulk',
            title='Norimaki — роздріб і BULK',
            scenario='Контейнер везе роздрібну фасовку разом із BULK 2,5 кг',
            products=[
                dict(key='Mizuho Norimaki Original 50 g — RETAIL', unit='1 пакет 50 г',
                     title='Norimaki\nOriginal', photo='tn_original',
                     badge='50 г · 36 шт/карт.',
                     desc='Рисовий крекер із водорістю норі, смажений за японською '
                     'рецептурою. Пакет 50 г.'),
                dict(key='Mizuho Norimaki Wasabi 55 g — RETAIL', unit='1 пакет 55 г',
                     title='Norimaki\nWasabi', photo='tn_wasabi',
                     badge='55 г · 36 шт/карт.',
                     desc='Той самий крекер із додаванням васабі. Найбільший пакет '
                     'лінійки — 55 г.'),
                dict(key='Mizuho Norimaki Original 2,5 kg x2 — BULK',
                     unit='1 пакет 2,5 кг', title='Norimaki Original\nBULK 2,5 кг',
                     photo='tn_original', badge='2,5 кг · 2 пак/карт.',
                     desc='Той самий крекер у форматі 2,5 кг — для HoReCa та фасування '
                     'на місці. На фото — роздрібна фасовка.'),
                dict(key='Mizuho Norimaki Wasabi 2,5 kg x2 — BULK',
                     unit='1 пакет 2,5 кг', title='Norimaki Wasabi\nBULK 2,5 кг',
                     photo='tn_wasabi', badge='2,5 кг · 2 пак/карт.',
                     desc='Васабі у форматі 2,5 кг — для HoReCa та фасування на місці. '
                     'На фото — роздрібна фасовка.'),
            ]),
        scenario='Контейнер везе тільки роздрібну фасовку',
        category='Рисові крекери, чипси та протеїнові снеки · Retail',
        stats=[('2', 'SKU із розрахованою\nсобівартістю'),
               ('28', 'позицій\nу прайс-листі'),
               ('1 013', 'картонів\nу 20′')],
        specs=[('Фасовка', '12 – 96 шт / картон'),
               ('Термін придатності', '8 – 12 місяців'),
               ('Місткість 20′', '1 013 картонів'),
               ('MOQ', 'від 300 картонів')],
        sheet='Thai-Nichi - Retail',
        name='Thai-Nichi\nIndustries Co., Ltd.',
        short='Thai-Nichi',
        headline='Norimaki — рисові крекери в норі',
        brand='Mizuho · Norimaki',
        brand_mark='THAI\nNICHI',
        country='Таїланд',
        port='FOB Bangkok',
        summary='Тайсько-японський виробник рисових снеків із найширшим прайсом серед '
                'усіх постачальників — 28 позицій. Одна фабрика закриває і крекери, '
                'і чипси, і протеїнові снеки.',
        products=[
            dict(key='Norimaki Original', unit='1 пакет 50 г', title='Norimaki\nOriginal',
                 photo='tn_original', badge='50 г · 36 шт/карт.',
                 desc='Рисовий крекер із водорістю норі, смажений за японською рецептурою. '
                 'Пакет 50 г.'),
            dict(key='Norimaki Wasabi', unit='1 пакет 55 г', title='Norimaki\nWasabi',
                 photo='tn_wasabi', badge='55 г · 36 шт/карт.',
                 desc='Той самий крекер із додаванням васабі. Найбільший пакет лінійки — '
                 '55 г.'),
        ],
    ),
    dict(
        id='tmk',
        scenario='Розрахунок на контейнер із роздрібною фасовкою',
        logo='logo_kokiri',
        category='Снеки з водорості норі · Retail і дитячий формат',
        stats=[('10', 'SKU із розрахованою\nсобівартістю'),
               ('45–60', 'днів\nвиробництва'),
               ('3 431', 'картонів\nу 40′HQ')],
        specs=[('Термін виробництва', '45 – 60 днів'),
               ('Оплата', '50% депозит + 50% до ETD'),
               ('Місткість 20′', '425 – 1 499 картонів'),
               ('Місткість 40′HQ', '959 – 3 431 картонів')],
        sheet='TMK Thailand Co., Ltd -Retail',
        name='TMK (Thailand)\nCo., Ltd.',
        short='TMK · KOKIRI',
        headline='KOKIRI Wow — снеки з водорості норі',
        brand='KOKIRI Wow',
        brand_mark='KOKIRI',
        country='Таїланд',
        port='FOB Bangkok',
        summary='Тайський виробник снеків із водоростей під брендом KOKIRI Wow. '
                'Найнижча собівартість одиниці в каталозі — за рахунок мінімальної '
                'фасовки та дитячого позиціонування.',
        note='Собівартість наведено за 1 пакетик (inner pack), а не за коробку: '
             'Wow ROLL — 2,5 г, Double Roll — 5 г.',
        products=[
            dict(key='KOKIRI Wow ROLL - Original', unit='1 пакетик 2,5 г', title='Wow ROLL\nOriginal',
                 photo='tmk_roll_orig', badge='2,5 г · пакетик',
                 desc='Лист норі, скручений навколо рисової палички. Пакетик на один укус, '
                 'у коробці 10 штук.'),
            dict(key='KOKIRI Wow ROLL - Hot Spicy', unit='1 пакетик 2,5 г', title='Wow ROLL\nHot Spicy',
                 photo='tmk_roll_spicy', badge='2,5 г · пакетик',
                 desc='Той самий рулетик норі з гострою приправою. У Таїланді — '
                 'найпродаваніший смак лінійки.'),
            dict(key='KOKIRI Wow ROLL - Spicy Squid', unit='1 пакетик 2,5 г', title='Wow ROLL\nSpicy Squid',
                 photo='tmk_roll_squid', badge='2,5 г · пакетик',
                 desc='Рулетик норі з приправою зі смаком гострого кальмара. '
                 'Найспецифічніший смак трійки.'),
            dict(key='KOKIRI Wow DOUBLE ROLL - Original', unit='1 пакетик 5 г',
                 title='Wow DOUBLE ROLL\nOriginal', photo='tmk_dbl_orig',
                 badge='5 г · пакетик',
                 desc='Подвійний рулетик норі — порція 5 г замість 2,5 г. У коробці 5 '
                 'пакетиків.'),
            dict(key='KOKIRI Wow DOUBLE ROLL - Hot Spicy', unit='1 пакетик 5 г',
                 title='Wow DOUBLE ROLL\nHot Spicy', photo='tmk_dbl_spicy',
                 badge='5 г · пакетик',
                 desc='Подвійний рулетик норі з гострою приправою. Крок нагору від '
                 'базового формату.'),
            dict(key='KOKIRI WOW SEAWEED - Original', unit='1 пакет 12 г', title='Wow SEAWEED\nOriginal',
                 photo='tmk_sw_orig', badge='12 г · 36 пак/карт.',
                 desc='Смажені листи норі, приправлені та нарізані на порційні шматочки. '
                 'Пакет 12 г.'),
            dict(key='KOKIRI WOW SEAWEED - Spicy', unit='1 пакет 12 г', title='Wow SEAWEED\nSpicy',
                 photo='tmk_sw_spicy', badge='12 г · 36 пак/карт.',
                 desc='Те саме з гострою приправою. Термін придатності 7 місяців — '
                 'потребує швидкого обігу.'),
            dict(key='KOKIRI WOW SEAWEED - Squid', unit='1 пакет 12 г', title='Wow SEAWEED\nSquid',
                 photo='tmk_sw_squid', badge='12 г · 36 пак/карт.',
                 desc='Смажена норі зі смаком кальмара. Термін придатності також 7 '
                 'місяців.'),
            dict(key='KOKIRI WOW MINI - Original', unit='1 пакет 10 г', title='Wow MINI\nOriginal',
                 photo='tmk_mini_orig', badge='10 г · 36 пак/карт.',
                 desc='Найменший пакет лінійки — 10 г. Формат для прикасової зони та '
                 'дитячих наборів.'),
            dict(key='KOKIRI WOW MINI - Hot Spicy', unit='1 пакет 10 г', title='Wow MINI\nHot Spicy',
                 photo='tmk_mini_spicy', badge='10 г · 36 пак/карт.',
                 desc='Гострий варіант міні-пакета 10 г. Ставиться в пару з Original на '
                 'одній полиці.'),
        ],
    ),
    dict(
        id='zek',
        scenario='Розрахунок на контейнер із роздрібною фасовкою',
        tagline='HanJin Food Co. Ltd',
        category='Темпура, сендвічі та присипка з норі · Retail',
        stats=[('14', 'SKU із розрахованою\nсобівартістю'),
               ('3', 'продуктові\nнапрямки'),
               ('2 370', 'картонів\nу 40′HQ')],
        specs=[('Фасовка', '8 – 36 шт / картон'),
               ('Термін придатності', '9 – 12 місяців'),
               ('Місткість 20′', '350 – 4 000 картонів'),
               ('Місткість 40′HQ', '500 – 2 370 картонів')],
        sheet='ZEK -Retail',
        name='HanJin Food Co. Ltd\nбренд ZEK',
        short='HanJin · ZEK',
        headline='ZEK — снеки з водорості норі',
        brand='ZEK',
        brand_mark='ZEK',
        country='Китай',
        port='FOB Qingdao',
        summary='Найбільший за глибиною розрахунку постачальник: 14 позицій у трьох '
                'напрямках — темпура, сендвіч і присипка з норі. Єдиний постачальник '
                'з відвантаженням із Циндао.',
        products=[
            dict(key='ZEK TEMPURA SEAWEED - Corn 30g', unit='1 пакет 30 г', title='Tempura Seaweed\nCorn',
                 photo='zek_tempura_corn30', badge='30 г · 24 шт/карт.',
                 desc='Листи норі в хрусткому клярі темпура зі смаком солодкої кукурудзи. '
                 'Пакет 30 г.'),
            dict(key='ZEK TEMPURA SEAWEED - Wasabi 30g', unit='1 пакет 30 г', title='Tempura Seaweed\nWasabi',
                 photo='zek_tempura_wasabi30', badge='30 г · 24 шт/карт.',
                 desc='Норі в клярі темпура з васабі. Гострий варіант тієї самої лінійки, '
                 '30 г.'),
            dict(key='ZEK TEMPURA SEAWEED - Mala Crawfish 30g', unit='1 пакет 30 г',
                 title='Tempura Seaweed\nMala Crawfish', photo='zek_tempura_mala30',
                 badge='30 г · 24 шт/карт.',
                 desc='Норі в темпурі з приправою мала — гостро-пряний китайський смак '
                 'раків. 30 г.'),
            dict(key='ZEK TEMPURA SEAWEED - Corn 50g', unit='1 пакет 50 г',
                 title='Tempura Seaweed\nCorn · 50 г', photo='zek_tempura_corn50',
                 badge='50 г · 24 шт/карт.',
                 desc='Кукурудзяна темпура у великому пакеті 50 г. Ціна за грам нижча, ніж '
                 'у 30 г.'),
            dict(key='ZEK TEMPURA SEAWEED - Wasabi 50g', unit='1 пакет 50 г',
                 title='Tempura Seaweed\nWasabi · 50 г', photo='zek_tempura_wasabi50',
                 badge='50 г · 24 шт/карт.',
                 desc='Темпура з васабі у пакеті 50 г — формат для домашнього споживання.'),
            dict(key='ZEK TEMPURA SEAWEED - Mala Crawfish 50g', unit='1 пакет 50 г',
                 title='Tempura Seaweed\nMala · 50 г', photo='zek_tempura_mala50',
                 badge='50 г · 24 шт/карт.',
                 desc='Темпура мала у пакеті 50 г. Замикає лінійку великих фасовок.'),
            dict(key='ZEK SANDWICH SEAWEED - Meat Floss 25g', unit='1 пакет 25 г',
                 title='Sandwich Seaweed\nMeat Floss', photo='zek_sandwich_meat25',
                 badge='25 г · 30 шт/карт.',
                 desc='Два листи норі з прошарком сушеної м’ясної стружки між ними. Пакет '
                 '25 г.'),
            dict(key='ZEK SANDWICH SEAWEED - Sesame 25g', unit='1 пакет 25 г',
                 title='Sandwich Seaweed\nSesame', photo='zek_sandwich_sesame25',
                 badge='25 г · 30 шт/карт.',
                 desc='Два листи норі з кунжутним прошарком. Вегетаріанський варіант '
                 'сендвіча.'),
            dict(key='ZEK SEAWEED TOPPING - Chicken Floss 35g', unit='1 пакет 35 г',
                 title='Seaweed Topping\nChicken Floss', photo='zek_topping_chicken35',
                 badge='35 г · 24 шт/карт.',
                 desc='Подрібнена норі з курячою стружкою — присипка до рису та супів. '
                 'Збагачена DHA.'),
            dict(key='ZEK SEAWEED TOPPING - Vegetables 35g', unit='1 пакет 35 г',
                 title='Seaweed Topping\nVegetables', photo='zek_topping_veg35',
                 badge='35 г · 24 шт/карт.',
                 desc='Норі-присипка з овочами. Паковання орієнтоване на дитячу аудиторію.'),
            dict(key='ZEK SEAWEED TOPPING - Sesame 35g', unit='1 пакет 35 г',
                 title='Seaweed Topping\nSesame', photo='zek_topping_sesame35',
                 badge='35 г · 24 шт/карт.',
                 desc='Норі-присипка з кунжутом — базовий смак напрямку, без м’ясних '
                 'добавок.'),
            dict(key='ZEK SEAWEED TOPPING - Vegetables 70g', unit='1 пакет 70 г',
                 title='Seaweed Topping\nVegetables · 70 г', photo='zek_topping_veg70',
                 badge='70 г · 24 шт/карт.',
                 desc='Овочева присипка у великому пакеті 70 г. Розрахована на родину або '
                 'HoReCa.'),
            dict(key='ZEK SEAWEED TOPPING - Sesame 70g', unit='1 пакет 70 г',
                 title='Seaweed Topping\nSesame · 70 г', photo='zek_topping_sesame70',
                 badge='70 г · 24 шт/карт.',
                 desc='Кунжутна присипка у пакеті 70 г — удвічі більший обсяг за той самий '
                 'смак.'),
            dict(key='ZEK SEAWEED TOPPING - Chicken Floss 70g', unit='1 пакет 70 г',
                 title='Seaweed Topping\nChicken · 70 г',
                 photo='zek_topping_chicken70', badge='70 г · 24 шт/карт.',
                 desc='Присипка з курячою стружкою, 70 г. Найдорожча позиція в розрахунку '
                 'ZEK.'),
        ],
    ),
]
