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
        sheet='SINGHA KAMEDA - Retail',
        name='Singha Kameda\n(Thailand) Co., Ltd.',
        short='Singha Kameda',
        brand='Arare Norimaki',
        brand_mark='SINGHA\nKAMEDA',
        country='Таїланд',
        port='FOB Bangkok',
        summary='Спільне підприємство японської Kameda Seika — світового лідера ринку '
                'рисових крекерів — та тайської групи Singha. Виробництво в Таїланді '
                'за японськими стандартами.',
        products=[
            dict(key='Original Flavour', title='Arare Norimaki\nOriginal',
                 photo='sk_original', badge='42 г · 24 шт/карт.',
                 desc='Японські рисові крекери арарe у хрусткій водорості норі. '
                      'Класичний смак соєвого соусу.'),
            dict(key='Wasabi Flavour', title='Arare Norimaki\nWasabi',
                 photo='sk_wasabi', badge='42 г · 24 шт/карт.',
                 desc='Ті самі крекери арарe в норі, але з пікантною ноткою '
                      'японського васабі.'),
        ],
    ),
    dict(
        id='thainichi',
        sheet='Thai-Nichi - Retail',
        name='Thai-Nichi\nIndustries Co., Ltd.',
        short='Thai-Nichi',
        brand='Mizuho · Norimaki',
        brand_mark='THAI\nNICHI',
        country='Таїланд',
        port='FOB Bangkok',
        summary='Тайсько-японський виробник рисових снеків із найширшим прайсом серед '
                'усіх постачальників — 28 позицій. Одна фабрика закриває і крекери, '
                'і чипси, і протеїнові снеки.',
        products=[
            dict(key='Norimaki Original', title='Norimaki\nOriginal',
                 photo='tn_original', badge='50 г · 36 шт/карт.',
                 desc='Рисовий крекер у норі за японською рецептурою. Хрусткий, '
                      'з делікатним смаком соєвого соусу.'),
            dict(key='Norimaki Wasabi', title='Norimaki\nWasabi',
                 photo='tn_wasabi', badge='55 г · 36 шт/карт.',
                 desc='Рисовий крекер у норі з гострим васабі. Найбільша фасовка '
                      'лінійки — 55 г.'),
        ],
    ),
    dict(
        id='tmk',
        sheet='TMK Thailand Co., Ltd -Retail',
        name='TMK (Thailand)\nCo., Ltd.',
        short='TMK · KOKIRI',
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
            dict(key='KOKIRI Wow ROLL - Original', title='Wow ROLL\nOriginal',
                 photo='tmk_roll_orig', badge='2,5 г · пакетик',
                 desc='Рулетик із водорості норі з рисовим крекером усередині. '
                      'У коробці 10 пакетиків.'),
            dict(key='KOKIRI Wow ROLL - Hot Spicy', title='Wow ROLL\nHot Spicy',
                 photo='tmk_roll_spicy', badge='2,5 г · пакетик',
                 desc='Той самий рулетик норі з гострою приправою. Найпопулярніший '
                      'смак лінійки.'),
            dict(key='KOKIRI Wow ROLL - Spicy Squid', title='Wow ROLL\nSpicy Squid',
                 photo='tmk_roll_squid', badge='2,5 г · пакетик',
                 desc='Рулетик норі зі смаком гострого кальмара — азійський смаковий '
                      'профіль.'),
            dict(key='KOKIRI Wow DOUBLE ROLL - Original',
                 title='Wow DOUBLE ROLL\nOriginal', photo='tmk_dbl_orig',
                 badge='5 г · пакетик',
                 desc='Подвійна порція рулетиків норі. У коробці 5 пакетиків.'),
            dict(key='KOKIRI Wow DOUBLE ROLL - Hot Spicy',
                 title='Wow DOUBLE ROLL\nHot Spicy', photo='tmk_dbl_spicy',
                 badge='5 г · пакетик',
                 desc='Подвійна порція рулетиків норі з гострою приправою.'),
            dict(key='KOKIRI WOW SEAWEED - Original', title='Wow SEAWEED\nOriginal',
                 photo='tmk_sw_orig', badge='12 г · 36 пак/карт.',
                 desc='Смажені листи норі, приправлені та нарізані на порційні '
                      'шматочки.'),
            dict(key='KOKIRI WOW SEAWEED - Spicy', title='Wow SEAWEED\nSpicy',
                 photo='tmk_sw_spicy', badge='12 г · 36 пак/карт.',
                 desc='Смажена норі з гострою приправою. Термін придатності '
                      '7 місяців.'),
            dict(key='KOKIRI WOW SEAWEED - Squid', title='Wow SEAWEED\nSquid',
                 photo='tmk_sw_squid', badge='12 г · 36 пак/карт.',
                 desc='Смажена норі зі смаком кальмара — позиція для спеціалізованої '
                      'роздрібу.'),
            dict(key='KOKIRI WOW MINI - Original', title='Wow MINI\nOriginal',
                 photo='tmk_mini_orig', badge='10 г · 36 пак/карт.',
                 desc='Міні-формат смаженої норі для імпульсної покупки біля каси.'),
            dict(key='KOKIRI WOW MINI - Hot Spicy', title='Wow MINI\nHot Spicy',
                 photo='tmk_mini_spicy', badge='10 г · 36 пак/карт.',
                 desc='Гострий міні-формат норі. Пара до Original для викладки '
                      'у прикасовій зоні.'),
        ],
    ),
    dict(
        id='zek',
        sheet='ZEK -Retail',
        name='HanJin Food Co. Ltd\nбренд ZEK',
        short='HanJin · ZEK',
        brand='ZEK',
        brand_mark='ZEK',
        country='Китай',
        port='FOB Qingdao',
        summary='Найбільший за глибиною розрахунку постачальник: 14 позицій у трьох '
                'напрямках — темпура, сендвіч і присипка з норі. Єдиний постачальник '
                'з відвантаженням із Циндао.',
        products=[
            dict(key='ZEK TEMPURA SEAWEED - Corn 30g', title='Tempura Seaweed\nCorn',
                 photo='zek_tempura_corn30', badge='30 г · 24 шт/карт.',
                 desc='Листи норі у хрусткому клярі темпура зі смаком солодкої '
                      'кукурудзи.'),
            dict(key='ZEK TEMPURA SEAWEED - Wasabi 30g', title='Tempura Seaweed\nWasabi',
                 photo='zek_tempura_wasabi30', badge='30 г · 24 шт/карт.',
                 desc='Норі у клярі темпура з японським васабі. Гострий, впізнаваний '
                      'смак.'),
            dict(key='ZEK TEMPURA SEAWEED - Mala Crawfish 30g',
                 title='Tempura Seaweed\nMala Crawfish', photo='zek_tempura_mala30',
                 badge='30 г · 24 шт/карт.',
                 desc='Норі у темпурі зі смаком раків мала — трендовий китайський '
                      'смак.'),
            dict(key='ZEK TEMPURA SEAWEED - Corn 50g',
                 title='Tempura Seaweed\nCorn · 50 г', photo='zek_tempura_corn50',
                 badge='50 г · 24 шт/карт.',
                 desc='Сімейна фасовка кукурудзяної темпури. Найкраща ціна за грам '
                      'у напрямку.'),
            dict(key='ZEK TEMPURA SEAWEED - Wasabi 50g',
                 title='Tempura Seaweed\nWasabi · 50 г', photo='zek_tempura_wasabi50',
                 badge='50 г · 24 шт/карт.',
                 desc='Сімейна фасовка темпури з васабі — формат для домашнього '
                      'споживання.'),
            dict(key='ZEK TEMPURA SEAWEED - Mala Crawfish 50g',
                 title='Tempura Seaweed\nMala · 50 г', photo='zek_tempura_mala50',
                 badge='50 г · 24 шт/карт.',
                 desc='Сімейна фасовка темпури мала. Замикає лінійку 50 г з трьох '
                      'смаків.'),
            dict(key='ZEK SANDWICH SEAWEED - Meat Floss 25g',
                 title='Sandwich Seaweed\nMeat Floss', photo='zek_sandwich_meat25',
                 badge='25 г · 30 шт/карт.',
                 desc='Сендвіч із двох листів норі з прошарком м’ясної стружки.'),
            dict(key='ZEK SANDWICH SEAWEED - Sesame 25g',
                 title='Sandwich Seaweed\nSesame', photo='zek_sandwich_sesame25',
                 badge='25 г · 30 шт/карт.',
                 desc='Сендвіч норі з кунжутним прошарком. Вегетаріанська '
                      'альтернатива.'),
            dict(key='ZEK SEAWEED TOPPING - Chicken Floss 35g',
                 title='Seaweed Topping\nChicken Floss', photo='zek_topping_chicken35',
                 badge='35 г · 24 шт/карт.',
                 desc='Норі-присипка до рису та страв із курячою стружкою. Збагачена '
                      'DHA.'),
            dict(key='ZEK SEAWEED TOPPING - Vegetables 35g',
                 title='Seaweed Topping\nVegetables', photo='zek_topping_veg35',
                 badge='35 г · 24 шт/карт.',
                 desc='Норі-присипка з овочами. Дитячий дизайн паковання.'),
            dict(key='ZEK SEAWEED TOPPING - Sesame 35g',
                 title='Seaweed Topping\nSesame', photo='zek_topping_sesame35',
                 badge='35 г · 24 шт/карт.',
                 desc='Норі-присипка з кунжутом. Базовий смак напрямку.'),
            dict(key='ZEK SEAWEED TOPPING - Vegetables 70g',
                 title='Seaweed Topping\nVegetables · 70 г', photo='zek_topping_veg70',
                 badge='70 г · 24 шт/карт.',
                 desc='Велика фасовка овочевої присипки для родинного споживання.'),
            dict(key='ZEK SEAWEED TOPPING - Sesame 70g',
                 title='Seaweed Topping\nSesame · 70 г', photo='zek_topping_sesame70',
                 badge='70 г · 24 шт/карт.',
                 desc='Велика фасовка кунжутної присипки — формат для родин '
                      'та HoReCa.'),
            dict(key='ZEK SEAWEED TOPPING - Chicken Floss 70g',
                 title='Seaweed Topping\nChicken · 70 г',
                 photo='zek_topping_chicken70', badge='70 г · 24 шт/карт.',
                 desc='Велика фасовка присипки з курячою стружкою. Замикає лінійку '
                      '70 г.'),
        ],
    ),
]

# Fifth supplier: present in the price comparison, but its SelfCost tab is empty.
KORIKO = dict(
    short='Nature Best Food Co., Ltd.',
    brand='KoriKo',
    summary='Вкладка «Koriko» у файлі Self Cost порожня — собівартість ще не '
            'розрахована. Нижче асортимент і ціни FOB як основа для розрахунку.',
    items=[
        ('kk_chewy', 'KoriKo Chewy\nOriginal', '4,5 г × 12 пакетів / постер · '
         '12 постерів у картоні', '$1,15 / постер'),
        ('kk_chewy_spicy', 'KoriKo Chewy\nHot & Spicy', '4,5 г × 12 пакетів / постер · '
         '12 постерів у картоні', '$1,15 / постер'),
        ('kk_bigsheet', 'Big Sheet 50 г\nOriginal / Squid', '5 г × 10 пакетів · '
         '12 наборів у картоні', '$1,30 / набір'),
        ('kk_roll', 'Roll Seaweed 36 г\nOriginal / Spicy', '10 рулетів у коробці · '
         '12 коробок у картоні', '$1,20 / коробка'),
        ('kk_sandwich', 'Sandwich Seaweed\n5 смаків', 'Original, Kimchi, Wasabi… · '
         '24 пакети в картоні', '$1,03 / пакет'),
        ('kk_fizze', 'Fizze Sheet 72 г\n3 смаки', '6 пакетів у коробці · '
         '8 коробок у картоні', '$1,30 / коробка'),
    ],
)
