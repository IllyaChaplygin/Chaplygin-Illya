# -*- coding: utf-8 -*-
"""Supplier / product catalogue.

Self-cost figures come from SelfCost.xlsx via data.json (never typed by hand).
`key` matches the "Наименование" column there; `photo` matches /home/user/work/photo.
"""
import json
import os

DATA = json.load(open(os.path.join(os.path.dirname(__file__), '..', 'data.json')))
PHOTO = os.path.join(os.path.dirname(__file__), '..', 'photo')

SCENARIOS = ["20'", "40'", 'LCL 17', 'LCL 34']
SCENARIO_LABEL = {
    "20'": "20′ контейнер",
    "40'": "40′ контейнер",
    'LCL 17': 'Збірний 17 м³',
    'LCL 34': 'Збірний 34 м³',
}
BASE = "40'"          # headline scenario shown on product cards


def photo(name):
    return os.path.join(PHOTO, name + '.png')


SUPPLIERS = [
    dict(
        id='singha',
        sheet='SINGHA KAMEDA - Retail',
        name='Singha Kameda (Thailand) Co., Ltd.',
        short='Singha Kameda',
        brand='Arare Norimaki · Okome Snacks',
        country='Таїланд',
        port='FOB Bangkok',
        logo_line='SINGHA\nKAMEDA',
        summary=('Спільне підприємство японської Kameda Seika — світового лідера ринку '
                 'рисових крекерів — та тайської групи Singha (Boon Rawd Brewery). '
                 'Виробництво в Таїланді за японськими стандартами якості.'),
        stats=[('2', 'SKU з розрахованою\nсобівартістю'), ('9', 'позицій\nу прайс-листі'),
               ('12', 'місяців\nтермін придатності')],
        facts=[('Країна', 'Таїланд'), ('Умови', 'FOB Bangkok'),
               ('Фасовка', '24 шт / картон'), ('Місткість 20′', '1 485 картонів'),
               ('Місткість 40′HC', '3 626 картонів'), ('Термін придатності', '12 місяців')],
        extra_title='Решта асортименту (собівартість ще не розрахована)',
        extra=['Arare Norimaki — Squid Chili, 42 г · FOB $0,77',
               'Arare Norimaki — Mala, 42 г · FOB $0,77',
               'Okome Snacks — Japanese Mix, 85 г · FOB $0,80',
               'Okome Snacks — Edo Mix, 85 г · FOB $0,80',
               'Okome Snacks — Cheesy Corn, 60 г · FOB $0,98',
               'Okome Snacks — Honey Shoyu, 40 г / 100 г · FOB $0,63 / $1,30'],
        extra_photo='sk_okome',
        products=[
            dict(key='Original Flavour', title='Arare Norimaki\nOriginal', photo='sk_original',
                 badge='42 г · пакет', specs='24 шт/карт. · FOB $0,77',
                 desc='Японські рисові крекери арарe, загорнуті у хрустку водорість норі. '
                      'Класичний смак соєвого соусу — база лінійки.'),
            dict(key='Wasabi Flavour', title='Arare Norimaki\nWasabi', photo='sk_wasabi',
                 badge='42 г · пакет', specs='24 шт/карт. · FOB $0,77',
                 desc='Ті самі крекери арарe в норі, але з пікантною ноткою японського '
                      'васабі. Найдинамічніший смак категорії.'),
        ],
    ),
    dict(
        id='thainichi',
        sheet='Thai-Nichi - Retail',
        name='Thai-Nichi Industries Co., Ltd.',
        short='Thai-Nichi',
        brand='Mizuho · Norimaki',
        country='Таїланд',
        port='FOB Bangkok',
        logo_line='THAI\nNICHI',
        summary=('Тайсько-японський виробник рисових снеків із найширшим прайсом серед '
                 'усіх постачальників — 28 позицій, з них 11 нових. Одна фабрика '
                 'закриває і рисові крекери, і чипси, і протеїнові снеки.'),
        stats=[('2', 'SKU з розрахованою\nсобівартістю'), ('28', 'позицій\nу прайс-листі'),
               ('11', 'нових SKU\n(позначені NEW)')],
        facts=[('Країна', 'Таїланд'), ('Умови', 'FOB Bangkok'),
               ('Фасовка', '36 шт / картон'), ('Місткість 20′', '1 013 картонів'),
               ('Місткість 40′HQ', '2 370 картонів'), ('MOQ', 'від 300 картонів')],
        extra_title='Ключові напрямки прайсу (собівартість ще не розрахована)',
        extra=['Norimaki California Roll / Cheese Ring · FOB $0,60 – $0,72',
               'Kakinotane та Kakinotane Mala, 50 г · FOB $0,36 – $0,43',
               'Hokkaido Milk / Matcha Cream, 35 г · FOB $0,49 – $0,51',
               'Rice & Beans Chips — 5 смаків, 90 г · FOB $0,63 – $0,78',
               'Soy Protein Rice Cracker — 3 смаки, 50 г · FOB $0,61',
               'Міні-формат Norimaki 15 г для каси · FOB $0,23'],
        extra_photo='tn_chips',
        products=[
            dict(key='Norimaki Original', title='Norimaki\nOriginal', photo='tn_original',
                 badge='50 г · пакет', specs='36 шт/карт. · FOB $0,60',
                 desc='Рисовий крекер у норі за японською рецептурою. Хрусткий, '
                      'з делікатним смаком соєвого соусу.'),
            dict(key='Norimaki Wasabi', title='Norimaki\nWasabi', photo='tn_wasabi',
                 badge='55 г · пакет', specs='36 шт/карт. · FOB $0,75',
                 desc='Рисовий крекер у норі з гострим васабі. Найбільша фасовка '
                      'лінійки — 55 г.'),
        ],
    ),
    dict(
        id='tmk',
        sheet='TMK Thailand Co., Ltd -Retail',
        name='TMK (Thailand) Co., Ltd.',
        short='TMK · KOKIRI',
        brand='KOKIRI Wow',
        country='Таїланд',
        port='FOB Bangkok',
        logo_line='KOKIRI',
        summary=('Тайський виробник снеків із водоростей під брендом KOKIRI Wow. '
                 'Найдешевша собівартість одиниці серед усіх постачальників — за рахунок '
                 'мінімальної фасовки (від 2,5 г) та дитячого позиціонування.'),
        stats=[('10', 'SKU з розрахованою\nсобівартістю'), ('9', 'лінійок\nу прайс-листі'),
               ('7,61 ₴', 'найнижча собівартість\nодиниці у деці')],
        facts=[('Країна', 'Таїланд'), ('Бренд', 'KOKIRI Wow'),
               ('Умови', 'FOB Bangkok'), ('Термін придатності', '7 – 12 місяців'),
               ('MOQ', '20′ контейнер'),
               ('Увага', 'собівартість — за 1 пакетик')],
        extra_title='Решта асортименту (собівартість ще не розрахована)',
        extra=['KOKIRI Wow Double Roll (Costco), 110 г · $4,84 / $4,18 промо',
               'KOKIRI Wow Seaweed, 30 г · $1,032 / $0,875 промо',
               'KOKIRI Wow Booga Chips, 20 г — 3 смаки · $0,528 / $0,447 промо',
               'KOKIRI Wow 7 Roll, 14 г · $0,634 / $0,540 промо',
               'KOKIRI Wow Sheet, 24 г · $1,901 / $1,620 промо'],
        extra_photo='tmk_line',
        note=('Собівартість наведено за 1 пакетик (inner pack), а не за коробку: '
              'Wow ROLL — 2,5 г, Double Roll — 5 г.'),
        products=[
            dict(key='KOKIRI Wow ROLL - Original', title='Wow ROLL\nOriginal',
                 photo='tmk_roll_orig', badge='2,5 г · пакетик',
                 specs='120 пак./карт. · FOB $0,122',
                 desc='Рулетик із водорості норі з рисовим крекером усередині. '
                      'У коробці 10 пакетиків — формат для дітей.'),
            dict(key='KOKIRI Wow ROLL - Hot Spicy', title='Wow ROLL\nHot Spicy',
                 photo='tmk_roll_spicy', badge='2,5 г · пакетик',
                 specs='120 пак./карт. · FOB $0,122',
                 desc='Той самий рулетик норі, але з гострою приправою. '
                      'Найпопулярніший смак лінійки у Таїланді.'),
            dict(key='KOKIRI Wow ROLL - Spicy Squid', title='Wow ROLL\nSpicy Squid',
                 photo='tmk_roll_squid', badge='2,5 г · пакетик',
                 specs='120 пак./карт. · FOB $0,122',
                 desc='Рулетик норі зі смаком гострого кальмара — азійський '
                      'смаковий профіль для сміливої аудиторії.'),
            dict(key='KOKIRI Wow DOUBLE ROLL - Original', title='Wow DOUBLE ROLL\nOriginal',
                 photo='tmk_dbl_orig', badge='5 г · пакетик',
                 specs='60 пак./карт. · FOB $0,218',
                 desc='Подвійна порція рулетиків норі у пакетику 5 г. '
                      'У коробці 5 пакетиків.'),
            dict(key='KOKIRI Wow DOUBLE ROLL - Hot Spicy', title='Wow DOUBLE ROLL\nHot Spicy',
                 photo='tmk_dbl_spicy', badge='5 г · пакетик',
                 specs='60 пак./карт. · FOB $0,218',
                 desc='Подвійна порція рулетиків норі з гострою приправою. '
                      'Крок нагору від базового формату 2,5 г.'),
            dict(key='KOKIRI WOW SEAWEED - Original', title='Wow SEAWEED\nOriginal',
                 photo='tmk_sw_orig', badge='12 г · пакет',
                 specs='36 пак./карт. · FOB $0,449',
                 desc='Смажені листи норі, приправлені та нарізані на порційні '
                      'шматочки. Класичний смак.'),
            dict(key='KOKIRI WOW SEAWEED - Spicy', title='Wow SEAWEED\nSpicy',
                 photo='tmk_sw_spicy', badge='12 г · пакет',
                 specs='36 пак./карт. · FOB $0,449',
                 desc='Смажена норі з гострою приправою. Термін придатності '
                      '7 місяців — планувати обіг.'),
            dict(key='KOKIRI WOW SEAWEED - Squid', title='Wow SEAWEED\nSquid',
                 photo='tmk_sw_squid', badge='12 г · пакет',
                 specs='36 пак./карт. · FOB $0,449',
                 desc='Смажена норі зі смаком кальмара. Найвужча ніша лінійки, '
                      'працює у спеціалізованій рознице.'),
            dict(key='KOKIRI WOW MINI - Original', title='Wow MINI\nOriginal',
                 photo='tmk_mini_orig', badge='10 г · пакет',
                 specs='36 пак./карт. · FOB $0,446',
                 desc='Міні-формат смаженої норі для імпульсної покупки біля каси. '
                      'Термін придатності 12 місяців.'),
            dict(key='KOKIRI WOW MINI - Hot Spicy', title='Wow MINI\nHot Spicy',
                 photo='tmk_mini_spicy', badge='10 г · пакет',
                 specs='36 пак./карт. · FOB $0,446',
                 desc='Гострий міні-формат норі. Пара до Original для викладки '
                      'у прикасовій зоні.'),
        ],
    ),
    dict(
        id='zek',
        sheet='ZEK -Retail',
        name='HanJin Food Co. Ltd — бренд ZEK',
        short='HanJin · ZEK',
        brand='ZEK',
        country='Китай',
        port='FOB Qingdao',
        logo_line='ZEK',
        summary=('Найбільший за глибиною розрахунку постачальник: 14 SKU з порахованою '
                 'собівартістю у трьох продуктових напрямках — темпура, сендвіч і '
                 'присипка з норі. Єдиний постачальник з відвантаженням із Циндао.'),
        stats=[('14', 'SKU з розрахованою\nсобівартістю'), ('35', 'позицій\nу прайс-листі'),
               ('3', 'продуктові\nнапрямки')],
        facts=[('Країна', 'Китай'), ('Умови', 'FOB Qingdao'),
               ('Напрямки', 'Tempura · Sandwich · Topping'),
               ('Фасовка', '12 – 30 шт / картон'),
               ('Термін придатності', '9 – 12 місяців'),
               ('Особливість', 'збагачення DHA та вітамінами')],
        extra_title='Решта асортименту (собівартість ще не розрахована)',
        extra=['ZEK Original / Bamboo Salt / Olive Oil Seaweed, 15 – 60 г · $0,52 – $2,09',
               'ZEK Seasoned Seaweed with Prebiotics, 21 г · $0,60',
               'ZEK Seaweed Topping, 48 г та 100 г · $0,57 – $1,16',
               'ZEK Sandwich Seaweed, 40 г та 60 г · $0,75 – $0,97'],
        extra_photo='zek_classic',
        products=[
            dict(key='ZEK TEMPURA SEAWEED - Corn 30g', title='Tempura Seaweed\nCorn',
                 photo='zek_tempura_corn30', badge='30 г · пакет',
                 specs='24 шт/карт. · FOB $0,43',
                 desc='Листи норі у хрусткому клярі темпура зі смаком солодкої '
                      'кукурудзи. Найм’якший смак напрямку.'),
            dict(key='ZEK TEMPURA SEAWEED - Wasabi 30g', title='Tempura Seaweed\nWasabi',
                 photo='zek_tempura_wasabi30', badge='30 г · пакет',
                 specs='24 шт/карт. · FOB $0,43',
                 desc='Норі у клярі темпура з японським васабі. Гострий, '
                      'впізнаваний смаковий профіль.'),
            dict(key='ZEK TEMPURA SEAWEED - Mala Crawfish 30g', title='Tempura Seaweed\nMala Crawfish',
                 photo='zek_tempura_mala30', badge='30 г · пакет',
                 specs='24 шт/карт. · FOB $0,43',
                 desc='Норі у темпурі зі смаком раків мала — трендовий китайський '
                      'смак для молодої аудиторії.'),
            dict(key='ZEK TEMPURA SEAWEED - Corn 50g', title='Tempura Seaweed\nCorn · 50 г',
                 photo='zek_tempura_corn50', badge='50 г · пакет',
                 specs='24 шт/карт. · FOB $0,65',
                 desc='Сімейна фасовка кукурудзяної темпури. Кращий показник '
                      'ціни за грам у напрямку.'),
            dict(key='ZEK TEMPURA SEAWEED - Wasabi 50g', title='Tempura Seaweed\nWasabi · 50 г',
                 photo='zek_tempura_wasabi50', badge='50 г · пакет',
                 specs='24 шт/карт. · FOB $0,65',
                 desc='Сімейна фасовка темпури з васабі — формат для домашнього '
                      'споживання та подарункових наборів.'),
            dict(key='ZEK TEMPURA SEAWEED - Mala Crawfish 50g', title='Tempura Seaweed\nMala · 50 г',
                 photo='zek_tempura_mala50', badge='50 г · пакет',
                 specs='24 шт/карт. · FOB $0,65',
                 desc='Сімейна фасовка темпури мала. Замикає лінійку 50 г '
                      'з трьох смаків.'),
            dict(key='ZEK SANDWICH SEAWEED - Meat Floss 25g', title='Sandwich Seaweed\nMeat Floss',
                 photo='zek_sandwich_meat25', badge='25 г · пакет',
                 specs='30 шт/карт. · FOB $0,31',
                 desc='Сендвіч із двох листів норі з прошарком м’ясної стружки. '
                      'Найнижчий FOB серед ZEK.'),
            dict(key='ZEK SANDWICH SEAWEED - Sesame 25g', title='Sandwich Seaweed\nSesame',
                 photo='zek_sandwich_sesame25', badge='25 г · пакет',
                 specs='30 шт/карт. · FOB $0,31',
                 desc='Сендвіч норі з кунжутним прошарком. Вегетаріанська '
                      'альтернатива м’ясному варіанту.'),
            dict(key='ZEK SEAWEED TOPPING - Chicken Floss 35g', title='Seaweed Topping\nChicken Floss',
                 photo='zek_topping_chicken35', badge='35 г · пакет',
                 specs='24 шт/карт. · FOB $0,40',
                 desc='Подрібнена норі-присипка до рису та страв із курячою '
                      'стружкою. Збагачена DHA та вітамінами.'),
            dict(key='ZEK SEAWEED TOPPING - Vegetables 35g', title='Seaweed Topping\nVegetables',
                 photo='zek_topping_veg35', badge='35 г · пакет',
                 specs='24 шт/карт. · FOB $0,40',
                 desc='Норі-присипка з овочами. Дитячий дизайн паковання — '
                      'позиція для полиці «здорове харчування».'),
            dict(key='ZEK SEAWEED TOPPING - Sesame 35g', title='Seaweed Topping\nSesame',
                 photo='zek_topping_sesame35', badge='35 г · пакет',
                 specs='24 шт/карт. · FOB $0,40',
                 desc='Норі-присипка з кунжутом. Базовий смак напрямку '
                      'для щоденного вжитку.'),
            dict(key='ZEK SEAWEED TOPPING - Vegetables 70g', title='Seaweed Topping\nVegetables · 70 г',
                 photo='zek_topping_veg70', badge='70 г · пакет',
                 specs='24 шт/карт. · FOB $0,81',
                 desc='Велика фасовка овочевої присипки. Найдорожча позиція '
                      'у розрахунку ZEK.'),
            dict(key='ZEK SEAWEED TOPPING - Sesame 70g', title='Seaweed Topping\nSesame · 70 г',
                 photo='zek_topping_sesame70', badge='70 г · пакет',
                 specs='24 шт/карт. · FOB $0,81',
                 desc='Велика фасовка кунжутної присипки — формат для родин '
                      'та HoReCa.'),
            dict(key='ZEK SEAWEED TOPPING - Chicken Floss 70g', title='Seaweed Topping\nChicken · 70 г',
                 photo='zek_topping_chicken70', badge='70 г · пакет',
                 specs='24 шт/карт. · FOB $0,81',
                 desc='Велика фасовка присипки з курячою стружкою. Замикає '
                      'лінійку 70 г.'),
        ],
    ),
]

# Fifth supplier: present in the price comparison, but its SelfCost tab is empty.
KORIKO = dict(
    name='Nature Best Food Co., Ltd.',
    short='Nature Best · KoriKo',
    brand='KoriKo',
    country='Таїланд',
    port='FOB Bangkok',
    summary=('Постачальник присутній у файлі Price Comparison (24 позиції), але вкладка '
             '«Koriko» у файлі Self Cost порожня — розрахунок собівартості ще не '
             'виконано. Нижче наведено асортимент і ціни FOB як основу для '
             'майбутнього розрахунку.'),
    items=[
        ('kk_chewy', 'KoriKo Chewy — Original', '4,5 г × 12 пакетів / постер',
         '12 постерів/карт.', '$1,15 / постер'),
        ('kk_chewy_spicy', 'KoriKo Chewy — Hot & Spicy', '4,5 г × 12 пакетів / постер',
         '12 постерів/карт.', '$1,15 / постер'),
        ('kk_bigsheet', 'Big Sheet 50 г — Original / Squid', '5 г × 10 пакетів',
         '12 наборів/карт.', '$1,30 / набір'),
        ('kk_roll', 'Roll Seaweed 36 г — Original / Spicy', '10 рулетів / коробка',
         '12 коробок/карт.', '$1,20 / коробка'),
        ('kk_sandwich', 'Sandwich Seaweed — 5 смаків', '24 пакети / картон',
         'Original, Kimchi, Wasabi…', '$1,03 / пакет'),
        ('kk_fizze', 'Fizze Sheet 72 г — 3 смаки', '6 пакетів / коробка',
         '8 коробок/карт.', '$1,30 / коробка'),
    ],
)
