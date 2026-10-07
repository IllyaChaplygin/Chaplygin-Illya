# -*- coding: utf-8 -*-
"""Ukrainian market findings for seasoned seaweed snacks / rice-cracker-norimaki,
gathered 2026-10 via open retail monitoring (marketplace listings, retailer
search). No customs-import database was available for this study — unlike the
earlier SIAS/rice projects, there is no "митна база" section here, and that
gap is stated on its own slide rather than silently skipped.

Every price below traces to a real, named listing found during the search.
Nothing is estimated or interpolated.
"""

METHOD_NOTE = (
    'На відміну від досліджень SIAS і готового рису, тут немає доступу до митної '
    'бази — обсяги імпорту категорії не показано. Дані нижче — відкритий моніторинг '
    'маркетплейсів і сайтів мереж, жовтень 2026.'
)

# Segment A: seasoned, ready-to-eat seaweed snacks (nori chips, seaweed rolls,
# tempura seaweed) — directly comparable to TMK/KOKIRI and HanJin/ZEK.
# Segment B: rice crackers wrapped in nori (arare/senbei) — comparable to
# Singha Kameda and Thai-Nichi. No listings were found for segment B at all.
COMPETITORS = [
    dict(brand='Tao Kae Noi', origin='Таїланд', product='Big Roll Classic',
         pack='18 г (6 × 3 г)', price_uah=154.85, price_was=163, per_g=154.85 / 18,
         channel='Маркетплейс (Prom.ua)', seller='«Смак 24»',
         note='Єдиний світовий бренд категорії, знайдений на українському ринку.'),
    dict(brand='AKURA', origin='імпорт, бренд сушi-аксесуарів', product='Nori Chips Original',
         pack='3 × 4,5 г (13,5 г)', price_uah=129, price_was=167, per_g=129 / 13.5,
         channel='Маркетплейс (Rozetka, продавець Rozetka)', seller='Rozetka',
         note='Той самий бренд, що й норі-листи для суші — вийшов і в снек-формат.'),
    dict(brand='AKURA', origin='імпорт, бренд сушi-аксесуарів', product='Nori Chips Kimchi',
         pack='3 × 4,5 г (13,5 г)', price_uah=89, price_was=167, per_g=89 / 13.5,
         channel='Маркетплейс (Rozetka, продавець Rozetka)', seller='Rozetka',
         note='Акційна ціна — майже вдвічі нижче базової 167 ₴.'),
    dict(brand='Seaweed Traditional', origin='AKURA · гурт/HoReCa', product='Опт / HoReCa',
         pack='72 × 4,5 г (324 г)', price_uah=1820, price_was=2600, per_g=1820 / 324,
         channel='B2B-опт (Prom.ua)', seller='«Exotic Food»',
         note='Гуртова/ресторанна ціна — довідково, не роздрібна полиця.'),
]

CHAINS_CHECKED = [
    dict(name='АТБ', found=False), dict(name='Сільпо', found=False),
    dict(name='Novus', found=False), dict(name='Varus', found=False),
    dict(name='Ашан', found=False), dict(name='Metro', found=False),
]

CHANNEL_SUMMARY = (
    'У каталогах шести перевірених мереж приправлений снек з водоростей не '
    'знайдено — там є лише листи норі для суші (інша категорія, інше '
    'використання). Категорія живе на маркетплейсах: Prom.ua, Rozetka '
    '(продавець-маркетплейс), Bigl.ua, Allo.ua.'
)

SEGMENT_B_NOTE = (
    'Рисові крекери в норі (арарe, сенбей — категорія Singha Kameda і '
    'Thai-Nichi) на українському ринку не знайдено жодного разу — ні в '
    'мережах, ні на маркетплейсах. Порожня ніша або засліплена зона пошуку: '
    'перевірити наживо на полицях варто, перш ніж вважати ринок порожнім.'
)
