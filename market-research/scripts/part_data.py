
# ══ ПРЕДМЕТ ДОСЛІДЖЕННЯ ═══════════════════════════════════════════════════
# Тільки готовий рис, який зберігається за кімнатної температури і
# розігрівається — мікрохвильовка 60–120 с або занурення пауча в окріп.
# НЕ входять: рис під заливання окропом, саморозігрів, сублімація,
# заморозка, охолоджені страви, суха крупа й варильні пакети.
#
# Кожна ціна — роздрібна за ОДНУ упаковку, перевірена на сайті продавця
# 23 вересня 2026 р. Оптові пороги (від 3–20 шт) до розрахунку не входять.

# (файл пакшоту, назва, грамів, ціна_від, ціна_до, канал)
# ── Гарнір: чистий рис у паучі, мікрохвильовка 90 с ───────────────────────
BENS = [
    ("md/bens_longgrain.jpg",     "Long Grain",        250, 45, 45,       "Сільпо"),
    ("md/bens_basmati250.jpg",    "Basmati",           250, 89.99, 89.99, "Сільпо"),
    ("md/bens_mediterran.jpg",    "Mediterran",        250, 89.99, 89.99, "Сільпо"),
    ("md/bens_risibisi.jpg",      "Risi Bisi",         250, 89.99, 89.99, "Сільпо"),
    ("md/bens_curry_indien.jpg",  "Indian Curry",      250, 89.99, 89.99, "Сільпо"),
    ("md/bens_curry_linsen.jpg",  "Curry з сочевицею", 220, 89.99, 89.99, "Сільпо"),
    ("md/bens_mexikanisch.jpg",   "Mexikanisch",       220, 109, 109,     "Сільпо"),
    ("md/bens_curryreis.jpg",     "Curryreis Indien",  220, 144, 144,     "Сільпо"),
    ("md/bens_basmati220.jpg",    "Basmati",           220, 144, 144,     "Сільпо"),
    ("md/bens_sweetchili.jpg",    "Sweet Chili",       220, 179, 179,     "Сільпо"),
    ("sku/bens_lang220_single.jpg", "Long Grain",      220, 139, 139,     "Edison Lee"),
]
# ── Страва в чаші: рис із наповнювачем, мікрохвильовка ────────────────────
OTTOGI = [
    # 뚝배기불고기밥: у «Тайякі» — «бульгогі», в «Апетітаріум» — «гостра яловичина».
    # Обидві картки — те саме фото пачки. Одна позиція.
    ("sku/ot_bulgogi.jpg",     "Бульгогі",            320, 252, 356, "Тайякі · Апетітаріум"),
    # На пачці — свинина з каракатицею, рис 180 г + соус 130 г = 310 г. Три продавці
    # ведуть її як три різні позиції («свинина» 310 г, «гостра свинина» 269 г,
    # «з соусом з осьминога» 280 г) — на всіх трьох картках те саме фото. Одна позиція.
    ("sku/ot_pork310.jpg",     "Свинина з восьминогом", 310, 252, 334,
     "Тайякі · Gurm. · Апетіт."),
    ("sku/ot_chicken_rib.jpg", "Гострі курячі ребра", 310, 260, 289, "Gurmissimo · Апетітаріум"),
    ("sku/ot_tuna.jpg",        "Тунець і майонез",    247, 225, 356, "Тайякі · Gurm. · Апетіт."),
    ("sku/ot_bibimbap.jpg",    "Пібімпаб",            269, 252, 334, "Тайякі · Апетітаріум"),
]
# (файл, бренд, назва, грамів, від, до, канал, колір)
BOWL_MORE = [
    ("md/bibigo_bowl.jpg", "Bibigo", "Білий рис", 210, 189, 189, "Тайякі Март", GREEN),
]
# ── Страва в реторт-паучі, українське виробництво ─────────────────────────
UA_RETORT = [
    ("sku/ua_portion_pack.jpg", "Portion", "Каша рисова з м'ясом курки", 350, 71.4, 85,
     "portion.com.ua · Prom", ROSE),
    ("sku/ua_myasnytsia.jpg", "М'ясниця", "Свинина 35 %, горошок, кукурудза", 350, 88, 88,
     "Belorfoods", TEAL),
    ("sku/ua_markel_soy.jpg", "Маркел", "Рис із соєвим м'ясом", 350, 80, 96,
     "UP Shop · СУХПАЙ · Мартел-shop", SLATE),
    ("sku/ua_pak_pork.jpg", "МАКРО", "Свинина та овочі", 350, 96.1, 135,
     "UP Shop · СУХПАЙ · Ліхтар", ORANGE),
    ("sku/ua_veres_pork.jpg", "Верес", "Свинина, горошок, кукурудза", 350, 100, 100,
     "МореПродуктів", GREEN),
    ("sku/ua_makro_chicken.jpg", "МАКРО", "Курятина 35 % і овочі", 350, 101, 101,
     "СУХПАЙ", ORANGE),
    ("sku/ua_makro_beef.jpg", "МАКРО", "Яловичина й солодкий перець", 350, 103.85, 107,
     "UP Shop · СУХПАЙ", ORANGE),
    ("sku/ua_khodoriv_350.jpg", "Ходорівський", "Свинина та овочі", 350, 105, 105,
     "Молочний склад", PLUM),
    ("sku/ua_veres_chicken.jpg", "Верес", "Каша рисова з куркою", 350, 108, 134,
     "Козуб · Смачна адреса · Продукт-Shop", GREEN),
    ("sku/ua_markel_eat.jpg", "Маркел", "Рис із рослинним фаршем", 350, 127, 144,
     "UP Shop · СУХПАЙ", SLATE),
    ("sku/ua_khodoriv_pork.jpg", "Ходорівський", "Свинина, горошок, кукурудза", 350, 140, 199,
     "Foodi Shop · Food Shop · МореПродуктів", PLUM),
]
# ── Страва в реторт-паучі, імпорт ─────────────────────────────────────────
IMPORT_RETORT = [
    ("md/clearspring.jpg", "Clearspring", "Brown & Wild Rice, тамарі", 250, 446, 446,
     "Скарби Азії", TEAL),
    ("sku/am_wild.jpg", "Adventure Menu", "Курка в томатному соусі", 400, 315, 357,
     "Freeride · MK-Sport · Лєєр", GREEN),
    ("sku/am_meatballs.jpg", "Adventure Menu", "Тефтелі з басматі", 400, 336, 466,
     "Freeride · OXO · Лєєр · MK · Palmer", GREEN),
    ("sku/am_korma400.jpg", "Adventure Menu", "Chicken Korma з рисом", 400, 376, 404,
     "ALANTUR · Modern Shop · MK-Sport", GREEN),
]

N_SKU = len(BENS) + len(OTTOGI) + len(BOWL_MORE) + len(UA_RETORT) + len(IMPORT_RETORT)
BRANDS = (["Ben's Original", "Ottogi"] +
          [b for _, b, *_ in BOWL_MORE + UA_RETORT + IMPORT_RETORT])
N_BRANDS = len(dict.fromkeys(BRANDS))

# ── Тип паковання ─────────────────────────────────────────────────────────
PACK_BY_BRAND = {"Ben's Original": "pouch", "Ottogi": "cup", "Bibigo": "cup",
                 "Portion": "pouch", "Маркел": "pouch", "Верес": "pouch",
                 "МАКРО": "pouch", "Ходорівський": "pouch", "М'ясниця": "pouch",
                 "Clearspring": "pouch", "Adventure Menu": "pouch"}
PACK_RU = {"pouch": "ПАУЧ · РЕТОРТ", "cup": "ЧАША · СТАКАН"}


def pack_of(brand, name=None, grams=None):
    return PACK_BY_BRAND[brand]


# ── Суміжна полиця мереж: те саме споживання, але бляшанка ────────────────
CHAIN_SHELF = [
    ("sku/ch_hapay_rice.jpg", "hapay!", "Каша рисова зі свининою", 340, 62.9, 62.9, "Ашан", GREEN),
    ("sku/ch_lappetit_pork.jpg", "L'appetit", "Каша рисова зі свининою", 340, 119.9, 119.9,
     "Ашан", ORANGE),
    ("sku/ch_lappetit_beef.jpg", "L'appetit", "Каша рисова з яловичиною", 340, 125.9, 125.9,
     "Ашан", ORANGE),
    ("sku/ch_hapay_plov.jpg", "hapay!", "Плов з качки та булгуру", 340, 117, 117, "Ашан", GREEN),
    ("sku/ch_foodfabrika.jpg", "Food Fabrika", "Плов з куркою", 250, 118.9, 118.9, "Восторг", PLUM),
    ("sku/ch_myastoria.jpg", "М'ясторія", "Плов з куркою та родзинками", 350, 147.3, 160,
     "Novus · МегаМаркет · Космос", ROSE),
]
