
# ══ РОЗРАХУНКИ ════════════════════════════════════════════════════════════
import collections


def as8(items, brand, c):
    return [(im, brand, nm, g, lo, hi, ch, c) for im, nm, g, lo, hi, ch in items]


L_SIDE = as8(BENS, "Ben's Original", ORANGE)                    # гарнір, пауч
L_BOWL = as8(OTTOGI, "Ottogi", PLUM) + BOWL_MORE                # страва в чаші
L_UA = UA_RETORT                                                # страва, пауч UA
L_IMP = IMPORT_RETORT                                           # страва, пауч імпорт
ALL8 = L_SIDE + L_BOWL + L_UA + L_IMP
assert len(ALL8) == N_SKU

COUNTRY = {"Ben's Original": "ЄС", "Ottogi": "Корея", "Bibigo": "Корея",
           "Portion": "Україна", "Маркел": "Україна", "Верес": "Україна",
           "МАКРО": "Україна", "Ходорівський": "Україна", "М'ясниця": "Україна",
           "Clearspring": "ЄС", "Adventure Menu": "ЄС"}
BY_PACK = collections.defaultdict(list)
for _x in ALL8:
    BY_PACK[pack_of(_x[1])].append(_x)
PACK_ORDER = ["pouch", "cup"]
N_UA = sum(1 for x in ALL8 if COUNTRY[x[1]] == "Україна")
N_SILPO = sum(1 for x in ALL8 if "Сільпо" in x[6])
PER100 = sorted(x[4] / x[3] * 100 for x in ALL8)
MED100 = st.median(PER100)
MED_UNIT = st.median([x[4] for x in ALL8])
# Топ-3 грамажі за кількістю позицій
G_TOP = collections.Counter(x[3] for x in ALL8).most_common(3)

# ── Імпорт готового рису, CN 1904 90 10 «рис приготовлений» ───────────────
# Дзеркальна статистика: експорт ЄС→Україна (Eurostat Comext DS-045409)
# та експорт Азії→Україна (UN Comtrade). Дзеркало не має прогалин
# української митниці 2019–2021 рр.
IMPORT = [(2019, 99.1, 294), (2020, 204.4, 553), (2021, 204.0, 567), (2022, 127.6, 434),
          (2023, 125.8, 428), (2024, 111.0, 475), (2025, 179.4, 823)]
IMPORT_PARTNERS_2025 = [("Польща", 94.4, 239), ("Болгарія", 78.3, 527), ("Італія", 5.6, 46),
                        ("Корея", 2.0, 10), ("Китай", 1.0, 4)]
EUR_UAH, USD_UAH = 47.15, 41.71          # НБУ, середній за 2025 р.
T_EU, V_EU = IMPORT[-1][1], IMPORT[-1][2]
T_ASIA = 3.1
T_TOTAL = T_EU + T_ASIA
T_PREV = IMPORT[-2][1] + 3.5
CIF_UAH = (V_EU * EUR_UAH + 14.0 * USD_UAH) / 1000            # млн грн
AVG_PACK_G = 300                          # медіана маси в межах дослідження
PACKS_K = T_TOTAL * 1000_000 / AVG_PACK_G / 1000              # тис. упаковок на рік
RETAIL_LO, RETAIL_HI = CIF_UAH * 2.2, CIF_UAH * 2.8           # роздріб, млн грн
POP_M = 29.0
EU_PER_CAP_USD = 1.43                     # ринок instant rice ЄС / населення ЄС
UA_PER_CAP_USD = RETAIL_LO * 1e6 / USD_UAH / (POP_M * 1e6)
GAP = EU_PER_CAP_USD / UA_PER_CAP_USD
GROWTH = (T_TOTAL / T_PREV - 1) * 100
