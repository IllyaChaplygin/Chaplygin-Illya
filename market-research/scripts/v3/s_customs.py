# ══ МИТНА БАЗА: склад коду, 2026 проти 2025, імпортери ════════════════════
import collections
from customs_calc import rows as C_ROWS, g as cg, seg as cseg, M as C_M, months as C_MONTHS

T = lambda kg: kg / 1000
KG = collections.Counter()
for r in C_ROWS:
    KG[cseg(r)] += cg(r, 'net_weight_kg') or 0
CODE_KG = sum(KG.values())
S1_KG = sum(v for k, v in KG.items() if k.startswith('S1'))
S2_KG = KG['S2_boilwater']
OTHER_KG = CODE_KG - KG['frozen'] - KG['chips'] - KG['dolma'] - KG['risotto_kit'] - S1_KG - S2_KG
COMP = [  # (підпис, кг, колір, пояснення)
    ("Заморожені сирі суміші рис + овочі", KG['frozen'], SLATE, "Oerlemans, «Рудь», АТБ-маркет: різото, паелья, гавайська суміш — у морозилку"),
    ("Рисові чіпси й снеки", KG['chips'], LGREY, "Ficosota «Livity»: солоний снек, не страва"),
    ("Долма в банці", KG['dolma'], ROSE, "Виноградне листя з рисом, Болгарія — інший формат і привід"),
    ("Різото «довести до готовності»", KG['risotto_kit'], NAVY_L, "Riso Gallo, Casa Rinaldi, Trevijano: рис сирий, варити 15–18 хв"),
    ("Інше: топінги, сирий рис, спорт-каша", OTHER_KG, MIST_D, "Посипки, промисловий рис, GymBeam, ттокпоккі, сублімат"),
    ("Сегмент 2 · рис під окріп", S2_KG, TEAL, "Китайський рис у чаші 144/174 г (відповідає лінійці Henan)"),
    ("Сегмент 1 · готовий рис для розігріву", S1_KG, ORANGE, "Ben's (Mars), CJ, Clearspring, Haldiram RTE"),
]


def s_composition():
    s = new_slide()
    header(s, "ОБСЯГ РИНКУ · МИТНА БАЗА", "Скільки готового рису ввозять в Україну",
           f"Код 1904901000: {num(T(CODE_KG))} т за 16 місяців, з них готовий рис для розігріву — {num(T(S1_KG),2)} т.")
    px0, pw = M, 7.95
    rect(s, px0, 1.86, pw, 5.06, MIST, rounded=True, adj=0.03)
    text(s, px0 + 0.28, 2.0, 7.4, 0.26, "З ЧОГО СКЛАДАЄТЬСЯ КОД, ТОНН ЗА 16 МІСЯЦІВ", size=10.5, bold=True, color=GREY)
    yy = 2.34
    mx = max(v for _, v, _, _ in COMP)
    for lb, kg, c, note in COMP:
        sh = kg / CODE_KG
        text(s, px0 + 0.28, yy, 5.6, 0.26, lb, size=12.5, bold=True, color=INK)
        text(s, px0 + 5.9, yy, 1.85, 0.26, f"{num(T(kg),2 if T(kg) < 10 else 1)} т · {num(sh*100,1)} %", size=12.5,
             bold=True, color=NAVY, align=PP_ALIGN.RIGHT)
        rect(s, px0 + 0.28, yy + 0.28, 7.4, 0.16, WHITE, rounded=True, adj=0.5)
        rect(s, px0 + 0.28, yy + 0.28, max(7.4 * kg / mx, 0.06), 0.16, c, rounded=True, adj=0.5)
        text(s, px0 + 0.28, yy + 0.47, 7.4, 0.2, note, size=8.5, color=GREY)
        yy += 0.655
    rx = px0 + pw + 0.22; rw = W - M - rx
    rect(s, rx, 1.86, rw, 3.30, NAVY, rounded=True, adj=0.05); rect(s, rx, 1.86, 0.09, 3.30, ORANGE, rounded=True, adj=0.5)
    text(s, rx + 0.3, 2.02, rw - 0.5, 0.3, "Наш ринок: сегменти 1–2", size=13, bold=True, color=AMBER)
    text(s, rx + 0.3, 2.38, rw - 0.5, 0.7, f"{num(T(S1_KG + S2_KG),1)} т", size=38, bold=True, color=WHITE)
    text(s, rx + 0.3, 3.12, rw - 0.5, 0.3, f"{num((S1_KG+S2_KG)/CODE_KG*100,1)} % коду за 16 місяців",
         size=11, color=RGBColor(0xB9, 0xC2, 0xDA))
    for i, (lb, v) in enumerate([("Сегмент 1 · розігрів", f"{num(T(S1_KG),2)} т"), ("Сегмент 2 · під окріп", f"{num(T(S2_KG),2)} т"),
                                 ("Ben's — одна партія", "676 кг")]):
        text(s, rx + 0.3, 3.62 + i * 0.46, rw - 1.9, 0.3, lb, size=11.5, color=RGBColor(0xD5, 0xDB, 0xEA))
        text(s, rx + rw - 1.6, 3.62 + i * 0.46, 1.4, 0.3, v, size=13, bold=True, color=WHITE, align=PP_ALIGN.RIGHT)
    rect(s, rx, 5.30, rw, 1.62, WHITE, line=MIST_D, lw=1.2, rounded=True, adj=0.06)
    text(s, rx + 0.24, 5.42, rw - 0.4, 0.26, "ЧОМУ НЕ 182 Т З EUROSTAT", size=9, bold=True, color=ROSE)
    text(s, rx + 0.24, 5.70, rw - 0.4, 1.2,
         "182 т — експорт ЄС→Україна за тим самим кодом. Три чверті — заморожені суміші, тому як обсяг "
         "готового рису ця цифра некоректна. Рахуємо за описом кожної декларації.", size=9.8, color=INK, line=1.2)
    foot(s, "Митна база (codexmb), вивантаження 28.09.2026, УКТ ЗЕД 1904901000, 01.2025–05.2026 (без 12–31.03.2025). "
            "Класифікація — за описом кожної декларації. Фактурна вартість USD без ПДВ і мита.")


def month_series():
    out = {}
    for m in C_MONTHS:
        d = C_M[m]
        out[m] = d
    return out


def s_dynamics():
    s = new_slide()
    header(s, "МИТНА БАЗА · 2026", "2026 проти 2025",
           "Увесь код зріс на 36 % завдяки замороженим сумішам. Наші сегменти — ×3 без разової партії Ben's.")
    ms = month_series()
    LAB = ['Січ', 'Лют', 'Бер', 'Кві', 'Тра', 'Чер', 'Лип', 'Сер', 'Вер', 'Жов', 'Лис', 'Гру', 'Січ', 'Лют', 'Бер', 'Кві', 'Тра']
    cx0, cw = M, 8.25
    # --- верхній графік: весь код, т
    rect(s, cx0, 1.80, cw, 2.46, MIST, rounded=True, adj=0.04)
    text(s, cx0 + 0.25, 1.88, 7.5, 0.26, "ВЕСЬ КОД ПО МІСЯЦЯХ, ТОНН", size=10.5, bold=True, color=GREY)
    layers = [('Заморожені суміші', SLATE), ('Снеки', LGREY), ('Інше', MIST_D), ('S1', ORANGE), ('S2', TEAL)]
    base_y, ph = 3.85, 1.38
    mx = max(sum(ms[m].values()) for m in C_MONTHS) / 1000
    bw = (cw - 0.7) / 17
    for i, m in enumerate(C_MONTHS):
        x = cx0 + 0.3 + i * bw
        y = base_y
        for key, col in layers:
            v = ms[m].get(key, 0) / 1000
            hh = ph * v / mx
            if hh > 0.004:
                rect(s, x + 0.04, y - hh, bw - 0.08, hh, col); y -= hh
        tot = sum(ms[m].values()) / 1000
        if tot > 20:
            text(s, x - 0.1, y - 0.22, bw + 0.2, 0.2, num(tot, 1), size=9, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
        text(s, x - 0.05, base_y + 0.04, bw + 0.1, 0.2, LAB[i], size=8.5, color=GREY, align=PP_ALIGN.CENTER)
    text(s, cx0 + 0.3, base_y + 0.24, 1.0, 0.2, "2025", size=9, bold=True, color=INK)
    text(s, cx0 + 0.3 + 12 * bw, base_y + 0.24, 1.0, 0.2, "2026", size=9, bold=True, color=INK)
    lx = cx0 + 3.55
    for nm, col in [('Заморожені суміші', SLATE), ('Снеки', LGREY), ('Інше', MIST_D), ('Наші сегменти 1–2', ORANGE)]:
        rect(s, lx, 1.93, 0.13, 0.13, col, rounded=True, adj=0.3)
        text(s, lx + 0.18, 1.90, 1.8, 0.2, nm, size=8.5, color=INK); lx += 0.42 + 0.062 * len(nm)
    # --- нижній графік: сегменти 1–2 по місяцях, кг
    rect(s, cx0, 4.36, cw, 2.58, MIST, rounded=True, adj=0.04)
    text(s, cx0 + 0.25, 4.44, 7.5, 0.26, "СЕГМЕНТИ 1–2 ПО МІСЯЦЯХ, КГ", size=10.5, bold=True, color=GREY)
    byb = collections.defaultdict(lambda: collections.Counter())
    for r in C_ROWS:
        sg = cseg(r); mth = cg(r, 'report_month'); kg = cg(r, 'net_weight_kg') or 0
        if sg == 'S1_bens': byb[mth]["Ben's (Mars)"] += kg
        elif sg == 'S2_boilwater': byb[mth]['Китай · під окріп'] += kg
        elif sg == 'S1_cj': byb[mth]['CJ'] += kg
        elif sg in ('S1_clearspring', 'S1_haldiram'): byb[mth]['Clearspring, Haldiram'] += kg
    BR = [("Ben's (Mars)", GOLD), ('Китай · під окріп', TEAL), ('CJ', GREEN), ('Clearspring, Haldiram', PLUM)]
    base2, ph2 = 6.25, 1.25
    mx2 = max(sum(byb[m].values()) for m in C_MONTHS)
    for i, m in enumerate(C_MONTHS):
        x = cx0 + 0.3 + i * bw
        y = base2
        for key, col in BR:
            v = byb[m].get(key, 0)
            hh = ph2 * v / mx2
            if hh > 0.004:
                rect(s, x + 0.04, y - hh, bw - 0.08, hh, col); y -= hh
        tot = sum(byb[m].values())
        if tot > 100:
            text(s, x - 0.15, y - 0.21, bw + 0.3, 0.2, num(tot), size=9, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
        text(s, x - 0.05, base2 + 0.04, bw + 0.1, 0.2, LAB[i], size=8.5, color=GREY, align=PP_ALIGN.CENTER)
    lx = cx0 + 0.3
    for nm, col in BR:
        rect(s, lx, base2 + 0.31, 0.13, 0.13, col, rounded=True, adj=0.3)
        text(s, lx + 0.18, base2 + 0.28, 2.2, 0.2, nm, size=8.5, color=INK); lx += 0.42 + 0.075 * len(nm)
    # --- праворуч: порівняння однакових місяців
    kx = cx0 + cw + 0.22; kw = W - M - kx
    cards = [("ВЕСЬ КОД · СІЧ, ЛЮТ, КВІ, ТРА", "27,2 → 37,0 т", "+36 %", SLATE,
              "$ за кг: 4,67 → 3,73 (−20 %)"),
             ("ЗАМОРОЖЕНІ СУМІШІ", "13,2 → 26,1 т", "×2,0", ROSE, "АТБ-маркет і «Рудь» — драйвер росту"),
             ("НАШІ СЕГМЕНТИ 1–2", "1,11 → 1,30 т", "+17 %", ORANGE,
              "Без разової партії Ben's (лют. 2025): 0,43 → 1,30 т, ×3,0"),
             ("ІНСТАНТ-ЛОКШИНА · 1902301000", "240,6 → 457,9 т", "+90 %", TEAL,
              "Суміжна категорія: масштаб, на який може вирости рис")]
    ch = 1.20
    for i, (lb, val, pct, c, nt) in enumerate(cards):
        y = 1.80 + i * (ch + 0.115)
        rect(s, kx, y, kw, ch, WHITE, line=MIST_D, lw=1.2, rounded=True, adj=0.07)
        rect(s, kx, y, 0.09, ch, c, rounded=True, adj=0.5)
        text(s, kx + 0.26, y + 0.1, kw - 1.0, 0.2, lb, size=8.3, bold=True, color=GREY)
        text(s, kx + 0.26, y + 0.33, kw - 1.2, 0.4, val, size=17, bold=True, color=NAVY)
        text(s, kx + kw - 1.1, y + 0.33, 0.95, 0.4, pct, size=17, bold=True, color=c, align=PP_ALIGN.RIGHT)
        text(s, kx + 0.26, y + 0.8, kw - 0.4, 0.42, nt, size=9, color=GREY, line=1.12)
    foot(s, "Митна база (codexmb), 28.09.2026. Порівняння — однакові місяці 2025 і 2026 (січень, лютий, квітень, травень): "
            "березня 2025 у вивантаженні неповний. Сегменти 1–2 — готовий рис для розігріву й рис під окріп.")


def s_importers():
    s = new_slide()
    header(s, "МИТНА БАЗА · ІМПОРТЕРИ", "Хто завозить",
           "П'ять основних імпортерів сегментів 1–2, і жоден не возить регулярно.")
    # таблиця імпортерів
    cols = [("ІМПОРТЕР", 2.75), ("ЩО ЗАВОЗИТЬ · ВИРОБНИК", 3.05), ("КРАЇНА", 1.2), ("КГ ЗА 16 МІС.", 1.2),
            ("$ ЗА КГ", 0.95), ("ПЕРІОД", 2.83)]
    R = [("Фоззі Коммерц (Fozzy Group)", "Ben's Original · Mars Austria", "ЄС / Британія", 676, 6.28, "лют. 2025 · 1 партія"),
         ("Вендінг Кінгз", "Рис у чаші 144/174 г (тип Henan)", "Китай", 2639, 2.79, "лют. 2025 – кві. 2026 · 4 партії"),
         ("Відус", "Hetbahn · CJ CheilJedang", "Корея", 76, 3.69, "тра. 2025 – кві. 2026 · 3 партії"),
         ("Бюро Він", "Clearspring · органічний рис", "ЄС", 63, 9.71, "січ. – лис. 2025 · 4 партії"),
         ("Тадж Махал Інтернешнл", "Haldiram · RTE Biryani", "Індія", 33, 2.79, "лис. 2025 · 1 партія")]
    x0 = M; y0 = 1.80
    xx = x0
    for lb, w in cols:
        text(s, xx + 0.12, y0, w - 0.1, 0.24, lb, size=9, bold=True, color=GREY); xx += w
    y = y0 + 0.30
    for i, row in enumerate(R):
        rect(s, x0, y, 11.98, 0.50, MIST if i % 2 == 0 else WHITE, rounded=True, adj=0.2)
        xx = x0
        for j, ((lb, w), val) in enumerate(zip(cols, row)):
            fmt = val if isinstance(val, str) else (num(val) if j == 3 else f"${val:.2f}")
            text(s, xx + 0.12, y + 0.12, w - 0.2, 0.3, fmt, size=11 if j else 11.5, bold=(j in (0, 3, 4)),
                 color=NAVY if j in (0, 3) else INK, line=1.0)
            xx += w
        y += 0.52
    # графік $/кг
    cy0 = y + 0.12
    rect(s, M, cy0, 11.98, 6.98 - cy0, MIST, rounded=True, adj=0.03)
    text(s, M + 0.26, cy0 + 0.1, 11.0, 0.26, "ЦІНА ЗА КГ: МИТНА ВАРТІСТЬ КОНКУРЕНТІВ ПРОТИ НАШОГО FOB, $", size=10.5, bold=True, color=GREY)
    BARS = [("Clearspring", 9.71, SLATE), ("Ben's", 6.28, SLATE), ("CJ Hetbahn", 3.69, SLATE),
            ("Китай · чаша", 2.79, SLATE),
            ("BSCM\nстакан 150", 0.500 / 0.150, ORANGE), ("BSCM\nпауч 150", 0.472 / 0.150, ORANGE),
            ("BSCM\nпауч 200", 0.550 / 0.200, ORANGE), ("BSCM\nпауч 240", 0.583 / 0.240, ORANGE),
            ("CM Premium\n250", 0.667 / 0.250, ORANGE), ("ONE'S\nбілий 210", 0.571 / 0.210, ORANGE),
            ("ONE'S fried\n200", 1.238 / 0.200, ORANGE)]
    base = 6.35; ph = 6.35 - (cy0 + 0.75)
    bw = (11.98 - 0.6) / len(BARS)
    for i, (lb, v, c) in enumerate(BARS):
        x = M + 0.3 + i * bw
        hh = ph * v / 10
        rect(s, x + 0.1, base - hh, bw - 0.2, hh, c, rounded=True, adj=0.08)
        text(s, x, base - hh - 0.26, bw, 0.24, f"${v:.2f}", size=11.5, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
        text(s, x, base + 0.05, bw, 0.4, lb, size=9, color=INK, align=PP_ALIGN.CENTER, line=1.0)
    lx = M + 8.0
    rect(s, lx, cy0 + 0.13, 0.14, 0.14, SLATE, rounded=True, adj=0.3); text(s, lx + 0.2, cy0 + 0.1, 1.9, 0.2, "митна декларація", size=9, color=INK)
    rect(s, lx + 1.9, cy0 + 0.13, 0.14, 0.14, ORANGE, rounded=True, adj=0.3); text(s, lx + 2.1, cy0 + 0.1, 1.7, 0.2, "наш FOB", size=9, color=INK)
    foot(s, "Митна база: фактурна вартість USD ÷ вага нетто, 01.2025–05.2026; ще «Азіяфудс» — 30 кг тайського рису (лют. 2025). Наш FOB — комерційні пропозиції BSCM, CM Premium, "
            "One's (для ONE'S білий і fried — оцінка із собівартості) ÷ вага упаковки. Рис у чаші (Китай) — сухий продукт під окріп.")
