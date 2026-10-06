# ══ МИТНА БАЗА: склад коду, 2025 проти 2026, імпортери ═════════════════════
import collections
from customs_calc import rows as C_ROWS, g as cg, seg as cseg, M as C_M, months as C_MONTHS

PERIOD = "січень 2025 – травень 2026"
PERIOD_S = "січ. 2025 – трав. 2026"
T = lambda kg: kg / 1000
KG = collections.Counter()
for r in C_ROWS:
    KG[cseg(r)] += cg(r, 'net_weight_kg') or 0
CODE_KG = sum(KG.values())
S1_KG = sum(v for k, v in KG.items() if k.startswith('S1'))
S2_KG = KG['S2_boilwater']
OTHER_KG = CODE_KG - KG['frozen'] - KG['chips'] - KG['dolma'] - KG['risotto_kit'] - S1_KG - S2_KG
COMP = [  # (підпис, кг, колір, пояснення)
    ("Заморожені суміші рис + овочі", KG['frozen'], SLATE, "Oerlemans, «Рудь», АТБ-маркет: гавайська суміш, паелья — у морозилку, не готова страва"),
    ("Рисові чіпси й снеки", KG['chips'], LGREY, "Ficosota «Livity»: солоний снек, не страва"),
    ("Долма в банці", KG['dolma'], ROSE, "Виноградне листя з рисом, Болгарія: інший формат і привід"),
    ("Різото «довести до готовності»", KG['risotto_kit'], NAVY_L, "Riso Gallo, Casa Rinaldi, Trevijano: рис сирий, варити 15–18 хв"),
    ("Інше: топінги, сирий рис, спорт-каша", OTHER_KG, MIST_D, "Посипки, промисловий рис, GymBeam, ттокпоккі, сублімат"),
    ("Сегмент 2 · рис під окріп", S2_KG, TEAL, "Китайський рис у чаші 144/174 г (тип Henan); 30 кг — тайський"),
    ("Сегмент 1 · готовий рис для розігріву", S1_KG, ORANGE, "Ben's (Mars), CJ Hetbahn, Clearspring, Haldiram RTE"),
]
CUST_FOOT = (f"Митна база (codexmb), вивантаження 28.09.2026, УКТ ЗЕД 1904901000, {PERIOD} (у базі немає 12–31 березня 2025). "
             "Класифікація — за описом кожної декларації. Вага нетто.")


def s_composition():
    s = new_slide()
    header(s, "ОБСЯГ РИНКУ · МИТНА БАЗА", "Скільки готового рису ввозять в Україну",
           f"Код 1904901000, {PERIOD}: {num(T(CODE_KG))} т, з них готовий рис для розігріву — лише {num(T(S1_KG), 2)} т.")
    px0, pw = M, 7.95
    rect(s, px0, 1.86, pw, 5.06, MIST, rounded=True, adj=0.03)
    text(s, px0 + 0.28, 2.0, 7.4, 0.26, "З ЧОГО СКЛАДАЄТЬСЯ КОД, ТОНН ЗА ПЕРІОД", size=10.5, bold=True, color=GREY)
    yy = 2.34
    mx = max(v for _, v, _, _ in COMP)
    for lb, kg, c, note in COMP:
        sh = kg / CODE_KG
        text(s, px0 + 0.28, yy, 5.6, 0.26, lb, size=12.5, bold=True, color=INK)
        text(s, px0 + 5.9, yy, 1.85, 0.26, f"{num(T(kg), 2 if T(kg) < 10 else 1)} т · {num(sh * 100, 1)} %", size=12.5,
             bold=True, color=NAVY, align=PP_ALIGN.RIGHT)
        rect(s, px0 + 0.28, yy + 0.28, 7.4, 0.16, WHITE, rounded=True, adj=0.5)
        rect(s, px0 + 0.28, yy + 0.28, max(7.4 * kg / mx, 0.06), 0.16, c, rounded=True, adj=0.5)
        text(s, px0 + 0.28, yy + 0.47, 7.4, 0.2, note, size=8.5, color=GREY)
        yy += 0.655
    rx = px0 + pw + 0.22; rw = W - M - rx
    rect(s, rx, 1.86, rw, 3.30, NAVY, rounded=True, adj=0.05); rect(s, rx, 1.86, 0.09, 3.30, ORANGE, rounded=True, adj=0.5)
    text(s, rx + 0.3, 2.02, rw - 0.5, 0.3, "Наш ринок: сегменти 1–2", size=13, bold=True, color=AMBER)
    text(s, rx + 0.3, 2.38, rw - 0.5, 0.7, f"{num(T(S1_KG + S2_KG), 1)} т", size=38, bold=True, color=WHITE)
    text(s, rx + 0.3, 3.12, rw - 0.5, 0.3, f"{num((S1_KG + S2_KG) / CODE_KG * 100, 1)} % коду за весь період",
         size=11, color=RGBColor(0xB9, 0xC2, 0xDA))
    for i, (lb, v) in enumerate([("Сегмент 1 · розігрів", f"{num(T(S1_KG), 2)} т"), ("Сегмент 2 · під окріп", f"{num(T(S2_KG), 2)} т"),
                                 ("Ben's — одна партія", "676 кг")]):
        text(s, rx + 0.3, 3.62 + i * 0.46, rw - 1.9, 0.3, lb, size=11.5, color=RGBColor(0xD5, 0xDB, 0xEA))
        text(s, rx + rw - 1.6, 3.62 + i * 0.46, 1.4, 0.3, v, size=13, bold=True, color=WHITE, align=PP_ALIGN.RIGHT)
    rect(s, rx, 5.30, rw, 1.62, WHITE, line=MIST_D, lw=1.2, rounded=True, adj=0.06)
    text(s, rx + 0.24, 5.42, rw - 0.4, 0.26, "ПЕРІОД І МЕТОД", size=9, bold=True, color=ROSE)
    text(s, rx + 0.24, 5.70, rw - 0.4, 1.2,
         "17 місяців: січень 2025 – травень 2026. У вивантаженні немає 12–31 березня 2025. "
         "Кожну декларацію віднесено до типу продукту за її описом. Показано вагу нетто, а не вартість.", size=10.5, color=INK, line=1.2)
    foot(s, CUST_FOOT + " Фактурна вартість USD без ПДВ і мита.")


# ── 2025 проти 2026: по типах продукту, середнє на місяць ──
def _year_kg():
    Y = collections.defaultdict(collections.Counter)
    for r in C_ROWS:
        sg = cseg(r); y = str(cg(r, 'report_month'))[:4]; kg = cg(r, 'net_weight_kg') or 0
        k = {'frozen': 'frozen', 'chips': 'chips', 'dolma': 'dolma', 'risotto_kit': 'risotto'}.get(sg, sg if sg.startswith('S') else 'other')
        if k.startswith('S1'): k = 'S1'
        Y[k][y] += kg
    return Y


YK = _year_kg()
MON = {'2025': 12, '2026': 5}


def _pm(k, y): return YK[k][y] / MON[y]


def s_dynamics():
    s = new_slide()
    header(s, "МИТНА БАЗА · 2025 І 2026", "Який рис ввозять і скільки",
           "Середнє на місяць: 2025 — 12 місяців, 2026 — січень–травень. Так порівняння чесне.")
    # ліва панель: масові типи, т/міс
    lx, lw = M, 7.35
    rect(s, lx, 1.80, lw, 5.14, MIST, rounded=True, adj=0.03)
    text(s, lx + 0.28, 1.92, lw - 0.5, 0.26, "ЩО ВВОЗЯТЬ МАСОВО · ТОНН НА МІСЯЦЬ", size=10.5, bold=True, color=GREY)
    for i, (nm, c, col) in enumerate([("2025", "2025", SLATE), ("2026 · січ–трав", "2026", ORANGE)]):
        rect(s, lx + 4.2 + i * 1.6, 1.97, 0.16, 0.16, col, rounded=True, adj=0.3)
        text(s, lx + 4.45 + i * 1.6, 1.94, 1.3, 0.24, nm, size=10, color=INK)
    types = [("Заморожені суміші", 'frozen', "рис + овочі, у морозилку"), ("Рисові чіпси", 'chips', "снек, не страва"),
             ("Долма в банці", 'dolma', "листя з рисом"), ("Різото-набори", 'risotto', "сирий рис, варити"),
             ("Інше", 'other', "топінги, сирий рис, каша")]
    mx = max(_pm(k, y) for _, k, _ in types for y in ('2025', '2026'))
    yy = 2.36; bx = lx + 0.28; bwmax = 5.3
    for nm, k, note in types:
        text(s, bx, yy, 3.5, 0.28, nm, size=13, bold=True, color=NAVY)
        text(s, bx + 2.45, yy + 0.03, 2.4, 0.24, note, size=9.5, color=GREY)
        for j, (y, col) in enumerate([('2025', SLATE), ('2026', ORANGE)]):
            v = _pm(k, y) / 1000
            ww = max(bwmax * (v * 1000) / mx, 0.05)
            rect(s, bx, yy + 0.34 + j * 0.30, ww, 0.24, col, rounded=True, adj=0.25)
            text(s, bx + ww + 0.1, yy + 0.33 + j * 0.30, 1.3, 0.26, f"{num(v, 1 if v >= 1 else 2)} т", size=11.5, bold=True, color=NAVY)
        yy += 0.90
    # права панель: наші сегменти, кг/міс
    rx = lx + lw + 0.22; rw = W - M - rx
    rect(s, rx, 1.80, rw, 3.26, WHITE, line=MIST_D, lw=1.2, rounded=True, adj=0.04)
    text(s, rx + 0.26, 1.92, rw - 0.4, 0.26, "НАШІ СЕГМЕНТИ · КГ НА МІСЯЦЬ", size=10.5, bold=True, color=GREY)
    ours = [("Сегмент 1 · готовий рис", 'S1', "розігрів у НВЧ"), ("Сегмент 2 · під окріп", 'S2_boilwater', "Henan-тип")]
    mx2 = max(_pm(k, y) for _, k, _ in ours for y in ('2025', '2026'))
    yy = 2.34
    for nm, k, note in ours:
        text(s, rx + 0.26, yy, rw - 0.4, 0.28, nm, size=12.5, bold=True, color=NAVY)
        text(s, rx + 0.26, yy + 0.27, rw - 0.4, 0.22, note, size=9.5, color=GREY)
        for j, (y, col) in enumerate([('2025', SLATE), ('2026', ORANGE)]):
            v = _pm(k, y)
            ww = max((rw - 1.9) * v / mx2, 0.05)
            rect(s, rx + 0.26, yy + 0.58 + j * 0.30, ww, 0.24, col, rounded=True, adj=0.25)
            text(s, rx + 0.26 + ww + 0.1, yy + 0.57 + j * 0.30, 1.4, 0.26, f"{num(v)} кг", size=11.5, bold=True, color=NAVY)
        yy += 1.38
    s12_25 = (YK['S1']['2025'] + YK['S2_boilwater']['2025']); s12_26 = (YK['S1']['2026'] + YK['S2_boilwater']['2026'])
    nb = s12_25 - 676
    insight(s, rx, 5.18, rw, 1.76, None,
            f"Сегменти 1–2: {num(T(s12_25), 2)} т у 2025 (з них 0,68 т — одна партія Ben's) і {num(T(s12_26), 2)} т за п'ять місяців 2026. "
            f"Без Ben's на місяць: {num(nb / 12)} кг → {num(s12_26 / 5)} кг, ×{num((s12_26 / 5) / (nb / 12), 1)}.", ORANGE, size=12)
    foot(s, CUST_FOOT + " Заморожені суміші в 2025 мали великі разові партії, тому середнє на місяць у 2026 нижче; тренд — за місяцями нестабільний.")


# ── імпортери: топ-10 за вагою ──
def _importers():
    imp = collections.defaultdict(lambda: dict(kg=0, seg=collections.Counter(), man=collections.Counter(), cty=collections.Counter(),
                                               n=0, months=set()))
    for r in C_ROWS:
        d = imp[cg(r, 'importer_display_name')]; kg = cg(r, 'net_weight_kg') or 0
        d['kg'] += kg; d['seg'][cseg(r)] += kg; d['n'] += 1; d['months'].add(cg(r, 'report_month'))
        d['man'][(cg(r, 'manufacturer_normalized') or '—')] += kg; d['cty'][cg(r, 'origin_country_normalized_ua') or '—'] += kg
    return imp


IMP_NAME = {'ТОВ "АТБ-маркет"': 'АТБ-маркет', 'ТОВ «Торгова фірма «Рудь»': 'ТФ «Рудь»', 'ТОВ "ФОЗЗІ КОММЕРЦ"': 'Фоззі Коммерц',
            'ТОВ "Метро Кеш Енд Кері Україна"': 'Метро Кеш енд Кері', 'ТОВ "ФУДКОМ"': 'Фудком', 'ТОВ "Таврія-В"': 'Таврія-В',
            'ТОВ ВІЧУНАЙ УКРАЇНА 03110 М КИЇВ ВУЛ СОЛ': 'Вічунай Україна', 'АТ "Рудь"': 'АТ «Рудь»',
            'ТОВ КОПІЙКА ЦЕНТР 65007 ОДЕСЬКА ОБЛ МІСТ': 'Копійка Центр', 'ТОВ "ВЕНДІНГ КІНГЗ"': 'Вендінг Кінгз'}
IMP_WHAT = {'АТБ-маркет': ("Заморожені суміші рис + овочі", "Oerlemans Foods", "Польща · ЄС"),
            'ТФ «Рудь»': ("Заморожені суміші рис + овочі", "Oerlemans Foods", "Польща"),
            'Фоззі Коммерц': ("Долма, різото, Ben's 676 кг", "Palirria, Mars Austria", "Греція · ЄС"),
            'Метро Кеш енд Кері': ("Рисові чіпси", "Ficosota Food", "ЄС (Болгарія)"),
            'Фудком': ("Рисові чіпси", "Ficosota Food", "ЄС (Болгарія)"),
            'Таврія-В': ("Рисові чіпси", "Ficosota Food", "ЄС (Болгарія)"),
            'Вічунай Україна': ("Заморожені суміші рис + овочі", "Oerlemans, Bonduelle", "Польща"),
            'АТ «Рудь»': ("Заморожені суміші рис + овочі", "Oerlemans Foods", "Польща"),
            'Копійка Центр': ("Рисові чіпси", "Ficosota Food", "ЄС (Болгарія)"),
            'Вендінг Кінгз': ("Рис у чаші під окріп (сегм. 2)", "не вказано в описі", "Китай")}


def s_importers():
    imp = _importers()
    tot = sum(d['kg'] for d in imp.values())
    top = sorted(imp.items(), key=lambda kv: -kv[1]['kg'])[:10]
    s = new_slide()
    header(s, "МИТНА БАЗА · ІМПОРТЕРИ", "Хто завозить",
           f"Топ-10 імпортерів коду за вагою, {PERIOD_S}. Усього в базі {len(imp)} імпортерів.")
    tw = 8.45
    cols = [("№", 0.45), ("ІМПОРТЕР", 1.95), ("ЩО ЗАВОЗИТЬ", 2.55), ("КРАЇНА", 1.2), ("ТОНН", 0.95), ("ЧАСТКА", 1.35)]
    xx = M
    for lb, w_ in cols:
        if lb == "ТОНН":
            text(s, xx - 0.05, 1.80, w_ - 0.10, 0.22, lb, size=9, bold=True, color=GREY, align=PP_ALIGN.RIGHT)
        elif lb == "ЧАСТКА":
            text(s, xx + 0.20, 1.80, w_, 0.22, lb, size=9, bold=True, color=GREY)
        else:
            text(s, xx + 0.08, 1.80, w_, 0.22, lb, size=9, bold=True, color=GREY)
        xx += w_
    y = 2.06; rh = 0.475; mxs = top[0][1]['kg']
    for i, (k, d) in enumerate(top):
        nm = next((v for kk, v in IMP_NAME.items() if k.startswith(kk[:20])), k[:22]); what, man, cty = IMP_WHAT[nm]
        rect(s, M, y, tw, rh - 0.04, MIST if i % 2 == 0 else WHITE, rounded=True, adj=0.2)
        text(s, M + 0.1, y + 0.11, 0.4, 0.26, str(i + 1), size=12, bold=True, color=GREY)
        text(s, M + 0.53, y + 0.10, 1.9, 0.28, nm, size=12.5, bold=True, color=NAVY)
        text(s, M + 2.48, y + 0.04, 2.55, 0.22, what, size=10.5, bold=True, color=INK)
        text(s, M + 2.48, y + 0.23, 2.55, 0.2, man, size=9, color=GREY)
        text(s, M + 5.03, y + 0.12, 1.2, 0.26, cty, size=10, color=INK)
        text(s, M + 6.23, y + 0.10, 0.85, 0.28, num(T(d['kg']), 1), size=12.5, bold=True, color=NAVY, align=PP_ALIGN.RIGHT)
        col = ORANGE if nm in ('Вендінг Кінгз', 'Фоззі Коммерц') else SLATE
        bw2 = max(0.8 * d['kg'] / mxs, 0.04)
        rect(s, M + 7.28, y + 0.13, bw2, 0.19, col, rounded=True, adj=0.4)
        text(s, M + 7.28 + bw2 + 0.07, y + 0.12, 0.7, 0.22, f"{num(d['kg'] / tot * 100, 1)} %", size=9.5, bold=True, color=GREY)
        y += rh
    top_sum = sum(d['kg'] for _, d in top)
    text(s, M + 0.1, y + 0.04, tw, 0.24, f"Топ-10 — {num(top_sum / tot * 100, 1)} % ваги коду; ще {len(imp) - 10} імпортерів — {num(T(tot - top_sum), 1)} т.",
         size=10.5, color=GREY)
    rx = M + tw + 0.22; rw = W - M - rx
    rect(s, rx, 1.80, rw, 5.14, NAVY, rounded=True, adj=0.04); rect(s, rx, 1.80, 0.09, 5.14, ORANGE, rounded=True, adj=0.5)
    text(s, rx + 0.3, 1.94, rw - 0.5, 0.3, "Наші сегменти 1–2", size=13, bold=True, color=AMBER)
    text(s, rx + 0.3, 2.28, rw - 0.5, 0.3, "імпортери, кг за період", size=10, color=RGBColor(0xB9, 0xC2, 0xDA))
    ours = [("Вендінг Кінгз", "Китай · під окріп", 2639), ("Фоззі Коммерц", "Ben's Original · одна партія", 676),
            ("Відус", "Hetbahn · CJ, Корея", 76), ("Бюро Він", "Clearspring · ЄС", 63),
            ("Тадж Махал", "Haldiram · Індія", 33), ("Азіяфудс", "рис, Таїланд", 30)]
    mxo = ours[0][2]; yy = 2.70
    for nm, nt, kg in ours:
        text(s, rx + 0.3, yy, rw - 1.5, 0.26, nm, size=11.5, bold=True, color=WHITE)
        text(s, rx + rw - 1.2, yy, 1.0, 0.26, num(kg), size=12.5, bold=True, color=AMBER, align=PP_ALIGN.RIGHT)
        text(s, rx + 0.3, yy + 0.25, rw - 0.5, 0.2, nt, size=9, color=RGBColor(0xB9, 0xC2, 0xDA))
        rect(s, rx + 0.3, yy + 0.50, max((rw - 0.6) * kg / mxo, 0.05), 0.07, ORANGE, rounded=True, adj=0.5)
        yy += 0.70
    text(s, rx + 0.3, yy + 0.02, rw - 0.5, 0.5, "Жоден не возить регулярно: 1–4 партії за 17 місяців.", size=10.5, bold=True, color=WHITE, line=1.15)
    foot(s, CUST_FOOT + " Частка — від ваги коду. Імпортер «Рудь» представлений двома юрособами (ТОВ і АТ).")
