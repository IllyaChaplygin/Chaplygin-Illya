# ══ МИТНА БАЗА: тільки готовий рис (сегменти 1–2) ═════════════════════════════
import collections
from customs_calc import rows as C_ROWS, g as cg, seg as cseg

PERIOD = "січень 2025 – травень 2026"
PERIOD_S = "січ. 2025 – трав. 2026"
T = lambda kg: kg / 1000
KG = collections.Counter()
for r in C_ROWS:
    KG[cseg(r)] += cg(r, 'net_weight_kg') or 0
CODE_KG = sum(KG.values())
S1_KG = sum(v for k, v in KG.items() if k.startswith('S1'))
S2_KG = KG['S2_boilwater']
RICE_KG = S1_KG + S2_KG                      # готовий рис: розігрів + під окріп
CUST_FOOT = (f"Митна база (codexmb), вивантаження 28.09.2026, УКТ ЗЕД 1904901000, {PERIOD} (у базі немає 12–31 березня 2025). "
             "Лише готовий рис сегментів 1–2; класифікація — за описом кожної декларації. Вага нетто.")

# ключ → (назва, країна, колір, імпортер, примітка)
PROD = [("S2_vend", "Рис у чаші під окріп (тип Henan)", "Китай", TEAL, "Вендінг Кінгз", "чаша 144/174 г, сегмент 2"),
        ("S1_bens", "Ben's Original", "ЄС · Британія", GOLD, "Фоззі Коммерц", "Mars Austria, пауч, сегмент 1"),
        ("S1_cj", "CJ Hetbahn", "Корея", GREEN, "Відус", "чаша 210 г, сегмент 1"),
        ("S1_clearspring", "Clearspring", "ЄС", PLUM, "Бюро Він", "органічний пауч 250 г, сегмент 1"),
        ("S2_other", "Рис швидкого приготування", "Таїланд", SLATE, "Азіяфудс", "пакет, сегмент 2"),
        ("S1_haldiram", "Haldiram RTE Biryani", "Індія", ROSE, "Тадж Махал", "пакет 200 г, сегмент 1")]


def _prod_data():
    D = {k: dict(kg=0, y={'2025': 0, '2026': 0}, decl=set(), months=[]) for k, *_ in PROD}
    for r in C_ROWS:
        s_ = cseg(r)
        if not (s_.startswith('S1') or s_ == 'S2_boilwater'): continue
        k = s_ if s_ != 'S2_boilwater' else ('S2_vend' if 'ВЕНДІНГ' in cg(r, 'importer_display_name') else 'S2_other')
        d = D[k]; kg = cg(r, 'net_weight_kg') or 0
        d['kg'] += kg; d['y'][cg(r, 'report_month')[:4]] += kg; d['decl'].add(cg(r, 'declaration_number')); d['months'].append(cg(r, 'report_month'))
    return D


PD = _prod_data()
assert abs(sum(d['kg'] for d in PD.values()) - (S1_KG + S2_KG)) < 0.5
YK = {'S1': {y: sum(PD[k]['y'][y] for k in PD if k.startswith('S1')) for y in ('2025', '2026')},
      'S2_boilwater': {y: PD['S2_vend']['y'][y] + PD['S2_other']['y'][y] for y in ('2025', '2026')}}
MON = {'2025': 12, '2026': 5}
MN = ['січ', 'лют', 'бер', 'кві', 'тра', 'чер', 'лип', 'сер', 'вер', 'жов', 'лис', 'гру']


def _mon(m): return f"{MN[int(m[5:]) - 1]}. {m[:4]}"


def s_composition():
    s = new_slide()
    header(s, "ОБСЯГ РИНКУ · МИТНА БАЗА", "Скільки готового рису ввозять в Україну",
           f"За {PERIOD}: {num(T(RICE_KG), 1)} т готового рису. Це {num(RICE_KG / CODE_KG * 100, 1)} % коду 1904901000.")
    px0, pw = M, 7.95
    rect(s, px0, 1.86, pw, 5.06, MIST, rounded=True, adj=0.03)
    text(s, px0 + 0.28, 2.0, 7.4, 0.26, "ЩО САМЕ ВВОЗЯТЬ, КГ ЗА ПЕРІОД", size=10.5, bold=True, color=GREY)
    yy = 2.34
    mx = max(PD[k]['kg'] for k, *_ in PROD)
    for k, nm, cty, col, imp, note in PROD:
        d = PD[k]
        text(s, px0 + 0.28, yy, 4.6, 0.28, nm, size=13, bold=True, color=INK)
        text(s, px0 + 4.2, yy, 3.5, 0.28, f"{num(d['kg'])} кг · {num(d['kg'] / RICE_KG * 100, 1)} %", size=13, bold=True, color=NAVY,
             align=PP_ALIGN.RIGHT)
        rect(s, px0 + 0.28, yy + 0.34, 7.4, 0.2, WHITE, rounded=True, adj=0.5)
        rect(s, px0 + 0.28, yy + 0.34, max(7.4 * d['kg'] / mx, 0.06), 0.2, col, rounded=True, adj=0.5)
        text(s, px0 + 0.28, yy + 0.56, 7.4, 0.2, f"{cty} · {note} · імпортер {imp}", size=9.5, color=GREY)
        yy += 0.755
    rx = px0 + pw + 0.22; rw = W - M - rx
    rect(s, rx, 1.86, rw, 3.30, NAVY, rounded=True, adj=0.05); rect(s, rx, 1.86, 0.09, 3.30, ORANGE, rounded=True, adj=0.5)
    text(s, rx + 0.3, 2.02, rw - 0.5, 0.3, "Готовий рис, сегменти 1–2", size=13, bold=True, color=AMBER)
    text(s, rx + 0.3, 2.38, rw - 0.5, 0.7, f"{num(T(RICE_KG), 1)} т", size=38, bold=True, color=WHITE)
    text(s, rx + 0.3, 3.12, rw - 0.5, 0.3, f"{num(sum(len(d['decl']) for d in PD.values()))} партій за {PERIOD_S}",
         size=11, color=RGBColor(0xB9, 0xC2, 0xDA))
    for i, (lb, v) in enumerate([("Сегмент 1 · розігрів", f"{num(T(S1_KG), 2)} т"), ("Сегмент 2 · під окріп", f"{num(T(S2_KG), 2)} т"),
                                 ("Ben's — одна партія", "676 кг")]):
        text(s, rx + 0.3, 3.62 + i * 0.46, rw - 1.9, 0.3, lb, size=11.5, color=RGBColor(0xD5, 0xDB, 0xEA))
        text(s, rx + rw - 1.6, 3.62 + i * 0.46, 1.4, 0.3, v, size=13, bold=True, color=WHITE, align=PP_ALIGN.RIGHT)
    rect(s, rx, 5.30, rw, 1.62, WHITE, line=MIST_D, lw=1.2, rounded=True, adj=0.06)
    text(s, rx + 0.24, 5.42, rw - 0.4, 0.26, "ПЕРІОД І МЕТОД", size=9, bold=True, color=ROSE)
    text(s, rx + 0.24, 5.70, rw - 0.4, 1.2,
         f"17 місяців: січень 2025 – травень 2026 (немає 12–31 березня 2025). У коді {num(T(CODE_KG), 1)} т; решта — "
         "заморожені суміші, чіпси, різото-набори, долма — це не готовий рис, їх не враховано.", size=10, color=INK, line=1.2)
    foot(s, CUST_FOOT)


def _pm(k, y): return PD[k]['y'][y] / MON[y]


def s_dynamics():
    s = new_slide()
    header(s, "МИТНА БАЗА · 2025 І 2026", "Скільки привозили у 2025 і 2026",
           "Середнє на місяць: 2025 — 12 місяців, 2026 — січень–травень. Так порівняння чесне.")
    lx, lw = M, 7.9
    rect(s, lx, 1.80, lw, 5.14, MIST, rounded=True, adj=0.03)
    text(s, lx + 0.28, 1.92, 4, 0.26, "КГ НА МІСЯЦЬ", size=10.5, bold=True, color=GREY)
    for i, (nm, col) in enumerate([("2025", SLATE), ("2026 · січ–трав", ORANGE)]):
        rect(s, lx + 4.4 + i * 1.5, 1.97, 0.16, 0.16, col, rounded=True, adj=0.3)
        text(s, lx + 4.65 + i * 1.5, 1.94, 1.3, 0.24, nm, size=10, color=INK)
    mx = max(_pm(k, y) for k, *_ in PROD for y in ('2025', '2026'))
    yy = 2.34; bx = lx + 0.28; bwmax = 4.3
    for k, nm, cty, col, imp, note in PROD:
        d = PD[k]
        text(s, bx, yy, 3.4, 0.26, nm, size=12, bold=True, color=NAVY)
        text(s, bx + 3.3, yy + 0.02, 4.2, 0.24, f"всього: 2025 — {num(d['y']['2025'])} кг · 2026 — {num(d['y']['2026'])} кг", size=9, color=GREY,
             align=PP_ALIGN.RIGHT)
        for j, (y, c) in enumerate([('2025', SLATE), ('2026', ORANGE)]):
            v = _pm(k, y)
            ww = max(bwmax * v / mx, 0.04)
            rect(s, bx, yy + 0.30 + j * 0.2, ww, 0.16, c, rounded=True, adj=0.3)
            text(s, bx + ww + 0.1, yy + 0.275 + j * 0.2, 1.4, 0.2, f"{num(v)} кг", size=9.5, bold=True, color=NAVY)
        yy += 0.755
    rx = lx + lw + 0.22; rw = W - M - rx
    s1 = YK['S1']; s2 = YK['S2_boilwater']
    cards = [("СЕГМЕНТ 1 · РОЗІГРІВ", ORANGE, f"{num(s1['2025'])} → {num(s1['2026'])} кг",
              f"У 2025 — {num(s1['2025'])} кг, з них 676 кг — одна партія Ben's. У 2026 за п'ять місяців — лише CJ Hetbahn, {num(s1['2026'])} кг."),
             ("СЕГМЕНТ 2 · ПІД ОКРІП", TEAL, f"×{num(_pm('S2_vend', '2026') / _pm('S2_vend', '2025'), 1)} на місяць",
              f"Китайський рис у чаші: {num(PD['S2_vend']['y']['2025'])} кг у 2025 і {num(PD['S2_vend']['y']['2026'])} кг за п'ять місяців 2026."),
             ("УСІ СЕГМЕНТИ 1–2 БЕЗ BEN'S", GREEN,
              f"×{num(((s1['2026'] + s2['2026']) / 5) / ((s1['2025'] + s2['2025'] - 676) / 12), 1)} на місяць",
              f"{num((s1['2025'] + s2['2025'] - 676) / 12)} → {num((s1['2026'] + s2['2026']) / 5)} кг на місяць.")]
    ch = 1.66
    for i, (lb, c, big, nt) in enumerate(cards):
        y = 1.80 + i * (ch + 0.08)
        rect(s, rx, y, rw, ch, WHITE, line=MIST_D, lw=1.2, rounded=True, adj=0.06)
        rect(s, rx, y, 0.09, ch, c, rounded=True, adj=0.5)
        text(s, rx + 0.26, y + 0.12, rw - 0.4, 0.22, lb, size=9, bold=True, color=GREY)
        text(s, rx + 0.26, y + 0.38, rw - 0.4, 0.45, big, size=19, bold=True, color=NAVY)
        text(s, rx + 0.26, y + 0.90, rw - 0.4, 0.72, nt, size=10.5, color=INK, line=1.15)
    foot(s, CUST_FOOT + " Ben's у 2025 — одна партія 676 кг (лютий).")


# ── імпортери готового рису ──
def s_importers():
    s = new_slide()
    header(s, "МИТНА БАЗА · ІМПОРТЕРИ", "Хто завозить готовий рис",
           f"Усього шість імпортерів за {PERIOD_S}, і жоден не возить регулярно: від однієї до чотирьох партій.")
    tw = 11.98
    cols = [("№", 0.45), ("ІМПОРТЕР", 2.0), ("ЩО ЗАВОЗИТЬ", 2.75), ("КРАЇНА", 1.35), ("ПАРТІЙ", 0.8), ("КОЛИ", 1.85), ("КГ", 0.85), ("ЧАСТКА", 1.93)]
    xx = M
    for lb, w_ in cols:
        text(s, xx + 0.1, 1.80, w_ - 0.12, 0.22, lb, size=9, bold=True, color=GREY, align=PP_ALIGN.RIGHT if lb == "КГ" else PP_ALIGN.LEFT)
        xx += w_
    y = 2.08; rh = 0.74
    mxs = max(d['kg'] for d in PD.values())
    rows_ = sorted(PROD, key=lambda p: -PD[p[0]]['kg'])
    for i, (k, nm, cty, col, imp, note) in enumerate(rows_):
        d = PD[k]
        rect(s, M, y, tw, rh - 0.06, MIST if i % 2 == 0 else WHITE, rounded=True, adj=0.15)
        mths = sorted(set(d['months']))
        when = _mon(mths[0]) if len(mths) == 1 else f"{_mon(mths[0])} – {_mon(mths[-1])}"
        vals = [str(i + 1), imp, nm, cty, str(len(d['decl'])), when]
        xx = M
        for j, ((lb, w_), v) in enumerate(zip(cols, vals)):
            if j == 2:
                text(s, xx + 0.1, y + 0.10, w_ - 0.15, 0.24, v, size=11.5, bold=True, color=INK)
                text(s, xx + 0.1, y + 0.36, w_ - 0.15, 0.22, note, size=9, color=GREY)
            else:
                text(s, xx + 0.1, y + 0.20, w_ - 0.12, 0.3, v, size=13 if j == 1 else 11.5, bold=(j in (1, 4)),
                     color=NAVY if j in (0, 1, 4) else INK)
            xx += w_
        text(s, xx + 0.0, y + 0.18, cols[6][1] - 0.05, 0.3, num(d['kg']), size=13, bold=True, color=NAVY, align=PP_ALIGN.RIGHT)
        xx += cols[6][1]
        bw2 = max(1.0 * d['kg'] / mxs, 0.05)
        rect(s, xx + 0.15, y + 0.22, bw2, 0.2, col, rounded=True, adj=0.4)
        text(s, xx + 0.15 + bw2 + 0.08, y + 0.2, 0.8, 0.24, f"{num(d['kg'] / RICE_KG * 100, 1)} %", size=10, bold=True, color=GREY)
        y += rh
    foot(s, CUST_FOOT + " Імпортери — за даними декларацій; Вендінг Кінгз і Азіяфудс не вказали виробника в описі.")
