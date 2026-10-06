# ══ СХОДИ ЦІН (розширені), ЦІНОУТВОРЕННЯ, КАРТА «ХТО З КИМ СТОЇТЬ» ═══════════
import json as _json
_PR = {p['n']: p for p in _json.load(open(V3 + 'pricing_rows.json'))}
# наші товарні групи: (назва, формат, id рядків моделі)
OURS = [("BSCM пауч 150 г", 'pouch', [5, 6, 7, 8]), ("BSCM пауч 200 г", 'pouch', [9, 10, 11, 12]),
        ("BSCM пауч 240 г", 'pouch', [1, 2, 3, 4]), ("CM Premium пауч 250 г", 'pouch', [23, 24, 25, 26]),
        ("BSCM стакан 150 г", 'cup', [13, 14, 15, 16, 31]), ("BSCM подвійний 2×125 г", 'cup', [17, 18, 19, 20, 21, 22, 32]),
        ("ONE'S лоток 210 г", 'cup', [30]), ("ONE'S fried rice 200 г", 'cup', [27, 28, 29])]
OURS_G = {"BSCM пауч 150 г": 150, "BSCM пауч 200 г": 200, "BSCM пауч 240 г": 240, "CM Premium пауч 250 г": 250,
          "BSCM стакан 150 г": 150, "BSCM подвійний 2×125 г": 250, "ONE'S лоток 210 г": 210, "ONE'S fried rice 200 г": 200}


def _wtxt(gs):
    gs = sorted({g_ for g_ in gs if g_})
    if not gs: return "маса н/д"
    f_ = lambda v: num(v, 0 if float(v).is_integer() else 1)
    return f"{f_(gs[0])} г" if len(gs) == 1 else f"{f_(gs[0])}–{f_(gs[-1])} г"


def r_sku(label, L, brand=None):
    brand = brand or L[0]['brand']
    return dict(nm=label, cty=COUNTRY[brand], wt=_wtxt([x['g'] for x in L]), pts=[x['lo'] for x in L],
                lo=min(x['lo'] for x in L), hi=max(x['hi'] for x in L), n=len(L), ours=False)


def r_ours(label, ids):
    v = [_PR[i]['shelf'] for i in ids]
    g_ = OURS_G[label]
    return dict(nm="НАШ · " + label, cty='—', wt=_wtxt([_PR[i]['g'] for i in ids]) if label.startswith(('BSCM пауч', 'CM', 'BSCM стакан', "ONE'S")) else "2×125 г",
                pts=v, lo=min(v), hi=max(v), n=len(v), ours=True)


def sel(brand, fmt=None, seg=None, g_=None):
    return [x for x in SKU if x['brand'] == brand and (fmt is None or x['fmt'] == fmt) and (seg is None or x['seg'] == seg)
            and (g_ is None or x['g'] in (g_ if isinstance(g_, (list, tuple)) else [g_]))]


def lad(title, eyebrow, dek, groups, axmax, ticks, note, window=True, y1=6.95, insight_txt=None):
    s = new_slide()
    header(s, eyebrow, title, dek)
    nrows = sum(len(g_[2]) for g_ in groups); nh = len(groups)
    y0 = 1.78
    ins_h = 0.78 if insight_txt else 0
    rh = min((y1 - y0 - nh * 0.34 - ins_h - 0.55) / nrows, 0.52)
    ybot = y0 + nh * 0.34 + nrows * rh
    AX, AW = 5.55, 6.55
    ax = lambda v: AX + AW * v / axmax
    for t in ticks:
        rect(s, ax(t), y0 + 0.02, 0.012, ybot - y0 - 0.02, MIST_D)
        text(s, ax(t) - 0.4, ybot + 0.03, 0.8, 0.2, num(t), size=10, color=GREY, align=PP_ALIGN.CENTER)
    if window:
        rect(s, ax(45), y0 + 0.02, ax(180) - ax(45), ybot - y0 - 0.02, RGBColor(0xFF, 0xF6, 0xE3))
        text(s, ax(45), ybot + 0.25, 4.2, 0.22, "вікно мережевої полиці 45–180 грн", size=10, bold=True, color=ORANGE)
    text(s, M, ybot + 0.03, 4.4, 0.2, "ГРН/УПАКОВКА · ВЕЛИКА ТОЧКА — МЕДІАНА, МАЛІ — ПОЗИЦІЇ", size=8, bold=True, color=GREY)
    y = y0
    for gname, gcol, rows in groups:
        rect(s, M, y + 0.03, W - 2 * M, 0.28, gcol, rounded=True, adj=0.3)
        text(s, M + 0.14, y + 0.03, 6.0, 0.28, gname, size=10.5, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
        for lbl, dx, al in (("КРАЇНА", 3.0, PP_ALIGN.LEFT), ("ВАГА", 3.78, PP_ALIGN.LEFT), ("SKU", 4.62, PP_ALIGN.RIGHT)):
            text(s, M + dx, y + 0.03, 0.8, 0.28, lbl, size=8, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE, align=al)
        y += 0.34
        for r in rows:
            cy = y + rh / 2
            if r['ours']:
                rect(s, M, y + 0.02, W - 2 * M, rh - 0.04, RGBColor(0xFF, 0xF1, 0xD2), rounded=True, adj=0.2)
            text(s, M + 0.14, cy - 0.13, 2.85, 0.28, r['nm'], size=11.5 if not r['ours'] else 11, bold=True, color=NAVY)
            text(s, M + 3.0, cy - 0.10, 0.8, 0.24, r['cty'], size=9.5, color=GREY)
            text(s, M + 3.78, cy - 0.10, 0.95, 0.24, r['wt'], size=9.5, color=INK)
            text(s, M + 4.5, cy - 0.10, 0.4, 0.24, str(r['n']), size=10, bold=True, color=GREY, align=PP_ALIGN.RIGHT)
            col = ORANGE if r['ours'] else gcol
            lo, hi = r['lo'], r['hi']
            x0, x1 = ax(lo), ax(min(hi, axmax))
            rect(s, x0, cy - 0.045, max(x1 - x0, 0.04), 0.09, col, rounded=True, adj=0.5)
            for p_ in r['pts']:
                dot(s, ax(p_), cy, 0.11, WHITE, line=col, lw=1.4)
            med = st.median(r['pts'])
            dot(s, ax(med), cy, 0.24, col, line=NAVY if r['ours'] else WHITE, lw=2.0 if r['ours'] else 1.5)
            lbl = num(lo) if round(lo) == round(hi) else f"{num(lo)}–{num(hi)}"
            text(s, x1 + 0.2, cy - 0.12, 1.5, 0.26, lbl, size=10.5, bold=True, color=NAVY)
            y += rh
    if insight_txt:
        insight(s, M, y1 - ins_h + 0.1, 11.98, ins_h - 0.1, None, insight_txt, ORANGE, size=11)
    foot(s, note)
    return s


def _ours_agg():
    bscm = [i for n_, f_, ids in OURS if n_.startswith('BSCM') for i in ids]
    cm = OURS[3][2]; on = OURS[6][2] + OURS[7][2]

    def agg(label, ids, wt):
        v = [_PR[i]['shelf'] for i in ids]
        return dict(nm="НАШ · " + label, cty='Таїланд' if 'ONE' not in label else 'Корея', wt=wt, pts=v, lo=min(v), hi=max(v), n=len(v), ours=True)
    return [agg('BSCM Foods', bscm, '150–250 г'), agg('CM Premium · Chefrey', cm, '250 г'), agg("ONE'S International", on, '200–210 г')]


def s_ladder1():
    sp = ['Seeds of Change', 'Маркел', 'Portion', "М'ясниця", 'МАКРО', 'Верес', 'Ходорівський', 'Clearspring']
    R1 = [r_sku("Ben's Original", sel("Ben's Original"))] + [r_sku(b, sel(b, 'pouch', 1)) for b in sp] + \
         [r_sku('Bibigo', sel('Bibigo')), r_sku('Ottogi', sel('Ottogi')), r_sku('Adventure Menu · пауч 400 г', sel('Adventure Menu', 'pouch', 1))]
    R1.sort(key=lambda r: st.median(r['pts']))
    groups = [("СЕГМЕНТ 1 · РОЗІГРІВ", GOLD, R1), ("НАШІ БРЕНДИ · ПОЛИЦЯ", NAVY, _ours_agg())]
    return lad("Ціна: готовий рис і наші бренди", "ЦІНА ЗА ОДНУ УПАКОВКУ · 1 З 3",
               "Сегмент 1 — розігрів у мікрохвильовці: усі бренди й наші позиції на одній шкалі.", groups, 520, [0, 100, 200, 300, 400, 500],
               "Роздрібна ціна за одну упаковку на сайтах продавців; кожна позиція — її мінімальна ціна. Наші бренди — полиця за фінмоделлю: "
               "ціна партнеру з ПДВ × 1,4 (20' навалом, курс 45). Смуга — від найдешевшої до найдорожчої позиції.")


def s_ladder2():
    R2 = [r_sku('Henan', sel('Henan')), r_sku('Gallina Blanca', sel('Gallina Blanca')), r_sku('Qiaoshanmei', sel('Qiaoshanmei'))]
    R3 = [r_sku(b, sel(b, None, 3)) for b in ('Haidilao', 'Zihaiguo', 'Rongcheng Haoji', 'Mo Xiao Xian', 'Forestia')]
    R2.sort(key=lambda r: st.median(r['pts'])); R3.sort(key=lambda r: st.median(r['pts']))
    groups = [("СЕГМЕНТ 2 · ПІД ОКРІП", TEAL, R2), ("СЕГМЕНТ 3 · САМОРОЗІГРІВ", ROSE, R3)]
    hm = st.median([x['lo'] for x in sel('Henan')])
    return lad("Ціна: під окріп і саморозігрів", "ЦІНА ЗА ОДНУ УПАКОВКУ · 2 З 3",
               f"Henan стоїть на рівні пауча, саморозігрів — у п'ять–вісім разів дорожче.", groups, 1250, [0, 250, 500, 750, 1000, 1250],
               "Роздрібна ціна за одну упаковку на сайтах продавців; кожна позиція — її мінімальна ціна.", window=False,
               insight_txt=f"Henan (чаша 144–174 г): медіана {num(hm)} грн проти {num(st.median([x['lo'] for x in by_fmt('pouch')]))} грн у пауча. "
                           f"Саморозігрів: Haidilao 380–1 190 грн, решта бренди — 735–935 грн; це інший привід (похід, дача), а не мережева полиця.")


def s_ladder3():
    ua = ['James Cook', '!FEST', 'Їжа в Похід', 'Харчі', 'SubliMate']
    im = ["Trek'n Eat", 'Adventure Food', 'Travellunch', 'Mountain House']
    R4 = [r_sku(b, sel(b, 'doypack', 4)) for b in ua + im] + [r_sku('Adventure Menu · дойпак 110 г', sel('Adventure Menu', 'doypack', 4))]
    R4.sort(key=lambda r: st.median(r['pts']))
    groups = [("СЕГМЕНТ 4 · СУБЛІМАЦІЯ", PLUM, R4)]
    return lad("Ціна: сублімація", "ЦІНА ЗА ОДНУ УПАКОВКУ · 3 З 3",
               "Український сублімат — від 55 до 370 грн, імпортний — від 270 до 890.", groups, 950, [0, 200, 400, 600, 800],
               "Роздрібна ціна за одну упаковку на сайтах продавців; кожна позиція — її мінімальна ціна. Вага — продукту в упаковці (сухого).",
               window=False,
               insight_txt="Український сублімат (James Cook, Харчі, Їжа в Похід, !FEST, SubliMate) — 80–140 г і 55–370 грн; імпортний "
                           "(Travellunch, Trek'n Eat, Mountain House, Adventure) — 110–250 г і 269–889 грн. Продається в туристичних магазинах.")


# ── ціноутворення ───────────────────────────────────────────────────────────
FIN = [("BSCM стакан 150 г", 13, 0.500), ("BSCM пауч 150 г", 5, 0.472), ("BSCM пауч 200 г", 9, 0.550),
       ("BSCM пауч 240 г", 1, 0.583), ("CM Premium 250 г", 26, 0.667), ("ONE'S білий 210 г", 30, None), ("ONE'S fried 200 г", 27, None)]


def s_fin():
    s = new_slide()
    header(s, "ЦІНОУТВОРЕННЯ", "З чого складається ціна на полиці",
           "Полиця — це собівартість плюс бонус мережі, наша маржа й націнка магазину. Закупівля (FOB) — окремо, справа.")
    parts = [("Собівартість (СС)", NAVY_L), ("Бонус мережі", GOLD), ("Наша маржа", ORANGE), ("Націнка магазину", LGREY)]
    sh = [0.4 / 1.4 * 100, 0.25 / 1.4 * 100, 0.35 / 1.4 * 100, 0.4 / 1.4 * 100]
    rect(s, M, 1.80, 11.98, 1.12, MIST, rounded=True, adj=0.1)
    text(s, M + 0.25, 1.88, 8, 0.24, "З КОЖНИХ 100 ГРН НА ПОЛИЦІ", size=10, bold=True, color=GREY)
    xx = M + 0.25; full = 11.48
    for (lb, col), v in zip(parts, sh):
        ww = full * v / 100
        rect(s, xx, 2.18, ww, 0.40, col)
        text(s, xx, 2.18, ww, 0.40, f"{num(v, 1)} %", size=12.5, bold=True, color=NAVY if col in (LGREY, ORANGE) else WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(s, xx + 0.04, 2.62, ww - 0.06, 0.24, lb, size=9.5, bold=True, color=INK, align=PP_ALIGN.CENTER)
        xx += ww
    y0 = 3.12; rh = 0.46
    bx = M + 2.45; scale = 4.9 / 340.0
    text(s, M, y0 - 0.02, 2.3, 0.22, "ТОВАР", size=8.5, bold=True, color=GREY)
    text(s, bx, y0 - 0.02, 5.4, 0.22, "ПОЛИЦЯ, ГРН ЗА ОДНУ УПАКОВКУ", size=8.5, bold=True, color=GREY)
    zx = M + 8.55
    rect(s, zx - 0.12, y0 - 0.04, 0.012, 3.6, MIST_D)
    text(s, zx, y0 - 0.02, 3.4, 0.22, "ЗАКУПІВЛЯ · ОКРЕМО ВІД ПОЛИЦІ", size=8.5, bold=True, color=NAVY)
    for lb, dx, w_, al in (("FOB, $", 0.0, 0.9, PP_ALIGN.LEFT), ("СС, $", 1.0, 0.9, PP_ALIGN.LEFT), ("СС, грн", 2.0, 1.0, PP_ALIGN.LEFT)):
        text(s, zx + dx, y0 + 0.2, w_, 0.2, lb, size=8.5, bold=True, color=GREY, align=al)
    y = y0 + 0.42
    for i, (nm, n_, fob) in enumerate(FIN):
        p = _PR[n_]; cc = p['cc_uah']; partner = p['partner']; shelf = p['shelf']
        segs = [cc, 0.25 * partner, 0.35 * partner, shelf - partner]
        if i % 2 == 0: rect(s, M, y - 0.03, 11.98, rh - 0.02, MIST, rounded=True, adj=0.2)
        text(s, M + 0.12, y + 0.07, 2.3, 0.3, nm, size=12, bold=True, color=NAVY)
        xx = bx
        for (lb, col), v in zip(parts, segs):
            ww = v * scale
            rect(s, xx, y + 0.05, ww, 0.32, col)
            if ww > 0.34:
                text(s, xx, y + 0.05, ww, 0.32, num(v), size=10, bold=True, color=NAVY if col in (LGREY, ORANGE) else WHITE,
                     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            xx += ww
        text(s, xx + 0.1, y + 0.05, 0.7, 0.32, f"{num(shelf)}", size=13, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
        text(s, zx, y + 0.07, 0.9, 0.3, f"${fob:.3f}" if fob else "н/д*", size=11.5, bold=True, color=NAVY if fob else GREY)
        text(s, zx + 1.0, y + 0.07, 0.9, 0.3, f"${p['cc']:.3f}", size=11.5, color=INK)
        text(s, zx + 2.0, y + 0.07, 1.0, 0.3, num(cc, 1), size=11.5, color=INK)
        y += rh
    lx = M
    for lb, col in parts:
        rect(s, lx, y + 0.14, 0.15, 0.15, col, rounded=True, adj=0.3)
        text(s, lx + 0.22, y + 0.10, 2.4, 0.24, lb, size=9.5, color=INK); lx += 0.55 + 0.085 * len(lb)
    foot(s, "Фінмодель покупця (20' FCL навалом, курс 45 грн/$, ціна вже з ПДВ): ціна партнеру = СС ÷ 0,4; бонус мережі = 25 % і наша маржа = 35 % "
            "від ціни партнеру; полиця = ціна партнеру × 1,4. СС — Self_Cost. FOB — комерційна пропозиція постачальника, до полиці не входить. "
            "*ONE'S не розкрив FOB.")


# ── карта «хто з ким стоїть» ────────────────────────────────────────────────
def _ours_rows_for(fmt):
    return [(n_, ids) for n_, f_, ids in OURS if f_ == fmt]


def stand_fmt(fmt, title, eyebrow, dek, spec_rows, axmax, ticks, note, insight_txt):
    rows = spec_rows
    rows.sort(key=lambda r: st.median(r['pts']))
    return lad(title, eyebrow, dek, [(f"{FMT_NAME[fmt].upper()} · ВІД ДЕШЕВИХ ДО ДОРОГИХ", FMT_COL[fmt], rows)],
               axmax, ticks, note, window=True, insight_txt=insight_txt)


def s_stand_pouch():
    P = by_fmt('pouch')
    rows = []
    for nm, gs in (("Ben's Original · 220 г", [220]), ("Ben's Original · 240–250 г", [240, 250])):
        rows.append(r_sku(nm, [x for x in P if x['brand'] == "Ben's Original" and x['g'] in gs]))
    rows.append(r_sku('Seeds of Change · 240 г', sel('Seeds of Change')))
    rows.append(r_sku('Clearspring · 250 г', sel('Clearspring')))
    for b in ('Portion', 'Маркел', "М'ясниця", 'МАКРО', 'Верес', 'Ходорівський'):
        rows.append(r_sku(b + ' · 350 г', sel(b, 'pouch')))
    rows.append(r_sku('Adventure Menu · 400 г', sel('Adventure Menu', 'pouch')))
    rows += [r_ours(n_, ids) for n_, ids in _ours_rows_for('pouch')]
    near = [r for r in rows if not r['ours'] and 100 <= st.median(r['pts']) <= 200]
    return stand_fmt('pouch', "Де ми стоїмо: пауч", "ХТО З КИМ СТОЇТЬ · ПАУЧ 1 З 2",
                     "Наш пауч 133–202 грн стоїть між українським реторт-паучем (медіана 101) і Seeds of Change (242). Помаранчеве — наші.", rows, 480,
                     [0, 100, 200, 300, 400],
                     "Роздрібна ціна за одну упаковку: конкуренти — мінімальна ціна позиції на сайтах продавців; наші — полиця за фінмоделлю з ПДВ.",
                     None)


def s_stand_cup():
    C = by_fmt('cup')
    rows = [r_sku('Henan · 144 г', [x for x in C if x['brand'] == 'Henan' and x['g'] == 144]),
            r_sku('Henan · 174 г', [x for x in C if x['brand'] == 'Henan' and x['g'] == 174]),
            r_sku('Gallina Blanca Yatekomo · 84 г', sel('Gallina Blanca')),
            r_sku('Bibigo · 210 г', sel('Bibigo')),
            r_sku('Ottogi · 217–247 г', [x for x in C if x['brand'] == 'Ottogi' and x['g'] < 250]),
            r_sku('Ottogi · 269–280 г', [x for x in C if x['brand'] == 'Ottogi' and 250 <= x['g'] < 300]),
            r_sku('Ottogi · 310–320 г', [x for x in C if x['brand'] == 'Ottogi' and x['g'] >= 300]),
            r_sku('Mo Xiao Xian · 275 г (саморозігрів)', sel('Mo Xiao Xian'))]
    rows += [r_ours(n_, ids) for n_, ids in _ours_rows_for('cup')]
    return stand_fmt('cup', "Де ми стоїмо: чаша", "ХТО З КИМ СТОЇТЬ · ЧАША 2 З 2",
                     "Наш стакан 150 г стоїть між Henan (медіана 124) і Bibigo (135–189); ONE'S fried rice — на рівні Ottogi. Помаранчеве — наші.", rows, 1000, [0, 200, 400, 600, 800, 1000],
                     "Роздрібна ціна за одну упаковку: конкуренти — мінімальна ціна позиції на сайтах продавців; наші — полиця за фінмоделлю з ПДВ.",
                     None)


def s_stand_matrix():
    s = new_slide()
    header(s, "ХТО З КИМ СТОЇТЬ · КАРТА", "Карта цін: формат × цінова смуга",
           "Усі бренди в клітинках за мінімальною ціною позиції; помаранчеве — наші товари за ціною полиці.")
    bands = [(0, 100, "до 100 грн"), (100, 150, "100–149"), (150, 200, "150–199"), (200, 300, "200–299"), (300, 500, "300–499"), (500, 5000, "500+")]
    cw0 = 1.15; cw = (11.98 - cw0) / 4
    for j, f in enumerate(FMT_ORDER):
        x = M + cw0 + j * cw
        rect(s, x + 0.04, 1.78, cw - 0.08, 0.32, FMT_COL[f], rounded=True, adj=0.25)
        text(s, x + 0.04, 1.78, cw - 0.08, 0.32, FMT_NAME[f] + f" · {len(by_fmt(f))}", size=11.5, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    rh = (6.98 - 2.14) / len(bands)
    for i, (lo, hi, lab) in enumerate(bands):
        y = 2.16 + i * rh
        text(s, M, y + rh / 2 - 0.13, cw0 - 0.1, 0.3, lab, size=11.5, bold=True, color=NAVY)
        for j, f in enumerate(FMT_ORDER):
            x = M + cw0 + j * cw
            rect(s, x + 0.04, y + 0.02, cw - 0.08, rh - 0.05, MIST, rounded=True, adj=0.05)
            L = [q for q in by_fmt(f) if lo <= q['lo'] < hi]
            bc = collections.Counter(q['brand'] for q in L)
            items = [f"{b} ×{n}" if n > 1 else b for b, n in bc.most_common()]
            ours = []
            for n_, f_, ids in OURS:
                if f_ != f: continue
                k = sum(1 for i_ in ids if lo <= _PR[i_]['shelf'] < hi)
                if k: ours.append(f"{n_}" + (f" ×{k}" if len(ids) > 1 else ""))
            yy = y + 0.06
            if items:
                txt = " · ".join(items)
                text(s, x + 0.14, yy, cw - 0.28, rh - 0.12 - 0.18 * min(len(ours), 3), txt, size=9.3, color=INK, line=1.05)
            elif not ours:
                text(s, x + 0.14, yy, cw - 0.28, 0.2, "—", size=10, color=LGREY)
            for k, o in enumerate(ours[:3]):
                oy = y + rh - 0.06 - 0.18 * (len(ours[:3]) - k)
                rect(s, x + 0.12, oy, cw - 0.24, 0.165, ORANGE, rounded=True, adj=0.4)
                text(s, x + 0.12, oy, cw - 0.24, 0.165, o, size=8, bold=True, color=NAVY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    foot(s, "Чорний текст — бренди конкурентів і кількість їхніх позицій у смузі (за мінімальною ціною на сайтах продавців); "
            "помаранчеві плашки — наші товарні групи (полиця за фінмоделлю з ПДВ), ×n — число смаків у смузі.")


def s_stand_neighbors():
    s = new_slide()
    header(s, "ХТО З КИМ СТОЇТЬ · СУСІДИ", "Найближчі за ціною до кожного нашого товару",
           "Конкуренти сегментів 1–2 (готовий рис і рис під окріп): три найближчі нижче й вище, плюс коридор ±20 %.")
    comp = [x for x in SKU if x['seg'] in (1, 2)]
    cols = [("НАШ ТОВАР", 2.55), ("ДЕШЕВШЕ ЗА НАС", 3.55), ("ДОРОЖЧЕ ЗА НАС", 3.55), ("КОРИДОР ±20 %", 2.33)]
    xx = M
    for lb, w_ in cols:
        text(s, xx + 0.12, 1.78, w_, 0.2, lb, size=8.5, bold=True, color=GREY); xx += w_
    y = 2.02; rh = 0.60

    def short(x):
        nm = x['name']
        nm = nm if len(nm) <= 22 else nm[:21] + '…'
        return f"{x['brand']} {nm} · {num(x['g']) + ' г' if x['g'] else ''} · {num(x['lo'])}"
    for i, (nm, f_, ids) in enumerate(OURS):
        v = [_PR[k]['shelf'] for k in ids]; c = st.median(v)
        lo_s, hi_s = min(v), max(v)
        below = sorted([x for x in comp if x['lo'] < lo_s], key=lambda q: -q['lo'])[:3]
        above = sorted([x for x in comp if x['lo'] > hi_s], key=lambda q: q['lo'])[:3]
        corr = [x for x in comp if 0.8 * c <= x['lo'] <= 1.2 * c]
        if i % 2 == 0: rect(s, M, y, 11.98, rh - 0.04, MIST, rounded=True, adj=0.1)
        rect(s, M, y, 0.07, rh - 0.04, ORANGE, rounded=True, adj=0.5)
        text(s, M + 0.18, y + 0.05, 2.35, 0.24, nm, size=11, bold=True, color=NAVY)
        text(s, M + 0.18, y + 0.29, 2.35, 0.24, f"{rngs(lo_s, hi_s)} грн" if round(lo_s) != round(hi_s) else f"{num(lo_s)} грн", size=12, bold=True, color=ORANGE)
        x_ = M + cols[0][1]
        for L, w_ in ((below, cols[1][1]), (above, cols[2][1])):
            paras = [(short(q), {}) for q in L] or [("— немає", {"color": LGREY})]
            text(s, x_ + 0.12, y + 0.03, w_ - 0.2, rh - 0.06, paras, size=8.6, color=INK, line=1.0)
            x_ += w_
        brs = list(dict.fromkeys(q['brand'] for q in sorted(corr, key=lambda q: abs(q['lo'] - c))))
        text(s, x_ + 0.12, y + 0.04, cols[3][1] - 0.2, 0.22, f"{len(corr)} позицій", size=11, bold=True, color=NAVY)
        text(s, x_ + 0.12, y + 0.27, cols[3][1] - 0.2, 0.3, ", ".join(brs[:4]) + ("…" if len(brs) > 4 else ""), size=8.2, color=GREY, line=1.0)
        y += rh
    foot(s, "Формат запису: бренд, назва, вага, мінімальна ціна на сайтах продавців, грн. Коридор — позиції сегментів 1–2 із ціною ±20 % від медіани нашого товару. "
            "Наші ціни — полиця за фінмоделлю з ПДВ. Саморозігрів і сублімат — інший привід, до сусідів не входять.")
