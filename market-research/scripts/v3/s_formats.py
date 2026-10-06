# ══ СЕГМЕНТАЦІЯ ЗА УПАКОВКОЮ ═══════════════════════════════════════════════
HERO = {'pouch': RTE + 'md/bens_basmati250.jpg', 'cup': RTE + 'sku/ot_bibimbap.jpg',
        'doypack': RTE + 'sku/mh_curry.jpg', 'box': RTE + 'sku/hd_tomato_beef272.jpg'}
HOW = {'pouch': "НВЧ 60–90 с або гаряча вода", 'cup': "НВЧ 60–90 с · окріп 5–10 хв · саморозігрів 15 хв",
       'doypack': "Окріп 8–10 хв (сублімат і сухий рис)", 'box': "Хімічний нагрівач, 15 хв, без техніки"}
BANDS = {
    'pouch': [(0, 225, '220 г'), (225, 300, '240–250 г'), (300, 360, '350 г'), (360, 999, '400 г')],
    'cup': [(0, 100, '84 г'), (100, 200, '144–174 г'), (200, 260, '210–247 г'), (260, 300, '269–280 г'),
            (300, 999, '310–320 г')],
    'doypack': [(0, 95, '80–90 г'), (95, 126, '100–125 г'), (126, 160, '133–146 г'), (160, 999, '250 г')],
    'box': [(0, 200, '165–187 г'), (200, 300, '272 г'), (300, 999, '360–440 г')],
}
PAL = [ORANGE, TEAL, NAVY_L, ROSE, GREEN, PLUM, GOLD, SLATE]
ALLN = len(SKU)


def med_lo(L): return st.median([x['lo'] for x in L])


def fmt_stats(f):
    L = by_fmt(f)
    gs = [x['g'] for x in L if x['g']]
    brands = collections.Counter(x['brand'] for x in L)
    return L, gs, brands


def s_overview():
    s = new_slide()
    header(s, "СЕГМЕНТАЦІЯ", "Сегментація за упаковкою",
           f"Чотири формати, {ALLN} позицій: спершу упаковка, далі вага, ціна й бренди.")
    cw = (11.98 - 3 * 0.16) / 4
    for i, f in enumerate(FMT_ORDER):
        L, gs, brands = fmt_stats(f)
        x = M + i * (cw + 0.16); c = FMT_COL[f]
        rect(s, x, 1.80, cw, 3.36, WHITE, line=MIST_D, lw=1.2, rounded=True, adj=0.045)
        rect(s, x, 1.82, cw, 0.40, c, rounded=True, adj=0.3)
        text(s, x + 0.16, 1.82, cw - 0.3, 0.40, FMT_FULL[f], size=11, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
        rect(s, x + 0.12, 2.30, cw - 0.24, 1.40, MIST, rounded=True, adj=0.06)
        pic(s, HERO[f], x + 0.2, 2.34, cw - 0.4, 1.32)
        text(s, x + 0.16, 3.72, 0.9, 0.42, str(len(L)), size=26, bold=True, color=c)
        text(s, x + 0.16 + 0.62 + 0.14 * (len(str(len(L))) - 1), 3.85, cw - 1.1, 0.3, f"позицій · {num(len(L)/ALLN*100)} %",
             size=10.5, color=GREY)
        facts = [("ГРАМАЖ", f"{num(min(gs))}–{num(max(gs))} г"), ("МЕДІАНА ЦІНИ", f"{num(med_lo(L))} грн"),
                 ("БРЕНДІВ", f"{len(brands)}")]
        for j, (lb, v) in enumerate(facts):
            text(s, x + 0.16, 4.27 + j * 0.28, 1.3, 0.24, lb, size=8, bold=True, color=GREY)
            text(s, x + 1.16, 4.24 + j * 0.28, cw - 1.3, 0.26, v, size=12, bold=True, color=NAVY, align=PP_ALIGN.RIGHT)
    # ── нижня смуга: формат × спосіб приготування (кількість і медіана ціни)
    y0 = 5.30
    text(s, M, y0, 9, 0.26, "ФОРМАТ × СПОСІБ ПРИГОТУВАННЯ · ПОЗИЦІЙ І МЕДІАНА ЦІНИ", size=10, bold=True, color=GREY)
    colw = [1.9, 2.52, 2.52, 2.52, 2.52]
    heads = ["", "1 · Розігрів", "2 · Під окріп", "3 · Саморозігрів", "4 · Сублімат"]
    xx = M
    for h_, w_ in zip(heads, colw):
        text(s, xx, y0 + 0.30, w_, 0.22, h_, size=10, bold=True, color=NAVY, align=PP_ALIGN.CENTER if h_ else PP_ALIGN.LEFT)
        xx += w_
    mxv = max(sum(1 for x in SKU if x['fmt'] == f and x['seg'] == sg) for f in FMT_ORDER for sg in (1, 2, 3, 4))
    for r, f in enumerate(FMT_ORDER):
        yy = y0 + 0.56 + r * 0.275
        xx = M
        text(s, xx, yy + 0.01, colw[0], 0.24, FMT_NAME[f], size=11.5, bold=True, color=FMT_COL[f])
        xx += colw[0]
        for sg in (1, 2, 3, 4):
            L_ = [x for x in SKU if x['fmt'] == f and x['seg'] == sg]
            n = len(L_)
            if n:
                a = 0.25 + 0.75 * n / mxv
                col = RGBColor(*[int(255 - (255 - v) * a) for v in (NAVY[0], NAVY[1], NAVY[2])])
                rect(s, xx + 0.10, yy, colw[1] - 0.2, 0.245, col, rounded=True, adj=0.3)
                text(s, xx, yy + 0.01, colw[1], 0.23, f"{n} · {num(st.median([q['lo'] for q in L_]))} грн", size=10.5, bold=True,
                     color=WHITE if a > 0.5 else NAVY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            else:
                text(s, xx, yy + 0.01, colw[1], 0.23, "—", size=10.5, color=LGREY, align=PP_ALIGN.CENTER)
            xx += colw[1]
    foot(s, "Формат визначено за фактичною упаковкою на фото картки; способи приготування: 1 — розігрів у НВЧ/воді, 2 — сухий рис під окріп, "
            "3 — саморозігрів, 4 — сублімація. Медіана — за мінімальною ціною позиції. Qiaoshanmei — плоский пакет під окріп, віднесено до дойпаку.")


def band_rows(f):
    L = by_fmt(f)
    rows = []
    for lo, hi, label in BANDS[f]:
        part = [x for x in L if x['g'] and lo <= x['g'] < hi]
        if part: rows.append((label, part))
    nd = [x for x in L if not x['g']]
    if nd: rows.append(('маса н/д', nd))
    return rows


def s_format(f):
    L, gs, brands = fmt_stats(f)
    c = FMT_COL[f]
    s = new_slide()
    titles = {'pouch': ("Пауч", "Від порційних 220 г до реторту 400 г."),
              'cup': ("Чаша", "Найширший діапазон ваги: від 84 до 320 г."),
              'doypack': ("Дойпак", "Стільки ж позицій, скільки в пауча, — але це туризм."),
              'box': ("Коробка з нагрівачем", "Найдорожчий формат: без техніки й окропу.")}
    header(s, f"ФОРМАТ {FMT_ORDER.index(f) + 1} З 4 · {FMT_FULL[f]}", titles[f][0], titles[f][1])
    # ліва картка
    lx, lw = M, 3.85
    rect(s, lx, 1.80, lw, 5.14, WHITE, line=MIST_D, lw=1.2, rounded=True, adj=0.03)
    rect(s, lx, 1.80, lw, 0.10, c, rounded=True, adj=0.5)
    rect(s, lx + 0.15, 2.02, lw - 0.3, 1.62, MIST, rounded=True, adj=0.05)
    pic(s, HERO[f], lx + 0.25, 2.06, lw - 0.5, 1.54)
    sellers = collections.Counter()
    for x in L:
        for t in x['seller'].replace(' · ', '·').split('·'):
            t = t.strip()
            if t and not t.startswith('ще'): sellers[t] += 1
    top_sellers = ", ".join(k for k, _ in sellers.most_common(3))
    _cc = collections.Counter()
    for _b in brands: _cc[COUNTRY[_b]] += 1
    ctry_txt = " · ".join(f"{k} {v}" for k, v in _cc.most_common())
    facts = [("ПОЗИЦІЙ · БРЕНДІВ", f"{len(L)} · {len(brands)}  ({num(len(L)/ALLN*100)} % асортименту)"),
             ("ГРАМАЖ", f"{num(min(gs))}–{num(max(gs))} г"),
             ("МЕДІАНА · ДІАПАЗОН", f"{num(med_lo(L))} грн · {num(min(x['lo'] for x in L))}–{num(max(x['hi'] for x in L))}"),
             ("ПРИГОТУВАННЯ", HOW[f]), ("ПОХОДЖЕННЯ БРЕНДІВ", ctry_txt), ("ПРОДАВЦІ", top_sellers)]
    yy = 3.76
    for lb, v in facts:
        text(s, lx + 0.22, yy, lw - 0.4, 0.16, lb, size=8, bold=True, color=GREY)
        hh = 0.40 if len(v) > 36 else 0.24
        text(s, lx + 0.22, yy + 0.17, lw - 0.4, hh, v, size=11, bold=True, color=NAVY, line=1.03)
        yy += 0.17 + hh + 0.07
    # права частина: рядки по вазі
    rows = band_rows(f)
    bx0 = lx + lw + 0.25; bw = W - M - bx0
    cols = [("ФОТО", 1.78), ("ГРАМАЖ", 1.30), ("ПОЗИЦІЙ · КОЛІР = БРЕНД", 3.35), ("ЦІНА, ГРН", 1.55)]
    xx = bx0
    for lb, w_ in cols:
        text(s, xx + 0.08, 1.82, w_, 0.2, lb, size=8.5, bold=True, color=GREY); xx += w_
    order = [b for b, _ in brands.most_common()]
    colr = {b: PAL[i % len(PAL)] for i, b in enumerate(order)}
    avail = 6.94 - 2.08 - 1.05
    rh = min(avail / len(rows), 1.2)
    mxn = max(len(p) for _, p in rows)
    y = 2.08
    for lab, part in rows:
        rect(s, bx0, y, bw, rh - 0.08, MIST, rounded=True, adj=0.08)
        # мініфото: до трьох брендів
        seen = []
        for x in sorted(part, key=lambda q: -collections.Counter(p['brand'] for p in part)[q['brand']]):
            if x['brand'] not in [q['brand'] for q in seen]: seen.append(x)
        seen = seen[:3]
        th = rh - 0.24; tw = min(0.56, th * 0.8, (1.64 - (len(seen) - 1) * 0.06) / max(len(seen), 1))
        for k, x in enumerate(seen):
            rect(s, bx0 + 0.10 + k * (tw + 0.06), y + 0.08, tw, rh - 0.24, WHITE, rounded=True, adj=0.1)
            pic(s, x['img'], bx0 + 0.12 + k * (tw + 0.06), y + 0.10, tw - 0.04, rh - 0.28)
        text(s, bx0 + cols[0][1] + 0.08, y + (rh - 0.08) / 2 - 0.2, cols[1][1], 0.4, lab, size=15, bold=True, color=NAVY,
             anchor=MSO_ANCHOR.MIDDLE)
        # стовпчик по брендах + легенда «колір = бренд»
        sx = bx0 + cols[0][1] + cols[1][1] + 0.08
        bc = collections.Counter(x['brand'] for x in part)
        full = 1.45
        cx_ = sx
        for b, n in sorted(bc.items(), key=lambda kv: -kv[1]):
            ww = full * n / mxn
            rect(s, cx_, y + 0.08, ww, 0.20, colr[b]); cx_ += ww
        text(s, sx + full + 0.06, y + 0.04, 0.5, 0.3, str(len(part)), size=14, bold=True, color=NAVY)
        lx_, ly_ = sx, y + 0.34
        maxw = cols[2][1] - 0.12
        for b, n in sorted(bc.items(), key=lambda kv: -kv[1]):
            lab = f"{b} ×{n}" if n > 1 else b
            iw = 0.16 + 0.066 * len(lab) + 0.12
            if lx_ + iw > sx + maxw and lx_ > sx:
                lx_ = sx; ly_ += 0.17
            rect(s, lx_, ly_ + 0.035, 0.10, 0.10, colr[b], rounded=True, adj=0.3)
            text(s, lx_ + 0.14, ly_, iw, 0.17, lab, size=8.5, color=INK)
            lx_ += iw
        lo = min(x['lo'] for x in part); hi = max(x['hi'] for x in part)
        px = bx0 + cols[0][1] + cols[1][1] + cols[2][1] + 0.08
        text(s, px, y + (rh - 0.08) / 2 - 0.16, cols[3][1] - 0.1, 0.3, rngs(lo, hi), size=14, bold=True, color=c,
             anchor=MSO_ANCHOR.MIDDLE)
        y += rh
    _P, _C, _D, _B = (by_fmt(k) for k in FMT_ORDER)
    _ua = [x for x in _P if COUNTRY[x['brand']] == 'Україна']
    _soc = [x for x in _P if x['brand'] == 'Seeds of Change']
    _hd = sum(1 for x in _B if x['brand'] == 'Haidilao')
    INS = {
        'pouch': f"Пауч — найдешевший формат: від 45 грн. 350 г — український реторт: {len(_ua)} позицій, {len(set(x['brand'] for x in _ua))} брендів, "
                 f"медіана {num(st.median([x['lo'] for x in _ua]))} грн. 220–250 г — Ben's Original. 240 г — органіка Seeds of Change "
                 f"({num(min(x['lo'] for x in _soc))}–{num(max(x['hi'] for x in _soc))} грн, перепродаж iHerb). Вагових смуг 150–199 і 300 г немає.",
        'cup': "Чаша — єдиний формат у всіх вагових смугах від 84 до 320 г. Найдешевша — Henan 144 г (83 грн), "
               "найдорожча готова — Ottogi 320 г (356 грн). Наш стакан 150 г стає між Henan (144–174 г) і Bibigo (210 г).",
        'doypack': f"{len(_D)} позицій, але {sum(1 for x in _D if x['seg'] == 4)} із них — сублімат для походів: вага 80–250 г, ціни 55–889 грн. "
                   f"Це туристичний канал, а не мережева полиця: без туристичних магазинів частка дойпаку падає з "
                   f"{num(len(_D) / ALLN * 100)} % до {num(MASS_BY_FMT['doypack'] / N_MASS * 100)} %.",
        'box': f"{len(_B)} позицій і {len(set(x['brand'] for x in _B))} бренди: Haidilao дає {_hd} із {len(_B)}. Нагрівач у коробці замінює "
               f"мікрохвильовку, але медіана {num(med_lo(_B))} грн — у {num(med_lo(_B) / med_lo(_P), 1)} раза вища за пауч ({num(med_lo(_P))} грн).",
    }
    iy = y + 0.04
    insight(s, bx0, iy, bw, 6.94 - iy, None, INS[f], c, size=10.5 if 6.94 - iy < 1.3 else 11.5)
    foot(s, "Позиції — за картками продавців; грамаж — вага готової страви (для сухого рису й сублімату — вага продукту). "
            "Медіана — за мінімальною ціною позиції. Ціни за одну упаковку, без опту.")
