# ══ ГОЛОВНЕ, КАТАЛОГ, «ДЕ МИ СТОЇМО», ФІНМОДЕЛЬ, ТАБЛИЦЯ ПОЛИЦІ ══════════════
def s_kpi():
    s = new_slide()
    header(s, "ГОЛОВНЕ", "Шість цифр, які описують ринок", "Усі ціни в колоді — за одну упаковку. Опт і ящики до розрахунку не входять.")
    _s12 = YK['S1']['2025'] + YK['S2_boilwater']['2025'] - 676
    _s12n = YK['S1']['2026'] + YK['S2_boilwater']['2026']
    _x = (_s12n / 5) / (_s12 / 12)
    _med = {f: st.median([x['lo'] for x in by_fmt(f)]) for f in FMT_ORDER}
    K = [("ПОЗИЦІЙ У ПРОДАЖУ", str(N), PLUM, f"{N_BR} брендів · 4 формати · {N_SELL} продавців. Кожна позиція має фото й ціну в каталозі."),
         ("НА ПОЛИЦІ МЕРЕЖ", "12", TEAL, "SKU на всю країну, 11 з них — Ben's Original у «Сільпо». Ритейл-аудит 9 мереж."),
         ("ІМПОРТ ЗА МИТНИЦЕЮ", f"{num(T(S1_KG + S2_KG), 1)} т", ORANGE,
          f"Сегменти 1–2 за {PERIOD_S} — лише {num((S1_KG + S2_KG) / CODE_KG * 100, 1)} % коду 1904901000. Решта — заморожені суміші, чіпси, різото й долма: не готовий рис."),
         ("2026 ПРОТИ 2025", f"×{num(_x, 1)}", GREEN, f"Наші сегменти без разової партії Ben's: {num(_s12 / 12)} → {num(_s12n / 5)} кг на місяць."),
         ("РОЗРИВ ЦІН ФОРМАТІВ", f"×{num(_med['box'] / _med['pouch'], 1)}", ROSE,
          f"Медіана коробки з нагрівачем {num(_med['box'])} грн проти {num(_med['pouch'])} грн у пауча."),
         ("ВХІД У КАТЕГОРІЮ", f"{num(_med['pouch'])} грн", GOLD,
          f"Медіана пауча — найдешевшого формату: {len(by_fmt('pouch'))} позицій, вага 220–400 г.")]
    cw = (11.98 - 2 * 0.18) / 3; ch = 2.46
    for i, (lb, v, c, note) in enumerate(K):
        x = M + (i % 3) * (cw + 0.18); y = 1.86 + (i // 3) * (ch + 0.18)
        rect(s, x, y, cw, ch, MIST, rounded=True, adj=0.06); rect(s, x, y, cw, 0.10, c, rounded=True, adj=0.5)
        text(s, x + 0.3, y + 0.34, cw - 0.6, 0.24, lb, size=10.5, bold=True, color=GREY)
        text(s, x + 0.3, y + 0.70, cw - 0.6, 0.8, v, size=44, bold=True, color=c)
        text(s, x + 0.3, y + 1.62, cw - 0.6, 0.8, note, size=11, color=INK, line=1.2)
    foot(s, f"Джерела: картки продавців 23–28.09.2026 і 06.10.2026 ({N} позицій); митна база codexmb, вивантаження 28.09.2026, {PERIOD}; ритейл-аудит 9 мереж, 28.09.2026.")


def _catalog(items, eyebrow, title, dek, note=None, cols=None, note_c=ORANGE):
    s = new_slide()
    header(s, eyebrow, title, dek)
    n = len(items)
    two = n > 7
    cols = cols or (math.ceil(n / 2) if two else n)
    gap = 0.12
    w = (11.98 - gap * (cols - 1)) / cols
    top = 1.80
    avail = 7.0 - top - (1.12 if note else 0)
    h = (avail - (gap if two else 0)) / (2 if two else 1)
    h = min(h, 3.9)
    for i, x in enumerate(items):
        r, ccol = divmod(i, cols)
        sku_cell(s, M + ccol * (w + gap), top + r * (h + gap), w, h, x['img'], x['name'], x['g'], x['lo'], x['hi'],
                 x['seller'], FMT_COL[x['fmt']] if False else SEG_COL[x['seg']], brand=x['brand'])
    if note:
        yb = top + (2 if two else 1) * h + (gap if two else 0) + 0.14
        insight(s, M, yb, 11.98, min(6.94 - yb, 0.98), None, note, note_c, size=11)
    foot(s, "Фото, вага й ціна — з карток продавців; ціна за одну упаковку (від–до серед продавців). Колір назви бренду: золотий — розігрів, бірюзовий — під окріп, рожевий — саморозігрів, фіолетовий — сублімат.")
    return s


def cat_slides(f):
    """Каталог позицій формату — слайди, що йдуть одразу за слайдом формату."""
    L = by_fmt(f)
    tag = f"{FMT_NAME[f].upper()} · КАТАЛОГ"
    if f == 'pouch':
        ben = [x for x in L if x['brand'] == "Ben's Original"]
        ua = [x for x in L if COUNTRY[x['brand']] == 'Україна']
        imp = [x for x in L if x not in ben and x not in ua]
        assert len(ben) + len(ua) + len(imp) == len(L)
        G = [(ben, "Пауч: Ben's Original", f"{len(ben)} позицій, 220–250 г — єдиний бренд готового гарніру на мережевій полиці.", None),
             (ua, "Пауч: український реторт", f"{len(ua)} позицій, 350 г: каша чи плов із м'ясом, не гарнір. 71–199 грн, готовий після розігріву.", None),
             (imp, "Пауч: органіка й імпорт", "Seeds of Change — органіка США, лише через перепродавців iHerb на Prom; Clearspring — ЄС; Adventure Menu — 400 г для походів.", None)]
    elif f == 'cup':
        a_ = [x for x in L if x['seg'] == 1]; b_ = [x for x in L if x['seg'] != 1]
        G = [(a_, "Чаша: готовий рис", "Ottogi і Bibigo — розігрів у мікрохвильовці 60–90 с.", None),
             (b_, "Чаша: під окріп і саморозігрів", "Henan, Gallina Blanca — окріп 5–10 хв; Mo Xiao Xian — нагрівач.", None)]
    elif f == 'doypack':
        d_imp = [x for x in L if x['brand'] in ('Travellunch', "Trek'n Eat", 'Mountain House', 'Adventure Food')]
        d_mid = [x for x in L if x['brand'] in ('Adventure Menu', 'SubliMate', 'Qiaoshanmei')]
        d_ua = [x for x in L if x['brand'] in ('James Cook', 'Їжа в Похід', 'Харчі', '!FEST')]
        assert len(d_imp) + len(d_mid) + len(d_ua) == len(L)
        G = [(d_imp, "Дойпак: імпортний сублімат", "Travellunch, Trek'n Eat, Mountain House, Adventure Food.", None),
             (d_mid, "Дойпак: Adventure Menu, SubliMate, Qiaoshanmei", "Adventure Menu — єдиний бренд у двох форматах (пауч 400 г і сублімат 110 г). Qiaoshanmei — сухий рис під окріп, не сублімат.", None),
             (d_ua, "Дойпак: український сублімат", "James Cook, Харчі, Їжа в Похід, !FEST — вага 80–100 г.", None)]
    else:
        G = [(L, "Коробка з нагрівачем", "Haidilao, Zihaiguo, Rongcheng Haoji, Forestia — без техніки й окропу, 15 хв.", None)]
    n = len(G)
    for i, (items, title, dek, note) in enumerate(G):
        eb = f"{tag} {i + 1} З {n}" if n > 1 else tag
        _catalog(items, eb, title, dek, note=note, note_c=FMT_COL[f])


# ── де ми стоїмо ────────────────────────────────────────────────────────────
def _brand_stat(brand, fmt=None, seg=None):
    L_ = [x for x in SKU if x['brand'] == brand and (fmt is None or x['fmt'] == fmt) and (seg is None or x['seg'] == seg)]
    return min(x['lo'] for x in L_), st.median([x['lo'] for x in L_]), max(x['hi'] for x in L_)


_ben = _brand_stat("Ben's Original"); _hn = _brand_stat('Henan'); _ot = _brand_stat('Ottogi'); _soc = _brand_stat('Seeds of Change')
_ua = [x for x in SKU if x['fmt'] == 'pouch' and COUNTRY[x['brand']] == 'Україна']
_uam = st.median([x['lo'] for x in _ua])
_cm = {f: st.median([x['lo'] for x in by_fmt(f)]) for f in FMT_ORDER}
STAND = [("Ben's Original · найдешевший 250 г", _ben[0], 'm'), ("Український реторт 350 г · найдешевший", min(x['lo'] for x in _ua), 'm'),
         ("Henan · найдешевша чаша", _hn[0], 'm'), ("Український реторт 350 г · медіана", _uam, 'm'),
         ("Ben's Original · медіана", _ben[1], 'm'), ("ПАУЧ · медіана категорії", _cm['pouch'], 'c'),
         ("НАШ · BSCM пауч 150 г", 133, 'o'), ("НАШ · BSCM стакан 150 г", 139, 'o'),
         ("Henan · медіана", _hn[1], 'm'), ("НАШ · BSCM пауч 200 г", 150, 'o'), ("НАШ · ONE'S лоток 210 г", 153, 'o'),
         ("НАШ · BSCM пауч 240 г", 157, 'o'), ("Bibigo · лоток 210 г", _brand_stat('Bibigo')[0], 'm'),
         ("НАШ · CM Premium пауч 250 г органік", 178, 'o'), ("Ben's Original · найдорожчий", _ben[2], 'm'),
         ("НАШ · BSCM подвійний стакан 2×125 г", 219, 'o'), ("ЧАША · медіана категорії", _cm['cup'], 'c'),
         ("Seeds of Change · медіана", _soc[1], 'm'), ("Ottogi · медіана", _ot[1], 'm'),
         ("ДОЙПАК · медіана категорії", _cm['doypack'], 'c'), ("НАШ · ONE'S fried rice 200 г", 332, 'o'),
         ("Ottogi · найдорожчий", _ot[2], 'm'), ("КОРОБКА · медіана категорії", _cm['box'], 'c'),
         ("Mo Xiao Xian · найдорожча чаша", 931, 'm'), ("Haidilao · максимум ринку", 1190, 'm')]
STAND = [(a_, round(b_), c_) for a_, b_, c_ in STAND]
STAND.sort(key=lambda r: r[1])


def s_stand():
    s = new_slide()
    header(s, "ЦІНА ЗА УПАКОВКУ", "Де ми стоїмо", "Ринкові якорі й наші позиції на одній шкалі. Помаранчеве — ми.")
    y0, y1 = 1.86, 6.82
    rh = (y1 - y0) / len(STAND)
    AX, AW, AMAX = 4.95, 6.6, 400
    ax = lambda v: AX + AW * min(v, AMAX) / AMAX
    for t in (0, 100, 200, 300, 400):
        rect(s, ax(t), y0, 0.012, y1 - y0, MIST_D)
        text(s, ax(t) - 0.4, y1 + 0.03, 0.8, 0.2, num(t), size=10, color=GREY, align=PP_ALIGN.CENTER)
    rect(s, ax(45), y0, ax(180) - ax(45), y1 - y0, RGBColor(0xFF, 0xF6, 0xE3))
    text(s, ax(45), y0 - 0.22, 3.2, 0.2, "вікно мережевої полиці 45–180 грн", size=9.5, bold=True, color=ORANGE)
    for i, (nm, v, k) in enumerate(STAND):
        y = y0 + i * rh; cy = y + rh / 2
        col = {'m': LGREY, 'c': NAVY, 'o': ORANGE}[k]
        text(s, M, cy - 0.11, AX - M - 0.15, 0.24, nm, size=10.5 if k != 'o' else 11, bold=(k != 'm'),
             color=NAVY if k != 'm' else INK, align=PP_ALIGN.RIGHT)
        bh = rh * 0.62
        rect(s, AX, cy - bh / 2, ax(v) - AX, bh, col, rounded=True, adj=0.2)
        if v > AMAX:
            rect(s, ax(AMAX) - 0.55, cy - bh / 2 - 0.02, 0.07, bh + 0.04, WHITE)
            rect(s, ax(AMAX) - 0.42, cy - bh / 2 - 0.02, 0.07, bh + 0.04, WHITE)
        text(s, ax(v) + 0.1 if v <= AMAX else ax(AMAX) + 0.1, cy - 0.11, 1.2, 0.24, f"{num(v)} грн", size=10.5, bold=(k != 'm'),
             color=NAVY)
    foot(s, "Ціни за одну упаковку на полиці. Наші позиції — полиця = собівартість × 3,5 (20' навалом, курс 45, бонус мережі 25 %, "
            "маржа 35 %, з ПДВ). Довгі стовпчики (понад 400 грн) скорочено; значення підписано.")


FIN = [("BSCM стакан 150 г", 13, 0.500, TEAL), ("BSCM пауч 150 г", 5, 0.472, GOLD), ("BSCM пауч 200 г", 9, 0.550, GOLD),
       ("BSCM пауч 240 г", 1, 0.583, GOLD), ("CM Premium 250 г", 26, 0.667, GREEN),
       ("ONE'S білий 210 г", 30, 0.571, PLUM), ("ONE'S fried 200 г", 27, 1.238, ROSE)]
RATE = 45


def s_fin():
    import json
    P = {p['n']: p for p in json.load(open(V3 + 'pricing_rows.json'))}
    s = new_slide()
    header(s, "ЦІНОУТВОРЕННЯ", "З чого складається ціна на полиці",
           "Кожна гривня на полиці розкладена: закупівля, наша маржа, бонус мережі й націнка магазину. FOB — окрема цифра.")
    parts = [("FOB постачальника", NAVY_L), ("Доставка, мито, витрати до складу", SLATE), ("Бонус мережі", GOLD),
             ("Наша маржа", ORANGE), ("Націнка магазину", LGREY)]
    # ── верхня смуга: 100 грн на полиці
    rect(s, M, 1.80, 11.98, 1.12, MIST, rounded=True, adj=0.1)
    text(s, M + 0.25, 1.88, 6, 0.24, "З КОЖНИХ 100 ГРН НА ПОЛИЦІ", size=10, bold=True, color=GREY)
    sh = [28.6 * 0.58, 28.6 * 0.42, 17.9, 25.0, 28.6]   # FOB і доставка — середня частка за моделлю (FOB ≈ 58 % собівартості)
    xx = M + 0.25; full = 11.48
    cost_share = 0.4 / 1.4
    fobs = []
    for nm_, n_, fob_, _ in FIN:
        fobs.append(fob_ * RATE / P[n_]['cc_uah'])
    fshare = sum(fobs) / len(fobs)
    sh = [cost_share * fshare * 100, cost_share * (1 - fshare) * 100, 25 / 1.4, 35 / 1.4, 40 / 1.4]
    for (lb, col), v in zip(parts, sh):
        ww = full * v / 100
        rect(s, xx, 2.18, ww, 0.40, col)
        text(s, xx, 2.18, ww, 0.40, f"{num(v)} %", size=12.5, bold=True, color=NAVY if col in (LGREY, ORANGE) else WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(s, xx + 0.04, 2.62, ww - 0.06, 0.24, lb, size=9, bold=True, color=INK, align=PP_ALIGN.CENTER)
        xx += ww
    # ── по товарах
    y0 = 3.10; rh = 0.46
    hx = [M, M + 2.5, M + 9.15, M + 10.35]
    text(s, hx[0], y0 - 0.02, 2.4, 0.22, "ТОВАР", size=8.5, bold=True, color=GREY)
    text(s, hx[1], y0 - 0.02, 6.4, 0.22, "ПОЛИЦЯ, ГРН ЗА ОДНУ УПАКОВКУ — З ЧОГО СКЛАДАЄТЬСЯ", size=8.5, bold=True, color=GREY)
    text(s, hx[2], y0 - 0.02, 1.2, 0.22, "FOB", size=8.5, bold=True, color=GREY, align=PP_ALIGN.RIGHT)
    text(s, hx[3], y0 - 0.02, 1.6, 0.22, "ПОЛИЦЯ / FOB", size=8.5, bold=True, color=GREY, align=PP_ALIGN.RIGHT)
    scale = 6.1 / 340.0
    y = y0 + 0.26
    for i, (nm, n_, fob, c) in enumerate(FIN):
        p = P[n_]; cc = p['cc_uah']; partner = p['partner']; shelf = p['shelf']
        fob_uah = fob * RATE
        segs = [fob_uah, cc - fob_uah, 0.25 * partner, 0.35 * partner, shelf - partner]
        if i % 2 == 0: rect(s, M, y - 0.03, 11.98, rh - 0.02, MIST, rounded=True, adj=0.2)
        text(s, hx[0] + 0.12, y + 0.08, 2.3, 0.3, nm, size=12, bold=True, color=NAVY)
        xx = hx[1]
        for (lb, col), v in zip(parts, segs):
            ww = v * scale
            rect(s, xx, y + 0.06, ww, 0.32, col)
            if ww > 0.38:
                text(s, xx, y + 0.06, ww, 0.32, num(v), size=10, bold=True, color=NAVY if col in (LGREY, ORANGE) else WHITE,
                     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            xx += ww
        text(s, xx + 0.1, y + 0.06, 0.9, 0.32, f"{num(shelf)}", size=13, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
        text(s, hx[2], y + 0.08, 1.2, 0.3, f"${fob:.3f}", size=11.5, bold=True, color=NAVY, align=PP_ALIGN.RIGHT)
        text(s, hx[3], y + 0.08, 1.6, 0.3, f"× {num(shelf / fob_uah, 1)}", size=12, bold=True, color=ORANGE, align=PP_ALIGN.RIGHT)
        y += rh
    # легенда
    lx = M
    for lb, col in parts:
        rect(s, lx, y + 0.14, 0.15, 0.15, col, rounded=True, adj=0.3)
        text(s, lx + 0.22, y + 0.10, 2.4, 0.24, lb, size=9.5, color=INK); lx += 0.5 + 0.075 * len(lb)
    foot(s, "Фінмодель покупця: курс 45 грн/$; собівартість (СС) — FOB плюс доставка, мито й витрати до складу; ціна партнеру = СС ÷ 0,4 "
            "(бонус мережі 25 % і наша маржа 35 % від неї); полиця = ціна партнеру × 1,4. FOB — комерційні пропозиції; для ONE'S — оцінка "
            "із собівартості ÷ 1,7. Усе з ПДВ, за одну упаковку.")


# ── таблиця полиці ──────────────────────────────────────────────────────────
NAMES = {1: "Jasmine", 2: "Mexican", 3: "Japanese Style", 4: "Hot & Spicy", 5: "Jasmine", 6: "Mexican", 7: "Japanese Style",
         8: "Hot & Spicy", 9: "Jasmine", 10: "Mexican", 11: "Japanese Style", 12: "Hot & Spicy",
         13: "Thai Jasmine", 14: "Thai Brown Jasmine", 15: "Thai Red Jasmine", 16: "Rice Berry",
         31: "Brown Jasmine + Red Quinoa", 17: "Jasmine · з олією", 18: "Jasmine Pathum · з олією",
         19: "Brown Jasmine · з олією", 20: "Brown Rice · з олією", 21: "Basmati* · з олією", 22: "Long Grain · з олією",
         32: "Jasmine · без олії", 23: "Garlic Fried Rice", 24: "Cilantro Lime Rice", 25: "Turmeric Basmati + Cumin",
         26: "Jasmine Rice", 27: "Kimchi Fried Rice", 28: "Vegetable Fried Rice", 29: "Japchae Fried Rice",
         30: "Cooked White Rice"}
TBL_L = [("ПАУЧ · РЕТОРТ 240 г", "BSCM Foods", GOLD, [1, 2, 3, 4]), ("ПАУЧ · РЕТОРТ 200 г", "BSCM Foods", GOLD, [9, 10, 11, 12]),
         ("ПАУЧ · РЕТОРТ 150 г", "BSCM Foods", GOLD, [5, 6, 7, 8]),
         ("ПАУЧ · РЕТОРТ 250 г, ОРГАНІК", "CM Premium · Chefrey", GREEN, [23, 24, 25, 26])]
TBL_R = [("СТАКАН 150 г", "BSCM Foods", TEAL, [13, 14, 15, 16, 31]),
         ("ПОДВІЙНИЙ СТАКАН 2×125 г", "BSCM Foods", TEAL, [17, 18, 19, 20, 21, 22, 32]),
         ("ЛОТОК 200–210 г", "ONE'S International", PLUM, [30, 27, 28, 29])]


def _tbl_block(s, x, w, blocks, P):
    cols = [("СМАК", w - 3.0), ("ВАГА", 0.62), ("ПАРТНЕРУ", 1.2), ("ПОЛИЦЯ", 1.18)]
    xx = x
    for lb, cw in cols:
        text(s, xx + (0.1 if lb == "СМАК" else 0), 1.84, cw - 0.1, 0.2, lb, size=8.5, bold=True, color=GREY,
             align=PP_ALIGN.LEFT if lb == "СМАК" else PP_ALIGN.RIGHT)
        xx += cw
    y = 2.10; rh = 0.208
    for gname, brand, c, ids in blocks:
        rect(s, x, y, w, 0.27, c, rounded=True, adj=0.3)
        text(s, x + 0.12, y, w * 0.6, 0.27, gname, size=9.5, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
        text(s, x + w * 0.5, y, w * 0.5 - 0.12, 0.27, brand, size=9, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE,
             align=PP_ALIGN.RIGHT)
        y += 0.30
        for k, n in enumerate(ids):
            row = next(p for p in P if p['n'] == n)
            if k % 2 == 0: rect(s, x, y, w, rh, MIST, rounded=True, adj=0.25)
            vals = [NAMES[n], f"{num(row['g'])} г", f"{num(row['partner'])}", f"{num(row['shelf'])} грн"]
            xx = x
            for j, ((lb, cw), v) in enumerate(zip(cols, vals)):
                text(s, xx + (0.1 if j == 0 else 0), y + 0.01, cw - 0.12, rh - 0.02, v, size=10 if j < 3 else 10.5,
                     bold=(j == 3), color=c if j == 3 else INK,
                     align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.RIGHT)
                xx += cw
            y += rh
        y += 0.08
    return y


def s_price_table():
    import json
    P = json.load(open(V3 + 'pricing_rows.json'))
    s = new_slide()
    header(s, "ПОЛИЦЯ", "Наша ціна на полиці: усі формати й смаки",
           "32 позиції трьох брендів. Ціна партнеру й роздрібна полиця — з ПДВ, за одну упаковку.")
    w = (11.98 - 0.3) / 2
    _tbl_block(s, M, w, TBL_L, P)
    _tbl_block(s, M + w + 0.3, w, TBL_R, P)
    foot(s, "Фінмодель покупця (20' FCL навалом): курс 45 грн/$, бонус мережі 25 %, маржа 35 %. Ціна партнеру = собівартість ÷ 0,4, "
            "полиця = ціна партнеру × 1,4. Варіанти змішаного контейнера (Bags+Cups) не показано. Basmati* — за комерційною пропозицією.")


# ── ритейл-аудит (два слайди) ────────────────────────────────────────────────
AUDIT_CHAINS = [  # (мережа, кімнатне, охолоджене, суміші, примітка)
    ("Сільпо", 11, 8, 2, "Ben's ×11 · Чіл Міл ×2 · боули ×6"),
    ("МегаМаркет · Ultramarket", 1, 1, 2, "The Local Food · рис японський"),
    ("NOVUS", 0, 1, 2, "«Майстри Смаку»"),
    ("Ашан", 0, 1, 1, "«Легко!»"),
    ("METRO", 0, 0, 0, "лише крупа 1–10 кг: Metro Chef, Aro"),
    ("ЕКО маркет", 0, 0, 0, "нічого не знайдено"),
    ("Epicentr", 0, 0, 0, "нічого не знайдено"),
    ("Космос", 0, 0, 0, "нічого не знайдено")]


def s_retail_matrix():
    s = new_slide()
    header(s, "РИТЕЙЛ-АУДИТ · МЕРЕЖІ", "Що з готового рису стоїть у мережах",
           "Каталоги 9 мереж, 28.09.2026: усього 26 позицій, а готового рису кімнатного зберігання — лише 12.")
    tw = 7.75
    cols = [("МЕРЕЖА", 2.55), ("КІМНАТНЕ ЗБЕРІГАННЯ", 1.45), ("ОХОЛОДЖЕНА КУЛІНАРІЯ", 1.45), ("СУМІШІ ПІД ВАРІННЯ", 1.4), ("", 0.9)]
    xx = M
    for lb, w_ in cols:
        text(s, xx + 0.1, 1.80, w_ - 0.12, 0.36, lb, size=8.5, bold=True, color=GREY, line=1.0,
             align=PP_ALIGN.LEFT if lb == "МЕРЕЖА" else PP_ALIGN.CENTER)
        xx += w_
    y = 2.22; rh = 0.40
    for i, (nm, a_, c_, m_, note) in enumerate(AUDIT_CHAINS):
        rect(s, M, y, tw, rh - 0.04, MIST if i % 2 == 0 else WHITE, rounded=True, adj=0.2)
        text(s, M + 0.14, y + 0.02, 2.45, 0.22, nm, size=12.5, bold=True, color=NAVY)
        text(s, M + 0.14, y + 0.21, 2.45, 0.16, note, size=8, color=GREY)
        xx = M + cols[0][1]
        for v, (lb, w_), col in zip((a_, c_, m_), cols[1:4], (TEAL, PLUM, GOLD)):
            if v:
                rect(s, xx + w_ / 2 - 0.30, y + 0.04, 0.60, rh - 0.12, col, rounded=True, adj=0.3)
                text(s, xx, y + 0.04, w_, rh - 0.12, str(v), size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            else:
                text(s, xx, y + 0.04, w_, rh - 0.12, "0", size=15, bold=True, color=LGREY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            xx += w_
        y += rh
    tot = [12, 11, 3]   # унікальні SKU (Жменька й Trapeza стоять у кількох мережах)
    rect(s, M, y + 0.02, tw, 0.40, NAVY, rounded=True, adj=0.2)
    text(s, M + 0.14, y + 0.02, 2.4, 0.40, "Унікальних SKU", size=12.5, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    xx = M + cols[0][1]
    for v, (lb, w_) in zip(tot, cols[1:4]):
        text(s, xx, y + 0.02, w_, 0.40, str(v), size=16, bold=True, color=AMBER, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        xx += w_
    # праворуч: Ben's у Сільпо
    rx = M + tw + 0.22; rw = W - M - rx
    bens = sorted([x for x in SKU if x['brand'] == "Ben's Original" and 'Сільпо' in x['seller']], key=lambda q: (q['lo'], q['g']))
    rect(s, rx, 1.80, rw, 4.08, WHITE, line=MIST_D, lw=1.2, rounded=True, adj=0.03)
    text(s, rx + 0.22, 1.90, rw - 0.4, 0.26, f"BEN'S ORIGINAL У «СІЛЬПО» · {len(bens)} ПОЗИЦІЙ", size=10, bold=True, color=GREY)
    yy = 2.24
    for q in bens:
        text(s, rx + 0.22, yy, 2.1, 0.24, q['name'], size=10.5, color=INK)
        text(s, rx + 2.15, yy, 0.6, 0.24, f"{num(q['g'])} г", size=10, color=GREY)
        text(s, rx + rw - 1.3, yy, 1.1, 0.24, f"{num(q['lo'])} грн", size=11, bold=True, color=NAVY, align=PP_ALIGN.RIGHT)
        yy += 0.318
    # нижня смуга
    by = 6.0
    cw = (tw - 0.3) / 3
    cards = [("12", "SKU кімнатного зберігання на всю країну", TEAL), ("11 з 12", "це Ben's Original в одній мережі", GOLD),
             ("0", "у METRO, Ашан, NOVUS, ЕКО маркет, Epicentr, Космос", ROSE)]
    for k, (big, lb, col) in enumerate(cards):
        x = M + k * (cw + 0.15)
        rect(s, x, by, cw, 0.94, MIST, rounded=True, adj=0.08); rect(s, x, by, 0.08, 0.94, col, rounded=True, adj=0.5)
        text(s, x + 0.24, by + 0.08, cw - 0.3, 0.4, big, size=22, bold=True, color=NAVY)
        text(s, x + 0.24, by + 0.52, cw - 0.3, 0.42, lb, size=9.5, color=INK, line=1.1)
    insight(s, rx, by, rw, 0.94, None, "Ben's у «Сільпо» — єдиний мережевий гарнір; інших немає.", ORANGE, size=11)
    foot(s, "Каталоги мереж на 28.09.2026: Ашан, Сільпо, METRO, NOVUS, МегаМаркет, Ultramarket (одна платформа з МегаМаркет), ЕКО маркет, Epicentr, Космос. "
            "Кімнатне зберігання — готовий рис у паучі/лотку; охолоджена — кулінарія 2–5 діб; суміші — сира крупа в пакеті для варіння.")


def s_retail_table():
    s = new_slide()
    header(s, "РИТЕЙЛ-АУДИТ · ПОЗИЦІЇ", "Що саме стоїть на полицях",
           "Усі 26 знайдених позицій: де, який формат, скільки SKU і яка ціна за упаковку.")
    cols = [("МЕРЕЖА", 2.75), ("ЩО САМЕ СТОЇТЬ", 4.2), ("ФОРМАТ", 1.55), ("SKU", 0.6), ("ЦІНА, ГРН", 1.45), ("ПОХОДЖЕННЯ", 1.43)]
    xx = M
    for lb, w_ in cols:
        text(s, xx + 0.12, 1.78, w_, 0.2, lb, size=8.5, bold=True, color=GREY, align=PP_ALIGN.LEFT if lb != "SKU" else PP_ALIGN.CENTER)
        xx += w_
    G = [("КІМНАТНЕ ЗБЕРІГАННЯ · НАШ СЕТ", "12 SKU", TEAL,
          [("Сільпо", "Ben's Original / Uncle Ben's Express", "пауч 220–250 г", 11, "45–179", "Mars · ЄС"),
           ("МегаМаркет · Ultramarket", "The Local Food, по-тайськи з куркою", "лоток 350 г", 1, "207", "Україна")]),
         ("ОХОЛОДЖЕНА КУЛІНАРІЯ · 2–5 ДІБ", "11 SKU", PLUM,
          [("Ашан", "«Легко!» рис з овочами та курячим фрікасе", "лоток 280 г", 1, "112", "Україна"),
           ("Сільпо", "«Чіл Міл» рис паровий · рис з яйцем і овочами", "лоток", 2, "75", "Україна"),
           ("Сільпо", "Боули з рисом · теріякі · рис по-китайськи", "кулінарія", 6, "139–439", "Україна"),
           ("NOVUS", "«Майстри Смаку» рис з овочами", "кулінарія", 1, "299", "Україна"),
           ("МегаМаркет · Ultramarket", "Рис японський з овочами", "кулінарія", 1, "366", "Україна")]),
         ("СУМІШ ПІД ВАРІННЯ · НЕ ГОТОВИЙ ПРОДУКТ", "3 SKU", GOLD,
          [("Ашан · NOVUS · МегаМаркет · Сільпо", "«Жменька» Tasty Mix басматі з овочами", "пакет 200 г", 1, "89–92", "Україна"),
           ("NOVUS · Сільпо", "Trapeza басматі з карі по-індійськи", "пакет 250 г", 1, "59–85", "Україна"),
           ("МегаМаркет · Ultramarket", "«Зерновита» рис з овочами", "пакет 500 г", 1, "78", "Україна")])]
    y = 2.04; rh = 0.355
    for gname, cnt, col, items in G:
        rect(s, M, y, 11.98, 0.32, col, rounded=True, adj=0.3)
        text(s, M + 0.16, y, 8, 0.32, gname, size=11, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
        text(s, M + 9, y, 2.8, 0.32, cnt, size=11, bold=True, color=WHITE, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
        y += 0.36
        for i, row in enumerate(items):
            if i % 2 == 0: rect(s, M, y, 11.98, rh - 0.02, MIST, rounded=True, adj=0.2)
            xx = M
            for j, ((lb, w_), v) in enumerate(zip(cols, row)):
                text(s, xx + 0.12, y + 0.045, w_ - 0.16, rh - 0.06, str(v), size=(11.5 if len(str(v)) < 26 else 9.5) if j != 4 else 12.5, bold=(j in (0, 3, 4)),
                     color=NAVY if j in (0, 3, 4) else INK, align=PP_ALIGN.CENTER if j == 3 else PP_ALIGN.LEFT)
                xx += w_
            y += rh
        y += 0.10
    foot(s, "Ашан, Сільпо, METRO, NOVUS, МегаМаркет, Ultramarket, ЕКО маркет, Epicentr, Космос — каталоги на 28.09.2026. "
            "METRO, ЕКО маркет, Epicentr і Космос — жодної позиції цих трьох груп.")


# ── канали продажу ───────────────────────────────────────────────────────────
CH_RETAIL3 = "Сільпо · МегаМаркет · Ultramarket"
CH_RETAIL15 = ("NOVUS · METRO · Таврія В · Cosmos · Восторг · Kharkiv · Чудо-маркет · Епіцентр · Ашан · Грона · Ідеал · "
               "ОнДе · Za Raz · ЕКО маркет · Торба")
CH_MARKET = "MAUDAU · Rozetka · Prom"
CH_ASIAN = ["Смак Кореї", "Тайякі Март", "Pulsar", "Апетітаріум", "Gurmissimo", "Edison Lee", "OMG! Asia", "СНЕКІС", "DCM",
            "Скарби Азії", "AsiaStyle", "ASIA FOODS", "Wèi Māo", "AsianFoods", "Asia Goods Store", "SushiPovar", "KladezGroup",
            "Sweet Svitt", "Товари з Іспанії", "Asia Foods Trade"]
CH_OUTDOOR = ["ALANTUR", "Highlander", "ForCamp", "Гайдамака", "Freeride", "MK-Sport", "Daruy", "Tactico", "Лєєр", "Activity",
              "ВсеОпт", "Шериф", "Суренж", "OXO", "Desna", "Kalush-Craft", "SportStorm", "Klever-Shop", "Terra Incognita", "110вольт",
              "Virnyy vybir", "Vkladovke.com.ua", "Експрес Шоп", "Клуб Мандрівників", "Вояджер", "Непереможні", "Вартовий",
              "Драйв Сенс", "Ліхтар", "Іграшка Мрії", "Military Style", "Націоналіст", "Modern Shop", "Palmer", "Kamanti",
              "Mr.Fishkin", "Висот-Нік", "UA-market", "Skaut.in.ua"]
CH_FOOD = ["Харчі ТМ", "Їжа в Похід", "UPcompany", "СУХПАЙ", "Belorfoods", "МореПродуктів", "Козуб", "Смачна адреса", "Продукт-Shop",
           "Молочний склад", "Foodi Shop", "Food Shop", "Мартел-shop", "portion.com.ua", "Euro-komplekt"]


def s_channels():
    s = new_slide()
    header(s, "КАНАЛИ ПРОДАЖУ", "Де продається категорія і кому",
           "Усі магазини, де ми знайшли готовий рис і суміжні формати, — по каналах.")
    CH = [("Мережевий роздріб", 3 + 15, ROSE, "ЦІЛЬОВИЙ КАНАЛ", [("3 мережі з категорією: ", CH_RETAIL3), ("ще 15 перевірених без категорії: ", CH_RETAIL15)]),
          ("Онлайн-маркетплейси", 3, GOLD, "ЦІЛЬОВИЙ КАНАЛ", [("", CH_MARKET + ". На Prom ще 7 перепродавців iHerb (Seeds of Change).")]),
          ("Азійські й етнічні фудшопи", len(CH_ASIAN), TEAL, "ЦІЛЬОВИЙ · ВХІД", [("", " · ".join(CH_ASIAN))]),
          ("Гастро- й фірмові інтернет-магазини", len(CH_FOOD), GREEN, "СУМІЖНИЙ", [("", " · ".join(CH_FOOD))]),
          ("Туристичні й військові магазини", len(CH_OUTDOOR), PLUM, "НЕ НАШ ПОКУПЕЦЬ", [("", " · ".join(CH_OUTDOOR))])]
    # ліва колонка — стовпчики
    x0, w = M, 3.55
    rect(s, x0, 1.80, w, 5.14, MIST, rounded=True, adj=0.03)
    text(s, x0 + 0.22, 1.92, w - 0.4, 0.26, "ТОЧОК ПРОДАЖУ В КАНАЛІ", size=10, bold=True, color=GREY)
    mx = max(c[1] for c in CH)
    y = 2.26
    for lb, n, col, tag, _ in CH:
        text(s, x0 + 0.22, y, w - 0.4, 0.46, lb, size=11.5, bold=True, color=NAVY, line=1.05)
        bw_ = max((w - 1.3) * n / mx, 0.1)
        rect(s, x0 + 0.22, y + 0.50, bw_, 0.34, col, rounded=True, adj=0.25)
        text(s, x0 + 0.22 + bw_ + 0.1, y + 0.46, 0.8, 0.4, str(n), size=20, bold=True, color=NAVY)
        y += 0.95
    # права колонка — списки
    rx = x0 + w + 0.2; rw = W - M - rx
    hs = [1.02, 0.78, 1.10, 1.00, 1.34]
    y = 1.80
    for (lb, n, col, tag, body), h_ in zip(CH, hs):
        rect(s, rx, y, rw, h_ - 0.07, WHITE, line=MIST_D, lw=1.2, rounded=True, adj=0.06)
        rect(s, rx, y, 0.08, h_ - 0.07, col, rounded=True, adj=0.5)
        text(s, rx + 0.24, y + 0.07, rw - 2.6, 0.26, f"{lb} · {n}", size=11.5, bold=True, color=NAVY)
        chip(s, rx + rw - 1.9, y + 0.08, tag, col, size=8, w=1.78, h=0.22)
        paras = [([(a_, {"bold": True, "color": col}), (b_, {})], {}) for a_, b_ in body]
        text(s, rx + 0.24, y + 0.36, rw - 0.4, h_ - 0.46, paras, size=9.5, color=INK, line=1.12)
        y += h_
    foot(s, "Покупець: мережі — масовий, щоденне харчування; маркетплейси — масовий, планове замовлення; азійські магазини — знає категорію; "
            "туристичні й військові — автономне харчування. Мереж перевірено 18 (17 через zakaz.ua + Сільпо); магазини — за картками продавців 23.09–06.10.2026.")
