# ══ ГОЛОВНЕ, КАТАЛОГ, «ДЕ МИ СТОЇМО», ФІНМОДЕЛЬ, ТАБЛИЦЯ ПОЛИЦІ ══════════════
def s_kpi():
    s = new_slide()
    header(s, "ГОЛОВНЕ", "Шість цифр, які описують ринок", "Усі ціни в колоді — за одну упаковку. Опт і ящики до розрахунку не входять.")
    K = [("ПОЗИЦІЙ У ПРОДАЖУ", "84", PLUM, "23 бренди · 4 формати · 64 продавці. Кожна позиція має фото й ціну в каталозі."),
         ("НА ПОЛИЦІ МЕРЕЖ", "12", TEAL, "SKU на всю країну, 11 з них — Ben's Original у «Сільпо». Ритейл-аудит 9 мереж."),
         ("ІМПОРТ ЗА МИТНИЦЕЮ", "3,5 т", ORANGE, "Сегменти 1–2 за 16 місяців — лише 1,4 % коду 1904901000. Решта — заморожені суміші й снеки."),
         ("2026 ПРОТИ 2025", "×3", GREEN, "Наші сегменти без разової партії Ben's: 0,43 → 1,30 т (однакові місяці)."),
         ("НАЙПОШИРЕНІШИЙ ФОРМАТ", "дойпак", ROSE, "30 позицій (36 %), але це туризм і армія, а не мережева полиця."),
         ("ВХІД У КАТЕГОРІЮ", "109 грн", GOLD, "Медіана пауча — найдешевший формат: 19 позицій, вага 220–400 г.")]
    cw = (11.98 - 2 * 0.18) / 3; ch = 2.46
    for i, (lb, v, c, note) in enumerate(K):
        x = M + (i % 3) * (cw + 0.18); y = 1.86 + (i // 3) * (ch + 0.18)
        rect(s, x, y, cw, ch, MIST, rounded=True, adj=0.06); rect(s, x, y, cw, 0.10, c, rounded=True, adj=0.5)
        text(s, x + 0.3, y + 0.34, cw - 0.6, 0.24, lb, size=10.5, bold=True, color=GREY)
        text(s, x + 0.3, y + 0.70, cw - 0.6, 0.8, v, size=44, bold=True, color=c)
        text(s, x + 0.3, y + 1.62, cw - 0.6, 0.8, note, size=11, color=INK, line=1.2)
    foot(s, "Джерела: картки продавців 23–28.09.2026 (84 позиції); митна база codexmb, вивантаження 28.09.2026; ритейл-аудит 9 мереж, 28.09.2026.")


def _catalog(items, eyebrow, title, dek, note=None, cols=None, note_c=ORANGE):
    s = new_slide()
    header(s, eyebrow, title, dek)
    n = len(items)
    two = n > 7
    cols = cols or (math.ceil(n / 2) if two else n)
    gap = 0.12
    w = (11.98 - gap * (cols - 1)) / cols
    top = 1.86
    avail = 6.94 - top - (1.12 if note else 0)
    h = (avail - (gap if two else 0)) / (2 if two else 1)
    h = min(h, 3.9)
    for i, x in enumerate(items):
        r, ccol = divmod(i, cols)
        sku_cell(s, M + ccol * (w + gap), top + r * (h + gap), w, h, x['img'], x['name'], x['g'], x['lo'], x['hi'],
                 x['seller'], FMT_COL[x['fmt']] if False else SEG_COL[x['seg']], brand=x['brand'])
    if note:
        yb = top + (2 if two else 1) * h + (gap if two else 0) + 0.14
        insight(s, M, yb, 11.98, min(6.94 - yb, 0.98), None, note, note_c, size=11)
    foot(s, "Фото, вага й ціна — з карток продавців; ціна за одну упаковку (від–до серед продавців). Колір підпису: золотий — розігрів, бірюзовий — під окріп, рожевий — саморозігрів, фіолетовий — сублімат.")
    return s


def cat_slides():
    P = by_fmt('pouch'); C = by_fmt('cup'); D = by_fmt('doypack'); B = by_fmt('box')
    ben = [x for x in P if x['brand'] == "Ben's Original"]
    p_oth = [x for x in P if x['brand'] != "Ben's Original"]
    cup_a = [x for x in C if x['seg'] == 1]; cup_b = [x for x in C if x['seg'] != 1]
    d_imp = [x for x in D if x['brand'] in ('Travellunch', "Trek'n Eat", 'Mountain House', 'Adventure Food')]
    d_mid = [x for x in D if x['brand'] in ('Adventure Menu', 'SubliMate', 'Qiaoshanmei')]
    d_ua = [x for x in D if x['brand'] in ('James Cook', 'Їжа в Похід', 'Харчі', '!FEST')]
    assert len(ben) + len(p_oth) == 19 and len(cup_a) + len(cup_b) == 23 and len(d_imp) + len(d_mid) + len(d_ua) == 30
    _catalog(ben, "ДОДАТОК · КАТАЛОГ · ПАУЧ", "Пауч: Ben's Original", "13 позицій, 220–250 г — єдиний бренд на мережевій полиці.")
    _catalog(p_oth, "ДОДАТОК · КАТАЛОГ · ПАУЧ", "Пауч: інші бренди", "Органіка, український реторт і туристичний Adventure Menu.",
             note="У пауча поза Ben's — лише шість позицій у трьох вагах: 250, 350 і 400 г. Пауч 150–199 г і 300 г не представлений.", note_c=GOLD)
    _catalog(cup_a, "ДОДАТОК · КАТАЛОГ · ЧАША", "Чаша: готовий рис", "Ottogi і Bibigo — розігрів у мікрохвильовці 60–90 с.")
    _catalog(cup_b, "ДОДАТОК · КАТАЛОГ · ЧАША", "Чаша: під окріп і саморозігрів", "Henan, Gallina Blanca — окріп 5–10 хв; Mo Xiao Xian — нагрівач.")
    _catalog(B, "ДОДАТОК · КАТАЛОГ · КОРОБКА", "Коробка з нагрівачем", "Haidilao, Zihaiguo, Rongcheng Haoji — без техніки й окропу, 15 хв.")
    _catalog(d_imp, "ДОДАТОК · КАТАЛОГ · ДОЙПАК", "Дойпак: імпортний сублімат", "Travellunch, Trek'n Eat, Mountain House, Adventure Food.")
    _catalog(d_mid, "ДОДАТОК · КАТАЛОГ · ДОЙПАК", "Дойпак: Adventure Menu, SubliMate, Qiaoshanmei",
             "Сублімат 110–140 г і китайський рис у плоскому пакеті під окріп.",
             note="Adventure Menu — єдиний бренд із двома форматами: готова страва в пауці 400 г і сублімат у дойпаку 110 г. "
                  "Qiaoshanmei — не сублімат, а сухий рис під окріп.", note_c=PLUM)
    _catalog(d_ua, "ДОДАТОК · КАТАЛОГ · ДОЙПАК", "Дойпак: український сублімат", "James Cook, Харчі, Їжа в Похід, !FEST — вага 80–100 г.")


# ── де ми стоїмо ────────────────────────────────────────────────────────────
STAND = [("Ben's Original · найдешевший 250 г", 45, 'm'), ("Henan · найдешевша чаша 144 г", 83, 'm'),
         ("Ben's Original · медіана", 90, 'm'), ("Маркел · реторт 350 г", 95, 'm'), ("Portion · реторт 350 г", 109, 'm'),
         ("ПАУЧ · медіана категорії", 109, 'c'), ("НАШ · BSCM пауч 150 г", 133, 'o'), ("НАШ · BSCM стакан 150 г", 139, 'o'),
         ("Henan · медіана", 140, 'm'), ("НАШ · BSCM пауч 200 г", 150, 'o'), ("НАШ · ONE'S лоток 210 г", 153, 'o'),
         ("НАШ · BSCM пауч 240 г", 157, 'o'), ("Bibigo · лоток 210 г", 162, 'm'),
         ("НАШ · CM Premium пауч 250 г органік", 178, 'o'), ("Ben's Original · найдорожчий 220 г", 179, 'm'),
         ("НАШ · BSCM подвійний стакан 2×125 г", 219, 'o'), ("ЧАША · медіана категорії", 229, 'c'),
         ("Ottogi · медіана", 255, 'm'), ("НАШ · ONE'S fried rice 200 г", 332, 'o'), ("ДОЙПАК · медіана категорії", 315, 'c'),
         ("Ottogi · найдорожчий", 356, 'm'), ("КОРОБКА · медіана категорії", 708, 'c'),
         ("Mo Xiao Xian · найдорожча чаша", 931, 'm'), ("Haidilao · максимум ринку", 1190, 'm')]
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


FIN = [("BSCM стакан 150 г", 139, 0.500, TEAL), ("BSCM пауч 150 г", 133, 0.472, GOLD), ("BSCM пауч 200 г", 150, 0.550, GOLD),
       ("BSCM пауч 240 г", 157, 0.583, GOLD), ("CM Premium 250 г", 178, 0.667, GREEN),
       ("ONE'S білий 210 г", 153, 0.571, PLUM), ("ONE'S fried 200 г", 332, 1.238, ROSE)]


def s_fin():
    s = new_slide()
    header(s, "ФІНМОДЕЛЬ ПРОТИ РИНКУ", "Наша полиця й FOB",
           "Полиця = собівартість × 3,5 за поточних умов. Вертикальні лінії — ринкові орієнтири.")
    y0 = 2.42; rh = 0.60; AX, AW, AMAX = 3.5, 7.6, 360
    ax = lambda v: AX + AW * v / AMAX
    refs = [(90, "Ben's 90"), (109, "пауч 109"), (162, "Bibigo 162"), (199, "Ottogi −22 % · 199")]
    for k, (v, lb) in enumerate(refs):
        rect(s, ax(v) - 0.01, y0 - 0.1, 0.02, rh * 7 + 0.15, NAVY_L if k == 1 else LGREY)
        text(s, ax(v) - 0.9, y0 - 0.46 + (0.2 if k % 2 else 0), 1.8, 0.2, lb, size=9.5, bold=(k == 1), color=NAVY,
             align=PP_ALIGN.CENTER)
    for i, (nm, v, fob, c) in enumerate(FIN):
        y = y0 + i * rh
        text(s, M, y + 0.12, AX - M - 0.15, 0.3, nm, size=13, bold=True, color=NAVY, align=PP_ALIGN.RIGHT)
        rect(s, AX, y + 0.08, ax(v) - AX, 0.38, c, rounded=True, adj=0.18)
        text(s, ax(v) + 0.12, y + 0.10, 3.2, 0.34, [([(f"{v} грн", {"bold": True, "size": 13, "color": NAVY}),
                                                      (f"   FOB ${fob:.3f}", {"size": 11.5, "color": GREY})], {})])
    for k in (0, 100, 200, 300):
        text(s, ax(k) - 0.4, y0 + 7 * rh + 0.08, 0.8, 0.2, str(k), size=10, color=GREY, align=PP_ALIGN.CENTER)
    text(s, AX, y0 + 7 * rh + 0.3, 8.3, 0.2, "ПОЛИЦЯ ЗА ПОТОЧНОЇ МОДЕЛІ, ГРН · 20' FCL НАВАЛОМ", size=9, bold=True, color=GREY)
    foot(s, "Фінмодель: курс 45 грн/$, бонус мережі 25 %, цільова маржа 35 %, націнка мережі ×1,4 (ціна партнеру з ПДВ). "
            "FOB — комерційні пропозиції; для ONE'S оцінка із собівартості. Перевірено поштучно.")


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


def channels_chart(s):
    for sh in list(s.shapes):
        if sh.shape_type == 13 and sh.width / 914400 > 3:
            sh._element.getparent().remove(sh._element)
    x0, w = M, 5.55
    rect(s, x0, 1.62, w, 5.30, MIST, rounded=True, adj=0.03)
    text(s, x0 + 0.25, 1.78, w - 0.4, 0.26, "ТОЧОК ПРОДАЖУ В КАНАЛІ", size=10.5, bold=True, color=GREY)
    ch = [("Туристичні й військові", 40, PLUM, "не наш покупець"), ("Азійські й етнічні фудшопи", 20, TEAL, "наш вхід"),
          ("Онлайн-маркетплейси", 3, GOLD, "MAUDAU, Rozetka, Prom"), ("Мережевий роздріб", 3, ROSE, "Сільпо, МегаМаркет, Ultramarket")]
    y = 2.18
    for lb, n, c, nt in ch:
        text(s, x0 + 0.25, y, w - 0.4, 0.3, lb, size=13, bold=True, color=NAVY)
        text(s, x0 + 0.25, y + 0.30, w - 0.4, 0.22, nt, size=9.5, color=GREY)
        rect(s, x0 + 0.25, y + 0.58, max((w - 1.4) * n / 40, 0.12), 0.46, c, rounded=True, adj=0.2)
        text(s, x0 + 0.25 + max((w - 1.4) * n / 40, 0.12) + 0.12, y + 0.56, 0.9, 0.5, str(n), size=24, bold=True, color=NAVY)
        y += 1.18
