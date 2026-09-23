
# ══ ФОРМАТ І ГРАМАЖ ═══════════════════════════════════════════════════════
def slide_format():
    s = slide()
    header(s, "ФОРМАТ І ГРАМАЖ",
           f"Пауч — {len(BY_PACK['pouch'])} позицій із {N_SKU}, масова вага 350 г",
           "Два формати паковання. Кожен тримає свою вагу — платформи майже не перетинаються.")
    PICS = {"pouch": "sku/ua_veres_chicken.jpg", "cup": "sku/ot_bibimbap.jpg"}
    NOTE = {"pouch": "Плаский реторт-пауч. Займає менше місця на полиці, дешевший у логістиці, "
                     "тримає 24 місяці. Формат, у якому працюють і Ben's, і всі українці.",
            "cup": "Жорстка чаша: їсти можна прямо з неї, але вона дорожча й об'ємніша. "
                   "На українському ринку — тільки корейський імпорт."}
    COL = {"pouch": ORANGE, "cup": PLUM}
    for i, k in enumerate(PACK_ORDER):
        it = BY_PACK[k]
        x = M + i * 3.07
        c = COL[k]
        pr = sorted(y[4] for y in it)
        ms = sorted(y[3] for y in it)
        rect(s, x, 1.88, 2.82, 4.94, WHITE, line=MIST_D, lw=1.2, rounded=True, adj=0.05)
        rect(s, x, 1.88, 2.82, 0.10, c, rounded=True, adj=0.5)
        rect(s, x + 0.16, 2.08, 2.50, 1.54, MIST, rounded=True, adj=0.06)
        pic(s, PICS[k], x + 0.26, 2.14, 2.30, 1.42)
        text(s, x + 0.24, 3.72, 2.40, 0.24, PACK_RU[k], size=10.5, bold=True, color=c)
        rect(s, x + 0.24, 4.02, 2.34, 0.18, MIST_D, rounded=True, adj=0.5)
        rect(s, x + 0.24, 4.02, 2.34 * len(it) / N_SKU, 0.18, c, rounded=True, adj=0.5)
        text(s, x + 0.24, 4.26, 2.40, 0.24,
             f"{pl(len(it), 'позиція', 'позиції', 'позицій')} · "
             f"{num(len(it) / N_SKU * 100)} % · "
             f"{pl(len(set(y[1] for y in it)), 'бренд', 'бренди', 'брендів')}",
             size=10.5, bold=True, color=NAVY)
        rect(s, x + 0.24, 4.58, 2.34, 0.02, MIST_D)
        for j, (lb, v) in enumerate([("ЗА УПАКОВКУ, МЕДІАНА", f"{num(st.median(pr))} грн"),
                                     ("СЕРЕДНЯ МАСА", f"{num(st.median(ms))} г"),
                                     ("ДІАПАЗОН",
                                      f"{num(min(pr))} – {num(max(y[5] for y in it))} грн"),
                                     ("МАСА", f"{num(min(ms))} – {num(max(ms))} г")]):
            yy = 4.68 + j * 0.46
            text(s, x + 0.24, yy, 1.90, 0.20, lb, size=7, bold=True, color=GREY)
            text(s, x + 0.24, yy + 0.16, 2.40, 0.22, v, size=10.5, bold=True,
                 color=c if j < 2 else INK)
        text(s, x + 0.24, 6.46, 2.40, 0.30, NOTE[k], size=7.5, color=GREY, line=1.22)
    # грамаж
    cnt = collections.Counter(x[3] for x in ALL8)
    tops = sorted(cnt.items(), key=lambda kv: kv[0])
    text(s, 6.88, 1.88, 6.00, 0.24, "ПОЗИЦІЙ У КОЖНІЙ ВАЗІ", size=9, bold=True, color=GREY)
    bx, by0 = 6.88, 4.46
    mxn = max(n for _, n in tops)
    bw = (5.78 - 0.10 * (len(tops) - 1)) / len(tops)
    for i, (g_, n) in enumerate(tops):
        x = bx + i * (bw + 0.10)
        h = 1.96 * n / mxn
        pk = collections.Counter(pack_of(y[1]) for y in ALL8 if y[3] == g_)
        c = COL[pk.most_common(1)[0][0]]
        rect(s, x, by0 - h, bw, h, c, rounded=True, adj=0.14)
        text(s, x, by0 - h - 0.24, bw, 0.22, str(n), size=11, bold=True, color=c,
             align=PP_ALIGN.CENTER)
        text(s, x, by0 + 0.06, bw, 0.22, f"{num(g_)} г", size=8, bold=True, color=INK,
             align=PP_ALIGN.CENTER)
    text(s, 6.88, 4.76, 5.90, 0.44,
         "350 г — український реторт-пауч. 250 і 220 г — Ben's Original. "
         "210–320 г — корейська чаша. 400 г — Adventure Menu.",
         size=8.5, color=GREY, line=1.24)
    insight(s, 6.88, 5.40, 5.90, 1.42, None,
            "Український пауч стоїть у вазі 350 г, Ben's Original — у 220 і 250 г, "
            "корейська чаша — у 210–320 г. Ваги майже не перетинаються.\n"
            "Вільне вікно — пауч 200–250 г українського виробництва: у цій вазі сьогодні "
            "тільки імпортний Ben's Original за 45–179 грн.", ORANGE)
    foot(s, f"Розподіл {N_SKU} позицій за типом паковання і масою з карток продавців, 23.09.2026.")


# ══ ЦІНА ЗА ОДНУ УПАКОВКУ ═════════════════════════════════════════════════
def slide_ladder():
    """Сходи брендів за ціною однієї упаковки. Крапка — медіана бренду."""
    s = slide()
    lo_all = min(x[4] for x in ALL8)
    hi_all = max(x[5] for x in ALL8)
    header(s, "ЦІНА ЗА ОДНУ УПАКОВКУ",
           f"Від {num(lo_all)} до {num(hi_all)} грн: сходи всіх {N_BRANDS} брендів",
           "Крапка — медіана бренду, смуга — від найдешевшої до найдорожчої його позиції.")
    SEGC = {}
    for _nm, _items, _c, _ in SEGMENTS:
        for _x in _items:
            SEGC[_x[1]] = (_c, _nm)
    rows = []
    for b in dict.fromkeys(x[1] for x in ALL8):
        it = [x for x in ALL8 if x[1] == b]
        rows.append((b, min(x[4] for x in it), max(x[5] for x in it),
                     st.median([x[4] for x in it]), len(it), COUNTRY[b], SEGC[b][0]))
    rows.sort(key=lambda r: r[3])
    LO, HI = 0, 500
    AX, AW = 4.70, 7.10

    def ax(v):
        return AX + AW * (v - LO) / (HI - LO)
    rect(s, M, 1.86, 11.98, 4.90, MIST, rounded=True, adj=0.04)
    for t in range(100, 501, 100):
        rect(s, ax(t), 2.10, 0.012, 3.94, RGBColor(0xDC, 0xE2, 0xEE))
        text(s, ax(t) - 0.40, 6.12, 0.80, 0.22, num(t), size=8.5, color=GREY,
             align=PP_ALIGN.CENTER)
    text(s, M + 0.26, 2.06, 2.30, 0.18, "БРЕНД", size=7, bold=True, color=GREY)
    text(s, M + 2.64, 2.06, 0.96, 0.18, "КРАЇНА", size=7, bold=True, color=GREY)
    text(s, M + 3.66, 2.06, 0.34, 0.18, "SKU", size=7, bold=True, color=GREY,
         align=PP_ALIGN.RIGHT)
    text(s, M + 0.26, 6.12, 3.20, 0.22, "ГРН ЗА ОДНУ УПАКОВКУ", size=8.5, bold=True, color=GREY)
    TOP, BOT = 2.46, 5.90
    step = min(0.372, (BOT - TOP) / max(1, len(rows) - 1))
    yy = TOP
    for b_, lo, hi, med, n, cty, c in rows:
        text(s, M + 0.26, yy - 0.11, 2.34, 0.24, b_, size=10.5, bold=True, color=NAVY)
        text(s, M + 2.64, yy - 0.09, 0.96, 0.20, cty, size=8, color=GREY)
        text(s, M + 3.60, yy - 0.09, 0.40, 0.20, str(n), size=8, bold=True, color=GREY,
             align=PP_ALIGN.RIGHT)
        x0, x1 = ax(lo), ax(hi)
        rect(s, x0, yy - 0.045, max(x1 - x0, 0.03), 0.09, c, rounded=True, adj=0.5)
        dot(s, ax(med), yy, 0.19, c)
        lbl = num(med) if lo == hi else f"{num(lo)} – {num(hi)}"
        text(s, x1 + 0.12, yy - 0.115, 1.40, 0.24, lbl, size=9, bold=True, color=c)
        yy += step
    rect(s, ax(45), 6.40, ax(200) - ax(45), 0.06, ORANGE)
    text(s, ax(45), 6.50, 4.20, 0.24, "вікно мережевої полиці: 45–200 грн", size=9,
         bold=True, color=ORANGE)
    lx = M + 0.26
    for _lb, _c in [("Гарнір", ORANGE), ("Чаша", PLUM), ("Пауч UA", GREEN),
                    ("Пауч імпорт", TEAL)]:
        dot(s, lx + 0.06, 6.52, 0.13, _c)
        w = 0.16 + 0.056 * len(_lb)
        text(s, lx + 0.18, 6.42, w + 0.10, 0.22, _lb, size=8, color=INK)
        lx += w + 0.28
    foot(s, "Роздрібна ціна за одну упаковку на сайті продавця, 23.09.2026. Оптові пороги "
            "(у частини продавців діє нижча ціна від 3–20 шт) у розрахунок не входять.")


# ══ ПОЛИЦЯ МЕРЕЖ ══════════════════════════════════════════════════════════
def slide_chain_shelf():
    s = slide()
    header(s, "ЩО СТОЇТЬ У МЕРЕЖІ ЗАМІСТЬ НАС", "Рис із м'ясом на полиці є — але в бляшанці",
           "Мережі не мають жодного пауча з готовим рисом. Натомість мають консерву за 63–160 грн.")
    for i, (im, b, nm, g, lo, hi, ch, c) in enumerate(CHAIN_SHELF):
        sku_cell(s, M + i * 2.02, 1.88, 1.92, 2.86, im, nm, g, lo, hi, ch, c, brand=b)
    insight(s, M, 4.94, 5.86, 1.94, "Чому це важливо",
            "Це прямий конкурент за той самий привід: гаряча страва з рисом і м'ясом без "
            "готування, уже на полиці мережі, з українським виробником.\n"
            "Бляшанку не можна поставити в мікрохвильовку — у цьому перевага пауча. "
            "Але орієнтир ціни задає саме вона, а не Ottogi за 252 грн.", ROSE)
    rect(s, M + 6.12, 4.94, 5.86, 1.94, MIST, rounded=True, adj=0.06)
    text(s, M + 6.40, 5.10, 5.30, 0.26, "Порівняння за упаковку", size=12, bold=True, color=NAVY)
    CMP = [("Каша рисова hapay! 340 г · бляшанка", 63, GREEN),
           ("Ben's Original 250 г · медіана", 90, PLUM),
           ("Український пауч 350 г · медіана", 101, ORANGE),
           ("Плов М'ясторія 350 г · лоток", 147, ROSE),
           ("Ottogi 247–320 г · медіана", 252, SLATE)]
    yy = 5.44
    mxv = max(v for _, v, _ in CMP)
    for lb, v, c in CMP:
        text(s, M + 6.40, yy - 0.04, 2.70, 0.20, lb, size=8.5, color=INK)
        rect(s, M + 9.20, yy + 0.01, 1.90 * v / mxv, 0.13, c, rounded=True, adj=0.5)
        text(s, M + 9.20 + 1.90 * v / mxv + 0.08, yy - 0.05, 0.80, 0.20, f"{num(v)} грн",
             size=8, bold=True, color=c)
        yy += 0.28
    foot(s, "Джерело: API 17 мереж zakaz.ua, 23.09.2026. Ці шість позицій — суміжна "
            f"категорія, у {N_SKU} позицій дослідження вони не входять: бляшанку не "
            "розігрівають у мікрохвильовці.")


# ══ КАНАЛИ ════════════════════════════════════════════════════════════════
GROUPS = [("МЕРЕЖЕВИЙ РОЗДРІБ", ORANGE, ["Сільпо"],
           "Ben's Original — 10 позицій. Інших брендів категорії немає."),
          ("АЗІЙСЬКІ ФУДШОПИ", PLUM,
           ["Тайякі Март", "Апетітаріум", "Gurmissimo", "Скарби Азії", "Edison Lee"],
           "Ottogi, Bibigo, Clearspring. Вузький канал, висока ціна."),
          ("PROM ТА ФІРМОВІ МАГАЗИНИ", GREEN,
           ["portion.com.ua", "СУХПАЙ", "UPcompany", "Мартел-shop", "Козуб Маркет",
            "Смачна адреса", "Продукт-Shop", "Belorfoods", "Euro-komplekt", "ALANTUR",
            "MK-Sport"],
           "Увесь український реторт-пауч і Adventure Menu.")]
N_SELLERS = sum(len(g[2]) for g in GROUPS)


def slide_channels():
    s = slide()
    header(s, "ДЕ ПРЕДСТАВЛЕНО", f"Три типи каналів, {N_SELLERS} продавців",
           "Мережа є лише в одного бренду. Українське виробництво в мережі не представлене взагалі.")
    yy = 1.90
    for gi, (title, c, chans, what) in enumerate(GROUPS):
        rh = 0.86 if gi < 2 else 1.14
        rect(s, M, yy, 11.98, rh, MIST if gi % 2 == 0 else WHITE, rounded=True, adj=0.12)
        rect(s, M, yy, 0.09, rh, c, rounded=True, adj=0.5)
        text(s, M + 0.30, yy + 0.14, 3.70, 0.22, title, size=9.5, bold=True, color=c)
        text(s, M + 0.30, yy + 0.40, 3.70, 0.56, what, size=8.5, color=GREY, line=1.20)
        cx, cy = M + 4.20, yy + 0.16
        for ch in chans:
            w = 0.24 + 0.068 * len(ch)
            if cx + w > M + 11.10:
                cx, cy = M + 4.20, cy + 0.34
            pill(s, cx, cy, ch, c, w=w, size=8.5)
            cx += w + 0.12
        text(s, M + 11.26, yy + rh / 2 - 0.17, 0.56, 0.34, str(len(chans)), size=16,
             bold=True, color=c, align=PP_ALIGN.RIGHT)
        yy += rh + 0.10
    rect(s, M, 5.14, 5.86, 1.02, MIST, rounded=True, adj=0.10)
    text(s, M + 0.26, 5.28, 5.34, 0.24, "Мережі без жодної позиції категорії",
         size=11, bold=True, color=NAVY)
    text(s, M + 0.26, 5.54, 5.34, 0.56,
         "АТБ, Novus, Metro, Varus, Ашан, Fozzy, «Таврія В», МегаМаркет, ЕКО маркет, "
         "Ultramarket та ще 7 мереж — перевірено через API zakaz.ua 23.09.2026.",
         size=8.5, color=INK, line=1.24)
    rect(s, M + 6.12, 5.14, 5.86, 1.02, MIST, rounded=True, adj=0.10)
    text(s, M + 6.38, 5.28, 5.34, 0.24, "Що відомо про походження", size=11, bold=True, color=NAVY)
    text(s, M + 6.38, 5.54, 5.34, 0.56,
         "Ben's Original виробляє Mars, заводи у Франції та Німеччині. Український пауч — "
         "Верес, Ходорівський м'ясокомбінат, «Маркел», «МАКРО», «М'ясниця», Portion. "
         "Імпортера в "
         "картках продавців не вказано.",
         size=8.5, color=INK, line=1.24)
    insight(s, M, 6.28, 11.98, 0.60, None,
            "Головна дірка каналу: український виробник робить найдешевший за грам продукт "
            "категорії — і не має жодної мережевої полиці.", GREEN)


# ══ ВИСНОВКИ ══════════════════════════════════════════════════════════════
def slide_conclusions():
    s = slide()
    header(s, "ВИСНОВКИ", "Чотири факти і три обмеження",
           "Усе нижче спирається на ціни, перевірені поштучно 23 вересня 2026 року.")
    FACTS = [("Категорія мала, але росте", ORANGE,
              f"{num(T_TOTAL)} т імпорту у 2025-му, +{num(GROWTH)} % за рік і рекорд за "
              "вартістю. На особу це у 20 разів менше за ЄС."),
             ("Гарнір тримає один бренд", PLUM,
              "Ben's Original — 10 позицій у «Сільпо» за 45–179 грн. Другого бренду "
              "чистого готового рису на ринку немає."),
             ("Українці виграють за ціною", GREEN,
              "Реторт-пауч 350 г за 71–199 грн, медіана 101 грн. Це більша порція за ту "
              "саму ціну, що імпортний гарнір 220–250 г."),
             ("Мережі закриті для категорії", ROSE,
              f"З {N_SKU} позицій у мережі продаються 10, і всі — Ben's. Український пауч "
              "живе на Prom і у фірмових магазинах.")]
    for i, (t, c, body) in enumerate(FACTS):
        x = M + (i % 2) * 6.12
        y = 1.88 + (i // 2) * 1.36
        rect(s, x, y, 5.86, 1.26, MIST, rounded=True, adj=0.08)
        rect(s, x, y, 0.09, 1.26, c, rounded=True, adj=0.5)
        text(s, x + 0.30, y + 0.16, 5.30, 0.26, f"{i + 1}. {t}", size=12.5, bold=True, color=NAVY)
        text(s, x + 0.30, y + 0.48, 5.30, 0.70, body, size=9.5, color=INK, line=1.28)
    insight(s, M, 4.72, 11.98, 1.56, "Чого дані не показують",
            "1. Продажів. Ні мережі, ні маркетплейси не віддають обсяги: ємність на слайді 03 "
            "порахована за імпортом, а не заміряна.\n"
            "2. Обсягу українського виробництва. Верес, «МАКРО», «Маркел», «М'ясниця», "
            "Portion "
            "виробляють у країні, тож у митну статистику не потрапляють — реальна категорія "
            "більша за 182 тонни.\n"
            "3. Імпортерів. У картках продавців їх не вказано; перелік дасть лише "
            "вивантаження Держмитслужби за УКТЗЕД 1904 90 10 за 2024–2025 рр.", ORANGE)
    foot(s, f"{N_SKU} позицій, {N_BRANDS} брендів, {N_SELLERS} продавців, 17 мереж. "
            "Зріз 23 вересня 2026 р. Предмет: рис кімнатного зберігання, який лише розігрівають.")


# ══ ЗБІРКА ════════════════════════════════════════════════════════════════
slide_cover()
slide_summary()
slide_volume()
slide_map()
slide_bens()
slide_bowls()
slide_ua()
slide_import()
slide_format()
slide_ladder()
slide_chain_shelf()
slide_channels()
slide_conclusions()

prs.save("Gotovyi_Rys_Rozbir_Brendiv.pptx")
print("saved ·", len(prs.slides._sldIdLst), "slides ·", N_SKU, "SKU ·", N_BRANDS, "brands ·",
      "median/unit", MED_UNIT, "· median/100g", round(MED100, 1))
