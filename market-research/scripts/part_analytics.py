
import math

# ── примітиви ─────────────────────────────────────────────────────────────
def kpi(s, x, y, w, h, label, value, note, c):
    rect(s, x, y, w, h, MIST, rounded=True, adj=0.08)
    rect(s, x, y, w, 0.09, c, rounded=True, adj=0.5)
    text(s, x + 0.24, y + 0.26, w - 0.48, 0.22, label, size=8.5, bold=True, color=GREY)
    text(s, x + 0.24, y + 0.50, w - 0.48, 0.44, value, size=23, bold=True, color=c)
    text(s, x + 0.24, y + 1.00, w - 0.48, h - 1.08, note, size=9, color=INK, line=1.24)


def insight(s, x, y, w, h, title, body, c=ORANGE):
    rect(s, x, y, w, h, NAVY, rounded=True, adj=0.06)
    rect(s, x, y, 0.09, h, c, rounded=True, adj=0.5)
    if title:
        text(s, x + 0.34, y + 0.18, w - 0.60, 0.28, title, size=12.5, bold=True, color=AMBER)
    text(s, x + 0.34, y + (0.54 if title else 0.22), w - 0.60,
         h - (0.70 if title else 0.40), body, size=10,
         color=RGBColor(0xD5, 0xDB, 0xEA), line=1.30)


def grid(s, items, y, w, h, x0=M, gap=0.10, brand=True):
    for i, (im, b, nm, g, lo, hi, ch, c) in enumerate(items):
        sku_cell(s, x0 + i * (w + gap), y, w, h, im, nm, g, lo, hi, ch, c,
                 brand=b if brand else None)


def brands_of(items):
    return list(dict.fromkeys(x[1] for x in items))


def rng_unit(items):
    return min(x[4] for x in items), max(x[5] for x in items)


def p100(items):
    v = [(lo / g * 100, hi / g * 100) for *_, g, lo, hi, _, _ in items if g]
    return min(a for a, _ in v), max(b for _, b in v)


def med_unit(items):
    return st.median([x[4] for x in items])


def med100(items):
    return st.median([x[4] / x[3] * 100 for x in items])


SEGMENTS = [("ГАРНІР У ПАУЧІ", L_SIDE, ORANGE,
             "Чистий рис без наповнювача. Мікрохвильовка 90 с."),
            ("СТРАВА В ЧАШІ", L_BOWL, PLUM,
             "Рис із м'ясом у жорсткій чаші. Мікрохвильовка 2 хв."),
            ("СТРАВА В ПАУЧІ · УКРАЇНА", L_UA, GREEN,
             "Рис із м'ясом у реторт-паучі. Розігрів у воді або НВЧ."),
            ("СТРАВА В ПАУЧІ · ІМПОРТ", L_IMP, TEAL,
             "Те саме, але завезене: ЄС, органіка й туристична лінійка.")]


# ══ 01 · ОБКЛАДИНКА ═══════════════════════════════════════════════════════
def slide_cover():
    s = slide(dark=True)
    cover_pic(s, BG_TITLE, 0, 0, W, H)
    s.shapes.add_picture(LOGO, Inches(M), Inches(0.52), Inches(1.94), Inches(0.84))
    rect(s, M, 2.02, 0.54, 0.09, ORANGE)
    text(s, M, 2.34, 7.00, 0.34, "ДОСЛІДЖЕННЯ РИНКУ · ВЕРЕСЕНЬ 2026", size=11.5,
         bold=True, color=AMBER)
    text(s, M, 2.78, 7.20, 1.70, "Готовий рис\nв Україні", size=46, bold=True,
         color=WHITE, line=1.04)
    text(s, M, 4.52, 7.00, 0.36, "Кімнатне зберігання · розігрів · без готування",
         size=14.5, color=RGBColor(0xC6, 0xCE, 0xE2))
    rect(s, M, 5.12, 5.90, 0.055, RGBColor(0x55, 0x60, 0x8C))
    text(s, M, 5.34, 6.90, 0.86,
         f"{pl(N_BRANDS, 'бренд', 'бренди', 'брендів')} · "
         f"{pl(N_SKU, 'позиція', 'позиції', 'позицій')} · ціна за одну упаковку\n"
         "Кожну ціну перевірено на сайті продавця 23 вересня 2026 р.",
         size=11.5, color=RGBColor(0x9F, 0xA9, 0xC4), line=1.40)
    for im, x, y, w, h in [("md/bens_basmati250.jpg", 8.26, 0.96, 2.35, 2.40),
                           ("sku/ot_bibimbap.jpg", 10.78, 0.96, 2.35, 2.40),
                           ("sku/ua_veres_chicken.jpg", 8.26, 3.50, 2.35, 2.10),
                           ("sku/ua_makro_beef.jpg", 10.78, 3.50, 2.35, 2.10)]:
        rect(s, x, y, w, h, WHITE, rounded=True, adj=0.06)
        pic(s, im, x + 0.16, y + 0.14, w - 0.32, h - 0.28)
    rect(s, 8.26, 5.76, 4.87, 1.02, RGBColor(0x26, 0x2F, 0x4D), rounded=True, adj=0.14)
    text(s, 8.50, 5.92, 4.44, 0.78, " · ".join(dict.fromkeys(x[1] for x in ALL8)),
         size=9.5, bold=True, color=WHITE, line=1.32)


# ══ 02 · ГОЛОВНЕ ══════════════════════════════════════════════════════════
def slide_summary():
    s = slide()
    header(s, "ГОЛОВНЕ", "Шість цифр про категорію",
           "Предмет: рис, що зберігається за кімнатної температури і лише розігрівається.")
    K = [("ОБСЯГ ІМПОРТУ 2025", f"{num(T_TOTAL)} т", ORANGE,
          f"Готовий рис, CN 1904 90 10.\n+{num(GROWTH)} % до 2024 року."),
         ("ЄМНІСТЬ РОЗДРІБУ", f"{num(RETAIL_LO)}–{num(RETAIL_HI)}", TEAL,
          "млн грн на рік — оцінка\nза імпортом × 2,2–2,8."),
         ("ВІДСТАВАННЯ ВІД ЄС", f"×{num(GAP)}", ROSE,
          f"{num(UA_PER_CAP_USD, 2)} $ на особу проти\n1,43 $ у ЄС."),
         ("ПОЗИЦІЙ У ПРОДАЖУ", f"{N_SKU}", PLUM,
          f"{N_BRANDS} брендів. {N_UA} позицій —\nукраїнського виробництва."),
         ("МЕДІАНА ЗА УПАКОВКУ", f"{num(MED_UNIT)} грн", GREEN,
          f"Діапазон {num(min(x[4] for x in ALL8))}–{num(max(x[5] for x in ALL8))} грн.\n"
          f"Половина позицій — до {num(MED_UNIT)} грн."),
         ("МАСОВА ВАГА", f"{num(G_TOP[0][0])} г", SLATE,
          f"{pl(G_TOP[0][1], 'позиція', 'позиції', 'позицій')}. "
          f"Далі {num(G_TOP[1][0])} г — {G_TOP[1][1]},\n"
          f"{num(G_TOP[2][0])} г — {G_TOP[2][1]}.")]
    for i, (lb, v, c, note) in enumerate(K):
        kpi(s, M + i * 2.04, 1.88, 1.86, 1.92, lb, v, note, c)
    insight(s, M, 4.02, 5.86, 2.36, "Що показало дослідження",
            "1. Українське виробництво в категорії вже є — і воно найдешевше: реторт-пауч "
            "350 г за 71–199 грн, медіана 101 грн за упаковку.\n"
            "2. Ben's Original — єдиний імпортний гарнір у мережі: 10 позицій у «Сільпо» "
            "за 45–179 грн.\n"
            "3. Корейська чаша — найдорожчий сегмент (медіана 252 грн) і продається "
            "лише у трьох азійських магазинах.", ORANGE)
    insight(s, M + 6.12, 4.02, 5.86, 2.36, "Де порожньо",
            f"1. Немає українського гарніру — чистого рису в паучі. Усі {N_UA} українських "
            "позицій це рис із м'ясом, тобто повноцінна страва.\n"
            "2. Немає чаші дешевше 189 грн: Bibigo 189, Ottogi 225–356.\n"
            "3. Немає жодної позиції категорії в АТБ, Novus, Metro, Varus, Ашан, Fozzy "
            "та ще 11 мережах — перевірено через API 17 мереж.", TEAL)
    foot(s, "Ціни: сайти продавців, 23.09.2026, за одну упаковку в роздріб. "
            "Імпорт: Eurostat Comext CN 1904 90 10 + UN Comtrade. "
            "Ємність роздрібу — оцінка, метод на слайді 03.")


# ══ 03 · ОБСЯГ РИНКУ ══════════════════════════════════════════════════════
def slide_volume():
    s = slide()
    header(s, "ОБСЯГ РИНКУ",
           f"{num(T_TOTAL)} тонн імпорту, {num(RETAIL_LO)}–{num(RETAIL_HI)} млн грн роздрібу",
           "Імпорт готового рису відновився: 2025 рік — найбільший за сім років за вартістю.")
    rect(s, M, 1.88, 7.30, 3.34, MIST, rounded=True, adj=0.05)
    text(s, M + 0.28, 2.06, 6.00, 0.24, "ІМПОРТ ГОТОВОГО РИСУ З ЄС, ТОНН НА РІК",
         size=9, bold=True, color=GREY)
    bx, by0, bh = M + 0.34, 4.56, 2.02
    mx = max(t for _, t, _ in IMPORT)
    for i, (yr, t, v) in enumerate(IMPORT):
        x = bx + i * 1.00
        h = bh * t / mx
        c = ORANGE if yr == 2025 else (NAVY_L if yr >= 2022 else RGBColor(0xC2, 0xC9, 0xDA))
        rect(s, x, by0 - h, 0.86, h, c, rounded=True, adj=0.10)
        text(s, x - 0.06, by0 - h - 0.24, 0.98, 0.22, num(t), size=8.5, bold=True,
             color=c if yr == 2025 else GREY, align=PP_ALIGN.CENTER)
        text(s, x - 0.06, by0 + 0.06, 0.98, 0.22, str(yr), size=8.5,
             bold=(yr == 2025), color=INK, align=PP_ALIGN.CENTER)
        text(s, x - 0.06, by0 + 0.26, 0.98, 0.20, f"€{num(v)}k", size=7.5, color=GREY,
             align=PP_ALIGN.CENTER)
    text(s, M + 0.34, 4.98, 7.00, 0.20,
         "Вартість 2025 року — 823 тис. € проти 475 тис. € у 2024-му, +73 %.",
         size=8.5, color=GREY)
    rect(s, M + 7.54, 1.88, 4.44, 3.34, WHITE, line=MIST_D, lw=1.2, rounded=True, adj=0.05)
    text(s, M + 7.80, 2.06, 3.92, 0.24, "ЗВІДКИ ЇДЕ, 2025 РІК", size=9, bold=True, color=GREY)
    yy = 2.36
    tot = sum(t for _, t, _ in IMPORT_PARTNERS_2025)
    for nm, t, v in IMPORT_PARTNERS_2025:
        w = 2.90 * t / tot
        text(s, M + 7.80, yy, 1.00, 0.22, nm, size=9, color=INK)
        rect(s, M + 8.80, yy + 0.05, max(w, 0.04), 0.13, ORANGE if t > 50 else TEAL,
             rounded=True, adj=0.5)
        text(s, M + 8.80 + max(w, 0.04) + 0.08, yy - 0.01, 0.92, 0.22, f"{num(t, 1)} т",
             size=8.5, bold=True, color=GREY)
        yy += 0.34
    text(s, M + 7.80, yy + 0.06, 3.92, 0.76,
         "Польща й Болгарія — 95 % тонажу: це заводи ЄС, що пакують рис у пауч.\n"
         "Корея і Китай разом 3 т — чаші Ottogi та Bibigo. Україна свій обсяг "
         "виробляє на місці, в імпорт він не входить.",
         size=8.5, color=GREY, line=1.26)
    rect(s, M, 5.26, 11.98, 1.44, NAVY, rounded=True, adj=0.07)
    rect(s, M, 5.26, 0.09, 1.44, TEAL, rounded=True, adj=0.5)
    text(s, M + 0.34, 5.42, 3.90, 0.26, "Як рахували ємність", size=12.5, bold=True, color=AMBER)
    STEPS = [(f"{num(T_TOTAL)} т", "імпорт 2025"),
             (f"{num(CIF_UAH)} млн грн", "CIF за курсом НБУ"),
             ("× 2,2–2,8", "мито 0 %, ПДВ, маржа"),
             (f"{num(RETAIL_LO)}–{num(RETAIL_HI)} млн", "роздріб на рік"),
             (f"≈ {num(PACKS_K)} тис.", "упаковок на рік")]
    for i, (v, lb) in enumerate(STEPS):
        x = M + 4.30 + i * 1.54
        text(s, x, 5.42, 1.48, 0.28, v, size=12, bold=True, color=WHITE)
        text(s, x, 5.70, 1.48, 0.36, lb, size=8, color=RGBColor(0x9F, 0xA9, 0xC4), line=1.20)
        if i < len(STEPS) - 1:
            text(s, x + 1.32, 5.44, 0.20, 0.24, "→", size=12, bold=True, color=ORANGE)
    text(s, M + 0.34, 5.78, 3.76, 0.80,
         "Це лише імпортна частина. Український реторт-пауч виробляють у країні, "
         "і в митну статистику він не потрапляє.",
         size=8.5, color=RGBColor(0xB9, 0xC2, 0xDA), line=1.26)
    foot(s, "Джерела: Eurostat Comext DS-045409, CN 1904 90 10 «рис приготовлений», "
            "експорт ЄС→Україна 2019–2025; UN Comtrade (Корея, Китай) за 2025 р.; "
            "курс НБУ, середній за 2025 р.: 47,15 грн/€. Ємність роздрібу — ОЦІНКА.")


# ══ 04 · ЧОТИРИ СЕГМЕНТИ ══════════════════════════════════════════════════
def slide_map():
    s = slide()
    header(s, "КАРТА КАТЕГОРІЇ",
           f"Чотири сегменти, {pl(N_BRANDS, 'бренд', 'бренди', 'брендів')}",
           "Усередині категорії — два формати паковання і дві ролі: гарнір або повна страва.")
    for i, (kind, items, c, note) in enumerate(SEGMENTS):
        x = M + i * 3.07
        lo, hi = rng_unit(items)
        br = brands_of(items)
        rect(s, x, 1.90, 2.82, 3.06, MIST, rounded=True, adj=0.06)
        rect(s, x, 1.90, 2.82, 0.10, c, rounded=True, adj=0.5)
        text(s, x + 0.22, 2.10, 2.40, 0.24, kind, size=9.5, bold=True, color=c)
        text(s, x + 0.22, 2.38, 2.40, 0.32,
             f"{pl(len(items), 'позиція', 'позиції', 'позицій')}", size=14, bold=True, color=NAVY)
        text(s, x + 0.22, 2.72, 2.40, 0.62, " · ".join(br), size=9, color=INK, line=1.22)
        text(s, x + 0.22, 3.38, 2.40, 0.44, note, size=8.5, color=GREY, line=1.20)
        rect(s, x + 0.22, 3.88, 2.38, 0.02, MIST_D)
        text(s, x + 0.22, 3.98, 2.40, 0.20, "ЗА УПАКОВКУ · МЕДІАНА", size=7,
             bold=True, color=GREY)
        text(s, x + 0.22, 4.16, 2.40, 0.28, f"{num(med_unit(items))} грн", size=15,
             bold=True, color=c)
        text(s, x + 0.22, 4.44, 2.40, 0.20, f"{num(lo)} – {num(hi)} грн", size=8.5, color=GREY)
        text(s, x + 0.22, 4.64, 2.40, 0.20,
             f"маса {rng(min(y[3] for y in items), max(y[3] for y in items))} г",
             size=8.5, bold=True, color=INK)
    insight(s, M, 5.14, 11.98, 1.42, "Як читати цю карту",
            "Український реторт-пауч уже виграє за ціною: медіана 101 грн проти 90 грн у "
            "Ben's Original — але в паучі 350 г проти 220–250 г, тобто це повна страва "
            "з м'ясом, а не гарнір. Інша полиця й інший привід.\n"
            "Ben's Original тримає нішу гарніру сам-один: жодного українського чистого рису "
            "в паучі на ринку немає. Корейська чаша — найдорожчий формат і найвужчий канал: "
            "три азійські магазини.", ORANGE)
    foot(s, f"{N_SKU} позицій, ціни за одну упаковку, 23.09.2026.")
