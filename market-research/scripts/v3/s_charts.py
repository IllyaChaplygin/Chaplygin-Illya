# ══ ЧАРТИ: формат × вага, топ поєднань, сходи цін, «де ми стоїмо», фінмодель ══
def tint(c, a):
    return RGBColor(*[int(255 - (255 - v) * a) for v in (c[0], c[1], c[2])])


OUR = [  # наші позиції: (формат, грамаж, підпис)
    ('pouch', 150, 'BSCM 150'), ('pouch', 200, 'BSCM 200'), ('pouch', 240, 'BSCM 240'),
    ('pouch', 250, 'CM Premium'), ('cup', 150, 'BSCM стакан'), ('cup', 200, "ONE'S 200/210"),
    ('cup', 250, 'BSCM 2×125')]
HB = [(80, 100, '80–99 г'), (100, 150, '100–149 г'), (150, 200, '150–199 г'), (200, 250, '200–249 г'),
      (250, 300, '250–299 г'), (300, 350, '300–349 г'), (350, 999, '350+ г')]


def s_heat():
    s = new_slide()
    header(s, "ФОРМАТ × ВАГА", "Формат × вага",
           "Скільки позицій у кожній вазі. Помаранчеві мітки — наші формати.")
    gx0, gw = 2.45, (W - M - 2.45) / 7
    cell_h, gy0 = 0.86, 2.30
    for j, (lo, hi, lab) in enumerate(HB):
        text(s, gx0 + j * gw, 1.88, gw, 0.26, lab, size=12, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    # порційне вікно
    rect(s, gx0 + 3 * gw - 0.03, 2.22, 2 * gw + 0.06, 4 * cell_h + 0.14, None, line=ORANGE, lw=2.5, rounded=True, adj=0.03)
    cnt_win = 0
    for i, f in enumerate(FMT_ORDER):
        y = gy0 + i * cell_h
        text(s, M, y + 0.14, 1.7, 0.4, FMT_NAME[f], size=15, bold=True, color=FMT_COL[f], align=PP_ALIGN.RIGHT)
        text(s, M, y + 0.46, 1.7, 0.25, {'pouch': 'ретрорт', 'cup': 'стакан', 'doypack': 'плоский пакет',
                                         'box': 'з нагрівачем'}[f].replace('ретрорт', 'реторт'),
             size=9, color=GREY, align=PP_ALIGN.RIGHT)
        for j, (lo, hi, lab) in enumerate(HB):
            n = sum(1 for x in SKU if x['fmt'] == f and x['g'] and lo <= x['g'] < hi)
            if 200 <= lo < 300: cnt_win += n
            ours = [t for (ff, g_, t) in OUR if ff == f and lo <= g_ < hi]
            x = gx0 + j * gw
            if n:
                a = 0.18 + 0.82 * min(n, 16) / 16
                col = tint(FMT_COL[f], a)
                rect(s, x + 0.04, y + 0.03, gw - 0.08, cell_h - 0.08, col, rounded=True, adj=0.08)
                dark = a > 0.55
                text(s, x, y + 0.04, gw, 0.4, str(n), size=21, bold=True, color=WHITE if dark else NAVY, align=PP_ALIGN.CENTER)
                text(s, x, y + 0.36, gw, 0.2, "позицій", size=8.5, color=WHITE if dark else GREY, align=PP_ALIGN.CENTER)
            else:
                rect(s, x + 0.04, y + 0.03, gw - 0.08, cell_h - 0.08, MIST, rounded=True, adj=0.08)
                text(s, x, y + 0.04, gw, 0.4, "0" if ours else "—", size=21, bold=True,
                     color=ORANGE if ours else LGREY, align=PP_ALIGN.CENTER)
            if ours:
                lab_t = " · ".join(ours)
                pw = min(gw - 0.14, 0.2 + 0.075 * len(lab_t))
                rect(s, x + (gw - pw) / 2, y + cell_h - 0.33, pw, 0.23, ORANGE, rounded=True, adj=0.5)
                text(s, x + (gw - pw) / 2, y + cell_h - 0.33, pw, 0.23, lab_t, size=8.3, bold=True, color=NAVY,
                     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, gx0 + 3 * gw, 2.04 - 0.02, 2 * gw, 0.01, "", size=6)
    iy = gy0 + 4 * cell_h + 0.22
    cw = (W - 2 * M - 0.32) / 3
    cards = [("ПОРЦІЙНЕ ВІКНО 200–299 г", f"{cnt_win} позицій",
              "Пауч і чаша: одна порція на людину. Тут стоять BSCM 200/240, CM Premium, ONE'S і подвійний стакан.", ORANGE),
             ("НАЙЩІЛЬНІША КЛІТИНКА", "дойпак 100–149 г — 16",
              "Це сухий сублімат для походів, інший привід і канал. Наші формати з ним не перетинаються.", PLUM),
             ("ВІЛЬНІ КЛІТИНКИ ДЛЯ НАС", "пауч 150–199 г — 0",
              "Жодного пауча 150–199 г. BSCM пауч 150 г заходить без прямого конкурента; найближчий — Ben's 220 г.", GREEN)]
    for k, (lb, big, txt, c) in enumerate(cards):
        x = M + k * (cw + 0.16)
        rect(s, x, iy, cw, 6.94 - iy, MIST, rounded=True, adj=0.06); rect(s, x, iy, 0.08, 6.94 - iy, c, rounded=True, adj=0.5)
        text(s, x + 0.24, iy + 0.1, cw - 0.4, 0.2, lb, size=8.5, bold=True, color=GREY)
        text(s, x + 0.24, iy + 0.32, cw - 0.4, 0.3, big, size=15, bold=True, color=NAVY)
        text(s, x + 0.24, iy + 0.66, cw - 0.4, 0.55, txt, size=9.3, color=INK, line=1.15)
    foot(s, "Вага — готової страви; для сухого рису й сублімату — вага продукту. Позиції без вказаної маси (Trek'n Eat ×2) не враховано. "
            "Наші формати — за комерційними пропозиціями BSCM, CM Premium, One's.")


TOP13 = [('pouch', 220, 129, "Ben's Original"), ('pouch', 250, 90, "Ben's Original · Clearspring"),
         ('cup', 174, 148, "Henan"), ('doypack', 85, 120, "Харчі · Їжа в Похід"),
         ('doypack', 110, 378, "Mountain House · Adventure Menu · SubliMate"), ('doypack', 125, 515, "Travellunch"),
         ('cup', 144, 135, "Henan"), ('doypack', 90, 133, "James Cook"),
         ('doypack', 146, 501, "Qiaoshanmei · Adventure Food"), ('doypack', 250, 811, "Travellunch"),
         ('box', 187, 665, "Haidilao"), ('box', 272, 799, "Haidilao"), ('pouch', 400, 356, "Adventure Menu")]


def s_top13():
    s = new_slide()
    header(s, "НАЙПОПУЛЯРНІШІ ФОРМАТИ", "Топ поєднань формат + вага",
           "Які упаковки ринок обирає найчастіше — і де стоять наші ваги.")
    rect(s, M, 1.80, 8.45, 5.14, MIST, rounded=True, adj=0.03)
    text(s, M + 0.26, 1.90, 8.0, 0.26, "ТОП-13 ПОЄДНАНЬ «ФОРМАТ + ГРАМАЖ», ПОЗИЦІЙ", size=10.5, bold=True, color=GREY)
    y = 2.24; rh = 0.355
    for f, g_, med, br in TOP13:
        n = sum(1 for x in SKU if x['fmt'] == f and x['g'] == g_)
        text(s, M + 0.22, y + 0.05, 1.75, 0.3, f"{FMT_NAME[f]} {g_} г", size=12.5, bold=True, color=NAVY)
        rect(s, M + 2.0, y + 0.06, 0.30 * n, 0.25, FMT_COL[f], rounded=True, adj=0.25)
        text(s, M + 2.0 + 0.30 * n + 0.08, y + 0.03, 0.4, 0.3, str(n), size=14, bold=True, color=NAVY)
        text(s, M + 4.35, y + 0.06, 1.6, 0.26, f"медіана {med} грн", size=10.5, color=INK)
        text(s, M + 5.95, y + 0.07, 2.4, 0.26, br, size=8.8, color=GREY)
        y += rh
    rx = M + 8.67; rw = W - M - rx
    rect(s, rx, 1.80, rw, 2.34, WHITE, line=MIST_D, lw=1.2, rounded=True, adj=0.05)
    text(s, rx + 0.2, 1.90, rw - 0.3, 0.24, "ЧАСТКА ФОРМАТУ", size=10, bold=True, color=GREY)
    for k, (lab, parts) in enumerate([("Усі 84 позиції", [('doypack', 30), ('cup', 23), ('pouch', 19), ('box', 12)]),
                                       ("Лише масовий канал · 46", [('cup', 23), ('pouch', 19), ('box', 2), ('doypack', 2)])]):
        yy = 2.20 + k * 0.92
        text(s, rx + 0.2, yy, rw - 0.3, 0.22, lab, size=9.5, bold=True, color=INK)
        tot = sum(n for _, n in parts); xx = rx + 0.2; full = rw - 0.4
        for f, n in parts:
            ww = full * n / tot
            rect(s, xx, yy + 0.26, ww, 0.36, FMT_COL[f])
            if ww > 0.38:
                text(s, xx, yy + 0.26, ww, 0.36, f"{num(n / tot * 100)} %", size=11 if ww > 0.55 else 9, bold=True, color=WHITE,
                     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            xx += ww
    ly = 3.58
    for f in FMT_ORDER:
        pass
    # наші формати
    oy = 4.28
    rect(s, rx, oy, rw, 2.66, NAVY, rounded=True, adj=0.05); rect(s, rx, oy, 0.09, 2.66, ORANGE, rounded=True, adj=0.5)
    text(s, rx + 0.28, oy + 0.12, rw - 0.4, 0.28, "НАШІ ВАГИ НА ЦІЙ КАРТІ", size=10.5, bold=True, color=AMBER)
    rows = [("Пауч 150 г · BSCM", 'pouch', 150), ("Пауч 200 г · BSCM", 'pouch', 200), ("Пауч 240 г · BSCM", 'pouch', 240),
            ("Пауч 250 г · CM Premium", 'pouch', 250), ("Стакан 150 г · BSCM", 'cup', 150), ("Лоток 210 г · ONE'S", 'cup', 210)]
    yy = oy + 0.5
    for lb, f, g_ in rows:
        n = sum(1 for x in SKU if x['fmt'] == f and x['g'] == g_)
        text(s, rx + 0.28, yy, rw - 1.5, 0.3, lb, size=10.5, color=WHITE)
        text(s, rx + rw - 1.25, yy, 1.05, 0.3, "0 конкурентів" if n == 0 else f"{n} в цій вазі", size=10.5, bold=True,
             color=AMBER if n == 0 else RGBColor(0xD5, 0xDB, 0xEA), align=PP_ALIGN.RIGHT)
        yy += 0.34
    foot(s, "Позиції — за картками продавців (84 SKU). Медіани — за даними колоди. Масовий канал — без туристичних і військових магазинів. "
            "«Конкурентів» — позицій того самого формату з точно такою вагою.")


# ── сходи цін ─────────────────────────────────────────────────────────────
def _ladder(title, eyebrow, dek, groups, axmax, ticks, note, window=True):
    s = new_slide()
    header(s, eyebrow, title, dek)
    nrows = sum(len(g[2]) for g in groups); nh = len(groups)
    y0, y1 = 1.78, 6.52
    rh = (y1 - y0 - nh * 0.34) / nrows
    AX, AW = 5.15, 7.05
    ax = lambda v: AX + AW * v / axmax
    # сітка й вікно полиці
    for t in ticks:
        rect(s, ax(t), y0 + 0.02, 0.012, y1 - y0 - 0.02, MIST_D)
        text(s, ax(t) - 0.4, y1 + 0.02, 0.8, 0.2, num(t), size=10, color=GREY, align=PP_ALIGN.CENTER)
    if window:
        rect(s, ax(45), y0 + 0.02, ax(180) - ax(45), y1 - y0 - 0.02, RGBColor(0xFF, 0xF6, 0xE3))
        text(s, ax(45), y1 + 0.26, 4.2, 0.22, "вікно мережевої полиці 45–180 грн", size=10, bold=True, color=ORANGE)
    text(s, M, y1 + 0.02, 3.4, 0.2, "ГРН ЗА ОДНУ УПАКОВКУ · ТОЧКА — МЕДІАНА БРЕНДУ", size=8.5, bold=True, color=GREY)
    y = y0
    for gname, gcol, rows in groups:
        rect(s, M, y + 0.03, W - 2 * M, 0.28, gcol, rounded=True, adj=0.3)
        text(s, M + 0.14, y + 0.03, 9.5, 0.28, gname, size=10.5, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
        y += 0.34
        for nm, lo, hi, med, n, cty, ours in rows:
            cy = y + rh / 2
            if ours:
                rect(s, M, y + 0.02, W - 2 * M, rh - 0.04, RGBColor(0xFF, 0xF1, 0xD2), rounded=True, adj=0.2)
            text(s, M + 0.14, cy - 0.13, 2.9, 0.28, nm, size=12 if ours else 11.5, bold=True, color=NAVY)
            text(s, M + 3.05, cy - 0.11, 0.9, 0.24, cty, size=9.5, color=GREY)
            text(s, M + 3.95, cy - 0.11, 0.55, 0.24, str(n), size=10, bold=True, color=GREY, align=PP_ALIGN.RIGHT)
            col = ORANGE if ours else gcol
            x0, x1 = ax(lo), ax(min(hi, axmax))
            rect(s, x0, cy - 0.06, max(x1 - x0, 0.04), 0.12, col, rounded=True, adj=0.5)
            dot(s, ax(med), cy, 0.24, col, line=NAVY if ours else WHITE, lw=2.0 if ours else 1.5)
            lbl = num(lo) if lo == hi else f"{num(lo)}–{num(hi)}"
            text(s, x1 + 0.12, cy - 0.12, 1.5, 0.26, lbl, size=10.5, bold=True, color=NAVY)
            y += rh
    foot(s, note)


def _ours_rows():
    import json
    P = json.load(open(V3 + 'pricing_rows.json'))
    def grp(prefix, excl=None):
        return [p for p in P if p['group'].startswith(prefix) and 'мікс' not in p['group']]
    bscm = [p for p in P if p['group'].startswith('BSCM') and 'мікс' not in p['group']]
    cm = [p for p in P if p['group'].startswith('CM Premium')]
    on = [p for p in P if p['group'].startswith("ONE'S")]
    def row(nm, L):
        v = [p['shelf'] for p in L]
        return (nm, round(min(v)), round(max(v)), round(st.median(v)), len(L), 'Таїланд' if 'BSCM' in nm or 'CM' in nm else 'Корея', True)
    return [row('BSCM Foods', bscm), row('CM Premium · Chefrey', cm), row("ONE'S International", on)]


def s_ladder1():
    R1 = [("Ben's Original", 45, 179, 90, 13, 'ЄС', False), ("Маркел", 94, 96, 95, 1, 'Україна', False),
          ("Portion", 109, 109, 109, 1, 'Україна', False), ("Bibigo", 135, 189, 162, 1, 'Корея', False),
          ("Ottogi", 225, 356, 255, 10, 'Корея', False), ("Clearspring", 252, 446, 349, 1, 'ЄС', False),
          ("Adventure Menu · пауч 400 г", 315, 404, 336, 3, 'ЄС', False)]
    R2 = [("Henan", 83, 156, 140, 8, 'Китай', False), ("Gallina Blanca", 229, 229, 229, 2, 'ЄС', False),
          ("Qiaoshanmei", 501, 501, 501, 2, 'Китай', False)]
    groups = [("СЕГМЕНТ 1 · ГОТОВИЙ РИС ДО РОЗІГРІВУ", GOLD, R1), ("СЕГМЕНТ 2 · СУХИЙ РИС ПІД ОКРІП", TEAL, R2),
              ("НАШІ БРЕНДИ · ПОЛИЦЯ ЗА ФІНМОДЕЛЛЮ (З ПДВ)", NAVY, _ours_rows())]
    _ladder("Ціна: сегменти 1–2 і наші бренди", "ЦІНА ЗА ОДНУ УПАКОВКУ 1 З 2",
            "Готовий рис, рис під окріп і наші бренди на одній шкалі.", groups, 520, [0, 100, 200, 300, 400, 500],
            "Роздрібна ціна за одну упаковку на сайтах продавців (конкуренти); наші бренди — полиця = собівартість × 3,5 "
            "(20' навалом, курс 45, бонус мережі 25 %, маржа 35 %). Смуга — від найдешевшої до найдорожчої позиції.")


def s_ladder2():
    R3 = [("Haidilao", 380, 1190, 680, 10, 'Китай', False), ("Mo Xiao Xian", 931, 931, 931, 2, 'Китай', False),
          ("Zihaiguo", 735, 735, 735, 1, 'Китай', False), ("Rongcheng Haoji", 735, 735, 735, 1, 'Китай', False)]
    R4 = [("James Cook", 55, 156, 133, 4, 'Україна', False), ("!FEST", 133, 140, 137, 1, 'Україна', False),
          ("Їжа в Похід", 130, 145, 138, 2, 'Україна', False), ("Харчі", 83, 321, 120, 4, 'Україна', False),
          ("SubliMate", 290, 370, 309, 3, 'Україна', False), ("Adventure Menu · дойпак 110 г", 378, 483, 433, 2, 'ЄС', False),
          ("Trek'n Eat", 269, 715, 492, 2, 'ЄС', False), ("Adventure Food", 492, 492, 492, 1, 'ЄС', False),
          ("Travellunch", 275, 889, 515, 7, 'ЄС', False), ("Mountain House", 699, 699, 699, 2, 'США', False)]
    groups = [("СЕГМЕНТ 3 · САМОРОЗІГРІВ", ROSE, R3), ("СЕГМЕНТ 4 · СУБЛІМАЦІЯ", PLUM, R4)]
    _ladder("Ціна: сегменти 3–4", "ЦІНА ЗА ОДНУ УПАКОВКУ 2 З 2",
            "Саморозігрів і сублімація — інший привід і канал, з нами не конкурують.", groups, 1250, [0, 250, 500, 750, 1000, 1250],
            "Роздрібна ціна за одну упаковку на сайтах продавців. Інший привід споживання і інший канал: туризм, армія, подарунок.")
