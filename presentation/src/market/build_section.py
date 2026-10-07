# -*- coding: utf-8 -*-
"""Appends the market section (Morskyi Dim style) to the user's supplier deck.
Slides 1..N of the source file are never touched — only new slides are added."""
import hashlib, os, sys, zipfile
from lxml import etree
from pptx import Presentation

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import mdata as D  # noqa: E402
from md import Deck, text, rect, caps, kpi, section_tag, footnote, NAVY, AMBER, GOLD, RED, GREEN, TEAL, PURPLE, PANEL, INK, MUTED, WHITE, SEG_COLOR, SUP_COLOR, tint  # noqa: E402
import slides_market as M  # noqa: E402
import slides_our as O  # noqa: E402
import slides_ext as X  # noqa: E402
import fix_supplier_prices as FX  # noqa: E402
from mdata import SEGS, rows_of, lines, med, gfmt, uah, rng, OUR  # noqa: E402

DEFAULT_SRC = '/root/.claude/uploads/c53f34b4-ae43-52db-aeb2-01bdcda48cc3/8a56c5c7-Snacks_Presentation_Final..pptx'
DEFAULT_OUT = os.path.join(HERE, '..', '..', 'Snacks_Presentation_Final_з_ринком.pptx')


def fm(v, d=1):
    return (('%.' + str(d) + 'f') % v).replace('.', ',')


def catalog_slides(deck, sgs, colors, eyebrow, title_base, per=12):
    items = []
    for sg in sgs:
        for l in lines(rows_of(sg)):
            items.append((l, SEG_COLOR[sg]))
    n = (len(items) + per - 1) // per
    size = (len(items) + n - 1) // n               # balanced pages: 13 -> 7 + 6, never 12 + 1
    for i in range(n):
        chunk = items[i * size:(i + 1) * size]
        lo = min(l['pmin'] for l, _ in chunk); hi = max(l['pmax'] for l, _ in chunk)
        t = '%s: каталог %d з %d — %s ₴' % (title_base, i + 1, n, rng(lo, hi)) if n > 1 else '%s: каталог — %s ₴' % (title_base, rng(lo, hi))
        M.catalog(deck, chunk, eyebrow, t)


def conclusions(deck):
    s = deck.slide('Висновки', 'Що показав зріз ринку і де ми стоїмо')
    O.sub_line(s, 'Ціни нашої полиці — за формулою рису: бонус мережі 25 %, наша маржа 30 % ціни партнеру, націнка магазину +40 %; ринок — онлайн-каталоги мереж на 07.10.2026 і митна база 2025–2026.')
    S = {sg: (rows_of(sg), lines(rows_of(sg))) for sg in SEGS}
    def ps(key_part, g=None):
        o = [x for x in OUR if key_part in x['key'] and (g is None or x['grams'] == g)][0]
        p = D.peer_stats(o)
        return o, p, 0
    roll, rp, rd = ps('Wow ROLL - Original')
    dbl, dp, dd = ps('DOUBLE ROLL - Original')
    sw, swp, swd = ps('WOW SEAWEED - Original')
    mini, mp_, md_ = ps('WOW MINI - Original')
    t30, t30p, t30d = ps('TEMPURA SEAWEED - Corn 30g')
    t50, t50p, t50d = ps('TEMPURA SEAWEED - Corn 50g')
    top35, top35p, top35d = ps('TOPPING - Vegetables 35g')
    top70, top70p, top70d = ps('TOPPING - Vegetables 70g')
    sand, sandp, sandd = ps('SANDWICH SEAWEED - Sesame')
    sk, skp, skd = ps('Original Flavour')
    tn50, tn50p, tn50d = ps('Norimaki Original')
    tn55, tn55p, tn55d = ps('Norimaki Wasabi')
    mrs = S['mini'][0]
    tk25 = next(l for l in lines(rows_of('tempura')) if l['brand'] == 'Tao Kae Noi' and l['grams'] == 25.0)
    tk40 = next(l for l in lines(rows_of('tempura')) if l['brand'] == 'Tao Kae Noi' and l['grams'] == 40.0)
    wwl = next(l for l in lines(rows_of('ricecr')) if l['brand'] == 'Want Want')
    ww60 = next(l for l in lines(rows_of('ricecr')) if l['brand'] == 'Want Want')
    top_mini = dict(M.pack_table(['mini']))['mini'][0]
    sd = [x for x in OUR if 'SANDWICH' in x['key']][0]
    cols = [
        ('ЩО КАЖЕ РИНОК', NAVY, [
            'Полиця мереж — це міні-пакет: %d SKU, %d брендів. Найходовіша упаковка — %s (%d %% присутності), типова ціна %d ₴.' % (len(mrs), len({r['brand'] for r in mrs}), gfmt(top_mini['grams']), round(100 * top_mini['share']), round(top_mini['pmed'])),
            'Чипси 8–40 г — лише 4 бренди (Norris, Hokkaido Club, Ock Dong Ja, Ocean Snack); найходовіша — 25 г, 47–88 ₴.',
            'Темпура — практично один бренд, Tao Kae Noi: 25 г — %d ₴, 40 г — %d ₴.' % (round(tk25['pmed']), round(tk40['pmed'])),
            'Рисові крекери в мережах — лише Want Want mini 60 г (%d ₴, 3 мережі) та Ultra Pop.' % round(ww60['pmed']),
            'Митна база: %s т нори-снеків за 17 місяців, %d %% з Кореї, $%s/кг; рисових крекерів у коді 1905 90 55 00 — лише %s т.' % (X.n1(X.C['seaweed']['all']['tons']), round(100 * X.C['seaweed']['origins'][0]['share']), X.n1(X.C['seaweed']['all']['usd_kg']), X.n1(X.C['rice']['rice_cracker']['tons'], 1))]),
        ('ДЕ МИ В ЦІНІ', AMBER, [
            'ZEK Tempura 30 г — %d ₴ (між Tao Kae Noi 25 г за %d ₴ і 40 г за %d ₴), 50 г — %d ₴. Topping 35 / 70 г — %d / %d ₴, нижче за Ock Dong Ja.' % (round(t30['shelf']), round(tk25['pmed']), round(tk40['pmed']), round(t50['shelf']), round(top35['shelf']), round(top70['shelf'])),
            'ZEK Sandwich 25 г — %d ₴: вище за Norris 25 г (%d ₴), нижче за Ock Dong Ja 25 г (%d ₴).' % (round(sd['shelf']), round(next(l for l in lines(rows_of('chips')) if l['brand'] == 'Norris' and l['grams'] == 25.0)['pmed']), round(next(l for l in lines(rows_of('chips')) if l['brand'] == 'Ock Dong Ja' and l['grams'] == 25.0)['pmed'])),
            'TMK ROLL 2,5 г — %d ₴, DOUBLE 5 г — %d ₴: у ціновому коридорі мініпакетів (типова 42 ₴). SEAWEED / MINI 10–12 г — %d–%d ₴: дорожче за всіх сусідів за вагою.' % (round(roll['shelf']), round(dbl['shelf']), round(sw['shelf']), round(mini['shelf'])),
            'Thai-Nichi 50 / 55 г — %d / %d ₴, Singha 42 г — %d ₴ проти Want Want 60 г — %d ₴.' % (round(tn50['shelf']), round(tn55['shelf']), round(sk['shelf']), round(ww60['pmed']))]),
        ('ЩО З ЦИМ РОБИТИ', GREEN, [
            'ZEK: темпура й топінг мають найчистіше позиціонування — ціна в коридорі ринку, прямий конкурент один.',
            'TMK: ROLL і DOUBLE можна виводити; SEAWEED / MINI 10–12 г перерахувати — навіть при марже 15 %% полиця %d–%d ₴ (див. чутливість).' % (min(round(r['shelf_by_margin']['0.15']) for r in D.PM['sensitivity'] if r['title'].startswith(('Wow SEAWEED', 'Wow MINI'))), max(round(r['shelf_by_margin']['0.15']) for r in D.PM['sensitivity'] if r['title'].startswith(('Wow SEAWEED', 'Wow MINI')))),
            'Thai-Nichi 50 г (%d ₴) — нижче Want Want (%d ₴); Thai-Nichi 55 г (%d ₴) і Singha (%d ₴) — вище: шукати дешевший контейнер або знижувати маржу.' % (round(tn50['shelf']), round(ww60['pmed']), round(tn55['shelf']), round(sk['shelf'])),
            'Для всіх: перед замовленням підтвердити ціну на живій полиці й окремо в азійських магазинах.']),
        ('ОБМЕЖЕННЯ ЗРІЗУ', MUTED, [
            'Митна база охоплює лише 4 коди УКТ ЗЕД і не розбита на SKU: товари під іншими кодами (наприклад, Want Want) у ній не видно.',
            'Ціни — онлайн-каталоги одного магазину кожної мережі; офлайн-акцій і знижок «в залі» немає.',
            'Арарe/норімакі не знайдено в мережах — це «не знайдено в каталогах», а не доказ відсутності.',
            'Фінмодель — за формулою рису: бонус мережі 25 %, наша маржа 30 %, полиця = ціна партнеру × 1,40; реальні умови мереж можуть відрізнятись.'])]
    w = 2.95
    for i, (ttl, col, pts) in enumerate(cols):
        x = 0.62 + i * (w + 0.1)
        rect(s, x, 1.85, w, 0.36, fill=col)
        text(s, x + 0.14, 1.92, w - 0.2, 0.22, ttl, size=8.2, color=WHITE, bold=True, spc=0.2)
        rect(s, x, 2.21, w, 4.85, fill=PANEL)
        y = 2.35
        for pt in pts:
            rect(s, x + 0.14, y + 0.07, 0.07, 0.07, fill=col)
            text(s, x + 0.3, y, w - 0.42, 1.2, pt, size=9.4 if len(pts) > 4 else 10, color=INK)
            y += 1.02 if len(pts) > 4 else 1.28
    return s


def media_hashes(path):
    with zipfile.ZipFile(path) as z:
        return {n: hashlib.md5(z.read(n)).hexdigest() for n in z.namelist() if n.startswith('ppt/media/')}


def slide_xml(path, n, strip_text=False):
    with zipfile.ZipFile(path) as z:
        root = etree.fromstring(z.read('ppt/slides/slide%d.xml' % n))
    if strip_text:
        for r in list(root.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}r')):
            r.getparent().remove(r)
    return etree.tostring(root, method='c14n')


def build(prs):
    deck = Deck(prs, len(prs.slides))
    M.cover(deck)
    M.method(deck)
    X.glossary(deck)
    X.passports(deck)
    X.customs_overview(deck)
    X.customs_seaweed(deck)
    X.customs_importers(deck)
    X.customs_rice(deck)
    M.segmentation(deck)
    # --- segment 1: mini
    ls = lines(rows_of('mini')); rs = rows_of('mini')
    M.seg_overview(deck, 'mini', 'Міні-пакет 4–5 г: найбільший сегмент, медіана %d ₴' % round(med(r['pmed'] for r in rs)),
                   note='Основа полиці: один–два аркуші в пакетику, ціна упаковки в усіх брендів близька. Сюди потрапляють наші TMK ROLL і DOUBLE ROLL.',
                   bins=[(0, 35), (35, 40), (40, 45), (45, 55), (55, 80), (80, 1000)])
    pk = dict(M.pack_table(['mini']))['mini'][0]
    M.popular_packs(deck, ['mini'], 'Міні-пакет: %s — найходовіша, %d %% присутності' % (gfmt(pk['grams']), round(100 * pk['share'])), 'Сегмент · Норі-снек у міні-пакеті · популярні упаковки')
    catalog_slides(deck, ['mini'], None, 'Сегмент · Норі-снек у міні-пакеті · %d позицій' % len(rs), 'Міні-пакет')
    M.where_sold(deck, ['mini'], 'Міні-пакет: де продається', 'Сегмент · Норі-снек у міні-пакеті · де продається')
    O.price_by_chain(deck, ['mini'], 'Міні-пакет: ціна конкурентів у кожній мережі', 'Сегмент · Норі-снек у міні-пакеті · ціни по мережах', must=('Akura',))
    # --- segment 2: chips
    rs = rows_of('chips')
    M.seg_overview(deck, 'chips', 'Чипси та сендвічі 8–40 г: %d позицій, медіана %d ₴' % (len(rs), round(med(r['pmed'] for r in rs))),
                   note='Чипси Hokkaido Club і Norris 25 г, сендвічі Norris 8 г, Ock Dong Ja 15–25 г, Ocean Snack 40 г. Сюди ставимо ZEK Sandwich і TMK SEAWEED/MINI.',
                   bins=[(0, 35), (35, 55), (55, 80), (80, 110), (110, 150), (150, 1000)])
    pk = dict(M.pack_table(['chips']))['chips'][0]
    M.popular_packs(deck, ['chips'], 'Чипси: %s — найходовіша, %d %% присутності' % (gfmt(pk['grams']), round(100 * pk['share'])), 'Сегмент · Норі-чипси та сендвічі · популярні упаковки')
    catalog_slides(deck, ['chips'], None, 'Сегмент · Норі-чипси та сендвічі · %d позицій' % len(rs), 'Чипси та сендвічі')
    M.where_sold(deck, ['chips'], 'Чипси та сендвічі: де продається', 'Сегмент · Норі-чипси та сендвічі · де продається')
    O.price_by_chain(deck, ['chips'], 'Чипси: ціна Hokkaido Club і Norris по мережах', 'Сегмент · Норі-чипси та сендвічі · ціни по мережах', must=('Hokkaido Club',))
    # --- segment 3: tempura & topping
    rs = rows_of('tempura')
    M.seg_overview(deck, 'tempura', 'Темпура й топінг: %d позицій, медіана %d ₴' % (len(rs), round(med(r['pmed'] for r in rs))),
                   note='Tao Kae Noi — єдиний бренд темпури в мережах; топінг — подрібнені норі Ock Dong Ja і Metro Chef. Конкуренти ZEK Tempura й Topping.',
                   bins=[(0, 80), (80, 110), (110, 130), (130, 160), (160, 250), (250, 1000)], rep_n=6)
    pk = dict(M.pack_table(['tempura']))['tempura'][0]
    M.popular_packs(deck, ['tempura'], 'Темпура й топінг: %s — найходовіша, %d %%' % (gfmt(pk['grams']), round(100 * pk['share'])), 'Сегмент · Темпура й топінг · популярні упаковки')
    catalog_slides(deck, ['tempura'], None, 'Сегмент · Темпура й топінг · %d позицій' % len(rs), 'Темпура й топінг')
    M.where_sold(deck, ['tempura'], 'Темпура й топінг: де продається', 'Сегмент · Темпура й топінг · де продається')
    O.price_by_chain(deck, ['tempura'], 'Темпура й топінг: ціна конкурентів у кожній мережі', 'Сегмент · Темпура й топінг · ціни по мережах', must=('Tao Kae Noi',))
    # --- segment 4/5: rice
    rs = rows_of('ricecr')
    M.seg_overview(deck, 'ricecr', 'Рисові крекери: %d позицій, медіана %d ₴' % (len(rs), round(med(r['pmed'] for r in rs))),
                   note='Сегмент Singha Kameda і Thai-Nichi. Прямий аналог у мережах — Want Want mini 60 г; решта — Ultra Pop (аніме-серія) й органічні Clearspring.',
                   bins=[(0, 100), (100, 125), (125, 135), (135, 145), (145, 150), (150, 1000)], rep_n=3)
    M.popular_packs(deck, ['ricecr', 'ricechip'], 'Рис: ходові упаковки крекерів і чипсів', 'Сегмент · Рисові крекери та чипси · популярні упаковки')
    catalog_slides(deck, ['ricecr', 'ricechip'], None, 'Сегмент · Рисові крекери та чипси · %d позицій' % (len(rows_of('ricecr')) + len(rows_of('ricechip'))), 'Рис: крекери та чипси')
    M.where_sold(deck, ['ricecr', 'ricechip'], 'Рисові крекери та чипси: де продається', 'Сегмент · Рисові крекери та чипси · де продається')
    # --- cross-segment maps
    O.brand_network(deck)
    O.brand_directory(deck)
    O.weight_map(deck)
    O.audit_networks(deck)
    O.channels(deck)
    # --- our price
    O.pricing_formula(deck, ['singha', 'thainichi', 'tmk'], 1)
    O.pricing_formula(deck, ['zek'], 2)
    cards = [l for l in lines(rows_of('ricecr'))]
    O.shelf_table(deck, ['singha', 'thainichi'], market_cards=cards)
    O.shelf_table(deck, ['tmk'])
    O.shelf_table(deck, ['zek'])
    O.verification(deck)
    O.scenarios(deck)
    O.sensitivity(deck)
    O.pack_price_map(deck)
    def Ln(sg, brand, g):
        return next(l for l in lines(rows_of(sg)) if l['brand'] == brand and l['grams'] == g)
    def O1(key, g=None):
        return next(o for o in OUR if key in o['key'] and (g is None or o['grams'] == g))
    P = lambda v: '%d ₴' % round(v)
    below = lambda v, sg: round(100 * sum(1 for r in rows_of(sg) if r['pmed'] < v) / len(rows_of(sg)))
    roll, dbl = O1('Wow ROLL - Original'), O1('DOUBLE ROLL - Original')
    D_ = lambda o, an: o['shelf'] - an['pmed']
    def vs(o):
        an = D.analog(o)
        d = o['shelf'] - an['pmed']
        return '%s ₴ проти %s %s — %s %d ₴ (%+d %%)' % (round(o['shelf']), an['brand'], gfmt(an['grams']) + ' ' + str(round(an['pmed'])), 'вище на' if d > 0 else 'нижче на', abs(round(d)), round(d / an['pmed'] * 100))
    O.stand(deck, ['mini', 'chips'], 'Хто з ким стоїть · TMK, пакети 2,5–12 г', 'Де ми стоїмо: TMK — від міні-пакета 2,5 г до 12 г', 20, 140, 20,
            ours_filter=lambda o: o['sup'] == 'tmk', comp_filter=lambda l: l['grams'] <= 15.5, seg_comp=['mini', 'chips'],
            sub='TMK на ринку: ROLL 2,5 г — %d ₴, DOUBLE ROLL 5 г — %d ₴, MINI 10 г — %d ₴, SEAWEED 12 г — %d ₴. Поруч — усі конкуренти до 15 г: хто на полицях, скільки позицій, типова ціна й у яких мережах стоїть (найходовіша упаковка ринку — 4,5 г, типова ціна %d ₴).' % (round(roll['shelf']), round(dbl['shelf']), round(O1('WOW MINI - Original')['shelf']), round(O1('WOW SEAWEED - Original')['shelf']), round(Ln('mini', 'Akura', 4.5)['pmed'])),
            takeaways=None)
    sw, mn, sd = O1('WOW SEAWEED - Original'), O1('WOW MINI - Original'), O1('SANDWICH SEAWEED - Sesame')
    n25 = Ln('chips', 'Norris', 25.0); h25 = Ln('chips', 'Hokkaido Club', 25.0); os40 = Ln('chips', 'Ocean Snack', 40.0)
    O.stand(deck, ['chips'], 'Хто з ким стоїть · чипси та сендвічі', 'Де ми стоїмо: чипси та сендвічі', 0, 160, 40,
            sub='ZEK Sandwich 25 г — %d ₴ (між Norris і Ock Dong Ja); TMK SEAWEED / MINI 10–12 г — %d–%d ₴: на рівні 40-грамових чипсів Ocean Snack (%d ₴).' % (round(sd['shelf']), round(sw['shelf']), round(mn['shelf']), round(os40['pmed'])),
            takeaways=[('ZEK SANDWICH 25 Г', '%s ₴ — вище за Norris 25 г (%s) і Hokkaido Club 25 г (%s), нижче за Ock Dong Ja 25 г (%s).' % (round(sd['shelf']), P(n25['pmed']), P(h25['pmed']), P(Ln('chips', 'Ock Dong Ja', 25.0)['pmed'])), AMBER),
                       ('TMK SEAWEED / MINI 10–12 Г', '%s ₴ за пакет 10–12 г — на рівні Ocean Snack 40 г (%s). Сусіди за вагою: Norris 8–10 г — 31–33 ₴, Ock Dong Ja 15 г — %s.' % (round(sw['shelf']), P(os40['pmed']), P(Ln('chips', 'Ock Dong Ja', 15.0)['pmed'])), RED)])
    t30, t50 = O1('TEMPURA SEAWEED - Corn 30g'), O1('TEMPURA SEAWEED - Corn 50g')
    tk25, tk40 = Ln('tempura', 'Tao Kae Noi', 25.0), Ln('tempura', 'Tao Kae Noi', 40.0)
    z35, z70 = O1('TOPPING - Vegetables 35g'), O1('TOPPING - Vegetables 70g')
    o35, o70, mc = Ln('tempura', 'Ock Dong Ja', 35.0), Ln('tempura', 'Ock Dong Ja', 70.0), Ln('tempura', 'Metro Chef', 30.0)
    O.stand(deck, ['tempura'], 'Хто з ким стоїть · темпура й топінг', 'Де ми стоїмо: темпура й топінг', 60, 260, 40,
            sub='ZEK Tempura 30 г — %d ₴ і 50 г — %d ₴ поряд із Tao Kae Noi 25 г (%d ₴) і 40 г (%d ₴); ZEK Topping 35 / 70 г — %d / %d ₴.' % (round(t30['shelf']), round(t50['shelf']), round(tk25['pmed']), round(tk40['pmed']), round(z35['shelf']), round(z70['shelf'])),
            takeaways=[('ZEK TEMPURA', '30 г — %s: між Tao Kae Noi 25 г (%s) і 40 г (%s). 50 г — %s: дорожче за Tao Kae Noi 40 г.' % (P(t30['shelf']), P(tk25['pmed']), P(tk40['pmed']), P(t50['shelf'])), PURPLE),
                       ('ZEK TOPPING', '35 г — %s, 70 г — %s. Ock Dong Ja: 35 г — %s, 70 г — %s; Metro Chef 30 г — %s. ZEK дешевший.' % (P(z35['shelf']), P(z70['shelf']), P(o35['pmed']), P(o70['pmed']), P(mc['pmed'])), TEAL)])
    ww = Ln('ricecr', 'Want Want', 60.0)
    sk, t50r, t55r = O1('Original Flavour'), O1('Norimaki Original'), O1('Norimaki Wasabi')
    O.stand(deck, ['ricecr'], 'Хто з ким стоїть · рисові крекери', 'Де ми стоїмо: рисові крекери', 100, 220, 20,
            sub='Thai-Nichi 50 г — %d ₴, 55 г — %d ₴; Singha 42 г — %d ₴. Єдиний аналог у мережах — Want Want mini 60 г: %d ₴.' % (round(t50r['shelf']), round(t55r['shelf']), round(sk['shelf']), round(ww['pmed'])),
            takeaways=[('THAI-NICHI', '50 г — %s і 55 г — %s проти Want Want 60 г — %s: різниця %+d і %+d ₴.' % (P(t50r['shelf']), P(t55r['shelf']), P(ww['pmed']), round(t50r['shelf'] - ww['pmed']), round(t55r['shelf'] - ww['pmed'])), RED),
                       ('SINGHA KAMEDA', '42 г — %s: на %d ₴ (%d %%) вище за найдорожчий крекер у мережах (Ultra Pop 60 г — 148 ₴) і на %d ₴ вище за Want Want.' % (P(sk['shelf']), round(sk['shelf'] - 148), round((sk['shelf'] / 148 - 1) * 100), round(sk['shelf'] - ww['pmed'])), RED),
                       ('ЩО ВАЖЛИВО', 'Це єдині аналоги, які знайшлися в мережах: Want Want mini — 3 мережі. Арарe/норімакі на кшталт наших не знайдено — полицю варто перевірити наживо.', NAVY)])
    O.neighbours(deck)
    conclusions(deck)
    return deck


def standalone(out=os.path.join(HERE, '..', '..', 'Snacks_Market_Research_Final.pptx')):
    """The market study on its own (no supplier slides): same slides, same page order, numbered from 1."""
    from pptx.util import Inches
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(10), Inches(5.625)
    build(prs)
    prs.save(out)
    print('standalone saved: %d slides' % len(prs.slides))


def main(src=DEFAULT_SRC, out=DEFAULT_OUT):
    prs = Presentation(src)
    n0 = len(prs.slides)
    # supplier slides 3-16: make every cost figure equal to the final SelfCost (text only)
    changed = FX.fix(prs)
    fixed_only = os.path.join(HERE, '..', '..', 'Snacks_Presentation_Final_prices_fixed.pptx')
    prs.save(fixed_only)
    deck = build(prs)
    prs.save(out)
    b, a = media_hashes(src), media_hashes(out)
    assert not (set(b) - set(a)), 'media dropped'
    assert all(a[k] == v for k, v in b.items()), 'media altered'
    for n in range(1, n0 + 1):
        price = n in FX.SLIDES
        assert slide_xml(src, n, price) == slide_xml(out, n, price), 'slide %d changed beyond price text' % n
        assert slide_xml(src, n, price) == slide_xml(fixed_only, n, price)
    print('ok: %d original slides (%d price cells corrected on slides %s) + %d new = %d; %d original media files identical'
          % (n0, len(changed), sorted({c[0] for c in changed}), len(prs.slides) - n0, len(prs.slides), len(b)))


def standalone_unused():
    pass


if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == 'standalone':
        standalone()
    else:
        main(*sys.argv[1:3])
