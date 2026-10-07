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
    for i in range(n):
        chunk = items[i * per:(i + 1) * per]
        lo = min(l['pmin'] for l, _ in chunk); hi = max(l['pmax'] for l, _ in chunk)
        t = '%s: каталог %d з %d — %s ₴' % (title_base, i + 1, n, rng(lo, hi)) if n > 1 else '%s: каталог — %s ₴' % (title_base, rng(lo, hi))
        M.catalog(deck, chunk, eyebrow, t)


def conclusions(deck):
    s = deck.slide('Висновки', 'Що показав зріз ринку і де ми стоїмо')
    S = {sg: (rows_of(sg), lines(rows_of(sg))) for sg in SEGS}
    def ps(key_part, g=None):
        o = [x for x in OUR if key_part in x['key'] and (g is None or x['grams'] == g)][0]
        p = D.peer_stats(o)
        return o, p, (o['per_g'] / p['per_g'] - 1) * 100
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
    cols = [
        ('ЩО КАЖЕ РИНОК', NAVY, [
            'Полиця мереж — це мініпакет 4–5 г: %d SKU, %d брендів, медіана %d ₴ (%s ₴/г).' % (len(mrs), len({r['brand'] for r in mrs}), round(med(r['pmed'] for r in mrs)), fm(med(r['per_g'] for r in mrs))),
            'Чипси 8–40 г — лише 4 бренди (Norris, Hokkaido Club, Ock Dong Ja, Ocean Snack): %s ₴.' % rng(min(r['pmin'] for r in S['chips'][0]), max(r['pmax'] for r in S['chips'][0])),
            'Темпура — практично один бренд, Tao Kae Noi: 25 г — %d ₴, 40 г — %d ₴.' % (round(tk25['pmed']), round(tk40['pmed'])),
            'Рисові крекери в мережах — лише Want Want mini 60 г (%d ₴, 3 мережі) і Ultra Pop.' % round(wwl['pmed'])]),
        ('ДЕ МИ В ЦІНІ', AMBER, [
            'ZEK Tempura 30 г — %d ₴, 50 г — %d ₴: за грам %s–%s ₴ проти %s–%s у Tao Kae Noi. Topping 35/70 г — %d/%d ₴, нижче за Ock Dong Ja.' % (round(t30['shelf']), round(t50['shelf']), fm(t50['per_g']), fm(t30['per_g']), fm(tk40['per_g']), fm(tk25['per_g']), round(top35['shelf']), round(top70['shelf'])),
            'TMK ROLL 2,5 г — %d ₴, DOUBLE 5 г — %d ₴: упаковка в коридорі мініпакетів, за грам %s…%s до медіани.' % (round(roll['shelf']), round(dbl['shelf']), O.pct(roll['per_g'], med(r['per_g'] for r in mrs)), O.pct(dbl['per_g'], med(r['per_g'] for r in mrs))),
            'TMK SEAWEED/MINI 10–12 г — %d–%d ₴ за упаковку, у 3–4 рази дорожче за сусідні за вагою чипси за грам.' % (round(sw['shelf']), round(mini['shelf'])),
            'Thai-Nichi 50–55 г — %d–%d ₴, Singha 42 г — %d ₴ проти Want Want 60 г — %d ₴.' % (round(tn50['shelf']), round(tn55['shelf']), round(sk['shelf']), round(wwl['pmed']))]),
        ('ЩО З ЦИМ РОБИТИ', GREEN, [
            'ZEK: темпура й топінг мають найчистіше позиціонування — ціна в коридорі ринку, прямий конкурент один.',
            'TMK: ROLL/DOUBLE можна виводити; SEAWEED/MINI 10–12 г перерахувати — навіть при марже 15 % полиця ~81–82 ₴ (див. чутливість).',
            'Thai-Nichi: упаковка на рівні Want Want; Singha — шукати дешевший контейнерний сценарій або зменшувати маржу.',
            'Для всіх: перед замовленням підтвердити ціну на живій полиці й окремо в азійських магазинах.']),
        ('ОБМЕЖЕННЯ ЗРІЗУ', MUTED, [
            'Немає митної бази — обсяги ввозу категорії невідомі.',
            'Ціни — онлайн-каталоги одного магазину кожної мережі; офлайн-акцій і знижок «в залі» немає.',
            'Арарe/норімакі не знайдено в мережах — це «не знайдено в каталогах», а не доказ відсутності.',
            'Фінмодель — за формулою рису (бонус 25 %, маржа 35 %, ×1,40); реальні умови мереж можуть відрізнятись.'])]
    w = 2.95
    for i, (ttl, col, pts) in enumerate(cols):
        x = 0.62 + i * (w + 0.1)
        rect(s, x, 1.5, w, 0.36, fill=col)
        text(s, x + 0.14, 1.57, w - 0.2, 0.22, ttl, size=8.2, color=WHITE, bold=True, spc=0.2)
        rect(s, x, 1.86, w, 5.3, fill=PANEL)
        y = 2.0
        for pt in pts:
            rect(s, x + 0.14, y + 0.07, 0.07, 0.07, fill=col)
            text(s, x + 0.3, y, w - 0.42, 1.2, pt, size=10, color=INK)
            y += 1.28
    return s


def media_hashes(path):
    with zipfile.ZipFile(path) as z:
        return {n: hashlib.md5(z.read(n)).hexdigest() for n in z.namelist() if n.startswith('ppt/media/')}


def slide_xml(path, n):
    with zipfile.ZipFile(path) as z:
        return etree.tostring(etree.fromstring(z.read('ppt/slides/slide%d.xml' % n)), method='c14n')


def build(prs):
    deck = Deck(prs, len(prs.slides))
    M.cover(deck)
    M.method(deck)
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
    # --- segment 2: chips
    rs = rows_of('chips')
    M.seg_overview(deck, 'chips', 'Чипси та сендвічі 8–40 г: %d позицій, медіана %d ₴' % (len(rs), round(med(r['pmed'] for r in rs))),
                   note='Чипси Hokkaido Club і Norris 25 г, сендвічі Norris 8 г, Ock Dong Ja 15–25 г, Ocean Snack 40 г. Сюди ставимо ZEK Sandwich і TMK SEAWEED/MINI.',
                   bins=[(0, 35), (35, 55), (55, 80), (80, 110), (110, 150), (150, 1000)])
    pk = dict(M.pack_table(['chips']))['chips'][0]
    M.popular_packs(deck, ['chips'], 'Чипси: %s — найходовіша, %d %% присутності' % (gfmt(pk['grams']), round(100 * pk['share'])), 'Сегмент · Норі-чипси та сендвічі · популярні упаковки')
    catalog_slides(deck, ['chips'], None, 'Сегмент · Норі-чипси та сендвічі · %d позицій' % len(rs), 'Чипси та сендвічі')
    M.where_sold(deck, ['chips'], 'Чипси та сендвічі: де продається', 'Сегмент · Норі-чипси та сендвічі · де продається')
    # --- segment 3: tempura & topping
    rs = rows_of('tempura')
    M.seg_overview(deck, 'tempura', 'Темпура й топінг: %d позицій, медіана %d ₴' % (len(rs), round(med(r['pmed'] for r in rs))),
                   note='Tao Kae Noi — єдиний бренд темпури в мережах; топінг — подрібнені норі Ock Dong Ja і Metro Chef. Конкуренти ZEK Tempura й Topping.',
                   bins=[(0, 80), (80, 110), (110, 130), (130, 160), (160, 250), (250, 1000)], rep_n=6)
    pk = dict(M.pack_table(['tempura']))['tempura'][0]
    M.popular_packs(deck, ['tempura'], 'Темпура й топінг: %s — найходовіша, %d %%' % (gfmt(pk['grams']), round(100 * pk['share'])), 'Сегмент · Темпура й топінг · популярні упаковки')
    catalog_slides(deck, ['tempura'], None, 'Сегмент · Темпура й топінг · %d позицій' % len(rs), 'Темпура й топінг')
    M.where_sold(deck, ['tempura'], 'Темпура й топінг: де продається', 'Сегмент · Темпура й топінг · де продається')
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
    O.weight_map(deck)
    O.audit_networks(deck)
    O.channels(deck)
    # --- our price
    O.pricing_formula(deck)
    cards = [l for l in lines(rows_of('ricecr'))]
    O.shelf_table(deck, ['singha', 'thainichi'], market_cards=cards)
    O.shelf_table(deck, ['tmk'])
    O.shelf_table(deck, ['zek'])
    O.pack_price_map(deck)
    O.per_g_map(deck)
    def Ln(sg, brand, g):
        return next(l for l in lines(rows_of(sg)) if l['brand'] == brand and l['grams'] == g)
    def O1(key, g=None):
        return next(o for o in OUR if key in o['key'] and (g is None or o['grams'] == g))
    P = lambda v: '%d ₴' % round(v)
    below = lambda v, sg: round(100 * sum(1 for r in rows_of(sg) if r['pmed'] < v) / len(rows_of(sg)))
    roll, dbl = O1('Wow ROLL - Original'), O1('DOUBLE ROLL - Original')
    mg = med(r['per_g'] for r in rows_of('mini'))
    O.stand(deck, ['mini'], 'Хто з ким стоїть · міні-пакет', 'Де ми стоїмо: міні-пакет 4–5 г', 20, 100, 20,
            sub='Наші TMK ROLL 2,5 г і DOUBLE ROLL 5 г — у ціновому коридорі мініпакетів; за грам ми дорожчі, бо ринок — це 4–5 г за ~42 ₴.',
            takeaways=[('ЦІНА УПАКОВКИ', 'ROLL 2,5 г — %s: дешевше за %d %% позицій сегмента. DOUBLE 5 г — %s: дорожче за %d %% позицій, нижче лише за Kimnori (77–98 ₴) і Clearspring (84–89 ₴).' % (P(roll['shelf']), 100 - below(roll['shelf'], 'mini'), P(dbl['shelf']), below(dbl['shelf'], 'mini')), NAVY),
                       ('ЦІНА ЗА ГРАМ', 'ROLL — %s ₴/г (%s), DOUBLE — %s ₴/г (%s) до медіани сегмента %s ₴/г. Найближчий за форматом аналог — Norris Рол 3 г: 30–32 ₴.' % (fm(roll['per_g']), O.pct(roll['per_g'], mg), fm(dbl['per_g']), O.pct(dbl['per_g'], mg), fm(mg)), AMBER)])
    sw, mn, sd = O1('WOW SEAWEED - Original'), O1('WOW MINI - Original'), O1('SANDWICH SEAWEED - Sesame')
    cm = med(r['per_g'] for r in rows_of('chips'))
    n25 = Ln('chips', 'Norris', 25.0); h25 = Ln('chips', 'Hokkaido Club', 25.0); os40 = Ln('chips', 'Ocean Snack', 40.0)
    O.stand(deck, ['chips'], 'Хто з ким стоїть · чипси та сендвічі', 'Де ми стоїмо: чипси та сендвічі', 0, 160, 40,
            sub='ZEK Sandwich 25 г стоїть між Norris і Ock Dong Ja; TMK SEAWEED/MINI 10–12 г по 122–123 ₴ — на рівні 40-грамових чипсів Ocean Snack.',
            takeaways=[('ZEK SANDWICH 25 Г', '%s ₴ — вище за Norris 25 г (%s) і Hokkaido Club 25 г (%s), нижче за Ock Dong Ja 25 г (%s). За грам %s ₴ проти %s–%s у лідерів 25 г.' % (round(sd['shelf']), P(n25['pmed']), P(h25['pmed']), P(Ln('chips', 'Ock Dong Ja', 25.0)['pmed']), fm(sd['per_g']), fm(h25['per_g']), fm(n25['per_g'])), AMBER),
                       ('TMK SEAWEED / MINI 10–12 Г', '%s ₴ за 10–12 г — стільки ж, скільки Ocean Snack 40 г (%s) і більше за сусідів за вагою (Norris 8–10 г: 31–33 ₴, Ock Dong Ja 15 г: %s). За грам %s–%s ₴ — у 3–4 рази вище за чипси.' % (round(sw['shelf']), P(os40['pmed']), P(Ln('chips', 'Ock Dong Ja', 15.0)['pmed']), fm(sw['per_g']), fm(mn['per_g'])), RED)])
    t30, t50 = O1('TEMPURA SEAWEED - Corn 30g'), O1('TEMPURA SEAWEED - Corn 50g')
    tk25, tk40 = Ln('tempura', 'Tao Kae Noi', 25.0), Ln('tempura', 'Tao Kae Noi', 40.0)
    z35, z70 = O1('TOPPING - Vegetables 35g'), O1('TOPPING - Vegetables 70g')
    o35, o70, mc = Ln('tempura', 'Ock Dong Ja', 35.0), Ln('tempura', 'Ock Dong Ja', 70.0), Ln('tempura', 'Metro Chef', 30.0)
    O.stand(deck, ['tempura'], 'Хто з ким стоїть · темпура й топінг', 'Де ми стоїмо: темпура й топінг', 60, 260, 40,
            sub='ZEK Tempura 30 і 50 г стоять поруч із Tao Kae Noi 25 і 40 г; ZEK Topping дешевший за Ock Dong Ja й Metro Chef.',
            takeaways=[('ZEK TEMPURA', '30 г — %s: між Tao Kae Noi 25 г (%s) і 40 г (%s). 50 г — %s. За грам %s–%s ₴ проти %s–%s у Tao Kae Noi 25–40 г.' % (P(t30['shelf']), P(tk25['pmed']), P(tk40['pmed']), P(t50['shelf']), fm(t50['per_g']), fm(t30['per_g']), fm(tk40['per_g']), fm(tk25['per_g'])), PURPLE),
                       ('ZEK TOPPING', '35 г — %s, 70 г — %s. Ock Dong Ja: 35 г — %s, 70 г — %s; Metro Chef 30 г — %s. За грам ZEK %s ₴ проти %s–%s.' % (P(z35['shelf']), P(z70['shelf']), P(o35['pmed']), P(o70['pmed']), P(mc['pmed']), fm(z35['per_g']), fm(o70['per_g']), fm(o35['per_g'])), TEAL)])
    ww = Ln('ricecr', 'Want Want', 60.0)
    sk, t50r, t55r = O1('Original Flavour'), O1('Norimaki Original'), O1('Norimaki Wasabi')
    O.stand(deck, ['ricecr'], 'Хто з ким стоїть · рисові крекери', 'Де ми стоїмо: рисові крекери', 100, 220, 20,
            sub='Thai-Nichi 50–55 г — на верхній межі Want Want mini 60 г (140–149 ₴) і вище; Singha 42 г — вище за весь діапазон мережевої полиці.',
            takeaways=[('THAI-NICHI', '50 г — %s і 55 г — %s проти Want Want 60 г — %s. За грам %s–%s ₴ проти %s: +%d…+%d %%.' % (P(t50r['shelf']), P(t55r['shelf']), P(ww['pmed']), fm(t50r['per_g']), fm(t55r['per_g']), fm(ww['per_g']), round((t50r['per_g'] / ww['per_g'] - 1) * 100), round((t55r['per_g'] / ww['per_g'] - 1) * 100)), RED),
                       ('SINGHA KAMEDA', '42 г — %s: на %d %% вище за найдорожчий крекер у мережах (149 ₴). За грам %s ₴ проти %s у Want Want: %s.' % (P(sk['shelf']), round((sk['shelf'] / 148 - 1) * 100), fm(sk['per_g']), fm(ww['per_g']), O.pct(sk['per_g'], ww['per_g'])), RED),
                       ('ЩО ВАЖЛИВО', 'Це єдині аналоги, які знайшлися в мережах: Want Want mini — 3 мережі. Арарe/норімакі на кшталт наших не знайдено — полицю варто перевірити наживо.', NAVY)])
    O.neighbours(deck)
    O.sensitivity(deck)
    conclusions(deck)
    return deck


def main(src=DEFAULT_SRC, out=DEFAULT_OUT):
    prs = Presentation(src)
    n0 = len(prs.slides)
    deck = build(prs)
    prs.save(out)
    # --- the original slides must be untouched: media byte-identical, slide XML canonically equal
    b, a = media_hashes(src), media_hashes(out)
    assert not (set(b) - set(a)), 'media dropped'
    assert all(a[k] == v for k, v in b.items()), 'media altered'
    for n in range(1, n0 + 1):
        assert slide_xml(src, n) == slide_xml(out, n), 'slide %d changed' % n
    print('ok: %d original slides + %d new = %d; %d original media files identical' % (n0, len(prs.slides) - n0, len(prs.slides), len(b)))


if __name__ == '__main__':
    main(*sys.argv[1:3])
