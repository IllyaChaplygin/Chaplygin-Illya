#!/usr/bin/env python3
"""Сборка Excel-книги со всеми листами и графиками."""
import csv, json, os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, ScatterChart, Reference, Series

HERE = os.path.dirname(__file__)
DATA = os.path.join(HERE, '..', 'data')
OUTX = os.path.join(DATA, 'land_jv_usa_research.xlsx')

HDR_FILL = PatternFill('solid', fgColor='1F3A5F')
HDR_FONT = Font(bold=True, color='FFFFFF', size=10)
TITLE_FONT = Font(bold=True, size=14, color='1F3A5F')
NOTE_FONT = Font(italic=True, size=9, color='666666')
THIN = Side(style='thin', color='D0D0D0')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def style_header(ws, row=1, ncols=None):
    ncols = ncols or ws.max_column
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL; cell.font = HDR_FONT
        cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
    ws.row_dimensions[row].height = 42
    ws.freeze_panes = ws.cell(row=row + 1, column=1)


def autosize(ws, maxw=52, minw=9):
    for col in range(1, ws.max_column + 1):
        best = minw
        for row in range(1, min(ws.max_row, 400) + 1):
            v = ws.cell(row=row, column=col).value
            if v is None: continue
            L = max(len(x) for x in str(v).split('\n'))
            best = max(best, min(L + 2, maxw))
        ws.column_dimensions[get_column_letter(col)].width = best


def read_csv(name):
    with open(os.path.join(DATA, name), encoding='utf-8') as f:
        return list(csv.DictReader(f))


def main():
    wb = Workbook(); wb.remove(wb.active)

    # ---------------- 0. Методика ----------------
    ws = wb.create_sheet('0. Методика')
    rows = [
        ('ИССЛЕДОВАНИЕ: земельные игроки США, привлекающие инвесторов в совместные проекты', ''),
        ('Дата среза', '5 октября 2026'),
        ('Рынок', 'США, национальный охват'),
        ('', ''),
        ('ГЛАВНЫЙ ВЫВОД', ''),
        ('Исходная гипотеза', 'Крупнейшие земельные игроки ведут агрессивную рекламу во всех каналах'),
        ('Результат проверки', 'ГИПОТЕЗА НЕ ПОДТВЕРДИЛАСЬ. Признаки обратно коррелированы.'),
        ('Доказательство', 'У всех 13 крупнейших землевладельцев выборки — 0 активных объявлений Google. '
                           'Рекламируются те, у кого земли мало или нет.'),
        ('Причина №1 (юридическая)', 'Reg D Rule 506(b) ЗАПРЕЩАЕТ general solicitation. '
                                     'Walton: 52 из 52 предложений — 506(b). Crow Holdings: 54 из 54 — 506(b).'),
        ('Причина №2 (экономическая)', 'Чек $1-20 млн не покупается кликом. CPL аккредитованного инвестора $50-500, '
                                       'стоимость профинансировавшего инвестора $3 500-4 500.'),
        ('Причина №3 (организационная)', 'Капитал берут там, где он лежит: институциональные LP, '
                                         'страховые балансы, сети wealth-менеджеров. Walton: 300+ каналов дистрибуции.'),
        ('', ''),
        ('ФИЛЬТР ВКЛЮЧЕНИЯ (нужно пройти все 5)', ''),
        ('F1', 'Основной актив — земля (raw / pre-development / entitled), а не готовые здания'),
        ('F2', 'Портфель ≥10 отдельных земельных объектов или предложений'),
        ('F3', 'Модель — привлечь капитал к участию в стоимости земли, а НЕ брокеридж за комиссию'),
        ('F4', 'Институциональный масштаб: раунд ≥$5 млн, чеки от $100K, принимаются $1 млн+'),
        ('F5', 'Не государственная структура (не муниципальный land bank)'),
        ('', ''),
        ('ИСКЛЮЧЕНЫ', ''),
        ('Земельные брокеры', 'Land Advisors, National Land Realty, Whitetail Properties — нарушают F3'),
        ('Маркетплейсы', 'Fundrise, CrowdStreet, RealtyMogul, EquityMultiple, Yieldstreet — нарушают F1/F2. '
                         'Вынесены в панель-бенчмарк.'),
        ('Муниципальные land banks', 'Detroit Land Bank Authority и аналоги — нарушают F5'),
        ('Застройщики-продавцы жилья', 'Brookfield Residential: 700 объявлений, но это продажа ДОМОВ, не привлечение инвесторов'),
        ('', ''),
        ('МОДЕЛЬ ОЦЕНКИ — 100 баллов', ''),
        ('Блок A. Масштаб земли', '30 баллов: акры (12) + число объектов (10) + число рынков (8)'),
        ('Блок B. Капитальная машина', '30 баллов: AUM (10) + привлечено (10) + число Reg D-эмитентов (10)'),
        ('Блок C. Маркетинговая интенсивность', '30 баллов: Google-реклама (8) + соцсети (5) + контент (7) + '
                                                 'дистрибуция/события (6) + органика (4)'),
        ('Блок D. Инфраструктура конверсии', '10 баллов: воронка (5) + статус 506(c), т.е. право рекламировать (5)'),
        ('', ''),
        ('СТАТУС ДАННЫХ', ''),
        ('FACT-M (измерено)', 'Снято напрямую: Google ATC, SEC EDGAR Form D, YouTube. Воспроизводимо.'),
        ('FACT-D (раскрыто)', 'Заявлено компанией или в отчётности: акры, AUM, число инвесторов.'),
        ('EST (оценка)', 'Рассчитано по бенчмаркам: бюджеты в долларах, распределение по каналам.'),
        ('', ''),
        ('ГЛАВНОЕ ОГРАНИЧЕНИЕ', ''),
        ('Почему нет точных медиабюджетов',
         'Рекламные расходы частных компаний в США не раскрываются НИГДЕ. Ни Meta Ad Library, '
         'ни Google ATC не показывают суммы для неполитической рекламы — только наличие, количество и креативы. '
         'Все долларовые значения в листах 4-5 — модельные оценки с точностью порядка величины (±50-70%). '
         'Использовать как ориентир структуры, а не как бюджетный план.'),
        ('Ограничение Meta', 'Meta Ad Library закрыла поиск по ключевым словам без авторизации (HTTP 403 / login gate). '
                             'Данные по Meta — неполные, помечены отдельно.'),
        ('Смещение полноты данных', 'Частные институциональные игроки раскрывают меньше публичных. '
                                    'Столбец DATA_COMPLETENESS_PCT показывает долю заполненных ключевых полей. '
                                    'Низкий балл при низкой полноте = непрозрачность, а не малый масштаб.'),
    ]
    for r in rows: ws.append(list(r))
    ws['A1'].font = TITLE_FONT
    for r in (5, 13, 20, 26, 32, 37):
        ws.cell(row=r, column=1).font = Font(bold=True, size=11, color='1F3A5F')
    ws.column_dimensions['A'].width = 34
    ws.column_dimensions['B'].width = 112
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=2, max_col=2):
        for c in row: c.alignment = Alignment(wrap_text=True, vertical='top')

    # ---------------- 1. Рейтинг ----------------
    comps = read_csv('companies.csv')
    ws = wb.create_sheet('1. Рейтинг 20')
    head = ['#', 'Уровень', 'Компания', 'Штаб-квартира', 'Домен', 'Акры/homesites', 'Объектов',
            'Рынков', 'AUM $млн', 'Привлечено $млн', 'Reg D-эмитентов', 'Форм D', '% 506(c)',
            'Продано по формам D $млн', 'Google-объявлений', 'YouTube подп.', 'YouTube видео',
            'Мин. чек $', 'Инвесторов', 'Стран', 'Каналов дистрибуции',
            'A: земля /30', 'B: капитал /30', 'C: маркетинг /30', 'D: конверсия /10',
            'ИТОГО /100', 'Полнота данных %', 'Модель', 'Как ищут инвесторов', 'Комментарий']
    ws.append(head)
    for c in comps:
        ws.append([
            int(c['rank']), int(c['tier']), c['name'], c['hq'], c['domain'],
            c['acres'] or None, c['projects'] or None, c['markets'] or None,
            c['aum_usd_m'] or None, c['raised_usd_m'] or None, c['regd_entities'] or None,
            c['formd_count'] or None, c['pct_506c'] or None, c['formd_sold_usd_m'] or None,
            c['google_ads'] or 0, c['yt_subs'] or None, c['yt_vids'] or None,
            c['min_check_usd'] or None, c['investors'] or None, c['countries'] or None,
            c['dist_channels'] or None,
            int(c['BLOCK_A']), int(c['BLOCK_B']), int(c['BLOCK_C']), int(c['BLOCK_D']),
            int(c['TOTAL']), int(c['DATA_COMPLETENESS_PCT']),
            c['model'], c['channel_model'], c['note'],
        ])
    style_header(ws)
    for r in range(2, ws.max_row + 1):
        for col in (6, 7, 8, 9, 10, 11, 12, 14, 16, 17, 18, 19, 21):
            ws.cell(row=r, column=col).number_format = '#,##0'
        ws.cell(row=r, column=13).number_format = '0.0'
        tot = ws.cell(row=r, column=26); tot.font = Font(bold=True)
        v = tot.value
        tot.fill = PatternFill('solid', fgColor='C6E0B4' if v >= 50 else ('FFE699' if v >= 30 else 'F8CBAD'))
        g = ws.cell(row=r, column=15)
        gv = g.value if isinstance(g.value, (int, float)) else 0
        g.fill = PatternFill('solid', fgColor='F8CBAD' if gv == 0 else 'C6E0B4')
        dc = ws.cell(row=r, column=27)
        if dc.value < 50: dc.fill = PatternFill('solid', fgColor='FFE699')
        for col in (28, 29, 30):
            ws.cell(row=r, column=col).alignment = Alignment(wrap_text=True, vertical='top')
    autosize(ws, maxw=46)
    for col in (28, 29, 30): ws.column_dimensions[get_column_letter(col)].width = 60

    ch = BarChart(); ch.type = 'bar'; ch.grouping = 'stacked'; ch.overlap = 100
    ch.title = 'Структура балла по блокам (A земля / B капитал / C маркетинг / D конверсия)'
    data = Reference(ws, min_col=22, max_col=25, min_row=1, max_row=ws.max_row)
    cats = Reference(ws, min_col=3, min_row=2, max_row=ws.max_row)
    ch.add_data(data, titles_from_data=True); ch.set_categories(cats)
    ch.height, ch.width = 17, 30
    ws.add_chart(ch, 'AE2')

    # ---------------- 2. Реклама: измерения ----------------
    ws = wb.create_sheet('2. Реклама (измерено)')
    ws.append(['Прямые измерения рекламной активности, Google Ads Transparency Center, 05.10.2026'])
    ws['A1'].font = TITLE_FONT
    ws.append(['ATC показывает НАЛИЧИЕ и КОЛИЧЕСТВО объявлений, но НЕ суммы. Суммы по неполитической рекламе не раскрываются нигде.'])
    ws['A2'].font = NOTE_FONT
    ws.append([])
    ws.append(['Компания', 'Группа', 'Домен', 'Google-объявлений', '% 506(c)', 'Право рекламировать', 'Комментарий'])
    style_header(ws, row=4)
    grp = {c['name']: ('Топ-20: земельный игрок' if int(c['tier']) == 1 else 'Топ-20: розничный фонд') for c in comps}
    rows2 = []
    for c in comps:
        pct = float(c['pct_506c']) if c['pct_506c'] else None
        if 'NYSE' in c['name'] or 'NASDAQ' in c['name']:
            right = 'Да (публичная компания)'
        elif pct is None:
            right = 'Нет данных'
        elif pct == 0:
            right = 'НЕТ — всё по 506(b)'
        elif pct >= 90:
            right = 'Да — почти всё 506(c)'
        else:
            right = f'Частично ({pct:.0f}% 506(c))'
        rows2.append([c['name'], grp[c['name']], c['domain'], int(c['google_ads'] or 0), pct, right, c['note'][:160]])
    for b in read_csv('benchmark.csv'):
        rows2.append([b['name'], 'БЕНЧМАРК (не земельный игрок)', b['domain'], int(b['google_ads'] or 0), None, '—', b['note'][:160]])
    rows2.sort(key=lambda r: -r[3])
    for r in rows2: ws.append(r)
    for r in range(5, ws.max_row + 1):
        g = ws.cell(row=r, column=4)
        gv = g.value if isinstance(g.value, (int, float)) else 0
        g.fill = PatternFill('solid', fgColor='F8CBAD' if gv == 0 else ('C6E0B4' if gv >= 20 else 'FFE699'))
        ws.cell(row=r, column=5).number_format = '0.0'
        ws.cell(row=r, column=7).alignment = Alignment(wrap_text=True, vertical='top')
    autosize(ws, maxw=40); ws.column_dimensions['G'].width = 70

    ch = BarChart(); ch.title = 'Активных объявлений Google (США)'
    ch.add_data(Reference(ws, min_col=4, min_row=4, max_row=ws.max_row), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=5, max_row=ws.max_row))
    ch.height, ch.width = 18, 30; ch.y_axis.title = 'объявлений'
    ws.add_chart(ch, 'I5')

    # ---------------- 3. SEC Form D ----------------
    ws = wb.create_sheet('3. SEC Form D')
    ws.append(['Выгрузка SEC EDGAR Form D — доказательная база по структурам, суммам и чекам'])
    ws['A1'].font = TITLE_FONT
    ws.append(['Имена эмитентов очищены от совпадений по названию. Суммы — по формам с 2019 г. Не равны AUM.'])
    ws['A2'].font = NOTE_FONT
    ws.append([])
    ws.append(['Компания', 'Reg D-эмитентов', 'Форм D всего', 'Форм D с 2024',
               '% 506(c) — право рекламировать', 'Продано $млн', 'Медианный мин. чек $',
               'Инвесторов', 'Google-объявлений'])
    style_header(ws, row=4)
    fd = [
        ('Crow Holdings', 44, 84, 17, 0.0, 38235, 5000000, 2350, 0),
        ('MLG Capital', 45, 87, 22, 82.4, 8194, 50000, 4879, 22),
        ('Walton Global', 44, 118, 6, 0.0, 8276, 500, 516, 0),
        ('DLP Capital', 22, 51, 21, 73.8, 4206, 200000, 11550, 16),
        ('GTIS Partners [набл.]', 41, 88, 7, 15.0, 3634, 100000, 4000, 0),
        ('GSP REI', 57, 78, 38, 9.2, 1698, 25000, 1647, 0),
        ('Caliber', 29, 106, 39, 57.7, 1321, 20000, 4547, 24),
        ('Cardone Capital [бенч.]', 23, 28, 7, 100.0, 831, 100000, 4526, 69),
        ('NexMetro [набл.]', 13, 35, 16, 31.4, 470, 10000, 835, 0),
        ('FarmTogether', 44, 60, 10, 93.3, 260, 15000, 5891, 24),
        ('AcreTrader', 100, 101, 1, 100.0, 186, 17125, 7035, 3),
        ('Urban Catalyst', 5, 8, 0, 66.7, 13, 100000, 27, 2),
        ('RealtyMogul [бенч.]', 131, 212, 0, None, None, None, None, 29),
        ('EquityMultiple [бенч.]', 100, 102, 0, None, None, None, None, 7),
        ('Fundrise [бенч.]', 42, 118, 13, None, None, None, None, 200),
        ('Hines [набл.]', 33, 228, 61, None, None, None, None, 4),
        ('Belpointe OZ [набл.]', 7, 30, 4, None, None, None, None, 0),
    ]
    for r in fd: ws.append(list(r))
    for r in range(5, ws.max_row + 1):
        for col in (2, 3, 4, 6, 7, 8, 9): ws.cell(row=r, column=col).number_format = '#,##0'
        ws.cell(row=r, column=5).number_format = '0.0'
        p = ws.cell(row=r, column=5).value
        if p is not None:
            ws.cell(row=r, column=5).fill = PatternFill(
                'solid', fgColor='C6E0B4' if p >= 70 else ('FFE699' if p >= 20 else 'F8CBAD'))
    autosize(ws)

    sc = ScatterChart(); sc.title = 'Право рекламировать (% 506c) против фактической рекламы'
    sc.x_axis.title = '% предложений 506(c)'; sc.y_axis.title = 'активных объявлений Google'
    xref = Reference(ws, min_col=5, min_row=5, max_row=16)
    yref = Reference(ws, min_col=9, min_row=5, max_row=16)
    s = Series(yref, xref, title='компании'); s.marker.symbol = 'circle'; s.graphicalProperties.line.noFill = True
    sc.series.append(s); sc.height, sc.width = 12, 22
    ws.add_chart(sc, 'K5')

    # ---------------- 4. Модель бюджета ----------------
    sp = read_csv('spend_model.csv')
    ws = wb.create_sheet('4. Модель бюджета (оценка)')
    ws.append(['МОДЕЛЬНАЯ ОЦЕНКА бюджета привлечения капитала — НЕ факт'])
    ws['A1'].font = TITLE_FONT
    ws.append(['Бюджет = Годовой объём привлечения × Ставка затрат на привлечение. '
               'Ставка зависит от модели дистрибуции (бенчмарки индустрии). Точность — порядок величины.'])
    ws['A2'].font = NOTE_FONT
    ws.append([])
    ws.append(['#', 'Компания', 'Модель дистрибуции', 'Годовое привлечение $млн (оценка)',
               'Ставка от, %', 'Ставка до, %', 'Бюджет от $млн', 'Бюджет до $млн', 'Бюджет средн. $млн',
               'в т.ч. платное медиа $млн', 'Google-объявл. (факт)', '% 506(c) (факт)',
               'Инвесторов (факт)', 'Обоснование'])
    style_header(ws, row=4)
    for r in sp:
        ws.append([r['rank'], r['name'], r['dist_model'], float(r['annual_raise_usd_m_EST']),
                   float(r['acq_cost_rate_lo_pct']), float(r['acq_cost_rate_hi_pct']),
                   float(r['acq_budget_usd_m_lo_EST']), float(r['acq_budget_usd_m_hi_EST']),
                   float(r['acq_budget_usd_m_mid_EST']), float(r['paid_media_usd_m_mid_EST']),
                   r['google_ads_MEASURED'], r['pct_506c_MEASURED'], r['investors_MEASURED'], r['basis']])
    for r in range(5, ws.max_row + 1):
        for col in (4, 7, 8, 9, 10): ws.cell(row=r, column=col).number_format = '#,##0.00'
        ws.cell(row=r, column=14).alignment = Alignment(wrap_text=True, vertical='top')
    autosize(ws, maxw=34); ws.column_dimensions['N'].width = 70

    ch = BarChart(); ch.title = 'Оценка годового бюджета привлечения, $млн'
    ch.add_data(Reference(ws, min_col=9, min_row=4, max_row=ws.max_row), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, min_row=5, max_row=ws.max_row))
    ch.height, ch.width = 16, 28
    ws.add_chart(ch, 'P5')

    # ---------------- 5. Каналы ----------------
    mix = read_csv('channel_mix.csv')
    ws = wb.create_sheet('5. Бюджет по каналам')
    ws.append(['Распределение бюджета по каналам — модельная оценка на основе наблюдаемых сигналов'])
    ws['A1'].font = TITLE_FONT
    ws.append([])
    agg = {}
    for m in mix: agg[m['channel']] = agg.get(m['channel'], 0) + float(m['usd_m_mid_EST'])
    tot = sum(agg.values())
    ws.append(['СВОДКА ПО 20 КОМПАНИЯМ'])
    ws['A3'].font = Font(bold=True, size=11, color='1F3A5F')
    ws.append(['Канал', '$млн/год (оценка)', 'Доля бюджета, %'])
    style_header(ws, row=4, ncols=3)
    for ch_name, v in sorted(agg.items(), key=lambda x: -x[1]):
        ws.append([ch_name, round(v, 2), round(100 * v / tot, 1)])
    srow = ws.max_row
    ws.append(['ИТОГО', round(tot, 1), 100.0])
    for c in range(1, 4): ws.cell(row=ws.max_row, column=c).font = Font(bold=True)

    pie = PieChart(); pie.title = 'Куда идут деньги: совокупный бюджет 20 компаний'
    pie.add_data(Reference(ws, min_col=2, min_row=4, max_row=srow), titles_from_data=True)
    pie.set_categories(Reference(ws, min_col=1, min_row=5, max_row=srow))
    pie.height, pie.width = 12, 20
    ws.add_chart(pie, 'E4')

    ws.append([]); ws.append([])
    base = ws.max_row + 1
    ws.cell(row=base, column=1, value='ПО КОМПАНИЯМ').font = Font(bold=True, size=11, color='1F3A5F')
    ws.append(['Компания', 'Модель дистрибуции', 'Канал', 'Доля, %', '$млн/год (оценка)'])
    style_header(ws, row=base + 1, ncols=5)
    ws.freeze_panes = None
    for m in mix:
        ws.append([m['name'], m['dist_model'], m['channel'], float(m['pct']), float(m['usd_m_mid_EST'])])
    autosize(ws, maxw=48)

    # ---------------- 6. Доказательства ----------------
    ev = read_csv('evidence.csv')
    ws = wb.create_sheet('6. Доказательства')
    ws.append(['Компания', 'Показатель', 'Значение', 'Статус', 'Источник', 'Примечание'])
    style_header(ws)
    for e in ev:
        ws.append([e['company'], e['metric'], e['value'], e['status'], e['source'], e['note']])
    for r in range(2, ws.max_row + 1):
        st = ws.cell(row=r, column=4)
        st.fill = PatternFill('solid', fgColor={'FACT-M': 'C6E0B4', 'FACT-D': 'DDEBF7', 'EST': 'FFE699'}.get(st.value, 'FFFFFF'))
        ws.cell(row=r, column=6).alignment = Alignment(wrap_text=True, vertical='top')
    autosize(ws, maxw=46)
    ws.column_dimensions['E'].width = 62; ws.column_dimensions['F'].width = 56

    # ---------------- 7. Бенчмарк ----------------
    ws = wb.create_sheet('7. Бенчмарк-панель')
    ws.append(['У КОГО КОПИРОВАТЬ РЕКЛАМНУЮ МЕХАНИКУ'])
    ws['A1'].font = TITLE_FONT
    ws.append(['Это НЕ земельные игроки. Это те, кто реально умеет привлекать капитал платным трафиком.'])
    ws['A2'].font = NOTE_FONT
    ws.append([])
    ws.append(['Компания', 'Домен', 'Google-объявлений', 'YouTube подп.', 'Reg D-эмитентов', 'Форм D', 'Чему учиться'])
    style_header(ws, row=4)
    for b in read_csv('benchmark.csv'):
        ws.append([b['name'], b['domain'], int(b['google_ads'] or 0), b['yt_subs'] or None,
                   b['regd_entities'] or None, b['formd_count'] or None, b['note']])
    for r in range(5, ws.max_row + 1):
        for col in (3, 4, 5, 6): ws.cell(row=r, column=col).number_format = '#,##0'
        ws.cell(row=r, column=7).alignment = Alignment(wrap_text=True, vertical='top')
    autosize(ws, maxw=34); ws.column_dimensions['G'].width = 86

    add_channel_sheets(wb)
    add_marketing_sheets(wb)
    wb.save(OUTX)
    print('Сохранено:', OUTX)
    print('Листов:', len(wb.sheetnames), '->', ', '.join(wb.sheetnames))




def add_channel_sheets(wb):
    """Листы 8-10: каналы, языки/регионы, статус инструментов измерения."""
    import json as _j
    ch = _j.load(open(os.path.join(DATA, 'channels.json'), encoding='utf-8'))

    ws = wb.create_sheet('8. Каналы продвижения')
    ws.append(['КАНАЛЫ ПРОДВИЖЕНИЯ: что измерено по каждой компании'])
    ws['A1'].font = TITLE_FONT
    ws.append(['Google и Meta — количество активных объявлений. Соцсети — наличие канала по ссылкам с сайта. '
               'Суммы расходов не раскрывает ни одна из платформ.'])
    ws['A2'].font = NOTE_FONT
    ws.append([])
    rows = ch['matrix']
    cols = list(rows[0].keys())
    ws.append(cols)
    style_header(ws, row=4)
    for r in rows:
        ws.append([r[c] for c in cols])
    for r in range(5, ws.max_row + 1):
        g = ws.cell(row=r, column=2)
        gv = g.value if isinstance(g.value, (int, float)) else 0
        g.fill = PatternFill('solid', fgColor='F8CBAD' if gv == 0 else 'C6E0B4')
        rel = ws.cell(row=r, column=4)
        if isinstance(rel.value, str) and rel.value.startswith('НЕДОСТОВЕРНО'):
            rel.fill = PatternFill('solid', fgColor='FFE699')
            ws.cell(row=r, column=3).fill = PatternFill('solid', fgColor='FFE699')
        ws.cell(row=r, column=4).alignment = Alignment(wrap_text=True, vertical='top')
    autosize(ws, maxw=30)
    ws.column_dimensions['D'].width = 52

    ws = wb.create_sheet('9. Языки и регионы')
    ws.append(['ЯЗЫКИ И РЕГИОНЫ ПРОДВИЖЕНИЯ'])
    ws['A1'].font = TITLE_FONT
    ws.append([])
    ws.append(['Компания', 'Языки', 'Региональные порталы', 'Офисы', 'Регионы', 'Активность 2025-2026'])
    style_header(ws, row=3)
    for r in ch['lang_region']:
        ws.append([r['name'], r['langs'], r['portals'], r['offices'], r['regions'], r['activity_2026']])
    for r in range(4, ws.max_row + 1):
        for c in range(1, 7):
            ws.cell(row=r, column=c).alignment = Alignment(wrap_text=True, vertical='top')
    autosize(ws, maxw=30)
    for c, w in zip('BCDEF', (34, 30, 38, 30, 90)):
        ws.column_dimensions[c].width = w

    e = ch['eb5']
    ws.append([]); ws.append([])
    base = ws.max_row + 1
    ws.cell(row=base, column=1, value='ОТДЕЛЬНЫЙ КАНАЛ ИНОСТРАННОГО КАПИТАЛА: ' + e['channel']).font = Font(bold=True, size=11, color='1F3A5F')
    for k, lab in [('what', 'Что это'), ('why', 'Зачем смотреть'), ('geography', 'География'),
                   ('example', 'Примеры'), ('mechanics', 'Механика'), ('caveat', 'Оговорка')]:
        ws.append([lab, e[k]])
        ws.cell(row=ws.max_row, column=2).alignment = Alignment(wrap_text=True, vertical='top')

    ws = wb.create_sheet('10. Инструменты')
    ws.append(['СТАТУС ИНСТРУМЕНТОВ ИЗМЕРЕНИЯ — каждый проверен контрольным запросом'])
    ws['A1'].font = TITLE_FONT
    ws.append(['Контрольный запрос обязан вернуть много. Если инструмент возвращает ноль на Nike — '
               'его нули ничего не значат и в выводы не идут.'])
    ws['A2'].font = NOTE_FONT
    ws.append([])
    ws.append(['Источник', 'Статус', 'Контрольный запрос', 'Что покрывает', 'Примечание'])
    style_header(ws, row=4)
    for r in ch['instruments']:
        ws.append([r['source'], r['status'], r['control'], r['covers'], r['note']])
    for r in range(5, ws.max_row + 1):
        st = ws.cell(row=r, column=2)
        st.fill = PatternFill('solid', fgColor={'ВАЛИДЕН': 'C6E0B4'}.get(st.value, 'F8CBAD'))
        ws.cell(row=r, column=5).alignment = Alignment(wrap_text=True, vertical='top')
    autosize(ws, maxw=34)
    ws.column_dimensions['E'].width = 80


def add_marketing_sheets(wb):
    """Листы 11-14: разбор воронок, экономика крупного чека, конференции, города."""
    import json as _j
    fun = _j.load(open(os.path.join(DATA, 'funnels.json'), encoding='utf-8'))
    big = _j.load(open(os.path.join(DATA, 'bigticket.json'), encoding='utf-8'))
    metro = _j.load(open(os.path.join(DATA, 'metro.json'), encoding='utf-8'))

    # --- 11. Разбор воронок ---
    ws = wb.create_sheet('11. Разбор воронок')
    ws.append(['СЛЕПОК ВОРОНКИ: куда ведут трафик и чем цепляют'])
    ws['A1'].font = TITLE_FONT
    ws.append(['Пиксель на сайте = канал реально используется. Это точнее рекламных библиотек: '
               'библиотека может не показать объявления, но пиксель стоит только у того, кто льёт трафик.'])
    ws['A2'].font = NOTE_FONT
    ws.append([])
    cols = list(fun[0].keys())
    ws.append(cols); style_header(ws, row=4)
    for r in fun:
        ws.append([r[c] for c in cols])
    npaid = cols.index('Платных каналов, шт') + 1
    for r in range(5, ws.max_row + 1):
        c = ws.cell(row=r, column=npaid)
        v = c.value if isinstance(c.value, int) else 0
        c.fill = PatternFill('solid', fgColor='C6E0B4' if v >= 4 else ('FFE699' if v >= 1 else 'F8CBAD'))
        acc = ws.cell(row=r, column=cols.index('Спрашивают аккредитацию') + 1)
        if acc.value == 'ДА':
            acc.fill = PatternFill('solid', fgColor='DDEBF7')
        for cc in (3, 4, 5, 7):
            ws.cell(row=r, column=cc).alignment = Alignment(wrap_text=True, vertical='top')
    autosize(ws, maxw=26)
    for c, w in zip('CDEG', (40, 46, 44, 40)):
        ws.column_dimensions[c].width = w

    # --- 12. Экономика крупного чека ---
    ws = wb.create_sheet('12. Крупный чек')
    ws.append(['ЭКОНОМИКА КАНАЛА $1 МЛН+ — выбранный сегмент'])
    ws['A1'].font = TITLE_FONT
    ws.append(['Это каналы, которыми пользуются Walton, Crow, Domain и Bedrock. '
               'Ни один из них не виден в рекламных библиотеках, потому что это не реклама.'])
    ws['A2'].font = NOTE_FONT
    ws.append([])
    ws.append(['СЦЕНАРИЙ', 'Логика расчёта', 'Стоимость', 'Вердикт', 'Обоснование'])
    style_header(ws, row=4)
    for r in big['scenarios']:
        ws.append([r['scenario'], r['logic'], r['nominal_cost'], r['verdict'], r['why']])
    for r in range(5, ws.max_row + 1):
        v = ws.cell(row=r, column=4)
        v.fill = PatternFill('solid', fgColor='F8CBAD' if 'НЕ РАБОТАЕТ' in str(v.value) else 'C6E0B4')
        for cc in (2, 5):
            ws.cell(row=r, column=cc).alignment = Alignment(wrap_text=True, vertical='top')
    autosize(ws, maxw=30)
    ws.column_dimensions['B'].width = 56; ws.column_dimensions['E'].width = 88

    ws.append([]); ws.append([])
    base = ws.max_row + 1
    ws.cell(row=base, column=1, value='РАСЦЕНКИ PLACEMENT AGENTS').font = Font(bold=True, size=11, color='1F3A5F')
    ws.append(['Позиция', 'Значение', 'Источник']); style_header(ws, row=base + 1, ncols=3)
    for r in big['placement']:
        ws.append([r['item'], r['value'], r['src']])

    # --- 13. Конференции ---
    ws = wb.create_sheet('13. Конференции 2026')
    ws.append(['ГДЕ ФИЗИЧЕСКИ НАХОДЯТСЯ ДЕНЬГИ: профильные события 2026'])
    ws['A1'].font = TITLE_FONT
    ws.append(['Экономика этих событий устроена в вашу пользу наоборот: инвесторы от $100 млн проходят '
               'бесплатно, платит тот, кто ищет капитал. Организатор гарантирует нужную аудиторию в зале.'])
    ws['A2'].font = NOTE_FONT
    ws.append([])
    ws.append(['Событие', 'Когда', 'Где', 'Цена входа для управляющего', 'Комментарий'])
    style_header(ws, row=4)
    for r in big['conferences']:
        ws.append([r['event'], r['when'], r['where'], r['cost_manager'], r['note']])
    for r in range(5, ws.max_row + 1):
        ws.cell(row=r, column=5).alignment = Alignment(wrap_text=True, vertical='top')
    autosize(ws, maxw=34); ws.column_dimensions['E'].width = 64

    # --- 14. Города ---
    ws = wb.create_sheet('14. Города')
    ws.append(['РАЗРЕЗ ПО ГОРОДАМ: где сидят игроки двадцатки'])
    ws['A1'].font = TITLE_FONT
    ws.append(['Нью-Йорка в этом рынке нет: ни одна компания двадцатки там не базируется. '
               'Рынок земли сконцентрирован в Солнечном поясе, где есть пригодная к застройке земля.'])
    ws['A2'].font = NOTE_FONT
    ws.append([])
    ws.append(['Метрополия', 'Компаний', 'Сумма объявлений Google', 'Компании'])
    style_header(ws, row=4)
    for r in metro:
        ws.append([r['Метрополия'], r['Компаний'], r['Сумма объявлений Google'], r['Компании']])
    for r in range(5, ws.max_row + 1):
        ws.cell(row=r, column=4).alignment = Alignment(wrap_text=True, vertical='top')
    autosize(ws, maxw=30); ws.column_dimensions['D'].width = 62
    ch = BarChart(); ch.title = 'Компаний двадцатки по метрополиям'
    ch.add_data(Reference(ws, min_col=2, min_row=4, max_row=ws.max_row), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=5, max_row=ws.max_row))
    ch.height, ch.width = 14, 24
    ws.add_chart(ch, 'F5')


if __name__ == '__main__':
    main()
