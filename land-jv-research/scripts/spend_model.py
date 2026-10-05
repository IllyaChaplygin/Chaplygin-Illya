#!/usr/bin/env python3
"""
Модель распределения маркетинговых затрат по каналам.

ВАЖНО О СТАТУСЕ ЦИФР
--------------------
Рекламные расходы частных компаний в США не раскрываются. Ни Meta Ad Library,
ни Google Ads Transparency Center не показывают суммы для неполитической рекламы.
Поэтому все долларовые значения здесь — МОДЕЛЬНЫЕ ОЦЕНКИ, а не факты.

Что в модели ФАКТ (измерено):
  • наличие и количество активных объявлений Google (ATC)
  • доля предложений 506(c) — право на публичную рекламу (SEC Form D)
  • объём привлечённого капитала (SEC Form D / отчётность)
  • число инвесторов, медианный минимальный чек (SEC Form D)
  • объём собственного контента (YouTube, вебинары)

Что ОЦЕНКА:
  • процентное распределение бюджета по каналам
  • абсолютный бюджет в долларах

ЛОГИКА ОЦЕНКИ (привязка к реальному знаменателю)
------------------------------------------------
Вместо угадывания медиабюджета «из воздуха» бюджет привязан к стоимости
привлечения капитала — величине, по которой в индустрии есть бенчмарки:

  Годовой бюджет привлечения = Годовой объём привлечения × Ставка затрат на привлечение

Ставка зависит от модели дистрибуции (бенчмарки индустрии):
  • Дистрибьюторская модель (брокер-дилеры/RIA/агенты): 5–8% — почти всё
    уходит на комиссии посредникам, прямой медиабюджет < 0,5%
  • Прямая цифровая модель 506(c): 1,5–3,0% — основная часть на медиа и контент
  • Институциональная модель (LP, страховые): 0,3–1,0% — placement agents
    или вообще без них, маркетинга почти нет
  • Публичный рынок: 0,1–0,5% — IR, а не маркетинг

Бенчмарки стоимости лида (опубликованные, 2026):
  Meta CPL (accredited RE syndication) .......... $50–100
  Google Ads CPL (средний) ...................... $79
  LinkedIn CPL (financial services) ............. $90–180 (до $250–500 в синдикации)
  Подкаст-спонсорство CPL ....................... $60–200
  Стоимость одного ПРОФИНАНСИРОВАВШЕГО инвестора . $3 500–4 500
  Конверсия Meta-лида в инвестора за 90 дней .... ~2%
"""
import csv, json, os

HERE = os.path.dirname(__file__)
DATA = os.path.join(HERE, '..', 'data')

# Модели дистрибуции и ставка затрат на привлечение капитала (% от привлечения)
DIST_MODEL = {
    'distributor':   (0.050, 0.080, 'Дистрибьюторская сеть (BD/RIA/агенты)'),
    'direct_digital':(0.015, 0.030, 'Прямая цифровая воронка 506(c)'),
    'institutional': (0.003, 0.010, 'Институциональные LP / placement agents'),
    'public':        (0.001, 0.005, 'Публичный рынок (IR, не маркетинг)'),
}

# Распределение бюджета по каналам, % — выведено из НАБЛЮДАЕМЫХ сигналов
# (наличие рекламы, объём контента, события, структура дистрибуции)
CHANNEL_MIX = {
 'distributor': {
    'Комиссии дистрибьюторам и агентам': 72, 'События, роудшоу, конференции': 12,
    'Контент, печать, презентационные материалы': 7, 'PR и отраслевые медиа': 5,
    'Платный поиск и YouTube': 0, 'Платные соцсети': 1, 'Email и CRM': 3,
 },
 'direct_digital': {
    'Комиссии дистрибьюторам и агентам': 8, 'События, роудшоу, конференции': 10,
    'Контент, печать, презентационные материалы': 22, 'PR и отраслевые медиа': 6,
    'Платный поиск и YouTube': 24, 'Платные соцсети': 18, 'Email и CRM': 12,
 },
 'institutional': {
    'Комиссии дистрибьюторам и агентам': 45, 'События, роудшоу, конференции': 25,
    'Контент, печать, презентационные материалы': 12, 'PR и отраслевые медиа': 13,
    'Платный поиск и YouTube': 0, 'Платные соцсети': 2, 'Email и CRM': 3,
 },
 'public': {
    'Комиссии дистрибьюторам и агентам': 5, 'События, роудшоу, конференции': 22,
    'Контент, печать, презентационные материалы': 20, 'PR и отраслевые медиа': 35,
    'Платный поиск и YouTube': 8, 'Платные соцсети': 5, 'Email и CRM': 5,
 },
}

# name -> (модель дистрибуции, оценка годового привлечения $млн, обоснование)
PROFILE = {
 "Walton Global":                     ('distributor',   400, "$3,14 млрд распределено за историю; 89 000 инвесторов; 300+ каналов; 0% 506(c)"),
 "Millrose Properties (NYSE: MRP)":   ('public',       1500, "Публичный REIT, активы $9,7 млрд; капитал с рынка, не из рекламы"),
 "DLP Capital":                       ('direct_digital',500, "$4,21 млрд продано по формам D; 11 550 инвесторов; 74% 506(c); 16 объявлений"),
 "Caliber (NASDAQ: CWD)":             ('direct_digital',150, "$1,32 млрд продано; 4 547 инвесторов; 58% 506(c); 24 объявления"),
 "AcreTrader":                        ('direct_digital', 40, "$186 млн продано; 7 035 инвесторов; 100% 506(c), но лишь 3 объявления"),
 "MLG Capital":                       ('direct_digital',250, "Fund VI таргет $400 млн; 22 объявления; дистрибуция через advisors"),
 "Forestar Group (NYSE: FOR)":        ('public',        300, "Публичная, контроль D.R. Horton; 60 объявлений — B2B/HR, не инвесторы"),
 "Howard Hughes Holdings (NYSE: HHH)":('public',        500, "Публичная; потребительская реклама на уровне комьюнити"),
 "The St. Joe Company (NYSE: JOE)":   ('public',        200, "Публичная; JV с операторами"),
 "FarmTogether":                      ('direct_digital', 50, "$260 млн продано; 5 891 инвестор; 93% 506(c); 24 объявления"),
 "Domain Real Estate Partners":       ('institutional',1500, "$6 млрд развёрнуто; партнёрство с PGIM; 0 объявлений"),
 "BTI Partners":                      ('institutional', 200, "$3,6 млрд сделок; институциональные партнёры; 0 объявлений"),
 "GSP REI":                           ('direct_digital', 60, "57 Reg D-эмитентов, 38 форм с 2024; минимум $100K; 0 объявлений"),
 "Tejon Ranch Co (NYSE: TRC)":        ('public',         50, "Публичная; развитие через JV; 1 видео на YouTube"),
 "Allied Development":                ('direct_digital', 25, "30+ entitlement-проектов; воронка Book a Clarity Call; 5 объявлений"),
 "13th Floor Investments":            ('institutional', 300, "$5 млрд под управлением; 70+ сделок; JV с институционалами"),
 "Urban Catalyst":                    ('direct_digital', 30, "Fund I $131 млн / 356 инвесторов; 67% 506(c); 2 объявления"),
 "Five Point Holdings (NYSE: FPH)":   ('public',        150, "Публичная; JV с домостроителями; Hearthstone"),
 "Crow Holdings":                     ('institutional', 800, "44 Reg D-эмитента; AUM ~$33 млрд; институциональные LP"),
 "Bedrock Land Finance (TWG Global)": ('institutional', 800, "Цель $5 млрд; эксклюзивный партнёр Guggenheim"),
}


def main():
    comp = json.load(open(os.path.join(DATA, 'companies.json')))['companies']
    by_name = {c['name']: c for c in comp}

    rows, mix_rows = [], []
    for name, (model, annual_raise_m, basis) in PROFILE.items():
        lo_rate, hi_rate, model_label = DIST_MODEL[model]
        lo = annual_raise_m * lo_rate
        hi = annual_raise_m * hi_rate
        mid = (lo + hi) / 2
        c = by_name.get(name, {})
        mix = CHANNEL_MIX[model]

        rows.append({
            'rank': c.get('rank', ''), 'name': name, 'dist_model': model_label,
            'annual_raise_usd_m_EST': annual_raise_m,
            'acq_cost_rate_lo_pct': round(lo_rate*100, 2),
            'acq_cost_rate_hi_pct': round(hi_rate*100, 2),
            'acq_budget_usd_m_lo_EST': round(lo, 2),
            'acq_budget_usd_m_hi_EST': round(hi, 2),
            'acq_budget_usd_m_mid_EST': round(mid, 2),
            'paid_media_usd_m_mid_EST': round(mid * (mix['Платный поиск и YouTube'] + mix['Платные соцсети'])/100, 3),
            'google_ads_MEASURED': c.get('google_ads', ''),
            'pct_506c_MEASURED': c.get('pct_506c', ''),
            'investors_MEASURED': c.get('investors', ''),
            'basis': basis,
        })
        for ch, pct in mix.items():
            mix_rows.append({'name': name, 'dist_model': model_label, 'channel': ch,
                             'pct': pct, 'usd_m_mid_EST': round(mid*pct/100, 3)})

    rows.sort(key=lambda r: -r['acq_budget_usd_m_mid_EST'])
    with open(os.path.join(DATA, 'spend_model.csv'), 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys())); w.writeheader(); w.writerows(rows)
    with open(os.path.join(DATA, 'channel_mix.csv'), 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=list(mix_rows[0].keys())); w.writeheader(); w.writerows(mix_rows)
    json.dump({'spend': rows, 'mix': mix_rows, 'dist_model': {k: v[2] for k, v in DIST_MODEL.items()},
               'channel_mix': CHANNEL_MIX},
              open(os.path.join(DATA, 'spend_model.json'), 'w'), ensure_ascii=False, indent=1)

    print(f"{'Компания':36}{'модель':30}{'привл.$млн':>11}{'бюджет$млн':>12}{'медиа$млн':>11}{'Google':>7}")
    print('-'*110)
    for r in rows:
        print(f"{r['name'][:35]:36}{r['dist_model'][:29]:30}{r['annual_raise_usd_m_EST']:>11}"
              f"{r['acq_budget_usd_m_mid_EST']:>12.1f}{r['paid_media_usd_m_mid_EST']:>11.2f}{str(r['google_ads_MEASURED']):>7}")

    agg = {}
    for m in mix_rows:
        agg[m['channel']] = agg.get(m['channel'], 0) + m['usd_m_mid_EST']
    tot = sum(agg.values())
    print(f"\nСОВОКУПНЫЙ БЮДЖЕТ 20 КОМПАНИЙ (оценка): ${tot:,.0f} млн/год")
    print(f"{'Канал':46}{'$млн':>10}{'доля':>8}")
    print('-'*64)
    for ch, v in sorted(agg.items(), key=lambda x: -x[1]):
        print(f"{ch:46}{v:>10.1f}{100*v/tot:>7.1f}%")


if __name__ == '__main__':
    main()
