#!/usr/bin/env python3
"""
Экономика канала КРУПНОГО ЧЕКА ($1 млн и выше).

Это тот сегмент, который выбрал заказчик. Здесь собраны опубликованные расценки
на те каналы, которыми реально пользуются Walton, Crow Holdings, Domain и Bedrock —
и которые НЕ видны ни в одной рекламной библиотеке, потому что это не реклама.
"""
import csv, json, os

DATA = os.path.join(os.path.dirname(__file__), '..', 'data')

# Канал 1. Placement agents / размещающие агенты
PLACEMENT = [
 dict(item='Success fee, фонд до $100 млн', value='2,0–2,5% от привлечённого', src='pipelineroad.com/blog/placement-agent-fees-2026'),
 dict(item='Success fee, фонд $100–500 млн', value='1,5–2,0%', src='pipelineroad.com/blog/placement-agent-fees-2026'),
 dict(item='Success fee, фонд свыше $500 млн', value='1,0–1,5%', src='pipelineroad.com/blog/placement-agent-fees-2026'),
 dict(item='Success fee, недвижимость (отраслевой диапазон)', value='1–3% от привлечённого эквити', src='selltohomepros.com/blog/real-estate-fund-placement-explained'),
 dict(item='Ретейнер разовый', value='$25 000 – $100 000+', src='pipelineroad.com/blog/placement-agent-fees-2026'),
 dict(item='Ретейнер ежемесячный', value='$10 000 – $25 000/мес', src='lpbacked.com/resources/placement-agent-fees'),
 dict(item='Период эксклюзивности', value='6–18 месяцев', src='lpbacked.com/resources/placement-agent-fees'),
 dict(item='Требование к агенту', value='Регистрация FINRA, Series 7/79', src='selltohomepros.com/blog/real-estate-fund-placement-explained'),
 dict(item='ПОЛНАЯ стоимость привлечения (все издержки)', value='4–8% от привлечённого эквити', src='accreditedinvestorleadgeneration.com/capital-raising/real-cost-raising-capital-2026'),
 dict(item='Пример: фонд $250 млн при ставке 1,75%', value='$4 375 000 только агенту', src='pipelineroad.com/blog/placement-agent-fees-2026'),
]

# Канал 2. Конференции — физический доступ к family offices и институционалам
CONFERENCES = [
 dict(event='Opal Family Office Winter Forum', when='10.03.2026', where='New York Marriott Marquis',
      cost_manager='$2 895', note='Семейные офисы от $100 млн проходят бесплатно — платит тот, кто ищет деньги'),
 dict(event='Opal Family Office & Private Wealth Management Forum', when='27–29.07.2026', where='Newport, RI',
      cost_manager='$3 195 early / $3 495 standard', note='Три дня, основной летний слёт'),
 dict(event='Opal Family Office & Private Wealth Legacy Summit', when='08.2026', where='США',
      cost_manager='$3 195', note='Верифицированные family offices от $100 млн — бесплатно'),
 dict(event='Markets Group Private Wealth US Spring Retreat', when='11–13.05.2026', where='Carlsbad, CA',
      cost_manager='$7 500', note='Самый дорогой вход в выборке; квалифицированные инвесторы бесплатно'),
 dict(event='IMN Real Estate Family Office & Private Wealth East', when='26–27.10.2026', where='Loews Coral Gables, FL',
      cost_manager='уточняется у организатора', note='Профильный для недвижимости, Майами'),
 dict(event='IMN Real Estate Private Funds', when='24–26.06.2026', where='Newport, RI',
      cost_manager='уточняется у организатора', note='Фонды недвижимости и их LP'),
 dict(event='IMN Land & Homebuilding Capital Markets', when='01–02.06.2026', where='США',
      cost_manager='уточняется у организатора', note='ЕДИНСТВЕННОЕ профильное событие именно по земле и жилью'),
 dict(event='Opal Real Estate Investment Summit', when='04.2026', where='Флорида',
      cost_manager='уточняется у организатора', note='Недвижимость, Флорида'),
]

# Сравнение экономики: цифровой канал против канала крупного чека
# Допущение для расчёта: цель привлечь $20 млн
TARGET_RAISE = 20_000_000
SCENARIOS = [
 dict(scenario='Платный цифровой трафик под чек $1 млн+',
      logic='CPL аккредитованного инвестора $50–500; стоимость профинансировавшего инвестора $3 500–4 500. '
            'Для $20 млн нужно ~20 инвесторов по $1 млн.',
      nominal_cost='~$90 000 (20 × $4 500)',
      verdict='НЕ РАБОТАЕТ',
      why='Бенчмарк $3 500–4 500 получен на чеках $100–200 тыс. Конверсия холодного цифрового лида '
          'в чек $1 млн близка к нулю: такие решения не принимаются по клику. '
          'Доказательство: у Walton ($8,28 млрд) и Crow ($38,2 млрд, медианный чек $5 млн) — '
          'ноль объявлений и ни одного рекламного пикселя на сайте.'),
 dict(scenario='Placement agent',
      logic='1,5–2,5% success fee плюс ретейнер $10–25 тыс./мес на 6–18 мес.',
      nominal_cost='$500 000 (2,5%) + ретейнер $60–300 тыс.',
      verdict='РАБОТАЕТ, но дорого',
      why='Платите за готовый доступ к LP. Агент обязан быть зарегистрирован в FINRA. '
          'Полная стоимость привлечения со всеми издержками — 4–8% от эквити.'),
 dict(scenario='Собственная сеть дистрибьюторов',
      logic='Модель Walton: 300+ одобренных каналов — брокер-дилеры, RIA, family offices.',
      nominal_cost='5–8% от привлечённого, но платится по факту',
      verdict='РАБОТАЕТ, масштабируется',
      why='Дороже агента в процентах, но строит собственный актив, а не арендует чужой. '
          'Walton собрал так 89 000 инвесторов из 91 страны.'),
 dict(scenario='Конференц-цикл',
      logic='10 профильных событий в год по $3–7,5 тыс. за вход плюс перелёты и представительские.',
      nominal_cost='$30–75 тыс. вход + $50–100 тыс. сопутствующие',
      verdict='РАБОТАЕТ как вход в канал',
      why='Самый дешёвый способ физически попасть к family offices. Экономика событий устроена так, '
          'что инвесторы от $100 млн проходят бесплатно, а платит ищущий деньги — '
          'то есть организатор гарантирует вам нужную аудиторию в зале.'),
]


def main():
    with open(os.path.join(DATA, 'bigticket_placement.csv'), 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=['item', 'value', 'src']); w.writeheader(); w.writerows(PLACEMENT)
    with open(os.path.join(DATA, 'bigticket_conferences.csv'), 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=['event', 'when', 'where', 'cost_manager', 'note'])
        w.writeheader(); w.writerows(CONFERENCES)
    with open(os.path.join(DATA, 'bigticket_scenarios.csv'), 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=['scenario', 'logic', 'nominal_cost', 'verdict', 'why'])
        w.writeheader(); w.writerows(SCENARIOS)
    json.dump({'placement': PLACEMENT, 'conferences': CONFERENCES, 'scenarios': SCENARIOS,
               'target_raise': TARGET_RAISE},
              open(os.path.join(DATA, 'bigticket.json'), 'w'), ensure_ascii=False, indent=1)
    print('Канал крупного чека:')
    print(f'  расценок placement agents: {len(PLACEMENT)}')
    print(f'  профильных конференций 2026: {len(CONFERENCES)}')
    print(f'  сценариев сравнения: {len(SCENARIOS)}')
    for s in SCENARIOS:
        print(f"    {s['verdict']:22} {s['scenario']}  ->  {s['nominal_cost']}")


if __name__ == '__main__':
    main()
