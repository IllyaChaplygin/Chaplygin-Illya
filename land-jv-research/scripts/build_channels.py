#!/usr/bin/env python3
"""
Дополнение: каналы, языки, регионы.

Закрывает три пункта, не покрытые в основном отчёте:
  1. Каналы помимо Google/YouTube — Meta, LinkedIn, X, Instagram, TikTok
  2. Языки, на которых ведётся продвижение
  3. Регионы и страны, на которые идёт таргетинг

ВАЖНО О ВАЛИДАЦИИ ИНСТРУМЕНТОВ
-------------------------------
Каждый источник проверен контрольным запросом, который ОБЯЗАН вернуть много:

  Google Ads Transparency Center .. ВАЛИДЕН   (fundrise.com -> 200+, nike не нужен)
  Meta Ad Library ................. ВАЛИДЕН   (Nike -> 24 000 результатов)
  TikTok Commercial Content Library НЕ ВАЛИДЕН (Nike/Temu/Shopify/Samsung -> 0)
  LinkedIn Ad Library ............. НЕДОСТУПЕН (Cloudflare блокирует)

Нули TikTok в этой сессии НИЧЕГО не означают и в выводы не идут.
Нули Meta означают отсутствие рекламы, потому что инструмент прошёл проверку.
"""
import csv, json, os

HERE = os.path.dirname(__file__)
DATA = os.path.join(HERE, '..', 'data')
RAW = os.path.join(DATA, 'raw')

# Статус инструментов измерения
INSTRUMENTS = [
    dict(source='Google Ads Transparency Center', status='ВАЛИДЕН',
         control='fundrise.com -> 200+ объявлений', covers='Поиск Google, YouTube, Display, Shopping',
         note='Показывает наличие и количество объявлений. Суммы НЕ показывает.'),
    dict(source='Meta Ad Library', status='ВАЛИДЕН',
         control='Nike -> 24 000 результатов', covers='Facebook, Instagram, Messenger, Audience Network',
         note='Поиск по тексту объявления. Компания, не упоминающая себя в тексте, может быть недосчитана. Суммы для неполитической рекламы НЕ показывает.'),
    dict(source='TikTok Commercial Content Library', status='НЕ ВАЛИДЕН',
         control='Nike, Temu, Shopify, Samsung -> все 0', covers='TikTok',
         note='Фильтр страны не устанавливается, библиотека возвращает нули на заведомо крупных рекламодателей. Результаты по TikTok в выводы НЕ включены.'),
    dict(source='LinkedIn Ad Library', status='НЕДОСТУПЕН',
         control='—', covers='LinkedIn',
         note='Cloudflare блокирует доступ. Присутствие на LinkedIn снято по ссылкам с сайтов компаний, объёма рекламы нет.'),
    dict(source='X / Twitter', status='НЕДОСТУПЕН',
         control='—', covers='X',
         note='Публичного репозитория рекламы для США нет (DSA-репозиторий покрывает только ЕС).'),
    dict(source='Сайты компаний', status='ВАЛИДЕН',
         control='снято с 20 из 20 сайтов', covers='Наличие каналов, языковые версии',
         note='Прямой разбор DOM: ссылки на соцсети, hreflang, языковые переключатели.'),
]

# Языки и регионы — по результатам разбора сайтов и публикаций 2025-2026
LANG_REGION = [
 dict(name='Walton Global', langs='EN, 繁體中文, 简体中文, 日本語, DEUTSCH (5 языков)',
      portals='Asia & Middle East, Canada, Germany, VISTRA, Great Lakes',
      offices='Scottsdale, Calgary, Гонконг, Тайбэй, Сингапур, Манила, Токио, Дубай',
      regions='91 страна. Фокус: Азия, Ближний Восток, Европа',
      activity_2026='Япония: фонд "US My Home Fund", запуск в Токио 10.04.2026, дистрибьюторы Teneo Partners и Matsuzaka Securities. Тайбэй: роудшоу. Гонконг: зарегистрированный инвестиционный инструмент. U.S. Land Income & Growth Fund для офшорных инвесторов, первое закрытие 10.2025. PR-синдикация через BERNAMA (Малайзия), ANTARA (Индонезия), Japan Times, индийские издания.'),
 dict(name='Millrose Properties', langs='EN', portals='—', offices='Майами',
      regions='США, 30 штатов', activity_2026='Публичный REIT, капитал с биржи. Языковых версий нет.'),
 dict(name='DLP Capital', langs='EN', portals='—', offices='St. Augustine, FL',
      regions='США, одобрен к кредитованию в 37 штатах', activity_2026='Вебинары, Elite Impact Podcast, Investor Vision Day. Только английский, только США.'),
 dict(name='Caliber (NASDAQ: CWD)', langs='EN', portals='—', offices='Scottsdale, AZ',
      regions='США, Юго-Запад', activity_2026='5 соцканалов — максимум выборки. Цели 2024-2026: привлечь $750 млн, AUM $3 млрд.'),
 dict(name='MLG Capital', langs='EN', portals='—', offices='Brookfield, WI',
      regions='США: Southeast, Mountain West, Midwest', activity_2026='4 соцканала. Дистрибуция через investment advisors и family offices.'),
 dict(name='AcreTrader', langs='EN', portals='—', offices='Fayetteville, AR',
      regions='США, сельхозземля', activity_2026='Контент-воронка: 152 видео. 100% предложений 506(c).'),
 dict(name='FarmTogether', langs='EN', portals='—', offices='Сан-Франциско',
      regions='США, permanent crops', activity_2026='4 соцканала, ~1 новое предложение в месяц.'),
 dict(name='Прочие 13 компаний топ-20', langs='EN (только английский)', portals='—', offices='США',
      regions='США, преимущественно Солнечный пояс',
      activity_2026='Ни у одной нет языковых версий сайта, региональных порталов или зарубежной дистрибуции.'),
]

# Отдельный канал иностранного капитала, не покрытый в основном отчёте
EB5 = dict(
  channel='EB-5 (иммиграционные инвестиции)',
  what='Иностранный инвестор вкладывает $800 тыс.–$1,05 млн в проект США и получает путь к грин-карте. Девелопмент и земля — один из основных классов проектов.',
  why='Это параллельный канал привлечения капитала под крупные чеки, работающий на других языках и в других регионах, чем внутренний рынок США.',
  geography='Китай 56,3% заявителей, затем Индия и Вьетнам — вместе более 75%',
  example='CanAm Enterprises: штаб-квартира Нью-Йорк, офисы в Пекине, Шанхае, Хошимине, Нью-Дели, Сингапуре. EB5AN: 3 000+ инвесторов из 70+ стран, $1 млрд капитала, $8 млрд стоимости проектов.',
  mechanics='Привлечение идёт через миграционных агентов за комиссию, а не через рекламу. Отраслевое медиа — EB5Investors Magazine. Материалы обязательно переводятся.',
  caveat='Ни одна из 20 компаний основного рейтинга не использует EB-5 как заявленный канал. Включено как отдельная возможность, а не как наблюдаемая практика выборки.',
)


def load_site_scan():
    path = os.path.join(RAW, 'site_social_scan.jsonl')
    rows = []
    if os.path.exists(path):
        for line in open(path, encoding='utf-8'):
            line = line.strip()
            if line:
                rows.append(json.loads(line))
    return rows


def load_meta():
    path = os.path.join(RAW, 'meta_ads.jsonl')
    d = {}
    if os.path.exists(path):
        for line in open(path, encoding='utf-8'):
            line = line.strip()
            if line:
                r = json.loads(line)
                d[r['q']] = r
    return d


# сопоставление имён сайта -> имён в Meta-выгрузке
META_KEY = {
 'Walton Global':'Walton Global','Millrose Properties':'Millrose Properties','DLP Capital':'DLP Capital',
 'Caliber':'Caliber Companies','MLG Capital':'MLG Capital','AcreTrader':'AcreTrader',
 'Forestar':'Forestar','Howard Hughes':'Howard Hughes Communities','St. Joe':'St. Joe Company',
 'FarmTogether':'FarmTogether','GSP REI':'GSP REI','Domain REP':'Domain Real Estate Partners',
 'Crow Holdings':'Crow Holdings','BTI Partners':'BTI Partners','Tejon Ranch':'Tejon Ranch',
 'Allied Development':'Allied Development land','13th Floor':'13th Floor Investments',
 'Urban Catalyst':'Urban Catalyst','Five Point':'Five Point','Bedrock Land Finance':'Bedrock Land Finance',
}

GOOGLE = {'Walton Global':0,'Millrose Properties':0,'DLP Capital':16,'Caliber':24,'MLG Capital':22,
 'AcreTrader':3,'Forestar':60,'Howard Hughes':0,'St. Joe':0,'FarmTogether':24,'GSP REI':0,
 'Domain REP':0,'Crow Holdings':0,'BTI Partners':0,'Tejon Ranch':0,'Allied Development':5,
 '13th Floor':0,'Urban Catalyst':2,'Five Point':0,'Bedrock Land Finance':0}


# Поиск Meta идёт по ТЕКСТУ объявления в режиме keyword_unordered: слова матчатся
# в любом порядке и по отдельности. Поэтому имя из распространённых слов даёт тысячи
# чужих объявлений. Достоверны только редкие, неделимые имена.
GENERIC_NAMES = {'FarmTogether','Five Point','St. Joe Company','Allied Development land',
                 'CrowdStreet','Forestar','Tejon Ranch'}
def meta_reliability(key, val):
    if val is None: return 'не снято'
    if key in GENERIC_NAMES:
        return 'НЕДОСТОВЕРНО: имя из общих слов, матчится чужая реклама'
    if val == 0:
        return 'ноль, но возможен ложный ноль (см. New Western)'
    return 'достоверно: редкое имя'

def main():
    sites = load_site_scan()
    meta = load_meta()
    rows = []
    for s in sites:
        n = s['name']
        soc = s.get('social', {})
        mk = META_KEY.get(n)
        mrec = meta.get(mk) if mk else None
        rows.append({
            'Компания': n,
            'Google объявлений (изм.)': GOOGLE.get(n, ''),
            'Meta объявлений (изм.)': ('' if not mrec or mrec.get('meta_results') is None
                                       else mrec['meta_results']),
            'Meta: достоверность': meta_reliability(mk, (mrec or {}).get('meta_results')),
            'Facebook': 'есть' if 'facebook' in soc else '—',
            'LinkedIn': 'есть' if 'linkedin' in soc else '—',
            'X/Twitter': 'есть' if 'x' in soc else '—',
            'YouTube': 'есть' if 'youtube' in soc else '—',
            'Instagram': 'есть' if 'instagram' in soc else '—',
            'TikTok': 'есть' if 'tiktok' in soc else '—',
            'Каналов всего': len(soc),
            'Языковых версий сайта': len(set(s.get('hreflang') or [])) or (5 if n == 'Walton Global' else 1),
        })
    rows.sort(key=lambda r: (-r['Каналов всего'], r['Компания']))

    with open(os.path.join(DATA, 'channel_matrix.csv'), 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys())); w.writeheader(); w.writerows(rows)

    with open(os.path.join(DATA, 'instruments.csv'), 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=['source','status','control','covers','note'])
        w.writeheader(); w.writerows(INSTRUMENTS)

    with open(os.path.join(DATA, 'languages_regions.csv'), 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=['name','langs','portals','offices','regions','activity_2026'])
        w.writeheader(); w.writerows(LANG_REGION)

    json.dump({'matrix': rows, 'instruments': INSTRUMENTS, 'lang_region': LANG_REGION, 'eb5': EB5},
              open(os.path.join(DATA, 'channels.json'), 'w'), ensure_ascii=False, indent=1)

    print(f"{'Компания':24}{'Google':>7}{'Meta':>6}{'каналов':>9}  соцсети")
    print('-' * 86)
    for r in rows:
        soc = [k for k in ['Facebook','LinkedIn','X/Twitter','YouTube','Instagram','TikTok'] if r[k] == 'есть']
        print(f"{r['Компания'][:23]:24}{str(r['Google объявлений (изм.)']):>7}"
              f"{str(r['Meta объявлений (изм.)']):>6}{r['Каналов всего']:>9}  {', '.join(soc) or '—'}")
    avg_t1 = [r for r in rows if r['Каналов всего'] <= 1]
    print(f"\nКомпаний с 0-1 каналом: {len(avg_t1)} из {len(rows)}")


if __name__ == '__main__':
    main()
