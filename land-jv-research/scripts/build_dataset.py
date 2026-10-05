#!/usr/bin/env python3
"""
Сборка основного датасета по 20 компаниям + панель-бенчмарк.

Каждое поле помечено уровнем достоверности:
  M = measured   — прямое измерение (Google ATC, SEC EDGAR, YouTube)
  D = disclosed  — раскрыто компанией/в отчётности
  E = estimated  — модельная оценка
  N = no data

Запуск:  python3 scripts/build_dataset.py
Выход:   data/companies.csv, data/companies.json, data/scores.csv
"""
import csv, json, os

OUT = os.path.join(os.path.dirname(__file__), '..', 'data')
os.makedirs(OUT, exist_ok=True)

# ---------------------------------------------------------------------------
# ДАННЫЕ
# acres           — акры/homesites под контролем (D)
# projects        — число отдельных земельных объектов/проектов (M/D)
# markets         — число региональных рынков (D)
# aum_usd_m       — AUM / land under management, $млн (D)
# raised_usd_m    — накопленный привлечённый капитал или распределения, $млн (D/M)
# regd_entities   — число юрлиц-эмитентов Reg D в SEC EDGAR (M)
# formd_count     — число поданных форм D (M)
# pct_506c        — доля предложений 506(c), % (M) — право на публичную рекламу
# google_ads      — активных объявлений в Google Ads Transparency Center (M)
# yt_subs/yt_vids — YouTube (M)
# min_check_usd   — медианный минимальный чек по Form D (M) или заявленный (D)
# ---------------------------------------------------------------------------

COMPANIES = [
 # ---------------- TIER 1: институциональные земельные банкиры ----------------
 dict(tier=1, name="Walton Global", hq="Scottsdale, AZ", domain="walton.com",
      model="Land banking: покупка pre-development земли, синдикация долей, продажа застройщикам",
      acres=80500, projects=198, markets=34, aum_usd_m=4640, raised_usd_m=3140,
      regd_entities=44, formd_count=118, pct_506c=0.0, formd_sold_usd_m=8276,
      google_ads=0, yt_subs=None, yt_vids=None, min_check_usd=500,
      investors=89000, countries=91, dist_channels=300,
      channel_model="300+ одобренных каналов дистрибуции; брокер-дилеры, RIA, family offices; офисы Гонконг/Сингапур/Тайбэй/Токио/Манила/Дубай; Global Distributor Meet",
      note="Крупнейший частный land asset manager мира. 52/52 предложений — 506(b): публичная реклама ЗАПРЕЩЕНА юридически. $8,28 млрд продано по формам D."),

 dict(tier=1, name="Millrose Properties (NYSE: MRP)", hq="Miami, FL", domain="millroseproperties.com",
      model="Публичный land banking REIT: Homesite Option Purchase Platform",
      acres=143771, projects=877, markets=30, aum_usd_m=9700, raised_usd_m=16300,
      regd_entities=None, formd_count=None, pct_506c=None, formd_sold_usd_m=None,
      google_ads=0, yt_subs=None, yt_vids=None, min_check_usd=None,
      investors=None, countries=1, dist_channels=19,
      channel_model="Публичный рынок акций + IR; внешнее управление Kennedy Lewis; 19 контрагентов-застройщиков",
      note="143 771 homesites в 877 комьюнити в 30 штатах. Активы $9,7 млрд, ожидаемые takedown-поступления $16,3 млрд. Крупнейший земельный банк выборки."),

 dict(tier=1, name="Domain Real Estate Partners", hq="Chicago, IL", domain="domainrealestatepartners.com",
      model="Land banking / lot option financing + project-level equity для застройщиков",
      acres=None, projects=720, markets=None, aum_usd_m=6000, raised_usd_m=16000,
      regd_entities=None, formd_count=None, pct_506c=None, formd_sold_usd_m=None,
      google_ads=0, yt_subs=None, yt_vids=None, min_check_usd=None,
      investors=None, countries=1, dist_channels=None,
      channel_model="Институциональные LP: PGIM ($4 млрд совместно), страховые балансы. Клиенты — Lennar, PulteGroup, Toll Brothers",
      note="720+ проектов, ~$16 млрд совокупной стоимости, $6 млрд развёрнуто. Основан в 2015. Розничной рекламы нет."),

 dict(tier=1, name="Howard Hughes Holdings (NYSE: HHH)", hq="The Woodlands, TX", domain="howardhughes.com",
      model="Master-planned communities: владение землёй, продажа лотов, JV по коммерческим активам",
      acres=118000, projects=5, markets=5, aum_usd_m=None, raised_usd_m=None,
      regd_entities=None, formd_count=None, pct_506c=None, formd_sold_usd_m=None,
      google_ads=0, yt_subs=None, yt_vids=None, min_check_usd=None,
      investors=None, countries=1, dist_channels=None,
      channel_model="Публичный рынок + JV с институциональными партнёрами; потребительская реклама на уровне комьюнити",
      note="118 000+ акров в 5 штатах: Teravalis 37 000, Summerlin 22 500, Bridgeland 11 500."),

 dict(tier=1, name="The St. Joe Company (NYSE: JOE)", hq="Panama City Beach, FL", domain="joe.com",
      model="Крупнейший частный землевладелец Флориды; монетизация через JV с операторами",
      acres=165000, projects=12, markets=1, aum_usd_m=None, raised_usd_m=None,
      regd_entities=None, formd_count=None, pct_506c=None, formd_sold_usd_m=None,
      google_ads=0, yt_subs=1110, yt_vids=228, min_check_usd=None,
      investors=None, countries=1, dist_channels=None,
      channel_model="JV с профильными операторами (Minto/Latitude Margaritaville, Key International, InterMountain); публичный рынок",
      note="~165 000 акров, 90% в 15 милях от Мексиканского залива. JV-выручка $56,1 млн за Q1 2026. 12 отелей."),

 dict(tier=1, name="Tejon Ranch Co (NYSE: TRC)", hq="Lebec, CA", domain="tejonranch.com",
      model="Крупнейший единый земельный массив Калифорнии; развитие ТОЛЬКО через JV",
      acres=270000, projects=4, markets=1, aum_usd_m=None, raised_usd_m=None,
      regd_entities=None, formd_count=None, pct_506c=None, formd_sold_usd_m=None,
      google_ads=0, yt_subs=5, yt_vids=1, min_check_usd=None,
      investors=None, countries=1, dist_channels=None,
      channel_model="Прямые JV с девелоперами (Majestic Realty, Dedeaux Properties, 3 домостроителя); публичный рынок",
      note="270 000 акров — эталон модели 'не продаём, а участвуем'. YouTube: 5 подписчиков, 1 видео за всю историю."),

 dict(tier=1, name="Five Point Holdings (NYSE: FPH)", hq="Irvine, CA", domain="fivepoint.com",
      model="MPC-девелопер Калифорнии + контроль в land banking venture Hearthstone",
      acres=40000, projects=3, markets=3, aum_usd_m=None, raised_usd_m=None,
      regd_entities=4, formd_count=4, pct_506c=None, formd_sold_usd_m=None,
      google_ads=0, yt_subs=None, yt_vids=None, min_check_usd=None,
      investors=None, countries=1, dist_channels=None,
      channel_model="Публичный рынок + JV с домостроителями; с 31.07.2025 контроль в Hearthstone Residential Holdings",
      note="Valencia 21 500 homesites, Great Park 2 100 акров / 11 856 домов, Candlestick 280 акров / 7 200 домов."),

 dict(tier=1, name="Forestar Group (NYSE: FOR)", hq="Arlington, TX", domain="forestar.com",
      model="Крупнейший независимый девелопер жилых лотов: земля → горизонталь → продажа лотов",
      acres=91700, projects=None, markets=65, aum_usd_m=None, raised_usd_m=None,
      regd_entities=None, formd_count=None, pct_506c=None, formd_sold_usd_m=None,
      google_ads=60, yt_subs=None, yt_vids=None, min_check_usd=None,
      investors=None, countries=1, dist_channels=None,
      channel_model="Публичный рынок (контроль D.R. Horton). 60 объявлений Google — B2B/HR, не привлечение инвесторов",
      note="91 700 лотов (62 200 собственных + 29 500 контролируемых), 65 рынков в 24 штатах."),

 dict(tier=1, name="Crow Holdings", hq="Dallas, TX", domain="crowholdings.com",
      model="Девелопмент-фонды: земля под индустриальные и жилые проекты, институциональный капитал",
      acres=None, projects=None, markets=None, aum_usd_m=33000, raised_usd_m=38235,
      regd_entities=44, formd_count=84, pct_506c=0.0, formd_sold_usd_m=38235,
      google_ads=0, yt_subs=None, yt_vids=None, min_check_usd=5000000,
      investors=2350, countries=1, dist_channels=None,
      channel_model="Институциональные LP, фонды Development Opportunities / Industrial Development; без публичной рекламы",
      note="САМЫЕ КРУПНЫЕ ЧЕКИ ВЫБОРКИ: медианный минимальный чек $5 000 000. $38,2 млрд продано по формам D, 2 350 инвесторов, 54 из 54 предложений — 506(b). Эталон искомой модели — и при этом ноль рекламы."),

 dict(tier=1, name="BTI Partners", hq="Fort Lauderdale, FL", domain="btipartners.com",
      model="Land developer / land investor Флориды: MPC и waterfront",
      acres=12000, projects=6, markets=4, aum_usd_m=3600, raised_usd_m=650,
      regd_entities=1, formd_count=1, pct_506c=None, formd_sold_usd_m=None,
      google_ads=0, yt_subs=113, yt_vids=19, min_check_usd=None,
      investors=None, countries=1, dist_channels=None,
      channel_model="Институциональные партнёры; публичной инвест-воронки нет",
      note="~12 000 акров приобретено/контролируется, $3,6 млрд сделок, 18 000 юнитов построено. 2 091 акр в Джексонвилле ($53,2 млн, апрель 2026) + 7 850 акров в Clay County."),

 dict(tier=1, name="13th Floor Investments", hq="Miami, FL", domain="13fi.com",
      model="Вертикально интегрированный инвестор: земля + девелопмент, вход на любом уровне капстека",
      acres=None, projects=70, markets=2, aum_usd_m=5000, raised_usd_m=None,
      regd_entities=1, formd_count=1, pct_506c=None, formd_sold_usd_m=None,
      google_ads=0, yt_subs=None, yt_vids=None, min_check_usd=None,
      investors=None, countries=1, dist_channels=None,
      channel_model="JV с институциональными и семейными партнёрами (Adler Group, LeFrak, Related Group)",
      note="$5 млрд+ недвижимости под управлением, 70+ сделок за 17 лет, 40+ профессионалов. Таргет 2,0x / 20%+."),

 dict(tier=1, name="Bedrock Land Finance (TWG Global)", hq="Dallas, TX", domain="bedrocklandfinance.com",
      model="Land banking origination + A&D-кредитование домостроителей",
      acres=None, projects=None, markets=None, aum_usd_m=5000, raised_usd_m=None,
      regd_entities=None, formd_count=None, pct_506c=None, formd_sold_usd_m=None,
      google_ads=0, yt_subs=None, yt_vids=None, min_check_usd=None,
      investors=None, countries=1, dist_channels=None,
      channel_model="Эксклюзивный land banking партнёр Guggenheim Investments (май 2026); институциональный капитал",
      note="Цель — профинансировать не менее $5 млрд девелопмента. Розничной рекламы нет в принципе."),

 # ---------------- TIER 2: розничные земельные/девелоперские фонды ----------------
 dict(tier=2, name="DLP Capital", hq="St. Augustine, FL", domain="dlpcapital.com",
      model="Семейство фондов: девелопмент, кредитование, attainable housing; инвестор входит в фонд",
      acres=None, projects=25, markets=22, aum_usd_m=5500, raised_usd_m=4206,
      regd_entities=22, formd_count=51, pct_506c=73.8, formd_sold_usd_m=4206,
      google_ads=16, yt_subs=878, yt_vids=99, min_check_usd=200000,
      investors=11550, countries=1, dist_channels=None,
      channel_model="Контент-машина: ежеквартальные вебинары по каждому фонду, Elite Impact Podcast, Investor Vision Day, события, Inc. 5000 PR",
      note="31/42 предложений — 506(c). Медианный минимальный чек $200 000 — самый высокий в выборке. 11 550 инвесторов, $4,21 млрд продано. 26 000+ юнитов, 37 штатов."),

 dict(tier=2, name="Caliber (NASDAQ: CWD)", hq="Scottsdale, AZ", domain="caliberco.com",
      model="Альтернативный управляющий: земля, хоспиталити, OZ-фонды; розничный и HNW-капитал",
      acres=None, projects=None, markets=None, aum_usd_m=2900, raised_usd_m=1321,
      regd_entities=29, formd_count=106, pct_506c=57.7, formd_sold_usd_m=1321,
      google_ads=24, yt_subs=None, yt_vids=None, min_check_usd=20000,
      investors=4547, countries=1, dist_channels=None,
      channel_model="Вебинары, подкасты, видео-отзывы инвесторов, book-an-investor-call, события; LinkedIn/FB/X/IG/YouTube",
      note="$2,9 млрд AUM+AUD, $1,32 млрд продано по формам D, 4 547 инвесторов. 39 форм D с 2024 — высокая текущая активность."),

 dict(tier=2, name="GSP REI", hq="Paoli, PA", domain="gsprei.com",
      model="Land Entitlement Fund: покупка raw land → entitlements → продажа национальным домостроителям",
      acres=None, projects=29, markets=40, aum_usd_m=None, raised_usd_m=1698,
      regd_entities=57, formd_count=78, pct_506c=9.2, formd_sold_usd_m=1698,
      google_ads=0, yt_subs=None, yt_vids=None, min_check_usd=25000,
      investors=1647, countries=1, dist_channels=None,
      channel_model="Gated-материалы фонда, прямые продажи аккредитованным; публичной рекламы нет",
      note="57 Reg D-эмитентов, 78 форм D, 38 с 2024 — второй по текущей активности. $1,70 млрд продано, 1 647 инвесторов. Но лишь 9,2% предложений — 506(c): отсюда ноль рекламы."),

 dict(tier=2, name="MLG Capital", hq="Brookfield, WI", domain="mlgcapital.com",
      model="Частные фонды недвижимости + JV-equity опытным операторам",
      acres=None, projects=None, markets=None, aum_usd_m=8800, raised_usd_m=8194,
      regd_entities=45, formd_count=87, pct_506c=82.4, formd_sold_usd_m=8194,
      google_ads=22, yt_subs=193, yt_vids=15, min_check_usd=50000,
      investors=4879, countries=1, dist_channels=None,
      channel_model="Платный поиск + контент; дистрибуция через investment advisors и family offices",
      note="$8,19 млрд продано по формам D, 4 879 инвесторов, 82,4% предложений — 506(c). 33 145 юнитов, $8,8 млрд рыночной стоимости. Лучший в выборке баланс масштаба и юридического права рекламировать."),

 dict(tier=2, name="AcreTrader", hq="Fayetteville, AR", domain="acretrader.com",
      model="Фермерская земля: каждый участок — отдельное SPV, инвестор берёт долю",
      acres=None, projects=100, markets=None, aum_usd_m=None, raised_usd_m=186,
      regd_entities=100, formd_count=101, pct_506c=100.0, formd_sold_usd_m=186,
      google_ads=3, yt_subs=1210, yt_vids=152, min_check_usd=17125,
      investors=7035, countries=1, dist_channels=None,
      channel_model="Контент-маркетинг и образовательная воронка; платная реклама минимальна при 100% праве на неё",
      note="101 из 101 предложения — 506(c): полное право рекламировать, но лишь 3 объявления. Сознательный выбор контента вместо платного трафика."),

 dict(tier=2, name="FarmTogether", hq="San Francisco, CA", domain="farmtogether.com",
      model="Фермерская земля, premium permanent crops; SPV на объект",
      acres=None, projects=44, markets=None, aum_usd_m=None, raised_usd_m=260,
      regd_entities=44, formd_count=60, pct_506c=93.3, formd_sold_usd_m=260,
      google_ads=24, yt_subs=None, yt_vids=None, min_check_usd=15000,
      investors=5891, countries=1, dist_channels=None,
      channel_model="Платный поиск + контент; ~1 новое предложение в месяц",
      note="44 Reg D-эмитента, 56/60 — 506(c). Самый активный платный рекламодатель среди землевладельческих структур."),

 dict(tier=2, name="Urban Catalyst", hq="San Jose, CA", domain="urbancatalyst.com",
      model="Opportunity Zone девелопер: земля + ground-up проекты в центре Сан-Хосе",
      acres=None, projects=7, markets=1, aum_usd_m=None, raised_usd_m=131,
      regd_entities=5, formd_count=8, pct_506c=66.7, formd_sold_usd_m=13,
      google_ads=2, yt_subs=67, yt_vids=16, min_check_usd=100000,
      investors=356, countries=1, dist_channels=None,
      channel_model="OZ-ниша: вебинары, блог, инвест-спотлайты; Fund II закрыт в конце 2025, Fund III — 2027",
      note="Fund I: $131 млн от 356 инвесторов — средний чек ~$368 000. 7 ground-up проектов."),

 dict(tier=2, name="Allied Development", hq="USA", domain="alliedlandfund.com",
      model="Land entitlement fund: покупка земли → entitlements → продажа домостроителям",
      acres=None, projects=30, markets=None, aum_usd_m=None, raised_usd_m=None,
      regd_entities=2, formd_count=4, pct_506c=None, formd_sold_usd_m=None,
      google_ads=5, yt_subs=54, yt_vids=76, min_check_usd=100000,
      investors=None, countries=1, dist_channels=None,
      channel_model="Воронка 'Book a Clarity Call' (15–30 мин), gated-материалы, 15% preferred, квартальные выплаты",
      note="30+ завершённых entitlement-проектов. Минимум $100 000. Чистый пример 506(c)-воронки малого масштаба."),
]

WATCHLIST = [
 dict(tier=2, name="NexMetro Communities", hq="Phoenix, AZ", domain="nexmetro.com",
      model="Build-to-rent: покупка земли, строительство horizontal-комьюнити, фонды для инвесторов",
      acres=None, projects=None, markets=5, aum_usd_m=None, raised_usd_m=None,
      regd_entities=13, formd_count=35, pct_506c=31.4, formd_sold_usd_m=470,
      google_ads=0, yt_subs=None, yt_vids=None, min_check_usd=10000,
      investors=835, countries=1, dist_channels=None,
      channel_model="Direct Access Fund 2024/2025/2026, Dividend Fund; прямые отношения с инвесторами",
      note="13 Reg D-эмитентов, 35 форм D, 16 с 2024 — высокая текущая активность. Ежегодные фонды прямого доступа."),

 dict(tier=2, name="Belpointe OZ (NYSE American: OZ)", hq="Greenwich, CT", domain="belpointeoz.com",
      model="Публичный Opportunity Zone фонд: земля + ground-up девелопмент",
      acres=None, projects=None, markets=None, aum_usd_m=3000, raised_usd_m=None,
      regd_entities=7, formd_count=30, pct_506c=None, formd_sold_usd_m=None,
      google_ads=0, yt_subs=None, yt_vids=None, min_check_usd=None,
      investors=None, countries=1, dist_channels=None,
      channel_model="Публичный листинг — доступен неаккредитованным без Reg D-ограничений",
      note="Единственный публично торгуемый OZ-фонд. Рекламы не ведёт."),
 dict(tier=1, name="GTIS Partners", hq="New York, NY", domain="gtispartners.com",
      model="Residential land + development funds",
      acres=None, projects=None, markets=None, aum_usd_m=4500, raised_usd_m=None,
      regd_entities=41, formd_count=88, pct_506c=15.0, formd_sold_usd_m=3634,
      google_ads=0, yt_subs=None, yt_vids=None, min_check_usd=100000,
      investors=4000, countries=2, dist_channels=None,
      channel_model="Институциональные LP",
      note="41 Reg D-эмитент, но большинство — бразильские фонды, а не земля США. Не проходит F1 по географии."),
 dict(tier=1, name="The Adler Group", hq="Miami, FL", domain="adlergroup.com",
      model="Девелопер + Florida Land Bank fund совместно с Apollo",
      acres=None, projects=None, markets=1, aum_usd_m=None, raised_usd_m=None,
      regd_entities=None, formd_count=None, pct_506c=None, formd_sold_usd_m=None,
      google_ads=0, yt_subs=None, yt_vids=None, min_check_usd=None,
      investors=None, countries=1, dist_channels=None,
      channel_model="Партнёрство с Apollo Real Estate Advisors",
      note="18 млн кв. футов за историю. Florida Land Bank fund. Данных по фонду в открытых источниках недостаточно для скоринга."),
 dict(tier=1, name="Hines", hq="Houston, TX", domain="hines.com",
      model="Глобальный девелопер; земля — часть пайплайна, не основной продукт",
      acres=None, projects=None, markets=None, aum_usd_m=93000, raised_usd_m=None,
      regd_entities=33, formd_count=228, pct_506c=None, formd_sold_usd_m=None,
      google_ads=4, yt_subs=None, yt_vids=None, min_check_usd=None,
      investors=None, countries=30, dist_channels=None,
      channel_model="Институциональные LP + частный wealth-канал",
      note="Самый активный Reg D-филер выборки: 61 форма с 2024. Но земля не основной актив — не проходит F1."),
]

# --------------------- ПАНЕЛЬ-БЕНЧМАРК (не входят в рейтинг) ---------------------
# Это НЕ земельные игроки. Это те, у кого реально нужно копировать рекламную механику.
BENCHMARK = [
 dict(name="Fundrise", domain="fundrise.com", google_ads=200, yt_subs=None, regd_entities=43, formd_count=119,
      note="200+ активных объявлений Google — потолок выдачи ATC. Reg A+ позволяет массовую рекламу неаккредитованным."),
 dict(name="Cardone Capital", domain="cardonecapital.com", google_ads=69, yt_subs=3140000, regd_entities=23, formd_count=28,
      note="3,14 млн подписчиков YouTube, 6 800 видео. 23 из 23 предложений — 506(c), $831 млн продано, 4 526 инвесторов, минимальный чек $100 000. Чистейший пример связки личный бренд и капитал."),
 dict(name="Origin Investments", domain="origininvestments.com", google_ads=51, yt_subs=1590, regd_entities=None, formd_count=None,
      note="51 объявление + 192 видео. Классическая 506(c) контент-воронка."),
 dict(name="CrowdStreet", domain="crowdstreet.com", google_ads=32, yt_subs=14, regd_entities=24, formd_count=38,
      note="Маркетплейс сделок, не землевладелец."),
 dict(name="RealtyMogul", domain="realtymogul.com", google_ads=29, yt_subs=None, regd_entities=131, formd_count=212,
      note="131 Reg D-эмитент — рекорд выборки по числу структур."),
 dict(name="New Western", domain="newwestern.com", google_ads=200, yt_subs=None, regd_entities=None, formd_count=None,
      note="200+ объявлений. Маркетплейс инвест-недвижимости, не земля. Эталон объёма платного трафика."),
 dict(name="Brookfield Residential", domain="brookfieldresidential.com", google_ads=700, yt_subs=None, regd_entities=None, formd_count=None,
      note="700 объявлений — но это продажа ДОМОВ конечным покупателям, НЕ привлечение инвесторов. Ловушка при наивном чтении рекламных счётчиков."),
 dict(name="Land Advisors Organization", domain="landadvisors.com", google_ads=16, yt_subs=None, regd_entities=None, formd_count=None,
      note="Крупнейший земельный брокер США. Исключён по F3 (брокеридж), но полезен как источник сделок."),
]


# ---------------------------------------------------------------------------
# СКОРИНГ
# ---------------------------------------------------------------------------
def band(v, bands, na=0):
    """bands = [(threshold, score), ...] по убыванию порога."""
    if v is None:
        return na
    for t, s in bands:
        if v >= t:
            return s
    return 0

def score_company(c):
    s = {}
    is_public = "NYSE" in c['name'] or "NASDAQ" in c['name']

    # Блок A — масштаб земли (30)
    s['A1_acres'] = band(c['acres'], [(100000,12),(50000,10),(20000,8),(5000,6),(1000,4),(1,2)])
    s['A2_projects'] = band(c['projects'], [(100,10),(50,8),(20,6),(10,4),(5,2)])
    s['A3_markets'] = band(c['markets'], [(50,8),(20,6),(10,4),(3,2),(1,1)])

    # Блок B — капитальная машина (30)
    s['B1_aum'] = band(c['aum_usd_m'], [(5000,10),(2000,8),(1000,6),(500,4),(100,2),(1,1)])
    s['B2_raised'] = band(c['raised_usd_m'], [(3000,10),(1000,8),(500,6),(100,4),(25,2),(1,1)])
    if c['regd_entities'] is None:
        # публичная компания: листинг = постоянный публичный механизм привлечения
        s['B3_vehicles'] = 5 if is_public else 0
    else:
        s['B3_vehicles'] = band(c['regd_entities'], [(100,10),(50,8),(20,6),(10,4),(5,2),(1,1)])

    # Блок C — маркетинговая интенсивность (30)
    s['C1_google'] = band(c['google_ads'], [(100,8),(50,7),(20,6),(10,4),(1,2)])
    s['C2_social_paid'] = c.get('_c2', 0)
    s['C3_content'] = band(c['yt_vids'], [(1000,7),(100,5),(20,3),(1,1)]) if c['yt_vids'] else c.get('_c3', 0)
    s['C4_distribution'] = c.get('_c4', 0)
    s['C5_organic'] = band(c['yt_subs'], [(100000,4),(10000,3),(1000,2),(100,1)])

    # Блок D — инфраструктура конверсии (10)
    s['D1_funnel'] = c.get('_d1', 0)
    if is_public:
        s['D2_solicit'] = 4
    elif c['pct_506c'] is None:
        s['D2_solicit'] = c.get('_d2', 0)
    elif c['pct_506c'] >= 90:
        s['D2_solicit'] = 5
    elif c['pct_506c'] >= 40:
        s['D2_solicit'] = 3
    elif c['pct_506c'] > 0:
        s['D2_solicit'] = 1
    else:
        s['D2_solicit'] = 0

    s['TOTAL'] = sum(s.values())
    s['BLOCK_A'] = s['A1_acres'] + s['A2_projects'] + s['A3_markets']
    s['BLOCK_B'] = s['B1_aum'] + s['B2_raised'] + s['B3_vehicles']
    s['BLOCK_C'] = s['C1_google'] + s['C2_social_paid'] + s['C3_content'] + s['C4_distribution'] + s['C5_organic']
    s['BLOCK_D'] = s['D1_funnel'] + s['D2_solicit']
    s['TOTAL'] = s['BLOCK_A'] + s['BLOCK_B'] + s['BLOCK_C'] + s['BLOCK_D']
    key = ['acres','projects','markets','aum_usd_m','raised_usd_m','regd_entities',
           'pct_506c','google_ads','min_check_usd','investors']
    filled = sum(1 for k in key if c.get(k) is not None)
    s['DATA_COMPLETENESS_PCT'] = round(100*filled/len(key))
    return s

# Экспертные под-оценки там, где нет прямого измерения (помечены как E в evidence.csv)
MANUAL = {
 "Walton Global":                 dict(_c2=2, _c3=5, _c4=6, _d1=3, _d2=0),
 "Domain Real Estate Partners":   dict(_c2=0, _c3=1, _c4=4, _d1=1, _d2=0),
 "Millrose Properties (NYSE: MRP)": dict(_c2=0, _c3=1, _c4=3, _d1=3, _d2=4),
 "Howard Hughes Holdings (NYSE: HHH)": dict(_c2=3, _c3=5, _c4=4, _d1=3, _d2=4),
 "The St. Joe Company (NYSE: JOE)": dict(_c2=2, _c3=3, _c4=3, _d1=2, _d2=4),
 "Tejon Ranch Co (NYSE: TRC)":    dict(_c2=0, _c3=1, _c4=3, _d1=2, _d2=4),
 "Five Point Holdings (NYSE: FPH)": dict(_c2=1, _c3=2, _c4=3, _d1=2, _d2=4),
 "Bedrock Land Finance (TWG Global)": dict(_c2=0, _c3=0, _c4=4, _d1=1, _d2=0),
 "Forestar Group (NYSE: FOR)":    dict(_c2=1, _c3=2, _c4=3, _d1=2, _d2=4),
 "BTI Partners":                  dict(_c2=1, _c3=2, _c4=2, _d1=1, _d2=0),
 "13th Floor Investments":        dict(_c2=0, _c3=1, _c4=2, _d1=1, _d2=0),
 "The Adler Group":               dict(_c2=0, _c3=1, _c4=2, _d1=1, _d2=0),
 "DLP Capital":                   dict(_c2=4, _c3=7, _c4=5, _d1=5, _d2=3),
 "Caliber (NASDAQ: CWD)":         dict(_c2=4, _c3=5, _c4=4, _d1=5, _d2=4),
 "MLG Capital":                   dict(_c2=3, _c3=3, _c4=3, _d1=4, _d2=3),
 "Urban Catalyst":                dict(_c2=2, _c3=3, _c4=3, _d1=4, _d2=3),
 "AcreTrader":                    dict(_c2=2, _c3=5, _c4=2, _d1=4, _d2=3),
 "FarmTogether":                  dict(_c2=3, _c3=3, _c4=2, _d1=4, _d2=3),
 "Allied Development":            dict(_c2=2, _c3=2, _c4=1, _d1=5, _d2=5),
 "Crow Holdings":                 dict(_c2=0, _c3=1, _c4=4, _d1=1, _d2=0),
 "GSP REI":                       dict(_c2=1, _c3=1, _c4=2, _d1=4, _d2=3),
 "Belpointe OZ (NYSE American: OZ)": dict(_c2=1, _c3=1, _c4=2, _d1=3, _d2=4),
}

def main():
    for c in COMPANIES:
        c.update(MANUAL.get(c['name'], {}))
        c['scores'] = score_company(c)
    ranked = sorted(COMPANIES, key=lambda c: -c['scores']['TOTAL'])
    for i, c in enumerate(ranked, 1):
        c['rank'] = i

    cols = ['rank','tier','name','hq','domain','acres','projects','markets','aum_usd_m','raised_usd_m',
            'regd_entities','formd_count','pct_506c','formd_sold_usd_m','google_ads','yt_subs','yt_vids',
            'min_check_usd','investors','countries','dist_channels','model','channel_model','note']
    scols = ['BLOCK_A','BLOCK_B','BLOCK_C','BLOCK_D','TOTAL','DATA_COMPLETENESS_PCT','A1_acres','A2_projects','A3_markets',
             'B1_aum','B2_raised','B3_vehicles','C1_google','C2_social_paid','C3_content',
             'C4_distribution','C5_organic','D1_funnel','D2_solicit']

    with open(os.path.join(OUT,'companies.csv'),'w',newline='',encoding='utf-8') as f:
        w = csv.writer(f); w.writerow(cols + scols)
        for c in ranked:
            w.writerow([c.get(k,'') if c.get(k) is not None else '' for k in cols] +
                       [c['scores'][k] for k in scols])

    with open(os.path.join(OUT,'benchmark.csv'),'w',newline='',encoding='utf-8') as f:
        w = csv.writer(f); w.writerow(['name','domain','google_ads','yt_subs','regd_entities','formd_count','note'])
        for b in sorted(BENCHMARK, key=lambda x: -(x['google_ads'] or 0)):
            w.writerow([b['name'],b['domain'],b['google_ads'],b.get('yt_subs') or '',
                        b.get('regd_entities') or '', b.get('formd_count') or '', b['note']])

    json.dump({'companies':ranked,'benchmark':BENCHMARK},
              open(os.path.join(OUT,'companies.json'),'w'), ensure_ascii=False, indent=1)

    print(f"{'#':>2} {'Компания':40} {'A':>3} {'B':>3} {'C':>3} {'D':>3} {'ИТОГО':>6} {'данных%':>8}")
    print('-'*80)
    for c in ranked:
        s=c['scores']
        print(f"{c['rank']:>2} {c['name'][:40]:40} {s['BLOCK_A']:>3} {s['BLOCK_B']:>3} {s['BLOCK_C']:>3} {s['BLOCK_D']:>3} {s['TOTAL']:>6} {s['DATA_COMPLETENESS_PCT']:>7}%")
    print(f"\nВ основном рейтинге: {len(ranked)} | В списке наблюдения: {len(WATCHLIST)} | Бенчмарк: {len(BENCHMARK)}")

if __name__ == '__main__':
    main()
