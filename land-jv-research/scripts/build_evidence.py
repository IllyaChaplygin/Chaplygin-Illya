#!/usr/bin/env python3
"""
Журнал доказательств: по одной строке на каждый значимый факт.
status: FACT-M (измерено), FACT-D (раскрыто компанией/отчётностью), EST (оценка модели)
"""
import csv, os

OUT = os.path.join(os.path.dirname(__file__), '..', 'data')

ATC = "https://adstransparency.google.com/?region=US&domain={}"
SEC_FD = "https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&company={}&type=D&dateb=&owner=include&count=100"

E = []
def add(company, metric, value, status, source, note=""):
    E.append(dict(company=company, metric=metric, value=value, status=status, source=source, note=note))

# ---------------------------------------------------------------- измерения рекламы
MEASURED_ADS = [
 ("Walton Global","walton.com",0),("Millrose Properties","millroseproperties.com",0),
 ("DLP Capital","dlpcapital.com",16),("Caliber","caliberco.com",24),
 ("MLG Capital","mlgcapital.com",22),("AcreTrader","acretrader.com",3),
 ("Forestar Group","forestar.com",60),("Howard Hughes Holdings","howardhughes.com",0),
 ("The St. Joe Company","joe.com",0),("FarmTogether","farmtogether.com",24),
 ("GSP REI","gsprei.com",0),("Domain Real Estate Partners","domainrep.com",0),
 ("Crow Holdings","crowholdings.com",0),("BTI Partners","btipartners.com",0),
 ("Tejon Ranch Co","tejonranch.com",0),("Allied Development","alliedlandfund.com",5),
 ("13th Floor Investments","13fi.com",0),("Urban Catalyst","urbancatalyst.com",2),
 ("Five Point Holdings","fivepoint.com",0),("Bedrock Land Finance","bedrocklandfinance.com",0),
 ("Fundrise [бенчмарк]","fundrise.com",200),("Cardone Capital [бенчмарк]","cardonecapital.com",69),
 ("Origin Investments [бенчмарк]","origininvestments.com",51),("CrowdStreet [бенчмарк]","crowdstreet.com",32),
 ("RealtyMogul [бенчмарк]","realtymogul.com",29),("New Western [бенчмарк]","newwestern.com",200),
 ("Brookfield Residential [бенчмарк]","brookfieldresidential.com",700),
 ("Land Advisors [бенчмарк]","landadvisors.com",16),("LandGate [бенчмарк]","landgate.com",53),
 ("EquityMultiple [бенчмарк]","equitymultiple.com",7),("Yieldstreet [бенчмарк]","yieldstreet.com",2),
 ("Belpointe OZ","belpointeoz.com",0),("NexMetro","nexmetro.com",0),
 ("GTIS Partners","gtispartners.com",0),("Promised Land OZ","promisedland.fund",0),
 ("Hines","hines.com",4),
]
for name, dom, n in MEASURED_ADS:
    add(name, "Активных объявлений Google/YouTube (США)", n, "FACT-M", ATC.format(dom),
        "Прямое измерение Google Ads Transparency Center, 05.10.2026. ATC не раскрывает суммы — только наличие и количество.")

# ---------------------------------------------------------------- SEC Form D
FORMD = [
 # company, entities, formD, since2024, pct506c, sold_m, medMin, investors
 ("Walton Global", 44, 118, 6, 0.0, 8276, 500, 516),
 ("DLP Capital", 22, 51, 21, 73.8, 4206, 200000, 11550),
 ("Caliber", 29, 106, 39, 57.7, 1321, 20000, 4547),
 ("MLG Capital", 45, 87, 22, 82.4, 8194, 50000, 4879),
 ("GSP REI", 57, 78, 38, 9.2, 1698, 25000, 1647),
 ("Crow Holdings", 44, 84, 17, 0.0, 38235, 5000000, 2350),
 ("AcreTrader", 100, 101, 1, 100.0, 186, 17125, 7035),
 ("FarmTogether", 44, 60, 10, 93.3, 260, 15000, 5891),
 ("Urban Catalyst", 5, 8, 0, 66.7, 13, 100000, 27),
 ("NexMetro", 13, 35, 16, 31.4, 470, 10000, 835),
 ("GTIS Partners", 41, 88, 7, 15.0, 3634, 100000, 4000),
 ("Cardone Capital [бенчмарк]", 23, 28, 7, 100.0, 831, 100000, 4526),
 ("RealtyMogul [бенчмарк]", 131, 212, 0, None, None, None, None),
 ("EquityMultiple [бенчмарк]", 100, 102, 0, None, None, None, None),
 ("Fundrise [бенчмарк]", 42, 118, 13, None, None, None, None),
 ("Hines", 33, 228, 61, None, None, None, None),
 ("Belpointe OZ", 7, 30, 4, None, None, None, None),
 ("BTI Partners", 1, 1, 0, None, None, None, None),
 ("13th Floor Investments", 1, 1, 0, None, None, None, None),
 ("Allied Development", 2, 4, 0, None, None, None, None),
]
for (name, ent, fd, s24, pct, sold, mn, inv) in FORMD:
    q = name.split(' [')[0].replace(' ', '+')
    src = SEC_FD.format(q)
    add(name, "Число отдельных Reg D-эмитентов (после очистки от однофамильцев)", ent, "FACT-M", src,
        "SEC EDGAR, выгрузка 05.10.2026. Имена эмитентов отфильтрованы по включающему/исключающему паттерну.")
    add(name, "Подано форм D (всего)", fd, "FACT-M", src, "")
    add(name, "Подано форм D с 01.01.2024 (текущая активность)", s24, "FACT-M", src, "")
    if pct is not None:
        add(name, "Доля предложений 506(c), % — право на публичную рекламу", pct, "FACT-M", src,
            "Поле federalExemptionsExclusions в primary_doc.xml. 506(b) ЗАПРЕЩАЕТ general solicitation.")
    if sold is not None:
        add(name, "Сумма продано по формам D, $млн", sold, "FACT-M", src,
            "Сумма totalAmountSold по разобранным формам с 2019 г. Не равна AUM.")
    if mn is not None:
        add(name, "Медианный минимальный чек по формам D, $", mn, "FACT-M", src,
            "Медиана minimumInvestmentAccepted по ненулевым значениям.")
    if inv is not None:
        add(name, "Инвесторов по формам D (сумма totalNumberAlreadyInvested)", inv, "FACT-M", src,
            "Учитывает только инвесторов в этих конкретных предложениях; офшорные инвесторы могут не попадать.")

# ---------------------------------------------------------------- раскрытые показатели
D = [
 ("Walton Global","Акров под управлением",80500,"https://walton.com/our-global-approach/"),
 ("Walton Global","Master-planned communities",198,"https://walton.com/our-global-approach/"),
 ("Walton Global","Региональных рынков",34,"https://walton.com/our-global-approach/"),
 ("Walton Global","Land under management, $млн",4640,"https://walton.com/"),
 ("Walton Global","Распределено инвесторам, $млн",3140,"https://walton.com/"),
 ("Walton Global","Инвесторов",89000,"https://walton.com/"),
 ("Walton Global","Стран",91,"https://walton.com/"),
 ("Walton Global","Одобренных каналов дистрибуции",300,"https://walton.com/our-global-approach/"),
 ("Walton Global","Новый фонд American Builder Growth & Income, целевой размер $млн",500,"https://www.businesswire.com/news/home/20260212613090/en/Walton-Global-Launches-the-American-Builder-Growth-and-Income-Fund"),
 ("Millrose Properties","Homesites",143771,"https://www.fool.com/earnings/call-transcripts/2026/08/11/millrose-properties-mrp-q2-2026-earnings-call-transcript/"),
 ("Millrose Properties","Комьюнити",877,"https://www.fool.com/earnings/call-transcripts/2026/08/11/millrose-properties-mrp-q2-2026-earnings-call-transcript/"),
 ("Millrose Properties","Штатов",30,"https://www.fool.com/earnings/call-transcripts/2026/08/11/millrose-properties-mrp-q2-2026-earnings-call-transcript/"),
 ("Millrose Properties","Совокупные активы, $млн",9700,"https://www.gurufocus.com/news/9001874/millrose-properties-reports-second-quarter-2026-financial-results"),
 ("Millrose Properties","Ожидаемые takedown-поступления, $млн",16300,"https://www.fool.com/earnings/call-transcripts/2026/08/11/millrose-properties-mrp-q2-2026-earnings-call-transcript/"),
 ("Domain Real Estate Partners","Профинансировано проектов",720,"https://www.pgim.com/us/en/borrower/about/newsroom/press-releases/pgim-and-domain-real-estate-partners-surpass-4-billion-in-u-s--land-banking-transactions"),
 ("Domain Real Estate Partners","Совокупная стоимость проектов, $млн",16000,"https://www.pgim.com/us/en/borrower/about/newsroom/press-releases/pgim-and-domain-real-estate-partners-surpass-4-billion-in-u-s--land-banking-transactions"),
 ("Domain Real Estate Partners","Развёрнуто в портфеле, $млн",6000,"https://alternativecreditinvestor.com/2026/05/26/pgim-and-domain-hit-4bn-in-us-land-banking-deals/"),
 ("Howard Hughes Holdings","Акров",118000,"https://investor.howardhughes.com/news-releases/news-release-details/howard-hughes-communitiestm-celebrates-grand-opening-teravalistm"),
 ("The St. Joe Company","Акров во Флориде",165000,"https://en.wikipedia.org/wiki/St._Joe_Company"),
 ("The St. Joe Company","Выручка неконсолидированных JV, Q1 2026, $млн",56.1,"https://ir.joe.com/news-releases/news-release-details/st-joe-company-reports-first-quarter-2026-results-and-declares"),
 ("Tejon Ranch Co","Акров",270000,"https://umbrex.com/resources/company-profiles/tejon-ranch-company/"),
 ("Forestar Group","Лотов (владение + контроль)",91700,"https://www.businesswire.com/news/home/20260721189655/en/Forestar-Reports-Fiscal-2026-Third-Quarter-Results"),
 ("Forestar Group","Рынков",65,"https://www.businesswire.com/news/home/20260721189655/en/Forestar-Reports-Fiscal-2026-Third-Quarter-Results"),
 ("Five Point Holdings","Homesites Valencia",21500,"https://ir.fivepoint.com/news-releases/2026/01-29-2026-211033188"),
 ("Five Point Holdings","Great Park: акров",2100,"https://therealdeal.com/la/2026/07/29/five-point-sells-irvine-land-to-erickson-for-159-million/"),
 ("DLP Capital","AUM, $млн",5500,"https://dlpcapital.com/about"),
 ("DLP Capital","Аккредитованных инвесторов",4000,"https://dlpcapital.com/about"),
 ("DLP Capital","Юнитов в портфеле",26000,"https://dlpcapital.com/about"),
 ("Caliber","AUM + AUD, $млн",2900,"https://www.caliberco.com/"),
 ("Caliber","Привлечённый капитал, $млн",667,"https://www.caliberco.com/"),
 ("MLG Capital","Совокупная рыночная стоимость, $млн",8800,"https://mlgcapital.com/"),
 ("MLG Capital","Юнитов",33145,"https://mlgcapital.com/news/seven-recent-acquisitions-grow-mlgs-footprint-to-more-than-22300-units-nationwide/"),
 ("BTI Partners","Акров приобретено/контролируется",12000,"https://database.thesisdriven.com/developers/bti-partners"),
 ("BTI Partners","Объём сделок за историю, $млн",3600,"https://database.thesisdriven.com/developers/bti-partners"),
 ("BTI Partners","Покупка в Джексонвилле, апрель 2026, акров",2091,"https://www.jaxdailyrecord.com/news/2026/apr/24/fort-lauderdale-developer-purchases-nearly-2100-acres-in-west-jacksonville-for-master-planned-community/"),
 ("13th Floor Investments","Недвижимости под управлением, $млн",5000,"https://13fi.com/about"),
 ("13th Floor Investments","Сделок за историю",70,"https://13fi.com/about"),
 ("GSP REI","Завершённых land-проектов",29,"https://gsprei.com/track-record/"),
 ("GSP REI","Минимальный чек фонда, $",100000,"https://gsprei.com/land-entitlement-offerings/"),
 ("Allied Development","Завершённых entitlement-проектов",30,"https://alliedlandfund.com/"),
 ("Allied Development","Минимальный чек, $",100000,"https://alliedlandfund.com/"),
 ("Urban Catalyst","Fund I: привлечено, $млн",131,"https://www.prnewswire.com/news-releases/urban-catalyst-opportunity-zone-fund-raises-131m-in-first-fund-continues-strong-growth-in-2021-301208179.html"),
 ("Urban Catalyst","Fund I: инвесторов",356,"https://www.prnewswire.com/news-releases/urban-catalyst-opportunity-zone-fund-raises-131m-in-first-fund-continues-strong-growth-in-2021-301208179.html"),
 ("Bedrock Land Finance","Целевой объём финансирования, $млн",5000,"https://www.businesswire.com/news/home/20260519206916/en/Bedrock-Named-Exclusive-Land-Banking-Partner-for-Guggenheim-Investments"),
 ("Crow Holdings","AUM, $млн",33000,"https://www.crowholdings.com/"),
]
for name, metric, val, src in D:
    add(name, metric, val, "FACT-D", src, "Раскрыто компанией или в отчётности/пресс-релизе.")

# ---------------------------------------------------------------- YouTube
YT = [("DLP Capital",878,99),("AcreTrader",1210,152),("MLG Capital",193,15),
      ("Urban Catalyst",67,16),("Tejon Ranch Co",5,1),("The St. Joe Company",1110,228),
      ("BTI Partners",113,19),("Allied Development",54,76),
      ("Cardone Capital [бенчмарк]",3140000,6800),("Origin Investments [бенчмарк]",1590,192),
      ("RealtyMogul [бенчмарк]",987,19),("EquityMultiple [бенчмарк]",558,324),
      ("Harvest Returns [бенчмарк]",648,60),("New Western [бенчмарк]",402,25),
      ("CrowdStreet [бенчмарк]",14,None)]
for name, subs, vids in YT:
    add(name, "YouTube: подписчиков", subs, "FACT-M", "https://www.youtube.com/", "Снято 05.10.2026.")
    if vids: add(name, "YouTube: видео", vids, "FACT-M", "https://www.youtube.com/", "Снято 05.10.2026.")

# ---------------------------------------------------------------- бенчмарки затрат (основа оценок)
B = [
 ("Бенчмарк","Meta CPL, аккредитованный инвестор (синдикация), $","50-100","https://gowercrowd.com/real-estate-syndication/how-to-generate-investor-leads"),
 ("Бенчмарк","LinkedIn CPL, синдикация, $","250-500","https://gowercrowd.com/real-estate-syndication/how-to-generate-investor-leads"),
 ("Бенчмарк","LinkedIn CPL, financial services, $","90-180","https://postiv.ai/blog/linkedin-advertising-costs"),
 ("Бенчмарк","Google Ads CPL, средний 2026, $","79.14","https://foundrycro.com/blog/cost-per-lead-benchmarks-by-industry-2026/"),
 ("Бенчмарк","Meta CPL, средний 2026, $","26.43","https://foundrycro.com/blog/cost-per-lead-benchmarks-by-industry-2026/"),
 ("Бенчмарк","Подкаст-спонсорство CPL, $","60-200","https://atomicfunnels.com/funnel-real-estate-syndication/"),
 ("Бенчмарк","Стоимость одного профинансировавшего инвестора, $","3500-4500","https://gowercrowd.com/real-estate-syndication/how-to-generate-investor-leads"),
 ("Бенчмарк","Конверсия Meta-лида в инвестора за 90 дней, %","2","https://gowercrowd.com/real-estate-syndication/how-to-generate-investor-leads"),
 ("Бенчмарк","Rule 506(c): публичная реклама разрешена при верификации аккредитации","да","https://www.sec.gov/resources-small-businesses/exempt-offerings/general-solicitation-rule-506c"),
 ("Бенчмарк","Rule 506(b): general solicitation запрещён","да","https://www.sec.gov/resources-small-businesses/exempt-offerings/general-solicitation-rule-506c"),
]
for name, metric, val, src in B:
    add(name, metric, val, "FACT-D", src, "Опубликованный отраслевой бенчмарк — основа модельных оценок бюджета.")

def main():
    path = os.path.join(OUT, 'evidence.csv')
    with open(path, 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=['company','metric','value','status','source','note'])
        w.writeheader(); w.writerows(E)
    by = {}
    for e in E: by[e['status']] = by.get(e['status'], 0) + 1
    print(f"Записей в журнале доказательств: {len(E)}")
    for k, v in sorted(by.items()): print(f"  {k}: {v}")
    print(f"Файл: {path}")

if __name__ == '__main__':
    main()
