# -*- coding: utf-8 -*-
import openpyxl,sys
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
from openpyxl.formatting.rule import CellIsRule
from openpyxl.worksheet.properties import PageSetupProperties
SRC="/root/.claude/uploads/8bc0034a-4b53-5022-9bf5-e5697c987387/5927c2e7-RAMEN_FINMODEL_4_TABS_PRESENTATION.xlsx"
OUT=sys.argv[1]
NAVY='223A5E'; INK='172B46'; GREYT='5E6D7E'; BAND='EDF2F8'; BAND2='F2F5F9'; INP='FFF4D2'; KEY='E7EDF5'
BLUE='0000FF'; GREEN='008000'
def font(sz=11,b=False,c=INK,i=False): return Font(name='Arial',size=sz,bold=b,color=c,italic=i)
def fill(c): return PatternFill('solid',fgColor=c)
thin=Side(style='thin',color='D5DCE8')
BOX=Border(left=thin,right=thin,top=thin,bottom=thin)
EUR4='"€"0.0000'; EUR3='"€"0.000'; EUR2='"€"0.00'; UAH='#,##0.00" грн"'; PCT='0.0%'; INT='#,##0'
wb=openpyxl.load_workbook(SRC)
CFG=[dict(name='Цена 120 грн — 6 паллет',src='Финмодель 6 паллет',head='6 паллет',freight=2600,pallets=[1,1,1,1,1,1],rows='строки 6–11'),
     dict(name='Цена 120 грн — 33 паллет',src='Финмодель 33 паллет',head='33 паллеты (Full Truck)',freight=3700,pallets=[7,7,4,5,5,5],rows='строки 30–35')]
NOW=[0.75,0.75,0.65,0.65,0.65,0.65]
def build(cfg):
    idx=wb.sheetnames.index(cfg['name']); del wb[cfg['name']]
    ws=wb.create_sheet(cfg['name'],idx); S="'"+cfg['src']+"'"
    ws.sheet_properties.tabColor=NAVY; ws.sheet_view.showGridLines=False
    for col,w in zip("ABCDEFGHIJK",(4,42,15,15,15,15,15,15,15,17,15)): ws.column_dimensions[col].width=w
    def put(addr,val=None,f=None,fl=None,fmt=None,al=None,merge=None,border=False):
        c=ws[addr]
        if val is not None: c.value=val
        c.font=f or font()
        if fl: c.fill=fill(fl)
        if fmt: c.number_format=fmt
        c.alignment=al or Alignment(vertical='center',wrap_text=False)
        if border: c.border=BOX
        if merge:
            ws.merge_cells(merge)
            first=merge.split(':')[0]; last=merge.split(':')[1]
            if fl:
                for row in ws[merge]:
                    for cc in row: cc.fill=fill(fl)
        return c
    WR=Alignment(vertical='center',wrap_text=True)
    CT=Alignment(horizontal='center',vertical='center',wrap_text=True)
    RT=Alignment(horizontal='right',vertical='center')
    def band(r,text):
        put(f"A{r}",text,font(12,True,'FFFFFF'),NAVY,merge=f"A{r}:K{r}"); ws.row_dimensions[r].height=24
    def hdr(r,cells):
        for rng,t in cells:
            put(rng.split(':')[0],t,font(10,True,'FFFFFF'),'3E5A85',al=CT,merge=(rng if ':' in rng else None))
        ws.row_dimensions[r].height=30
    # ---- заголовок
    ws.row_dimensions[1].height=8
    put("A2",f"SIAS: какую закупочную цену запросить, чтобы встать на полку за 120 грн — {cfg['head']}",font(17,True,INK),merge="A2:K2"); ws.row_dimensions[2].height=30
    put("A3","Читайте сверху вниз: 1 — результат · 2 — исходные данные · 3 — расчёт по шагам от полки к закупке · 4 — куда идёт каждая гривна · 5 — сравнение с ценой сейчас · 6 — чувствительность · 7 — допущения.",font(10,False,GREYT),al=WR,merge="A3:K3"); ws.row_dimensions[3].height=30
    # ---- 1. результат
    band(5,"1. РЕЗУЛЬТАТ — что просить у SIAS")
    rows1=[(6,"Предельная закупка (математический максимум), €/уп.","=F44",EUR4,False),
           (7,"ЦЕНА ДЛЯ ЗАПРОСА SIAS, €/уп. (вниз до евроцента)","=F45",EUR2,True),
           (8,"Своя цена для проверки, €/уп. — необязательно: впишите, и разделы 4–5 пересчитаются",None,EUR2,False),
           (9,"Цена, по которой считаются разделы 4 и 5, €/уп.","=IF(ISNUMBER(F8),F8,F7)",EUR2,False),
           (10,"Полка при этой цене, грн","=F59",UAH,False),
           (11,"Наша маржа при этой цене","=F60",PCT,False),
           (12,"Запас до целевой полки, грн","=F17-F59",UAH,False)]
    for r,lab,fm,nf,big in rows1:
        put(f"B{r}",lab,font(11,big),KEY if big else None,al=WR,merge=f"B{r}:E{r}")
        if r==8: put(f"F{r}",None,font(12,True,BLUE),INP,nf,RT,border=True)
        else: put(f"F{r}",fm,font(18 if big else 11,True,INK),KEY if big else None,nf,RT,border=big)
        ws.row_dimensions[r].height=34 if big else (30 if r==8 else 20)
    put("G7","одна цена для всех шести SKU",font(9,False,GREYT,True),merge="G7:K7")
    put("G8","жёлтая ячейка — можно менять",font(9,False,GREYT,True),merge="G8:K8")
    put("B13",'="Почему такая цена: чтобы полка была "&ROUND(F17,0)&" грн при бонусе сети "&ROUND(F19*100,0)&"% и нашей марже "&ROUND(F18*100,0)&"%, себестоимость пачки в Украине не может быть выше "&ROUND(F37,2)&" грн (≈ €"&ROUND(F38,3)&"). После вычета финансирования, расходов партии и импортных налогов на сам товар остаётся €"&ROUND(F44,4)&" за пачку → запрос €"&ROUND(F45,2)&"."',font(10,False,INK),KEY,al=WR,merge="B13:K13"); ws.row_dimensions[13].height=44
    # ---- 2. параметры
    band(15,"2. ИСХОДНЫЕ ДАННЫЕ")
    put("B16","Жёлтая ячейка с синим шрифтом — исходное число (можно менять) · зелёный шрифт — ссылка на вкладку «"+cfg['src']+"» · чёрный — формула.",font(9,False,GREYT,True),merge="B16:K16")
    P=[(17,"Целевая цена на полке","С НДС, за одну пачку. Это цель всего расчёта.",120,'0.00','грн',"Задано вами",'inp'),
       (18,"Наша целевая маржа","Доля цены партнёру, которая остаётся нам после себестоимости и бонуса.",0.30,PCT,'%',"Задано вами (вкладка «"+cfg['src']+"», колонка маржи 30 %)",'inp'),
       (19,"Бонус торговой сети","Доля цены партнёру, которую отдаём сети.",f"={S}!D4",PCT,'%',"Вкладка «"+cfg['src']+"», ячейка D4",'lnk'),
       (20,"Множитель полки","Во сколько раз полка дороже цены партнёру (наценка магазина).",f"={S}!D6",'0.00','×',"Вкладка «"+cfg['src']+"», ячейка D6",'lnk'),
       (21,"Курс","Пересчёт EUR → грн.",f"={S}!D3",'0.00','грн/€',"Вкладка «"+cfg['src']+"», ячейка D3",'lnk'),
       (22,"Финансирование себестоимости","Процент на кредитные деньги от суммы расходов.",0.025,PCT,'%',"Self-Cost_Sias, ячейка AJ2",'inp'),
       (23,"Импортный НДС","Входит в себестоимость, как в Self-Cost.",0.20,PCT,'%',"Self-Cost_Sias, ячейка U1",'inp'),
       (24,"Пошлина","Для этого товара 0 % (EUR.1).",0,PCT,'%',"Self-Cost_Sias",'inp'),
       (25,"Перевозка партии","Фрахт за всю партию.",cfg['freight'],'#,##0.00','€',"Self-Cost_Sias, итоговая строка блока, колонка V",'inp'),
       (26,"Брокер партии","Услуги брокера за партию.",288,'#,##0.00','€',"Self-Cost_Sias, итоговая строка блока, колонка X",'inp'),
       (27,"Прочие расходы партии","Решения, терминал, карантин и т. п.",42,'#,##0.00','€',"Self-Cost_Sias, итоговая строка блока, колонка AB",'inp'),
       (28,"Доля фрахта в таможенной базе","Часть фрахта, на которую начисляется импортный НДС.",0.7,'0.00','доля',"Параметр таможенной базы в Self-Cost",'inp'),
       (29,"Пачек на паллете","Из спецификации SIAS.",1600,INT,'шт',"Self-Cost_Sias (20 пачек × 80 коробов)",'inp'),
       (30,"Пачек в коробе","Нужно для цены короба.",20,INT,'шт',"Self-Cost_Sias",'inp')]
    for r,lab,desc,val,nf,unit,srcn,kind in P:
        put(f"B{r}",lab,font(11,True))
        put(f"C{r}",desc,font(9,False,GREYT),al=WR,merge=f"C{r}:E{r}")
        put(f"F{r}",val,font(11,True,BLUE if kind=='inp' else GREEN),INP if kind=='inp' else BAND2,nf,RT,border=True)
        put(f"G{r}",unit,font(10,False,GREYT)); put(f"H{r}",srcn,font(9,False,GREYT),al=WR,merge=f"H{r}:K{r}")
        ws.row_dimensions[r].height=26
    # ---- 3. шаги
    band(32,"3. РАСЧЁТ ПО ШАГАМ — от полки 120 грн к закупочной цене (идём от конца к началу)")
    hdr(33,[("A33","№"),("B33","Что считаем"),("C33:E33","Как считаем"),("F33","Результат"),("G33","Ед."),("H33:K33","Почему так")])
    ST=[("Целевая полка","заданная цель","=F17",'0.00','грн',"Это цель: цена, по которой пачка должна стоять на полке магазина (с НДС)."),
        ("Цена, по которой продаём партнёру (магазину)","полка ÷ множитель полки","=F34/F20",'0.00','грн',"Магазин добавляет к цене закупки у нас наценку (множитель). Значит, нам можно продавать ему не дороже, чем полка ÷ множитель."),
        ("Доля цены партнёру на себестоимость","100 % − бонус сети − наша маржа","=1-F19-F18",PCT,'%',"Из цены партнёру часть уходит сети бонусом, часть — наша целевая маржа. Остаток — то, что можно потратить на себестоимость."),
        ("Максимальная себестоимость пачки в Украине","цена партнёру × доля на себестоимость","=F35*F36",'0.00','грн',"Товар + логистика + налоги + финансирование вместе не должны быть дороже. Иначе маржа или полка не сойдутся."),
        ("То же в евро","макс. себестоимость ÷ курс","=F37/F21",EUR4,'€',"Расходы партии у нас считаются в EUR."),
        ("Без процента за финансирование","€ ÷ (1 + финансирование)","=F38/(1+F22)",EUR4,'€',"В себестоимость входит процент за кредитные деньги от суммы расходов; убираем его, чтобы получить сумму расходов без процента."),
        ("Расходы на всю партию, не зависящие от цены товара","перевозка + брокер + прочие + НДС на фрахт","=F25+F26+F27+F25*F28*(F24+F23*(1+F24))",'#,##0.00','€',"Эти расходы нужно оплатить за партию независимо от цены товара. Налог на фрахт считается по доле фрахта в таможенной базе."),
        ("Пачек в партии","из раздела 5","=D71",INT,'шт',"Сумма пачек по всем шести SKU: паллеты × пачек на паллете."),
        ("Расходы партии на одну пачку","расходы партии ÷ пачек","=F40/F41",EUR4,'€',"Делим расходы партии на пачки поровну — так принято в модели."),
        ("Остаётся на товар вместе с импортными налогами","€ без финансирования − расходы на пачку","=F39-F42",EUR4,'€',"Это деньги, которые можно отдать за товар и налоги, начисляемые на цену товара."),
        ("Предельная закупка (максимум)","остаток ÷ (1 + пошлина) ÷ (1 + НДС)","=F43/((1+F24)*(1+F23))",EUR4,'€',"Импортный НДС и пошлина начисляются на цену товара, поэтому делим на (1+пошлина)×(1+НДС) и получаем чистую цену товара — это математический максимум."),
        ("ЦЕНА ДЛЯ ЗАПРОСА SIAS","предельная закупка, округление вниз до евроцента","=ROUNDDOWN(F44,2)",EUR2,'€',"Округляем вниз, чтобы не превысить предел. Это цена для переговоров — одна для всех шести SKU.")]
    for i,(lab,how,fm,nf,unit,why) in enumerate(ST):
        r=34+i; last=(i==len(ST)-1)
        base=KEY if last else (BAND2 if i%2==0 else None)
        put(f"A{r}",i+1,font(10,True,GREYT),base,al=CT)
        put(f"B{r}",lab,font(11,True),base,al=WR)
        put(f"C{r}",how,font(10,False,GREYT),base,al=WR,merge=f"C{r}:E{r}")
        put(f"F{r}",fm,font(12 if last else 11,True),KEY if last else base,nf,RT)
        put(f"G{r}",unit,font(10,False,GREYT),base)
        put(f"H{r}",why,font(9,False,INK),base,al=WR,merge=f"H{r}:K{r}")
        ws.row_dimensions[r].height=46
    # ---- 4. проверка
    band(47,"4. ПРОВЕРКА — куда идёт каждая гривна при цене из ячейки F9 (расчёт вперёд, от закупки к полке)")
    hdr(48,[("A48","№"),("B48","Статья"),("C48:E48","Как считаем"),("F48","грн / пачку"),("G48","% полки"),("H48:K48","Доля полки")])
    CK=[("Товар от поставщика","цена закупки € × курс","=F9*F21",False,'3E8E96'),
        ("Логистика, брокер, прочие","(перевозка + брокер + прочие) ÷ пачек × курс","=(F25+F26+F27)/F41*F21",False,'46538A'),
        ("Пошлина","(цена + фрахт таможни на пачку) × пошлина × курс","=(F9+F25*F28/F41)*F24*F21",False,'46538A'),
        ("Импортный НДС","(цена + фрахт таможни на пачку) × (1+пошлина) × НДС × курс","=(F9+F25*F28/F41)*(1+F24)*F23*F21",False,'46538A'),
        ("Финансирование","сумма четырёх статей выше × % финансирования","=SUM(F49:F52)*F22",False,'46538A'),
        ("ИТОГО · ПОЛНАЯ СЕБЕСТОИМОСТЬ (СС)","сумма статей выше","=SUM(F49:F53)",True,'223A5E'),
        ("Бонус торговой сети","цена партнёру × % бонуса","=F57*F19",False,'C4820A'),
        ("Наша прибыль после бонуса","цена партнёру − СС − бонус","=F57-F54-F55",False,'F9A50B'),
        ("ИТОГО · ЦЕНА ПАРТНЁРУ (с НДС)","СС ÷ (1 − бонус − маржа)","=F54/(1-F19-F18)",True,'223A5E'),
        ("Наценка магазина","цена партнёру × (множитель − 1)","=F57*(F20-1)",False,'A9B0C8'),
        ("ИТОГО · ПОЛКА (с НДС)","цена партнёру × множитель","=F57*F20",True,'223A5E')]
    for i,(lab,how,fm,bold,colr) in enumerate(CK):
        r=49+i; base=KEY if bold else (BAND2 if i%2==0 else None)
        put(f"A{r}",i+1,font(10,True,GREYT),base,al=CT)
        put(f"B{r}",lab,font(11,bold),base)
        put(f"C{r}",how,font(9,False,GREYT),base,al=WR,merge=f"C{r}:E{r}")
        put(f"F{r}",fm,font(11,bold),base,UAH,RT); put(f"G{r}",f"=F{r}/$F$59",font(10,bold,GREYT),base,PCT,RT)
        put(f"H{r}",f'=REPT("█",ROUND(G{r}*38,0))',font(9,False,colr),base,merge=f"H{r}:K{r}")
        ws.row_dimensions[r].height=24 if len(how)<50 else 30
    put("B60","Проверка: наша маржа при этой цене",font(11,True)); put("C60","прибыль ÷ цена партнёру (должна совпасть с целевой)",font(9,False,GREYT),merge="C60:E60"); put("F60","=F56/F57",font(11,True),None,PCT,RT)
    put("B61","Маржа, если полка ровно как цель",font(11,True)); put("C61","(цена партнёру для цели − СС − бонус) ÷ цена партнёру",font(9,False,GREYT),merge="C61:E61"); put("F61","=(F35-F54-F35*F19)/F35",font(11,True),None,PCT,RT)
    put("H60","Запрос округлён вниз → полка чуть ниже цели, маржа чуть выше 30 %.",font(9,False,GREYT,True),merge="H60:K61",al=WR)
    # ---- 5. SKU
    band(63,"5. СРАВНЕНИЕ С ЦЕНОЙ СЕЙЧАС — по каждому SKU")
    hdr(64,[("A64","№"),("B64","Продукт SIAS"),("C64","Паллет"),("D64","Пачек"),("E64","Цена сейчас, €/уп."),("F64","Цена в расчёте, €/уп."),("G64","Снижение"),("H64","Товар в партии сейчас, €"),("I64","Товар в партии при расчётной цене, €"),("J64","Полка сейчас при марже 30 %, грн"),("K64","До цели, грн")])
    ws.row_dimensions[64].height=44
    for i in range(6):
        r=65+i; base=BAND2 if i%2==0 else None
        put(f"A{r}",i+1,font(10,True,GREYT),base,al=CT)
        put(f"B{r}",f"={S}!C{11+i}",font(11,False,GREEN),base)
        put(f"C{r}",cfg['pallets'][i],font(11,True,BLUE),INP,INT,CT,border=True)
        put(f"D{r}",f"=C{r}*$F$29",font(11),base,INT,RT)
        put(f"E{r}",NOW[i],font(11,True,BLUE),INP,EUR2,RT,border=True)
        put(f"F{r}","=$F$9",font(11),base,EUR2,RT)
        put(f"G{r}",f"=1-F{r}/E{r}",font(11,True),base,PCT,RT)
        put(f"H{r}",f"=D{r}*E{r}",font(11),base,'#,##0',RT)
        put(f"I{r}",f"=D{r}*F{r}",font(11),base,'#,##0',RT)
        put(f"J{r}",f"={S}!T{11+i}",font(11,False,GREEN),base,UAH,RT)
        put(f"K{r}",f"=J{r}-$F$17",font(11),base,'+#,##0.00;-#,##0.00',RT)
        ws.row_dimensions[r].height=24
    put("B71","ИТОГО",font(11,True),KEY); 
    for col in "ACEFGJK": put(f"{col}71",None,font(11,True),KEY)
    put("C71","=SUM(C65:C70)",font(11,True),KEY,INT,CT); put("D71","=SUM(D65:D70)",font(11,True),KEY,INT,RT)
    put("G71","=1-I71/H71",font(11,True),KEY,PCT,RT); put("H71","=SUM(H65:H70)",font(11,True),KEY,'#,##0',RT); put("I71","=SUM(I65:I70)",font(11,True),KEY,'#,##0',RT)
    ws.row_dimensions[71].height=24
    put("B72","Название SKU, полка сейчас (колонка T, маржа 30 %) — из вкладки «"+cfg['src']+"»; паллеты и цены сейчас — из Self-Cost_Sias ("+cfg['rows']+"), введены вручную.",font(9,False,GREYT,True),merge="B72:K72")
    # ---- 6. чувствительность
    band(74,"6. ЧУВСТВИТЕЛЬНОСТЬ — какая закупка нужна при другой полке и маржи (€/уп., вниз до евроцента)")
    hdr(75,[("A75",""),("B75","Маржа ↓   /   Полка, грн →")])
    shelves=[105,110,120,130,140]; margins=[0.35,0.30,0.25,0.20,0.15]
    for j,sh in enumerate(shelves):
        col="CDEFG"[j]; put(f"{col}75",sh,font(11,True,'FFFFFF'),'3E5A85',INT,CT,border=True)
    for i,m in enumerate(margins):
        r=76+i
        put(f"B{r}",m,font(11,True,BLUE),INP,'0%',CT,border=True)
        for j in range(5):
            col="CDEFG"[j]
            put(f"{col}{r}",f"=ROUNDDOWN((({col}$75/$F$20*(1-$F$19-$B{r}))/$F$21/(1+$F$22)-$F$40/$F$41)/((1+$F$24)*(1+$F$23)),2)",font(12,True),None,EUR2,CT,border=True)
        ws.row_dimensions[r].height=26
    rng="C76:G80"
    ws.conditional_formatting.add(rng,CellIsRule(operator='greaterThanOrEqual',formula=['MAX($E$65:$E$70)'],fill=fill('9FD8D6'),stopIfTrue=True))
    ws.conditional_formatting.add(rng,CellIsRule(operator='greaterThanOrEqual',formula=['MIN($E$65:$E$70)'],fill=fill('D3EEEC'),stopIfTrue=True))
    ws.conditional_formatting.add(rng,CellIsRule(operator='greaterThanOrEqual',formula=['0.5'],fill=fill('FCE3A6'),stopIfTrue=True))
    ws.conditional_formatting.add(rng,CellIsRule(operator='lessThan',formula=['0.5'],fill=fill('EBBDBD')))
    put("H75","Как читать",font(10,True,NAVY),merge="H75:K75")
    put("H76","Каждая клетка — максимальная закупка для заданных полки и маржи (формулы шага 3). Жёлтые числа слева и сверху можно менять.\nЗелёный — хватает даже при текущей цене SIAS; светло-зелёный — только для SKU по €0,65; жёлтый — нужна скидка до €0,50…0,65; красный — глубже (< €0,50).\nЦель расчёта: полка 120 грн, маржа 30 %.",font(9,False,GREYT),al=WR,merge="H76:K80")
    # ---- 7. допущения
    band(82,"7. ДОПУЩЕНИЯ И ЧТО НЕ УЧТЕНО")
    B7=["Бонус сети, множитель полки и курс — условия модели (вкладка «"+cfg['src']+"»); сетями не подтверждены.",
        "Импортный НДС входит в себестоимость (как в Self-Cost), а цена партнёру и полка — с НДС. Если импортный НДС возмещается, допустимая закупка выше — см. контрольный расчёт ниже.",
        "Расходы партии делятся на пачки поровну; в Self-Cost они делятся пропорционально стоимости SKU, поэтому «СС сейчас» там чуть отличается от «СС при запросе» (≈1 %).",
        "Не учтены: листинг-фи и промо-поддержка сетей, отсрочка платежа, брак и возвраты, колебания курса.",
        "Согласие SIAS на эту цену не получено — это цена для переговоров."]
    for i,t in enumerate(B7):
        put(f"B{83+i}","• "+t,font(10,False,INK),al=WR,merge=f"B{83+i}:K{83+i}"); ws.row_dimensions[83+i].height=30
    put("B89","Контрольный расчёт (не основной): НДС нейтрален — полка без НДС, импортный НДС не в себестоимости",font(11,True,NAVY),BAND2,al=WR,merge="B89:E89")
    put("F89","=ROUNDDOWN(((F17/F20/(1+F23))*(1-F19-F18)/F21/(1+F22)-(F25+F26+F27)/F41)/(1+F24),2)",font(12,True),BAND2,EUR2,RT,border=True)
    put("G89",'="против €"&ROUND(F45,2)&" в основном расчёте ("&IF(F89>=F45,"+","−")&"€"&ROUND(ABS(F89-F45),2)&")"',font(9,False,GREYT,True),al=WR,merge="G89:K89")
    ws.row_dimensions[89].height=34
    # печать
    ws.page_setup.orientation='landscape'; ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0
    ws.sheet_properties.pageSetUpPr=PageSetupProperties(fitToPage=True)
    ws.freeze_panes="A4"
for c in CFG: build(c)
from openpyxl.workbook.properties import CalcProperties
wb.calculation=CalcProperties(fullCalcOnLoad=True)
wb.save(OUT); print("saved",OUT)
