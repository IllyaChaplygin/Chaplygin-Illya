# -*- coding: utf-8 -*-
import openpyxl,sys
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.workbook.properties import CalcProperties
from openpyxl.utils import get_column_letter as CL
SRC="/root/.claude/uploads/8bc0034a-4b53-5022-9bf5-e5697c987387/5927c2e7-RAMEN_FINMODEL_4_TABS_PRESENTATION.xlsx"
OUT=sys.argv[1]
NAVY='223A5E'; INK='172B46'; GREYT='5E6D7E'; KEY='E7EDF5'; BAND2='F2F5F9'; INP='FFF4D2'; BLUE='0000FF'; GREEN='008000'; AFTER='FCE3A6'
def font(sz=11,b=False,c=INK,i=False): return Font(name='Arial',size=sz,bold=b,color=c,italic=i)
def fill(c): return PatternFill('solid',fgColor=c)
thin=Side(style='thin',color='D5DCE8'); BOX=Border(left=thin,right=thin,top=thin,bottom=thin)
UAH='#,##0.00'; EUR='"€"0.000'; PCT='0.0%'; INT='#,##0'
wb=openpyxl.load_workbook(SRC)
CFG=[dict(name='Цена 120 грн — 6 паллет',src='Финмодель 6 паллет',head='6 паллет',freight=2600,pallets=6),
     dict(name='Цена 120 грн — 33 паллет',src='Финмодель 33 паллет',head='33 паллеты (Full Truck)',freight=3700,pallets=33)]
def build(cfg):
    idx=wb.sheetnames.index(cfg['name']); del wb[cfg['name']]
    ws=wb.create_sheet(cfg['name'],idx); S="'"+cfg['src']+"'"
    ws.sheet_properties.tabColor=NAVY; ws.sheet_view.showGridLines=False
    ws.column_dimensions['A'].width=2; ws.column_dimensions['B'].width=40
    for col in "CDEFGHIJKLMN": ws.column_dimensions[col].width=12
    ws.column_dimensions['O'].width=2
    CT=Alignment(horizontal='center',vertical='center',wrap_text=True); LF=Alignment(horizontal='left',vertical='center',wrap_text=True)
    def put(a,v=None,f=None,fl=None,fmt=None,al=None,border=False,merge=None):
        c=ws[a]; c.value=v; c.font=f or font(); c.alignment=al or CT
        if fl: c.fill=fill(fl)
        if fmt: c.number_format=fmt
        if border: c.border=BOX
        if merge:
            ws.merge_cells(merge)
            for row in ws[merge]:
                for cc in row:
                    if fl: cc.fill=fill(fl)
                    if border: cc.border=BOX
        return c
    ws.row_dimensions[1].height=8
    put("B2",f"SIAS · было → стало · от себестоимости в Украине до полки · {cfg['head']}",font(16,True),al=LF,merge="B2:N2"); ws.row_dimensions[2].height=28
    # видимые настройки: строка 14: C:D полка, E:F маржа, G:H бонус, I:J множитель, K:L курс
    SH,MG,BO,MK,RT="C14","E14","G14","I14","K14"
    # скрытые: строка 17: C:D ПДВ, E:F фин., G:H пошлина, I:J перевозка, K:L брокер, M:N прочие ; строка 19: C:D доля фрахта, E:F паллет, G:H пачек на паллете
    VT,FN,DT,FR,BR,OT="C17","E17","G17","I17","K17","M17"; FS,PL,PP="C19","E19","G19"
    BATCH=f"({FR}+{BR}+{OT}+{FR}*{FS}*({DT}+{VT}*(1+{DT})))"; PACKS=f"({PL}*{PP})"
    # скрытый расчёт запроса и СС «после»: C21 запрос, C22 предел, F22.. СС после ₴
    ASK="C21"
    put("B3",f'="Полка "&ROUND({SH},0)&" ₴ достижима при себестоимости "&ROUND(D9,2)&" ₴ (сейчас "&ROUND(D7,2)&"–"&ROUND(D8,2)&" ₴). Для этого закупка SIAS должна быть €"&ROUND({ASK},2)&" вместо €"&MIN(C20,E20)&"–€"&MAX(C20,E20)&" (−"&ROUND((1-{ASK}/MIN(C20,E20))*100,0)&"…"&ROUND((1-{ASK}/MAX(C20,E20))*100,0)&"%)."',font(11,True,NAVY),KEY,al=LF,merge="B3:N3"); ws.row_dimensions[3].height=30
    groups=["Себестоимость в Украине (СС)","Цена партнёру (магазину)","Бонус сети","Наша прибыль","Наценка магазина","ПОЛКА с НДС"]
    put("B5","",font(9,True,'FFFFFF'),'3E5A85'); put("B6","Вариант",font(9,True,'FFFFFF'),'3E5A85')
    for i,g in enumerate(groups):
        c1=CL(3+2*i); c2=CL(4+2*i)
        put(f"{c1}5",g,font(9,True,'FFFFFF'),'3E5A85',merge=f"{c1}5:{c2}5")
        put(f"{c1}6","€",font(9,True,'FFFFFF'),'5A74A0'); put(f"{c2}6","₴",font(9,True,'FFFFFF'),'5A74A0')
    ws.row_dimensions[5].height=32; ws.row_dimensions[6].height=18
    # колонки: СС C(€) D(₴); партнёр E F; бонус G H; прибыль I J; наценка K L; полка M N
    # NB: СС ₴ = D ; в формулах ниже используем D как базу
    rows=[(7,'="ДО · Vegetable, Spicy, Beef, Chicken · закупка SIAS €"&C20',f"={S}!E13",False),
          (8,'="ДО · Gouda Cheese, Carbonara Spicy · закупка SIAS €"&E20',f"={S}!E11",False),
          (9,'="ПОСЛЕ · запрос SIAS €"&ROUND(C21,2)&" (все 6 SKU)"',None,True)]
    for r,lab,cc_link,after in rows:
        fl=AFTER if after else (BAND2 if r==7 else None)
        put(f"B{r}",lab,font(10,True,NAVY if after else INK),fl,al=LF,border=True)
        if after:
            put(f"D{r}","=C23",font(12,True,BLUE if False else INK),fl,UAH,CT,border=True)
            put(f"C{r}",f"=D{r}/{RT}",font(10,False,GREYT),fl,EUR,CT,border=True)
        else:
            put(f"C{r}",cc_link,font(10,False,GREEN),fl or BAND2,EUR,CT,border=True); put(f"D{r}",f"=C{r}*{RT}",font(12,True),fl,UAH,CT,border=True)
        put(f"F{r}",f"=D{r}/(1-{BO}-{MG})",font(12,True),fl,UAH,CT,border=True)
        put(f"H{r}",f"=F{r}*{BO}",font(12),fl,UAH,CT,border=True)
        put(f"J{r}",f"=F{r}-D{r}-H{r}",font(12),fl,UAH,CT,border=True)
        put(f"L{r}",f"=F{r}*({MK}-1)",font(12),fl,UAH,CT,border=True)
        put(f"N{r}",f"=F{r}*{MK}",font(13,True),fl or KEY,UAH,CT,border=True)
        for ce,cu in (("E","F"),("G","H"),("I","J"),("K","L"),("M","N")):
            put(f"{ce}{r}",f"={cu}{r}/{RT}",font(10,False,GREYT),fl,EUR,CT,border=True)
        ws.row_dimensions[r].height=32
    put("B10","ИЗМЕНЕНИЕ: ПОСЛЕ − ДО (4 SKU)",font(10,True,NAVY),KEY,al=LF,border=True)
    for ci in range(3,15):
        col=CL(ci)
        put(f"{col}10",f"={col}9-{col}7",font(10,True,'B53A3A'),KEY,('+"€"0.000;-"€"0.000' if ci%2==1 else '+#,##0.00;-#,##0.00'),CT,border=True)
    ws.row_dimensions[10].height=24
    put("B11","Читайте по строке слева направо: себестоимость в Украине → цена партнёру = СС ÷ (1 − бонус − маржа) → бонус, наша прибыль, наценка → полка. Каждая цифра дана в € и ₴.",font(9,False,GREYT,True),al=LF,merge="B11:N11"); ws.row_dimensions[11].height=26
    # видимые настройки
    put("B12","ИСХОДНЫЕ ДАННЫЕ — жёлтое можно менять, зелёное из вкладки «"+cfg['src']+"»",font(9,True,NAVY),al=LF,merge="B12:N12")
    for c in "BCDEFGHIJKLMN": ws[f"{c}12"].border=Border(bottom=Side(style='medium',color=NAVY))
    P1=[("Полка, ₴",120,UAH,'inp'),("Наша маржа",0.30,PCT,'inp'),("Бонус сети",f"={S}!D4",PCT,'lnk'),("Множитель полки",f"={S}!D6",'0.00"×"','lnk'),("Курс ₴/€",f"={S}!D3",'0.00','lnk')]
    for i,(lab,val,nf,kind) in enumerate(P1):
        c1=CL(3+2*i); c2=CL(4+2*i)
        put(f"{c1}13",lab,font(8,True,GREYT),merge=f"{c1}13:{c2}13")
        put(f"{c1}14",val,font(11,True,BLUE if kind=='inp' else GREEN),INP if kind=='inp' else BAND2,nf,CT,border=True,merge=f"{c1}14:{c2}14")
    ws.row_dimensions[13].height=20; ws.row_dimensions[14].height=22
    put("B14","Цель расчёта и условия модели",font(9,True,GREYT),al=LF)
    put("B15","Допущения: бонус и множитель — по модели (не подтверждены сетями); без листинг-фи, промо и отсрочки; согласие SIAS на цену не получено — это цена для переговоров.",font(8,False,GREYT,True),al=LF,merge="B15:N15"); ws.row_dimensions[15].height=26
    # ---- скрытый расчёт (строки 16–23)
    P2=[("Импортный ПДВ",0.20,PCT),("Финансирование",0.025,PCT),("Пошлина",0,PCT),("Перевозка партии, €",cfg['freight'],INT),("Брокер, €",288,INT),("Прочие, €",42,INT)]
    for i,(lab,val,nf) in enumerate(P2):
        c1=CL(3+2*i); c2=CL(4+2*i)
        put(f"{c1}16",lab,font(8,True,GREYT),merge=f"{c1}16:{c2}16"); put(f"{c1}17",val,font(11,True,BLUE),INP,nf,CT,border=True,merge=f"{c1}17:{c2}17")
    P3=[("Доля фрахта в таможне",0.7,'0.00'),("Паллет в партии",cfg['pallets'],INT),("Пачек на паллете",1600,INT)]
    for i,(lab,val,nf) in enumerate(P3):
        c1=CL(3+2*i); c2=CL(4+2*i)
        put(f"{c1}18",lab,font(8,True,GREYT),merge=f"{c1}18:{c2}18"); put(f"{c1}19",val,font(11,True,BLUE),INP,nf,CT,border=True,merge=f"{c1}19:{c2}19")
    put("B20","Закупка SIAS сейчас, €: 4 SKU / 2 SKU",font(9,False,GREYT),al=LF); put("C20",0.65,font(11,True,BLUE),INP,'0.00',CT,border=True); put("E20",0.75,font(11,True,BLUE),INP,'0.00',CT,border=True)
    put("B21","Запрос SIAS, €/уп. (вниз до евроцента)",font(9,False,GREYT),al=LF)
    put("C21",f"=ROUNDDOWN((({SH}/{MK}*(1-{BO}-{MG}))/{RT}/(1+{FN})-{BATCH}/{PACKS})/((1+{DT})*(1+{VT})),2)",font(11,True),AFTER,'"€"0.00',CT,border=True,merge="C21:D21")
    put("B22","Предельная закупка до округления, €/уп.",font(9,False,GREYT),al=LF)
    put("C22",f"=(({SH}/{MK}*(1-{BO}-{MG}))/{RT}/(1+{FN})-{BATCH}/{PACKS})/((1+{DT})*(1+{VT}))",font(11),None,'"€"0.0000',CT,border=True,merge="C22:D22")
    put("B23","Себестоимость в Украине при запросе, ₴/уп.",font(9,False,GREYT),al=LF)
    put("C23",f"=(C21*{RT}+({FR}+{BR}+{OT})/{PACKS}*{RT}+(C21+{FR}*{FS}/{PACKS})*{DT}*{RT}+(C21+{FR}*{FS}/{PACKS})*(1+{DT})*{VT}*{RT})*(1+{FN})",font(11),None,UAH,CT,border=True,merge="C23:D23")
    for r in range(16,24): ws.row_dimensions[r].hidden=True
    ws.page_setup.orientation='landscape'; ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=1
    ws.sheet_properties.pageSetUpPr=PageSetupProperties(fitToPage=True)
# убрать пустые тёмные блоки-«призраки» (L2:M7, N2:O5) на вкладках 1–2: заливка 44546A без текста
from openpyxl.styles import PatternFill as _PF
for _n in wb.sheetnames[:2]:
    _ws=wb[_n]
    for _rng in [str(r) for r in _ws.merged_cells.ranges]:
        _cells=[c for row in _ws[_rng] for c in row]
        if all(c.value is None for c in _cells) and any(c.fill.fill_type=='solid' and c.fill.fgColor.type=='rgb' and c.fill.fgColor.rgb=='FF44546A' for c in _cells) and _cells[0].row<=8 and _cells[0].column>=12:
            _ws.unmerge_cells(_rng)
            for c in _cells: c.fill=_PF(fill_type=None)
for c in CFG: build(c)
wb.calculation=CalcProperties(fullCalcOnLoad=True)
wb.save(OUT); print("saved",OUT)
