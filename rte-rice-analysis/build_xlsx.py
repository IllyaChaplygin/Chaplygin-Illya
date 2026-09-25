# -*- coding: utf-8 -*-
import data as d
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

INK="FF1C2235"; INK2="FF303A5D"; MUT="FF717890"; BG="FFF1F4FA"; BRAND="FFF9A50B"
TEAL="FF1E9EA6"; PINK="FFC94F7C"; GREEN="FF37A169"; AMB="FFC4820A"; PUR="FF5D4B96"
F=lambda **k: Font(name="Segoe UI",**k)
thin=Side(style="thin",color="FFD5DBEA")
BOX=Border(bottom=thin)

wb=Workbook(); wb.remove(wb.active)

def head(ws,row,cols,widths,fill=INK2):
    for i,(c,w) in enumerate(zip(cols,widths),1):
        cell=ws.cell(row,i,c); cell.font=F(sz=9,bold=True,color="FFFFFFFF")
        cell.fill=PatternFill("solid",fgColor=fill)
        cell.alignment=Alignment(vertical="center",wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width=w
    ws.row_dimensions[row].height=30
    ws.freeze_panes=ws.cell(row+1,1)

def title(ws,t,sub):
    ws["A1"]=t; ws["A1"].font=F(sz=14,bold=True,color=INK)
    ws["A2"]=sub; ws["A2"].font=F(sz=9,color=MUT)
    ws.row_dimensions[1].height=20

def zebra(ws,r0,r1,ncol):
    for r in range(r0,r1):
        if (r-r0)%2: 
            for c in range(1,ncol+1): ws.cell(r,c).fill=PatternFill("solid",fgColor=BG)

# ───────────── 1. БРЕНДИ ─────────────
ws=wb.create_sheet("1. Бренди")
title(ws,"Матриця брендів — ринок готового рису України",
      "23 бренди · 84 позиції · зріз 22–23.09.2026. Ціни роздрібні за одну упаковку, без оптових порогів.")
cols=["Бренд","Країна","Технологія","Формат","SKU","Ціна min, грн","Ціна max, грн",
      "Медіана, грн","Вага min, г","Вага max, г","грн/100 г min","грн/100 г max","Канал"]
head(ws,4,cols,[20,10,15,13,6,12,12,12,12,12,12,12,30])
r=5
for b in sorted(d.BRANDS,key=lambda x:x[7]):
    br,co,te,fm,sku,lo,hi,md,wl,wh,pl,ph,ch=b
    vals=[br,co,te,fm,sku,lo,hi,md,wl or None,wh or None,pl or None,ph or None,ch]
    for i,v in enumerate(vals,1):
        c=ws.cell(r,i,v); c.font=F(sz=9,color=INK2,bold=(i==1)); c.border=BOX
        if i in (6,7,8,9,10,11,12): c.number_format="# ##0"
    r+=1
zebra(ws,5,r,13)
ws.auto_filter.ref=f"A4:M{r-1}"

# ───────────── 2. ФОРМАТИ ─────────────
ws=wb.create_sheet("2. Формати")
title(ws,"Аналітика форматів упаковки",
      "Розподіл 84 позицій. грн/100 г = медіана ціни формату / медіану ваги формату.")
head(ws,4,["Формат","Позицій","Частка позицій","Брендів","Медіана ціни, грн",
           "Ціна min","Ціна max","Медіана ваги, г","грн/100 г","Коментар"],
     [22,10,14,10,17,11,11,16,12,52])
notes=["Формат мережевої полиці. Найнижча ціна входу в категорію.",
       "Формат азійського каналу. Прямі конкуренти Ottogi і Henan.",
       "Формат туристичного каналу. Потребує окропу, не полиця.",
       "Найдорожчий формат. Купується як подарунок або в дорогу."]
r=5
for (n,pos,sh,br,med,lo,hi,w,col),note in zip(d.FORMATS,notes):
    ws.cell(r,1,n).font=F(sz=9,bold=True,color=INK)
    for i,v in enumerate([pos,sh/100,br,med,lo,hi,w],2):
        c=ws.cell(r,i,v); c.font=F(sz=9,color=INK2)
        c.number_format="0 %" if i==3 else "# ##0"
    c=ws.cell(r,9,f"=E{r}/H{r}*100"); c.font=F(sz=9,bold=True,color=INK); c.number_format="# ##0"
    ws.cell(r,10,note).font=F(sz=9,color=MUT)
    for i in range(1,11): ws.cell(r,i).border=BOX
    r+=1
ws.cell(r+1,1,"Висновок").font=F(sz=10,bold=True,color=INK)
ws.cell(r+2,1,"Дойпак лідирує за кількістю позицій (36 %), але це артефакт туристичного каналу. "
        "За ціною для покупця виграє пауч: 109 грн проти 315 грн.").font=F(sz=9,color=INK2)

# ───────────── 3. НАША ЛІНІЙКА vs РИНОК ─────────────
ws=wb.create_sheet("3. Наша лінійка")
title(ws,"Наша лінійка проти ринку — 32 SKU",
      "Полиця розрахована за фінмоделлю. Жовті клітинки на аркуші «5. Калькулятор» керують розрахунком.")
head(ws,4,["Група","SKU","Вага, г","Постачальник","Формат","СС 20′ FCL, $","СС 40′ FCL, $",
           "СС 20′, грн","Ціна партнеру, грн","Бонус мережі, грн","Полиця 20′, грн",
           "Полиця 40′, грн","грн/100 г","Ринковий орієнтир","Позиція"],
     [30,30,9,13,10,13,13,12,17,16,14,14,12,26,26])
MK="'5. Калькулятор'!"
r=5
for g,sku,wt,c20,c40,sup,fmt in d.OWN:
    ref = {"Пауч":"Ben's 45–179 · медіана 90","Чаша":"Ottogi 225–356 · Henan 83–156",
           "Лоток":"Bibigo 135–189"}[fmt]
    for i,v in enumerate([g,sku,wt,sup,fmt,c20,c40],1):
        c=ws.cell(r,i,v); c.font=F(sz=9,color=INK2,bold=(i==2))
        if i in (6,7): c.number_format="0.000"
        if i==3: c.number_format="# ##0"
    ws.cell(r,8,f"=F{r}*{MK}$B$4").number_format="# ##0.00"
    ws.cell(r,9,f"=H{r}/{MK}$B$8").number_format="# ##0.00"
    ws.cell(r,10,f"=I{r}*{MK}$B$5").number_format="# ##0.00"
    ws.cell(r,11,f"=I{r}*{MK}$B$7").number_format="# ##0"
    ws.cell(r,12,f"=G{r}*{MK}$B$4/{MK}$B$8*{MK}$B$7").number_format="# ##0"
    ws.cell(r,13,f"=K{r}/C{r}*100").number_format="# ##0"
    ws.cell(r,14,ref)
    ws.cell(r,15,f'=IF(K{r}<=180,IF(K{r}<=109,"у низу вікна полиці","у вікні полиці 45–180"),"вище вікна полиці")')
    for i in range(8,16):
        ws.cell(r,i).font=F(sz=9,color=INK2,bold=(i in (11,13)))
    for i in range(1,16): ws.cell(r,i).border=BOX
    r+=1
zebra(ws,5,r,15)
ws.conditional_formatting.add(f"K5:K{r-1}",
    CellIsRule(operator="greaterThan",formula=["180"],
               fill=PatternFill("solid",fgColor="FFFBE4EC"),font=F(sz=9,bold=True,color=PINK)))
ws.conditional_formatting.add(f"K5:K{r-1}",
    CellIsRule(operator="lessThanOrEqual",formula=["140"],
               fill=PatternFill("solid",fgColor="FFE8F6EE"),font=F(sz=9,bold=True,color=GREEN)))
ws.auto_filter.ref=f"A4:O{r-1}"

# ───────────── 4. КАНАЛИ ─────────────
ws=wb.create_sheet("4. Канали")
title(ws,"Карта каналів — 54 ідентифіковані продавці",
      "Розширено відносно першого зрізу. Туристичний канал перелічено поіменно частково (28 із ~40).")
head(ws,4,["Канал","Продавець","Які бренди тут стоять","Статус для нас"],[34,30,46,30])
r=5
STAT={"Мережевий роздріб (офлайн+онлайн)":"Ціль 2-ї хвилі",
      "Онлайн-супермаркети й маркетплейси":"Ціль 1-ї хвилі",
      "Азійські / етнічні фудшопи":"ЦІЛЬ ПЕРШОГО ВХОДУ",
      "Туристичні та військові":"Не наш привід споживання"}
for ch,v in d.CHANNELS.items():
    for name in v["found"]:
        ws.cell(r,1,ch).font=F(sz=9,color=MUT)
        ws.cell(r,2,name).font=F(sz=9,bold=True,color=INK)
        ws.cell(r,3,v["note"]).font=F(sz=9,color=INK2)
        c=ws.cell(r,4,STAT[ch]); c.font=F(sz=9,bold=("ЦІЛЬ ПЕРШОГО" in STAT[ch]),
                                          color=TEAL if "ПЕРШОГО" in STAT[ch] else MUT)
        for i in range(1,5): ws.cell(r,i).border=BOX
        r+=1
r+=1
ws.cell(r,1,"Мережі без жодної позиції категорії").font=F(sz=10,bold=True,color=INK); r+=1
ws.cell(r,1,", ".join(d.CHANNELS["Мережевий роздріб (офлайн+онлайн)"]["empty"])).font=F(sz=9,color=MUT)
ws.auto_filter.ref="A4:D4"

# ───────────── 5. КАЛЬКУЛЯТОР ─────────────
ws=wb.create_sheet("5. Калькулятор")
title(ws,"Калькулятор цілі: яка ціна потрібна від постачальника",
      "Змінюйте жовті клітинки. Усі інші аркуші перераховуються автоматично.")
YEL=PatternFill("solid",fgColor="FFFFF3CD")
par=[("Курс, грн/$",45.0,"0.00"),("Бонус мережі, %",0.25,"0 %"),("Цільова маржа, %",0.35,"0 %"),
     ("Націнка мережі (x)",1.40,"0.00"),("Частка ціни для собівартості",None,"0 %"),
     ("Коеф. FOB → СС, 20′ FCL",d.K_LANDED_20,"0.0000"),
     ("Коеф. FOB → СС, 40′ FCL",d.K_LANDED_40,"0.0000")]
r=4
for n,v,fmt in par:
    ws.cell(r,1,n).font=F(sz=10,color=INK2)
    c=ws.cell(r,2, v if v is not None else "=1-B5-B6")
    c.font=F(sz=10,bold=True,color=INK); c.number_format=fmt
    if v is not None and r<=7: c.fill=YEL
    r+=1
ws.column_dimensions["A"].width=38; ws.column_dimensions["B"].width=14
ws.cell(11,1,"Полиця = СС_грн × націнка / частка_СС").font=F(sz=9,italic=True,color=MUT)
ws.cell(12,1,"Множник СС → полиця").font=F(sz=10,color=INK2)
c=ws.cell(12,2,"=B7/B8"); c.font=F(sz=10,bold=True,color=INK); c.number_format="0.00"
ws.cell(13,1,"Множник FOB → полиця, 20′ FCL").font=F(sz=10,color=INK2)
c=ws.cell(13,2,"=B9*B4*B12"); c.font=F(sz=10,bold=True,color=INK); c.number_format="0.0"
ws.cell(14,1,"Множник FOB → полиця, 40′ FCL").font=F(sz=10,color=INK2)
c=ws.cell(14,2,"=B10*B4*B12"); c.font=F(sz=10,bold=True,color=INK); c.number_format="0.0"

head(ws,17,["Цільова полиця, грн","Що це на ринку","Потрібна СС, грн","Потрібна СС, $",
            "Потрібен FOB 20′, $","Потрібен FOB 40′, $"],[20,40,18,16,20,20])
tg=[(98.4,"Ben's Original медіана 41 грн/100 г при 240 г"),
    (109,"Медіана формату «пауч» — найдешевший вхід у мережу"),
    (119,"Агресивний сценарій для стакана 150 г"),
    (129,"РЕКОМЕНДОВАНО для стакана 150 г — нижче Henan-медіани"),
    (139,"Поточна полиця стакана 150 г без переговорів"),
    (155,"Паритет із Bibigo 210 г (135–189 грн)"),
    (179,"Верх вікна мережевої полиці"),
    (199,"Ottogi −22 % — стеля для лотка ONE'S 200 г")]
r=18
for v,note in tg:
    c=ws.cell(r,1,v); c.font=F(sz=10,bold=True,color=INK); c.number_format="# ##0"; c.fill=YEL
    ws.cell(r,2,note).font=F(sz=9,color=MUT)
    for i,f in enumerate([f"=A{r}/$B$12",f"=C{r}/$B$4",f"=A{r}/$B$13",f"=A{r}/$B$14"],3):
        cc=ws.cell(r,i,f); cc.font=F(sz=10,bold=(i>=5),color=INK2)
        cc.number_format="# ##0.00" if i==3 else "0.000"
    for i in range(1,7): ws.cell(r,i).border=BOX
    r+=1
zebra(ws,18,r,6)
ws.cell(r+1,1,"Правило для переговорів").font=F(sz=10,bold=True,color=INK)
ws.cell(r+2,1,"Кожні $0,10 зниження FOB знімають ~27 грн з полиці (20′ FCL). "
        "Перехід із 20′ на 40′ FCL дає ~−11 % собівартості без переговорів із постачальником.").font=F(sz=9,color=INK2)

# ───────────── 6. ДЖЕРЕЛА І ЗАСТЕРЕЖЕННЯ ─────────────
ws=wb.create_sheet("6. Застереження")
title(ws,"Що тут факт, що припущення, а що оцінка","Читати до використання цифр у переговорах.")
head(ws,4,["Показник","Статус","Джерело / метод","Вплив, якщо хибне"],[34,16,52,52])
rows=[("Ціни конкурентів на полиці","ФАКТ","Картки продавців, 17 мереж zakaz.ua, «Сільпо», Prom, Rozetka, MAUDAU, 22–23.09.2026","—"),
 ("Формати, ваги, кількість SKU","ФАКТ","Картки продавців. 2 позиції без маси виключено з грн/100 г","Незначний"),
 ("Обсяг імпорту 182 т","ФАКТ","Eurostat Comext CN 1904 90 10 + UN Comtrade","—"),
 ("Ємність роздрібу 87–110 млн грн","ОЦІНКА","Імпорт × 2,2–2,8. Заміряних продажів немає","Похибка ±30 % у бізнес-кейсі"),
 ("Собівартість BSCM і CM Premium","ФАКТ","Комерційні пропозиції, RTE_Rice_Suppliers.pptx","—"),
 ("FOB ONE'S International","ОЦІНКА","Виведено з СС за коефіцієнтом 1,7074. Постачальник FOB не розкрив","Рекомендація по Fried Rice 200 г може бути хибною"),
 ("Коеф. FOB → СС = 1,7074","РОЗРАХУНОК","СС / FOB для BSCM Jasmine 240 г, перевірено на 15 SKU","Зсув усіх цільових FOB"),
 ("Мито 10 % у собівартості","ПРИПУЩЕННЯ","Зазначено в RTE_Rice_Suppliers. Первинне джерело не наведено","±27 грн на полиці"),
 ("Імпортний ПДВ 20 % у СС","РИЗИК","Ймовірне задвоєння: ПДВ і у витратах, і в ціні «вже з ПДВ»","СС −16,7 %, полиця пауча 157 → 131 грн"),
 ("Бонус мережі 25 %","ПРИПУЩЕННЯ","Взято з прикладу по локшині. Умови «Сільпо» не запитані","За 30 % полиця +17 %"),
 ("«Маржа 35 %»","ВИЗНАЧЕННЯ","Валова націнка. Не містить ФОП, маркетингу, лістингу, списань","Не є чистою рентабельністю"),
 ("Оборотність Ben's у «Сільпо»","НЕМАЄ ДАНИХ","Мережа обсягів не віддає","Неможливо порахувати план продажів")]
r=5
CL={"ФАКТ":GREEN,"ОЦІНКА":AMB,"ПРИПУЩЕННЯ":PUR,"РИЗИК":PINK,"РОЗРАХУНОК":TEAL,"НЕМАЄ ДАНИХ":MUT,"ВИЗНАЧЕННЯ":INK2}
for n,st,src,imp in rows:
    ws.cell(r,1,n).font=F(sz=9,bold=True,color=INK)
    c=ws.cell(r,2,st); c.font=F(sz=9,bold=True,color=CL[st])
    ws.cell(r,3,src).font=F(sz=9,color=INK2)
    ws.cell(r,4,imp).font=F(sz=9,color=MUT)
    for i in range(1,5):
        ws.cell(r,i).border=BOX; ws.cell(r,i).alignment=Alignment(wrap_text=True,vertical="top")
    ws.row_dimensions[r].height=28
    r+=1
zebra(ws,5,r,4)

wb.save("out/RTE_Rice_Benchmark.xlsx")
print("saved, sheets:",wb.sheetnames)
