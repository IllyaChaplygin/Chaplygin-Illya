# -*- coding: utf-8 -*-
import openpyxl,copy,sys
from openpyxl.styles import Font
from openpyxl.workbook.properties import CalcProperties
SRC="/home/user/Chaplygin-Illya/RAMEN_FINMODEL_INTERNAL_UA.xlsx"; OUT=sys.argv[1]
wb=openpyxl.load_workbook(SRC); wv=openpyxl.load_workbook(SRC,data_only=True)
for n in ('Финмодель 6 паллет','Финмодель 33 паллет','Сценарии'): del wb[n]
CONST={}
CFG=(('Цена 120 грн — 33 паллет','33 pallets',33,19,'33 pallets (Full Truck)'),('Цена 120 грн — 6 паллет','6 pallets',6,20,'6 pallets'))
for old,new,pal,ar,ttl in CFG:
    ws=wb[old]; v=wv[old]; cc=ar+2; r1=cc+2; r2=cc+3
    ws.title=new
    # константы вместо ссылок на внутренние вкладки
    ws['D14']=v['D14'].value; ws['E14']=v['E14'].value; ws['F14']=v['F14'].value
    for a in ('D14','E14','F14'): ws[a]._style=copy.copy(ws['C14']._style)
    ws['D14'].number_format='0.0%'; ws['E14'].number_format='0.00"×"'; ws['F14'].number_format='0.00'
    ws['B%d'%r1]=v['B%d'%r1].value; ws['B%d'%r2]=v['B%d'%r2].value
    for a in ('B%d'%r1,'B%d'%r2):
        f=ws[a].font; ws[a].font=Font(name=f.name,sz=f.sz,b=f.b,color='FF0000FF'); ws[a].fill=copy.copy(ws['C14'].fill)
    # убираем доп. расходы (внутренний параметр)
    for a in ('G13','G14'): ws[a].value=None; ws[a]._style=copy.copy(ws['H13']._style)
    ws['C4']=f"=$B${r1}*F14"; ws['C6']=f"=$B${r2}*F14"
    vat,fin,dut,fr,br,ot,sh,pl,pk=[v[a].value for a in ('B16','C16','D16','E16','F16','G16','B18','C18','D18')]
    n_=pl*pk
    batch=f"({fr}+{br}+{ot}+{fr}*{sh}*({dut}+{vat}*(1+{dut})))/{n_}"
    core=f"((B14/E14*(1-D14-C14))/F14/(1+{fin})-{batch})/((1+{dut})*(1+{vat}))"
    ws[f'B{ar}']=f"=ROUNDDOWN({core},2)"; ws[f'B{ar+1}']=f"={core}"
    ws[f'B{cc}']=f"=(B{ar}*F14+({fr}+{br}+{ot})/{n_}*F14+(B{ar}+{fr}*{sh}/{n_})*{dut}*F14+(B{ar}+{fr}*{sh}/{n_})*(1+{dut})*{vat}*F14)*(1+{fin})"
    CONST[new]=(vat,fin,dut,fr,br,ot,sh,n_)
    # английские подписи
    T={'A1':f'SIAS · from purchase price to shelf price, one line · {ttl}','A2':'REQUEST TO SIAS, €/pack','H2':'one price for all 6 SKUs','A3':'Scenario',
       'B3':'SIAS purchase\nper pack','C3':'Landed cost in Ukraine\nproduct + costs in Ukraine','D3':'Partner price\ncost ÷ (1 − bonus − our share)',
       'G3':'Store markup\n+40 % on partner price','H3':'SHELF PRICE incl. VAT\npartner price × 1.40',
       'A4':'BEFORE · Vegetable, Spicy, Beef, Chicken\n(4 SKUs)','A6':'BEFORE · Gouda Cheese, Carbonara Spicy\n(2 SKUs)','A8':'AFTER · SIAS request\n(all 6 SKUs)','A10':'CHANGE: AFTER − BEFORE (4 SKUs)',
       'A12':'INPUTS — yellow cells can be changed','B13':'Shelf price, ₴','C13':'Our share of partner price','D13':'Chain bonus','E13':'Shelf multiplier','F13':'Rate ₴/€','A14':'Target and model terms',
       'B15':'Import VAT','C15':'Financing','D15':'Customs duty','E15':'Batch freight, €','F15':'Broker, €','G15':'Other, €','B17':'Freight share in customs value','C17':'Pallets in batch','D17':'Packs per pallet',
       f'A{ar}':'Request to SIAS, €/pack (rounded down to the euro cent)',f'A{ar+1}':'Maximum purchase price before rounding, €/pack',f'A{cc}':'Landed cost in Ukraine at the request, ₴/pack',
       f'A{r1}':'Landed cost today, 4 SKUs (Vegetable, Spicy, Beef, Chicken), €/pack',f'A{r2}':'Landed cost today, 2 SKUs (Gouda, Carbonara), €/pack'}
    for k,t in T.items(): ws[k]=t
    for rr_ in (15,16,17,18):
        for cc_ in 'ABCDEFGH':
            ws[f'{cc_}{rr_}'].value=None; ws[f'{cc_}{rr_}']._style=copy.copy(ws['I1']._style)
        ws.row_dimensions[rr_].height=3
    ws.row_dimensions[13].height=26
    ws['D2']='="now €"&B5&" (4 SKUs) / €"&B7&" (2 SKUs)"'
    ws['F2']=f'="reduction −"&ROUND((1-B{ar}/B5)*100,0)&" % / −"&ROUND((1-B{ar}/B7)*100,0)&" %"'
    ws['E3']='="Chain bonus"&CHAR(10)&ROUND(D14*100,0)&" % of partner price"'
    ws['F3']='="Left for us"&CHAR(10)&"costs & margin"'
    for a in ('B4','B6','B8'): pass
# проверка: нет ссылок на удалённые вкладки
for ws in wb:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value,str) and c.value.startswith('=') and ('Финмодель' in c.value or 'Сценарии' in c.value): print('BAD REF',ws.title,c.coordinate,c.value)
            if isinstance(c.value,str) and not c.value.startswith('=') and any('а'<=ch<='я' or 'А'<=ch<='Я' for ch in c.value): print('CYR',ws.title,c.coordinate,c.value[:60])

# ---------- лист «Landed cost blocks» ----------
import json
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
BL=json.load(open("/tmp/claude-0/-home-user-Chaplygin-Illya/8bc0034a-4b53-5022-9bf5-e5697c987387/scratchpad/v10/blocks.json"))
bs=wb.create_sheet('Landed cost blocks',0)
thin=Side(style='thin',color='C9D3E0'); BR=Border(left=thin,right=thin,top=thin,bottom=thin)
def st(c,sz=10,b=False,col='FF172B46',fill=None,fmt=None,al='center'):
    c.font=Font(name='Arial',sz=sz,b=b,color=col)
    if fill: c.fill=PatternFill('solid',fgColor=fill)
    if fmt: c.number_format=fmt
    c.alignment=Alignment(horizontal=al,vertical='center',wrap_text=True); c.border=BR
bs.sheet_view.showGridLines=False
bs.merge_cells('A1:H1'); bs['A1']='What the landed cost in Ukraine is made of · ₴ per pack'; bs['A1'].font=Font(name='Arial',sz=16,b=True,color='FF172B46'); bs.row_dimensions[1].height=26
bs.merge_cells('A2:H2'); bs['A2']='SIAS purchase + delivery + import VAT + financing = landed cost. Delivery = freight, broker, terminal and certification of the batch, per pack. Customs duty 0 %.'
bs['A2'].font=Font(name='Arial',sz=10,color='FF5E6D7E'); bs['A2'].alignment=Alignment(wrap_text=True,vertical='center'); bs.row_dimensions[2].height=30
HEAD=['','SIAS purchase\npack price × rate','Delivery\nfreight, broker, terminal','Import VAT\n20 % on goods and freight','Financing\n2.5 % of the sum','Landed cost, ₴ (sum of blocks)','Landed cost, €','Check vs model sheet']
r=4
for sh,key,pal,fr_ask,fr_cc in (('33 pallets','33',33,'B19','B21'),('6 pallets','6',6,'B20','B22')):
    bs.cell(r,1,f'{pal} PALLETS' + (' · FULL TRUCK' if pal==33 else '')); bs.cell(r,1).font=Font(name='Arial',sz=11,b=True,color='FF223A5E'); r+=1
    for i,h in enumerate(HEAD,1):
        c=bs.cell(r,i,h); st(c,9,True,'FFFFFFFF','3E5A85')
    bs.row_dimensions[r].height=44; r+=1
    T=f"'{sh}'!"
    rows=[('BEFORE · 4 SKUs\n(Vegetable, Spicy, Beef, Chicken)',BL[key]['before4'],f"{T}C4"),('BEFORE · 2 SKUs\n(Gouda, Carbonara)',BL[key]['before2'],f"{T}C6"),('AFTER · SIAS request\n(all 6 SKUs)',None,f"{T}C8")]
    for nm,vals,ref in rows:
        bs.cell(r,1,nm); st(bs.cell(r,1),10,True,'FF172B46','FCE3A6' if vals is None else 'F2F5F9',None,'left')
        if vals is not None:
            for j in range(4):
                bs.cell(r,2+j,vals[j]); st(bs.cell(r,2+j),11,False,'FF0000FF','FFF4D2','#,##0.00" ₴"')
        else:
            ask=f"{T}{fr_ask}"; vat_,fin_,dut_,fr_,br_,ot_,sh_,n_=CONST[sh]
            bs.cell(r,2,f"={ask}*{T}F14")
            bs.cell(r,3,f"=({fr_}+{br_}+{ot_})/{n_}*{T}F14")
            bs.cell(r,4,f"=({ask}+{fr_}*{sh_}/{n_})*(1+{dut_})*{vat_}*{T}F14")
            bs.cell(r,5,f"=(B{r}+C{r}+D{r}+({ask}+{fr_}*{sh_}/{n_})*{dut_}*{T}F14)*{fin_}")
            for j in range(4): st(bs.cell(r,2+j),11,False,'FF172B46','FCE3A6','#,##0.00" ₴"')
        bs.cell(r,6,f"=SUM(B{r}:E{r})"); st(bs.cell(r,6),13,True,'FF172B46','E7EDF5' if vals is not None else 'FCE3A6','#,##0.00" ₴"')
        bs.cell(r,7,f"=F{r}/{T}F14"); st(bs.cell(r,7),10,False,'FF5E6D7E','F2F5F9','"€"0.000')
        bs.cell(r,8,f"=ROUND(F{r}-{ref},4)"); st(bs.cell(r,8),9,False,'FF5E6D7E','F2F5F9','0.00;-0.00;"OK"')
        bs.row_dimensions[r].height=34; r+=1
    r+=1
bs.merge_cells(start_row=r,start_column=1,end_row=r,end_column=8)
bs.cell(r,1,'"Before": per our cost file (logistics allocated by value). "After": at the requested price, logistics spread evenly per pack. The smaller the batch, the higher the delivery cost per pack.')
bs.cell(r,1).font=Font(name='Arial',sz=9,color='FF5E6D7E'); bs.cell(r,1).alignment=Alignment(wrap_text=True,vertical='top'); bs.row_dimensions[r].height=28
bs.column_dimensions['A'].width=38
for col in 'BCDEFGH': bs.column_dimensions[col].width=17
wb.properties.title='SIAS financial model'; wb.properties.creator=''; wb.properties.lastModifiedBy=''; wb.properties.description=None; wb.properties.keywords=None
wb.calculation=CalcProperties(fullCalcOnLoad=True)
wb.save(OUT); print('saved',OUT)
