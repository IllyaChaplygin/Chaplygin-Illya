# -*- coding: utf-8 -*-
import openpyxl,copy,sys
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
from openpyxl.workbook.properties import CalcProperties
SRC,OUT,LANG=sys.argv[1],sys.argv[2],sys.argv[3]
wb=openpyxl.load_workbook(SRC)
if LANG=='ua':
    PAIRS=((('Финмодель 6 паллет','Цена 120 грн — 6 паллет',20),('Финмодель 33 паллет','Цена 120 грн — 33 паллет',19)))
    L=dict(title='ДО → ПІСЛЯ · ціна на полиці при поточній закупівлі та при запиті SIAS',
      sub=lambda P:f'="Бонус мережі "&ROUND(\'{P}\'!D14*100,0)&" % · маржа "&ROUND(\'{P}\'!C14*100,0)&" % · націнка ×"&\'{P}\'!E14',
      g1='ДО · поточна закупівля SIAS',g2='ПІСЛЯ · запит SIAS (одна ціна на 6 SKU)',g3='ЗМІНА',
      h=['№','Позиція','Закупівля SIAS: до → після','Вага, г','Закупівля, €/пач.','СС, ₴/пач.','Ціна партнеру, ₴','Бонус мережі, ₴','Залишається після бонусу, ₴','Полиця з ПДВ, ₴','Закупівля, €/пач.','СС, ₴/пач.','Ціна партнеру, ₴','Бонус мережі, ₴','Залишається після бонусу, ₴','Полиця з ПДВ, ₴','Δ полиці, ₴','Δ полиці, %'],
      note='СС «до» — реальна (Self-Cost), «після» — розрахунок на запит SIAS (логістика рівномірно на пачку). Бонус, маржа та націнка — з аркуша «Цена 120 грн».')
else:
    PAIRS=((('Model 6 pallets','Price 120 UAH - 6 pallets',20),('Model 33 pallets','Price 120 UAH - 33 pallets',19)))
    L=dict(title='BEFORE → AFTER · shelf price at the current purchase price and at the SIAS request',
      sub=lambda P:f'="Chain bonus "&ROUND(\'{P}\'!D14*100,0)&" % · margin "&ROUND(\'{P}\'!C14*100,0)&" % · markup ×"&\'{P}\'!E14',
      g1='BEFORE · current SIAS purchase price',g2='AFTER · SIAS request (one price for all 6 SKUs)',g3='CHANGE',
      h=['#','Item','SIAS purchase: before → after','Weight, g','Purchase, €/pack','Landed cost, ₴/pack','Partner price, ₴','Chain bonus, ₴','Left after bonus, ₴','Shelf price incl. VAT, ₴','Purchase, €/pack','Landed cost, ₴/pack','Partner price, ₴','Chain bonus, ₴','Left after bonus, ₴','Shelf price incl. VAT, ₴','Δ shelf, ₴','Δ shelf, %'],
      note='"Before" landed cost per Self-Cost (incl. the uplift from G14 on the "Price 120 UAH" sheets); "after" is calculated at the SIAS request (logistics spread evenly per pack). Bonus, margin and markup come from the "Price 120 UAH" sheets.')
thin=Side(style='thin',color='C9D3E0'); BR=Border(left=thin,right=thin,top=thin,bottom=thin)
def st(c,sz=10,b=False,col='FF222222',fill=None,fmt=None,al='center'):
    c.font=Font(name='Arial',sz=sz,b=b,color=col)
    if fill: c.fill=PatternFill('solid',fgColor=fill)
    if fmt: c.number_format=fmt
    c.alignment=Alignment(horizontal=al,vertical='center',wrap_text=True); c.border=BR
EUR='"€"0.000'; UAH='#,##0.00" ₴"'; DUAH='+#,##0.00" ₴";-#,##0.00" ₴"'; DPCT='+0.0%;-0.0%'
for mn,pn,ar in PAIRS:
    ws=wb[mn]; P=f"'{pn}'!"
    # единые условия: маржа на аркуші моделі = маржа з аркуша «Цена 120 грн»
    if LANG=='ua':
        ws['D5']=f"={P}C14"
        ws['E4']='="Бонус мережі "&ROUND(D4*100,0)&"%; полиця = ціна партнера × "&D6&"."'
        ws['E5']='=ROUND(D5*100,0)&"% = (ціна партнеру − собівартість − бонус) / повна ціна партнеру. Береться з аркуша «Цена 120 грн»."'
        ws['E8']='="100% − бонус − маржа. За поточних умов: 100% − "&ROUND(D4*100,0)&"% − "&ROUND(D5*100,0)&"% = "&ROUND(D8*100,0)&"%."'
    r0=42
    ws.merge_cells(start_row=r0,start_column=2,end_row=r0,end_column=18); c=ws.cell(r0,2,L['title']); c.font=Font(name='Arial',sz=14,b=True,color='FF1F3864'); c.alignment=Alignment(horizontal='left',vertical='center'); ws.row_dimensions[r0].height=26
    ws.merge_cells(start_row=r0+1,start_column=2,end_row=r0+1,end_column=18); c=ws.cell(r0+1,2,L['sub'](pn)); c.font=Font(name='Arial',sz=10,color='FF5E6D7E'); c.alignment=Alignment(horizontal='left',vertical='center'); ws.row_dimensions[r0+1].height=20
    g=r0+2
    for (c1,c2,txt,fill) in ((5,10,L['g1'],'FF203864'),(11,16,L['g2'],'FFB07A12'),(17,18,L['g3'],'FFA33A3A')):
        ws.merge_cells(start_row=g,start_column=c1,end_row=g,end_column=c2)
        for cc in range(c1,c2+1): st(ws.cell(g,cc),10,True,'FFFFFFFF',fill)
        ws.cell(g,c1).value=txt
    ws.row_dimensions[g].height=26
    h=g+1
    for i,t in enumerate(L['h'],1):
        fill='FF44546A' if i<=4 else ('FF2F75B5' if i<=10 else ('FFD9A441' if i<=16 else 'FFC26A6A'))
        st(ws.cell(h,i,t),9,True,'FFFFFFFF',fill)
    ws.row_dimensions[h].height=42
    for k,r in enumerate(range(11,17)):
        o=h+1+k; two=(r in (11,12))
        buy_before=f"={P}$B$7" if two else f"={P}$B$5"
        vals={1:f"=A{r}",2:f"=C{r}",4:f"=D{r}",5:buy_before,6:f"=F{r}",7:f"=G{r}",8:f"=H{r}",9:f"=I{r}",10:f"=K{r}",
              11:f"={P}$B${ar}",12:f"={P}$C$8",13:f"=L{o}/(1-$D$4-$D$5)",14:f"=M{o}*$D$4",15:f"=M{o}-L{o}-N{o}",16:f"=M{o}*$D$6",17:f"=P{o}-J{o}",18:f"=Q{o}/J{o}"}
        vals[3]=f'="€"&ROUND(E{o},2)&" → €"&ROUND(K{o},2)'
        for col,f in vals.items():
            c=ws.cell(o,col,f)
            if col<=4: st(c,10,col==2,'FF203864','FFF2F5F9',None,'left' if col==2 else 'center')
            elif col<=10: st(c,11,col==10,'FF203864' if col==10 else 'FF222222','FFD6E4F0' if col==10 else 'FFF5F8FC',EUR if col==5 else UAH)
            elif col<=16: st(c,11,col==16,'FF203864' if col==16 else 'FF222222','FFFCE3A6' if col==16 else 'FFFEF4DC',EUR if col==11 else UAH)
            else: st(c,11,True,'FFA33A3A','FFFBE9E9',DUAH if col==17 else DPCT)
        ws.row_dimensions[o].height=30
    n=h+8
    ws.merge_cells(start_row=n,start_column=2,end_row=n,end_column=18); c=ws.cell(n,2,L['note']); c.font=Font(name='Arial',sz=9,color='FF5E6D7E'); c.alignment=Alignment(wrap_text=True,vertical='top'); ws.row_dimensions[n].height=30
wb.calculation=CalcProperties(fullCalcOnLoad=True); wb.save(OUT); print('saved',OUT)
