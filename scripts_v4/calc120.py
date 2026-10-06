# -*- coding: utf-8 -*-
"""Відтворення вкладок 3 і 4 фінмоделі (RAMEN_FINMODEL_4_TABS_PRESENTATION.xlsx) та їх чутливостей."""
import math
RATE=51.2; BONUS=0.25; MARKUP=1.40; TARGET=120.0; MARGIN=0.30
FIN=0.025; VAT=0.20; DUTY=0.0; FRT_BASE=0.70
BROKER=288.0; OTHER=42.0; PACK_PALLET=1600; PACK_BOX=20
SKUS=[("Gouda Cheese","123 г",123,0.75),("Carbonara Spicy","131 г",131,0.75),
      ("Vegetable Flavor","113 г",113,0.65),("Spicy Flavor","112,5 г",112.5,0.65),
      ("Beef Flavor","113 г",113,0.65),("Chicken Flavor","113 г",113,0.65)]
# Self-Cost_Sias.xlsx, колонка AJ (EUR/уп.) — «СС зараз»
SC={"6":{"freight":2600.0,"pallets":[1,1,1,1,1,1],"cc":[1.308515625,1.308515625]+[1.1340468750000001]*4},
    "33":{"freight":3700.0,"pallets":[7,7,4,5,5,5],"cc":[1.0181311542669584]*2+[0.8823803336980306]*4}}
LAB={"6":"6 палет","33":"33 палети · повна машина"}

def units(sc): return sum(SC[sc]["pallets"])*PACK_PALLET
def batch_costs(sc):
    f=SC[sc]["freight"]
    return f+BROKER+OTHER+f*FRT_BASE*(DUTY+VAT*(1+DUTY))
def max_cc_uah(shelf=TARGET,margin=MARGIN,bonus=BONUS,markup=MARKUP):
    return shelf/markup*(1-bonus-margin)
def max_price_eur(sc,shelf=TARGET,margin=MARGIN,bonus=BONUS,markup=MARKUP):
    n=units(sc)
    return (max_cc_uah(shelf,margin,bonus,markup)/RATE/(1+FIN)-batch_costs(sc)/n)/((1+DUTY)*(1+VAT))
def ask_price(sc,**kw): return math.floor(max_price_eur(sc,**kw)*100+1e-9)/100
def breakdown(sc,p,margin=MARGIN,bonus=BONUS,markup=MARKUP):
    """Повний розклад ціни на упаковку при закупівлі p EUR/уп. (H68–H84)."""
    n=units(sc); f=SC[sc]["freight"]
    prod=p*RATE
    logi=(f+BROKER+OTHER)/n*RATE
    duty=(p+f*FRT_BASE/n)*DUTY*RATE
    vat=(p+f*FRT_BASE/n)*(1+DUTY)*VAT*RATE
    fin=(prod+logi+duty+vat)*FIN
    cc=prod+logi+duty+vat+fin
    partner=cc/(1-bonus-margin); bon=partner*bonus; prof=partner-cc-bon
    return dict(prod=prod,logi=logi,duty=duty,vat=vat,fin=fin,cc=cc,partner=partner,bonus=bon,profit=prof,
                markup=partner*(markup-1),shelf=partner*markup,margin=prof/partner)
def forward_current(sc,margin=MARGIN,bonus=BONUS,markup=MARKUP):
    """Полиця за поточною СС (Self-Cost) — вкладки 1–2."""
    out=[]
    for cc in SC[sc]["cc"]:
        c=cc*RATE; part=c/(1-bonus-margin); out.append(dict(cc=c,partner=part,shelf=part*markup))
    return out
def margin_at_shelf(cc_uah,shelf=TARGET,bonus=BONUS,markup=MARKUP):
    part=shelf/markup; return (part-cc_uah-part*bonus)/part

if __name__=="__main__":
    for sc in ("6","33"):
        mp=max_price_eur(sc); a=ask_price(sc); b=breakdown(sc,a)
        print(sc,'units',units(sc),'batch',round(batch_costs(sc),2),'max €',round(mp,10),'ask',a,
              'shelf',round(b['shelf'],10),'margin',round(b['margin'],6),'cc',round(b['cc'],6))
        print('   breakdown',{k:round(v,4) for k,v in b.items()})
        fc=forward_current(sc); print('   current shelf @30%',[round(x['shelf'],4) for x in fc])
        print('   margin at 120 with current CC',[round(margin_at_shelf(x['cc'])*100,2) for x in fc])
