# -*- coding: utf-8 -*-
"""Ручний аудит сегментів 1/2: Buldak і Toomba — локшина із соусом (воду зливають), а не з бульйоном."""
import json,statistics as st,collections,os
SC=os.path.dirname(os.path.abspath(__file__))
d=json.load(open(f"{SC}/segmented_before_reclass.json",encoding='utf-8')); rows=d['rows']
TO_SAUCE={  # EAN -> причина
 '08801073113428':'Samyang Buldak 2× spicy Hot Chicken 140 г — сухий формат із соусом',
 '08801073110502':'Samyang Buldak Hot Chicken 140 г — сухий формат із соусом',
 '08801073116474':'Samyang Buldak з сиром 130 г — родина Buldak, соус',
 '08801073113312':'Samyang з сиром гостра 130 г — родина Buldak, соус',
 '08801043015479':'Nongshim Toomba K-Pop Spicy & Creamy 137 г — stir-fry з вершковим соусом',
}
n=0
for r in rows:
    if r['ean'] in TO_SAUCE and r['seg']==1:
        r['seg']=2; n+=1
print('перенесено',n)
json.dump(d,open(f"{SC}/segmented.json","w",encoding='utf-8'),ensure_ascii=False,indent=1)
SS={}
for k in '1234':
    v=[r for r in rows if r['seg']==int(k)]; pm=[r['pmed'] for r in v]; q=st.quantiles(pm,n=4)
    bl=collections.Counter(r['tm'] for r in v).most_common()
    SS[k]={'n':len(v),'brands':len(set(r['tm'] for r in v)),'chains':len(set(c for r in v for c in r['chains'])),
           'lo':round(min(r['pmin'] for r in v)),'hi':round(max(r['pmax'] for r in v)),'med':round(st.median(pm)),
           'p25':round(q[0]),'p75':round(q[2]),'g':round(st.median([r['per100'] for r in v])),
           'wlo':round(min(r['w'] for r in v)),'whi':round(max(r['w'] for r in v)),
           'blist':[[b,c] for b,c in bl],'prices':pm}
    print(k,{x:SS[k][x] for x in ('n','brands','chains','lo','hi','med','p25','p75','g','wlo','whi')})
json.dump(SS,open(f"{SC}/segstats.json","w",encoding='utf-8'),ensure_ascii=False,indent=1)
