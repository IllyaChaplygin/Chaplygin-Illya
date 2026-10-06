import pickle, collections, json
from cls import seg, g, ALL, hdr, idx
rows=[r for r in ALL if g(r,'hs_code_normalized')=='1904901000']
CLASS={'frozen':'Заморожені суміші','chips':'Снеки','dolma':'Інше','risotto_kit':'Інше','topping':'Інше','raw_or_industrial':'Інше',
       'porridge':'Інше','tteok':'Інше','sublimate':'Інше','S2_boilwater':'S2','S1_bens':'S1','S1_cj':'S1','S1_clearspring':'S1','S1_haldiram':'S1'}
months=['2025-%02d'%m for m in range(1,13)]+['2026-%02d'%m for m in range(1,6)]
M={m:collections.Counter() for m in months}
tot=collections.Counter(); usd=collections.Counter()
for r in rows:
    s=seg(r); c=CLASS.get(s,'Інше'); m=g(r,'report_month')
    M[m][c]+=g(r,'net_weight_kg') or 0
    tot[s]+=g(r,'net_weight_kg') or 0; usd[s]+=g(r,'invoice_value_usd') or 0
if __name__=='__main__':
    print({k:round(v,1) for k,v in tot.items()})
    print('--- monthly t')
    for m in months:
        print(m, {k:round(v/1000,2) for k,v in M[m].items()}, round(sum(M[m].values())/1000,2))
    cmp_m=['2025-01','2025-02','2025-04','2025-05'], ['2026-01','2026-02','2026-04','2026-05']
    for lab,ms in zip(('2025','2026'),cmp_m):
        d=collections.Counter()
        for m in ms:
            for k,v in M[m].items(): d[k]+=v
        print(lab, {k:round(v/1000,2) for k,v in d.items()}, round(sum(d.values())/1000,2))
    # segments 1-2 detail rows
    print('--- S1/S2 per importer')
    imp=collections.defaultdict(lambda:[0,0,[], set()])
    for r in rows:
        s=seg(r)
        if s.startswith('S'):
            k=(g(r,'importer_display_name')[:28], s)
            imp[k][0]+=g(r,'net_weight_kg'); imp[k][1]+=g(r,'invoice_value_usd'); imp[k][2].append(g(r,'report_month')); imp[k][3].add((g(r,'manufacturer_normalized') or '')[:25]+'|'+(g(r,'origin_country_normalized_ua') or ''))
    for k,v in imp.items(): print(k, round(v[0],1), round(v[1]), round(v[1]/v[0],2), min(v[2]), max(v[2]), v[3])
