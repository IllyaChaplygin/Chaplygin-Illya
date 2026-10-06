import pickle, collections
hdr,ALL = pickle.load(open('raw.pkl','rb'))
idx={h:i for i,h in enumerate(hdr)}
def g(r,k): return r[idx[k]]
def seg(r):
    d=(g(r,'product_description') or '').upper(); m=(g(r,'manufacturer_normalized') or '').upper()
    if 'ЗАМОРОЖ' in d or 'VICI' in d or 'MEXICAN PAN' in d or ('ARB' in d.split() and 'ТАЙСЬКИЙ' not in d and 'КРУГЛИЙ' not in d) or 'БАСКСКАЯ' in d: return 'frozen'
    if 'ЧІПС' in d or 'ЧИПС' in d or 'CHIPS' in d: return 'chips'
    if 'ПОСИПК' in d: return 'topping'
    if 'ДОЛМ' in d: return 'dolma'
    if 'TREK' in d and 'EAT' in d or 'СУБЛІМАТ' in d or 'ЗАЛИТИ ОКРОПОМ' in d and 'СУБЛ' in d: return 'sublimate'
    if 'RICE PORRIDGE' in d or 'GYMBEAM' in m: return 'porridge'
    if 'YOPOKKI' in d or 'ТТОКПОКК' in d: return 'tteok'
    if ('НЕПРИГОТОВАН' in d or 'ДЛЯ ПРИГОТУВАННЯ РІЗОТТО' in d or 'РИЗОТО CORDERO' in d or 'CASA RINALDI' in d or 'RISO GALLO' in d or 'PRINCIPATO' in d
        or 'TREVIJANO' in d or 'TREVIJANO' in m or 'TARTUFO' in m or 'РІЗОТО ШВИДКОГО' in d): return 'risotto_kit'
    if 'РИС 10 ХВ' in d or 'АШАН РИС' in d or 'РИС ТАЙСЬКИЙ' in d or 'РИС КРУГЛИЙ' in d or 'ХАРЧОВІЙ ПРОМИСЛОВ' in d or 'ОЗЕРЯНКА' in d: return 'raw_or_industrial'
    if 'MARS AUSTRIA' in m or 'UNCLE BEN' in d: return 'S1_bens'
    if 'CJCHEIL' in m: return 'S1_cj'
    if 'CLEARSPRING' in m: return 'S1_clearspring'
    if 'RTE BIRYANI' in d or 'HALDIRAM' in m: return 'S1_haldiram'
    if 'INSTANT RICE' in d or 'РИС ШВИДКОГО ПРИГОТУВАННЯ' in d: return 'S2_boilwater'
    return 'unclassified'
if __name__=='__main__':
    rows=[r for r in ALL if g(r,'hs_code_normalized')=='1904901000']
    c=collections.Counter(); w=collections.Counter()
    for r in rows:
        s=seg(r); c[s]+=1; w[s]+=g(r,'net_weight_kg') or 0
    for k in sorted(w,key=lambda k:-w[k]): print(f'{k:18s} {c[k]:4d} {w[k]:10.1f}')
    print('--- unclassified'); seen=set()
    for r in rows:
        if seg(r)=='unclassified':
            d=(g(r,'product_description') or '')[:150]
            if d in seen: continue
            seen.add(d); print(g(r,'manufacturer_normalized'),'|',d)
    print('--- S1/S2/sublimate rows')
    for r in rows:
        if seg(r).startswith('S') or seg(r)=='sublimate':
            print(seg(r), g(r,'report_month'), round(g(r,'net_weight_kg'),1), round(g(r,'invoice_value_usd'),0), g(r,'importer_display_name')[:35], '|', g(r,'origin_country_normalized_ua'), '|', (g(r,'manufacturer_normalized') or '')[:30], '|', (g(r,'product_description') or '')[:70])
