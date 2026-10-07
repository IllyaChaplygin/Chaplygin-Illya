import re,json
STORE={'Ашан':'Auchan','Зараз':'Zaraz','Космос':'Kosmos','Таврія В':'Tavria V','Таврія':'Tavria','Клас':'Klass','Восторг':'Vostorg','ЧудоМаркет':'ChudoMarket','Сільпо':'Silpo','Торба':'Torba','Ідеал':'Ideal','МегаМаркет':'MegaMarket','ЕКО':'EKO','Епіцентр':'Epicentr','Суші Повар':'Sushi Povar','Тайякі Март':'Taiyaki Mart','ЧудоМ.':'ChudoM.'}
# truncated store tokens
TRUNC={'Кос…':'Kos…','Таврія …':'Tavria …','Таврія В…':'Tavria V…','Чудо…':'Chudo…','Ашан…':'Auchan…'}
def store(tok):
    t=tok.strip()
    if t in STORE: return STORE[t]
    if t in TRUNC: return TRUNC[t]
    m=re.match(r'^«(.+)»$',t)
    if m and m.group(1) in STORE: return STORE[m.group(1)]
    return t
def r_weight(t):
    m=re.match(r'^(\d+(?:,\d+)?(?:–\d+(?:,\d+)?)?) г(?: ·| \+|$)(.*)$',t)
    if not m: return None
    n=m.group(1); rest=t[len(n)+2:]
    if rest.startswith(' +'): return f"{n} g +"+rest[2:]
    if rest.startswith(' · '):
        toks=rest[3:].split(' · ')
        out=[]
        for k in toks:
            mm=re.match(r'^(.*?)( \+\d+)$',k)
            if mm: out.append(store(mm.group(1))+mm.group(2))
            else: out.append(store(k))
        return f"{n} g · "+" · ".join(out)
    return f"{n} g"+rest
def num(s):
    s=re.sub(r'(\d),(\d)',r'\1.\2',s)
    return s
