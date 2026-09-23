import json,time,sys
sys.argv=['x']; exec(open('sweep2.py').read().split('TERMS={')[0])
TERMS=["tilda рис готовий","tilda microwave rice","veetee рис","riso gallo рис готовий","oryza рис готовий",
 "kupiec рис готовий","sonko рис готовий","рис готовий 250 г пакет","рис 2 хвилини мікрохвильовка",
 "zihaiguo рис","自嗨锅","саморозігрівальний горщик рис","саморозігрівальна каша рис","mo xiao xian",
 "莫小仙","haidilao 360","рис з овочами сублімований","сублімований плов","плов сублімований",
 "adventure menu рис","lyofood рис","travellunch рис курка","trek'n eat рис","rations рис",
 "сухпай рис","рис каррі готовий","рис теріякі готовий","рис швидкого приготування курка",
 "рис швидкого приготування яловичина","рис швидкого приготування чашка","японський рис готовий",
 "sato рис","sato no gohan","тайський рис готовий","рис готовий корея 210","cj рис 210","ottogi рис 210",
 "рис готовий ottogi білий","bibigo","рисова каша готова з м'ясом","ризото швидкого приготування"]
out=[]
for t in TERMS:
    try:
        c=cards(t); out+=c; print(f"{t:40} {len(c)}",file=sys.stderr)
    except Exception as e: print(t,"ERR",file=sys.stderr)
    time.sleep(0.35)
json.dump(out,open("sweep3.json","w"),ensure_ascii=False,indent=1)
