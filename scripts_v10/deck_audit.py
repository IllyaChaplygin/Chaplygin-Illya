import sys,re,math,openpyxl
from pptx import Presentation
sys.path.insert(0,'/tmp/claude-0/-home-user-Chaplygin-Illya/8bc0034a-4b53-5022-9bf5-e5697c987387/scratchpad')
import calc120 as M
X=openpyxl.load_workbook('/home/user/Chaplygin-Illya/RAMEN_FINMODEL_INTERNAL_UA.xlsx',data_only=True)
lang=sys.argv[2]; P=Presentation(sys.argv[1])
def texts(sl):
    out=[]
    def rec(shs):
        for sh in shs:
            if sh.shape_type==6: rec(sh.shapes)
            elif sh.has_text_frame and sh.text_frame.text.strip(): out.append(sh.text_frame.text.strip())
    rec(sl.shapes); return out
def f2(x): s=f"{x:,.2f}".replace(',',' '); return s.replace('.',',') if lang=='ua' else s.replace(' ',',') if False else (s.replace('.',',') if lang=='ua' else f"{x:,.2f}")
def e(x,n=3): s=f"{x:.{n}f}"; return '€'+(s.replace('.',',') if lang=='ua' else s)
fails=0
slides=list(P.slides)
def need(i,t):
    global fails
    if t not in set(texts(slides[i-1])): fails+=1; print('MISSING slide',i,repr(t))
for i,sheet,ar in ((27,'Цена 120 грн — 33 паллет',19),(28,'Цена 120 грн — 6 паллет',20)):
    w=X[sheet]
    for r,pre in ((4,''),(6,''),(8,'')):
        for c in 'BCDEFGH':
            need(i,f2(w[f'{c}{r}'].value)+' ₴'); need(i,e(w[f'{c}{r+1}'].value))
    for c in 'BCDEFGH':
        v=w[f'{c}10'].value; s=f"{v:+,.2f}".replace(',',' ').replace('.',',' if lang=='ua' else '.').replace('-','−'); need(i,s+' ₴')
        d=w[f'{c}11'].value; need(i,('+' if d>0 else '−')+e(abs(d)))
    ask=w[f'B{ar}'].value; need(i,e(ask,2))
    need(i,f"{f2(w['H4'].value)} / {f2(w['H6'].value)} ₴"); need(i,f"{f2(w['H8'].value)} ₴")
    print(i,'ask',ask,'shelf',[round(w[f'H{r}'].value,2) for r in (4,6,8)],'tile texts ok?')
# сетка чувствительности — по независимой формуле
BON=0.35
for sc,col in (('33',0),('6',1)):
    pass
print('deck fails',fails)
