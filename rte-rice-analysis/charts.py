# -*- coding: utf-8 -*-
import matplotlib; matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch
import numpy as np, data as d

plt.rcParams.update({
 "font.family":"DejaVu Sans","font.size":9,
 "axes.edgecolor":d.LINE,"axes.labelcolor":d.MUTED,
 "xtick.color":d.MUTED,"ytick.color":d.MUTED,
 "text.color":d.INK,"axes.facecolor":"none","figure.facecolor":"none",
 "savefig.facecolor":"none","axes.grid":False,
})
OUT="out/"

def strip(ax, bottom=True, left=False):
    for s in ("top","right","left","bottom"): ax.spines[s].set_visible(False)
    if bottom: ax.spines["bottom"].set_visible(True); ax.spines["bottom"].set_color(d.LINE)
    if left:   ax.spines["left"].set_visible(True);   ax.spines["left"].set_color(d.LINE)

# ============ CHART 1: формати — 3 панелі, small multiples ============
def chart_formats():
    F=d.FORMATS
    names=[f[0] for f in F]; cols=[f[8] for f in F]
    pos=[f[1] for f in F]; share=[f[2] for f in F]
    med=[f[4] for f in F]; w=[f[7] for f in F]
    per100=[f[4]/f[7]*100 for f in F]
    y=np.arange(len(F))[::-1]

    fig,axes=plt.subplots(1,3,figsize=(13.4,3.5),gridspec_kw={"width_ratios":[1,1,1],"wspace":0.55})
    panels=[
      (pos,   "ПОЗИЦІЙ У ПРОДАЖУ",        lambda v,i:f"{v}  ·  {share[i]} %"),
      (med,   "МЕДІАНА ЦІНИ ЗА УПАКОВКУ",  lambda v,i:f"{v} грн"),
      (w,     "ТИПОВА ВАГА УПАКОВКИ",      lambda v,i:f"{v} г"),
    ]
    for ax,(vals,title,lab) in zip(axes,panels):
        for yy,v,c in zip(y,vals,cols):
            ax.barh(yy,v,height=0.52,color=c,zorder=3,
                    joinstyle="round",edgecolor="none")
        for yy,v,c,i in zip(y,vals,cols,range(len(vals))):
            ax.text(v+max(vals)*0.035,yy,lab(v,i),va="center",ha="left",
                    fontsize=8.5,color=d.INK2,fontweight="bold")
        ax.set_yticks(y); ax.set_yticklabels(names,fontsize=8.5,color=d.INK2)
        ax.set_xlim(0,max(vals)*1.42); ax.set_xticks([])
        ax.set_title(title,fontsize=7.5,color=d.MUTED,loc="left",pad=11,
                     fontweight="bold")
        strip(ax,bottom=False)
        ax.tick_params(axis="y",length=0)
        if ax is not axes[0]: ax.set_yticklabels([])
    fig.text(0.007,-0.04,"Пауч має найменше позицій — але найнижчу ціну входу в категорію. "
             "Дойпак лідирує за кількістю лише тому, що туристичні магазини заводять десятки дрібних брендів.",
             fontsize=8,color=d.MUTED)
    fig.savefig(OUT+"c1_formats.png",dpi=220,bbox_inches="tight",transparent=True)
    plt.close(fig)

# ============ CHART 2: цінові сходи за 100 г — ринок vs ми ============
def chart_ladder():
    rows=[
      ("Суміжна полиця · консерва",18.5,47.6,40.0,d.MUTED2,"market"),
      ("Ben's Original · пауч",18,81,41,d.C_AMBER,"market"),
      ("Пауч · медіана категорії",None,None,43.6,d.C_AMBER,"ref"),
      ("Henan · чаша (окріп)",58,106,80,d.C_TEAL,"market"),
      ("Bibigo · лоток 210 г",64,90,77,d.C_TEAL,"market"),
      ("Ottogi · чаша",79,130,88,d.C_TEAL,"market"),
      ("Чаша · медіана категорії",None,None,105,d.C_TEAL,"ref"),
      ("Дойпак · медіана",None,None,286,d.C_PURPLE,"ref"),
      ("Коробка · медіана",None,None,308,d.C_PINK,"ref"),
    ]
    # our groups
    groups={}
    for g,sku,wt,c20,c40,sup,fmt in d.OWN:
        groups.setdefault(g,[]).append(d.shelf(c20)/wt*100)
    ours=[(g,min(v),max(v)) for g,v in groups.items()]
    ours.sort(key=lambda x:x[1])

    fig,ax=plt.subplots(figsize=(13.4,6.0))
    labels=[];ypos=[];y=0
    # market block
    for name,lo,hi,mid,col,kind in rows:
        if lo is not None:
            ax.plot([lo,hi],[y,y],lw=7,color=col,alpha=0.22,solid_capstyle="round",zorder=2)
        ax.plot([mid],[y],"o",ms=9,color=col,zorder=4,
                markeredgecolor="white",markeredgewidth=2)
        txt=f"{mid:.0f}" if lo is None else f"{mid:.0f}  ({lo:.0f}–{hi:.0f})"
        ax.text(mid+9,y,txt,va="center",fontsize=8,color=d.MUTED)
        labels.append(name); ypos.append(y); y-=1
    y-=1.35
    sep=y+0.95
    for g,lo,hi in ours:
        ax.plot([lo,hi],[y,y],lw=7,color=d.C_BRAND,alpha=0.30,solid_capstyle="round",zorder=2)
        ax.plot([lo],[y],"o",ms=9,color=d.C_BRAND,zorder=4,markeredgecolor="white",markeredgewidth=2)
        if hi-lo>1: ax.plot([hi],[y],"o",ms=9,color=d.C_BRAND,zorder=4,markeredgecolor="white",markeredgewidth=2)
        ax.text(hi+9,y,f"{lo:.0f}–{hi:.0f}" if hi-lo>1 else f"{lo:.0f}",
                va="center",fontsize=8,color=d.INK2,fontweight="bold")
        labels.append("НАШ · "+g); ypos.append(y); y-=1

    ax.axhline(sep,color=d.LINE,lw=1,ls=(0,(4,4)),zorder=1)
    ax.text(3,sep-0.32,"НАША ФІНМОДЕЛЬ  ·  20′ FCL навалом, маржа 35 %, бонус 25 %, націнка 1,40",
            fontsize=7.5,color=d.C_BRAND,fontweight="bold",va="top")
    # shelf window band
    ax.axvspan(18,50,color=d.C_AMBER,alpha=0.055,zorder=0)
    ax.text(34,1.05,"ВІКНО МЕРЕЖЕВОЇ ПОЛИЦІ  18–50 грн/100 г",fontsize=7.5,
            color=d.C_AMBER,fontweight="bold",ha="center")

    ax.set_yticks(ypos); ax.set_yticklabels(labels,fontsize=8.5,color=d.INK2)
    for t,(n,*_ ) in zip(ax.get_yticklabels(),[(l,) for l in labels]):
        if n.startswith("НАШ"): t.set_color(d.INK); t.set_fontweight("bold")
    ax.set_xlim(0,330); ax.set_ylim(y+0.45,1.75)
    ax.set_xlabel("ГРН ЗА 100 Г НА ПОЛИЦІ",fontsize=7.5,color=d.MUTED,fontweight="bold")
    ax.xaxis.set_label_coords(0.5,-0.055)
    for xv in (50,100,150,200,250,300):
        ax.axvline(xv,color=d.LINE,lw=0.8,zorder=0)
    strip(ax); ax.tick_params(axis="y",length=0)
    fig.savefig(OUT+"c2_ladder.png",dpi=220,bbox_inches="tight",transparent=True)
    plt.close(fig)

# ============ CHART 3: ранжир за упаковку — від найдешевшої до найдорожчої ============
def chart_pack():
    market=[("Ben's Original · найдешевший 250 г",45),("Henan · найдешевша чаша 144 г",83),
            ("Маркел · реторт 350 г",95),("Ben's Original · медіана",90),
            ("Portion · реторт 350 г",109),("ПАУЧ · медіана категорії",109),
            ("Henan · медіана",140),("Bibigo · лоток 210 г",162),
            ("Ben's Original · найдорожчий 220 г",179),("ЧАША · медіана категорії",229),
            ("Ottogi · медіана",255),("ДОЙПАК · медіана категорії",315),
            ("Ottogi · найдорожчий",356),("КОРОБКА · медіана категорії",708),
            ("Mo Xiao Xian · найдорожча чаша",931),("Haidilao · максимум ринку",1190)]
    groups={}
    for g,sku,wt,c20,c40,sup,fmt in d.OWN: groups.setdefault(g,[]).append(d.shelf(c20))
    ours=[("НАШ · "+g.replace("BSCM — ","BSCM ").replace("CM Premium — ","CM Premium ")
           .replace("ONE'S / KBROS — ","ONE'S "),min(v)) for g,v in groups.items()]

    rows=[(n,v,False) for n,v in market]+[(n,v,True) for n,v in ours]
    rows.sort(key=lambda r:r[1])
    fig,ax=plt.subplots(figsize=(13.4,7.4))
    y=np.arange(len(rows))[::-1]
    for yy,(n,v,mine) in zip(y,rows):
        is_ref=n.startswith(("ЧАША","ДОЙПАК","КОРОБКА","ПАУЧ"))
        c=d.C_BRAND if mine else (d.INK2 if is_ref else "#C6CEE0")
        ax.barh(yy,v,height=0.58,color=c,zorder=3)
        ax.text(v+18,yy,f"{v:,.0f} грн".replace(","," "),va="center",fontsize=8.5,
                color=d.INK if mine else d.MUTED,
                fontweight="bold" if (mine or is_ref) else "normal")
    ax.set_yticks(y)
    ax.set_yticklabels([r[0] for r in rows],fontsize=8.5)
    for t,(n,v,mine) in zip(ax.get_yticklabels(),rows):
        t.set_color(d.INK if mine else d.INK2)
        if mine: t.set_fontweight("bold")
    ax.axvspan(45,180,color=d.C_AMBER,alpha=0.07,zorder=0)
    ax.text(112,len(rows)-0.35,"ВІКНО МЕРЕЖЕВОЇ ПОЛИЦІ  45–180 грн",fontsize=7.5,
            color=d.C_AMBER,fontweight="bold",ha="center")
    ax.set_xlim(0,1390); ax.set_ylim(-0.8,len(rows)+0.3)
    ax.set_xticks([0,200,400,600,800,1000,1200])
    ax.set_xticklabels(["0","200","400","600","800","1 000","1 200"],fontsize=8)
    ax.set_xlabel("ГРН ЗА ОДНУ УПАКОВКУ НА ПОЛИЦІ",fontsize=7.5,color=d.MUTED,fontweight="bold")
    strip(ax); ax.tick_params(axis="y",length=0)
    fig.savefig(OUT+"c3_pack.png",dpi=220,bbox_inches="tight",transparent=True)
    plt.close(fig)

# ============ CHART 4: FOB -> полиця, міст до цілі ============
def chart_fob():
    targets=[("Медіана пауча 109 грн",109),("Ben's медіана 98 грн",98.4),
             ("Під Ottogi 199 грн",199),("Bibigo-рівень 162 грн",162)]
    skus=[("BSCM стакан 150 г",0.500,d.shelf(0.881437262584171),150,d.C_TEAL),
          ("BSCM пауч 150 г",0.472,d.shelf(0.842849973738682),150,d.C_AMBER),
          ("BSCM пауч 200 г",0.550,d.shelf(0.952556220056099),200,d.C_AMBER),
          ("BSCM пауч 240 г",0.583,d.shelf(0.995418606770833),240,d.C_AMBER),
          ("CM Premium 250 г",0.667,d.shelf(1.13300100081699),250,d.C_GREEN),
          ("ONE'S білий 210 г",0.571,d.shelf(0.971476418890301),210,d.C_PURPLE),
          ("ONE'S fried 200 г",1.238,d.shelf(2.1048630417902),200,d.C_PINK)]
    fig,ax=plt.subplots(figsize=(13.4,4.6))
    y=np.arange(len(skus))[::-1]
    for yy,(n,fob,sh,wt,c) in zip(y,skus):
        ax.barh(yy,sh,height=0.5,color=c,zorder=3)
        ax.text(sh+14,yy,f"{sh:.0f} грн   ·   FOB ${fob:.3f}",
                va="center",fontsize=8.5,color=d.INK2,fontweight="bold")
    ax.set_yticks(y); ax.set_yticklabels([s[0] for s in skus],fontsize=8.5,color=d.INK2)
    for xv,(lab,tv) in zip([109,98.4,162,199],targets):
        pass
    refs=[("Ben's 98",98.4,(0,(3,3)),0.95),("медіана пауча 109",109,"-",0.35),
          ("Bibigo 162",162,(0,(3,3)),0.95),("Ottogi −22 % 199",199,(0,(3,3)),0.35)]
    for lab,tv,st,dy in refs:
        ax.axvline(tv,color=d.INK2 if st=="-" else d.MUTED2,lw=1.3,ls=st,zorder=1)
        ax.text(tv,len(skus)-1+dy,lab,fontsize=7,
                color=d.INK2 if st=="-" else d.MUTED,
                ha="center",va="bottom",fontweight="bold" if st=="-" else "normal")
    ax.set_xlim(0,430); ax.set_ylim(-0.7,len(skus)+0.45)
    ax.set_xlabel("ПОЛИЦЯ ЗА ПОТОЧНОЇ МОДЕЛІ, ГРН  ·  20′ FCL навалом",
                  fontsize=7.5,color=d.MUTED,fontweight="bold")
    strip(ax); ax.tick_params(axis="y",length=0); ax.set_xticks([0,100,200,300,400])
    fig.savefig(OUT+"c4_fob.png",dpi=220,bbox_inches="tight",transparent=True)
    plt.close(fig)

# ============ CHART 5: канали ============
def chart_channels():
    data=[("Туристичні та військові",40,d.C_PURPLE),
          ("Азійські / етнічні фудшопи",20,d.C_TEAL),
          ("Онлайн-супермаркети, МП",3,d.C_AMBER),
          ("Продуктовий роздріб",3,d.C_PINK)]
    fig,ax=plt.subplots(figsize=(6.6,2.9))
    y=np.arange(len(data))[::-1]
    for yy,(n,v,c) in zip(y,data):
        ax.barh(yy,v,height=0.5,color=c,zorder=3)
        ax.text(v+0.7,yy,str(v),va="center",fontsize=9,color=d.INK2,fontweight="bold")
    ax.set_yticks(y); ax.set_yticklabels([x[0] for x in data],fontsize=8.5,color=d.INK2)
    ax.set_xlim(0,46); ax.set_xticks([])
    ax.set_title("ТОЧОК ПРОДАЖУ В КАНАЛІ",fontsize=7.5,
                 color=d.MUTED,loc="left",pad=10,fontweight="bold")
    strip(ax,bottom=False); ax.tick_params(axis="y",length=0)
    fig.savefig(OUT+"c5_channels.png",dpi=220,bbox_inches="tight",transparent=True)
    plt.close(fig)

# ============ CHART 6: модель формат × грамаж ============
def chart_fmt_weight():
    import matplotlib.colors as mcolors
    M=[v for _,v,_ in d.FMT_WEIGHT]; names=[n for n,_,_ in d.FMT_WEIGHT]
    nb=len(d.WBANDS); nf=len(M)
    vmax=max(max(r) for r in M)
    ramp=mcolors.LinearSegmentedColormap.from_list("amb",["#FBF7EF","#F6E3BC","#E8B85F","#C4820A","#8A5A06"])
    fig,ax=plt.subplots(figsize=(13.4,4.10))
    for i in range(nf):
        for j in range(nb):
            v=M[i][j]
            y=nf-1-i
            if v==0:
                ax.add_patch(plt.Rectangle((j+0.035,y+0.035),0.93,0.93,
                             facecolor="#F7F8FB",edgecolor="none",zorder=2))
                ax.text(j+0.5,y+0.5,"—",ha="center",va="center",fontsize=8.5,color="#C6CEE0",zorder=4)
            else:
                ax.add_patch(plt.Rectangle((j+0.035,y+0.035),0.93,0.93,
                             facecolor=ramp(v/vmax),edgecolor="none",zorder=2))
                ax.text(j+0.5,y+0.58,str(v),ha="center",va="center",fontsize=14,
                        color="white" if v/vmax>0.48 else d.INK,fontweight="bold",zorder=4)
                ax.text(j+0.5,y+0.27,"позицій",ha="center",va="center",fontsize=6.2,
                        color="white" if v/vmax>0.48 else d.MUTED,zorder=4)
    # порційне вікно
    ax.add_patch(plt.Rectangle((3,-0.06),2,nf+0.12,fill=False,edgecolor=d.C_BRAND,
                 lw=2.2,zorder=6,joinstyle="round"))
    ax.text(4,nf+0.60,"ПОРЦІЙНЕ ВІКНО 200–299 Г  ·  22 ПОЗИЦІЇ  ·  ТУТ ЖИВЕ ПОЛИЦЯ",
            ha="center",fontsize=8,color=d.C_BRAND,fontweight="bold")
    ax.set_xlim(0,nb); ax.set_ylim(0,nf+0.80)
    ax.set_xticks([j+0.5 for j in range(nb)]); ax.set_xticklabels(d.WBANDS,fontsize=8.5,color=d.INK2)
    ax.set_yticks([nf-1-i+0.5 for i in range(nf)])
    ax.set_yticklabels(names,fontsize=9,color=d.INK2,fontweight="bold")
    ax.tick_params(length=0,pad=30); ax.xaxis.tick_top()
    for sp in ax.spines.values(): sp.set_visible(False)
    fig.savefig(OUT+"c6_fmtweight.png",dpi=220,bbox_inches="tight",transparent=True)
    plt.close(fig)


# ============ CHART 7: ТОП комбінацій формат + грамаж ============
def chart_combos():
    C=d.COMBOS
    fig,(ax,ax2)=plt.subplots(1,2,figsize=(13.4,5.0),
                              gridspec_kw={"width_ratios":[2.45,1],"wspace":0.34})
    y=np.arange(len(C))[::-1]
    for yy,(fmt,wt,n,price,p100,who,c) in zip(y,C):
        ax.barh(yy,n,height=0.60,color=c,zorder=3)
        ax.text(n+0.13,yy,f"{n}",va="center",fontsize=10,color=d.INK,fontweight="bold")
        ax.text(n+0.52,yy,f"медіана {price} грн",va="center",
                fontsize=7.8,color=d.MUTED)
        ax.text(n+2.60,yy,who,va="center",fontsize=7,color=d.MUTED2)
    ax.set_yticks(y)
    ax.set_yticklabels([f"{f}  {w} г" for f,w,*_ in C],fontsize=9,color=d.INK2,fontweight="bold")
    ax.set_xlim(0,10.6); ax.set_ylim(-0.75,len(C)-0.15)
    ax.set_xticks([0,2,4,6,8]); ax.tick_params(axis="y",length=0)
    ax.set_xlabel("ПОЗИЦІЙ У ПРОДАЖУ",fontsize=7.5,color=d.MUTED,fontweight="bold")
    ax.set_title("ТОП-13 ПОЄДНАНЬ «ФОРМАТ + ГРАМАЖ»",fontsize=8,color=d.INK2,
                 loc="left",pad=12,fontweight="bold")
    strip(ax)
    # праворуч — частка форматів у двох зрізах
    for k,(title,data_,total) in enumerate([
            ("УСІ 84 ПОЗИЦІЇ",d.SHARE_ALL,84),
            ("ЛИШЕ МАСОВИЙ КАНАЛ · 46",d.SHARE_MASS,46)]):
        base=0.0; yb=1.0-k*0.64
        for nm,v,c in data_:
            ax2.barh(yb,v/total,left=base,height=0.30,color=c,zorder=3)
            if v/total>0.10:
                ax2.text(base+v/total/2,yb,f"{round(v/total*100)} %",ha="center",va="center",
                         fontsize=9,color="white",fontweight="bold",zorder=5)
            base+=v/total
        ax2.text(0,yb+0.235,title,fontsize=7.5,color=d.MUTED,fontweight="bold")
        lx=0.0
        for nm,v,c in data_:
            if v/total<0.06: continue
            ax2.text(lx,yb-0.268,f"{nm} {v}",fontsize=7,color=c,fontweight="bold")
            lx+=0.30
    ax2.set_xlim(0,1); ax2.set_ylim(0.00,1.40)
    ax2.set_xticks([]); ax2.set_yticks([])
    for sp in ax2.spines.values(): sp.set_visible(False)
    ax2.set_title("ЧАСТКА ФОРМАТУ",fontsize=8,color=d.INK2,loc="left",pad=12,fontweight="bold")
    fig.savefig(OUT+"c7_combos.png",dpi=220,bbox_inches="tight",transparent=True)
    plt.close(fig)


# ============ CHART 8: азійські бренди вже на полиці мереж ============
def chart_asian_shelf():
    import matplotlib.colors as mcolors
    B=d.ASIAN_ON_SHELF; CH=d.RETAIL_CHAINS
    nb, nc = len(B), len(CH)
    vmax=max(max(v.values()) for _,v,_,_ in B)
    ramp=mcolors.LinearSegmentedColormap.from_list("t",["#F2FAFA","#CFEBEC","#7FCBCF","#1E9EA6","#136A70"])
    fig,ax=plt.subplots(figsize=(13.4,4.25))
    for i,(brand,counts,has_rice,what) in enumerate(B):
        y=nb-1-i
        for j,ch in enumerate(CH):
            v=counts.get(ch,0)
            if v==0:
                ax.add_patch(plt.Rectangle((j+0.04,y+0.06),0.92,0.88,
                             facecolor="#F7F8FB",edgecolor="none",zorder=2))
                ax.text(j+0.5,y+0.5,"—",ha="center",va="center",fontsize=9,
                        color="#C6CEE0",zorder=4)
            else:
                ax.add_patch(plt.Rectangle((j+0.04,y+0.06),0.92,0.88,
                             facecolor=ramp(v/vmax),edgecolor="none",zorder=2))
                ax.text(j+0.5,y+0.5,str(v),ha="center",va="center",fontsize=13,
                        color="white" if v/vmax>0.46 else d.INK,fontweight="bold",zorder=4)
        ax.text(nc+0.30,y+0.5,what,va="center",fontsize=8.5,color=d.MUTED,zorder=4)
    ax.set_xlim(0,nc+2.9); ax.set_ylim(0,nb+0.05)
    ax.set_xticks([j+0.5 for j in range(nc)]+[nc+0.30])
    ax.set_xticklabels(CH+["ЩО САМЕ СТОЇТЬ"],fontsize=8.5,color=d.INK2,fontweight="bold")
    ax.set_yticks([nb-1-i+0.5 for i in range(nb)])
    ax.set_yticklabels([b for b,_,_,_ in B],fontsize=10,color=d.INK,fontweight="bold")
    ax.tick_params(length=0,pad=10); ax.xaxis.tick_top()
    ax.get_xticklabels()[-1].set_ha("left")
    for sp in ax.spines.values(): sp.set_visible(False)
    fig.text(0.012,-0.045,"163 позиції азійських брендів уже стоять у національних мережах — "
             "і жодна з них не рис.",fontsize=10,color=d.C_PINK,fontweight="bold")
    fig.savefig(OUT+"c8_asian_shelf.png",dpi=220,bbox_inches="tight",transparent=True)
    plt.close(fig)


for f in (chart_formats,chart_asian_shelf,chart_combos,chart_fmt_weight,chart_pack,chart_fob,chart_channels):
    f(); print("ok",f.__name__)
