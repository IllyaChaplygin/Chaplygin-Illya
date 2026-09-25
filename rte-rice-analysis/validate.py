# -*- coding: utf-8 -*-
"""Geometry validator: overflow + text-on-text collision + font-fit estimate."""
from pptx import Presentation
from pptx.util import Emu
import sys
W,H=13.333,7.5
def inch(v): return Emu(v).inches
prs=Presentation(sys.argv[1] if len(sys.argv)>1 else "out/RTE_Rice_Market_Analysis_v2.pptx")
issues=0
for i,s in enumerate(prs.slides,1):
    boxes=[]
    for sh in s.shapes:
        x,y,w,h=inch(sh.left),inch(sh.top),inch(sh.width),inch(sh.height)
        # overflow
        if x< -0.02 or y< -0.02 or x+w>W+0.02 or y+h>H+0.02:
            if not (w>=W-0.1 and h>=H-0.1):  # ignore full-bleed bg
                print(f"[S{i:02}] OVERFLOW {sh.shape_type} '{(sh.text_frame.text[:34] if sh.has_text_frame else sh.name)}' "
                      f"x{x:.2f} y{y:.2f} w{w:.2f} h{h:.2f}")
                issues+=1
        if sh.has_text_frame and sh.text_frame.text.strip():
            txt=sh.text_frame.text.strip()
            sizes=[r.font.size.pt for p in sh.text_frame.paragraphs for r in p.runs if r.font.size]
            pt=max(sizes) if sizes else 10
            # crude fit check: chars per line at ~0.52*pt/72 inch avg advance
            cpl=max(1,int(w/(0.50*pt/72)))
            lines=sum(max(1,-(-len(p.text)//cpl)) for p in sh.text_frame.paragraphs if p.text)
            need=lines*(pt*1.3/72)
            if need>h+0.16:
                print(f"[S{i:02}] TEXT MAY OVERFLOW BOX  '{txt[:40]}'  need~{need:.2f}in box h={h:.2f}in")
                issues+=1
            boxes.append((x,y,w,h,txt[:30],pt))
    # text-on-text overlap
    for a in range(len(boxes)):
        for b in range(a+1,len(boxes)):
            x1,y1,w1,h1,t1,_=boxes[a]; x2,y2,w2,h2,t2,_=boxes[b]
            ox=min(x1+w1,x2+w2)-max(x1,x2); oy=min(y1+h1,y2+h2)-max(y1,y2)
            if ox>0.06 and oy>0.06:
                print(f"[S{i:02}] TEXT OVERLAP  '{t1}' <> '{t2}'  ({ox:.2f}x{oy:.2f}in)")
                issues+=1
print("\nissues:",issues)
