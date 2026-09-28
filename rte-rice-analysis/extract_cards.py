# -*- coding: utf-8 -*-
"""Pull every product card out of the original deck: photo plus its text.

Copying their slides is what breaks the file, so the card slides get rebuilt
from scratch instead. That needs the photos and the text that belongs to each
one. A card is a column of text boxes with a price in it; the photo is the
picture sitting above that column.
"""
import json, os, re
from pptx import Presentation
from pptx.util import Emu

ORIG = "/root/.claude/uploads/df3fafbd-cd5a-5689-8909-9c95fa4cbf11/c731572d-RTE_Rice_Market_Research..pptx"
MEDIA = "card_media"
os.makedirs(MEDIA, exist_ok=True)

CARD_SLIDES = list(range(7, 17))          # brand cards + the adjacent shelf
prs = Presentation(ORIG)
out = {}

for oi in CARD_SLIDES:
    sl = prs.slides[oi]
    boxes = []
    for sh in sl.shapes:
        if not (sh.has_text_frame and sh.text_frame.text.strip()):
            continue
        x, y, w = Emu(sh.left).inches, Emu(sh.top).inches, Emu(sh.width).inches
        if y < 1.80 or y > 6.95 or w > 3.2:
            continue
        boxes.append((round(x, 2), y, sh.text_frame.text.strip()))

    cols = {}
    for x, y, t in boxes:
        cols.setdefault(x, []).append((y, t))

    cards = []
    for x, items in cols.items():
        items.sort()
        cur = [items[0]]
        for it in items[1:]:
            if it[0] - cur[-1][0] > 0.45:
                cards.append((x, cur)); cur = [it]
            else:
                cur.append(it)
        cards.append((x, cur))

    kept = []
    for x, c in cards:
        ny, ntxt = c[0]
        if ny < 2.30 or len(c) < 3 or len(ntxt) > 45:
            continue
        if not any("грн" in t for _, t in c):
            continue
        kept.append({"x": x, "name_y": ny,
                     "lines": [t for _, t in c]})

    # attach the picture whose centre sits in this card's column, above its text
    pics = [(Emu(sh.left).inches, Emu(sh.top).inches,
             Emu(sh.width).inches, Emu(sh.height).inches, sh)
            for sh in sl.shapes if "PICTURE" in str(sh.shape_type)]
    for card in kept:
        best, bestd = None, 9e9
        for px, py, pw, ph, sh in pics:
            cx = px + pw / 2
            if py + ph > card["name_y"] + 0.05:
                continue
            dist = abs(cx - (card["x"] + 0.7))
            if dist < bestd:
                best, bestd = sh, dist
        if best is not None and bestd < 1.6:
            blob = best.image.blob
            fn = f"{MEDIA}/s{oi+1:02}_{card['x']:.2f}.{best.image.ext}"
            with open(fn, "wb") as fh:
                fh.write(blob)
            card["photo"] = fn
            card["photo_wh"] = [round(Emu(best.width).inches, 3),
                                round(Emu(best.height).inches, 3)]
        else:
            card["photo"] = None
    kept.sort(key=lambda c: (c["name_y"], c["x"]))
    out[oi] = kept
    withphoto = sum(1 for c in kept if c["photo"])
    print(f"слайд {oi+1:02}: карток {len(kept)}, з фото {withphoto}")

json.dump(out, open("cards.json", "w"), ensure_ascii=False, indent=1)
print("разом карток:", sum(len(v) for v in out.values()))
