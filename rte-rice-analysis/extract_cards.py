# -*- coding: utf-8 -*-
"""Pull every product card out of the original deck: photo plus its text.

Copying their slides is what breaks the file, so the card slides get rebuilt
from scratch instead. That needs the photos and the text that belongs to each
one. A card is a column of text boxes with a price in it; the photo is the
picture sitting above that column.
"""
import io, json, os, re
from PIL import Image
from pptx import Presentation
from pptx.util import Emu

MAXPX = 340            # cards show these at about an inch; anything more is dead weight


def _is_blank(blob):
    """True for an all-white image.

    Some cards sit on a white rectangle that is itself a picture. Picking the
    nearest picture lands on that plate instead of the product shot, and the
    card ends up showing an empty box.
    """
    try:
        im = Image.open(io.BytesIO(blob)).convert("L")
    except Exception:
        return True
    lo, hi = im.getextrema()
    return lo > 246


def _shrink(blob, path):
    """Save the photo at the size it is actually shown, as PNG.

    The deck that opens carries eight small PNGs; the one that fails carries
    sixty photos at full size, half of them JPEG. Normalising to small PNGs
    removes that difference and cuts the file by an order of magnitude.
    """
    im = Image.open(io.BytesIO(blob))
    if max(im.size) > MAXPX:
        sc = MAXPX / max(im.size)
        im = im.resize((max(1, int(im.width * sc)), max(1, int(im.height * sc))),
                       Image.LANCZOS)
    if im.mode in ("RGBA", "LA", "P"):          # flatten transparency onto white
        bg = Image.new("RGB", im.size, (255, 255, 255))
        im = im.convert("RGBA")
        bg.paste(im, mask=im.split()[-1])
        im = bg
    elif im.mode != "RGB":
        im = im.convert("RGB")
    im.save(path, "JPEG", quality=80, optimize=True, progressive=False)

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
        cand = []
        for px, py, pw, ph, sh in pics:
            cx = px + pw / 2
            if py + ph > card["name_y"] + 0.05:
                continue
            dist = abs(cx - (card["x"] + 0.7))
            if dist < 1.6:
                cand.append((dist, sh))
        cand.sort(key=lambda t: t[0])
        best, bestd = None, 9e9
        for dist, sh in cand:
            if _is_blank(sh.image.blob):      # біла плашка, не фото товару
                continue
            best, bestd = sh, dist
            break
        if best is not None:
            fn = f"{MEDIA}/s{oi+1:02}_{card['x']:.2f}_{card['name_y']:.2f}.jpg"
            _shrink(best.image.blob, fn)
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
