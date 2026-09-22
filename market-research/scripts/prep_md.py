"""Photo preparation for the Morskyi Dim deck.

Pack shots come from the research file (cards verified on Ukrainian shop
fronts on 22.09.2026); brand furniture — logo, wave background — is lifted
from the Morskyi Dim Nissin overview so the two decks share one identity.
Every shot is trimmed to the product edge and keeps its own proportions.
"""
import os

import numpy as np
from PIL import Image, ImageFilter

os.makedirs("md", exist_ok=True)
F = "unf/ppt/media"


def trim(im, thr=247, pad=0.02):
    a = np.asarray(im.convert("RGB")).min(axis=2)
    ys, xs = np.where(a < thr)
    if not len(xs):
        return im
    x0, x1, y0, y1 = xs.min(), xs.max(), ys.min(), ys.max()
    m = int(pad * max(x1 - x0, y1 - y0))
    return im.crop((max(0, x0 - m), max(0, y0 - m),
                    min(im.width, x1 + m), min(im.height, y1 + m)))


def pack(src, dst, h=880):
    if not os.path.exists(src):
        print("MISSING", src); return
    im = trim(Image.open(src).convert("RGB"))
    im = im.resize((max(1, round(im.width * h / im.height)), h), Image.LANCZOS)
    im = im.filter(ImageFilter.UnsharpMask(radius=1.2, percent=55, threshold=3))
    im.save(dst, quality=95)
    print(f"{dst:30} {im.size}  ratio {im.width / im.height:.2f}")


# ── the 14 comparable side-dish cards + the bowls ─────────────────────────
MAP = {
    # Ben's Original — Сільпо catalogue
    "bens_basmati250":   "image.png",
    "bens_mediterran":   "image12.png",
    "bens_longgrain":    "image13.png",
    "bens_risibisi":     "image14.png",
    "bens_curry_indien": "image15.png",
    "bens_curry_linsen": "image16.png",
    "bens_sweetchili":   "image17.png",
    "bens_mexikanisch":  "image9.png",
    "bens_curryreis":    "image19.png",
    "bens_basmati220":   "image11.png",
    "bens_stickybowl":   "image10.png",
    "bens_bio240":       "image22.png",
    "bens_lang220":      "image23.png",
    # the other three brands
    "bibigo_bowl":       "image2.png",
    "bibigo_bowl2":      "image24.png",
    "clearspring":       "image4.png",
    "ottogi_burger":     "image5.png",
    "ottogi_chicken":    "image26.png",
    "ottogi_hamburg":    "image27.png",
    "ottogi_octopus":    "image28.png",
    "ottogi_jjampong":   "image29.png",
    "ottogi_tuna":       "image30.png",
    "ottogi_kimchi":     "image31.png",
}
for name, f in MAP.items():
    pack(f"{F}/{f}", f"md/{name}.jpg")

# ── our own supply, from the supplier deck ────────────────────────────────
S = "un/ppt/media"
OURS = {"own_pouch_jasmine": "image9.jpeg", "own_pouch_mexican": "image10.jpeg",
        "own_pouch_spicy": "image12.jpeg", "own_chefrey": "image4.jpeg",
        "own_spar_cups": "image37.png", "own_cup": "image13.jpeg"}
for name, f in OURS.items():
    pack(f"{S}/{f}", f"md/{name}.jpg")
