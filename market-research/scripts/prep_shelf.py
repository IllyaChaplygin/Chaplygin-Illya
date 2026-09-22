"""Photo preparation for the SHELF deck.

Two sources: pack shots of what is actually on sale in Ukraine (downloaded from
the shops that sell them) and our own suppliers' product photography (lifted
from the RTE_Rice_Suppliers deck's media). Everything is trimmed of its white
studio surround, keeps its own proportions, and is sharpened once.
"""
import os
import shutil

import numpy as np
from PIL import Image, ImageEnhance, ImageFilter

os.makedirs("assets", exist_ok=True)
M = "un/ppt/media"


def trim(im, thr=246, pad=0.015):
    a = np.asarray(im.convert("RGB")).min(axis=2)
    ys, xs = np.where(a < thr)
    if not len(xs):
        return im
    x0, x1, y0, y1 = xs.min(), xs.max(), ys.min(), ys.max()
    m = int(pad * max(x1 - x0, y1 - y0))
    return im.crop((max(0, x0 - m), max(0, y0 - m),
                    min(im.width, x1 + m), min(im.height, y1 + m)))


def pack(src, dst, h=900, do_trim=True):
    if not os.path.exists(src):
        print("MISSING", src); return
    im = Image.open(src).convert("RGB")
    if do_trim:
        im = trim(im)
    if im.height != h:
        im = im.resize((max(1, round(im.width * h / im.height)), h), Image.LANCZOS)
    im = im.filter(ImageFilter.UnsharpMask(radius=1.3, percent=60, threshold=3))
    im.save(dst, quality=94)
    print(f"{dst:38} {im.size}  ratio {im.width / im.height:.2f}")


# ── what is actually sold in Ukraine ──────────────────────────────────────
pack("img/ua_hetbahn.jpg", "assets/ua_bibigo.jpg")
pack("img/ua_selfheat.jpg", "assets/ua_selfheat_187.jpg")
pack("img/ua_selfheat2.jpg", "assets/ua_selfheat_272.jpg")

# ── substitutes, from the chains' own catalogue images ────────────────────
pack("img/sub_rice_bags.jpg", "assets/sub_rice_bags.jpg")
pack("img/sub_noodle_cup.jpg", "assets/sub_noodle_cup.jpg")
pack("img/sub_noodle_shin.jpg", "assets/sub_noodle_shin.jpg")
pack("img/sub_tteok.jpg", "assets/sub_tteok.jpg")
pack("img/sub_porridge.jpg", "assets/sub_porridge.jpg")

# ── our own supply, from the supplier deck's media ────────────────────────
OURS = {
    "own_pouch_jasmine":  "image9.jpeg",     # BSCM Jasmine pouch
    "own_pouch_mexican":  "image10.jpeg",    # BSCM Special Mexican pouch
    "own_pouch_spicy":    "image12.jpeg",    # BSCM Hot & Spicy pouch
    "own_chefrey_turm":   "image4.jpeg",     # Chefrey organic turmeric basmati
    "own_chefrey_garlic": "image5.jpeg",     # Chefrey organic garlic fried rice
    "own_chefrey_lime":   "image6.jpeg",     # Chefrey organic cilantro lime
    "own_chefrey_jasm":   "image7.jpeg",     # Chefrey organic jasmine
    "own_cup":            "image13.jpeg",    # BSCM single/double cup
    "own_cup2":           "image21.jpeg",
    "own_spar_cups":      "image37.png",     # SPAR private-label 2x125 cups
    "own_kbros_kimchi":   "image42.jpg",    # KBROS Kimchi Fried Rice tray
    "own_kbros_veg":      "image43.jpg",    # KBROS Vegetable Fried Rice tray
    "own_kbros_japchae":  "image44.jpg",    # KBROS Japchae Fried Rice tray
    "own_kbros_white":    "image45.jpg",    # KBROS Sticky White Rice tray
    "own_kbros_line":     "image41.jpg",     # KBROS line-up
}
for name, f in OURS.items():
    pack(f"{M}/{f}", f"assets/{name}.jpg")

# ── full-bleed backgrounds ────────────────────────────────────────────────
for src, dst in [(f"{M}/image1.jpeg", "assets/bg_field.jpg"),
                 (f"{M}/image8.jpeg", "assets/bg_plant.jpg")]:
    if os.path.exists(src):
        im = Image.open(src).convert("RGB")
        im = ImageEnhance.Color(im).enhance(0.96)
        im.save(dst, quality=92)
        print(f"{dst:38} {im.size}")
