"""v6 assets — pack shots kept whole, factory photography prepared.

Correction from v5: the circular mask cropped the packs. Here every pack shot
keeps its own proportions — the white studio surround is trimmed away, nothing
of the product is cut, and the slide places it inside a blush tile with air
around it.
"""
import os

import numpy as np
from PIL import Image, ImageEnhance, ImageFilter

os.makedirs("assets/packs", exist_ok=True)
os.makedirs("assets/hero", exist_ok=True)
M = "unpacked/ppt/media"


def trim(im, thr=243, pad=0.02):
    a = np.asarray(im.convert("RGB")).min(axis=2)
    ys, xs = np.where(a < thr)
    if not len(xs):
        return im
    x0, x1, y0, y1 = xs.min(), xs.max(), ys.min(), ys.max()
    m = int(pad * max(x1 - x0, y1 - y0))
    return im.crop((max(0, x0 - m), max(0, y0 - m),
                    min(im.width, x1 + m), min(im.height, y1 + m)))


def pack(src, dst, target_h=1000):
    """Whole pack shot on white, trimmed to its own edges — never cropped."""
    im = trim(Image.open(src).convert("RGB"))
    scale = target_h / im.height
    im = im.resize((max(1, round(im.width * scale)), target_h), Image.LANCZOS)
    im = im.filter(ImageFilter.UnsharpMask(radius=1.4, percent=65, threshold=3))
    im.save(dst)
    print(dst, im.size, f"ratio {im.width / im.height:.2f}")


# ZEK — from the original deck's own media
pack(f"{M}/image4.png", "assets/packs/zek_chicken35.png")
pack(f"{M}/image3.png", "assets/packs/zek_veg35.png")
pack(f"{M}/image5.png", "assets/packs/zek_sesame35.png")
pack(f"{M}/image6.png", "assets/packs/zek_veg70.png")
pack(f"{M}/image7.png", "assets/packs/zek_sesame70.png")
pack(f"{M}/image8.png", "assets/packs/zek_chicken70.png")

# Mishima / Takaokaya
pack("assets/sku_ms_shrimp.png", "assets/packs/ms_shrimp.png")
pack("assets/sku_ms_nori.png", "assets/packs/ms_nori.png")
pack("assets/sku_ms_wasabi.png", "assets/packs/ms_wasabi.png")
pack("assets/sku_tk_wasabi.png", "assets/packs/tk_wasabi.png")


def hero(src, dst, w=2200, sat=1.05, con=1.04):
    im = Image.open(src).convert("RGB")
    if im.width != w:
        im = im.resize((w, round(im.height * w / im.width)), Image.LANCZOS)
    im = ImageEnhance.Color(im).enhance(sat)
    im = ImageEnhance.Contrast(im).enhance(con)
    im = im.filter(ImageFilter.UnsharpMask(radius=1.5, percent=55, threshold=3))
    im.save(dst, quality=93)
    print(dst, im.size)


# real Mishima plant photography, from the manufacturer's own site
hero("factory/ms_08.jpg", "assets/hero/mishima_plant.jpg")      # facade + logo
hero("factory/ms_06.jpg", "assets/hero/mishima_line.jpg")       # production line
hero("factory/ms_07.jpg", "assets/hero/mishima_aerial.jpg")     # aerial of the works
