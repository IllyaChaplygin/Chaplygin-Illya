"""Circular product portraits for every supplier — ZEK included.

The deck previously ran two visual languages: ZEK on the original cream cards,
Mishima/Takaokaya on a redesigned system. Under PANTRY_LIGHT there is one
language, so every pack shot in the deck gets the same treatment: trimmed to
its product, centred, masked to a perfect circle with an anti-aliased edge.
"""
import os

import numpy as np
from PIL import Image, ImageDraw, ImageFilter

os.makedirs("assets/circles", exist_ok=True)
SIZE = 900


def trim_white(im, thr=243, pad_frac=0.04):
    """Crop away the flat white studio surround so the product fills the circle."""
    a = np.asarray(im.convert("RGB")).min(axis=2)
    ys, xs = np.where(a < thr)
    if len(xs) == 0:
        return im
    x0, x1, y0, y1 = xs.min(), xs.max(), ys.min(), ys.max()
    m = int(pad_frac * max(x1 - x0, y1 - y0))
    return im.crop((max(0, x0 - m), max(0, y0 - m),
                    min(im.width, x1 + m), min(im.height, y1 + m)))


def circle_portrait(src, dst, trim=True, fill_frac=0.94, bg=(255, 255, 255)):
    """Product centred on white, masked to a circle. fill_frac controls how much
    of the circle the product occupies — kept constant so a row of portraits
    reads as one family rather than as photos of assorted sizes."""
    im = Image.open(src).convert("RGB")
    if trim:
        im = trim_white(im)

    inner = round(SIZE * fill_frac)
    scale = min(inner / im.width, inner / im.height)
    im = im.resize((max(1, round(im.width * scale)), max(1, round(im.height * scale))),
                   Image.LANCZOS)
    im = im.filter(ImageFilter.UnsharpMask(radius=1.4, percent=60, threshold=3))

    canvas = Image.new("RGB", (SIZE, SIZE), bg)
    canvas.paste(im, ((SIZE - im.width) // 2, (SIZE - im.height) // 2))

    mask = Image.new("L", (SIZE * 4, SIZE * 4), 0)
    ImageDraw.Draw(mask).ellipse([0, 0, SIZE * 4, SIZE * 4], fill=255)
    mask = mask.resize((SIZE, SIZE), Image.LANCZOS)

    out = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    out.paste(canvas, (0, 0))
    out.putalpha(mask)
    out.save(dst)
    print(dst)


M = "unpacked/ppt/media"

# ZEK — pack shots pulled from the original deck's media
circle_portrait(f"{M}/image4.png", "assets/circles/zek_chicken35.png")
circle_portrait(f"{M}/image3.png", "assets/circles/zek_veg35.png")
circle_portrait(f"{M}/image5.png", "assets/circles/zek_sesame35.png")
circle_portrait(f"{M}/image6.png", "assets/circles/zek_veg70.png")
circle_portrait(f"{M}/image7.png", "assets/circles/zek_sesame70.png")
circle_portrait(f"{M}/image8.png", "assets/circles/zek_chicken70.png")
circle_portrait(f"{M}/image2.png", "assets/circles/zek_tempura.png", trim=False, fill_frac=0.98)
circle_portrait(f"{M}/image1.png", "assets/circles/zek_nori.png", fill_frac=0.99)

# Mishima / Takaokaya — same treatment, rebuilt from the untrimmed sources
circle_portrait("assets/sku_ms_shrimp.png", "assets/circles/ms_shrimp.png")
circle_portrait("assets/sku_ms_nori.png", "assets/circles/ms_nori.png")
circle_portrait("assets/sku_ms_wasabi.png", "assets/circles/ms_wasabi.png")
circle_portrait("assets/sku_tk_wasabi.png", "assets/circles/tk_wasabi.png", fill_frac=0.98)
