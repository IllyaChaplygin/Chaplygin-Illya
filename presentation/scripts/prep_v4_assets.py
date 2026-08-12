"""Circular product portraits for the v4 "gastronomy magazine" redesign.

Kinfolk/Bon Appétit/Cherry Bombe layouts favour a circular product portrait
over a hard-edged rectangle — it reads as a considered editorial crop rather
than a UI card. Pre-mask each SKU photo to a soft-edged circle (alpha PNG) so
it drops straight into pptx with no shape-fill XML surgery.
"""
import os

from PIL import Image, ImageDraw, ImageFilter

os.makedirs("assets/circles", exist_ok=True)

SIZE = 900          # working resolution; pptx scales down, keeps it crisp
FEATHER = 3          # soft anti-aliased edge


def circle_crop(src, dst, zoom=1.0, size=SIZE):
    im = Image.open(src).convert("RGB")
    s = min(im.size)
    cx, cy = im.width / 2, im.height / 2
    half = s / 2 / zoom
    im = im.crop((cx - half, cy - half, cx + half, cy + half)).resize((size, size), Image.LANCZOS)

    mask = Image.new("L", (size * 4, size * 4), 0)
    ImageDraw.Draw(mask).ellipse([0, 0, size * 4, size * 4], fill=255)
    mask = mask.resize((size, size), Image.LANCZOS)

    out = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    out.paste(im, (0, 0))
    out.putalpha(mask)
    out.save(dst)
    print(dst, out.size)


circle_crop("assets/sku_ms_shrimp.png", "assets/circles/ms_shrimp.png", zoom=1.06)
circle_crop("assets/sku_ms_nori.png", "assets/circles/ms_nori.png", zoom=1.06)
circle_crop("assets/sku_ms_wasabi.png", "assets/circles/ms_wasabi.png", zoom=1.06)
circle_crop("assets/sku_tk_wasabi.png", "assets/circles/tk_wasabi.png", zoom=1.02)
