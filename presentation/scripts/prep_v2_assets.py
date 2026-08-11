"""Assets for the v2 redesign: full-bleed profile hero photos.

Mishima: a tall, tight crop of the rice-bowl shot for the right-column
full-bleed panel (ratio 4.58:7.5 = 0.611).
Takaokaya: no real photography exists anywhere reachable, so the profile
hero is a large, moody vector rendering of two bottles on a dark ground —
honestly a visualization, not a photo, and captioned as such on the slide.
"""
import sys

from PIL import Image, ImageEnhance, ImageFilter

RATIO = 4.58 / 7.5  # 0.6107

# ---- Mishima: tall crop of the rice-bowl hero, full bleed --------------
im = Image.open("src/x734.png").convert("RGB")   # 958 x 1436, ratio 0.667
w, h = im.size
nh = h
nw = round(nh * RATIO)
x0 = round((w - nw) * 0.52)
crop = im.crop((x0, 0, x0 + nw, h))
crop = crop.resize((1100, round(1100 / RATIO)), Image.LANCZOS)
crop = ImageEnhance.Color(crop).enhance(1.05)
crop = ImageEnhance.Contrast(crop).enhance(1.04)
crop = crop.filter(ImageFilter.UnsharpMask(radius=1.5, percent=60, threshold=3))
crop.save("assets/mishima_profile_hero.jpg", quality=93)
print("mishima_profile_hero.jpg", crop.size, round(crop.width / crop.height, 3))
