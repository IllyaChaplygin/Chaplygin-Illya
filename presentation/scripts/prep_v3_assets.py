"""v3 asset touch-ups: richer colour grade on the Mishima hero, a soft studio
glow behind the Takaokaya bottle illustration. Same crops as v2, just graded
for the punchier v3 palette instead of sitting flat next to it.
"""
import sys

from PIL import Image, ImageDraw, ImageEnhance, ImageFilter

RATIO = 4.58 / 7.5

# ---- Mishima: same crop as v2, pushed further (contrast/saturation + vignette)
im = Image.open("src/x734.png").convert("RGB")
w, h = im.size
nh = h
nw = round(nh * RATIO)
x0 = round((w - nw) * 0.52)
crop = im.crop((x0, 0, x0 + nw, h)).resize((1200, round(1200 / RATIO)), Image.LANCZOS)

crop = ImageEnhance.Color(crop).enhance(1.16)
crop = ImageEnhance.Contrast(crop).enhance(1.12)
crop = ImageEnhance.Brightness(crop).enhance(1.01)
crop = crop.filter(ImageFilter.UnsharpMask(radius=1.6, percent=75, threshold=3))

# subtle vignette: darken the corners so the bowl/spoon reads as the lit subject
vig = Image.new("L", crop.size, 0)
vd = ImageDraw.Draw(vig)
vd.ellipse([-crop.width * 0.28, -crop.height * 0.14,
            crop.width * 1.28, crop.height * 1.10], fill=255)
vig = vig.filter(ImageFilter.GaussianBlur(crop.width * 0.12))
dark = ImageEnhance.Brightness(crop).enhance(0.72)
crop = Image.composite(crop, dark, vig)

crop.save("assets/mishima_profile_hero.jpg", quality=94)
print("mishima_profile_hero.jpg", crop.size, round(crop.width / crop.height, 3))
