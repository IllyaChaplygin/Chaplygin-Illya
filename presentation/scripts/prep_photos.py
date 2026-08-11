"""Prepare Mishima photos from the retail catalog for the deck."""
import numpy as np
from collections import deque
from PIL import Image, ImageFilter, ImageEnhance

OUT = "assets"
import os
os.makedirs(OUT, exist_ok=True)


def key_out_dark(im, bg=(255, 255, 255), thresh=42, feather=1.6):
    """Flood-fill the dark studio background from the borders and replace it."""
    a = np.asarray(im.convert("RGB")).astype(np.int16)
    h, w = a.shape[:2]
    lum = a.max(axis=2)
    dark = lum <= thresh
    seen = np.zeros((h, w), bool)
    dq = deque()
    for x in range(w):
        for y in (0, h - 1):
            if dark[y, x] and not seen[y, x]:
                seen[y, x] = True
                dq.append((y, x))
    for y in range(h):
        for x in (0, w - 1):
            if dark[y, x] and not seen[y, x]:
                seen[y, x] = True
                dq.append((y, x))
    while dq:
        y, x = dq.popleft()
        for dy, dx in ((1, 0), (-1, 0), (0, 1), (0, -1)):
            ny, nx = y + dy, x + dx
            if 0 <= ny < h and 0 <= nx < w and dark[ny, nx] and not seen[ny, nx]:
                seen[ny, nx] = True
                dq.append((ny, nx))
    mask = Image.fromarray((seen * 255).astype(np.uint8))
    mask = mask.filter(ImageFilter.GaussianBlur(feather))
    out = Image.new("RGB", im.size, bg)
    out.paste(im.convert("RGB"), (0, 0), Image.eval(mask, lambda v: 255 - v))
    return out


def crop_ratio(im, ratio, anchor_y=0.5, anchor_x=0.5):
    """Center-ish crop to the given w/h ratio."""
    w, h = im.size
    if w / h > ratio:
        nw, nh = int(h * ratio), h
    else:
        nw, nh = w, int(w / ratio)
    x0 = int((w - nw) * anchor_x)
    y0 = int((h - nh) * anchor_y)
    return im.crop((x0, y0, x0 + nw, y0 + nh))


def finish(im, target_w, sharpen=1.25, saturation=1.04):
    if im.width != target_w:
        im = im.resize((target_w, round(im.height * target_w / im.width)), Image.LANCZOS)
    im = im.filter(ImageFilter.UnsharpMask(radius=1.4, percent=int(60 * sharpen), threshold=3))
    im = ImageEnhance.Color(im).enhance(saturation)
    return im


src = {k: Image.open(f"src/x{k}.png") for k in (734, 769, 749, 662)}

# 1. Slide 4 hero — the boxed furikake line-up on a linen table (portrait source).
lineup = src[769]
w, h = lineup.size                       # 851 x 1276
hero4 = lineup.crop((0, 585, w, 1235))   # keep the boxes + branch tips + table edge
hero4 = crop_ratio(hero4, 3.28 / 2.38, anchor_y=0.35)
finish(hero4, 1400).save(f"{OUT}/mishima_lineup.jpg", quality=92)

# 1b. Wide crop of the same line-up for the cost slide's lower strip.
wide = lineup.crop((0, 600, w, 1120))
finish(crop_ratio(wide, 2.83 / 1.20, anchor_y=0.45), 1300).save(
    f"{OUT}/mishima_lineup_wide.jpg", quality=92)

# 2. Slide 5 hero — furikake sprinkled over rice (portrait source, light background).
rice = src[734]
hero5 = crop_ratio(rice, 2.83 / 2.18, anchor_y=0.52)
finish(hero5, 1400).save(f"{OUT}/mishima_rice.jpg", quality=92)

# 3. Ochazuke bowl and the sauce line-up — cut out of their black studio background.
och = key_out_dark(src[749])
och = crop_ratio(och, 2.83 / 1.78, anchor_y=0.5)
finish(och, 1250, saturation=1.06).save(f"{OUT}/mishima_ochazuke.jpg", quality=92)

sauces = key_out_dark(src[662])
sauces = crop_ratio(sauces, 1.53 / 1.18, anchor_y=0.48)
finish(sauces, 900, saturation=1.06).save(f"{OUT}/mishima_sauces.jpg", quality=92)

och_small = key_out_dark(src[749])
och_small = crop_ratio(och_small, 1.53 / 1.18, anchor_y=0.5)
finish(och_small, 900, saturation=1.06).save(f"{OUT}/mishima_ochazuke_small.jpg", quality=92)

# 4. Flavour pack shots for the cost-table legend (native ~155-190 px, upscaled x4).
import pymupdf, io
doc = pymupdf.open("catalog.pdf")
packs = {"shrimp": 508, "wasabi": 512, "nori": 502}
for name, xref in packs.items():
    raw = doc.extract_image(xref)
    im = Image.open(io.BytesIO(raw["image"])).convert("RGB")
    im = im.resize((im.width * 4, im.height * 4), Image.LANCZOS)
    im = im.filter(ImageFilter.UnsharpMask(radius=3, percent=95, threshold=2))
    a = np.asarray(im).astype(np.int16)          # lift the near-white paper to pure white
    a = np.clip((a - 6) * (255 / 242), 0, 255).astype(np.uint8)
    im = Image.fromarray(a)
    im = crop_ratio(im, 1.0)
    im.save(f"{OUT}/pack_{name}.png")
    print("pack", name, im.size)

for f in sorted(os.listdir(OUT)):
    print(f, Image.open(f"{OUT}/{f}").size)
