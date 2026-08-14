"""Full-bleed opener images for the two suppliers with no plant photograph.

Mishima's opener uses the manufacturer's own factory photo. ZEK and Takaokaya
have none reachable, so their openers are composed still lifes of the actual
packs on the same linen ground — same frame, same crop, same caption slot, so
the three openers read as one sequence. Product sits right of centre because
the text panel occupies the left of the page.
"""
import sys

from PIL import Image, ImageDraw, ImageFilter

sys.path.insert(0, ".")

W, H = 2400, 1350                       # 16:9, generous for a 13.33in bleed
LINEN = (250, 247, 242)
BLUSH = (241, 221, 211)
SHADOW = (196, 176, 164)


def still_life(items, out, bg=BLUSH):
    """items: (path, centre-x frac, baseline-y frac, height frac)"""
    canvas = Image.new("RGB", (W, H), bg)
    shade = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    sd = ImageDraw.Draw(shade)

    placed = []
    for path, cxf, basef, hf in items:
        im = Image.open(path).convert("RGBA")
        th = round(H * hf)
        tw = max(1, round(im.width * th / im.height))
        im = im.resize((tw, th), Image.LANCZOS)
        cx, base = round(W * cxf), round(H * basef)
        x, y = cx - tw // 2, base - th
        sd.ellipse([x + tw * 0.02, base - th * 0.035, x + tw * 0.98, base + th * 0.055],
                   fill=SHADOW + (110,))
        placed.append((im, x, y))

    canvas.paste(Image.alpha_composite(
        Image.new("RGBA", (W, H), bg + (255,)),
        shade.filter(ImageFilter.GaussianBlur(26))).convert("RGB"), (0, 0))
    for im, x, y in placed:
        canvas.paste(im, (x, y), im)
    canvas.save(out, quality=94)
    print(out, canvas.size)


P = "assets/packs"
# Staged in the band right of the text panel (x > 0.55) and above the caption
# tab (y < 0.89). Widths are measured from the real pack ratio so the three
# only just kiss each other — a magazine trio, not a pile.
still_life([
    (f"{P}/zek_veg35.png",     0.634, 0.835, 0.36),
    (f"{P}/zek_sesame35.png",  0.921, 0.825, 0.34),
    (f"{P}/zek_chicken35.png", 0.779, 0.875, 0.44),   # front of the group
], "assets/hero/zek_still.jpg")

# Takaokaya: the drawn fresh-lock bottles, re-staged wide on the same ground
from make_bottles import bottle, render                                    # noqa: E402

BW, BH = 1333, 750
body = "".join([
    bottle("nori",   cx=BW * 0.634, base=BH * 0.790, h=BH * 0.40, seed=3,  grams="50 g"),
    bottle("yuzu",   cx=BW * 0.921, base=BH * 0.770, h=BH * 0.35, seed=11, grams="50 g"),
    bottle("wasabi", cx=BW * 0.779, base=BH * 0.845, h=BH * 0.47, seed=7,  grams="50 g"),
])
render("assets/hero/takaokaya_still.jpg", 13.33, 7.5, body, bg="#F1DDD3", dpi=180)
