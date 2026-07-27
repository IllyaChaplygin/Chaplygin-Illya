"""Prepare product photos for the deck.

De-background, slice multi-flavour line shots, fix mirroring, and grade every
shot: the supplier catalogue images are flat, low-contrast scans that look dull
on a light slide.
"""
import os
from collections import deque

from PIL import Image, ImageChops, ImageEnhance, ImageFilter, ImageOps

SRC = '/home/user/work/img'
DST = '/home/user/work/photo'
MAXPX = 1400
os.makedirs(DST, exist_ok=True)


def load(sheet, row):
    return Image.open('%s/%s__r%02d.png' % (SRC, sheet, row)).convert('RGB')


def whiten_border(im, thresh=70):
    """Flood-fill dark background inwards from the border and repaint it white."""
    im = im.copy()
    px = im.load()
    w, h = im.size
    seen = [[False] * w for _ in range(h)]
    q = deque()

    def dark(x, y):
        r, g, b = px[x, y]
        return r < thresh and g < thresh and b < thresh

    for x in range(w):
        for y in (0, h - 1):
            if not seen[y][x] and dark(x, y):
                seen[y][x] = True
                q.append((x, y))
    for y in range(h):
        for x in (0, w - 1):
            if not seen[y][x] and dark(x, y):
                seen[y][x] = True
                q.append((x, y))
    while q:
        x, y = q.popleft()
        px[x, y] = (255, 255, 255)
        for dx, dy in ((1, 0), (-1, 0), (0, 1), (0, -1)):
            nx, ny = x + dx, y + dy
            if 0 <= nx < w and 0 <= ny < h and not seen[ny][nx] and dark(nx, ny):
                seen[ny][nx] = True
                q.append((nx, ny))
    return im


def trim(im, tol=248):
    g = im.convert('L')
    diff = ImageChops.difference(g, Image.new('L', im.size, 255))
    bbox = diff.point(lambda p: 255 if p > (255 - tol) else 0).getbbox()
    return im.crop(bbox) if bbox else im


def pad(im, frac=0.04):
    w, h = im.size
    p = max(6, int(frac * max(w, h)))
    out = Image.new('RGB', (w + 2 * p, h + 2 * p), (255, 255, 255))
    out.paste(im, (p, p))
    return out


def slice_h(im, n, i, inset=0.0):
    """Take vertical slice i of n across the width."""
    w, h = im.size
    step = w / n
    x0 = int(step * (i + inset))
    x1 = int(step * (i + 1 - inset))
    return im.crop((x0, 0, x1, h))


def quad(im, col, row):
    w, h = im.size
    return im.crop((w * col // 2, h * row // 2, w * (col + 1) // 2, h * (row + 1) // 2))


def border_is_dark(im, thresh=70, frac=0.5):
    px = im.load()
    w, h = im.size
    edge = ([(x, 0) for x in range(w)] + [(x, h - 1) for x in range(w)] +
            [(0, y) for y in range(h)] + [(w - 1, y) for y in range(h)])
    dark = sum(1 for x, y in edge if max(px[x, y]) < thresh)
    return dark / len(edge) >= frac


def grade(im):
    """Lift the flat catalogue scans: clean whites, wider tonal range, richer colour.

    Small sources are JPEG-compressed thumbnails — on those, auto-contrast and
    unsharp masking amplify the block artefacts into coloured speckle, so they
    get a much gentler pass.
    """
    fragile = max(im.size) < 220
    if not fragile:
        # stretch each channel, ignoring the extremes so a stray pixel can't anchor it
        im = ImageOps.autocontrast(im, cutoff=(0.4, 0.2), preserve_tone=True)
        im = ImageEnhance.Color(im).enhance(1.22)
        im = ImageEnhance.Contrast(im).enhance(1.08)
        im = ImageEnhance.Brightness(im).enhance(1.03)
        im = im.filter(ImageFilter.UnsharpMask(radius=1.4, percent=55, threshold=4))
    else:
        im = ImageEnhance.Color(im).enhance(1.12)
        im = ImageEnhance.Contrast(im).enhance(1.04)
        im = ImageEnhance.Brightness(im).enhance(1.04)

    # anything already near-white becomes pure white so the cut-out reads clean
    px = im.load()
    w, h = im.size
    for y in range(h):
        for x in range(w):
            r, g, b = px[x, y]
            if r > 243 and g > 243 and b > 243:
                px[x, y] = (255, 255, 255)
    return im


def resample(im):
    """Only ever downscale.

    Enlarging here does not add detail — it just bakes in interpolation blur and
    hands PowerPoint a file whose pixel count no longer says how sharp the shot
    really is. The deck sizes each picture from its true resolution instead
    (see theme.picture), so the original pixels are what must be preserved.
    """
    if max(im.size) <= MAXPX:
        return im
    s = MAXPX / max(im.size)
    return im.resize((int(im.width * s), int(im.height * s)), Image.LANCZOS)


def save(im, name, dark_bg=False):
    if dark_bg or border_is_dark(im):
        im = whiten_border(im)
    im = pad(resample(grade(trim(im))), frac=0.02)
    im.save('%s/%s.png' % (DST, name))
    return name


HANJIN = 'HanJin_Food_Co_Ltd'
TMK = 'TMK_(Thailand)_Co_Ltd'
TN = 'Thai-Nichi_Industries_Co_Ltd'
SK = 'SINGHA_KAMEDA'


def main():
    # --- Singha Kameda: one 2x2 line shot -> Original (top-left), Wasabi (top-right)
    sk = load(SK, 3)
    # the line shot sits on dark plates that bleed into each quadrant; a higher
    # cut-off clears them without eating the (also dark) pack artwork
    save(whiten_border(quad(sk, 0, 0), thresh=110), 'sk_original')
    save(whiten_border(quad(sk, 1, 0), thresh=110), 'sk_wasabi')
    save(sk, 'sk_line')
    save(load(SK, 7), 'sk_okome')

    # --- Thai-Nichi: individual pack shots
    save(load(TN, 4), 'tn_original')
    save(load(TN, 5), 'tn_wasabi')
    save(load(TN, 8), 'tn_kakinotane')
    save(load(TN, 19), 'tn_chips')

    # --- TMK / KOKIRI. Brand colour code: green = Original, red = Spicy, purple = Squid.
    roll = whiten_border(load(TMK, 3))          # left->right: green, red, purple
    save(slice_h(roll, 3, 0), 'tmk_roll_orig')
    save(slice_h(roll, 3, 1), 'tmk_roll_spicy')
    save(slice_h(roll, 3, 2), 'tmk_roll_squid')

    dbl = whiten_border(load(TMK, 4))           # two boxes: green left, red right
    save(dbl.crop((0, 0, int(dbl.width * 0.56), dbl.height)), 'tmk_dbl_orig')
    save(dbl.crop((int(dbl.width * 0.42), 0, dbl.width, dbl.height)), 'tmk_dbl_spicy')

    sw = whiten_border(load(TMK, 6))            # left->right: red, green, purple
    save(slice_h(sw, 3, 1), 'tmk_sw_orig')
    save(slice_h(sw, 3, 0), 'tmk_sw_spicy')
    save(slice_h(sw, 3, 2), 'tmk_sw_squid')

    mini = whiten_border(load(TMK, 10))         # green left, red right
    save(slice_h(mini, 2, 0), 'tmk_mini_orig')
    save(slice_h(mini, 2, 1), 'tmk_mini_spicy')

    save(load(TMK, 3), 'tmk_line', dark_bg=True)

    # --- ZEK / HanJin
    zek = {
        'zek_tempura_corn30': 32, 'zek_tempura_wasabi30': 34, 'zek_tempura_mala30': 33,
        'zek_tempura_corn50': 35, 'zek_tempura_wasabi50': 36, 'zek_tempura_mala50': 37,
        'zek_sandwich_meat25': 25, 'zek_sandwich_sesame25': 26,
        'zek_topping_chicken35': 14, 'zek_topping_veg35': 15, 'zek_topping_sesame35': 16,
        'zek_topping_veg70': 19, 'zek_topping_sesame70': 20, 'zek_topping_chicken70': 21,
    }
    for name, row in zek.items():
        im = load(HANJIN, row)
        if row == 25:  # this shot is mirrored in the source workbook
            im = im.transpose(Image.FLIP_LEFT_RIGHT)
        save(im, name)
    save(load(HANJIN, 3), 'zek_classic')

    print('written', len(os.listdir(DST)), 'files ->', DST)


if __name__ == '__main__':
    main()
