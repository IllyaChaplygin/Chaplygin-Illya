# -*- coding: utf-8 -*-
"""Prepare company assets pulled from the suppliers' own websites.

Only material from a supplier's own site is used for that supplier, so nothing
on a profile slide can be misattributed. TMK and ZEK sites were unreachable, so
they keep the KOKIRI logo from the quotation and a typographic mark.
"""
import os

import numpy as np
from PIL import Image, ImageChops, ImageEnhance, ImageFilter, ImageOps

SK = '/tmp/web/sk'
TN = '/tmp/web/tn'
TMK = '/tmp/web/tmk'
KK = '/tmp/web/kk'
DST = '/home/user/Chaplygin-Illya/presentation/src/photo'


def trim_white(im, tol=246):
    g = im.convert('L')
    diff = ImageChops.difference(g, Image.new('L', im.size, 255))
    bbox = diff.point(lambda p: 255 if p > (255 - tol) else 0).getbbox()
    return im.crop(bbox) if bbox else im


def drop_black(im, thresh=52):
    """Whiten a black studio backdrop without eating the dark nori in the pack.

    Flood-filling inward from the corners only clears background that actually
    touches the edge, so the near-black seaweed inside the bag survives.
    """
    from PIL import ImageDraw
    im = im.convert('RGB')
    for xy in [(0, 0), (im.width - 1, 0), (0, im.height - 1),
               (im.width - 1, im.height - 1), (im.width // 2, 0),
               (im.width // 2, im.height - 1)]:
        ImageDraw.floodfill(im, xy, (255, 255, 255), thresh=thresh)
    return im


def drop_grey(im, lum=182, sat=26):
    """Clear a flat grey/gradient backdrop without touching coloured artwork."""
    a = np.asarray(im.convert('RGB')).astype(int)
    mx, mn = a.max(axis=2), a.min(axis=2)
    bg = (mx - mn <= sat) & (a.mean(axis=2) >= lum)
    a[bg] = 255
    return Image.fromarray(a.astype('uint8'))


def pad(im, frac=0.03):
    p = max(4, int(frac * max(im.size)))
    out = Image.new('RGB', (im.width + 2 * p, im.height + 2 * p), (255, 255, 255))
    out.paste(im, (p, p))
    return out


def save(im, name, cap=1600):
    if max(im.size) > cap:
        s = cap / max(im.size)
        im = im.resize((int(im.width * s), int(im.height * s)), Image.LANCZOS)
    im.convert('RGB').save('%s/%s.png' % (DST, name))
    print('%-22s %sx%s' % (name, im.width, im.height))


def main():
    os.makedirs(DST, exist_ok=True)

    # --- Singha Kameda: corporate lockup sits at the left of a banner strip
    strip = Image.open('%s/-01.jpg' % SK).convert('RGB')
    a = np.asarray(strip.convert('L')).astype(int)
    ink = [i for i, v in enumerate((a < 150).sum(axis=0)) if v > 2]
    lock = strip.crop((max(0, ink[0] - 6), 0, min(strip.width, ink[-1] + 6),
                       strip.height))
    save(pad(trim_white(drop_grey(lock))), 'logo_singha')

    plant = Image.open('%s/banner_manufacturing.png' % SK).convert('RGB')
    plant = ImageEnhance.Color(plant).enhance(1.10)
    save(ImageEnhance.Contrast(plant).enhance(1.06), 'plant_singha')

    for name, f in [('cert_haccp', 'BSI-Assurance-Mark-HACCP-and-GMP-Red.jpg'),
                    ('cert_brc', 'brc.jpg'), ('cert_halal', 'halal.jpg'),
                    ('cert_nsf', 'nsf.jpg'), ('cert_organic', 'organic.jpg'),
                    ('cert_sedex', 'sedex.jpg')]:
        save(pad(trim_white(Image.open('%s/%s' % (SK, f)).convert('RGB')), 0.02),
             name)

    # --- Thai-Nichi: Mizuho brand mark and their own styled product photography
    save(trim_white(Image.open(
        '%s/691c0a_beedd78c99264da6ac960f47ce13b536_mv2.png' % TN).convert('RGB')),
        'logo_mizuho')
    save(Image.open('%s/691c0a_12106c5a50294e338f36ea47980f7ce0_mv2.jpg' % TN),
         'plant_thainichi')

    # --- TMK: their plant, from the KOKIRI site
    plant = Image.open(
        '%s/279558619-681682616275854-6421290885006178567-n.jpg' % TMK).convert('RGB')
    plant = ImageEnhance.Color(plant).enhance(1.12)
    save(ImageEnhance.Contrast(plant).enhance(1.05), 'plant_tmk')

    # KOKIRI pack shots at 1080-2000px, replacing the 43px slices from the
    # quotation. The Wow Seaweed shots are photographed on black, which would
    # punch a hole in a cream card, so the backdrop is flooded out.
    for name, f in [('tmk_sw_orig', '2025_07_1-2.png'),
                    ('tmk_sw_spicy', '2025_07_2.png'),
                    ('tmk_sw_squid', '2025_07_4-1.png')]:
        save(pad(trim_white(drop_black(Image.open('%s/%s' % (KK, f)))), 0.02), name)
    for name, f in [('tmk_mini_orig', '2025_10_1-6.png'),
                    ('tmk_mini_spicy', '2025_10_2-7.png')]:
        im = Image.open('%s/%s' % (KK, f)).convert('RGB')
        # the top third of these is the site's own KOKIRI banner, not the pack
        im = im.crop((0, int(im.height * 0.27), im.width, im.height))
        save(pad(trim_white(im), 0.03), name)

    # --- ZEK: no photography of the plant is published anywhere we can verify,
    # and their own pack art blurred up still read as smudged lettering. So the
    # column gets a made surface instead — roasted-nori grain in the supplier's
    # terracotta. It is decoration, and looks like decoration.
    save(nori_surface(), 'texture_zek')


def nori_surface(w=1100, h=1450, seed=11):
    """A made surface for the ZEK column — deep terracotta, matte, barely there.

    Reads as pressed paper under raking light: a long diagonal fall-off carries
    the panel, a wide soft mottle breaks the flatness, and a fine horizontal
    grain keeps it from looking like a gradient fill. All three are held at low
    amplitude on purpose — the white brand plate has to stay the loud thing.
    """
    rng = np.random.default_rng(seed)
    yy, xx = np.mgrid[0:h, 0:w]

    # light falls from the top-left corner and drops away across the panel
    field = 1.0 - (xx / w * 0.42 + yy / h * 0.58)
    field = 0.20 + 0.72 * field ** 1.25

    mottle = Image.fromarray(
        np.clip(rng.normal(0.5, 0.5, (h // 90, w // 90)) * 255, 0, 255).astype('uint8'), 'L')
    mottle = mottle.resize((w, h), Image.BICUBIC).filter(ImageFilter.GaussianBlur(60))
    field += (np.asarray(mottle).astype(float) / 255 - 0.5) * 0.14

    grain = rng.normal(0, 1, (h, w // 4))
    grain = np.asarray(Image.fromarray(
        np.clip(grain * 40 + 128, 0, 255).astype('uint8'), 'L')
        .resize((w, h), Image.BILINEAR)).astype(float)
    field += (grain / 255 - 0.5) * 0.07

    grey = Image.fromarray(np.clip(field * 255, 0, 255).astype('uint8'), 'L')
    return ImageOps.colorize(grey, black=(0x4E, 0x16, 0x0C), white=(0xCB, 0x55, 0x33))


if __name__ == '__main__':
    main()
