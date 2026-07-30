# -*- coding: utf-8 -*-
"""Prepare company assets pulled from the suppliers' own websites.

Only material from a supplier's own site is used for that supplier, so nothing
on a profile slide can be misattributed. TMK and ZEK sites were unreachable, so
they keep the KOKIRI logo from the quotation and a typographic mark.
"""
import os

import numpy as np
from PIL import Image, ImageChops, ImageEnhance

SK = '/tmp/web/sk'
TN = '/tmp/web/tn'
TMK = '/tmp/web/tmk'
KK = '/tmp/web/kk'
KK3 = '/tmp/web/kk3'
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

    # Double Roll: the site's own marketing collages carry a clean box shot at
    # the left before the text and repeated-pouch panels start. 600px source,
    # a real jump from the 69-73px quotation crop.
    for name, f in [('tmk_dbl_orig', 'dbl_orig.jpg'), ('tmk_dbl_spicy', 'dbl_spicy.jpg')]:
        im = Image.open('%s/%s' % (KK3, f)).convert('RGB').crop((15, 175, 180, 468))
        save(im, name)

    # ZEK gets no image at all: no verified photography of the company exists,
    # and a generated stand-in for one reads as fake no matter how it's made.
    # Its profile column is built as pure typography in build.py instead.


if __name__ == '__main__':
    main()
