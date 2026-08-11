"""Zoom the cut-out photos onto their subject while keeping the exact card ratio."""
import numpy as np
from PIL import Image, ImageFilter


def zoom_to_subject(path, target_ratio, margin=0.06, out_w=1200):
    im = Image.open(path).convert("RGB")
    a = np.asarray(im).min(axis=2)
    ys, xs = np.where(a < 244)                     # anything that is not paper white
    x0, x1, y0, y1 = xs.min(), xs.max(), ys.min(), ys.max()
    m = int(margin * max(x1 - x0, y1 - y0))
    x0, y0 = max(0, x0 - m), max(0, y0 - m)
    x1, y1 = min(im.width, x1 + m), min(im.height, y1 + m)

    cw, ch = x1 - x0, y1 - y0
    if cw / ch < target_ratio:                     # widen
        cw = ch * target_ratio
    else:                                          # heighten
        ch = cw / target_ratio
    cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
    left, top = cx - cw / 2, cy - ch / 2
    left = min(max(0, left), max(0, im.width - cw))
    top = min(max(0, top), max(0, im.height - ch))

    if cw > im.width or ch > im.height:            # pad on white when the crop overflows
        canvas = Image.new("RGB", (max(im.width, int(cw)), max(im.height, int(ch))), "white")
        canvas.paste(im, ((canvas.width - im.width) // 2, (canvas.height - im.height) // 2))
        im = canvas
        left = (im.width - cw) / 2
        top = (im.height - ch) / 2

    im = im.crop((round(left), round(top), round(left + cw), round(top + ch)))
    im = im.resize((out_w, round(out_w / target_ratio)), Image.LANCZOS)
    im = im.filter(ImageFilter.UnsharpMask(radius=1.2, percent=45, threshold=3))
    im.save(path, quality=93)
    print(path, im.size, round(im.width / im.height, 3))


zoom_to_subject("assets/mishima_ochazuke.jpg", 2.83 / 1.78, margin=0.05, out_w=1250)
zoom_to_subject("assets/mishima_ochazuke_small.jpg", 1.53 / 1.18, margin=0.05, out_w=950)
zoom_to_subject("assets/mishima_sauces.jpg", 1.53 / 1.18, margin=0.05, out_w=950)
