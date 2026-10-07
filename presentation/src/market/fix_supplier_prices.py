# -*- coding: utf-8 -*-
"""Make every cost figure on supplier slides 3-16 equal to the SelfCost workbook.

Each product card has four rows (40' container, Збірний 34 м³, 20' container, Збірний 17 м³) with a
$ box and a grn box. The boxes are located by their row labels, the product by the card order
(same order as the SelfCost blocks), and only the run text is replaced — fonts, colours and
positions stay exactly as the user left them. $ is rounded to cents, grn = $ x 45 (one rate for
every row; the Squid row in SelfCost carries 46 by mistake)."""
import json, os, re
HERE = os.path.dirname(os.path.abspath(__file__))
DATA = json.load(open(os.path.join(HERE, '..', 'data.json'), encoding='utf-8'))
RATE = 45.0
EMU = 914400.0

# slide number -> (sheet, [index of the product on each card, left to right])
SLIDES = {
    3: ('SINGHA KAMEDA - Retail', [0, 1]),
    4: ('SINGHA KAMEDA - Retail+Bulk', [0, 1, 2, 3]),
    6: ('Thai-Nichi - Retail', [0, 1]),
    7: ('Thai-Nichi - Retail+Bulk', [0, 1, 2, 3]),
    9: ('TMK Thailand Co., Ltd -Retail', [0, 1, 2, 3]),
    10: ('TMK Thailand Co., Ltd -Retail', [4, 5, 6]),
    11: ('TMK Thailand Co., Ltd -Retail', [7, 8, 9]),
    13: ('ZEK -Retail', [0, 1, 2, 3]),
    14: ('ZEK -Retail', [4, 5, 6, 7]),
    15: ('ZEK -Retail', [8, 9, 10]),
    16: ('ZEK -Retail', [11, 12, 13]),
}
ROWS = [("40'", '40′ контейнер'), ('LCL 34', 'збірний 34 м³'), ("20'", '20′ контейнер'), ('LCL 17', 'збірний 17 м³')]


def usd_txt(v):
    return ('$%.2f' % v).replace('.', ',')


def uah_txt(v):
    return ('%.2f ₴' % v).replace('.', ',')


def set_text(shape, new):
    """Replace the visible text but keep the first run's formatting."""
    p = shape.text_frame.paragraphs[0]
    runs = p.runs
    runs[0].text = new
    for r in runs[1:]:
        r._r.getparent().remove(r._r)
    for extra in shape.text_frame.paragraphs[1:]:
        extra._p.getparent().remove(extra._p)


def fix(prs, log=None):
    changed = []
    for n, (sheet, idxs) in SLIDES.items():
        s = prs.slides[n - 1]
        shapes = [sh for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.strip()]
        heads = sorted([sh for sh in shapes if sh.text_frame.text.strip().lower() == '40′ контейнер'], key=lambda a: a.left)
        assert len(heads) == len(idxs), (n, len(heads), len(idxs))
        for k, (head, idx) in enumerate(zip(heads, idxs)):
            xmax = heads[k + 1].left / EMU - 0.1 if k + 1 < len(heads) else 99.0
            cost = DATA[sheet][idx]['cost']
            x0 = head.left / EMU
            region = [sh for sh in shapes if x0 - 0.1 <= sh.left / EMU < xmax and sh.top / EMU >= head.top / EMU - 0.05]
            labels = {}
            for sh in region:
                t = sh.text_frame.text.strip().lower()
                for key, lab in ROWS:
                    if t == lab and key not in labels and not (key == "40'" and sh is not head):
                        labels[key] = sh
            assert len(labels) == 4, (n, x0, sorted(labels))
            for key, lab in ROWS:
                lb = labels[key]
                ly = lb.top / EMU
                if key == "40'":
                    cand = [sh for sh in region if ly < sh.top / EMU < ly + 0.35 and re.fullmatch(r'\$[\d.,]+|[\d\s.,]+\s*₴', sh.text_frame.text.strip())]
                else:
                    cand = [sh for sh in region if abs(sh.top / EMU - ly) < 0.07 and sh.left > lb.left and re.fullmatch(r'\$[\d.,]+|[\d\s.,]+\s*₴', sh.text_frame.text.strip())]
                usd_box = [c for c in cand if c.text_frame.text.strip().startswith('$')]
                uah_box = [c for c in cand if c.text_frame.text.strip().endswith('₴')]
                assert len(usd_box) == 1 and len(uah_box) == 1, (n, x0, key, [c.text_frame.text for c in cand])
                v = cost[key]['usd']
                for box, new in ((usd_box[0], usd_txt(v)), (uah_box[0], uah_txt(v * RATE))):
                    old = box.text_frame.text.strip()
                    if old != new:
                        changed.append((n, DATA[sheet][idx]['name'][:36], key, old, new))
                        set_text(box, new)
    if log is not None:
        log.extend(changed)
    return changed
