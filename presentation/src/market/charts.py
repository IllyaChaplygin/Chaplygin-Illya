# -*- coding: utf-8 -*-
"""Chart primitives for the market-research deck: horizontal bars and a price
ladder. Kept separate from theme.py since the product catalogue deck never
needed them — this is the one deck in the repo that does.
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'deck'))
from theme import (BODY, HEAD, INK, MUTED, RULE, WHITE, deepen, label, mix,  # noqa: E402
                   rect, text, tint)
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN  # noqa: E402


def hbar_chart(s, x, y, w, h, rows, max_val, value_fmt=lambda v: '%.1f' % v,
              name_w=2.35, val_w=0.62):
    """rows: list of dict(name, value, color, sub=None). One bar per row,
    label left, value right, track width proportional to max_val."""
    track_x = x + name_w
    track_w = w - name_w - val_w
    row_h = h / len(rows)
    for i, r in enumerate(rows):
        ry = y + i * row_h
        cy = ry + row_h / 2
        text(s, x, ry, name_w - 0.08, row_h, r['name'], size=8.2, color=INK,
             bold=True, anchor=MSO_ANCHOR.MIDDLE)
        if r.get('sub'):
            text(s, x, ry + row_h * 0.52, name_w - 0.08, row_h * 0.42, r['sub'],
                 size=6.3, color=MUTED, anchor=MSO_ANCHOR.MIDDLE)
        bar_h = row_h * 0.46
        rect(s, track_x, cy - bar_h / 2, track_w, bar_h, fill=tint(INK, 0.06),
             radius=0.4)
        bw = max(0.04, track_w * min(1.0, r['value'] / max_val))
        rect(s, track_x, cy - bar_h / 2, bw, bar_h, fill=r['color'], radius=0.4)
        text(s, track_x + bw + 0.08, ry, val_w + 0.4, row_h, value_fmt(r['value']),
             size=8.5, font=HEAD, bold=True, color=deepen(r['color'], 0.92),
             anchor=MSO_ANCHOR.MIDDLE)


def price_ladder(s, x, y, w, h, points, lo, hi, axis_label='₴ / г'):
    """points: list of dict(label, value, color). Rows are assigned
    automatically (greedy interval packing on value, not caller input) so
    labels whose values are too close to share a baseline stack instead of
    overlapping."""
    axis_y = y + h - 0.30
    rect(s, x, axis_y, w, 0.014, fill=RULE)
    lw = 1.35
    min_gap = lw / w * (hi - lo) * 0.95
    row_last = []
    placed = []
    for p in sorted(points, key=lambda p: p['value']):
        for ri, last in enumerate(row_last):
            if p['value'] - last >= min_gap:
                row_last[ri] = p['value']
                placed.append((p, ri))
                break
        else:
            row_last.append(p['value'])
            placed.append((p, len(row_last) - 1))
    row_gap = (h - 0.46) / max(len(row_last), 1)
    for p, ri in placed:
        px = x + (p['value'] - lo) / (hi - lo) * w
        py = axis_y - 0.10 - ri * row_gap
        d = 0.11
        rect(s, px - d / 2, py - d / 2, d, d, fill=p['color'], radius=0.5)
        rect(s, px - 0.006, py, 0.012, axis_y - py, fill=mix(WHITE, p['color'], 0.55))
        text(s, px - lw / 2, py - 0.30, lw, 0.18, p['label'], size=6.6, bold=True,
             color=deepen(p['color'], 0.9), align=PP_ALIGN.CENTER)
        text(s, px - lw / 2, py - 0.14, lw, 0.14,
             ('%.1f ₴/г' % p['value']), size=6.2, color=MUTED, align=PP_ALIGN.CENTER)
    ticks = 5
    for i in range(ticks + 1):
        v = lo + (hi - lo) * i / ticks
        tx = x + w * i / ticks
        rect(s, tx, axis_y, 0.012, 0.07, fill=RULE)
        text(s, tx - 0.35, axis_y + 0.09, 0.70, 0.14, '%.0f' % v, size=6.5,
             color=MUTED, align=PP_ALIGN.CENTER)
    label(s, x, axis_y + 0.24, w, axis_label, color=MUTED, size=6.5,
         align=PP_ALIGN.CENTER)
