# -*- coding: utf-8 -*-
"""Final deck with the user's own slides left completely untouched.

Slides going blank after PowerPoint's repair means it rejected the shape tree
of specific slides. Every edit I make to their slides — rewriting a run,
deleting a shape, stamping a badge — is a candidate. So this build makes
none of them: their 19 slides are carried over byte-for-byte and only
reordered, which touches presentation.xml alone and never a slide part.

The analysis lives entirely on the generated slides, which is where the
segmentation, the format work and the benchmark already are.
"""
from pptx import Presentation
import re
from pptx.util import Emu
import data as d
import deckkit as K
import slides as S

ORIG = "/root/.claude/uploads/df3fafbd-cd5a-5689-8909-9c95fa4cbf11/c731572d-RTE_Rice_Market_Research..pptx"
OUT = "RTE_Rice_Market_Research_FINAL.pptx"

prs = Presentation(ORIG)
n_orig = len(prs.slides._sldIdLst)
assert n_orig == 19

# ── generate the new slides (they are appended, then reordered) ────
BUILD = ["sl_02", "sl_04", "sl_05", "sl_06", "sl_15", "sl_16", "sl_07",
         "sl_17", "sl_18", "sl_03", "sl_09", "sl_10", "sl_11", "sl_12",
         "sl_13", "sl_14"]
made = {fn: getattr(S, fn)(prs) for fn in BUILD}
DIV = None

ids = list(prs.slides._sldIdLst)
orig = ids[:n_orig]
new = {fn: ids[n_orig + i] for i, fn in enumerate(BUILD)}

# ── narrative order: theirs interleaved with mine, nothing of theirs edited ──
order = [
    ("o", 0),                      # титул
    ("o", 1),                      # головне
    ("o", 2),                      # обсяг ринку
    ("o", 3),                      # карта ринку
    ("n", "sl_04"), ("n", "sl_05"),   # матриця брендів 1–2
    ("o", 4),                      # формати упаковки
    ("n", "sl_06"),                # аналітика форматів
    ("o", 5),                      # грамаж
    ("n", "sl_15"), ("n", "sl_16"),   # формат × грамаж, найпопулярніші
    ("o", 6),                      # цінові сходи
    ("o", 7), ("o", 8), ("o", 9), ("o", 10),      # картки брендів
    ("o", 11), ("o", 12), ("o", 13), ("o", 14), ("o", 15),
    ("o", 16),                     # суміжна полиця
    ("n", "sl_17"), ("n", "sl_18"),   # ритейл-аудит, азійські бренди
    ("n", "sl_03"),                # канали
    ("o", 17),                     # де представлено
    ("o", 18),                     # висновки дослідження
    ("n", "sl_02"),                # шість висновків
    ("n", "sl_07"),                # ранжир за упаковку
    ("n", "sl_09"),                # фінмодель проти ринку
    ("n", "sl_10"),                # аудит фінмоделі
    ("n", "sl_11"), ("n", "sl_12"), ("n", "sl_13"), ("n", "sl_14"),
]

seq = [orig[r] if k == "o" else new[r] for k, r in order]
assert len(seq) == len(set(id(e) for e in seq)) == n_orig + len(BUILD)

lst = prs.slides._sldIdLst
for el in list(lst):
    lst.remove(el)
for el in seq:
    lst.append(el)

# ── page numbers on the generated slides only ──────────────────────
mine = {id(made[fn]._element) for fn in BUILD}
for pos, sl in enumerate(prs.slides, 1):
    if id(sl._element) not in mine:
        continue
    for sh in sl.shapes:
        if (sh.has_text_frame
                and re.fullmatch(r"\d\d", sh.text_frame.text.strip())
                and Emu(sh.left).inches > 11.5 and Emu(sh.top).inches < 0.9):
            p = sh.text_frame.paragraphs[0]
            if p.runs:
                p.runs[0].text = f"{pos:02}"
                for r in p.runs[1:]:
                    r.text = ""

prs.save(OUT)
print(f"{OUT}: {len(prs.slides._sldIdLst)} слайдів, "
      f"{n_orig} оригінальних не редаговано")
