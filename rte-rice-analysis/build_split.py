# -*- coding: utf-8 -*-
"""Build two decks that isolate where the corruption comes from.

A — the user's original 19 slides, byte-for-byte untouched, with the new
    slides appended after them. Nothing of theirs is edited.
B — only the new slides, in a presentation of their own.

If A opens, the fault is in the edits made to the original slides.
If A fails but B opens, the fault is in combining them.
If B fails, the fault is in the slide-building code.
"""
from pptx import Presentation
import data as d
import deckkit as K
import slides as S

ORIG = "/root/.claude/uploads/df3fafbd-cd5a-5689-8909-9c95fa4cbf11/c731572d-RTE_Rice_Market_Research..pptx"

NEW = ["sl_02", "sl_04", "sl_05", "sl_06", "sl_15", "sl_16", "sl_07",
       "sl_17", "sl_18", "sl_03", "sl_09", "sl_10", "sl_11", "sl_12",
       "sl_13", "sl_14"]


def add_new_slides(prs, start_no):
    """Append the generated slides, numbering them from start_no."""
    made = []
    for i, fn in enumerate(NEW):
        sl = getattr(S, fn)(prs)
        made.append(sl)
    # renumber the page badge on each
    import re
    from pptx.util import Emu
    for i, sl in enumerate(made, start_no):
        for sh in sl.shapes:
            if (sh.has_text_frame
                    and re.fullmatch(r"\d\d", sh.text_frame.text.strip())
                    and Emu(sh.left).inches > 11.5 and Emu(sh.top).inches < 0.9):
                p = sh.text_frame.paragraphs[0]
                if p.runs:
                    p.runs[0].text = f"{i:02}"
                    for r in p.runs[1:]:
                        r.text = ""
    return made


# ── A: original untouched + new slides appended ────────────────────
a = Presentation(ORIG)
add_new_slides(a, len(a.slides._sldIdLst) + 1)
a.save("A_original_plus_new.pptx")
print("A:", len(a.slides._sldIdLst), "слайдів (19 оригінальних недоторкані)")

# ── B: only the new slides ─────────────────────────────────────────
b = K.new_deck()
add_new_slides(b, 1)
b.save("B_only_new.pptx")
print("B:", len(b.slides._sldIdLst), "слайдів (лише нові)")
