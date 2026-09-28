# -*- coding: utf-8 -*-
"""Rebuild the deck's zip so everything that can come from the original does.

python-pptx re-serialises the whole package. That is normally fine, but it
puts every part under suspicion. Here only the parts that actually changed —
the slides, the presentation and the bits that index them — are taken from
the generated file. Media, masters, layouts, theme and props are copied
byte-for-byte out of the user's original, and entry order and compression
follow the original too.
"""
import sys, zipfile, shutil

ORIG = "/root/.claude/uploads/df3fafbd-cd5a-5689-8909-9c95fa4cbf11/c731572d-RTE_Rice_Market_Research..pptx"
SRC = sys.argv[1] if len(sys.argv) > 1 else "RTE_Rice_Market_Research_FULL.pptx"
OUT = sys.argv[2] if len(sys.argv) > 2 else "RTE_Rice_Market_Research_FINAL.pptx"

# parts that must come from the generated file
def is_generated(name):
    return (name.startswith("ppt/slides/")
            or name in ("ppt/presentation.xml",
                        "ppt/_rels/presentation.xml.rels",
                        "[Content_Types].xml")
            or name.startswith("docProps/"))

zo = zipfile.ZipFile(ORIG)
zs = zipfile.ZipFile(SRC)
orig_names = set(zo.namelist())
src_names = zs.namelist()

kept = reused = 0
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as out:
    # content types first, as the OPC spec wants
    out.writestr("[Content_Types].xml", zs.read("[Content_Types].xml"))
    for name in src_names:
        if name == "[Content_Types].xml":
            continue
        if not is_generated(name) and name in orig_names:
            info = zo.getinfo(name)
            data = zo.read(name)
            zi = zipfile.ZipInfo(name, date_time=info.date_time)
            zi.compress_type = info.compress_type
            zi.external_attr = info.external_attr
            out.writestr(zi, data)
            reused += 1
        else:
            out.writestr(name, zs.read(name))
            kept += 1

print(f"{OUT}: з оригіналу побайтово {reused} частин, згенеровано {kept}")

# sanity: same part set as the generated file
zc = zipfile.ZipFile(OUT)
assert set(zc.namelist()) == set(src_names), "набір частин розійшовся"
import hashlib
same = sum(1 for n in zc.namelist()
           if n in orig_names and not is_generated(n)
           and hashlib.md5(zc.read(n)).hexdigest() == hashlib.md5(zo.read(n)).hexdigest())
print(f"   ідентичних оригіналу: {same}")
