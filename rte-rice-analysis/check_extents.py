# -*- coding: utf-8 -*-
"""Find shapes with a zero or negative extent — PowerPoint rejects those."""
import sys, zipfile, re
from lxml import etree

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
EMU = 914400

for path in sys.argv[1:]:
    z = zipfile.ZipFile(path)
    bad = []
    total = 0
    for n in sorted(x for x in z.namelist()
                    if re.match(r"ppt/slides/slide\d+\.xml$", x)):
        root = etree.fromstring(z.read(n))
        for ext in root.iter(f"{{{A}}}ext"):
            total += 1
            cx, cy = int(ext.get("cx", 0)), int(ext.get("cy", 0))
            if cx <= 0 or cy <= 0:
                owner = ext.getparent().getparent().getparent()
                name = ""
                for c in owner.iter(f"{{{P}}}cNvPr"):
                    name = c.get("name", ""); break
                txt = " ".join((t.text or "") for t in owner.iter(f"{{{A}}}t"))[:44]
                bad.append((n, cx / EMU, cy / EMU, name, txt))
        for off in root.iter(f"{{{A}}}off"):
            x, y = int(off.get("x", 0)), int(off.get("y", 0))
            if abs(x) > EMU * 60 or abs(y) > EMU * 60:
                bad.append((n, x / EMU, y / EMU, "OFFSET", ""))
    print(f"{path}: ext={total}  проблемних={len(bad)}")
    for b in bad[:14]:
        print(f"   {b[0]}  cx={b[1]:.2f}in cy={b[2]:.2f}in  {b[3]}  {b[4]!r}")
