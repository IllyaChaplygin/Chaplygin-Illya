# -*- coding: utf-8 -*-
"""Diff two .pptx by the XML constructs they use.

Compares the set of (element, attribute) pairs and attribute-value shapes that
appear in each file's slides. Anything present only in the suspect file is a
construct the known-good file never needed — a candidate for what breaks it.
"""
import sys, zipfile, re, collections
from lxml import etree


def constructs(path):
    z = zipfile.ZipFile(path)
    elems = collections.Counter()
    attrs = collections.Counter()
    numeric = collections.defaultdict(list)
    for n in sorted(x for x in z.namelist()
                    if re.match(r"ppt/slides/slide\d+\.xml$", x)):
        root = etree.fromstring(z.read(n))
        for el in root.iter():
            if isinstance(el, etree._Comment):
                continue
            tag = etree.QName(el).localname
            elems[tag] += 1
            for k, v in el.attrib.items():
                key = etree.QName(k).localname if "}" in k else k
                attrs[f"{tag}@{key}"] += 1
                if v.lstrip("-").isdigit():
                    numeric[f"{tag}@{key}"].append(int(v))
    return elems, attrs, numeric


good_e, good_a, good_n = constructs(sys.argv[1])
bad_e, bad_a, bad_n = constructs(sys.argv[2])

print(f"A (good)   {sys.argv[1]}")
print(f"B (suspect){sys.argv[2]}\n")

only_b_e = sorted(set(bad_e) - set(good_e))
only_b_a = sorted(set(bad_a) - set(good_a))
print("елементи лише в B:", only_b_e or "—")
print("атрибути лише в B:", only_b_a or "—")

print("\nчислові атрибути з підозрілим діапазоном у B:")
flagged = False
for k, vals in sorted(bad_n.items()):
    lo, hi = min(vals), max(vals)
    glo, ghi = (min(good_n[k]), max(good_n[k])) if good_n.get(k) else (None, None)
    bad_range = (
        (k.endswith("@sz") and (lo < 100 or hi > 40000))
        or (k.endswith("@id") and lo <= 0)
        or (k.endswith("@val") and abs(hi) > 100000000)
        or lo < -914400 * 60 or hi > 914400 * 60
    )
    if bad_range:
        flagged = True
        print(f"  {k}: B={lo}..{hi}  A={glo}..{ghi}")
if not flagged:
    print("  — нічого")

print("\nнайбільші розбіжності в частоті елементів (B/A):")
for k in sorted(set(bad_e) | set(good_e)):
    a, b = good_e.get(k, 0), bad_e.get(k, 0)
    if a and b / a > 3 or (b and not a):
        print(f"  {k}: A={a} B={b}")
