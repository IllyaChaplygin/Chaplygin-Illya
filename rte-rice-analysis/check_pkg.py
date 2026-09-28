# -*- coding: utf-8 -*-
"""Validate the OPC package of a .pptx: dangling rels, bad refs, missing parts."""
import sys, zipfile, posixpath, re, collections
from lxml import etree

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}
R = NS["r"]
P_ = NS["p"]
REL_ATTRS = [f"{{{R}}}embed", f"{{{R}}}link", f"{{{R}}}id", f"{{{R}}}pict",
             f"{{{R}}}dm", f"{{{R}}}lo", f"{{{R}}}qs", f"{{{R}}}cs"]

path = sys.argv[1]
z = zipfile.ZipFile(path)
names = set(z.namelist())
problems = []

def rels_for(part):
    d, f = posixpath.split(part)
    return posixpath.join(d, "_rels", f + ".rels")

def load_rels(part):
    rp = rels_for(part)
    if rp not in names:
        return {}
    root = etree.fromstring(z.read(rp))
    out = {}
    for rel in root:
        rid = rel.get("Id"); tgt = rel.get("Target"); mode = rel.get("TargetMode")
        if mode == "External":
            out[rid] = ("EXTERNAL", tgt)
        else:
            base = posixpath.dirname(part)
            out[rid] = ("INTERNAL", posixpath.normpath(posixpath.join(base, tgt)))
    return out

# 1. every part is well-formed XML
for n in sorted(names):
    if n.endswith((".xml", ".rels")):
        try:
            etree.fromstring(z.read(n))
        except Exception as e:
            problems.append(f"MALFORMED XML  {n}: {str(e)[:90]}")

# 2. content-type overrides point at parts that exist, and every xml part is typed
ct = etree.fromstring(z.read("[Content_Types].xml"))
overrides = {o.get("PartName").lstrip("/"): o.get("ContentType")
             for o in ct.findall("ct:Override", NS)}
defaults = {d.get("Extension").lower() for d in ct.findall("ct:Default", NS)}
for pn in overrides:
    if pn not in names:
        problems.append(f"CONTENT-TYPE points at missing part: /{pn}")
for n in names:
    if n == "[Content_Types].xml" or "/_rels/" in n:
        continue
    ext = n.rsplit(".", 1)[-1].lower()
    if n not in overrides and ext not in defaults:
        problems.append(f"UNTYPED PART (no Override, no Default for .{ext}): {n}")

# 3. every relationship target exists
for n in sorted(names):
    if not n.endswith(".rels"):
        continue
    owner = n.replace("/_rels/", "/").replace(".rels", "")
    if owner.startswith("_rels/"):
        owner = owner[len("_rels/"):]
    for rid, (kind, tgt) in load_rels(owner).items():
        if kind == "INTERNAL" and tgt not in names:
            problems.append(f"DANGLING REL  {owner}  {rid} -> {tgt}")

# 4. presentation.xml slide list resolves, and in order
pres = "ppt/presentation.xml"
proot = etree.fromstring(z.read(pres))
prels = load_rels(pres)
sld_ids = proot.find("p:sldIdLst", NS)
slide_parts = []
for sld in sld_ids:
    rid = sld.get(f"{{{R}}}id")
    if rid not in prels:
        problems.append(f"sldIdLst references missing rId {rid}")
        continue
    slide_parts.append(prels[rid][1])
dupe_ids = [s.get("id") for s in sld_ids]
if len(set(dupe_ids)) != len(dupe_ids):
    problems.append(f"DUPLICATE sldId id= values in sldIdLst")

# 5. every r:* reference inside each slide resolves to one of its own rels
for sp in slide_parts:
    if sp not in names:
        problems.append(f"MISSING SLIDE PART {sp}")
        continue
    srels = load_rels(sp)
    root = etree.fromstring(z.read(sp))
    for node in root.iter():
        for attr in REL_ATTRS:
            rid = node.get(attr)
            if rid and rid not in srels:
                tag = node.tag.split("}")[-1]
                problems.append(f"BAD REF  {sp}  <{tag} {attr.split('}')[-1]}=\"{rid}\"> "
                                f"not in its rels")


# 6. degenerate text structures — legal XML that PowerPoint still refuses
A = NS["a"]
empty_runs = 0
total_runs = 0
for sp in slide_parts:
    if sp not in names:
        continue
    root = etree.fromstring(z.read(sp))
    for r in root.iter(f"{{{A}}}r"):
        total_runs += 1
        t = r.find(f"{{{A}}}t")
        if t is None:
            problems.append(f"RUN WITHOUT TEXT  {sp}  <a:r> has no <a:t>")
        elif not (t.text or ""):
            empty_runs += 1
    for tb in root.iter(f"{{{P_}}}txBody"):
        if tb.find(f"{{{A}}}p") is None:
            problems.append(f"EMPTY TXBODY  {sp}  <p:txBody> has no <a:p>")
    ids = [e.get("id") for e in root.iter(f"{{{P_}}}cNvPr") if e.get("id")]
    dup = [k for k, v in collections.Counter(ids).items() if v > 1]
    if dup:
        problems.append(f"DUPLICATE SHAPE ID  {sp}  {','.join(dup[:4])}")
if total_runs and empty_runs > max(10, total_runs * 0.01):
    problems.append(f"TOO MANY EMPTY RUNS  {empty_runs} of {total_runs} — "
                    f"editing left blank <a:r> behind; remove them instead")

print(f"{path}")
print(f"  parts={len(names)}  slides={len(slide_parts)}  "
      f"runs={total_runs} (порожніх {empty_runs})")
if problems:
    print(f"  ПРОБЛЕМ: {len(problems)}")
    seen = {}
    for p in problems:
        key = p.split("  ")[0]
        seen.setdefault(key, []).append(p)
    for key, items in seen.items():
        print(f"\n  [{key}] x{len(items)}")
        for it in items[:6]:
            print("    ", it)
        if len(items) > 6:
            print(f"     ... ще {len(items)-6}")
else:
    print("  OK — структурних проблем не знайдено")
