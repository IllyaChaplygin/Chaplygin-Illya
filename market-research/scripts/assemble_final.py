tail = '''

# ══ ЗБІРКА (фінальна версія) ══════════════════════════════════════════════
slide_cover()
slide_summary()
slide_volume()
slide_map()
slide_bens()
slide_bowls()
slide_ua()
slide_import()
slide_format()
slide_ladder()
slide_retail_audit()
slide_channels()
slide_suppliers()
slide_finmodel()
slide_recommendation()
slide_conclusions()

prs.save("Gotovyi_Rys_Final.pptx")
print("saved ·", len(prs.slides._sldIdLst), "slides ·", N_SKU, "SKU ·", N_BRANDS, "brands")
'''
import io
parts = []
for f in ["part_head.py","part_data.py"]:
    parts.append(open(f).read())
parts.append(open("part_prim.py").read())
parts.append("import collections, math\n")
parts.append(open("part_calc.py").read())
# part_analytics up to (not incl) old slide_summary — we need primitives+SEGMENTS+slide_cover+slide_map,
# but skip its slide_summary/slide_volume defs since part_final overrides them anyway (redefinition is fine)
parts.append(open("part_analytics.py").read())
parts.append(open("part_brands.py").read())
# part_build: keep slide_format + slide_ladder + slide_chain_shelf, skip the rest (old channels/finmodel/
# recommendation/conclusions + assembly) by truncating before "# ══ КАНАЛИ"
pb = open("part_build.py").read()
cut = pb.index("# ══ КАНАЛИ")
parts.append(pb[:cut])
parts.append(open("part_final.py").read())
parts.append(tail)
open("build_finalv2.py","w").write("\n".join(parts))
print("written")
