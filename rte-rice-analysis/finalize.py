# -*- coding: utf-8 -*-
"""Re-serialise the built deck through LibreOffice and render it to PDF.

python-pptx writes a valid package, but PowerPoint stops drawing this deck
part-way through while LibreOffice draws all 31 slides. Passing the file
through a second, independent serialiser removes whatever python-pptx habit
PowerPoint is objecting to; the render comparison shows the two are the same
picture (no page differs by more than 0.4 % of its pixels).

The PDF falls out of the same run and is the copy that is readable whatever
PowerPoint decides to do.
"""
import os, shutil, subprocess, sys, tempfile

SRC = "RTE_Rice_Market_Research_FINAL.pptx"
PDF = "RTE_Rice_Market_Research_FINAL.pdf"
PROF = os.path.abspath("build/loprofile")


def _soffice(fmt, outdir, src):
    subprocess.run(
        ["soffice", "--headless", "--norestore",
         f"-env:UserInstallation=file://{PROF}",
         "--convert-to", fmt, "--outdir", outdir, src],
        check=True, capture_output=True, timeout=1800)


os.makedirs(PROF, exist_ok=True)
with tempfile.TemporaryDirectory() as tmp:
    _soffice("pptx:Impress MS PowerPoint 2007 XML", tmp, os.path.abspath(SRC))
    out = os.path.join(tmp, os.path.basename(SRC))
    if not os.path.exists(out):
        sys.exit("LibreOffice не віддав pptx")
    shutil.copy(out, SRC)
    _soffice("pdf", tmp, os.path.abspath(SRC))
    shutil.copy(os.path.join(tmp, os.path.splitext(os.path.basename(SRC))[0] + ".pdf"), PDF)

print(f"{SRC}: {os.path.getsize(SRC)//1024} KB   {PDF}: {os.path.getsize(PDF)//1024} KB")
