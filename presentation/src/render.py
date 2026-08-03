"""Rasterise the deck to PNGs so the layout can be eyeballed before delivery."""
import os
import shutil
import sys

sys.path.insert(0, '/root/.claude/skills/pptx/scripts')
from office.soffice import run_soffice  # noqa: E402

SRC = sys.argv[1] if len(sys.argv) > 1 else '/home/user/work/deck_out.pptx'
OUT = '/home/user/work/render'
DPI = int(sys.argv[2]) if len(sys.argv) > 2 else 110


def main():
    shutil.rmtree(OUT, ignore_errors=True)
    os.makedirs(OUT)
    r = run_soffice(['--headless', '--convert-to', 'pdf', '--outdir', OUT, SRC],
                    capture_output=True, text=True, timeout=600)
    pdf = os.path.join(OUT, os.path.splitext(os.path.basename(SRC))[0] + '.pdf')
    if not os.path.exists(pdf):
        print('CONVERT FAILED\n', r.stdout, r.stderr)
        sys.exit(1)

    import fitz
    doc = fitz.open(pdf)
    for i, page in enumerate(doc, 1):
        page.get_pixmap(dpi=DPI).save(os.path.join(OUT, 'slide%02d.png' % i))
    print('rendered %d slides at %d dpi -> %s' % (doc.page_count, DPI, OUT))


if __name__ == '__main__':
    main()
