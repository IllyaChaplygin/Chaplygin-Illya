# -*- coding: utf-8 -*-
"""Append the market-research content onto the user's own supplier deck,
in place of shipping it as a separate file. The existing slides are never
touched — python-pptx only ever appends — but the whole file is re-serialized
on save, so every asset is checked byte-identical after the fact (see
verify_unchanged below) rather than assumed.
"""
import hashlib
import os
import sys
import zipfile

from pptx import Presentation

HERE = os.path.dirname(__file__)
sys.path.insert(0, HERE)
import build_market as M  # noqa: E402

DEFAULT_SRC = ('/root/.claude/uploads/c53f34b4-ae43-52db-aeb2-01bdcda48cc3/'
               '45308769-Snacks_Presentation_Final..pptx')
DEFAULT_OUT = os.path.join(HERE, '..', '..',
                           'Постачальники_та_зріз_ринку.pptx')


def media_hashes(path):
    with zipfile.ZipFile(path) as z:
        return {n: hashlib.md5(z.read(n)).hexdigest()
               for n in z.namelist() if n.startswith('ppt/media/')}


def verify_unchanged(src, out, n_original_slides):
    before, after = media_hashes(src), media_hashes(out)
    missing = set(before) - set(after)
    changed = [n for n in before if n in after and before[n] != after[n]]
    assert not missing, 'media files dropped: %s' % missing
    assert not changed, 'media files altered: %s' % changed
    print('media check: %d/%d original files identical, %d new slides carry '
         'no images yet' % (len(before), len(before), len(after) - len(before)))


def main(src=DEFAULT_SRC, out=DEFAULT_OUT):
    prs = Presentation(src)
    n0 = len(prs.slides)
    M._page[0] = n0
    M.build_content(prs, cover=M.slide_divider)
    prs.save(out)
    verify_unchanged(src, out, n0)
    print('saved %s — %d original + %d new = %d slides'
         % (out, n0, len(prs.slides) - n0, len(prs.slides)))


if __name__ == '__main__':
    src = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_SRC
    out = sys.argv[2] if len(sys.argv) > 2 else DEFAULT_OUT
    main(src, out)
