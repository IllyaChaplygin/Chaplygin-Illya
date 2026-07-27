"""Design system: light consulting-report layout — white space, hairlines, one accent.

The colour comes from the product photography; the page furniture stays quiet.
"""
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

# ------------------------------------------------------------------- palette ----
INK = RGBColor(0x0E, 0x2A, 0x24)        # headings
BODY_TX = RGBColor(0x38, 0x4A, 0x44)    # body copy
MUTED = RGBColor(0x7C, 0x8B, 0x86)      # labels, captions
ACCENT = RGBColor(0x0B, 0x7D, 0x57)     # emerald — fills carrying white text
ACCENT_TX = RGBColor(0x07, 0x6B, 0x4A)  # emerald — text on white
BRIGHT = RGBColor(0x00, 0xB8, 0x7C)     # thin rules and marks only
TINT = RGBColor(0xEA, 0xF6, 0xF1)       # pale mint panels
PANEL = RGBColor(0xF4, 0xF7, 0xF6)      # photo wells
RULE = RGBColor(0xE1, 0xE8, 0xE5)       # hairlines
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

HEAD = 'Cambria'
BODY = 'Calibri'

SW, SH = 10.0, 5.625
M = 0.40
TOP_RULE = 0.055
CONTENT_TOP = 1.22
CONTENT_BOTTOM = 5.16
HAIRLINE = 0.01


def fmt_usd(v):
    return ('$%.3f' % v).replace('.', ',')


def fmt_uah(v):
    return ('%.2f ₴' % v).replace('.', ',')


# ---------------------------------------------------------------- primitives ----
def rect(slide, x, y, w, h, fill=None, radius=None, line=None, lw=0.75):
    kind = MSO_SHAPE.ROUNDED_RECTANGLE if radius is not None else MSO_SHAPE.RECTANGLE
    sh = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    if radius is not None:
        sh.adjustments[0] = radius
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill
    sh.shadow.inherit = False
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(lw)
    sh.text_frame.text = ''
    return sh


def hline(slide, x, y, w, color=RULE):
    return rect(slide, x, y, w, HAIRLINE, fill=color)


def text(slide, x, y, w, h, runs, size=10, font=BODY, color=BODY_TX, bold=False,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, line_spacing=1.0, italic=False,
         wrap=True, spc=0):
    """runs: a string, or a list of (text, {overrides}). `spc` is letter spacing in pt."""
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    items = runs if isinstance(runs, list) else [(runs, {})]
    # a "\n" inside a run is a soft break that ignores paragraph spacing — split
    # it into real paragraphs so line_spacing applies to every line
    paras = []
    for item in items:
        content, over = item if isinstance(item, tuple) else (item, {})
        paras += [(line, over) for line in content.split('\n')]
    for i, (content, over) in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = over.get('align', align)
        p.line_spacing = over.get('line_spacing', line_spacing)
        r = p.add_run()
        r.text = content
        f = r.font
        f.name = over.get('font', font)
        f.size = Pt(over.get('size', size))
        f.bold = over.get('bold', bold)
        f.italic = over.get('italic', italic)
        f.color.rgb = over.get('color', color)
        tracking = over.get('spc', spc)
        if tracking:
            r.font._rPr.set('spc', str(int(tracking * 100)))
    return box


def label(slide, x, y, w, content, color=MUTED, size=6.5, align=PP_ALIGN.LEFT,
          h=0.16):
    """Small letter-spaced caps — the workhorse label of a consulting layout.

    Keep word_wrap on: a wrap="none" body auto-sizes and LibreOffice then centres
    it on the box regardless of paragraph alignment.
    """
    return text(slide, x, y, w, h, content.upper(), size=size, color=color, bold=True,
                align=align, spc=0.9)


def picture(slide, path, bx, by, bw, bh):
    """Scale to fit (contain) inside the box, centred."""
    from PIL import Image
    with Image.open(path) as im:
        iw, ih = im.size
    scale = min(bw / iw, bh / ih)
    w, h = iw * scale, ih * scale
    return slide.shapes.add_picture(path, Inches(bx + (bw - w) / 2),
                                    Inches(by + (bh - h) / 2), Inches(w), Inches(h))


# ----------------------------------------------------------------- furniture ----
def blank(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    rect(s, 0, 0, SW, SH, fill=WHITE)
    rect(s, 0, 0, SW, TOP_RULE, fill=ACCENT)
    return s


def header(slide, title, eyebrow=None, tag=None):
    if eyebrow:
        label(slide, M, 0.34, 5.6, eyebrow, color=ACCENT_TX, size=7.5)
    text(slide, M, 0.54, SW - 2 * M - 2.6, 0.46, title, size=20, font=HEAD, color=INK,
         bold=True, anchor=MSO_ANCHOR.MIDDLE)
    if tag:
        text(slide, SW - M - 2.6, 0.54, 2.6, 0.46, tag, size=9, color=MUTED,
             align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    hline(slide, M, 1.08, SW - 2 * M)


def footer(slide, page, note=None):
    hline(slide, M, 5.26, SW - 2 * M)
    if note:
        text(slide, M, 5.34, SW - 2 * M - 0.6, 0.16, note, size=6.5, color=MUTED,
             italic=True)
    text(slide, SW - M - 0.6, 5.34, 0.6, 0.16, str(page), size=7, color=MUTED,
         bold=True, align=PP_ALIGN.RIGHT)
