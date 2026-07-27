"""Design system for the suppliers deck: palette, type scale and shape primitives."""
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

# ---------------------------------------------------------------- palette ----
INK = RGBColor(0x14, 0x23, 0x2B)        # deep slate-teal — header bars, covers
INK_SOFT = RGBColor(0x1E, 0x33, 0x3C)
PRIMARY = RGBColor(0x0E, 0x5C, 0x51)    # deep nori green
PRIMARY_L = RGBColor(0xE4, 0xEF, 0xEC)
ACCENT = RGBColor(0xC1, 0x51, 0x2A)     # terracotta — prices, emphasis
ACCENT_L = RGBColor(0xF8, 0xEC, 0xE6)
GOLD = RGBColor(0xD9, 0xA4, 0x41)
CREAM = RGBColor(0xF7, 0xF4, 0xEE)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
TEXT = RGBColor(0x1C, 0x2A, 0x2F)
MUTED = RGBColor(0x64, 0x75, 0x7B)
LINE = RGBColor(0xE2, 0xE6, 0xE3)

HEAD = 'Cambria'
BODY = 'Calibri'

# ------------------------------------------------------------------ grid ----
SW, SH = 10.0, 5.625          # slide size, inches
M = 0.4                       # outer margin
BAR_H = 0.72                  # header bar height
BODY_TOP = 1.00               # first usable y below the header


def fmt_usd(v):
    # three decimals throughout: unit costs run from $0,169 to $2,72 and a mixed
    # precision made columns of figures impossible to scan
    return ('$%.3f' % v).replace('.', ',')


def fmt_uah(v):
    return ('%.2f ₴' % v).replace('.', ',')


def fmt_num(v, d=2):
    return ('%.*f' % (d, v)).replace('.', ',')


# ------------------------------------------------------------ primitives ----
def _plain(shape):
    shape.line.fill.background()
    shape.shadow.inherit = False
    return shape


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


def text(slide, x, y, w, h, runs, size=10, font=BODY, color=TEXT, bold=False,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, space=0, line_spacing=1.0,
         italic=False, wrap=True):
    """runs: a string, or a list of (text, {overrides}) tuples rendered as paragraphs."""
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    paras = runs if isinstance(runs, list) else [(runs, {})]
    for i, item in enumerate(paras):
        content, over = item if isinstance(item, tuple) else (item, {})
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = over.get('align', align)
        p.line_spacing = over.get('line_spacing', line_spacing)
        if i:
            p.space_before = Pt(over.get('space_before', space))
        r = p.add_run()
        r.text = content
        f = r.font
        f.name = over.get('font', font)
        f.size = Pt(over.get('size', size))
        f.bold = over.get('bold', bold)
        f.italic = over.get('italic', italic)
        f.color.rgb = over.get('color', color)
    return box


def picture(slide, path, bx, by, bw, bh):
    """Insert a picture scaled to fit (contain) inside the given box, centred."""
    from PIL import Image
    with Image.open(path) as im:
        iw, ih = im.size
    scale = min(bw / iw, bh / ih)
    w, h = iw * scale, ih * scale
    return slide.shapes.add_picture(path, Inches(bx + (bw - w) / 2),
                                    Inches(by + (bh - h) / 2), Inches(w), Inches(h))


# ------------------------------------------------------------- furniture ----
def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def page_bg(slide, fill=WHITE):
    rect(slide, 0, 0, SW, SH, fill=fill)


def header(slide, title, eyebrow=None, tag=None):
    """Dark title bar with an accent rule; optional right-hand tag."""
    rect(slide, 0, 0, SW, BAR_H, fill=INK)
    rect(slide, M, BAR_H - 0.045, 1.9, 0.045, fill=ACCENT)
    ty, th = (0.10, 0.24) if eyebrow else (0.18, 0.38)
    if eyebrow:
        text(slide, M, ty, SW - 2 * M - 2.2, 0.20, eyebrow.upper(), size=7.5,
             color=GOLD, bold=True)
        text(slide, M, 0.28, SW - 2 * M - 2.2, 0.36, title, size=18, font=HEAD,
             color=WHITE, bold=True, anchor=MSO_ANCHOR.TOP)
    else:
        text(slide, M, ty, SW - 2 * M - 2.2, th, title, size=18, font=HEAD,
             color=WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    if tag:
        text(slide, SW - M - 2.6, 0.24, 2.6, 0.26, tag, size=8.5, color=WHITE,
             bold=True, align=PP_ALIGN.RIGHT)


def footer(slide, page, note=None):
    rect(slide, M, SH - 0.30, SW - 2 * M, 0.01, fill=LINE)
    if note:
        text(slide, M, SH - 0.25, SW - 2 * M - 0.6, 0.20, note, size=6.5, color=MUTED)
    text(slide, SW - M - 0.6, SH - 0.25, 0.6, 0.20, str(page), size=7.5,
         color=MUTED, bold=True, align=PP_ALIGN.RIGHT)


def chip(slide, x, y, label, fill=ACCENT_L, color=ACCENT, w=None, size=6.5, h=0.22):
    w = w or (0.075 * len(label) + 0.22)
    rect(slide, x, y, w, h, fill=fill, radius=0.5)
    text(slide, x, y, w, h, label.upper(), size=size, color=color, bold=True,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, wrap=False)
    return w


def stat(slide, x, y, w, h, value, label, accent=PRIMARY):
    rect(slide, x, y, w, h, fill=CREAM, radius=0.10, line=LINE)
    rect(slide, x, y, 0.045, h, fill=accent)
    text(slide, x + 0.18, y + 0.13, w - 0.30, 0.36, value, size=17, font=HEAD,
         bold=True, color=accent)
    text(slide, x + 0.18, y + 0.52, w - 0.30, h - 0.60, label, size=7.5, color=MUTED,
         line_spacing=1.05)


def kv_table(slide, x, y, w, rows, rh=0.235, key_w=0.42, size=7.5):
    """Zebra key/value list used on profile slides."""
    for i, (k, v) in enumerate(rows):
        yy = y + i * rh
        if i % 2 == 0:
            rect(slide, x, yy, w, rh, fill=CREAM)
        text(slide, x + 0.10, yy, w * key_w - 0.10, rh, k, size=size, color=MUTED,
             anchor=MSO_ANCHOR.MIDDLE)
        text(slide, x + w * key_w, yy, w * (1 - key_w) - 0.10, rh, v, size=size,
             color=TEXT, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    return y + len(rows) * rh
