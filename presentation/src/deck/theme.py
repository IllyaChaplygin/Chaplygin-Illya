"""Design system lifted from the Avocado deck: same palette, type and card style."""
import copy

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

# ------------------------------------- palette (from Avocado_Product_Final.pptx) ----
DARK = RGBColor(0x1A, 0x3A, 0x2A)      # header bars, covers, headings
GREEN = RGBColor(0x2E, 0x6E, 0x4A)     # pills, badges, highlighted row
BRIGHT = RGBColor(0x6D, 0xC0, 0x4B)    # accent
MINT = RGBColor(0xE8, 0xF5, 0xE0)      # price panel fill
PAGE = RGBColor(0xF7, 0xFA, 0xF4)      # slide background
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
TEXT = RGBColor(0x1A, 0x2E, 0x1C)      # body copy
SOFT = RGBColor(0x3D, 0x5A, 0x40)      # secondary copy

HEAD = 'Cambria'
BODY = 'Calibri'

SW, SH = 10.0, 5.625
M = 0.4
BAR_H = 0.72


def fmt_usd(v):
    return ('$%.3f' % v).replace('.', ',')


def fmt_uah(v):
    return ('%.2f ₴' % v).replace('.', ',')


# ------------------------------------------------------------------ primitives ----
_SHADOW = (
    '<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
    '<a:outerShdw blurRad="76200" dist="25400" dir="2700000" algn="bl" rotWithShape="0">'
    '<a:srgbClr val="000000"><a:alpha val="12000"/></a:srgbClr></a:outerShdw></a:effectLst>'
)
_SHADOW_XML = None


def _add_shadow(shape):
    """Avocado cards carry a soft 12% drop shadow; python-pptx can only clear one."""
    global _SHADOW_XML
    if _SHADOW_XML is None:
        from pptx.oxml import parse_xml
        _SHADOW_XML = parse_xml(_SHADOW)
    spPr = shape._element.spPr
    for old in spPr.findall(qn('a:effectLst')):
        spPr.remove(old)
    spPr.append(copy.deepcopy(_SHADOW_XML))


def rect(slide, x, y, w, h, fill=None, radius=None, line=None, lw=0.75, shadow=False):
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
    if shadow:
        _add_shadow(sh)
    return sh


def text(slide, x, y, w, h, runs, size=10, font=BODY, color=TEXT, bold=False,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, line_spacing=1.0, italic=False,
         wrap=True):
    """runs: a string, or a list of (text, {overrides}) tuples rendered as paragraphs."""
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    items = runs if isinstance(runs, list) else [(runs, {})]
    # a "\n" inside a run becomes a soft break that ignores paragraph spacing —
    # split it into real paragraphs so line_spacing applies to every line
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
    return box


def picture(slide, path, bx, by, bw, bh):
    """Scale to fit (contain) inside the box, centred."""
    from PIL import Image
    with Image.open(path) as im:
        iw, ih = im.size
    scale = min(bw / iw, bh / ih)
    w, h = iw * scale, ih * scale
    return slide.shapes.add_picture(path, Inches(bx + (bw - w) / 2),
                                    Inches(by + (bh - h) / 2), Inches(w), Inches(h))


# ------------------------------------------------------------------- furniture ----
def blank(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    rect(s, 0, 0, SW, SH, fill=PAGE)
    return s


def header(slide, title, tag=None):
    rect(slide, 0, 0, SW, BAR_H, fill=DARK)
    text(slide, M, 0, SW - 2 * M - 2.4, BAR_H, title, size=18, font=HEAD, color=WHITE,
         bold=True, anchor=MSO_ANCHOR.MIDDLE)
    if tag:
        text(slide, SW - M - 2.4, 0, 2.4, BAR_H, tag, size=9.5, color=BRIGHT,
             bold=True, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)


def pill(slide, x, y, w, h, label, fill=GREEN, color=WHITE, size=9, radius=0.22):
    rect(slide, x, y, w, h, fill=fill, radius=radius)
    text(slide, x, y, w, h, label, size=size, color=color, bold=True,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, wrap=False)
