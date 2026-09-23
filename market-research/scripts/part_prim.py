# ── примітиви ─────────────────────────────────────────────────────────────
def rect(s, x, y, w, h, fill, line=None, lw=1.0, rounded=False, adj=0.10):
    sh = s.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE,
        Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid(); sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line; sh.line.width = Pt(lw)
    sh.shadow.inherit = False
    if rounded:
        sh.adjustments[0] = adj
    return sh


def dot(s, cx, cy, d, fill):
    return rect(s, cx - d / 2, cy - d / 2, d, d, fill, rounded=True, adj=0.5)


def text(s, x, y, w, h, body, size=11, bold=False, italic=False, color=INK,
         font=FONT, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, line=None):
    box = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    items = body if isinstance(body, list) else [body]
    for i, item in enumerate(items):
        txt, over = item if isinstance(item, tuple) else (item, {})
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = over.get("align", align)
        if line:
            p.line_spacing = line
        for chunk in (txt if isinstance(txt, list) else [txt]):
            ctxt, cover = chunk if isinstance(chunk, tuple) else (chunk, {})
            r = p.add_run(); r.text = ctxt
            f = r.font
            f.name = cover.get("font", over.get("font", font))
            f.size = Pt(cover.get("size", over.get("size", size)))
            f.bold = cover.get("bold", over.get("bold", bold))
            f.italic = cover.get("italic", over.get("italic", italic))
            f.color.rgb = cover.get("color", over.get("color", color))
    return box


def pic(s, path, x, y, w, h):
    iw, ih = Image.open(path).size
    sc = min(w / iw, h / ih)
    pw, ph = iw * sc, ih * sc
    return s.shapes.add_picture(path, Inches(x + (w - pw) / 2), Inches(y + (h - ph) / 2),
                                Inches(pw), Inches(ph))


def cover_pic(s, path, x, y, w, h):
    iw, ih = Image.open(path).size
    sc = max(w / iw, h / ih)
    pw, ph = iw * sc, ih * sc
    p = s.shapes.add_picture(path, Inches(x), Inches(y), Inches(pw), Inches(ph))
    p.crop_left = p.crop_right = max(0.0, (pw - w) / pw / 2)
    p.crop_top = p.crop_bottom = max(0.0, (ph - h) / ph / 2)
    p.left, p.top, p.width, p.height = Inches(x), Inches(y), Inches(w), Inches(h)
    return p


def slide(dark=False):
    s = prs.slides.add_slide(BLANK)
    rect(s, 0, 0, W, H, NAVY if dark else WHITE)
    PG["n"] += 1
    return s


def header(s, eyebrow, title, dek=None):
    rect(s, 0, 0, W, HDR, NAVY)
    rect(s, 0, HDR, W, 0.055, ORANGE)
    s.shapes.add_picture(LOGO, Inches(M), Inches(0.26), Inches(1.52), Inches(0.66))
    text(s, 2.52, 0.34, 7.40, 0.24, eyebrow, size=9.5, bold=True, color=AMBER)
    text(s, 2.52, 0.60, 8.80, 0.40, title, size=19, bold=True, color=WHITE)
    text(s, W - M - 0.70, 0.44, 0.70, 0.36, f"{PG['n']:02d}", size=17, bold=True,
         color=NAVY_L, align=PP_ALIGN.RIGHT)
    if dek:
        text(s, M, HDR + 0.26, 11.9, 0.34, dek, size=12, color=GREY, line=1.20)


def foot(s, src):
    text(s, M, 7.08, 11.9, 0.26, src, size=8, color=GREY, line=1.14)


def pill(s, x, y, label, bg=GREEN, w=None, size=8.5, fg=WHITE):
    w = w or (0.24 + 0.072 * len(label))
    rect(s, x, y, w, 0.26, bg, rounded=True, adj=0.5)
    text(s, x, y, w, 0.26, label, size=size, bold=True, color=fg,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    return w


def num(v, d=0):
    """1190 → «1 190», 89.99 → «90»; десяткова кома."""
    return f"{v:,.{d}f}".replace(",", "\u00a0").replace(".", ",")


def pl(n, one, few, many):
    """Українська множина: 1 позиція, 3 позиції, 5 позицій."""
    n = int(n)
    if n % 10 == 1 and n % 100 != 11:
        return f"{n} {one}"
    if 2 <= n % 10 <= 4 and not 12 <= n % 100 <= 14:
        return f"{n} {few}"
    return f"{n} {many}"


def weights_of(items, idx=3):
    """Грамажі формату: до трьох — переліком «220, 240 і 250 г», більше — «217–320 г»."""
    ws = sorted({x[idx] for x in items})
    if len(ws) == 1:
        return f"{num(ws[0])} г"
    if len(ws) <= 3:
        return ", ".join(num(w) for w in ws[:-1]) + f" і {num(ws[-1])} г"
    return f"{num(ws[0])}–{num(ws[-1])} г"


def fmt_title(items, fmt):
    """«13 позицій: пауч 220, 240 і 250 г»."""
    return (f"{pl(len(items), 'позиція', 'позиції', 'позицій')}: "
            f"{fmt} {weights_of(items)}")


def rng(lo, hi):
    a, b = num(lo), num(hi)
    return a if a == b else f"{a}–{b}"


def money(lo, hi):
    return num(lo) if lo == hi else f"{num(lo)}–{num(hi)}"


def sku_cell(s, x, y, w, h, img, name, grams, lo, hi, chan, c, brand=None):
    """Клітинка SKU: пакшот, назва, маса, ціна, ціна за 100 г, канал."""
    rect(s, x, y, w, h, WHITE, line=MIST_D, lw=1.0, rounded=True, adj=0.07)
    ph = h - 1.30                      # текстовий блок завжди 1,30″
    rect(s, x + 0.08, y + 0.08, w - 0.16, ph, MIST, rounded=True, adj=0.08)
    if img:
        pic(s, img, x + 0.14, y + 0.12, w - 0.28, ph - 0.08)
    else:
        text(s, x + 0.08, y + ph / 2 - 0.04, w - 0.16, 0.24, "фото немає", size=8,
             color=GREY, align=PP_ALIGN.CENTER)
    yy = y + ph + 0.12
    text(s, x + 0.13, yy, w - 0.26, 0.36, name, size=9, bold=True, color=NAVY, line=1.02)
    gt = f"{num(grams, 0 if float(grams).is_integer() else 1)} г" if grams else "маса н/д"
    if brand:
        text(s, x + 0.13, yy + 0.38, w - 0.26, 0.18,
             [([(brand.upper() + "  ·  ", {"bold": True, "color": c}), gt], {})],
             size=7.5, color=GREY)
    else:
        text(s, x + 0.13, yy + 0.38, w - 0.26, 0.18, gt, size=7.5, color=GREY)
    text(s, x + 0.13, yy + 0.56, w - 0.26, 0.26, money(lo, hi) + " грн", size=11.5,
         bold=True, color=c)
    text(s, x + 0.13, yy + 0.84, w - 0.26, 0.18, chan[:34], size=7.5, color=GREY)


def brand_head(s, x, y, brand, origin, fmt, n, mlo, mhi, plo, phi, c):
    rect(s, x, y, 11.98, 0.66, MIST, rounded=True, adj=0.20)
    rect(s, x, y, 0.09, 0.66, c, rounded=True, adj=0.5)
    text(s, x + 0.30, y + 0.07, 3.00, 0.28, brand, size=14, bold=True, color=NAVY)
    text(s, x + 0.30, y + 0.36, 3.00, 0.22, origin, size=8.5, color=GREY)
    for i, (lb, vl) in enumerate([("ФОРМАТ", fmt), ("ПОЗИЦІЙ", f"{n}"),
                                  ("ЗА УПАКОВКУ", f"{money(plo, phi)} грн"),
                                  ("МАСА", f"{rng(mlo, mhi)} г")]):
        cx = x + 3.60 + i * 2.10
        text(s, cx, y + 0.09, 2.00, 0.18, lb, size=7.5, bold=True, color=GREY)
        text(s, cx, y + 0.29, 2.00, 0.28, vl, size=11.5, bold=True,
             color=c if i >= 2 else INK)


