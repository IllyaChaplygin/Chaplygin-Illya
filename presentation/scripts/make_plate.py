"""PANTRY LIGHT — Plate I: a specimen chart of unit cost.

The design philosophy expressed as an artefact rather than a slide: fifteen
seasonings drawn as specimens in a naturalist's plate, each circle a
constant-scale observation, each annotated with the one property the whole
supplier study exists to measure — cost per gram. Specimens still awaiting
photography are drawn as empty rings, the way a plate leaves a slot open for
a species collected but not yet illustrated.
"""
import base64
import os

import cairosvg

W, H = 2480, 3508                       # A4 at 300 dpi
LINEN, BLUSH = "#FAF7F2", "#F1DDD3"
CHERRY, INK, GREY, RULE = "#C32733", "#231C18", "#6A5F56", "#DFD4C9"
SERIF, SANS = "Liberation Serif", "Liberation Sans"
C = "assets/circles"

# (code, latin label, uah-per-gram, circle png or None)
SPECIMENS = [
    ("I.i",    "ZEK · CHICKEN FLOSS",  0.90, f"{C}/zek_chicken35.png"),
    ("I.ii",   "ZEK · VEGETABLES",     0.90, f"{C}/zek_veg35.png"),
    ("I.iii",  "ZEK · SESAME",         0.90, f"{C}/zek_sesame35.png"),
    ("I.iv",   "ZEK · VEGETABLES 70",  0.91, f"{C}/zek_veg70.png"),
    ("I.v",    "ZEK · SESAME 70",      0.91, f"{C}/zek_sesame70.png"),
    ("II.i",   "MISHIMA · NORI",       1.12, f"{C}/ms_nori.png"),
    ("II.ii",  "MISHIMA · KIMCHI",     1.33, None),
    ("II.iii", "MISHIMA · SHRIMP",     1.46, f"{C}/ms_shrimp.png"),
    ("II.iv",  "MISHIMA · WASABI",     1.80, f"{C}/ms_wasabi.png"),
    ("III.i",  "TAKAOKAYA · CURRY",    1.59, None),
    ("III.ii", "TAKAOKAYA · TAMAGO",   1.59, None),
    ("III.iii","TAKAOKAYA · NORI",     1.72, None),
    ("III.iv", "TAKAOKAYA · WASABI",   1.72, f"{C}/tk_wasabi.png"),
    ("III.v",  "TAKAOKAYA · YUZU",     1.81, None),
    ("III.vi", "TAKAOKAYA · KATSUO",   2.12, None),
]


def b64(path):
    with open(path, "rb") as fh:
        return base64.b64encode(fh.read()).decode()


def uah(v):
    return f"{v:.2f}".replace(".", ",")


parts = [f'<rect width="{W}" height="{H}" fill="{LINEN}"/>']

# ── masthead ──────────────────────────────────────────────────────────────
M = 240
parts.append(f'<rect x="{M}" y="300" width="150" height="7" fill="{CHERRY}"/>')
parts.append(f'<text x="{M + 190}" y="316" font-family="{SANS}" font-size="30" '
             f'font-weight="bold" letter-spacing="7" fill="{CHERRY}">PLATE I · UNIT COST</text>')
parts.append(f'<text x="{M}" y="470" font-family="{SERIF}" font-size="132" '
             f'font-weight="bold" fill="{INK}">Pantry Light</text>')
parts.append(f'<text x="{M}" y="556" font-family="{SERIF}" font-size="46" font-style="italic" '
             f'fill="{CHERRY}">Fifteen seasonings, observed at constant scale</text>')
parts.append(f'<line x1="{M}" y1="640" x2="{W - M}" y2="640" stroke="{RULE}" stroke-width="3"/>')
parts.append(f'<text x="{M}" y="700" font-family="{SANS}" font-size="26" letter-spacing="4" '
             f'fill="{GREY}">ГРИВНЯ ЗА ГРАМ · ЗБІРНИЙ ВАНТАЖ · 2026</text>')
parts.append(f'<text x="{W - M}" y="700" text-anchor="end" font-family="{SANS}" font-size="26" '
             f'letter-spacing="4" fill="{GREY}">n = 15</text>')

# ── specimen grid: 5 columns, constant circle scale ───────────────────────
COLS, D = 5, 300
cell_w = (W - 2 * M) / COLS
row_y = [1010, 1650, 2290]

for i, (code, label, val, img) in enumerate(SPECIMENS):
    c, r = i % COLS, i // COLS
    cx = M + cell_w * (c + 0.5)
    cy = row_y[r]

    parts.append(f'<circle cx="{cx:.1f}" cy="{cy}" r="{D / 2 + 26}" fill="{BLUSH}"/>')
    if img and os.path.exists(img):
        parts.append(f'<image x="{cx - D / 2:.1f}" y="{cy - D / 2}" width="{D}" height="{D}" '
                     f'xlink:href="data:image/png;base64,{b64(img)}"/>')
    else:
        # collected, not yet illustrated
        parts.append(f'<circle cx="{cx:.1f}" cy="{cy}" r="{D / 2}" fill="none" '
                     f'stroke="{CHERRY}" stroke-width="2.4" stroke-dasharray="9 11"/>')
        parts.append(f'<circle cx="{cx:.1f}" cy="{cy}" r="7" fill="{CHERRY}"/>')

    parts.append(f'<text x="{cx:.1f}" y="{cy + D / 2 + 96}" text-anchor="middle" '
                 f'font-family="{SANS}" font-size="21" letter-spacing="2.6" '
                 f'fill="{GREY}">{code}</text>')
    parts.append(f'<text x="{cx:.1f}" y="{cy + D / 2 + 142}" text-anchor="middle" '
                 f'font-family="{SANS}" font-size="18" letter-spacing="1.5" '
                 f'fill="{INK}">{label}</text>')
    parts.append(f'<text x="{cx:.1f}" y="{cy + D / 2 + 210}" text-anchor="middle" '
                 f'font-family="{SERIF}" font-size="52" font-weight="bold" '
                 f'fill="{CHERRY}">{uah(val)}</text>')

# ── the measurement: every specimen on one axis ───────────────────────────
AX_Y, LO, HI = 3170, 0.80, 2.25
ax0, ax1 = M, W - M


def axis_x(v):
    return ax0 + (ax1 - ax0) * (v - LO) / (HI - LO)


parts.append(f'<line x1="{M}" y1="2950" x2="{W - M}" y2="2950" stroke="{RULE}" stroke-width="3"/>')
parts.append(f'<text x="{M}" y="3022" font-family="{SANS}" font-size="24" letter-spacing="4" '
             f'fill="{CHERRY}" font-weight="bold">РОЗПОДІЛ ПИТОМОЇ СОБІВАРТОСТІ</text>')

parts.append(f'<line x1="{ax0}" y1="{AX_Y}" x2="{ax1}" y2="{AX_Y}" '
             f'stroke="{INK}" stroke-width="2"/>')
tick = LO
while tick <= HI + 1e-9:
    x = axis_x(tick)
    major = abs(tick * 100 - round(tick * 100 / 25) * 25) < 1e-6
    parts.append(f'<line x1="{x:.1f}" y1="{AX_Y}" x2="{x:.1f}" y2="{AX_Y + (26 if major else 14)}" '
                 f'stroke="{INK if major else RULE}" stroke-width="2"/>')
    if major:
        parts.append(f'<text x="{x:.1f}" y="{AX_Y + 72}" text-anchor="middle" '
                     f'font-family="{SANS}" font-size="24" fill="{GREY}">{uah(tick)}</text>')
    tick += 0.05

for code, label, val, img in SPECIMENS:
    x = axis_x(val)
    marker = (f'fill="{CHERRY}"' if img
              else f'fill="none" stroke="{CHERRY}" stroke-width="2.4"')
    parts.append(f'<line x1="{x:.1f}" y1="{AX_Y - 62}" x2="{x:.1f}" y2="{AX_Y - 8}" '
                 f'stroke="{CHERRY}" stroke-width="2"/>')
    parts.append(f'<circle cx="{x:.1f}" cy="{AX_Y - 74}" r="9" {marker}/>')

parts.append(f'<text x="{W - M}" y="3022" text-anchor="end" font-family="{SANS}" font-size="22" '
             f'letter-spacing="2" fill="{GREY}">● спостережено · ○ очікує фотофіксації</text>')

# ── colophon ──────────────────────────────────────────────────────────────
parts.append(f'<line x1="{M}" y1="3330" x2="{W - M}" y2="3330" stroke="{RULE}" stroke-width="3"/>')
parts.append(f'<text x="{M}" y="3392" font-family="{SERIF}" font-size="28" font-style="italic" '
             f'fill="{GREY}">Одна земля, одна форма, два голоси.</text>')
parts.append(f'<text x="{W - M}" y="3392" text-anchor="end" font-family="{SANS}" font-size="22" '
             f'letter-spacing="4" fill="{GREY}">PANTRY LIGHT · I</text>')

svg = (f'<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" '
       f'width="{W}" height="{H}" viewBox="0 0 {W} {H}">' + "".join(parts) + "</svg>")

cairosvg.svg2png(bytestring=svg.encode(), write_to="PANTRY_LIGHT_plate.png",
                 output_width=W, output_height=H)
print("PANTRY_LIGHT_plate.png", W, "x", H)
