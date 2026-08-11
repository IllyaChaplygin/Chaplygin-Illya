"""Editorial vector renderings of the Takaokaya fresh-lock furikake bottles.

The supplier provided no photography and outbound web access is blocked in this
environment, so the product visuals are drawn in a flat editorial style that
matches the deck (warm paper, terracotta accents) and are captioned on the
slides as visualisations rather than photographs.
"""
import os
import random

import cairosvg

os.makedirs("assets", exist_ok=True)

INK = "#2A221E"
PAPER = "#FBF7F3"
GLASS = "#EFEFEA"
GLASS_EDGE = "#D6D4CD"

# flavour -> (label line 1, label line 2, cap colour, seasoning base, flake tones)
FLAVOURS = {
    "nori":   ("NORI", "", "#1E5B45", "#2F3C31", ["#16281C", "#3E5B3A", "#8FA07E", "#E9E3D2"]),
    "wasabi": ("NORI", "WASABI", "#74A83E", "#3B4B2E", ["#22361F", "#5C7C3A", "#A8C07E", "#F0EBD8"]),
    "yuzu":   ("YUZU", "KOSHO", "#C9971F", "#4A4626", ["#33301A", "#7E7328", "#C4A63C", "#F2E9C9"]),
    "katsuo": ("NORI", "KATSUO", "#A8452C", "#40332A", ["#2A2019", "#6E4A33", "#A97C55", "#EADCC6"]),
    "tamago": ("NORI", "TAMAGO", "#E0AE4B", "#4C452C", ["#2E2A1B", "#7A6B2F", "#D6B košt"[:7], "#F5EBCB"]),
    "curry":  ("CURRY", "", "#C87A1E", "#4E3A22", ["#2E2116", "#7A4E20", "#BC8437", "#EFDFC0"]),
    "garlic": ("GARLIC", "CHILI OIL", "#B02418", "#4A2B24", ["#2C1A17", "#7C3226", "#B8654A", "#EFD9CC"]),
}
FLAVOURS["tamago"] = ("NORI", "TAMAGO", "#E0AE4B", "#4C452C",
                      ["#2E2A1B", "#7A6B2F", "#D6B24A", "#F5EBCB"])


def flakes(seed, x, y, w, h, tones, density=0.62):
    """Fine seasoning granules: mostly dark flakes, a few pale sesame seeds."""
    rnd = random.Random(seed)
    weights = [0.50, 0.28, 0.14, 0.08]      # dark, mid, light, sesame
    out = []
    for _ in range(int(w * h * density)):
        fx = rnd.uniform(x + 1, x + w - 1)
        fy = rnd.uniform(y + 1, y + h - 1)
        rx = rnd.uniform(w * 0.009, w * 0.021)
        ry = rx * rnd.uniform(0.45, 0.75)
        rot = rnd.uniform(-40, 40)
        tone = rnd.choices(tones, weights=weights[:len(tones)])[0]
        op = rnd.uniform(0.6, 0.95)
        out.append(f'<ellipse cx="{fx:.1f}" cy="{fy:.1f}" rx="{rx:.1f}" ry="{ry:.1f}" '
                   f'fill="{tone}" opacity="{op:.2f}" transform="rotate({rot:.0f} {fx:.1f} {fy:.1f})"/>')
    return "\n".join(out)


def bottle(key, cx, base, h, seed=1, label=True, grams=None):
    """SVG group for one bottle. cx/base in user units, h = overall height."""
    l1, l2, cap, seasoning, tones = FLAVOURS[key]
    bw = h * 0.415                       # body width
    x0, x1 = cx - bw / 2, cx + bw / 2
    body_top = base - h * 0.755          # where the straight body starts
    sh_top = base - h * 0.855            # shoulder top (neck base)
    neck_w = bw * 0.50
    cap_h = h * 0.125
    cap_w = neck_w * 1.30
    cap_top = base - h
    r = bw * 0.10
    uid = f"{key}{seed}"

    body = (f'M {x0:.1f} {base - r:.1f} '
            f'L {x0:.1f} {body_top:.1f} '
            f'C {x0:.1f} {body_top - h * 0.03:.1f} {cx - neck_w / 2:.1f} {sh_top + h * 0.045:.1f} '
            f'{cx - neck_w / 2:.1f} {sh_top:.1f} '
            f'L {cx - neck_w / 2:.1f} {cap_top + cap_h * 0.55:.1f} '
            f'L {cx + neck_w / 2:.1f} {cap_top + cap_h * 0.55:.1f} '
            f'L {cx + neck_w / 2:.1f} {sh_top:.1f} '
            f'C {cx + neck_w / 2:.1f} {sh_top + h * 0.045:.1f} {x1:.1f} {body_top - h * 0.03:.1f} '
            f'{x1:.1f} {body_top:.1f} '
            f'L {x1:.1f} {base - r:.1f} '
            f'Q {x1:.1f} {base:.1f} {x1 - r:.1f} {base:.1f} '
            f'L {x0 + r:.1f} {base:.1f} '
            f'Q {x0:.1f} {base:.1f} {x0:.1f} {base - r:.1f} Z')

    fill_top = body_top + h * 0.055
    fill_h = base - fill_top - h * 0.012
    lab_top = body_top + h * 0.18
    lab_h = h * 0.345
    lab_x0, lab_x1 = x0 + bw * 0.055, x1 - bw * 0.055

    parts = [f'<defs>',
             f'<clipPath id="clipBody{uid}"><path d="{body}"/></clipPath>',
             f'<linearGradient id="glass{uid}" x1="0" y1="0" x2="1" y2="0">'
             f'<stop offset="0" stop-color="#E4E4DE"/><stop offset="0.18" stop-color="#FCFCFA"/>'
             f'<stop offset="0.62" stop-color="#F0F0EA"/><stop offset="1" stop-color="#DCDBD4"/>'
             f'</linearGradient>',
             f'<linearGradient id="cap{uid}" x1="0" y1="0" x2="1" y2="0">'
             f'<stop offset="0" stop-color="{cap}" stop-opacity="0.78"/>'
             f'<stop offset="0.30" stop-color="{cap}"/>'
             f'<stop offset="1" stop-color="{cap}" stop-opacity="0.86"/></linearGradient>',
             f'<linearGradient id="seed{uid}" x1="0" y1="0" x2="0" y2="1">'
             f'<stop offset="0" stop-color="{seasoning}" stop-opacity="0.92"/>'
             f'<stop offset="1" stop-color="{seasoning}"/></linearGradient>',
             f'<linearGradient id="sheen{uid}" x1="0" y1="0" x2="1" y2="0">'
             f'<stop offset="0" stop-color="#FFFFFF" stop-opacity="0.05"/>'
             f'<stop offset="0.45" stop-color="#FFFFFF" stop-opacity="0.34"/>'
             f'<stop offset="1" stop-color="#FFFFFF" stop-opacity="0.02"/></linearGradient>',
             f'</defs>']

    # contact shadow
    parts.append(f'<ellipse cx="{cx:.1f}" cy="{base + h * 0.012:.1f}" rx="{bw * 0.62:.1f}" '
                 f'ry="{h * 0.020:.1f}" fill="#3A2A20" opacity="0.16"/>')
    # glass
    parts.append(f'<path d="{body}" fill="url(#glass{uid})" stroke="{GLASS_EDGE}" stroke-width="{h*0.006:.2f}"/>')
    # seasoning inside
    parts.append(f'<g clip-path="url(#clipBody{uid})">')
    parts.append(f'<rect x="{x0:.1f}" y="{fill_top:.1f}" width="{bw:.1f}" height="{fill_h:.1f}" fill="url(#seed{uid})"/>')
    parts.append(flakes(seed, x0, fill_top, bw, fill_h, tones))
    parts.append(f'<rect x="{x0:.1f}" y="{fill_top:.1f}" width="{bw:.1f}" height="{h*0.012:.1f}" '
                 f'fill="#FFFFFF" opacity="0.20"/>')
    # glass sheen over everything inside the silhouette
    parts.append(f'<rect x="{x0 + bw*0.07:.1f}" y="{sh_top:.1f}" width="{bw*0.16:.1f}" '
                 f'height="{base - sh_top:.1f}" fill="url(#sheen{uid})"/>')
    parts.append(f'<rect x="{x1 - bw*0.15:.1f}" y="{sh_top:.1f}" width="{bw*0.08:.1f}" '
                 f'height="{base - sh_top:.1f}" fill="url(#sheen{uid})" opacity="0.55"/>')
    parts.append('</g>')

    # label
    parts.append(f'<rect x="{lab_x0:.1f}" y="{lab_top:.1f}" width="{lab_x1-lab_x0:.1f}" '
                 f'height="{lab_h:.1f}" rx="{h*0.010:.1f}" fill="#FDFBF7"/>')
    parts.append(f'<path d="M {lab_x0:.1f} {lab_top + h*0.010:.1f} '
                 f'a {h*0.010:.1f} {h*0.010:.1f} 0 0 1 {h*0.010:.1f} {-h*0.010:.1f} '
                 f'L {lab_x1 - h*0.010:.1f} {lab_top:.1f} '
                 f'a {h*0.010:.1f} {h*0.010:.1f} 0 0 1 {h*0.010:.1f} {h*0.010:.1f} '
                 f'L {lab_x1:.1f} {lab_top + lab_h*0.30:.1f} L {lab_x0:.1f} {lab_top + lab_h*0.30:.1f} Z" '
                 f'fill="{cap}"/>')
    if label:
        fs_small = lab_h * 0.135
        parts.append(f'<text x="{cx:.1f}" y="{lab_top + lab_h*0.195:.1f}" text-anchor="middle" '
                     f'font-family="Liberation Sans, Arial" font-size="{fs_small:.1f}" '
                     f'letter-spacing="{fs_small*0.18:.2f}" fill="#FFFFFF" opacity="0.95">FURIKAKE</text>')
        lines = [l for l in (l1, l2) if l]
        avail = (lab_x1 - lab_x0) * 0.88
        fs = lab_h * (0.185 if len(lines) > 1 else 0.235)
        fs = min(fs, avail / (max(len(l) for l in lines) * 0.62))
        if len(lines) > 1:
            baselines = [lab_top + lab_h * 0.50, lab_top + lab_h * 0.50 + fs * 1.14]
        else:
            baselines = [lab_top + lab_h * 0.585]
        for ln, ty in zip(lines, baselines):
            parts.append(f'<text x="{cx:.1f}" y="{ty:.1f}" text-anchor="middle" '
                         f'font-family="Liberation Sans, Arial" font-weight="bold" '
                         f'font-size="{fs:.1f}" letter-spacing="{fs*0.04:.2f}" fill="{INK}">{ln}</text>')
        if grams:
            parts.append(f'<text x="{cx:.1f}" y="{lab_top + lab_h*0.925:.1f}" text-anchor="middle" '
                         f'font-family="Liberation Sans, Arial" font-size="{lab_h*0.125:.1f}" '
                         f'letter-spacing="{lab_h*0.012:.2f}" fill="#8A817A">{grams}</text>')

    # cap
    parts.append(f'<rect x="{cx - cap_w/2:.1f}" y="{cap_top:.1f}" width="{cap_w:.1f}" '
                 f'height="{cap_h:.1f}" rx="{h*0.016:.1f}" fill="url(#cap{uid})"/>')
    parts.append(f'<rect x="{cx - cap_w/2:.1f}" y="{cap_top + cap_h*0.66:.1f}" width="{cap_w:.1f}" '
                 f'height="{cap_h*0.16:.1f}" fill="#000000" opacity="0.14"/>')
    parts.append(f'<rect x="{cx - cap_w/2:.1f}" y="{cap_top:.1f}" width="{cap_w*0.22:.1f}" '
                 f'height="{cap_h:.1f}" rx="{h*0.016:.1f}" fill="#FFFFFF" opacity="0.18"/>')
    return "\n".join(parts)


def render(path, w_in, h_in, body, bg=PAPER, dpi=300):
    W, H = w_in * 100, h_in * 100          # user units: 100 per inch
    svg = (f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" '
           f'viewBox="0 0 {W} {H}"><rect width="{W}" height="{H}" fill="{bg}"/>{body}</svg>')
    cairosvg.svg2png(bytestring=svg.encode(), write_to=path,
                     output_width=round(w_in * dpi), output_height=round(h_in * dpi))
    print(path)


# Hero: three bottles, staggered heights, generous margins.
W, H = 328, 218
hero = "".join([
    bottle("nori",   cx=W * 0.215, base=H * 0.900, h=H * 0.700, seed=3, grams="50 g"),
    bottle("yuzu",   cx=W * 0.795, base=H * 0.885, h=H * 0.660, seed=11, grams="50 g"),
    bottle("wasabi", cx=W * 0.505, base=H * 0.945, h=H * 0.800, seed=7, grams="50 g"),
])
render("assets/takaokaya_hero.jpg", 3.28, 2.18, hero)

# Single tall bottle for the cost slide.
W2, H2 = 156, 205
render("assets/takaokaya_bottle.jpg", 1.56, 2.05,
       bottle("katsuo", cx=W2 * 0.5, base=H2 * 0.92, h=H2 * 0.80, seed=5, grams="45 g"))

# Compact unlabelled tiles per flavour.
for key in FLAVOURS:
    render(f"assets/tile_{key}.jpg", 0.80, 1.00,
           bottle(key, cx=40, base=94, h=80, seed=abs(hash(key)) % 97, label=False))
