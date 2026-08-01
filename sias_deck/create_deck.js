const pptxgen = require("pptxgenjs");
const path = require("path");
const A = (p) => path.join(__dirname, p);

// ─────────────────────────────────────────────────────────────────────
// EDITORIAL SYSTEM — warm paper, hairlines, one red, no boxed cards
// ─────────────────────────────────────────────────────────────────────
const INK      = "17130F";
const INK_SOFT = "4A423A";
const PAPER    = "FCFAF6";
const RED      = "C4101F";
const GOLD     = "9C7C42";
const GREY     = "8A8177";
const HAIR     = "DCD4C6";

const FONT   = "Calibri";
const FONT_H = "Cambria";
const FONT_KR= "Malgun Gothic";

const PW = 13.333, PH = 7.5;
const M  = 0.72;                 // page margin
const COL = PW - M * 2;          // 11.893

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";

// ─────────────────────────────────────────────────────────────────────
// DATA — SKU names exactly as they appear in SelfCost.xlsx
// ─────────────────────────────────────────────────────────────────────
const SKUS = [
  { key: "gouda",     line: "INSTANT NOODLE · RAMEN BAG", name: "GOUDA CHESSE",    kr: "고다치즈",   g: "123 G",   img: A("packs/gouda.png") },
  { key: "carbonara", line: "INSTANT NOODLE · RAMEN BAG", name: "CARBONARA SPICY", kr: "까르보나라", g: "131 G",   img: A("packs/carbonara.png") },
  { key: "vegetable", line: "INSTANT NOODLE · RAMEN BAG", name: "VEGETABLE FLAVOR",kr: "야채맛",     g: "113 G",   img: A("packs/vegetable.png") },
  { key: "spicy",     line: "INSTANT NOODLE · RAMEN BAG", name: "SPICY FLAVOR",    kr: "매운맛",     g: "112.5 G", img: A("packs/spicy.png") },
  { key: "beef",      line: "INSTANT NOODLE · RAMEN BAG", name: "BEEF FLAVOR",     kr: "소고기맛",   g: "113 G",   img: A("packs/beef.png") },
  { key: "chicken",   line: "INSTANT NOODLE · RAMEN BAG", name: "CHICKEN FLAVOR",  kr: "치킨맛",     g: "113 G",   img: A("packs/chicken.png") },
];

// landed cost of one pack — [EUR, UAH]; duty 0% under EUR.1
const COST = {
  gouda:     { p6: [1.31, 67.00], p12: [1.11, 56.88], p33: [1.02, 52.13] },
  carbonara: { p6: [1.31, 67.00], p12: [1.11, 56.88], p33: [1.02, 52.13] },
  vegetable: { p6: [1.13, 58.06], p12: [0.96, 49.29], p33: [0.88, 45.18] },
  spicy:     { p6: [1.13, 58.06], p12: [0.96, 49.29], p33: [0.88, 45.18] },
  beef:      { p6: [1.13, 58.06], p12: [0.96, 49.29], p33: [0.88, 45.18] },
  chicken:   { p6: [1.13, 58.06], p12: [0.96, 49.29], p33: [0.88, 45.18] },
};

const SPLIT = {
  p6:  { gouda: 1, carbonara: 1, vegetable: 1, spicy: 1, beef: 1, chicken: 1 },
  p12: { gouda: 3, carbonara: 3, vegetable: 1, spicy: 2, beef: 2, chicken: 1 },
  p33: { gouda: 7, carbonara: 7, vegetable: 4, spicy: 5, beef: 5, chicken: 5 },
};

const PER_PALLET = 1600;   // 20 шт/короб × 80 короб/палета

// ─────────────────────────────────────────────────────────────────────
// PRIMITIVES
// ─────────────────────────────────────────────────────────────────────
function rect(s, x, y, w, h, o = {}) {
  const sh = { x, y, w, h, fill: o.fill ? { color: o.fill } : { type: "none" }, line: o.line || { type: "none" } };
  if (o.fill && o.transparency !== undefined) sh.fill.transparency = o.transparency;
  s.addShape("rect", sh);
}

function txt(s, t, x, y, w, h, o = {}) {
  s.addText(t, {
    x, y, w, h,
    fontFace: o.font || FONT,
    fontSize: o.size || 11,
    color: o.color || INK,
    bold: !!o.bold, italic: !!o.italic,
    align: o.align || "left", valign: o.valign || "top",
    lineSpacingMultiple: o.ls || 1.05,
    charSpacing: o.cs, margin: 0,
  });
}

function pic(s, p, x, y, w, h, o = {}) {
  const im = { path: p, x, y, w, h, sizing: { type: o.fit || "contain", w, h } };
  if (o.shadow) im.shadow = o.shadow;
  s.addImage(im);
}

const lift = () => ({ type: "outer", color: "2A1F16", opacity: 0.26, blur: 15, offset: 7, angle: 90 });

function rule(s, x, y, w, c) { rect(s, x, y, w, 0.009, { fill: c || HAIR }); }

const eur = (n) => n.toFixed(2).replace(".", ",") + " €";
const uah = (n) => n.toFixed(2).replace(".", ",") + " ₴";
const num = (n) => n.toLocaleString("uk-UA").replace(/ /g, " ");
const palletWord = (n) => (n === 1 ? "палета" : n <= 4 ? "палети" : "палет");

function frFlag(s, x, y, w, h) {
  const g = w / 3;
  rect(s, x, y, g, h, { fill: "002654" });
  rect(s, x + g, y, g, h, { fill: "FFFFFF" });
  rect(s, x + 2 * g, y, g, h, { fill: "ED2939" });
}
function uaFlag(s, x, y, w, h) {
  rect(s, x, y, w, h / 2, { fill: "0057B7" });
  rect(s, x, y + h / 2, w, h / 2, { fill: "FFD700" });
}

// large editorial folio
function folio(s, n, o = {}) {
  if (!o.noLabel) txt(s, "CHOI'S RAMYEON — SIAS FRANCE", M, PH - 0.46, 6, 0.24, { size: 7.5, color: GREY, cs: 2 });
  txt(s, String(n).padStart(2, "0"), PW - M - 0.9, PH - 0.62, 0.9, 0.42,
    { size: 21, color: HAIR, align: "right", font: FONT_H, bold: true });
}

// chapter head: hairline above, tiny tracked kicker, oversized serif title
function head(s, kicker, title, sub, opts = {}) {
  rule(s, M, 0.62, COL, HAIR);
  txt(s, kicker, M, 0.76, 7, 0.26, { size: 8.5, bold: true, color: RED, cs: 2.6 });
  txt(s, title, M - 0.05, 1.06, opts.tw || 9.6, 0.86, { font: FONT_H, bold: true, size: opts.size || 44, color: INK });
  if (sub) txt(s, sub, M, 1.98, opts.sw || 8.6, 0.3, { size: 11.5, italic: true, color: GOLD, font: FONT_H });
}

// ═════════════════════════════════════════════════════════════════════
// 01 · COVER
// ═════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: INK };
  pic(s, A("assets/cover_bg.jpg"), 0, 0, PW, PH, { fit: "cover" });

  rule(s, M, 0.64, COL, "5A4A38");
  txt(s, "SIAS FRANCE", M, 0.8, 5, 0.28, { size: 10, bold: true, color: "D9C69E", cs: 3.4 });
  txt(s, "ПРАЙС-КАТАЛОГ · 2026", PW - M - 5, 0.8, 5, 0.28, { size: 10, color: "9C9084", cs: 2.4, align: "right" });

  txt(s, "라면", M, 1.44, 3, 0.4, { font: FONT_KR, bold: true, size: 19, color: "C2A468", cs: 5 });
  txt(s, "CHOI'S", M - 0.07, 1.74, 8, 0.92, { font: FONT_H, bold: true, size: 58, color: "FFFFFF" });
  txt(s, "RAMYEON", M - 0.07, 2.54, 9, 0.92, { font: FONT_H, bold: true, size: 58, color: "FFFFFF" });

  txt(s, "Локшина швидкого приготування, вироблена у Франції.\nСобівартість пачки для 6, 12 та 33 палет.",
    M, 3.5, 7.4, 0.62, { font: FONT_H, italic: true, size: 13.5, color: "E6DCCB", ls: 1.3 });

  rule(s, M, 4.3, COL, "4E4034");
  txt(s, "6 СМАКІВ", M, 4.42, 2.4, 0.26, { size: 9.5, bold: true, color: "CBBEAA", cs: 1.6 });
  txt(s, "ФОРМАТ BAG · 112.5–131 G", M + 2.5, 4.42, 3.6, 0.26, { size: 9.5, bold: true, color: "CBBEAA", cs: 1.6 });
  txt(s, "EXW ROYE", M + 6.4, 4.42, 2.2, 0.26, { size: 9.5, bold: true, color: "CBBEAA", cs: 1.6 });
  txt(s, "EUR.1 · 0% МИТО", PW - M - 2.6, 4.42, 2.6, 0.26, { size: 9.5, bold: true, color: "E0C98F", cs: 1.6, align: "right" });

  // lineup box matches the artwork's 4.474 aspect exactly, so nothing crops
  const lh = 2.62, lw = lh * 4.474;
  pic(s, A("assets/cover_lineup.png"), (PW - lw) / 2, PH - lh, lw, lh, { fit: "contain" });
}

// ═════════════════════════════════════════════════════════════════════
// 02 · THE MAKER — image in a field of paper, wide margins
// ═════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: PAPER };

  rule(s, M, 0.62, COL, HAIR);
  txt(s, "ВИРОБНИК", M, 0.76, 6, 0.26, { size: 8.5, bold: true, color: RED, cs: 2.6 });

  txt(s, "SIAS", M - 0.05, 1.14, 4.6, 0.92, { font: FONT_H, bold: true, size: 60, color: INK });
  txt(s, "France", M - 0.05, 2.02, 4.6, 0.86, { font: FONT_H, bold: true, italic: true, size: 52, color: RED });
  txt(s, "최씨라면", M, 3.06, 3.4, 0.36, { font: FONT_KR, bold: true, size: 17, color: GOLD, cs: 3 });

  rule(s, M, 3.6, 4.4, HAIR);
  txt(s,
    "Європейська виробнича площадка корейської групи SIAS. Тут, у Roye, " +
    "виробляють лінійку Choi's Ramyeon за корейськими рецептурами.",
    M, 3.74, 4.4, 0.62, { size: 12, color: INK_SOFT, ls: 1.34 });

  const facts = [
    ["16 000 м²", "виробнича площадка", false],
    ["IFS · BRC", "харчова безпека", false],
    ["EXW Roye", "умови відвантаження", false],
    ["0 %", "мито в Україну · EUR.1", true],
  ];
  let fy = 4.55;
  facts.forEach((f) => {
    rule(s, M, fy, 4.4, HAIR);
    txt(s, f[0], M, fy + 0.13, 2.2, 0.34, { font: FONT_H, bold: true, size: f[2] ? 19 : 16, color: f[2] ? RED : INK });
    txt(s, f[1], M + 2.2, fy + 0.2, 2.2, 0.26, { size: 9.5, color: GREY, align: "right" });
    fy += 0.56;
  });
  rule(s, M, fy, 4.4, HAIR);

  // the photograph, generous white around it
  pic(s, A("assets/plant_mural.jpg"), 5.85, 1.14, 6.76, 3.55, { fit: "cover" });
  rule(s, 5.85, 4.86, 6.76, HAIR);
  txt(s, "Завод SIAS · Roye, Hauts-de-France", 5.85, 4.98, 4.4, 0.28, { size: 9.5, italic: true, color: GREY });
  frFlag(s, PW - M - 1.18, 4.96, 0.46, 0.3);
  uaFlag(s, PW - M - 0.46, 4.96, 0.46, 0.3);

  folio(s, 2);
}

// ═════════════════════════════════════════════════════════════════════
// 03-05 · ONE SLIDE PER SHIPMENT VOLUME
// ═════════════════════════════════════════════════════════════════════
function volumeSlide(page, kicker, tier, big, small, lead) {
  const s = pres.addSlide();
  s.background = { color: PAPER };

  const pallets = Object.values(SPLIT[tier]).reduce((a, b) => a + b, 0);

  rule(s, M, 0.62, COL, HAIR);
  txt(s, kicker, M, 0.76, 7, 0.26, { size: 8.5, bold: true, color: RED, cs: 2.6 });

  txt(s, big, M - 0.06, 1.02, 3.2, 0.96, { font: FONT_H, bold: true, size: 62, color: INK });
  txt(s, small, M + 2.5, 1.42, 4.4, 0.42, { font: FONT_H, italic: true, size: 21, color: RED });
  txt(s, lead, M, 2.02, 6.4, 0.3, { size: 11, italic: true, color: GOLD, font: FONT_H });

  // running totals, set as type rather than boxes
  const tot = [
    [num(pallets * PER_PALLET), "пачок разом"],
    ["1 600", "пачок у палеті"],
    ["0 %", "мито · EUR.1"],
  ];
  tot.forEach((t, i) => {
    const x = PW - M - (3 - i) * 1.95;
    txt(s, t[0], x, 1.16, 1.8, 0.44, { font: FONT_H, bold: true, size: 22, color: i === 2 ? RED : INK, align: "right" });
    txt(s, t[1], x, 1.64, 1.8, 0.26, { size: 8.5, color: GREY, align: "right", cs: 0.6 });
  });

  rule(s, M, 2.44, COL, GOLD);

  // 3 × 2 editorial grid — no cards, hairlines only
  const gx = 0.35, cw = (COL - 2 * gx) / 3, ch = 2.0, y0 = 2.66;

  SKUS.forEach((sk, i) => {
    const c = i % 3, r = Math.floor(i / 3);
    const x = M + c * (cw + gx);
    const y = y0 + r * (ch + 0.18);
    const [e, u] = COST[sk.key][tier];
    const pal = SPLIT[tier][sk.key];

    pic(s, sk.img, x, y + 0.02, 1.28, ch - 0.06, { shadow: lift() });

    const tx = x + 1.42, tw = cw - 1.42;
    txt(s, sk.line, tx, y + 0.04, tw, 0.18, { size: 7, color: GREY, cs: 1.4 });
    txt(s, sk.name, tx, y + 0.22, tw, 0.46, { font: FONT_H, bold: true, size: 13.5, color: INK, ls: 1.0 });
    txt(s, sk.kr, tx, y + 0.7, tw, 0.2, { font: FONT_KR, size: 9, color: GOLD, cs: 1 });

    rule(s, tx, y + 0.94, tw - 0.08, HAIR);
    txt(s, sk.g, tx, y + 1.02, 1.1, 0.24, { size: 11, bold: true, color: INK });
    txt(s, pal + " " + palletWord(pal), tx + 1.1, y + 1.05, tw - 1.18, 0.22, { size: 8.5, color: GREY, align: "right" });

    txt(s, "СОБІВАРТІСТЬ ЗА 1 ПАЧКУ", tx, y + 1.32, tw, 0.18, { size: 6.8, bold: true, color: GOLD, cs: 1.2 });
    txt(s, eur(e), tx, y + 1.5, 1.25, 0.4, { font: FONT_H, bold: true, size: 23, color: RED });
    txt(s, uah(u), tx + 1.25, y + 1.62, tw - 1.33, 0.24, { size: 10.5, color: INK_SOFT, align: "right" });
  });

  rule(s, M, 6.94, COL, HAIR);
  txt(s, "Собівартість 1 пачки: EXW Roye + логістика + брокер + митне оформлення + ПДВ-передфінансування · мито 0% (EUR.1) · курс 1 € ≈ 51,2 ₴",
    M, 7.06, 10.4, 0.24, { size: 8, italic: true, color: GREY });

  folio(s, page, { noLabel: true });
}

volumeSlide(3, "ВАРІАНТ ПОСТАЧАННЯ · 01", "p6",  "6",  "палет",              "пробна партія · по 1 палеті на кожен смак");
volumeSlide(4, "ВАРІАНТ ПОСТАЧАННЯ · 02", "p12", "12", "палет",              "середній обсяг");
volumeSlide(5, "ВАРІАНТ ПОСТАЧАННЯ · 03", "p33", "33", "палети · Full Truck", "повне авто · найнижча собівартість");

// ═════════════════════════════════════════════════════════════════════
// 06 · WHERE THE PRODUCT IS SOLD
// ═════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: PAPER };
  head(s, "РИНОК", "Де представлена продукція", "мережі та дистриб'ютори, з якими працює SIAS",
    { size: 38, tw: 11.9 });

  rule(s, M, 2.5, COL, GOLD);

  const LOGOS = [
    "carrefour", "leclerc", "systemeu", "aldi", "eurospin", "bofrost",
    "sysco", "hellofresh", "wismettac", "geia", "bibars", "dipsa",
    "forezia", "ttfoods", "gotiger", "foodex", "senko",
  ];
  const cols = 6, gx = 0.24;
  const cw = (COL - (cols - 1) * gx) / cols;
  const ch = 1.2, y0 = 2.72;

  LOGOS.forEach((name, i) => {
    const c = i % cols, r = Math.floor(i / cols);
    const x = M + c * (cw + gx);
    const y = y0 + r * (ch + 0.24);
    pic(s, A(`logos/${name}.png`), x + 0.12, y + 0.12, cw - 0.24, ch - 0.24);
    if (r < 2) rule(s, x, y + ch + 0.12, cw, HAIR);
  });

  folio(s, 6);
}

// ═════════════════════════════════════════════════════════════════════
pres.writeFile({ fileName: A("SIAS_France_Presentation.pptx") })
  .then(() => console.log("Deck written."))
  .catch((e) => { console.error(e); process.exit(1); });
