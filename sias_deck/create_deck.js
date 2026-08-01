const pptxgen = require("pptxgenjs");
const path = require("path");
const A = (p) => path.join(__dirname, p);

// ─────────────────────────────────────────────────────────────────────
// ONE SYSTEM FOR EVERY PAGE — warm paper, hairlines, a single red
// ─────────────────────────────────────────────────────────────────────
const INK      = "17130F";
const INK_SOFT = "4A423A";
const PAPER    = "FCFAF6";
const RED      = "C4101F";
const GOLD     = "9C7C42";
const GREY     = "8A8177";
const HAIR     = "DCD4C6";

const FONT    = "Calibri";
const FONT_H  = "Cambria";
const FONT_KR = "Malgun Gothic";

const PW = 13.333, PH = 7.5;
const M   = 0.72;
const COL = PW - M * 2;              // 11.893

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";

// ─────────────────────────────────────────────────────────────────────
// PRODUCT DATA — names exactly as in SelfCost.xlsx
// ─────────────────────────────────────────────────────────────────────
const SKUS = [
  { key: "gouda",     name: "GOUDA CHESSE",     kr: "고다치즈",   g: "123 G",   img: A("packs/gouda.png") },
  { key: "carbonara", name: "CARBONARA SPICY",  kr: "까르보나라", g: "131 G",   img: A("packs/carbonara.png") },
  { key: "vegetable", name: "VEGETABLE FLAVOR", kr: "야채맛",     g: "113 G",   img: A("packs/vegetable.png") },
  { key: "spicy",     name: "SPICY FLAVOR",     kr: "매운맛",     g: "112.5 G", img: A("packs/spicy.png") },
  { key: "beef",      name: "BEEF FLAVOR",      kr: "소고기맛",   g: "113 G",   img: A("packs/beef.png") },
  { key: "chicken",   name: "CHICKEN FLAVOR",   kr: "치킨맛",     g: "113 G",   img: A("packs/chicken.png") },
];

const COST = {   // landed cost of one pack — [EUR, UAH]; duty 0% under EUR.1
  gouda:     { p6: [1.31, 67.00], p12: [1.11, 56.88], p33: [1.02, 52.13] },
  carbonara: { p6: [1.31, 67.00], p12: [1.11, 56.88], p33: [1.02, 52.13] },
  vegetable: { p6: [1.13, 58.06], p12: [0.96, 49.29], p33: [0.88, 45.18] },
  spicy:     { p6: [1.13, 58.06], p12: [0.96, 49.29], p33: [0.88, 45.18] },
  beef:      { p6: [1.13, 58.06], p12: [0.96, 49.29], p33: [0.88, 45.18] },
  chicken:   { p6: [1.13, 58.06], p12: [0.96, 49.29], p33: [0.88, 45.18] },
};

const SPLIT = {  // pallets per SKU in each shipment
  p6:  { gouda: 1, carbonara: 1, vegetable: 1, spicy: 1, beef: 1, chicken: 1 },
  p12: { gouda: 3, carbonara: 3, vegetable: 1, spicy: 2, beef: 2, chicken: 1 },
  p33: { gouda: 7, carbonara: 7, vegetable: 4, spicy: 5, beef: 5, chicken: 5 },
};

const PER_PALLET = 1600;   // 20 шт/короб × 80 короб/палета

// ─────────────────────────────────────────────────────────────────────
// PRIMITIVES
// ─────────────────────────────────────────────────────────────────────
function rect(s, x, y, w, h, o = {}) {
  s.addShape("rect", {
    x, y, w, h,
    fill: o.fill ? { color: o.fill } : { type: "none" },
    line: o.line || { type: "none" },
  });
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

const lift = () => ({ type: "outer", color: "3C2E20", opacity: 0.22, blur: 14, offset: 6, angle: 90 });
const rule = (s, x, y, w, c) => rect(s, x, y, w, 0.009, { fill: c || HAIR });

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

function folio(s, n, o = {}) {
  if (!o.noLabel)
    txt(s, "CHOI'S RAMYEON — SIAS FRANCE", M, PH - 0.46, 6, 0.24, { size: 7.5, color: GREY, cs: 2 });
  txt(s, String(n).padStart(2, "0"), PW - M - 0.9, PH - 0.62, 0.9, 0.42,
    { size: 21, color: HAIR, align: "right", font: FONT_H, bold: true });
}

function head(s, kicker, title, sub, o = {}) {
  rule(s, M, 0.62, COL, HAIR);
  txt(s, kicker, M, 0.76, 7, 0.26, { size: 8.5, bold: true, color: RED, cs: 2.6 });
  txt(s, title, M - 0.05, 1.06, o.tw || 11.9, 0.82, { font: FONT_H, bold: true, size: o.size || 40, color: INK });
  if (sub) txt(s, sub, M, 1.96, o.sw || 9.6, 0.3, { size: 11.5, italic: true, color: GOLD, font: FONT_H });
}

// ═════════════════════════════════════════════════════════════════════
// 01 · COVER — same paper as every other page
// ═════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: PAPER };

  rule(s, M, 0.62, COL, HAIR);
  txt(s, "SIAS FRANCE · ROYE, ФРАНЦІЯ", M, 0.76, 6, 0.26, { size: 9, bold: true, color: RED, cs: 2.4 });
  txt(s, "ПРАЙС-КАТАЛОГ · 2026", PW - M - 5, 0.76, 5, 0.26, { size: 9, color: GREY, cs: 2.2, align: "right" });

  txt(s, "라면", M, 1.42, 3, 0.4, { font: FONT_KR, bold: true, size: 19, color: GOLD, cs: 5 });
  txt(s, "CHOI'S", M - 0.07, 1.72, 8, 0.92, { font: FONT_H, bold: true, size: 58, color: INK });
  txt(s, "RAMYEON", M - 0.07, 2.52, 9, 0.92, { font: FONT_H, bold: true, size: 58, color: RED });

  rule(s, M, 3.6, COL, HAIR);
  txt(s,
    "Локшина швидкого приготування у форматі Bag, вироблена у Франції.\n" +
    "Собівартість однієї пачки для трьох обсягів постачання.",
    M, 3.76, 8.4, 0.64, { font: FONT_H, italic: true, size: 14, color: INK_SOFT, ls: 1.3 });

  const meta = [["6 SKU", M], ["112,5–131 Г", M + 2.35], ["20 ШТ / КОРОБ", M + 5.0], ["EXW ROYE", M + 8.0]];
  meta.forEach(([t, x]) => txt(s, t, x, 4.62, 2.6, 0.26, { size: 9.5, bold: true, color: INK_SOFT, cs: 1.5 }));
  txt(s, "EUR.1 · 0% МИТО", PW - M - 2.6, 4.62, 2.6, 0.26, { size: 9.5, bold: true, color: RED, cs: 1.5, align: "right" });

  const lh = 2.5, lw = lh * 4.474;
  pic(s, A("assets/cover_lineup.png"), (PW - lw) / 2, PH - lh, lw, lh);
}

// ═════════════════════════════════════════════════════════════════════
// 02 · THE MAKER
// ═════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: PAPER };

  rule(s, M, 0.62, COL, HAIR);
  txt(s, "ВИРОБНИК", M, 0.76, 6, 0.26, { size: 8.5, bold: true, color: RED, cs: 2.6 });

  txt(s, "SIAS", M - 0.05, 1.12, 4.6, 0.9, { font: FONT_H, bold: true, size: 58, color: INK });
  txt(s, "France", M - 0.05, 1.96, 4.6, 0.84, { font: FONT_H, bold: true, italic: true, size: 50, color: RED });
  txt(s, "최씨라면", M, 2.96, 3.4, 0.34, { font: FONT_KR, bold: true, size: 16, color: GOLD, cs: 3 });

  rule(s, M, 3.48, 4.5, HAIR);
  txt(s,
    "Завод у Roye — європейська площадка корейської групи SIAS, що працює " +
    "з 1997 року. Лінійку Choi's Ramyeon виробляють тут за оригінальними " +
    "корейськими рецептурами.",
    M, 3.62, 4.5, 0.86, { size: 11.5, color: INK_SOFT, ls: 1.32 });

  const facts = [
    ["16 000 м²", "виробнича площадка", false],
    ["IFS · BRC", "сертифікати харчової безпеки", false],
    ["EXW Roye", "умови відвантаження", false],
    ["0 %", "ввізного мита в Україну · EUR.1", true],
  ];
  let fy = 4.66;
  facts.forEach((f) => {
    rule(s, M, fy, 4.5, HAIR);
    txt(s, f[0], M, fy + 0.13, 2.0, 0.34, { font: FONT_H, bold: true, size: f[2] ? 19 : 16, color: f[2] ? RED : INK });
    txt(s, f[1], M + 1.9, fy + 0.2, 2.6, 0.26, { size: 9, color: GREY, align: "right" });
    fy += 0.54;
  });
  rule(s, M, fy, 4.5, HAIR);

  // photograph — the plant's Choi's Ramyeon mural
  pic(s, A("assets/plant_mural.jpg"), 5.95, 1.12, 6.66, 4.79, { fit: "cover" });
  rule(s, 5.95, 6.08, 6.66, HAIR);
  txt(s, "Завод SIAS · Roye, Hauts-de-France", 5.95, 6.2, 4.4, 0.28, { size: 9.5, italic: true, color: GREY });
  frFlag(s, PW - M - 1.18, 6.18, 0.46, 0.3);
  uaFlag(s, PW - M - 0.46, 6.18, 0.46, 0.3);

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

  txt(s, big, M - 0.06, 1.0, 3.2, 0.94, { font: FONT_H, bold: true, size: 60, color: INK });
  txt(s, small, M + 2.45, 1.38, 4.6, 0.42, { font: FONT_H, italic: true, size: 21, color: RED });
  txt(s, lead, M, 1.98, 6.6, 0.3, { size: 11, italic: true, color: GOLD, font: FONT_H });

  const tot = [
    [num(pallets), "палет у постачанні"],
    [num(pallets * PER_PALLET), "пачок разом"],
    ["0 %", "мито · EUR.1"],
  ];
  tot.forEach((t, i) => {
    const x = PW - M - (3 - i) * 2.0;
    txt(s, t[0], x, 1.14, 1.85, 0.44, { font: FONT_H, bold: true, size: 22, color: i === 2 ? RED : INK, align: "right" });
    txt(s, t[1], x, 1.62, 1.85, 0.26, { size: 8.5, color: GREY, align: "right", cs: 0.6 });
  });

  rule(s, M, 2.42, COL, GOLD);

  const gx = 0.35, cw = (COL - 2 * gx) / 3, ch = 2.02, y0 = 2.64;

  SKUS.forEach((sk, i) => {
    const c = i % 3, r = Math.floor(i / 3);
    const x = M + c * (cw + gx);
    const y = y0 + r * (ch + 0.18);
    const [e, u] = COST[sk.key][tier];
    const pal = SPLIT[tier][sk.key];

    pic(s, sk.img, x, y + 0.02, 1.28, ch - 0.06, { shadow: lift() });

    const tx = x + 1.42, tw = cw - 1.42;
    txt(s, "INSTANT NOODLE · RAMEN BAG", tx, y + 0.03, tw, 0.18, { size: 7, color: GREY, cs: 1.4 });
    txt(s, sk.name, tx, y + 0.21, tw, 0.44, { font: FONT_H, bold: true, size: 13.5, color: INK, ls: 1.0 });
    txt(s, sk.kr + "   ·   " + sk.g, tx, y + 0.68, tw, 0.22, { font: FONT_KR, size: 9.5, color: GOLD, cs: 0.6 });

    // volume for this SKU — the pallet split, stated plainly
    rule(s, tx, y + 0.94, tw - 0.08, HAIR);
    rect(s, tx, y + 1.06, 0.09, 0.09, { fill: RED });
    txt(s, pal + " " + palletWord(pal), tx + 0.19, y + 1.0, 1.15, 0.24, { size: 11.5, bold: true, color: INK });
    txt(s, num(pal * PER_PALLET) + " пачок", tx + 1.32, y + 1.04, tw - 1.4, 0.22, { size: 8.5, color: GREY, align: "right" });

    txt(s, "СОБІВАРТІСТЬ ЗА 1 ПАЧКУ", tx, y + 1.34, tw, 0.18, { size: 6.8, bold: true, color: GOLD, cs: 1.2 });
    txt(s, eur(e), tx, y + 1.52, 1.25, 0.4, { font: FONT_H, bold: true, size: 23, color: RED });
    txt(s, uah(u), tx + 1.25, y + 1.64, tw - 1.33, 0.24, { size: 10.5, color: INK_SOFT, align: "right" });
  });

  rule(s, M, 6.94, COL, HAIR);
  txt(s,
    "Собівартість однієї пачки включає ціну EXW Roye, міжнародну логістику, брокерські послуги, митне оформлення та ПДВ-передфінансування. " +
    "Ставка ввізного мита 0% за сертифікатом EUR.1. Курс 1 € ≈ 51,2 ₴.",
    M, 7.06, 11.0, 0.24, { size: 8, italic: true, color: GREY });

  folio(s, page, { noLabel: true });
}

volumeSlide(3, "ОБСЯГ ПОСТАЧАННЯ · 01", "p6",  "6",  "палет",
  "Пробна партія — по одній палеті на кожен смак.");
volumeSlide(4, "ОБСЯГ ПОСТАЧАННЯ · 02", "p12", "12", "палет",
  "Робочий обсяг з акцентом на вершкових смаках — Gouda та Carbonara.");
volumeSlide(5, "ОБСЯГ ПОСТАЧАННЯ · 03", "p33", "33", "палети · Full Truck",
  "Повне авто — найнижча собівартість пачки.");

// ═════════════════════════════════════════════════════════════════════
// 06 · WHERE THE PRODUCT IS SOLD
// ═════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: PAPER };
  head(s, "РИНОК", "Де представлена продукція",
    "роздрібні мережі та дистриб'ютори, з якими працює SIAS", { size: 38 });

  rule(s, M, 2.46, COL, GOLD);
  pic(s, A("assets/logo_wall.png"), M, 2.72, COL, COL / 2.865);

  folio(s, 6);
}

// ═════════════════════════════════════════════════════════════════════
pres.writeFile({ fileName: A("SIAS_France_Presentation.pptx") })
  .then(() => console.log("Deck written."))
  .catch((e) => { console.error(e); process.exit(1); });
