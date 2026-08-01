const pptxgen = require("pptxgenjs");
const path = require("path");
const A = (p) => path.join(__dirname, p);

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
// PALETTE — lifted straight off the plant's Choi's Ramyeon mural
// ─────────────────────────────────────────────────────────────────────
const PINK     = "F04860";   // the painted wall
const PINK_DK  = "D2354E";
const PINK_LT  = "F4788C";   // watermark tint on pink
const YELLOW   = "FDB930";   // noodle
const IVORY    = "FDF6EC";
const INK      = "1C1512";
const INK_SOFT = "4C4038";
const GREY     = "8A7F76";
const HAIR     = "E2D6C6";
const CRIMSON  = "C8102E";   // prices

const FONT    = "Calibri";
const FONT_H  = "Cambria";
const FONT_KR = "Malgun Gothic";

const PW = 13.333, PH = 7.5;
const M   = 0.72;
const COL = PW - M * 2;

const RIBBON_ASPECT = 17.0;
const RIB_H = PW / RIBBON_ASPECT;    // 0.784 in

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
    txt(s, "CHOI'S RAMYEON — SIAS FRANCE", M, PH - 0.44, 6, 0.24, { size: 7.5, color: GREY, cs: 2 });
  txt(s, String(n).padStart(2, "0"), PW - M - 0.9, PH - 0.6, 0.9, 0.42,
    { size: 21, color: "E7DBCB", align: "right", font: FONT_H, bold: true });
}

// The signature: a pink masthead block with the noodle ribbon spilling out of it
function masthead(s, kicker, title, o = {}) {
  const blockH = o.blockH || 1.72;
  rect(s, 0, 0, PW, blockH, { fill: PINK });
  txt(s, "라면", PW - M - 3.2, blockH - 1.02, 3.2, 0.72,
    { font: FONT_KR, bold: true, size: 44, color: PINK_LT, align: "right" });
  txt(s, kicker, M, 0.42, 7, 0.26, { size: 8.5, bold: true, color: YELLOW, cs: 2.6 });
  if (title)
    txt(s, title, M - 0.04, 0.72, o.tw || 9.4, 0.74, { font: FONT_H, bold: true, size: o.size || 36, color: "FFFFFF" });
  pic(s, A("assets/noodle_ribbon.png"), 0, blockH - 0.2, PW, RIB_H);
  return blockH - 0.2 + RIB_H;    // where the page body may start
}

// ═════════════════════════════════════════════════════════════════════
// 01 · COVER
// ═════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: IVORY };
  rect(s, 0, 0, PW, 4.3, { fill: PINK });

  txt(s, "라면", PW - M - 4.2, 1.5, 4.2, 1.6, { font: FONT_KR, bold: true, size: 104, color: PINK_LT, align: "right" });

  txt(s, "SIAS FRANCE · ROYE, ФРАНЦІЯ", M, 0.5, 6, 0.26, { size: 9, bold: true, color: YELLOW, cs: 2.4 });
  txt(s, "ПРАЙС-КАТАЛОГ · 2026", PW - M - 5, 0.5, 5, 0.26, { size: 9, color: "FBD3DA", cs: 2.2, align: "right" });

  txt(s, "CHOI'S", M - 0.07, 1.42, 8, 0.86, { font: FONT_H, bold: true, size: 54, color: "FFFFFF" });
  txt(s, "RAMYEON", M - 0.07, 2.18, 9, 0.86, { font: FONT_H, bold: true, size: 54, color: YELLOW });

  txt(s,
    "Локшина швидкого приготування у форматі Bag, вироблена у Франції.\n" +
    "Собівартість однієї пачки для трьох обсягів постачання.",
    M, 3.1, 8.4, 0.6, { font: FONT_H, italic: true, size: 13.5, color: "FFEFF2", ls: 1.28 });

  const meta = [["6 SKU", M], ["112,5–131 Г", M + 2.3], ["20 ШТ / КОРОБ", M + 4.9], ["EXW ROYE", M + 7.8]];
  meta.forEach(([t, x]) => txt(s, t, x, 3.86, 2.6, 0.24, { size: 9, bold: true, color: "FFE3E8", cs: 1.4 }));
  txt(s, "EUR.1 · 0% МИТО", PW - M - 2.6, 3.86, 2.6, 0.24, { size: 9, bold: true, color: YELLOW, cs: 1.4, align: "right" });

  pic(s, A("assets/noodle_ribbon.png"), 0, 4.1, PW, RIB_H);

  const lh = 2.5, lw = lh * 4.474;
  pic(s, A("assets/cover_lineup.png"), (PW - lw) / 2, PH - lh, lw, lh);
}

// ═════════════════════════════════════════════════════════════════════
// 02 · THE MAKER
// ═════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: IVORY };
  masthead(s, "ВИРОБНИК", "SIAS France · Roye");

  const top = 2.52;
  txt(s, "최씨라면", M, top, 3.4, 0.34, { font: FONT_KR, bold: true, size: 16, color: PINK_DK, cs: 3 });
  txt(s,
    "Завод у Roye — європейська площадка корейської групи SIAS, що працює " +
    "з 1997 року. Лінійку Choi's Ramyeon виробляють тут за оригінальними " +
    "корейськими рецептурами.",
    M, top + 0.46, 5.5, 0.9, { size: 11.5, color: INK_SOFT, ls: 1.32 });

  const facts = [
    ["16 000 м²", "виробнича площадка", false],
    ["IFS · BRC", "сертифікати харчової безпеки", false],
    ["EXW Roye", "умови відвантаження", false],
    ["0 %", "ввізного мита в Україну · EUR.1", true],
  ];
  let fy = top + 1.5;
  facts.forEach((f) => {
    rule(s, M, fy, 5.5, HAIR);
    txt(s, f[0], M, fy + 0.13, 2.2, 0.34, { font: FONT_H, bold: true, size: f[2] ? 20 : 16, color: f[2] ? PINK_DK : INK });
    txt(s, f[1], M + 2.2, fy + 0.2, 3.3, 0.26, { size: 9, color: GREY, align: "right" });
    fy += 0.56;
  });
  rule(s, M, fy, 5.5, HAIR);

  const px = 6.62, pw = PW - M - px, ph = pw / 1.392;
  pic(s, A("assets/plant_mural.jpg"), px, top, pw, ph, { fit: "cover" });
  txt(s, "Завод SIAS · Roye, Hauts-de-France", px, top + ph + 0.14, 2.6, 0.28, { size: 9.5, italic: true, color: GREY });
  frFlag(s, px + 2.72, top + ph + 0.12, 0.46, 0.3);
  uaFlag(s, px + 3.3, top + ph + 0.12, 0.46, 0.3);

  folio(s, 2);
}

// ═════════════════════════════════════════════════════════════════════
// 03-05 · ONE SLIDE PER SHIPMENT VOLUME
// ═════════════════════════════════════════════════════════════════════
function volumeSlide(page, kicker, tier, big, small, lead) {
  const s = pres.addSlide();
  s.background = { color: IVORY };

  const pallets = Object.values(SPLIT[tier]).reduce((a, b) => a + b, 0);
  const blockH = 1.98;
  rect(s, 0, 0, PW, blockH, { fill: PINK });
  txt(s, kicker, M, 0.4, 7, 0.26, { size: 8.5, bold: true, color: YELLOW, cs: 2.6 });

  txt(s, big, M - 0.06, 0.66, 3.2, 0.9, { font: FONT_H, bold: true, size: 56, color: "FFFFFF" });
  txt(s, small, M + 2.3, 1.0, 4.6, 0.42, { font: FONT_H, italic: true, size: 20, color: YELLOW });
  txt(s, lead, M, 1.56, 7.2, 0.3, { size: 10.5, italic: true, color: "FFE3E8", font: FONT_H });

  const tot = [
    [num(pallets), "палет у постачанні"],
    [num(pallets * PER_PALLET), "пачок разом"],
    ["0 %", "мито · EUR.1"],
  ];
  tot.forEach((t, i) => {
    const x = PW - M - (3 - i) * 2.0;
    txt(s, t[0], x, 0.72, 1.85, 0.44, { font: FONT_H, bold: true, size: 22, color: i === 2 ? YELLOW : "FFFFFF", align: "right" });
    txt(s, t[1], x, 1.2, 1.85, 0.26, { size: 8.5, color: "FBD3DA", align: "right", cs: 0.6 });
  });

  pic(s, A("assets/noodle_ribbon.png"), 0, blockH - 0.2, PW, RIB_H);

  const gx = 0.35, cw = (COL - 2 * gx) / 3, ch = 2.04, y0 = 2.68;

  SKUS.forEach((sk, i) => {
    const c = i % 3, r = Math.floor(i / 3);
    const x = M + c * (cw + gx);
    const y = y0 + r * (ch + 0.16);
    const [e, u] = COST[sk.key][tier];
    const pal = SPLIT[tier][sk.key];

    pic(s, sk.img, x, y + 0.02, 1.28, ch - 0.06, { shadow: lift() });

    const tx = x + 1.42, tw = cw - 1.42;
    txt(s, "INSTANT NOODLE · RAMEN BAG", tx, y + 0.03, tw, 0.18, { size: 7, color: GREY, cs: 1.4 });
    txt(s, sk.name, tx, y + 0.21, tw, 0.44, { font: FONT_H, bold: true, size: 13.5, color: INK, ls: 1.0 });
    txt(s, sk.kr + "   ·   " + sk.g, tx, y + 0.68, tw, 0.22, { font: FONT_KR, size: 9.5, color: PINK_DK, cs: 0.6 });

    rule(s, tx, y + 0.94, tw - 0.08, HAIR);
    rect(s, tx, y + 1.06, 0.09, 0.09, { fill: PINK });
    txt(s, pal + " " + palletWord(pal), tx + 0.19, y + 1.0, 1.15, 0.24, { size: 11.5, bold: true, color: INK });
    txt(s, num(pal * PER_PALLET) + " пачок", tx + 1.32, y + 1.04, tw - 1.4, 0.22, { size: 8.5, color: GREY, align: "right" });

    txt(s, "СОБІВАРТІСТЬ ЗА 1 ПАЧКУ", tx, y + 1.34, tw, 0.18, { size: 6.8, bold: true, color: PINK_DK, cs: 1.2 });
    txt(s, eur(e), tx, y + 1.52, 1.25, 0.4, { font: FONT_H, bold: true, size: 23, color: CRIMSON });
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
  s.background = { color: IVORY };
  masthead(s, "РИНОК", "Де представлена продукція", { size: 34, tw: 9.4 });

  txt(s, "роздрібні мережі та дистриб'ютори, з якими працює SIAS",
    M, 2.56, 9.6, 0.3, { size: 11, italic: true, color: GREY, font: FONT_H });

  pic(s, A("assets/logo_wall.png"), M, 3.0, COL, COL / 2.936);

  folio(s, 6);
}

// ═════════════════════════════════════════════════════════════════════
pres.writeFile({ fileName: A("SIAS_France_Presentation.pptx") })
  .then(() => console.log("Deck written."))
  .catch((e) => { console.error(e); process.exit(1); });
