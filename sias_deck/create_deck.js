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
// DARK, PRODUCT-FORWARD — colour comes from the packs, not the page
// (Buldak-style presentation, SIAS label energy)
// ─────────────────────────────────────────────────────────────────────
const BG      = "141110";   // warm near-black
const PANEL   = "1F1A17";
const CREAM   = "F6F1E9";
const MUTED   = "9C9088";
const FAINT   = "5E554E";
const HAIR    = "312A26";
const CORAL   = "F0485F";   // SIAS red, for price only
const AMBER   = "F5A623";   // noodle, for small labels
const PAPER   = "F6F1E9";

const FONT    = "Arial";
const FONT_KR = "Malgun Gothic";

const PW = 13.333, PH = 7.5;
const M   = 0.78;
const COL = PW - M * 2;

// ─────────────────────────────────────────────────────────────────────
function rect(s, x, y, w, h, o = {}) {
  s.addShape(o.round ? "roundRect" : "rect", Object.assign({
    x, y, w, h,
    fill: o.fill ? { color: o.fill } : { type: "none" },
    line: o.line || { type: "none" },
  }, o.round ? { rectRadius: o.round } : {}));
}

function txt(s, t, x, y, w, h, o = {}) {
  s.addText(t, {
    x, y, w, h,
    fontFace: o.font || FONT,
    fontSize: o.size || 11,
    color: o.color || CREAM,
    bold: !!o.bold, italic: !!o.italic,
    align: o.align || "left", valign: o.valign || "top",
    lineSpacingMultiple: o.ls || 1.06,
    charSpacing: o.cs, margin: 0,
  });
}

function pic(s, p, x, y, w, h, o = {}) {
  const im = { path: p, x, y, w, h, sizing: { type: o.fit || "contain", w, h } };
  if (o.shadow) im.shadow = o.shadow;
  s.addImage(im);
}

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
    txt(s, "CHOI'S RAMYEON — SIAS FRANCE", M, PH - 0.44, 6, 0.22, { size: 7.5, color: FAINT, cs: 2 });
  txt(s, String(n).padStart(2, "0"), PW - M - 0.8, PH - 0.56, 0.8, 0.34,
    { size: 15, color: FAINT, align: "right", bold: true });
}

// page head: hairline, amber kicker, heavy uppercase headline
function head(s, kicker, title, o = {}) {
  rule(s, M, 0.6, COL);
  txt(s, kicker, M, 0.74, 7, 0.24, { size: 8, bold: true, color: AMBER, cs: 2.8 });
  if (title) txt(s, title, M - 0.03, 1.02, o.tw || 9.6, 0.7, { size: o.size || 34, bold: true, color: CREAM, cs: -0.4 });
}

// ═════════════════════════════════════════════════════════════════════
// 01 · COVER
// ═════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: BG };

  rule(s, M, 0.6, COL);
  txt(s, "SIAS FRANCE · ROYE", M, 0.74, 6, 0.24, { size: 8, bold: true, color: AMBER, cs: 2.8 });
  txt(s, "ПРАЙС-КАТАЛОГ · 2026", PW - M - 5, 0.74, 5, 0.24, { size: 8, color: MUTED, cs: 2.4, align: "right" });

  txt(s, "최씨라면", M, 1.34, 4, 0.36, { font: FONT_KR, bold: true, size: 15, color: CORAL, cs: 3 });
  txt(s, "CHOI'S RAMYEON", M - 0.05, 1.7, 11.6, 1.0, { size: 66, bold: true, color: CREAM, cs: -2 });

  txt(s, "Локшина швидкого приготування, вироблена у Франції.", M, 2.84, 8.6, 0.3,
    { size: 13.5, color: MUTED });
  txt(s, "Собівартість однієї пачки для трьох обсягів постачання.", M, 3.14, 8.6, 0.3,
    { size: 13.5, color: MUTED });

  pic(s, A("assets/noodle_line.png"), M, 3.66, COL, COL / 37.778);

  const meta = [["6", "SKU"], ["112,5–131", "ГРАМІВ"], ["20", "ШТ / КОРОБ"], ["0 %", "МИТО · EUR.1"]];
  meta.forEach((mt, i) => {
    const x = M + i * 2.5;
    txt(s, mt[0], x, 3.92, 2.3, 0.36, { size: 21, bold: true, color: i === 3 ? CORAL : CREAM });
    txt(s, mt[1], x, 4.3, 2.3, 0.22, { size: 8, color: MUTED, cs: 1.4 });
  });
  txt(s, "EXW ROYE, FRANCE", PW - M - 3, 3.98, 3, 0.26, { size: 9.5, bold: true, color: MUTED, cs: 1.6, align: "right" });

  const lw = COL + 0.9, lh = lw / 4.359;
  pic(s, A("assets/lineup_dark.png"), (PW - lw) / 2, PH - lh - 0.16, lw, lh);
}

// ═════════════════════════════════════════════════════════════════════
// 02 · THE MAKER
// ═════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: BG };
  head(s, "ВИРОБНИК", "SIAS FRANCE · ROYE");

  const top = 2.06;
  txt(s,
    "Європейська площадка корейської групи SIAS, що працює з 1997 року. " +
    "Лінійку Choi's Ramyeon виробляють тут за оригінальними корейськими рецептурами.",
    M, top, 5.2, 0.9, { size: 12, color: MUTED, ls: 1.4 });

  const facts = [
    ["16 000 м²", "виробнича площадка", false],
    ["IFS · BRC", "сертифікати харчової безпеки", false],
    ["EXW Roye", "умови відвантаження", false],
    ["0 %", "ввізного мита в Україну · EUR.1", true],
  ];
  let fy = top + 1.12;
  facts.forEach((f) => {
    rule(s, M, fy, 5.2);
    txt(s, f[0], M, fy + 0.16, 2.2, 0.32, { size: f[2] ? 19 : 15, bold: true, color: f[2] ? CORAL : CREAM });
    txt(s, f[1], M + 2.2, fy + 0.22, 3.0, 0.24, { size: 9, color: MUTED, align: "right" });
    fy += 0.62;
  });
  rule(s, M, fy, 5.2);

  const px = 6.5, pw = PW - M - px, ph = pw / 1.392;
  pic(s, A("assets/plant_mural.jpg"), px, top - 0.1, pw, ph, { fit: "cover" });
  txt(s, "Завод SIAS · Roye, Hauts-de-France", px, top + ph + 0.06, 2.8, 0.26, { size: 9, color: MUTED });
  frFlag(s, px + 2.9, top + ph + 0.04, 0.44, 0.29);
  uaFlag(s, px + 3.46, top + ph + 0.04, 0.44, 0.29);

  folio(s, 2);
}

// ═════════════════════════════════════════════════════════════════════
// 03-05 · SHIPMENT VOLUMES
// ═════════════════════════════════════════════════════════════════════
function volumeSlide(page, kicker, tier, big, small, lead) {
  const s = pres.addSlide();
  s.background = { color: BG };

  const pallets = Object.values(SPLIT[tier]).reduce((a, b) => a + b, 0);

  rule(s, M, 0.6, COL);
  txt(s, kicker, M, 0.74, 7, 0.24, { size: 8, bold: true, color: AMBER, cs: 2.8 });

  txt(s, big + " " + small, M - 0.03, 1.0, 7.4, 0.74, { size: 40, bold: true, color: CREAM, cs: -1 });
  txt(s, lead, M, 1.82, 7.4, 0.28, { size: 11, color: MUTED });

  const tot = [
    [num(pallets), "палет"],
    [num(pallets * PER_PALLET), "пачок разом"],
    ["0 %", "мито · EUR.1"],
  ];
  tot.forEach((t, i) => {
    const x = PW - M - (3 - i) * 1.95;
    txt(s, t[0], x, 1.06, 1.8, 0.4, { size: 21, bold: true, color: i === 2 ? CORAL : CREAM, align: "right" });
    txt(s, t[1], x, 1.5, 1.8, 0.24, { size: 8.5, color: MUTED, align: "right", cs: 0.8 });
  });

  rule(s, M, 2.24, COL);

  const gx = 0.42, cw = (COL - 2 * gx) / 3, ch = 2.14, y0 = 2.44;

  SKUS.forEach((sk, i) => {
    const c = i % 3, r = Math.floor(i / 3);
    const x = M + c * (cw + gx);
    const y = y0 + r * (ch + 0.2);
    const [e, u] = COST[sk.key][tier];
    const pal = SPLIT[tier][sk.key];

    pic(s, sk.img, x, y + 0.02, 1.36, ch - 0.1);

    const tx = x + 1.52, tw = cw - 1.52;
    txt(s, sk.name, tx, y + 0.06, tw, 0.42, { size: 12.5, bold: true, color: CREAM, cs: -0.2, ls: 1.1 });
    txt(s, sk.kr + "   ·   " + sk.g, tx, y + 0.52, tw, 0.22, { font: FONT_KR, size: 9, color: MUTED });

    rule(s, tx, y + 0.84, tw - 0.06);
    rect(s, tx, y + 0.96, 0.08, 0.08, { fill: AMBER });
    txt(s, pal + " " + palletWord(pal), tx + 0.18, y + 0.9, 1.2, 0.24, { size: 11, bold: true, color: CREAM });
    txt(s, num(pal * PER_PALLET) + " пачок", tx + 1.36, y + 0.93, tw - 1.42, 0.22, { size: 8.5, color: MUTED, align: "right" });

    txt(s, "СОБІВАРТІСТЬ ЗА 1 ПАЧКУ", tx, y + 1.26, tw, 0.18, { size: 7, bold: true, color: AMBER, cs: 1.2 });
    txt(s, eur(e), tx, y + 1.46, 1.5, 0.46, { size: 27, bold: true, color: CORAL });
    txt(s, uah(u), tx + 1.5, y + 1.6, tw - 1.56, 0.24, { size: 10.5, color: MUTED, align: "right" });
  });

  txt(s,
    "Собівартість однієї пачки включає ціну EXW Roye, логістику, брокерські послуги, митне оформлення та ПДВ-передфінансування · мито 0% за EUR.1 · курс 1 € ≈ 51,2 ₴",
    M, 7.08, 11.2, 0.22, { size: 7.5, color: FAINT });

  folio(s, page, { noLabel: true });
}

volumeSlide(3, "ОБСЯГ ПОСТАЧАННЯ · 01", "p6",  "6",  "ПАЛЕТ",
  "Пробна партія — по одній палеті на кожен смак.");
volumeSlide(4, "ОБСЯГ ПОСТАЧАННЯ · 02", "p12", "12", "ПАЛЕТ",
  "Робочий обсяг з акцентом на вершкових смаках — Gouda та Carbonara.");
volumeSlide(5, "ОБСЯГ ПОСТАЧАННЯ · 03", "p33", "33", "ПАЛЕТИ · FULL TRUCK",
  "Повне авто — найнижча собівартість пачки.");

// ═════════════════════════════════════════════════════════════════════
// 06 · WHERE THE PRODUCT IS SOLD
// ═════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: BG };
  head(s, "РИНОК", "ДЕ ПРЕДСТАВЛЕНА ПРОДУКЦІЯ");

  txt(s, "роздрібні мережі та дистриб'ютори, з якими працює SIAS",
    M, 1.82, 9.6, 0.28, { size: 11, color: MUTED });

  const ww = COL, wh = ww / 3.04;
  rect(s, M - 0.16, 2.34, ww + 0.32, wh + 0.32, { fill: PAPER, round: 0.1 });
  pic(s, A("assets/logo_wall.png"), M, 2.5, ww, wh);

  folio(s, 6);
}

// ═════════════════════════════════════════════════════════════════════
pres.writeFile({ fileName: A("SIAS_France_Presentation.pptx") })
  .then(() => console.log("Deck written."))
  .catch((e) => { console.error(e); process.exit(1); });
