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
// DARK BUT ALIVE — warm ember ground, SIAS flames and characters as the
// personality, colour otherwise reserved for the packs
// ─────────────────────────────────────────────────────────────────────
const CREAM   = "F6F1E9";
const MUTED   = "9C9088";
const FAINT   = "5E554E";
const HAIR    = "342C27";
const CORAL   = "F0485F";
const AMBER   = "F5A623";
const PAPER   = "F6F1E9";

const FONT    = "Arial";
const FONT_KR = "Malgun Gothic";

const PW = 13.333, PH = 7.5;
const M   = 0.78;
const COL = PW - M * 2;

const BGIMG = () => A("assets/bg_dark.jpg");

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

function page(s) { s.background = { color: "141110" }; pic(s, BGIMG(), 0, 0, PW, PH, { fit: "cover" }); }

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
  page(s);

  pic(s, A("assets/sias_logo.png"), M, 0.52, 1.5, 0.64);
  txt(s, "ВИРОБНИК · SIAS FRANCE, ROYE", M + 1.72, 0.72, 5, 0.24, { size: 8.5, bold: true, color: AMBER, cs: 2.4 });
  txt(s, "ПРАЙС-КАТАЛОГ · 2026", PW - M - 5, 0.72, 5, 0.24, { size: 8.5, color: MUTED, cs: 2.2, align: "right" });
  rule(s, M, 1.34, COL);

  txt(s, "최씨라면", M, 1.56, 4, 0.34, { font: FONT_KR, bold: true, size: 14, color: CORAL, cs: 3 });
  txt(s, "CHOI'S RAMYEON", M - 0.05, 1.88, 10.4, 0.96, { size: 62, bold: true, color: CREAM, cs: -2 });

  txt(s, "Локшина швидкого приготування від корейської групи SIAS,", M, 2.94, 8.8, 0.3, { size: 13, color: MUTED });
  txt(s, "вироблена на власному заводі у Франції.", M, 3.24, 8.8, 0.3, { size: 13, color: MUTED });

  // the point that matters commercially
  rect(s, M, 3.74, 0.055, 0.5, { fill: CORAL });
  txt(s, "SIAS — виробника цього рівня складно знайти на ринку.", M + 0.24, 3.74, 8.6, 0.26,
    { size: 12, bold: true, color: CREAM });
  txt(s, "Прямий контракт із заводом, без посередників.", M + 0.24, 4.0, 8.6, 0.26, { size: 12, color: MUTED });

  pic(s, A("assets/noodle_line.png"), M, 4.46, COL, COL / 37.778);

  const meta = [["6", "SKU"], ["112,5–131", "ГРАМІВ"], ["0 %", "МИТО · EUR.1"]];
  meta.forEach((mt, i) => {
    const x = M + i * 2.6;
    txt(s, mt[0], x, 4.66, 2.4, 0.34, { size: 20, bold: true, color: i === 2 ? CORAL : CREAM });
    txt(s, mt[1], x, 5.02, 2.4, 0.22, { size: 8, color: MUTED, cs: 1.4 });
  });
  txt(s, "EXW ROYE, FRANCE", PW - M - 3, 4.72, 3, 0.26, { size: 9.5, bold: true, color: MUTED, cs: 1.6, align: "right" });

  pic(s, A("assets/mascot_panel.jpg"), 10.62, 1.62, 1.95, 2.59);

  const lw = COL + 0.9, lh = lw / 4.359;
  pic(s, A("assets/lineup_dark.png"), (PW - lw) / 2, PH - lh - 0.06, lw, lh);
}

// ═════════════════════════════════════════════════════════════════════
// 02 · THE MAKER
// ═════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  page(s);
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
function volumeSlide(page_no, kicker, tier, big, small, lead) {
  const s = pres.addSlide();
  page(s);

  const pallets = Object.values(SPLIT[tier]).reduce((a, b) => a + b, 0);

  rule(s, M, 0.6, COL);
  txt(s, kicker, M, 0.74, 7, 0.24, { size: 8, bold: true, color: AMBER, cs: 2.8 });

  pic(s, A("assets/flames.png"), M, 0.98, 0.52, 0.64);
  txt(s, big + " " + small, M + 0.66, 1.0, 5.1, 0.74, { size: 40, bold: true, color: CREAM, cs: -1 });
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

  folio(s, page_no, { noLabel: true });
}

volumeSlide(3, "ОБСЯГ ПОСТАЧАННЯ · 01", "p6",  "6",  "ПАЛЕТ",
  "Пробна партія — по одній палеті на кожен смак.");
volumeSlide(4, "ОБСЯГ ПОСТАЧАННЯ · 02", "p12", "12", "ПАЛЕТ",
  "Робочий обсяг з акцентом на вершкових смаках — Gouda та Carbonara.");
volumeSlide(5, "ОБСЯГ ПОСТАЧАННЯ · 03", "p33", "33", "ПАЛЕТИ",
  "Повне авто (Full Truck) — найнижча собівартість пачки.");

// ═════════════════════════════════════════════════════════════════════
// 06 · WHERE THE PRODUCT IS SOLD
// ═════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  page(s);
  head(s, "РИНОК", "ДЕ ПРЕДСТАВЛЕНА ПРОДУКЦІЯ");
  txt(s, "роздрібні мережі та дистриб'ютори, з якими працює SIAS",
    M, 1.82, 9.6, 0.28, { size: 11, color: MUTED });

  const ww = COL, wh = ww / 3.04;
  rect(s, M - 0.16, 2.34, ww + 0.32, wh + 0.32, { fill: PAPER, round: 0.1 });
  pic(s, A("assets/logo_wall.png"), M, 2.5, ww, wh);

  folio(s, 6);
}

// ═════════════════════════════════════════════════════════════════════
// 07 · MADE IN THE EU, BUILT THE KOREAN WAY
// ═════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  page(s);
  head(s, "ВИРОБНИЦТВО", "ЗАВОД У ЄС — ТЕХНОЛОГІЯ З КОРЕЇ", { size: 30 });

  txt(s, "Змінилася географія, а не технологія: рецептури, обладнання та стандарти перенесені з корейських заводів групи.",
    M, 1.86, 10.4, 0.28, { size: 11.5, color: MUTED });

  rule(s, M, 2.26, COL);

  // the 150 m line
  pic(s, A("assets/line150.png"), M - 0.1, 2.42, 7.4, 2.34);
  txt(s, "ЛІНІЯ 150 МЕТРІВ", M, 4.74, 4, 0.24, { size: 9, bold: true, color: AMBER, cs: 1.8 });
  txt(s, "Три послідовні процеси по 50 м: заміс тіста → формування пружної\nлокшини → сушіння та пакування.",
    M, 4.98, 6.6, 0.54, { size: 10.5, color: MUTED, ls: 1.3 });

  const pts = [
    ["KNOW-HOW З КОРЕЇ", "Рецептури, обладнання та виробничі процеси перенесені з корейських майданчиків SIAS — у групі їх сім."],
    ["ТОЙ САМИЙ КОНТРОЛЬ", "Тести та інспекції на кожному етапі за корейським стандартом якості."],
    ["КОРЕЙСЬКІ СОУСИ ТА СМАКИ", "Власні рецептури соусів і приправ — смак не адаптований під «європейську» версію."],
  ];
  let py = 2.44;
  pts.forEach((p) => {
    rect(s, 8.5, py + 0.06, 0.09, 0.09, { fill: CORAL });
    txt(s, p[0], 8.72, py, 4.1, 0.24, { size: 10, bold: true, color: CREAM, cs: 0.6 });
    txt(s, p[1], 8.72, py + 0.26, 4.1, 0.62, { size: 9.5, color: MUTED, ls: 1.3 });
    py += 1.02;
  });

  // the commercial consequence of making it inside the EU
  rect(s, 8.5, 5.62, 4.05, 1.16, { fill: "20191A", round: 0.08 });
  rect(s, 8.5, 5.62, 0.06, 1.16, { fill: CORAL });
  txt(s, "ВИРОБНИЦТВО В ЄС", 8.74, 5.78, 3.6, 0.22, { size: 8.5, bold: true, color: AMBER, cs: 1.6 });
  txt(s, "Без морського фрахту, коротке плече доставки та сертифікат EUR.1 — 0% ввізного мита в Україну.",
    8.74, 6.04, 3.66, 0.6, { size: 9.5, color: MUTED, ls: 1.3 });

  pic(s, A("assets/mascot_panel2.jpg"), M, 5.72, 1.0, 1.32);

  txt(s, "Made in France · savoir-faire coréen", M + 1.2, 5.8, 5.2, 0.26, { size: 10, italic: true, color: FAINT });
  frFlag(s, M + 1.2, 6.24, 0.44, 0.29);
  txt(s, "+", M + 1.74, 6.22, 0.3, 0.3, { size: 12, color: MUTED });
  rect(s, M + 2.06, 6.24, 0.44, 0.29, { fill: "FFFFFF" });
  txt(s, "KR", M + 2.06, 6.24, 0.44, 0.29, { size: 9, bold: true, color: "C60C30", align: "center", valign: "middle" });
  txt(s, "→", M + 2.64, 6.22, 0.3, 0.3, { size: 12, color: MUTED });
  uaFlag(s, M + 2.98, 6.24, 0.44, 0.29);

  folio(s, 7);
}

// ═════════════════════════════════════════════════════════════════════
pres.writeFile({ fileName: A("SIAS_France_Presentation.pptx") })
  .then(() => console.log("Deck written."))
  .catch((e) => { console.error(e); process.exit(1); });
