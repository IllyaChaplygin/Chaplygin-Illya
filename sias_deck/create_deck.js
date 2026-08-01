const pptxgen = require("pptxgenjs");
const path = require("path");

const A = (p) => path.join(__dirname, p);

// ─────────────────────────────────────────────────────────────────────
// EDITORIAL PALETTE — warm paper, one loud red, muted gold hairlines
// ─────────────────────────────────────────────────────────────────────
const INK = "14100E";
const INK_SOFT = "3A322C";
const PAPER = "F7F2EA";
const CARD = "FFFFFF";
const RED = "D80F19";
const RED_DEEP = "8E0B12";
const GOLD = "A8874E";
const GOLD_LT = "E6D6B8";
const GREY = "6E6862";
const HAIR = "DCD3C6";

const FONT = "Calibri";
const FONT_H = "Cambria";
const FONT_KR = "Malgun Gothic";

const PW = 13.333, PH = 7.5;
const M = 0.6;                       // page margin
const COLW = PW - M * 2;             // 12.133

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";

// ─────────────────────────────────────────────────────────────────────
// PRODUCT DATA
// ─────────────────────────────────────────────────────────────────────
const SKUS = [
  { key: "gouda",     name: "Гауда",      kr: "고다치즈",   weight: "123 г",   accent: "F07C1E",
    note: "Bibimmyeon · вершковий сир гауда", img: A("packs/gouda.png") },
  { key: "carbonara", name: "Карбонара",  kr: "까르보나라", weight: "131 г",   accent: "D96A88",
    note: "Ramyeon · вершкова, гостра HOT",   img: A("packs/carbonara.png") },
  { key: "vegetable", name: "Овочева",    kr: "야채맛",     weight: "113 г",   accent: "4E9E2A",
    note: "Ramyeon · без гостроти",           img: A("packs/vegetable.png") },
  { key: "spicy",     name: "Гостра",     kr: "매운맛",     weight: "112,5 г", accent: "C8101C",
    note: "Ramyeon · гострота MEDIUM",        img: A("packs/spicy.png") },
  { key: "beef",      name: "Яловичина",  kr: "소고기맛",   weight: "113 г",   accent: "A9541A",
    note: "Ramyeon · гострота MILD",          img: A("packs/beef.png") },
  { key: "chicken",   name: "Курка",      kr: "치킨맛",     weight: "113 г",   accent: "D9AA05",
    note: "Ramyeon · без гостроти",           img: A("packs/chicken.png") },
];

// Landed cost per pack, from SelfCost.xlsx (duty 0% — EUR.1)
const COST = {
  gouda:     { p6: [1.31, 67.00, 10.64], p12: [1.11, 56.88, 9.03], p33: [1.02, 52.13, 8.28] },
  carbonara: { p6: [1.31, 67.00,  9.99], p12: [1.11, 56.88, 8.48], p33: [1.02, 52.13, 7.77] },
  vegetable: { p6: [1.13, 58.06, 10.04], p12: [0.96, 49.29, 8.52], p33: [0.88, 45.18, 7.81] },
  spicy:     { p6: [1.13, 58.06, 10.08], p12: [0.96, 49.29, 8.56], p33: [0.88, 45.18, 7.84] },
  beef:      { p6: [1.13, 58.06, 10.04], p12: [0.96, 49.29, 8.52], p33: [0.88, 45.18, 7.81] },
  chicken:   { p6: [1.13, 58.06, 10.04], p12: [0.96, 49.29, 8.52], p33: [0.88, 45.18, 7.81] },
};

// Pallet split per SKU in each shipment scenario
const SPLIT = {
  p6:  { gouda: 1, carbonara: 1, vegetable: 1, spicy: 1, beef: 1, chicken: 1 },
  p12: { gouda: 3, carbonara: 3, vegetable: 1, spicy: 2, beef: 2, chicken: 1 },
  p33: { gouda: 7, carbonara: 7, vegetable: 4, spicy: 5, beef: 5, chicken: 5 },
};

const PACKS_PER_PALLET = 1600; // 20 шт/короб × 80 короб/палета

// ─────────────────────────────────────────────────────────────────────
// DRAWING HELPERS
// ─────────────────────────────────────────────────────────────────────
function rect(s, x, y, w, h, o = {}) {
  const sh = { x, y, w, h, fill: o.fill ? { color: o.fill } : { type: "none" }, line: o.line || { type: "none" } };
  if (o.fill && o.transparency !== undefined) sh.fill.transparency = o.transparency;
  s.addShape("rect", sh);
}

function roundRect(s, x, y, w, h, o = {}) {
  const sh = {
    x, y, w, h,
    rectRadius: o.radius !== undefined ? o.radius : 0.08,
    fill: o.fill ? { color: o.fill } : { type: "none" },
    line: o.line || { type: "none" },
  };
  if (o.fill && o.transparency !== undefined) sh.fill.transparency = o.transparency;
  if (o.shadow) sh.shadow = o.shadow;
  s.addShape("roundRect", sh);
}

function txt(s, t, x, y, w, h, o = {}) {
  s.addText(t, {
    x, y, w, h,
    fontFace: o.font || FONT,
    fontSize: o.size || 12,
    color: o.color || INK,
    bold: !!o.bold,
    italic: !!o.italic,
    align: o.align || "left",
    valign: o.valign || "top",
    lineSpacingMultiple: o.ls || 1.06,
    charSpacing: o.cs,
    margin: 0,
  });
}

function pic(s, p, x, y, w, h, o = {}) {
  const im = { path: p, x, y, w, h, sizing: { type: o.fit || "contain", w, h } };
  if (o.shadow) im.shadow = o.shadow;
  s.addImage(im);
}

const packShadow = () => ({ type: "outer", color: "241A14", opacity: 0.3, blur: 14, offset: 6, angle: 90 });
const cardShadow = () => ({ type: "outer", color: "2B211A", opacity: 0.13, blur: 12, offset: 3, angle: 90 });

function hairline(s, x, y, w, color) {
  rect(s, x, y, w, 0.011, { fill: color || HAIR });
}

function euro(n) { return n.toFixed(2).replace(".", ",") + " €"; }
function uah(n) { return n.toFixed(2).replace(".", ",") + " ₴"; }
function thin(n) { return n.toLocaleString("uk-UA").replace(/ /g, " "); }

function palletWord(n) {
  if (n === 1) return "палета";
  if (n >= 2 && n <= 4) return "палети";
  return "палет";
}

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

function runner(s, n, light) {
  txt(s, "CHOI'S RAMYEON  ·  SIAS FRANCE  ·  2026", M, PH - 0.4, 6, 0.24,
    { size: 8, color: light ? "9C918A" : GREY, cs: 1.4 });
  txt(s, String(n).padStart(2, "0"), PW - M - 0.6, PH - 0.42, 0.6, 0.26,
    { size: 10, color: light ? "9C918A" : GREY, align: "right", font: FONT_H, bold: true });
}

// Chapter head: small red tick, kicker, big serif title
function head(s, kicker, title, sub) {
  rect(s, M, 0.5, 0.19, 0.19, { fill: RED });
  txt(s, kicker, M + 0.31, 0.45, 8, 0.28, { size: 10, bold: true, color: GREY, cs: 1.6 });
  txt(s, title, M, 0.82, 8.8, 0.58, { font: FONT_H, bold: true, size: 32, color: INK });
  if (sub) txt(s, sub, M + 0.03, 1.42, 8.8, 0.3, { size: 12.5, italic: true, color: GOLD, font: FONT_H });
}

// ═════════════════════════════════════════════════════════════════════
// 01 — COVER
// ═════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: INK };
  pic(s, A("assets/cover_bg.jpg"), 0, 0, PW, PH, { fit: "cover" });

  txt(s, "SIAS FRANCE", M, 0.5, 5, 0.3, { size: 11.5, bold: true, color: GOLD_LT, cs: 3 });
  hairline(s, M, 0.92, 1.5, "8A6F3E");
  txt(s, "ПРАЙС-КАТАЛОГ ДЛЯ КОМАНДИ  ·  2026", M, 1.06, 6, 0.28, { size: 10, color: "C4B8AA", cs: 1.8 });

  frFlag(s, PW - M - 1.25, 0.52, 0.5, 0.33);
  uaFlag(s, PW - M - 0.6, 0.52, 0.5, 0.33);

  txt(s, "라면", M, 1.5, 3, 0.42, { font: FONT_KR, bold: true, size: 23, color: GOLD_LT, cs: 4 });
  txt(s, "CHOI'S RAMYEON", M - 0.06, 1.86, 12.4, 1.0, { font: FONT_H, bold: true, size: 64, color: "FFFFFF" });

  hairline(s, M, 2.98, 0.72, GOLD);
  txt(s, "6 СМАКІВ  ·  ФОРМАТ BAG  ·  112,5–131 Г  ·  20 ШТ/КОРОБ", M, 3.12, 9, 0.26,
    { size: 10.5, bold: true, color: "E4DACB", cs: 1 });
  txt(s, "Собівартість пачки для трьох обсягів постачання — 6, 12 та 33 палети (Full Truck).\nВиробництво SIAS у Франції, сертифікат EUR.1: 0% ввізного мита в Україну.",
    M, 3.46, 9.6, 0.62, { font: FONT_H, italic: true, size: 14.5, color: "EFE5D6", ls: 1.28 });

  // the product itself carries the cover
  pic(s, A("assets/cover_lineup.png"), 0.07, 4.16, 13.2, 3.34, { fit: "contain" });
}


// ═════════════════════════════════════════════════════════════════════
// 02 — THE MAKER (full-bleed hero photo + overlapping title block)
// ═════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: PAPER };

  // full-bleed hero
  pic(s, A("assets/factory_hero.jpg"), 0, 0, PW, 4.0, { fit: "cover" });

  // title block straddling the photo edge
  roundRect(s, M, 3.02, 4.75, 1.5, { fill: RED, radius: 0.1, shadow: cardShadow() });
  txt(s, "ВИРОБНИК · 01", M + 0.3, 3.2, 4.2, 0.24, { size: 9.5, bold: true, color: "F7C9CC", cs: 1.6 });
  txt(s, "SIAS France", M + 0.28, 3.48, 4.2, 0.6, { font: FONT_H, bold: true, size: 30, color: "FFFFFF" });
  txt(s, "Roye · Hauts-de-France", M + 0.3, 4.08, 4.2, 0.28, { size: 10.5, italic: true, color: "FBD9DC" });

  // korean accent on the cream, beneath the title block
  txt(s, "최씨라면", M + 0.02, 4.76, 3.2, 0.42, { font: FONT_KR, bold: true, size: 21, color: GOLD, cs: 2 });

  txt(s,
    "Європейська виробнича площадка корейської групи SIAS: 16 000 м², власна лінія сушеної " +
    "локшини, сертифікати IFS Food і BRC Food. Choi's Ramyeon виробляють тут за корейськими " +
    "рецептурами, адаптованими під європейського споживача.",
    5.65, 4.2, 7.1, 1.3, { size: 12, color: INK_SOFT, ls: 1.32 });

  const facts = [
    ["EXW Roye", "умови відвантаження"],
    ["16 000 м²", "виробнича площадка"],
    ["IFS · BRC", "харчова безпека"],
    ["0 %", "мито в Україну за EUR.1"],
  ];
  const fw = (COLW - 3 * 0.18) / 4;
  facts.forEach((f, i) => {
    const x = M + i * (fw + 0.18);
    const hot = i === 3;
    roundRect(s, x, 5.62, fw, 1.2, { fill: hot ? RED : CARD, radius: 0.09, shadow: cardShadow() });
    txt(s, f[0], x + 0.2, 5.78, fw - 0.4, 0.42, { font: FONT_H, bold: true, size: hot ? 24 : 20, color: hot ? "FFFFFF" : RED });
    txt(s, f[1], x + 0.2, 6.26, fw - 0.4, 0.4, { size: 9.5, color: hot ? "FBD9DC" : GREY, ls: 1.14 });
  });

  runner(s, 2);
}
// ═════════════════════════════════════════════════════════════════════
// 04-06 — ONE SLIDE PER SHIPMENT SCENARIO, ALL SIX SKUs
// ═════════════════════════════════════════════════════════════════════
function scenarioSlide(pageNo, chapter, tier, title, lead) {
  const s = pres.addSlide();
  s.background = { color: PAPER };

  const totalPallets = Object.values(SPLIT[tier]).reduce((a, b) => a + b, 0);
  const totalPacks = totalPallets * PACKS_PER_PALLET;

  rect(s, M, 0.5, 0.19, 0.19, { fill: RED });
  txt(s, chapter, M + 0.31, 0.45, 7, 0.28, { size: 10, bold: true, color: GREY, cs: 1.6 });
  txt(s, title, M, 0.8, 7.4, 0.6, { font: FONT_H, bold: true, size: 34, color: INK });
  txt(s, lead, M + 0.03, 1.44, 7.4, 0.3, { size: 12, italic: true, color: GOLD, font: FONT_H });

  // scenario summary, top-right
  const chips = [
    [thin(totalPallets), "палет у постачанні"],
    [thin(totalPacks), "пачок разом"],
    ["0 %", "мито · EUR.1"],
  ];
  const cw2 = 1.62, cg = 0.14;
  chips.forEach((c, i) => {
    const x = PW - M - (3 - i) * (cw2 + cg) + cg;
    const hot = i === 2;
    roundRect(s, x, 0.5, cw2, 1.1, { fill: hot ? RED : CARD, radius: 0.09, shadow: cardShadow() });
    txt(s, c[0], x + 0.12, 0.62, cw2 - 0.24, 0.4, { font: FONT_H, bold: true, size: 19, color: hot ? "FFFFFF" : INK, align: "center" });
    txt(s, c[1], x + 0.1, 1.06, cw2 - 0.2, 0.42, { size: 8.2, color: hot ? "FBD9DC" : GREY, align: "center", ls: 1.12 });
  });

  hairline(s, M, 1.86, COLW, GOLD);

  // 3 × 2 grid of product/price cards
  const gx = 0.28, gy = 0.20;
  const cw = (COLW - 2 * gx) / 3;      // 3.858
  const ch = 2.34;
  const y0 = 1.96;

  SKUS.forEach((sk, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = M + col * (cw + gx);
    const y = y0 + row * (ch + gy);
    const [eur, grn, kg] = COST[sk.key][tier];
    const pal = SPLIT[tier][sk.key];

    roundRect(s, x, y, cw, ch, { fill: CARD, radius: 0.1, shadow: cardShadow() });

    // packshot, left
    const imgW = 1.5;
    pic(s, sk.img, x + 0.12, y + 0.14, imgW, ch - 0.28, { shadow: packShadow() });

    // info column, right
    const ix = x + imgW + 0.24;
    const iw = cw - imgW - 0.24 - 0.16;   // 1.958

    txt(s, sk.name, ix, y + 0.13, iw, 0.3, { font: FONT_H, bold: true, size: 15, color: INK });
    txt(s, sk.kr, ix, y + 0.44, iw, 0.2, { font: FONT_KR, bold: true, size: 9.5, color: sk.accent });

    // weight left, pallet chip right — same row, no collision with the name
    txt(s, sk.weight, ix, y + 0.68, 0.95, 0.26, { size: 12.5, bold: true, color: RED });
    roundRect(s, ix + iw - 0.8, y + 0.69, 0.8, 0.25, { fill: PAPER, radius: 0.12 });
    txt(s, pal + " " + palletWord(pal), ix + iw - 0.8, y + 0.69, 0.8, 0.25,
      { align: "center", valign: "middle", size: 7.5, bold: true, color: INK_SOFT });

    txt(s, "20 шт/короб · 80 короб/палета", ix, y + 0.96, iw, 0.2, { size: 8.2, color: GREY });
    txt(s, thin(pal * PACKS_PER_PALLET) + " пачок у постачанні", ix, y + 1.15, iw, 0.2, { size: 8.2, color: GREY });

    // price block
    const by = y + 1.38, bh = 0.82;
    roundRect(s, ix, by, iw, bh, { fill: RED_DEEP, radius: 0.07 });
    txt(s, "СОБІВАРТІСТЬ / ПАЧКА", ix + 0.13, by + 0.07, iw - 0.26, 0.18, { size: 7, bold: true, color: "F0BFC3", cs: 0.7 });
    txt(s, euro(eur), ix + 0.13, by + 0.24, iw - 0.26, 0.36, { font: FONT_H, bold: true, size: 22, color: "FFFFFF" });
    txt(s, uah(grn), ix + 0.13, by + 0.6, iw * 0.5, 0.2, { size: 9.5, color: "F6DBDD" });
    txt(s, euro(kg) + "/кг", ix + iw * 0.48, by + 0.6, iw * 0.5 - 0.13, 0.2, { size: 9.5, bold: true, color: "F0BFC3", align: "right" });
  });

  txt(s,
    "Собівартість на 1 пачку: EXW Roye + логістика + брокер + митне оформлення + ПДВ-передфінансування, ставка ввізного мита 0% (EUR.1). Курс: 1 € ≈ 51,2 ₴.",
    M, 6.92, COLW, 0.24, { size: 8.2, italic: true, color: GREY });

  runner(s, pageNo);
}

scenarioSlide(3, "ВАРІАНТ ПОСТАЧАННЯ · 02", "p6",
  "6 палет", "пробна партія · по 1 палеті на кожен смак");

scenarioSlide(4, "ВАРІАНТ ПОСТАЧАННЯ · 03", "p12",
  "12 палет", "середній обсяг · акцент на вершкових смаках");

scenarioSlide(5, "ВАРІАНТ ПОСТАЧАННЯ · 04", "p33",
  "33 палети · Full Truck", "повне авто · найнижча собівартість пачки");


// ═════════════════════════════════════════════════════════════════════
// 06 — CLIENTS & RETAIL NETWORKS
// ═════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: PAPER };
  head(s, "ПАРТНЕРИ · 05", "Нам довіряють", "рітейл, дистрибуція, HoReCa та meal-kit по всій Європі");

  hairline(s, M, 1.92, COLW, GOLD);

  const LOGOS = [
    "carrefour", "leclerc", "systemeu", "aldi", "eurospin", "bofrost",
    "sysco", "hellofresh", "wismettac", "geia", "bibars", "dipsa",
    "forezia", "ttfoods", "gotiger", "foodex", "senko",
  ];
  const cols = 6, gap = 0.16;
  const cw = (COLW - (cols - 1) * gap) / cols;   // 1.889
  const chh = 1.16, y0 = 2.16;

  LOGOS.forEach((name, i) => {
    const c = i % cols, r = Math.floor(i / cols);
    const x = M + c * (cw + gap);
    const y = y0 + r * (chh + gap);
    roundRect(s, x, y, cw, chh, { fill: CARD, radius: 0.09, shadow: cardShadow() });
    pic(s, A(`logos/${name}.png`), x + 0.2, y + 0.19, cw - 0.4, chh - 0.38);
  });

  // category strip under the grid
  const cats = [
    ["Рітейл-мережі", "Carrefour · E.Leclerc · Système U · ALDI · EuroSpin"],
    ["Дистрибуція та HoReCa", "Sysco · Wismettac · Dipsa · Geia · Bofrost"],
    ["Meal-kit та спеціалізовані", "HelloFresh · Forezia · T&T Foods · BiBars"],
  ];
  const catW = (COLW - 2 * 0.3) / 3;
  cats.forEach((c, i) => {
    const x = M + i * (catW + 0.3);
    rect(s, x, 6.28, 0.12, 0.12, { fill: RED });
    txt(s, c[0], x + 0.24, 6.22, catW - 0.24, 0.26, { size: 11, bold: true, color: INK });
    txt(s, c[1], x + 0.24, 6.48, catW - 0.24, 0.4, { size: 9.2, color: GREY, ls: 1.18 });
  });

  runner(s, 6);
}

// ═════════════════════════════════════════════════════════════════════
// 07 — COMPANY & PRODUCTION (short closing)
// ═════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: PAPER };
  head(s, "ПРО ВИРОБНИКА · 06", "Компанія та виробництво", "коротко про групу SIAS і завод у Франції");

  hairline(s, M, 1.92, COLW, GOLD);

  // left — key figures
  const facts = [
    ["1997", "рік заснування\nгрупи SIAS (Корея)"],
    ["9", "заводів: 7 у Кореї\nта 2 у Франції"],
    ["$180 млн", "річний оборот\nгрупи"],
    ["2 100+", "рецептур і понад\n1 000 клієнтів"],
  ];
  const fw = (6.3 - 0.18) / 2, fh = 1.32;
  facts.forEach((f, i) => {
    const x = M + (i % 2) * (fw + 0.18);
    const y = 2.2 + Math.floor(i / 2) * (fh + 0.18);
    roundRect(s, x, y, fw, fh, { fill: CARD, radius: 0.09, shadow: cardShadow() });
    txt(s, f[0], x + 0.18, y + 0.16, fw - 0.36, 0.44, { font: FONT_H, bold: true, size: 22, color: RED });
    txt(s, f[1], x + 0.18, y + 0.66, fw - 0.36, 0.55, { size: 9.5, color: GREY, ls: 1.16 });
  });

  txt(s,
    "Choi's Ramyeon виробляють на власному заводі SIAS у Roye (Франція) — 16 000 м² площадки, " +
    "лінія сушеної локшини завдовжки 150 метрів, від замісу тіста до пакування. " +
    "Сертифікати IFS Food і BRC Food; походження ЄС дає сертифікат EUR.1 — 0% ввізного мита в Україну.",
    M, 5.2, 6.3, 1.4, { size: 11.5, color: INK_SOFT, ls: 1.3 });

  // right — production line + certification
  const rx = 7.2, rw = PW - M - rx;
  roundRect(s, rx, 2.2, rw, 2.5, { fill: CARD, radius: 0.1, shadow: cardShadow() });
  pic(s, A("assets/production_line.jpg"), rx + 0.16, 2.36, rw - 0.32, 1.85);
  txt(s, "Виробнича лінія сушеної локшини · 150 метрів", rx + 0.16, 4.3, rw - 0.32, 0.3,
    { size: 9.5, italic: true, color: GREY, align: "center" });

  roundRect(s, rx, 4.88, rw, 1.32, { fill: CARD, radius: 0.1, shadow: cardShadow() });
  txt(s, "СЕРТИФІКАТИ ТА ПОХОДЖЕННЯ", rx + 0.22, 5.02, rw - 0.44, 0.24, { size: 9.5, bold: true, color: GOLD, cs: 1.4 });
  pic(s, A("assets/cert_ifs_brc.jpg"), rx + 0.22, 5.34, 1.6, 0.46);
  pic(s, A("assets/cert_france_relance.jpg"), rx + 2.0, 5.28, 0.58, 0.58);
  frFlag(s, rx + 2.85, 5.44, 0.46, 0.3);
  txt(s, "→", rx + 3.4, 5.4, 0.3, 0.3, { size: 12, color: GREY });
  uaFlag(s, rx + 3.75, 5.44, 0.46, 0.3);
  txt(s, "EUR.1 · 0%", rx + 4.35, 5.42, 1.3, 0.3, { size: 11, bold: true, color: RED });

  roundRect(s, rx, 6.36, rw, 0.62, { fill: INK, radius: 0.09 });
  txt(s, "EXW ROYE, FRANCE  ·  6 / 12 / 33 ПАЛЕТИ  ·  0% МИТО", rx, 6.36, rw, 0.62,
    { align: "center", valign: "middle", size: 11, bold: true, color: "FFFFFF", cs: 1 });

  runner(s, 7);
}

// ═════════════════════════════════════════════════════════════════════
pres.writeFile({ fileName: A("SIAS_France_Presentation.pptx") })
  .then(() => console.log("Deck written."))
  .catch((e) => { console.error(e); process.exit(1); });
