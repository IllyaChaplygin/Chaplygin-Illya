const pptxgen = require("pptxgenjs");
const path = require("path");

const A = (p) => path.join(__dirname, p);

// ---------- PALETTE (matches internal "snacks" catalog style: cream + one accent) ----------
const RED = "C81420";        // SIAS brand red, editorial tone
const RED_DEEP = "7A0A10";   // deep red for highlighted price box
const INK = "1C1A18";
const CREAM = "FAF7F2";      // catalog cream background
const CARD = "FFFFFF";
const GOLD = "B8892B";
const GREY = "746F69";
const LGREY = "E7E1D9";
const HAIRLINE = "DED6CB";

const FONT = "Calibri";
const FONT_H = "Cambria";
const FONT_KR = "Malgun Gothic";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
const PW = 13.333, PH = 7.5;

// ---------- flavor accent colors sampled from real packaging ----------
const SKUS = [
  { key: "gouda", name: "Гауда", sub: "сир", kr: "고다치즈", weight: "123 г", img: A("products/gouda_123g.png"), accent: "F99023", desc: "Вершковий смак сиру гауда з ноткою чеддеру — м'який, без гостроти." },
  { key: "carbonara", name: "Карбонара", sub: "пікантна", kr: "까르보나라", weight: "131 г", img: A("products/carbonara_131g.png"), accent: "E0879D", desc: "Вершковий соус карбонара з чорним перцем — гострота рівня HOT." },
  { key: "vegetable", name: "Овочева", sub: "без гостроти", kr: "야채맛", weight: "113 г", img: A("products/vegetable_113g.png"), accent: "5DB42F", desc: "Овочевий бульйон із сушеними овочами — м'який смак, без перцю." },
  { key: "spicy", name: "Гостра", sub: "spicy", kr: "매운맛", weight: "112,5 г", img: A("products/spicy_112.5g.png"), accent: "D3010E", desc: "Гострий бульйон з перцем чилі — рівень гостроти MEDIUM." },
  { key: "beef", name: "Яловичина", sub: "mild", kr: "소고기맛", weight: "113 г", img: A("products/beef_113g.png"), accent: "C85D19", desc: "Насичений яловичий бульйон — м'який смак, рівень гостроти MILD." },
  { key: "chicken", name: "Курка", sub: "mild", kr: "치킨맛", weight: "113 г", img: A("products/chicken_113g.png"), accent: "F4C803", desc: "Ніжний курячий бульйон із зеленню — без гостроти, для всієї родини." },
];

const PRICING = {
  gouda:     { p33: { eur: 1.02, kg: 8.28, uah: 52.13 }, p12: { eur: 1.11, kg: 9.03, uah: 56.88 }, p6: { eur: 1.31, kg: 10.64, uah: 67.00 } },
  carbonara: { p33: { eur: 1.02, kg: 7.77, uah: 52.13 }, p12: { eur: 1.11, kg: 8.48, uah: 56.88 }, p6: { eur: 1.31, kg: 9.99,  uah: 67.00 } },
  vegetable: { p33: { eur: 0.88, kg: 7.81, uah: 45.18 }, p12: { eur: 0.96, kg: 8.52, uah: 49.29 }, p6: { eur: 1.13, kg: 10.04, uah: 58.06 } },
  spicy:     { p33: { eur: 0.88, kg: 7.84, uah: 45.18 }, p12: { eur: 0.96, kg: 8.56, uah: 49.29 }, p6: { eur: 1.13, kg: 10.08, uah: 58.06 } },
  beef:      { p33: { eur: 0.88, kg: 7.81, uah: 45.18 }, p12: { eur: 0.96, kg: 8.52, uah: 49.29 }, p6: { eur: 1.13, kg: 10.04, uah: 58.06 } },
  chicken:   { p33: { eur: 0.88, kg: 7.81, uah: 45.18 }, p12: { eur: 0.96, kg: 8.52, uah: 49.29 }, p6: { eur: 1.13, kg: 10.04, uah: 58.06 } },
};

// =====================================================================
// HELPERS
// =====================================================================
function bg(slide, color) { slide.background = { color }; }

function rect(slide, x, y, w, h, opts = {}) {
  const o = { x, y, w, h, fill: opts.fill ? { color: opts.fill } : { type: "none" }, line: opts.line || { type: "none" } };
  if (opts.transparency !== undefined) o.fill.transparency = opts.transparency;
  slide.addShape("rect", o);
}

function roundRect(slide, x, y, w, h, opts = {}) {
  const o = {
    x, y, w, h,
    rectRadius: opts.radius !== undefined ? opts.radius : 0.09,
    fill: opts.fill ? { color: opts.fill } : { type: "none" },
    line: opts.line || { type: "none" },
  };
  if (opts.shadow) o.shadow = opts.shadow;
  slide.addShape("roundRect", o);
}

function txt(slide, text, x, y, w, h, opts = {}) {
  const o = {
    x, y, w, h,
    fontFace: opts.font || FONT,
    fontSize: opts.size || 14,
    color: opts.color || INK,
    bold: !!opts.bold,
    italic: !!opts.italic,
    align: opts.align || "left",
    valign: opts.valign || "top",
    lineSpacingMultiple: opts.lineSpacing || 1.08,
    charSpacing: opts.charSpacing,
    margin: opts.margin !== undefined ? opts.margin : 0,
  };
  if (opts.transparency !== undefined) o.transparency = opts.transparency;
  slide.addText(text, o);
}

function pic(slide, p, x, y, w, h, opts = {}) {
  const o = { x, y, w, h, sizing: { type: opts.sizingType || "cover", w, h } };
  if (opts.shadow) o.shadow = opts.shadow;
  if (opts.rotate) o.rotate = opts.rotate;
  if (opts.transparency !== undefined) o.transparency = opts.transparency;
  slide.addImage(Object.assign({ path: p }, o));
}

function softShadow(opts = {}) {
  return { type: "outer", color: "3A2E28", opacity: opts.opacity || 0.22, blur: opts.blur || 10, offset: opts.offset || 3, angle: 90 };
}

function euro(n) { return n.toFixed(2).replace(".", ",") + " €"; }
function uah(n) { return n.toFixed(2).replace(".", ",") + " ₴"; }

function frFlag(slide, x, y, w, h) {
  const seg = w / 3;
  rect(slide, x, y, seg, h, { fill: "002654" });
  rect(slide, x + seg, y, seg, h, { fill: "FFFFFF" });
  rect(slide, x + 2 * seg, y, seg, h, { fill: "ED2939" });
}

function uaFlag(slide, x, y, w, h) {
  rect(slide, x, y, w, h / 2, { fill: "0057B7" });
  rect(slide, x, y + h / 2, w, h / 2, { fill: "FFD700" });
}

function pageNum(slide, n, light) {
  txt(slide, String(n).padStart(2, "0"), PW - 0.85, PH - 0.42, 0.5, 0.3, { size: 9.5, color: light ? "FFFFFF" : GREY, align: "right" });
  txt(slide, "CHOI'S RAMYEON · SIAS FRANCE", 0.55, PH - 0.42, 4.5, 0.3, { size: 8.5, color: light ? "FFFFFF" : GREY, charSpacing: 1 });
}

function kicker(slide, num, label, opts = {}) {
  rect(slide, 0.55, 0.56, 0.22, 0.22, { fill: opts.color || RED });
  txt(slide, num + "  ·  " + label, 0.88, 0.5, 9, 0.35, { size: 11, bold: true, color: opts.textColor || GREY, charSpacing: 1.4 });
}

// =====================================================================
// SLIDE 1 — MAGAZINE COVER
// =====================================================================
{
  const s = pres.addSlide();
  bg(s, INK);
  pic(s, A("assets/hero_bowl.jpg"), 0, 0, PW, PH, {});
  rect(s, 0, 0, PW, PH, { fill: "1C0A08", transparency: 42 });
  rect(s, 0, 4.35, PW, 3.15, { fill: "0D0605", transparency: 22 });

  // Korean watermark accent
  txt(s, "라면", 8.55, 0.25, 5.2, 3.6, { font: FONT_KR, bold: true, size: 205, color: "FFFFFF", transparency: 82, align: "right" });

  // cover kicker lines (magazine cover-lines)
  txt(s, "SIAS FRANCE · 2026", 0.65, 0.5, 5, 0.4, { size: 12.5, bold: true, color: "F2D9A8", charSpacing: 2.2 });
  txt(s, "6 СМАКІВ", 0.65, 0.92, 3, 0.35, { size: 11, color: "FFFFFF", charSpacing: 1.5 });
  txt(s, "EXW ROYE → УКРАЇНА", PW - 4.15, 0.5, 3.6, 0.35, { size: 11, color: "FFFFFF", align: "right", charSpacing: 1.2 });
  txt(s, "EUR.1 · 0% МИТО", PW - 4.15, 0.86, 3.6, 0.35, { size: 12.5, bold: true, color: "F2D9A8", align: "right", charSpacing: 1.4 });

  txt(s, "CHOI'S", 0.6, 3.55, 9.5, 1.15, { font: FONT_H, bold: true, size: 76, color: "FFFFFF", charSpacing: 1 });
  txt(s, "RAMYEON", 0.6, 4.55, 11.5, 1.25, { font: FONT_H, bold: true, size: 76, color: "FFFFFF", charSpacing: 1 });

  rect(s, 0.65, 5.85, 0.5, 0.045, { fill: GOLD });
  txt(s, "Корейська локша швидкого приготування виробництва Франції — прайс-каталог для команди", 0.62, 6.0, 10.6, 0.5,
    { italic: true, size: 14.5, color: "EFE6D8", font: FONT_H });

  txt(s, "Bag · 6 SKU · вершкові, овочеві, гострі та м'ясні смаки", 0.65, 6.62, 9, 0.4, { size: 11.5, color: "CFC3B2" });
}

// =====================================================================
// SLIDE 2 — SIAS FRANCE PROFILE (single accent panel, reference-style)
// =====================================================================
{
  const s = pres.addSlide();
  bg(s, CREAM);

  rect(s, 8.15, 0, PW - 8.15, PH, { fill: RED });

  rect(s, 0.55, 0.58, 0.22, 0.22, { fill: RED });
  txt(s, "ПРОФІЛЬ ВИРОБНИКА · 01", 0.88, 0.52, 6, 0.35, { size: 11, bold: true, color: GREY, charSpacing: 1.4 });

  txt(s, "SIAS France", 0.55, 0.95, 7.3, 0.95, { font: FONT_H, bold: true, size: 40, color: INK });
  txt(s, "Choi's Ramyeon · Instant Noodle Bag · Retail", 0.58, 1.72, 7.3, 0.4, { size: 13.5, color: GREY, italic: true, font: FONT_H });

  roundRect(s, 0.58, 2.28, 2.05, 0.46, { fill: "FFFFFF", radius: 0.23 });
  txt(s, "EXW ROYE, FRANCE", 0.58, 2.28, 2.05, 0.46, { align: "center", valign: "middle", size: 10.5, bold: true, color: RED });

  txt(s, "ПРО ВИРОБНИЦТВО", 0.58, 3.02, 6, 0.3, { size: 10.5, bold: true, color: GOLD, charSpacing: 1.3 });
  txt(s,
    "SIAS France — європейська виробнича площадка корейської групи SIAS (з 1997 року на ринку " +
    "азійської гастрономії). Лінійку Choi's Ramyeon виробляють на заводах у Roye та Wisches — " +
    "за корейськими рецептурами, адаптованими під європейського споживача.",
    0.58, 3.35, 7.15, 1.35, { size: 13, color: INK, lineSpacing: 1.3 });

  const stats = [
    ["2020", "рік запуску заводу\nSIAS Roye (Франція)"],
    ["16 000 м²", "площа виробничого\nмайданчика"],
    ["0%", "мито в Україну за\nсертифікатом EUR.1"],
  ];
  const sw = 2.33, sgap = 0.13;
  stats.forEach((st, i) => {
    const x = 0.58 + i * (sw + sgap);
    roundRect(s, x, 4.62, sw, 1.15, { fill: i === 2 ? RED : "FFFFFF", radius: 0.1, shadow: softShadow({ opacity: 0.1 }) });
    txt(s, st[0], x + 0.16, 4.72, sw - 0.32, 0.46, { font: FONT_H, bold: true, size: 22, color: i === 2 ? "FFFFFF" : RED });
    txt(s, st[1], x + 0.16, 5.18, sw - 0.32, 0.55, { size: 9, color: i === 2 ? "FCE2E4" : GREY, lineSpacing: 1.12 });
  });

  txt(s, "СЕРТИФІКАТИ ЯКОСТІ", 0.58, 6.05, 6, 0.28, { size: 10.5, bold: true, color: GOLD, charSpacing: 1.3 });
  pic(s, A("assets/cert_ifs_brc.jpg"), 0.58, 6.34, 1.75, 0.5, { sizingType: "contain" });
  pic(s, A("assets/cert_france_relance.jpg"), 2.5, 6.26, 0.6, 0.6, { sizingType: "contain" });

  // right red panel
  roundRect(s, 8.55, 0.6, 4.2, 1.55, { fill: "FFFFFF", radius: 0.12 });
  pic(s, A("assets/sias_logo.jpg"), 8.85, 0.95, 3.6, 0.86, { sizingType: "contain" });

  roundRect(s, 8.55, 2.35, 4.2, 3.55, { fill: "FFFFFF", radius: 0.12 });
  pic(s, A("assets/factory.jpg"), 8.72, 2.52, 3.86, 2.75, { sizingType: "cover" });
  txt(s, "Виробничий майданчик · Roye, Франція", 8.72, 5.35, 3.86, 0.45, { italic: true, size: 9.5, color: GREY, align: "center" });

  txt(s, "ВИРОБНИЦТВО", 8.55, 6.25, 4.2, 0.3, { align: "center", size: 10.5, bold: true, color: "F9D9DB", charSpacing: 1.5 });
  txt(s, "Франція", 8.55, 6.5, 4.2, 0.55, { align: "center", font: FONT_H, bold: true, size: 22, color: "FFFFFF" });
  frFlag(s, 10.35, 7.12, 0.5, 0.32);
  uaFlag(s, 10.95, 7.12, 0.5, 0.32);

  pageNum(s, 2);
}

// =====================================================================
// PRODUCT PRICING SLIDES (2 SKU per slide, 3 volume tiers each)
// =====================================================================
function priceRow(slide, x, y, w, label, d, opts = {}) {
  txt(slide, label, x, y, w * 0.42, 0.3, { size: 10, bold: true, color: opts.color || GREY });
  txt(slide, euro(d.eur), x + w * 0.42, y, w * 0.3, 0.3, { size: 12, bold: true, color: opts.color || INK, align: "right" });
  txt(slide, uah(d.uah), x + w * 0.72, y, w * 0.28, 0.3, { size: 9.5, color: opts.color ? opts.color : GREY, align: "right" });
}

function productCard(slide, sk, x, y, w) {
  const d = PRICING[sk.key];
  const pad = 0.16;
  let cy = y + pad;

  const imgH = 1.5;
  const cx = x + w / 2, cyc = cy + imgH / 2;
  slide.addShape("ellipse", { x: cx - 1.0, y: cyc - 0.88, w: 2.0, h: 1.76, fill: { color: sk.accent, transparency: 88 }, line: { type: "none" } });
  pic(slide, sk.img, x + w / 2 - 0.78, cy, 1.56, imgH, { sizingType: "contain" });
  cy += imgH + 0.12;

  rect(slide, x + 0.35, cy + 0.05, 0.12, 0.12, { fill: sk.accent });
  txt(slide, sk.weight + "  ·  20 ШТ/КОРОБ", x + 0.55, cy, w - 0.9, 0.24, { size: 9.5, bold: true, color: GREY, charSpacing: 0.6 });
  cy += 0.32;

  txt(slide, sk.name, x + 0.35, cy, w - 1.6, 0.38, { font: FONT_H, bold: true, size: 17.5, color: INK });
  txt(slide, sk.kr, x + w - 1.15, cy + 0.05, 0.9, 0.3, { font: FONT_KR, bold: true, size: 12.5, color: sk.accent, align: "right" });
  cy += 0.36;

  txt(slide, sk.sub, x + 0.35, cy, w - 0.7, 0.22, { italic: true, size: 10.5, color: GOLD });
  cy += 0.26;

  txt(slide, sk.desc, x + 0.35, cy, w - 0.7, 0.46, { size: 10, color: INK, lineSpacing: 1.18 });
  cy += 0.5;

  txt(slide, "СОБІВАРТІСТЬ · 1 ПАЧКА " + sk.weight, x + 0.35, cy, w - 0.7, 0.22, { size: 9, bold: true, color: GREY, charSpacing: 0.6 });
  cy += 0.26;

  const boxH = 0.66;
  roundRect(slide, x + 0.35, cy, w - 0.7, boxH, { fill: RED_DEEP, radius: 0.08 });
  txt(slide, "33 ПАЛЕТИ · FULL TRUCK", x + 0.55, cy + 0.07, w - 1.1, 0.2, { size: 8.5, bold: true, color: "F2C9CC", charSpacing: 0.8 });
  txt(slide, euro(d.p33.eur), x + 0.55, cy + 0.28, (w - 1.1) * 0.46, 0.32, { font: FONT_H, bold: true, size: 19, color: "FFFFFF" });
  txt(slide, uah(d.p33.uah) + "  ·  " + euro(d.p33.kg) + "/кг", x + 0.55 + (w - 1.1) * 0.4, cy + 0.35, (w - 1.1) * 0.6, 0.26, { size: 9.5, color: "F6DBDD", align: "right" });
  cy += boxH + 0.12;

  priceRow(slide, x + 0.35, cy, w - 0.7, "12 палет", d.p12);
  cy += 0.28;
  rect(slide, x + 0.35, cy, w - 0.7, 0.012, { fill: HAIRLINE });
  cy += 0.08;
  priceRow(slide, x + 0.35, cy, w - 0.7, "6 палет", d.p6);
  cy += 0.28;

  return cy + pad - y; // total card height used
}

function pricingSlide(idx, num, title, theme, tag, skuKeys) {
  const s = pres.addSlide();
  bg(s, CREAM);
  rect(s, 0, 0, PW, 0.08, { fill: RED });

  kicker(s, num, "CHOI'S RAMYEON");
  txt(s, title, 0.55, 0.9, 8.6, 0.46, { font: FONT_H, bold: true, size: 27, color: INK });
  txt(s, theme, 0.58, 1.4, 8.6, 0.32, { italic: true, size: 13.5, color: GOLD, font: FONT_H });

  roundRect(s, 9.6, 0.5, 3.2, 0.62, { fill: "FFFFFF", radius: 0.31, shadow: softShadow({ opacity: 0.08 }) });
  txt(s, tag, 9.6, 0.5, 3.2, 0.62, { align: "center", valign: "middle", size: 9.5, bold: true, color: RED, lineSpacing: 1.1 });

  const cw = 5.92, gap = 0.4, y0 = 1.8, CARD_HEIGHT = 5.05;
  skuKeys.forEach((key, i) => {
    const sk = SKUS.find((x) => x.key === key);
    const x = 0.55 + i * (cw + gap);
    roundRect(s, x, y0, cw, CARD_HEIGHT, { fill: CARD, radius: 0.12, shadow: softShadow() });
    productCard(s, sk, x, y0, cw);
  });

  txt(s,
    "Собівартість включає EXW, логістику, брокера, митне оформлення і ПДВ-передфінансування, за ставкою мита 0% (EUR.1). Курс: 1 € ≈ 51,2 ₴.",
    0.55, 6.94, 12.2, 0.22, { size: 8.5, italic: true, color: GREY });

  pageNum(s, idx);
}

pricingSlide(3, "02", "Гауда та Карбонара", "вершкові смаки", "EXW ROYE → УКРАЇНА\nEUR.1 · 0% МИТО", ["gouda", "carbonara"]);
pricingSlide(4, "03", "Овочева та Гостра", "легкі й пікантні смаки", "EXW ROYE → УКРАЇНА\nEUR.1 · 0% МИТО", ["vegetable", "spicy"]);
pricingSlide(5, "04", "Яловичина та Курка", "м'ясні смаки", "EXW ROYE → УКРАЇНА\nEUR.1 · 0% МИТО", ["beef", "chicken"]);

// =====================================================================
pres.writeFile({ fileName: A("SIAS_France_Presentation.pptx") }).then(() => {
  console.log("Deck written.");
}).catch((e) => {
  console.error(e);
  process.exit(1);
});
