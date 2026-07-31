const pptxgen = require("pptxgenjs");
const path = require("path");

const A = (p) => path.join(__dirname, p);

// ---------- PALETTE ----------
const RED = "D80F19";
const RED_DARK = "7A0A10";
const INK = "1A1A1A";
const CHARCOAL = "262626";
const PAPER = "FFFFFF";
const OFFWHITE = "F7F5F3";
const GOLD = "C9962C";
const GOLD_LIGHT = "F2E2BB";
const GREEN = "17784A";
const GREEN_LIGHT = "DCEFE3";
const NAVY = "0B2265";
const GREY = "6B6B6B";
const LGREY = "E9E5E1";

const FONT = "Calibri";
const FONT_H = "Cambria";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
const PW = 13.333, PH = 7.5;

function bg(slide, color) { slide.background = { color }; }

function rect(slide, x, y, w, h, opts = {}) {
  slide.addShape("rect", {
    x, y, w, h,
    fill: opts.fill ? { color: opts.fill } : { type: "none" },
    line: opts.line || { type: "none" },
  });
}

function roundRect(slide, x, y, w, h, opts = {}) {
  const shapeOpts = {
    x, y, w, h,
    rectRadius: opts.radius !== undefined ? opts.radius : 0.08,
    fill: opts.fill ? { color: opts.fill } : { type: "none" },
    line: opts.line || { type: "none" },
  };
  if (opts.shadow) shapeOpts.shadow = opts.shadow;
  slide.addShape("roundRect", shapeOpts);
}

function txt(slide, text, x, y, w, h, opts = {}) {
  slide.addText(text, {
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
  });
}

function bullets(slide, items, x, y, w, h, opts = {}) {
  const paras = items.map((it, i) => ({
    text: it,
    options: {
      bullet: { code: "25AA", indent: 14 },
      color: opts.color || INK,
      fontFace: opts.font || FONT,
      fontSize: opts.size || 13,
      breakLine: true,
      paraSpaceAfter: opts.gap !== undefined ? opts.gap : 10,
      bold: !!opts.bold,
    },
  }));
  slide.addText(paras, { x, y, w, h, valign: "top", lineSpacingMultiple: 1.12 });
}

function pageNum(slide, n, dark) {
  txt(slide, String(n).padStart(2, "0"), PW - 0.75, PH - 0.42, 0.5, 0.3, {
    size: 10, color: dark ? "FFFFFF" : GREY, align: "right",
  });
  txt(slide, "SIAS FRANCE", 0.5, PH - 0.42, 3, 0.3, {
    size: 9, color: dark ? "FFFFFF" : GREY, charSpacing: 1,
  });
}

function sectionTag(slide, num, label, opts = {}) {
  const color = opts.color || RED;
  txt(slide, num, 0.55, 0.5, 1.4, 0.5, { font: FONT_H, italic: true, size: 20, color, bold: true });
  txt(slide, label, 0.55, 0.88, 9.5, 0.7, { font: FONT_H, size: 30, bold: true, color: opts.titleColor || INK });
}

function frFlag(slide, x, y, w, h) {
  const seg = w / 3;
  rect(slide, x, y, seg, h, { fill: "002654" });
  rect(slide, x + seg, y, seg, h, { fill: "FFFFFF" });
  rect(slide, x + 2 * seg, y, seg, h, { fill: "ED2939" });
  slide.addShape("rect", { x, y, w, h, fill: { type: "none" }, line: { color: "D9D9D9", width: 0.75 } });
}

function uaFlag(slide, x, y, w, h) {
  rect(slide, x, y, w, h / 2, { fill: "0057B7" });
  rect(slide, x, y + h / 2, w, h / 2, { fill: "FFD700" });
  slide.addShape("rect", { x, y, w, h, fill: { type: "none" }, line: { color: "D9D9D9", width: 0.75 } });
}

function euBadge(slide, x, y, d) {
  slide.addShape("ellipse", { x, y, w: d, h: d, fill: { color: NAVY }, line: { type: "none" } });
  txt(slide, "EU", x, y + d / 2 - 0.16, d, 0.32, { align: "center", bold: true, color: "FFD700", size: d * 26, font: FONT_H });
}

function pic(slide, p, x, y, w, h, opts = {}) {
  const o = { x, y, w, h, sizing: { type: opts.sizingType || "cover", w, h } };
  if (opts.shadow) o.shadow = opts.shadow;
  slide.addImage(Object.assign({ path: p }, o));
}

// =====================================================================
// DATA
// =====================================================================
const SKUS = [
  { key: "gouda", name: "GOUDA CHEESE", nameRu: "Гауда (сыр)", weight: "123 г", img: A("products/gouda_123g.png"), tier: 0.75 },
  { key: "carbonara", name: "CARBONARA SPICY", nameRu: "Карбонара пикантная", weight: "131 г", img: A("products/carbonara_131g.png"), tier: 0.75 },
  { key: "vegetable", name: "VEGETABLE", nameRu: "Овощная", weight: "113 г", img: A("products/vegetable_113g.png"), tier: 0.65 },
  { key: "spicy", name: "SPICY", nameRu: "Острая", weight: "112,5 г", img: A("products/spicy_112.5g.png"), tier: 0.65 },
  { key: "beef", name: "BEEF", nameRu: "Говядина", weight: "113 г", img: A("products/beef_113g.png"), tier: 0.65 },
  { key: "chicken", name: "CHICKEN", nameRu: "Курица", weight: "113 г", img: A("products/chicken_113g.png"), tier: 0.65 },
];

// pricing[scenario][sku] = { pallets, costEur, costKg, costUah }
const PRICING = {
  "6": {
    gouda:      { pallets: 1, costEur: 1.31, costKg: 10.64, costUah: 67.0 },
    carbonara:  { pallets: 1, costEur: 1.31, costKg: 9.99,  costUah: 67.0 },
    vegetable:  { pallets: 1, costEur: 1.13, costKg: 10.04, costUah: 58.06 },
    spicy:      { pallets: 1, costEur: 1.13, costKg: 10.08, costUah: 58.06 },
    beef:       { pallets: 1, costEur: 1.13, costKg: 10.04, costUah: 58.06 },
    chicken:    { pallets: 1, costEur: 1.13, costKg: 10.04, costUah: 58.06 },
  },
  "12": {
    gouda:      { pallets: 3, costEur: 1.11, costKg: 9.03, costUah: 56.88 },
    carbonara:  { pallets: 3, costEur: 1.11, costKg: 8.48, costUah: 56.88 },
    vegetable:  { pallets: 1, costEur: 0.96, costKg: 8.52, costUah: 49.29 },
    spicy:      { pallets: 2, costEur: 0.96, costKg: 8.56, costUah: 49.29 },
    beef:       { pallets: 2, costEur: 0.96, costKg: 8.52, costUah: 49.29 },
    chicken:    { pallets: 1, costEur: 0.96, costKg: 8.52, costUah: 49.29 },
  },
  "33": {
    gouda:      { pallets: 7, costEur: 1.02, costKg: 8.28, costUah: 52.13 },
    carbonara:  { pallets: 7, costEur: 1.02, costKg: 7.77, costUah: 52.13 },
    vegetable:  { pallets: 4, costEur: 0.88, costKg: 7.81, costUah: 45.18 },
    spicy:      { pallets: 5, costEur: 0.88, costKg: 7.84, costUah: 45.18 },
    beef:       { pallets: 5, costEur: 0.88, costKg: 7.81, costUah: 45.18 },
    chicken:    { pallets: 5, costEur: 0.88, costKg: 7.81, costUah: 45.18 },
  },
};

// =====================================================================
// SLIDE 1 — COVER
// =====================================================================
{
  const s = pres.addSlide();
  bg(s, INK);
  pic(s, A("assets/hero_bowl.jpg"), 6.55, 0, 6.783, 7.5, {});
  rect(s, 0, 0, 6.6, 7.5, { fill: INK });
  // subtle red panel
  rect(s, 0, 5.35, 6.6, 2.15, { fill: RED_DARK });

  txt(s, "SIAS FRANCE  ·  CHOI'S RAMYEON", 0.7, 0.62, 5.6, 0.4, { size: 12, color: "E8B7B9", charSpacing: 2, bold: true });
  txt(s, "Корейская лапша\nевропейского\nпроизводства", 0.68, 1.55, 5.7, 2.9, {
    font: FONT_H, size: 40, bold: true, color: "FFFFFF", lineSpacing: 1.05,
  });
  rect(s, 0.72, 4.35, 0.55, 0.055, { fill: GOLD });
  txt(s, "Instant Noodle Bag · коммерческое предложение", 0.7, 4.55, 5.6, 0.5, { size: 15, color: "F0D9AE", italic: true, font: FONT_H });

  txt(s, "Сделано во Франции", 0.7, 5.62, 2.8, 0.35, { size: 13.5, bold: true, color: "FFFFFF" });
  txt(s, "Сертификат EUR.1 — 0% ввозной пошлины в Украину", 0.7, 5.98, 5.6, 0.6, { size: 12.5, color: "F2D9DA" });

  frFlag(s, 0.72, 6.55, 0.55, 0.37);
  uaFlag(s, 1.42, 6.55, 0.55, 0.37);
  txt(s, "EU  →  UA · EXW Roye, France", 2.15, 6.55, 4, 0.4, { size: 11.5, color: "E8B7B9", valign: "middle" });

  txt(s, "2026", PW - 1.55, 0.5, 1.0, 0.4, { align: "right", color: "FFFFFF", size: 13, bold: true, font: FONT_H });
}

// =====================================================================
// SLIDE 2 — ABOUT SIAS
// =====================================================================
{
  const s = pres.addSlide();
  bg(s, PAPER);
  pic(s, A("assets/rice_bowls.jpg"), 8.55, 0, 4.783, 7.5, {});
  rect(s, 0, 0, 0.16, 7.5, { fill: RED });

  sectionTag(s, "01", "О компании SIAS");
  txt(s, "ЮЖНОКОРЕЙСКАЯ ГРУППА · С 1997 ГОДА", 0.58, 1.62, 7.6, 0.35, { size: 11, bold: true, color: GOLD, charSpacing: 1.2 });

  txt(s,
    "SIAS входит в южнокорейскую группу SIAS — одного из ключевых игроков " +
    "агропродовольственного рынка. Компания более 25 лет специализируется " +
    "на замороженных продуктах и блюдах корейской и японской гастрономии, " +
    "разрабатывая рецептуры под заказ для производителей, дистрибьюторов " +
    "и розничных сетей (B2B, MDD/RHD).",
    0.58, 2.05, 7.55, 1.85, { size: 14.5, color: INK, lineSpacing: 1.28 });

  const stats = [
    ["$180 млн", "годовой оборот группы"],
    ["2 100+", "рецептур в портфеле"],
    ["1 000+", "клиентов по всему миру"],
    ["60%+", "продаж — экспорт, в основном в ЕС"],
  ];
  const sw = 1.82, gap = 0.15;
  stats.forEach((st, i) => {
    const x = 0.58 + i * (sw + gap);
    roundRect(s, x, 4.15, sw, 1.65, { fill: OFFWHITE, radius: 0.09 });
    txt(s, st[0], x + 0.12, 4.32, sw - 0.24, 0.55, { font: FONT_H, bold: true, size: 19, color: RED });
    txt(s, st[1], x + 0.12, 5.0, sw - 0.24, 0.7, { size: 10, color: GREY, lineSpacing: 1.15 });
  });

  txt(s, "«Импортируем корейское know-how в сердце Европы, адаптируя рецептуры под европейского потребителя»", 0.58, 6.0, 7.55, 0.9,
    { italic: true, size: 13, color: CHARCOAL, font: FONT_H });

  pageNum(s, 2);
}

// =====================================================================
// SLIDE 3 — GROUP & HISTORY
// =====================================================================
{
  const s = pres.addSlide();
  bg(s, PAPER);
  sectionTag(s, "02", "Группа SIAS и её развитие");

  const points = [
    ["1997", "SIAS KOREA", "7 промышленных площадок в Корее · основа группы"],
    ["2020", "SIAS ROYE", "Собственный завод во Франции — в центре Европы"],
    ["2022", "YOMIA BY SIAS", "Национальный французский бренд для розницы"],
    ["2024", "SIAS ALSACE", "Завод в Wisches: производство сушёной лапши Choi's"],
  ];
  const cw = 2.9, gap = 0.18, startX = 0.58, y0 = 1.85;
  points.forEach((p, i) => {
    const x = startX + i * (cw + gap);
    roundRect(s, x, y0, cw, 3.35, { fill: i >= 2 ? OFFWHITE : "FBEAEA", radius: 0.09 });
    s.addShape("ellipse", { x: x + 0.22, y: y0 + 0.25, w: 0.5, h: 0.5, fill: { color: i % 2 === 0 ? RED : INK }, line: { type: "none" } });
    txt(s, p[0], x + 0.22, y0 + 0.25, 0.5, 0.5, { align: "center", valign: "middle", color: "FFFFFF", bold: true, size: 11 });
    txt(s, p[1], x + 0.22, y0 + 0.98, cw - 0.44, 0.5, { font: FONT_H, bold: true, size: 15.5, color: INK });
    txt(s, p[2], x + 0.22, y0 + 1.5, cw - 0.44, 1.7, { size: 11.5, color: GREY, lineSpacing: 1.25 });
  });

  // group footprint
  roundRect(s, 0.58, 5.5, 12.2, 1.45, { fill: INK, radius: 0.1 });
  const cols = [
    ["7", "заводов в Ю. Корее"],
    ["2", "завода во Франции (Roye, Wisches)"],
    ["60", "сотрудников на площадках во Франции"],
    ["4", "производственные линии FR"],
  ];
  const cw2 = 12.2 / 4;
  cols.forEach((c, i) => {
    const x = 0.58 + i * cw2;
    txt(s, c[0], x + 0.25, 5.68, cw2 - 0.4, 0.55, { font: FONT_H, bold: true, size: 26, color: GOLD });
    txt(s, c[1], x + 0.25, 6.28, cw2 - 0.4, 0.55, { size: 10.5, color: "FFFFFF" });
  });

  pageNum(s, 3);
}

// =====================================================================
// SLIDE 4 — SIAS FRANCE FACILITY
// =====================================================================
{
  const s = pres.addSlide();
  bg(s, PAPER);
  sectionTag(s, "03", "Производство SIAS France");

  roundRect(s, 0.58, 1.85, 6.35, 3.85, { fill: OFFWHITE, radius: 0.1 });
  pic(s, A("assets/factory.jpg"), 0.78, 2.05, 5.95, 2.9, { sizingType: "cover" });
  txt(s, "Завод SIAS · Roye, Франция (регион О-де-Франс)", 0.78, 5.05, 5.95, 0.55, { size: 11, italic: true, color: GREY });

  const info = [
    "16 000 м² площадки, из них 3 000 м² — производство, склад и офисы",
    "Стратегическое расположение — «перекрёсток Европы»: рядом Лондон, Брюссель, Кёльн, Париж",
    "2024: автоматизация линии гёдза и оптимизация линии риса",
    "Поддержка France Relance, сертификаты IFS Food и BRC Food",
  ];
  bullets(s, info, 7.2, 1.95, 5.55, 2.6, { size: 13, gap: 12 });

  pic(s, A("assets/europe_map.png"), 7.55, 4.55, 2.05, 2.05, { sizingType: "contain" });
  const certs = ["IFS Food", "BRC Food", "France Relance"];
  certs.forEach((c, i) => {
    const x = 9.85, y = 4.7 + i * 0.68;
    s.addShape("ellipse", { x, y, w: 0.16, h: 0.16, fill: { color: GREEN }, line: { type: "none" } });
    txt(s, c, x + 0.3, y - 0.09, 2.9, 0.4, { size: 13, bold: true, color: INK });
  });
  txt(s, "Гарантии надёжности и государственная поддержка", 9.85, 6.75, 3.1, 0.6, { size: 10, color: GREY, italic: true });

  pageNum(s, 4);
}

// =====================================================================
// SLIDE 5 — MADE IN EU + EUR.1 = 0% DUTY  (KEY SLIDE)
// =====================================================================
{
  const s = pres.addSlide();
  bg(s, INK);
  rect(s, 0, 0, PW, 1.5, { fill: RED_DARK });
  txt(s, "04", 0.55, 0.42, 1.4, 0.5, { font: FONT_H, italic: true, size: 20, color: "F0AEB2", bold: true });
  txt(s, "Сделано в ЕС — ключевое преимущество для Украины", 0.55, 0.78, 12.2, 0.6, { font: FONT_H, size: 26, bold: true, color: "FFFFFF" });

  // Big 0% callout
  roundRect(s, 0.58, 1.95, 3.9, 4.55, { fill: GREEN, radius: 0.14 });
  txt(s, "0%", 0.58, 2.35, 3.9, 1.6, { font: FONT_H, bold: true, size: 92, color: "FFFFFF", align: "center" });
  txt(s, "ВВОЗНАЯ ПОШЛИНА", 0.58, 3.95, 3.9, 0.45, { bold: true, size: 15, color: "FFFFFF", align: "center", charSpacing: 1.5 });
  txt(s, "при импорте в Украину благодаря\nсертификату EUR.1", 0.58, 4.45, 3.9, 1.0, { size: 12.5, color: "E7F3EC", align: "center", lineSpacing: 1.25 });
  rect(s, 1.1, 5.55, 2.85, 0.02, { fill: "FFFFFF" });
  txt(s, "Уже учтено в расчёте себестоимости\nв этом предложении", 0.58, 5.7, 3.9, 0.7, { size: 11, italic: true, color: "D7ECDF", align: "center" });

  // Right explanation
  const rx = 4.85;
  txt(s, "Как это работает", rx, 2.0, 7.9, 0.45, { font: FONT_H, bold: true, size: 18, color: "FFFFFF" });
  const steps = [
    ["Страна происхождения — ЕС", "Линейка Choi's Ramyeon Bag производится на заводах SIAS во Франции (Roye / Wisches)."],
    ["Сертификат EUR.1", "SIAS предоставляет сертификат перемещения товара EUR.1, подтверждающий преференциальное происхождение продукции."],
    ["Соглашение об ассоциации Украина–ЕС (DCFTA)", "EUR.1 даёт право на 0% ввозной пошлины при таможенном оформлении в Украине — вместо стандартной ставки."],
    ["Прозрачный расчёт", "Себестоимость в этом предложении уже построена по ставке пошлины 0% — цифры на следующих слайдах отражают реальные условия поставки."],
  ];
  let sy = 2.55;
  steps.forEach((st, i) => {
    s.addShape("ellipse", { x: rx, y: sy, w: 0.42, h: 0.42, fill: { color: GOLD }, line: { type: "none" } });
    txt(s, String(i + 1), rx, sy, 0.42, 0.42, { align: "center", valign: "middle", bold: true, color: INK, size: 15 });
    txt(s, st[0], rx + 0.58, sy - 0.03, 7.3, 0.4, { bold: true, size: 14, color: "FFFFFF" });
    txt(s, st[1], rx + 0.58, sy + 0.36, 7.3, 0.75, { size: 11.5, color: "D8D8D8", lineSpacing: 1.2 });
    sy += 1.08;
  });

  frFlag(s, rx, 6.95, 0.42, 0.28);
  uaFlag(s, rx + 0.52, 6.95, 0.42, 0.28);
  txt(s, "France (EXW Roye)  →  Ukraine, 0% duty (EUR.1)", rx + 1.05, 6.9, 5, 0.4, { size: 10.5, color: "E7C7C9", valign: "middle" });

  pageNum(s, 5, true);
}

// =====================================================================
// SLIDE 6 — OUR CLIENTS
// =====================================================================
{
  const s = pres.addSlide();
  bg(s, PAPER);
  sectionTag(s, "05", "Нам доверяют");
  txt(s, "Партнёр розничных сетей, дистрибьюторов и HoReCa Европы", 0.58, 1.6, 10, 0.4, { size: 13, italic: true, color: GREY });

  roundRect(s, 0.58, 2.1, 12.2, 4.7, { fill: OFFWHITE, radius: 0.1 });
  pic(s, A("assets/clients_logos.jpg"), 0.9, 2.35, 11.55, 4.2, { sizingType: "contain" });

  pageNum(s, 6);
}

// =====================================================================
// SLIDE 7 — PRODUCT RANGE
// =====================================================================
{
  const s = pres.addSlide();
  bg(s, PAPER);
  sectionTag(s, "06", "Choi's Ramyeon — Instant Noodle Bag");
  txt(s, "6 вкусов · пакет (bag) · 20 шт/короб · 80 коробов/паллета", 0.58, 1.6, 11, 0.4, { size: 13, italic: true, color: GREY });

  const cw = 1.92, gap = 0.155, startX = 0.58, y0 = 2.25, imgH = 2.55;
  SKUS.forEach((sk, i) => {
    const x = startX + i * (cw + gap);
    roundRect(s, x, y0, cw, 4.55, { fill: OFFWHITE, radius: 0.1 });
    pic(s, sk.img, x + 0.18, y0 + 0.18, cw - 0.36, imgH, { sizingType: "contain" });
    rect(s, x + 0.15, y0 + imgH + 0.32, cw - 0.3, 0.014, { fill: LGREY });
    txt(s, sk.nameRu, x + 0.14, y0 + imgH + 0.42, cw - 0.28, 0.7, { bold: true, size: 12.5, color: INK, align: "center", font: FONT_H });
    txt(s, sk.weight, x + 0.14, y0 + imgH + 1.0, cw - 0.28, 0.35, { size: 12, color: RED, align: "center", bold: true });
  });

  txt(s, "Made in France  ·  без ГМО  ·  без консервантов  ·  формула clean label",
    0.58, 6.95, 12.2, 0.4, { size: 11.5, color: GREY, align: "center", italic: true });

  pageNum(s, 7);
}

// =====================================================================
// SLIDES 8-10 — PRICING PER PALLET SCENARIO
// =====================================================================
function euro(n) { return n.toFixed(2).replace(".", ",") + " €"; }
function uah(n) { return n.toFixed(2).replace(".", ",") + " грн"; }

function pricingSlide(num, scenarioKey, title, subtitle, badgeText, idx) {
  const s = pres.addSlide();
  bg(s, PAPER);
  sectionTag(s, num, title);
  txt(s, subtitle, 0.58, 1.6, 8.6, 0.4, { size: 13, italic: true, color: GREY });

  roundRect(s, 10.35, 0.55, 2.42, 1.15, { fill: RED, radius: 0.1 });
  txt(s, badgeText, 10.35, 0.66, 2.42, 0.48, { align: "center", bold: true, size: 21, color: "FFFFFF", font: FONT_H });
  txt(s, "EXW Roye → Украина, DDP расчёт", 10.35, 1.2, 2.42, 0.42, { align: "center", size: 8.5, color: "F6D3D5" });

  const cw = 1.92, gap = 0.155, startX = 0.58, y0 = 2.1, imgH = 1.55;
  const data = PRICING[scenarioKey];
  SKUS.forEach((sk, i) => {
    const x = startX + i * (cw + gap);
    const d = data[sk.key];
    roundRect(s, x, y0, cw, 4.35, { fill: OFFWHITE, radius: 0.1 });
    pic(s, sk.img, x + 0.24, y0 + 0.16, cw - 0.48, imgH, { sizingType: "contain" });
    txt(s, sk.nameRu, x + 0.12, y0 + imgH + 0.26, cw - 0.24, 0.5, { bold: true, size: 11.5, color: INK, align: "center", font: FONT_H });
    txt(s, sk.weight, x + 0.12, y0 + imgH + 0.68, cw - 0.24, 0.28, { size: 10.5, color: GREY, align: "center" });

    roundRect(s, x + 0.12, y0 + imgH + 0.98, cw - 0.24, 1.28, { fill: "FFFFFF", radius: 0.08 });
    txt(s, "СЕБЕСТОИМОСТЬ / ПАЧКА", x + 0.12, y0 + imgH + 1.04, cw - 0.24, 0.26, { align: "center", size: 7.5, color: GREY, bold: true, charSpacing: 0.5 });
    txt(s, euro(d.costEur), x + 0.12, y0 + imgH + 1.26, cw - 0.24, 0.38, { align: "center", bold: true, size: 18, color: RED, font: FONT_H });
    txt(s, uah(d.costUah), x + 0.12, y0 + imgH + 1.63, cw - 0.24, 0.24, { align: "center", size: 9, color: GREY });
    rect(s, x + 0.3, y0 + imgH + 1.92, cw - 0.6, 0.012, { fill: LGREY });
    txt(s, euro(d.costKg) + " / кг", x + 0.12, y0 + imgH + 1.99, cw - 0.24, 0.28, { align: "center", bold: true, size: 11, color: GOLD });

    roundRect(s, x + cw - 0.72, y0 + 0.05, 0.6, 0.32, { fill: INK, radius: 0.16 });
    txt(s, d.pallets + " пал.", x + cw - 0.72, y0 + 0.05, 0.6, 0.32, { align: "center", valign: "middle", size: 8, color: "FFFFFF", bold: true });
  });

  txt(s,
    "Себестоимость включает: EXW-цену, международную логистику, брокерские услуги, таможенное оформление и НДС-предфинансирование, при ставке ввозной пошлины 0% (сертификат EUR.1). Курс расчёта: 1 € ≈ 51,2 грн.",
    0.58, 6.68, 12.2, 0.5, { size: 9, italic: true, color: GREY });

  pageNum(s, idx);
}

pricingSlide("07", "6", "Себестоимость — 6 паллет", "Пробная партия · 1 паллета на каждый SKU · 9 600 пачек", "6 ПАЛЛЕТ", 8);
pricingSlide("08", "12", "Себестоимость — 12 паллет", "Средний объём · 19 200 пачек", "12 ПАЛЛЕТ", 9);
pricingSlide("09", "33", "Себестоимость — Full Truck", "Полный тираж · 33 паллеты · 52 800 пачек · максимальная экономия", "33 ПАЛЛЕТЫ", 10);

// =====================================================================
// SLIDE 11 — ECONOMIES OF SCALE (CHART)
// =====================================================================
{
  const s = pres.addSlide();
  bg(s, PAPER);
  sectionTag(s, "10", "Эффект масштаба поставки");
  txt(s, "Себестоимость за 1 кг продукта, €, в зависимости от объёма поставки", 0.58, 1.6, 11, 0.4, { size: 13, italic: true, color: GREY });

  const cats = ["6 паллет", "12 паллет", "33 паллеты (Full Truck)"];
  const chartData = [
    { name: "Тариф 0,65 €/уп. (Курица, Говядина, Овощная, Острая)", labels: cats, values: [10.04, 8.52, 7.81] },
    { name: "Тариф 0,75 €/уп. (Гауда, Карбонара)", labels: cats, values: [10.64, 9.03, 8.28] },
  ];
  s.addChart("bar", chartData, {
    x: 0.58, y: 2.15, w: 8.1, h: 4.55,
    barDir: "col",
    chartColors: [RED, GOLD],
    showTitle: false,
    showLegend: true,
    legendPos: "b",
    legendColor: INK,
    legendFontSize: 11,
    showValue: true,
    dataLabelPosition: "outEnd",
    dataLabelColor: INK,
    dataLabelFontSize: 10,
    dataLabelFormatCode: '0.00" €"',
    catAxisLabelColor: INK,
    catAxisLabelFontSize: 11,
    valAxisLabelColor: GREY,
    valAxisLabelFontSize: 10,
    valAxisTitle: "€ / кг",
    showValAxisTitle: true,
    valGridLine: { color: LGREY, size: 0.75 },
    catGridLine: { style: "none" },
    barGapWidthPct: 45,
  });

  const rx = 9.0;
  roundRect(s, rx, 2.15, 3.75, 4.55, { fill: INK, radius: 0.1 });
  txt(s, "ЭКОНОМИЯ ОБЪЁМА", rx + 0.3, 2.4, 3.15, 0.35, { bold: true, size: 12, color: GOLD, charSpacing: 1 });
  const savings = [
    ["-22%", "себестоимость/кг тарифа 0,65 € с ростом 6 → 33 паллеты"],
    ["-22%", "себестоимость/кг тарифа 0,75 € с ростом 6 → 33 паллеты"],
    ["0%", "пошлина на всех объёмах — сертификат EUR.1"],
  ];
  let sy2 = 2.95;
  savings.forEach((r) => {
    txt(s, r[0], rx + 0.3, sy2, 1.3, 0.6, { font: FONT_H, bold: true, size: 26, color: "FFFFFF" });
    txt(s, r[1], rx + 0.3, sy2 + 0.6, 3.15, 0.7, { size: 10, color: "D8D8D8", lineSpacing: 1.2 });
    sy2 += 1.2;
  });

  pageNum(s, 11);
}

// =====================================================================
// SLIDE 12 — WHY SIAS (PRODUCTION STRENGTH)
// =====================================================================
{
  const s = pres.addSlide();
  bg(s, PAPER);
  sectionTag(s, "11", "Почему SIAS");

  roundRect(s, 0.58, 1.85, 12.2, 2.15, { fill: OFFWHITE, radius: 0.1 });
  pic(s, A("assets/production_line.png"), 0.9, 2.0, 11.55, 1.6, { sizingType: "contain" });
  txt(s, "Полная производственная линия — 150 метров, от замеса теста до упаковки", 0.9, 3.62, 11.55, 0.35, { size: 10.5, italic: true, color: GREY, align: "center" });

  const cards = [
    ["Эффективное производство", "Автоматизированные линии обеспечивают высокую производительность без потери качества."],
    ["Контроль качества", "Регулярные тесты и проверки на каждом этапе — уверенность потребителей в продукте."],
    ["Быстрая поставка", "Большая производственная мощность формирует быструю и надёжную цепочку поставок."],
    ["Технологии и R&D", "Современное оборудование и разработки для улучшения вкуса и качества лапши."],
  ];
  const cw = 2.9, gap = 0.18, y0 = 4.35;
  cards.forEach((c, i) => {
    const x = 0.58 + i * (cw + gap);
    roundRect(s, x, y0, cw, 2.55, { fill: i % 2 === 0 ? "FBEAEA" : OFFWHITE, radius: 0.1 });
    rect(s, x + 0.22, y0 + 0.25, 0.42, 0.42, { fill: RED });
    txt(s, String(i + 1), x + 0.22, y0 + 0.25, 0.42, 0.42, { align: "center", valign: "middle", bold: true, color: "FFFFFF", size: 15 });
    txt(s, c[0], x + 0.22, y0 + 0.85, cw - 0.44, 0.55, { bold: true, size: 13.5, color: INK, font: FONT_H });
    txt(s, c[1], x + 0.22, y0 + 1.38, cw - 0.44, 1.1, { size: 10.5, color: GREY, lineSpacing: 1.22 });
  });

  pageNum(s, 12);
}

// =====================================================================
// SLIDE 13 — THANK YOU / CONTACT
// =====================================================================
{
  const s = pres.addSlide();
  bg(s, INK);
  rect(s, 0, 0, PW, 7.5, { fill: INK });
  rect(s, 0, 0, PW, 0.12, { fill: RED });

  txt(s, "Спасибо за внимание", 0.9, 2.35, 11.5, 1.0, { font: FONT_H, bold: true, size: 42, color: "FFFFFF" });
  txt(s, "Готовы обсудить объёмы поставки, условия EXW/DDP и сроки производства", 0.9, 3.3, 11.0, 0.6, { size: 15, italic: true, color: "D9B8BA", font: FONT_H });

  roundRect(s, 0.9, 4.35, 5.6, 1.85, { fill: RED_DARK, radius: 0.1 });
  txt(s, "SIAS FRANCE", 1.15, 4.55, 5.1, 0.4, { bold: true, size: 16, color: "FFFFFF", font: FONT_H });
  txt(s, "Rue du Champ Macret · BP 20041\n80700 ROYE, France", 1.15, 4.95, 5.1, 0.7, { size: 11.5, color: "F0D9DA", lineSpacing: 1.3 });
  txt(s, "SIRET 883 646 739 00028 · TVA FR69 883 646 739", 1.15, 5.7, 5.1, 0.4, { size: 9.5, color: "CFA9AB" });

  roundRect(s, 6.75, 4.35, 5.65, 1.85, { fill: GREEN, radius: 0.1 });
  txt(s, "0% пошлина · EUR.1", 7.0, 4.55, 5.1, 0.4, { bold: true, size: 16, color: "FFFFFF", font: FONT_H });
  txt(s, "Made in France · IFS & BRC Food certified", 7.0, 4.98, 5.1, 0.4, { size: 12, color: "E7F3EC" });
  txt(s, "Партия готова к отгрузке от 6 паллет до Full Truck (33 паллеты)", 7.0, 5.4, 5.1, 0.65, { size: 10.5, color: "CBE7D6", lineSpacing: 1.25 });

  frFlag(s, 0.9, 6.55, 0.5, 0.34);
  uaFlag(s, 1.55, 6.55, 0.5, 0.34);

  pageNum(s, 13, true);
}

// =====================================================================
pres.writeFile({ fileName: A("SIAS_France_Presentation.pptx") }).then(() => {
  console.log("Deck written.");
}).catch((e) => {
  console.error(e);
  process.exit(1);
});
