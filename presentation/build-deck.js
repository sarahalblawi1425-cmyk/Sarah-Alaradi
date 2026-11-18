// Mirsah (مرساة) pitch deck v4 — self-explanatory pre-read, website palette
const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const path = require("path");
const lu = require("react-icons/lu");
const { applyTheme } = require(process.env.SKILL + "/scripts/apply_theme.js");

const SHOTS = process.argv[2], OUT = process.argv[3];

// ---- website palette ----
const NAVY = "0B2545", DEEP = "0B1D33", NAVY2 = "13315C", LINE_D = "1E3A5F", TEAL = "0E7C86", TEAL_D = "0B5C66",
  ICE = "8FD3DA", ICE_BG = "DDF2F3", BG = "F4F6F9", LINE = "DDE3EB", MUTED = "4A5568", INK = "0F1B2D",
  SOFT = "C9D6E8", WHITE = "FFFFFF", RED = "E5383B", ORANGE = "EA6A2A", AMBER = "E0A526", GREEN = "3FA36B";
const F = "Arial";
const THEME = { name: "Mirsah", headFontFace: F, bodyFontFace: F,
  colors: { dk1: INK, lt1: WHITE, dk2: NAVY, lt2: BG, accent1: TEAL, accent2: ICE, accent3: AMBER, accent4: ORANGE, accent5: RED, accent6: GREEN, hlink: TEAL_D, folHlink: NAVY2 } };

const W = 13.333, H = 7.5, M = 0.6, R = W - M, CW = W - 2 * M;
const ar = (o = {}) => Object.assign({ fontFace: F, rtlMode: true, lang: "ar-SA", align: "right", isTextBox: true, margin: 0, valign: "top", color: INK }, o);
const ctr = (o = {}) => ar(Object.assign({ align: "center" }, o));

async function icon(Comp, color, size = 256) {
  const svg = RDS.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: String(size) }));
  return "image/png;base64," + (await sharp(Buffer.from(svg)).resize(size, size).png().toBuffer()).toString("base64");
}
async function rings(color, opacity) {
  const c = `#${color}`;
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="1200" height="1200" viewBox="0 0 1200 1200" fill="none">
    <g stroke="${c}" stroke-opacity="${opacity}" stroke-width="2">
      <circle cx="600" cy="600" r="150"/><circle cx="600" cy="600" r="300"/><circle cx="600" cy="600" r="450"/><circle cx="600" cy="600" r="590"/>
      <line x1="600" y1="0" x2="600" y2="1200"/><line x1="0" y1="600" x2="1200" y2="600"/></g>
    <path d="M600 600 L1190 600 A590 590 0 0 0 1017 183 Z" fill="${c}" fill-opacity="${opacity * 0.6}"/></svg>`;
  return "image/png;base64," + (await sharp(Buffer.from(svg)).png().toBuffer()).toString("base64");
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.rtlMode = true;
  pres.title = "مرساة — منصة وطنية للتأمين البحري وتقييم مخاطر المسارات";
  pres.theme = { headFontFace: F, bodyFontFace: F };

  const ANCHOR_ICE = await icon(lu.LuAnchor, ICE);
  const RINGS_D = await rings(ICE, 0.16);
  const RINGS_L = await rings(NAVY, 0.06);

  const footer = (dark) => ({ text: { text: "مرساة  ·  نقل ثون 5  ·  المسار الرابع", options: ar({ x: 4.5, y: 7.05, w: 4.33, h: 0.3, fontSize: 10, color: dark ? "6E8DB3" : "8A97A8", align: "center" }) } });
  const logo = [
    { rect: { x: M, y: 0.42, w: 1.45, h: 0.5, fill: { color: NAVY }, rectRadius: 0.1 } },
    { image: { data: ANCHOR_ICE, x: M + 1.03, y: 0.5, w: 0.34, h: 0.34 } },
    { text: { text: "مرساة", options: ar({ x: M + 0.1, y: 0.42, w: 0.88, h: 0.5, fontSize: 16, bold: true, color: WHITE, valign: "middle" }) } },
  ];
  const titlePh = (color) => ({ placeholder: { options: { name: "title", type: "title", x: M + 1.8, y: 0.72, w: CW - 1.8, h: 0.75, fontFace: F, fontSize: 28, bold: true, color, align: "right", valign: "middle", rtlMode: true, margin: 0 }, text: "" } });

  pres.defineSlideMaster({ title: "LIGHT", background: { color: BG },
    objects: [{ image: { data: RINGS_L, x: -2.6, y: 4.0, w: 6.0, h: 6.0 } }, ...logo, footer(false), titlePh(NAVY)],
    slideNumber: { x: M, y: 7.05, w: 0.5, h: 0.3, fontFace: F, fontSize: 10, color: "8A97A8" } });
  pres.defineSlideMaster({ title: "DARK", background: { color: NAVY },
    objects: [{ image: { data: RINGS_D, x: -3.0, y: 3.6, w: 7.0, h: 7.0 } },
      { rect: { x: M, y: 0.42, w: 1.45, h: 0.5, fill: { color: NAVY2 }, rectRadius: 0.1 } },
      { image: { data: ANCHOR_ICE, x: M + 1.03, y: 0.5, w: 0.34, h: 0.34 } },
      { text: { text: "مرساة", options: ar({ x: M + 0.1, y: 0.42, w: 0.88, h: 0.5, fontSize: 16, bold: true, color: WHITE, valign: "middle" }) } },
      footer(true), titlePh(WHITE)],
    slideNumber: { x: M, y: 7.05, w: 0.5, h: 0.3, fontFace: F, fontSize: 10, color: "6E8DB3" } });
  pres.defineSlideMaster({ title: "COVER", background: { color: WHITE }, objects: [{ image: { data: RINGS_L, x: 9.2, y: -2.2, w: 6.0, h: 6.0 } }] });

  // ---- helpers ----
  const kicker = (s, t, dark = false) => s.addText(t, ar({ x: M + 1.8, y: 0.36, w: CW - 1.8, h: 0.35, fontSize: 13, bold: true, color: dark ? ICE : TEAL }));
  const card = (s, x, y, w, h, fill = WHITE, line = LINE, shadow = true) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: fill }, line: { color: line, width: 1 },
    shadow: shadow ? { type: "outer", color: "0B2545", opacity: 0.08, blur: 8, offset: 2, angle: 90 } : undefined });
  async function bubble(s, Comp, x, y, d, bg = ICE_BG, fg = TEAL_D) {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: bg }, line: { color: bg } });
    const p = d * 0.25; s.addImage({ data: await icon(Comp, fg), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
  }
  function badge(s, n, x, y, d = 0.42, bg = TEAL, fg = WHITE) {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: bg }, line: { color: bg } });
    s.addText(String(n), { x, y, w: d, h: d, fontFace: F, fontSize: 13, bold: true, color: fg, align: "center", valign: "middle", margin: 0, isTextBox: true });
  }
  const shot = (s, file, x, y, w, h, alt) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x - 0.06, y: y - 0.06, w: w + 0.12, h: h + 0.12, rectRadius: 0.08, fill: { color: WHITE }, line: { color: LINE },
      shadow: { type: "outer", color: "0B2545", opacity: 0.18, blur: 14, offset: 4, angle: 90 } });
    s.addImage({ path: path.join(SHOTS, file), x, y, w, h, altText: alt });
  };

  let s;
  pres.addSection({ title: "البداية" });
  // ================= 1. COVER =================
  s = pres.addSlide({ masterName: "COVER", sectionTitle: "البداية" });
  shot(s, "map.png", M + 0.06, 0.95, 5.3 * 1736 / 1460, 5.3, "خريطة رادار الخطر من النموذج الأولي");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: R - 1.1, y: 1.0, w: 1.1, h: 1.1, rectRadius: 0.2, fill: { color: NAVY }, line: { color: NAVY } });
  s.addImage({ data: ANCHOR_ICE, x: R - 0.88, y: 1.22, w: 0.66, h: 0.66 });
  s.addText("مرساة", ar({ x: 7.3, y: 2.25, w: R - 7.3, h: 1.3, fontSize: 72, bold: true, color: NAVY }));
  s.addText("منصة وطنية للتأمين البحري وتقييم مخاطر المسارات", ar({ x: 7.3, y: 3.6, w: R - 7.3, h: 0.9, fontSize: 22, bold: true, color: TEAL }));
  s.addText("نصنع مؤشر خطر سعودياً للمسارات البحرية، ونجمع لصاحب البضاعة ومالك السفينة الخطر والتكلفة وعروض المؤمّنين السعوديين في مكان واحد.", ar({ x: 7.3, y: 4.55, w: R - 7.3, h: 0.95, fontSize: 15, color: MUTED, lineSpacingMultiple: 1.15 }));
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.3, y: 5.75, w: R - 7.3, h: 0.5, rectRadius: 0.25, fill: { color: ICE_BG }, line: { color: ICE_BG } });
  s.addText("نقل ثون 5  ·  المسار الرابع: الخدمات اللوجستية وسلاسل الإمداد", ctr({ x: 7.35, y: 5.75, w: R - 7.4, h: 0.5, fontSize: 11.5, bold: true, color: TEAL_D, valign: "middle" }));
  pres.addSection({ title: "المشكلة" });
  // ================= 2. PROBLEM: geopolitics + numbers =================
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "المشكلة" });
  kicker(s, "المشكلة");
  s.addText("الأوضاع الجيوسياسية رفعت تأمين الشحن البحري بشكل حاد", { placeholder: "title" });
  s.addText("حين تعبر السفينة منطقة توتر أمني مثل البحر الأحمر أو مضيق هرمز، يُضاف قسط يُعرف في السوق باسم «تأمين مخاطر الحرب». ومع تصاعد الأوضاع الجيوسياسية في المنطقة بين 2023 و2026، تضاعف هذا القسط عشرات المرات.",
    ar({ x: M, y: 1.58, w: CW, h: 0.95, fontSize: 15, color: MUTED, lineSpacingMultiple: 1.2 }));
  {
    const cw = 7.0, cx = M, cy = 2.65, ch = 3.55;
    card(s, cx, cy, cw, ch);
    s.addText("قسط تأمين جسم السفينة: % من قيمة السفينة لكل رحلة", ar({ x: cx + 0.3, y: cy + 0.15, w: cw - 0.6, h: 0.4, fontSize: 14, bold: true, color: NAVY }));
    s.addChart(pres.charts.BAR, [{ name: "% من قيمة السفينة", labels: ["2023", "2024", "ينبع 2026", "جنوب ينبع 2026"], values: [0.05, 1, 3, 7] }], {
      x: cx + 0.15, y: cy + 0.55, w: cw - 0.3, h: ch - 0.7, barDir: "col", chartColors: ["3FA36B", "E0A526", "EA6A2A", "E5383B"],
      showValue: true, dataLabelPosition: "outEnd", dataLabelColor: INK, dataLabelFontSize: 13, dataLabelFontBold: true, dataLabelFontFace: "+mn-lt",
      dataLabelFormatCode: '0.0#"%"', catAxisLabelColor: MUTED, catAxisLabelFontSize: 12, catAxisLabelFontFace: "+mn-lt", catAxisOrientation: "maxMin",
      catAxisLineShow: false, valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false, barGapWidthPct: 55,
    });
    // right: dollar impact
    const rx = cx + cw + 0.35, rw = R - rx;
    s.addText("ماذا يعني ذلك لناقلة قيمتها 100 مليون دولار؟", ar({ x: rx, y: cy, w: rw, h: 0.4, fontSize: 14, bold: true, color: TEAL_D }));
    const rows = [["2023", "نحو 50 ألف $", "قبل تصاعد التوترات في البحر الأحمر", "2E7D4F"], ["2026 · ينبع", "نحو 3 ملايين $", "بعد أن كان القسط أقل من 1% في يوليو 2026", ORANGE], ["2026 · جنوب ينبع وهرمز", "نحو 7 ملايين $", "وفي هرمز بين 6% و9% من قيمة السفينة", RED]];
    rows.forEach((r, i) => {
      const y = cy + 0.5 + i * 1.03;
      card(s, rx, y, rw, 0.9);
      s.addShape(pres.shapes.OVAL, { x: rx + rw - 0.42, y: y + 0.2, w: 0.18, h: 0.18, fill: { color: r[3] }, line: { color: r[3] } });
      s.addText(r[0], ar({ x: rx + 0.2, y: y + 0.1, w: rw - 0.75, h: 0.32, fontSize: 12.5, bold: true, color: MUTED }));
      s.addText(r[1], ar({ x: rx + 0.2, y: y + 0.42, w: 2.0, h: 0.42, fontSize: 19, bold: true, color: r[3], valign: "middle" }));
      s.addText(r[2], ar({ x: rx + 2.2, y: y + 0.42, w: rw - 2.4 - 0.35, h: 0.42, fontSize: 11, color: MUTED, valign: "middle" }));
    });
  }
  s.addText("والأقساط لا تثبت: تُسعَّر لفترات قصيرة (7 أيام عادةً) وتُراجع يومياً.", ar({ x: M, y: 6.3, w: CW, h: 0.38, fontSize: 15, bold: true, color: NAVY }));
  s.addText("المصادر: S&P Global وPolicyholder Pulse والشرق الأوسط (2023–2024)، رويترز عبر Insurance Journal (سبتمبر 2026). أرقام عروض سوقية لتأمين جسم السفينة، وليست أسعاراً متفقاً عليها.", ar({ x: M, y: 6.68, w: CW, h: 0.28, fontSize: 9.5, color: "8A97A8" }));
  s.addNotes("2023: نحو 0.05% من قيمة السفينة لعبور البحر الأحمر. نهاية 2023 و2024: بين 0.7% و1%. 2026: ينبع نحو 3% (من أقل من 1% في يوليو)، جنوب ينبع نحو 7%، هرمز بين 6% و9%.");
  // ================= 3. PROBLEM: user pain =================
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "المشكلة" });
  kicker(s, "المشكلة");
  s.addText("والنتيجة: الشاحن يدفع أكثر، ويقرر دون صورة كاملة", { placeholder: "title" });
  const pains = [
    [lu.LuGlobe, "مرجع دولي واحد", "يُسعَّر الخطر على قوائم لجنة الحرب المشتركة في لندن، ولا تملك المملكة حتى اليوم مؤشراً وطنياً لمسارات تجارتها."],
    [lu.LuPlaneTakeoff, "اعتماد خارجي", "حضور محدود للمؤمّن السعودي: التأمين البحري والنقل (مع الطيران والمسؤولية) نحو 4.5% فقط من أقساط السوق."],
    [lu.LuScale, "لا توجد مقارنة", "الأسعار الفعلية غير معلنة، والأرقام المتداولة عروض قد تختلف عن السعر الحقيقي."],
    [lu.LuPhoneCall, "شركة شركة", "يتواصل الشاحن مع كل مؤمّن على حدة، ولا يرى أثر تغيير الميناء أو المسار على القسط."],
  ];
  pains[1][0] = lu.LuArrowUpRight;
  const pw = (CW - 3 * 0.3) / 4;
  for (let i = 0; i < 4; i++) {
    const x = R - pw - i * (pw + 0.3), y = 1.85;
    card(s, x, y, pw, 4.55);
    await bubble(s, pains[i][0], x + pw - 1.15, y + 0.35, 0.85, i === 1 ? "FDE8E2" : ICE_BG, i === 1 ? "C2410C" : TEAL_D);
    s.addText(pains[i][1], ar({ x: x + 0.3, y: y + 1.45, w: pw - 0.6, h: 0.55, fontSize: 19, bold: true, color: NAVY }));
    s.addText(pains[i][2], ar({ x: x + 0.3, y: y + 2.1, w: pw - 0.6, h: 2.2, fontSize: 14, color: MUTED, lineSpacingMultiple: 1.25 }));
  }
  s.addText("مصدر نسبة 4.5%: GlobalData، أقساط 2023", ar({ x: M, y: 6.55, w: CW, h: 0.3, fontSize: 10, color: "8A97A8" }));
  pres.addSection({ title: "الحل" });
  // ================= 4. SOLUTION =================
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "الحل" });
  kicker(s, "الحل");
  s.addText("لذلك نقترح مرساة", { placeholder: "title" });
  s.addText("منصة وطنية تقوم على فكرة واحدة: أن نصنع مؤشر خطر سعودياً خاصاً بنا للمسارات البحرية، ببياناتنا ومنهجيتنا، لا بالاعتماد على تقييم لجنة الحرب المشتركة، ونضعه في موقع واحد يجمع لصاحب البضاعة ومالك السفينة الخطر والتكلفة والمؤمّنين السعوديين.",
    ar({ x: M, y: 1.6, w: CW, h: 0.95, fontSize: 15.5, color: MUTED, lineSpacingMultiple: 1.2 }));
  const ins = ["الحوادث الأمنية البحرية", "تتبع السفن (AIS) الوطني", "بيانات الموانئ السعودية", "أقساط ومطالبات المؤمّنين السعوديين"];
  const outs = ["درجة خطر لكل منطقة ومسار", "قسط تقديري للبضاعة وللسفينة", "مقارنة محايدة بين المؤمّنين", "تنبيه بأفضل وقت للتأمين"];
  s.addText("بياناتنا الوطنية", ar({ x: 9.2, y: 2.65, w: R - 9.2, h: 0.4, fontSize: 14, bold: true, color: TEAL }));
  s.addText("ما يحصل عليه المستخدم", ar({ x: M, y: 2.65, w: 3.7, h: 0.4, fontSize: 14, bold: true, color: TEAL }));
  ins.forEach((t, i) => {
    const y = 3.1 + i * 0.8;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 9.2, y, w: R - 9.2, h: 0.66, rectRadius: 0.1, fill: { color: WHITE }, line: { color: LINE } });
    s.addText(t, ar({ x: 9.35, y, w: R - 9.5, h: 0.66, fontSize: 14, color: NAVY, valign: "middle" }));
  });
  outs.forEach((t, i) => {
    const y = 3.1 + i * 0.8;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y, w: 3.7, h: 0.66, rectRadius: 0.1, fill: { color: ICE_BG }, line: { color: ICE_BG } });
    s.addText(t, ar({ x: M + 0.15, y, w: 3.4, h: 0.66, fontSize: 14, bold: true, color: TEAL_D, valign: "middle" }));
  });
  s.addShape(pres.shapes.OVAL, { x: 5.27, y: 2.95, w: 2.8, h: 2.8, fill: { color: NAVY }, line: { color: ICE, width: 4 } });
  s.addImage({ data: ANCHOR_ICE, x: 6.32, y: 3.25, w: 0.7, h: 0.7 });
  s.addText("مؤشر مرساة", ctr({ x: 5.27, y: 4.0, w: 2.8, h: 0.5, fontSize: 21, bold: true, color: WHITE }));
  s.addText("سعودي ومستقل", ctr({ x: 5.27, y: 4.47, w: 2.8, h: 0.4, fontSize: 14, color: ICE }));
  s.addShape(pres.shapes.LINE, { x: 8.2, y: 4.35, w: 0.85, h: 0, line: { color: TEAL, width: 2.5, beginArrowType: "triangle" } });
  s.addShape(pres.shapes.LINE, { x: 4.42, y: 4.35, w: 0.72, h: 0, line: { color: TEAL, width: 2.5, beginArrowType: "triangle" } });
  s.addText("مؤشر نصنعه نحن ونملكه، والتأكيد النهائي للسعر يبقى من شركة التأمين.", ar({ x: M, y: 6.45, w: CW, h: 0.35, fontSize: 13, bold: true, color: TEAL_D }));
  // ================= 4b. HULL vs CARGO =================
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "الحل" });
  kicker(s, "نوعا التأمين");
  s.addText("تأمين السفينة وتأمين البضاعة: مرساة تخدم الاثنين", { placeholder: "title" });
  s.addText("في الرحلة الواحدة يوجد تأمينان مختلفين، لكلٍّ منهما دافع وطريقة حساب. ومرساة تقدّم مساراً واضحاً لكل منهما.", ar({ x: M, y: 1.58, w: CW, h: 0.5, fontSize: 15, color: MUTED }));
  {
    const cw2 = (CW - 0.4) / 2, y0 = 2.2, hh = 3.6;
    const cols = [
      [lu.LuPackage, "تأمين البضاعة", "Cargo War Risk", [["من يدفع؟", "صاحب البضاعة: المصدّر أو المستورد"], ["كيف يُحسب؟", "نسبة من قيمة الشحنة"], ["ما يؤثر فيه", "نوع البضاعة ومسارها، وتنتقل إليه أيضاً تكلفة السفينة عبر أجرة الشحن"], ["في مرساة", "مقارنة عروض المؤمّنين لكل شحنة، وتنبيه بأفضل سعر"]]],
      [lu.LuShip, "تأمين جسم السفينة", "Hull War Risk", [["من يدفع؟", "مالك السفينة أو مشغّلها"], ["كيف يُحسب؟", "نسبة من قيمة السفينة لكل رحلة أو فترة خطر"], ["ما يؤثر فيه", "منطقة العبور، وقيمة السفينة، ومدة بقائها في منطقة الخطر"], ["في مرساة", "قسط تقديري لكل رحلة، ومؤمّنو الهياكل المشاركون في الوعاء الوطني"]]],
    ];
    for (let k = 0; k < 2; k++) {
      const x = R - cw2 - k * (cw2 + 0.4), c = cols[k];
      card(s, x, y0, cw2, hh);
      await bubble(s, c[0], x + cw2 - 1.05, y0 + 0.25, 0.8, k === 0 ? TEAL : NAVY, WHITE);
      s.addText(c[1], ar({ x: x + 0.3, y: y0 + 0.25, w: cw2 - 1.5, h: 0.48, fontSize: 21, bold: true, color: NAVY }));
      s.addText(c[2], ar({ x: x + 0.3, y: y0 + 0.7, w: cw2 - 1.5, h: 0.3, fontSize: 11, color: MUTED }));
      c[3].forEach((r, i) => {
        const y = y0 + 1.2 + i * 0.58;
        if (i) s.addShape(pres.shapes.LINE, { x: x + 0.3, y: y - 0.04, w: cw2 - 0.6, h: 0, line: { color: LINE, width: 0.75 } });
        s.addText(r[0], ar({ x: x + cw2 - 1.65, y, w: 1.35, h: 0.52, fontSize: 12.5, bold: true, color: i === 3 ? TEAL_D : MUTED, valign: "middle" }));
        s.addText(r[1], ar({ x: x + 0.3, y, w: cw2 - 2.0, h: 0.52, fontSize: 12.5, bold: i === 3, color: i === 3 ? TEAL_D : INK, valign: "middle" }));
      });
    }
  }
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y: 5.98, w: CW, h: 0.75, rectRadius: 0.12, fill: { color: NAVY }, line: { color: NAVY } });
  s.addText([{ text: "مؤشر واحد للمسارين:  ", options: { bold: true, color: ICE } }, { text: "درجة خطر المنطقة نفسها تُترجم إلى قسط للسفينة وقسط للبضاعة. وأسعار تأمين البضاعة لا تُنشر علناً، وهذه الفجوة التي تسدها مرساة.", options: { color: WHITE } }],
    ar({ x: M + 0.3, y: 5.98, w: CW - 0.6, h: 0.75, fontSize: 13.5, valign: "middle" }));
  // ================= 11. WHY A SAUDI INDEX (+ knowledge pyramid) =================
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "الحل" });
  kicker(s, "لماذا مؤشر سعودي؟");
  s.addText("نصنع مؤشر خطر سعودياً.. لماذا؟", { placeholder: "title" });
  {
    card(s, M, 1.8, 5.2, 4.85);
    s.addText("هرم المعرفة", ctr({ x: M, y: 1.92, w: 5.2, h: 0.42, fontSize: 16, bold: true, color: NAVY }));
    // triangle split into 4 bands
    const TW = 4.4, TH = 3.3, tx = M + (5.2 - TW) / 2, ty = 2.4, gap = 0.05;
    const cols = ["7FCFD7", "47B4C0", "1A93A0", "0E7C86"];
    const px = 100, Wp = TW * px, Hp = TH * px;
    let polys = "";
    for (let k = 0; k < 4; k++) {
      const y1 = (k / 4) * Hp + (k ? gap * px / 2 : 0), y2 = ((k + 1) / 4) * Hp - (k < 3 ? gap * px / 2 : 0);
      const half = (y) => (y / Hp) * Wp / 2;
      polys += `<polygon fill="#${cols[k]}" points="${Wp / 2 - half(y1)},${y1} ${Wp / 2 + half(y1)},${y1} ${Wp / 2 + half(y2)},${y2} ${Wp / 2 - half(y2)},${y2}"/>`;
    }
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${Wp * 3}" height="${Hp * 3}" viewBox="0 0 ${Wp} ${Hp}">${polys}</svg>`;
    const png = await sharp(Buffer.from(svg)).png().toBuffer();
    s.addImage({ data: "image/png;base64," + png.toString("base64"), x: tx, y: ty, w: TW, h: TH, altText: "هرم المعرفة: البيانات ثم المعلومات ثم المعرفة ثم الحكمة" });
    const labels = ["الحكمة", "المعرفة", "المعلومات", "البيانات"];
    labels.forEach((l, k) => {
      const yc = ty + (k + 0.5) * TH / 4, wc = TW * (k + 0.5) / 4;
      s.addText(l, ctr({ x: tx + TW / 2 - Math.max(wc, 1.0) / 2, y: yc - 0.22 + (k === 0 ? 0.12 : 0), w: Math.max(wc, 1.0), h: 0.44, fontSize: k === 0 ? 12 : 15, bold: true, color: k === 0 ? NAVY : WHITE, valign: "middle" }));
    });
    s.addText("المعلومات حاضرة في كل شيء: في كل خطوة وكل قرار. ومن يملك البيانات يملك القرار.", ctr({ x: M + 0.3, y: 5.85, w: 4.6, h: 0.7, fontSize: 13, bold: true, color: TEAL_D }));
    const why = [
      [lu.LuDatabase, "نملك بياناتنا", "البيانات أساس الهرم، ومنها تُبنى المعلومات والمعرفة والقرار. ومع سنوات التشغيل تتراكم لدى المملكة بيانات متميزة عن مناطقها وعن النقل البحري تحديداً، وتكون كلها في منصة واحدة."],
      [lu.LuFlag, "استقلال لوجستي", "قدرة وطنية تقلل اعتماد خطة المملكة اللوجستية على المراجع الخارجية."],
      [lu.LuEarth, "مرجع عالمي مستقبلاً", "يبدأ بالمناطق المحيطة بالمملكة ويتوسع، وقد تعتمد عليه دول أخرى."],
      [lu.LuTrainFront, "يكمل منظومة النقل", "جزء مما تبنيه المملكة من موانئ وقطارات وبنية لوجستية نحو المركز العالمي."],
    ];
    const hs = [1.6, 1.05, 1.05, 1.05];
    let y = 1.8;
    for (let i = 0; i < 4; i++) {
      const x = 6.2;
      await bubble(s, why[i][0], R - 0.8, y + 0.05, 0.8);
      s.addText(why[i][1], ar({ x, y, w: R - x - 1.0, h: 0.42, fontSize: 17, bold: true, color: NAVY }));
      s.addText(why[i][2], ar({ x, y: y + 0.42, w: R - x - 1.0, h: hs[i] - 0.5, fontSize: 13.5, color: MUTED }));
      y += hs[i] + 0.1;
    }
  }
  // ================= 5. STRATEGY =================
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "الحل" });
  kicker(s, "الاستراتيجية");
  s.addText("نبدأ وطنياً، ثم نتوسع إقليمياً وعالمياً", { placeholder: "title" });
  const strat = [
    [lu.LuLayers, "طبقة بيانات وطنية", "تملكها المملكة: تجمع بيانات الموانئ والحوادث وحركة السفن والمؤمّنين المحليين، وتزداد قيمتها مع كل رحلة."],
    [lu.LuGitCompare, "منهجية مستقلة", "نصنع مقياسنا الخاص بمنهجية شفافة، ولا نعتمد على تقييم لجنة الحرب المشتركة في لندن."],
    [lu.LuEarth, "من الإقليم إلى العالم", "يبدأ بالمناطق المحيطة بالمملكة، ثم يتوسع ليصبح مرجعاً تستفيد منه دول أخرى."],
  ];
  const sw3 = (CW - 2 * 0.6) / 3;
  for (let i = 0; i < 3; i++) {
    const x = R - sw3 - i * (sw3 + 0.6), y = 1.85;
    card(s, x, y, sw3, 3.6);
    badge(s, i + 1, x + sw3 - 0.72, y + 0.3);
    await bubble(s, strat[i][0], x + 0.3, y + 0.3, 0.9, i === 2 ? TEAL : ICE_BG, i === 2 ? WHITE : TEAL_D);
    s.addText(strat[i][1], ar({ x: x + 0.3, y: y + 1.45, w: sw3 - 0.6, h: 0.55, fontSize: 21, bold: true, color: NAVY }));
    s.addText(strat[i][2], ar({ x: x + 0.3, y: y + 2.1, w: sw3 - 0.6, h: 1.5, fontSize: 15, color: MUTED, lineSpacingMultiple: 1.2 }));
    if (i < 2) s.addShape(pres.shapes.LINE, { x: x - 0.5, y: y + 1.85, w: 0.4, h: 0, line: { color: TEAL, width: 2.5, endArrowType: "triangle" }, flipH: true });
  }
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y: 5.62, w: CW, h: 1.1, rectRadius: 0.12, fill: { color: NAVY }, line: { color: NAVY } });
  s.addText([{ text: "الهدف:  ", options: { bold: true, color: ICE } }, { text: "أن تملك المملكة قراءتها الوطنية للخطر. فما دامت رؤية 2030 تجعل المملكة مركزاً لوجستياً عالمياً، فالأجدر أن يكون لها مقياس خطر خاص بها، أساساً لتقدير الأقساط وقرارات الشحن والتأمين.", options: { color: WHITE } }], ar({ x: M + 0.3, y: 5.62, w: CW - 0.6, h: 1.1, fontSize: 15, valign: "middle", lineSpacingMultiple: 1.15 }));
  pres.addSection({ title: "الموقع" });
  // ================= 6. HOW THE WEBSITE WORKS =================
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "الموقع" });
  kicker(s, "موقع مرساة");
  s.addText("كيف يستخدم الشاحن مرساة؟ أربع خطوات", { placeholder: "title" });
  const steps = [
    [lu.LuRadar, "الرادار", "خريطة حية لمناطق الخطر وحركة السفن، ومسارات مقترحة لكل منها درجة خطر ومدة وقسط متوقع."],
    [lu.LuShieldCheck, "شركات التأمين", "كل مؤمّن سعودي: ما يغطيه للبضاعة والسفينة، وأنواع وثائقه، ومشاركته في الوعاء الوطني."],
    [lu.LuBellRing, "مشاريعي", "يحفظ الشاحن شحنته، ويصله تنبيه عند تغيّر القسط أو وصوله للسعر الذي حدده."],
    [lu.LuSparkles, "الذكاء الاصطناعي", "يقترح أفضل وقت للتأمين مع درجة ثقة، ويرشّح المؤمّن الأنسب مع سبب الترتيب."],
  ];
  for (let i = 0; i < 4; i++) {
    const x = R - pw - i * (pw + 0.3), y = 1.85;
    card(s, x, y, pw, 4.3);
    badge(s, i + 1, x + pw - 0.7, y + 0.28);
    await bubble(s, steps[i][0], x + pw / 2 - 0.55, y + 0.6, 1.1, i === 3 ? TEAL : ICE_BG, i === 3 ? WHITE : TEAL_D);
    s.addText(steps[i][1], ctr({ x: x + 0.2, y: y + 1.9, w: pw - 0.4, h: 0.5, fontSize: 19, bold: true, color: NAVY }));
    s.addText(steps[i][2], ar({ x: x + 0.3, y: y + 2.5, w: pw - 0.6, h: 1.65, fontSize: 14, color: MUTED, lineSpacingMultiple: 1.2 }));
  }
  s.addText("مناطق الخطر يحددها مؤشر مرساة من بيانات رسمية، والذكاء الاصطناعي للتوصية والتنبيه فقط.", ar({ x: M, y: 6.35, w: CW, h: 0.35, fontSize: 13, bold: true, color: TEAL_D }));
  // ================= 7-9. SCREENS =================
  async function screen(k, title, file, aspect, callouts, maxW = 8.55, maxH = 4.95) {
    s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "الموقع" });
    kicker(s, k);
    s.addText(title, { placeholder: "title" });
    let w = maxW, h = w / aspect; if (h > maxH) { h = maxH; w = h * aspect; }
    const ix = M + 0.06, iy = 1.7;
    shot(s, file, ix, iy, w, h, title);
    s.addText("لقطة من النموذج الأولي التجريبي لمرساة، والأسماء والأرقام توضيحية وليست من مصدر رسمي", ar({ x: ix, y: iy + h + 0.14, w, h: 0.3, fontSize: 10, color: "8A97A8" }));
    const cx = ix + w + 0.45, cw = R - cx;
    callouts.forEach((c, i) => {
      const y = 1.75 + i * 1.62;
      badge(s, i + 1, R - 0.42, y);
      s.addText(c[0], ar({ x: cx, y: y - 0.02, w: cw - 0.6, h: 0.5, fontSize: 16, bold: true, color: NAVY, valign: "middle" }));
      s.addText(c[1], ar({ x: cx, y: y + 0.52, w: cw - 0.6, h: 0.95, fontSize: 13.5, color: MUTED }));
    });
  }
  await screen("الشاشة 1 · الرادار", "الخطر وحركة السفن على مسارك، لحظة بلحظة", "search_row.png", 2544 / 1460, [
    ["تصفّح الرادار مباشرة", "خريطة حية لمناطق الخطر وحركة السفن، متاحة لأي زائر دون إدخال رحلة."],
    ["أدخل ميناءيك", "يختار الشاحن ميناء الانطلاق والوصول السعودي، ونوع الشحنة وتاريخ الإبحار."],
    ["قارن المسارات", "لكل مسار: أعلى درجة خطر، والمدة، والمسافة، وعدد مناطق الخطر التي يعبرها."],
  ]);
  await screen("الشاشة 2 · الأسعار", "كم سيكلف التأمين؟ ومتى أؤمّن؟", "prices_top.png", 2880 / 1880, [
    ["قسط البضاعة التقديري", "يكتب صاحب البضاعة قيمة الشحنة، فيرى أقل قسط ومتوسطه وتغيّره خلال 14 يوماً."],
    ["توصية التوقيت", "الذكاء الاصطناعي يقترح متى يؤمّن الشاحن، مع درجة ثقة واضحة."],
    ["تنبيه السعر", "يحفظ الشحنة ويحدد سعراً، فيصله إشعار عند نزول القسط."],
  ]);
  await screen("الشاشة 3 · شركات التأمين", "المؤمّنون السعوديون والمشاركون في الوعاء الوطني", "insurers_top.png", 2880 / 1340, [
    ["مؤمّنو الوعاء الوطني", "علامة واضحة على المؤمّنين المشاركين في الوعاء الوطني لتأمين مخاطر الحرب."],
    ["ما يغطيه كل مؤمّن", "أنواع الشحنات (حاويات، زيوت، سيارات، مبردة) والوثائق ونطاق استرشادي للقسط."],
    ["اطلب عرضاً رسمياً", "ينتقل الشاحن مباشرة إلى المؤمّن السعودي، ولا يوجد ترتيب مدفوع."],
  ]);
  // ================= 9b. VIDEO + QR =================
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "الموقع" });
  kicker(s, "جرّب مرساة بنفسك");
  s.addText("شاهد النموذج الأولي وهو يعمل", { placeholder: "title" });
  {
    const vw = 8.4, vh = vw * 9 / 16, vx = M + 0.06, vy = 1.75;
    const cover = "data:image/jpeg;base64," + require("fs").readFileSync(path.join(SHOTS, "demo_cover.jpeg")).toString("base64");
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: vx - 0.06, y: vy - 0.06, w: vw + 0.12, h: vh + 0.12, rectRadius: 0.08, fill: { color: WHITE }, line: { color: LINE },
      shadow: { type: "outer", color: "0B2545", opacity: 0.18, blur: 14, offset: 4, angle: 90 } });
    s.addMedia({ type: "video", path: path.join(SHOTS, "demo.mp4"), x: vx, y: vy, w: vw, h: vh, cover });
    s.addText("فيديو قصير يشرح طريقة استخدام منصة مرساة، من اختيار الميناء إلى طلب عرض التأمين", ar({ x: vx, y: vy + vh + 0.15, w: vw, h: 0.35, fontSize: 12, color: MUTED }));
    const qx = vx + vw + 0.45, qw = R - qx;
    card(s, qx, vy - 0.06, qw, vh + 0.12);
    const qs = 2.1;
    s.addImage({ path: path.join(SHOTS, "qr_canva.png"), x: qx + (qw - qs) / 2, y: vy + 0.2, w: qs, h: qs, altText: "رمز QR لفتح موقع مرساة" });
    s.addText("امسح الرمز وجرّب الموقع", ctr({ x: qx + 0.15, y: vy + 2.42, w: qw - 0.3, h: 0.4, fontSize: 15, bold: true, color: NAVY }));
    ["اختر ميناءيك وشاهد خطر كل مسار", "تفقّد أسعار التأمين وتغيّرها", "قارن المؤمّنين واطلب عرضاً"].forEach((t2, i) => {
      const y = vy + 2.9 + i * 0.42;
      badge(s, i + 1, qx + qw - 0.5, y + 0.04, 0.3);
      s.addText(t2, ar({ x: qx + 0.2, y, w: qw - 0.8, h: 0.38, fontSize: 12, color: INK, valign: "middle" }));
    });
  }
  pres.addSection({ title: "القيمة والأثر" });
  // ================= 12. TRACK 4 ANSWER =================
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "القيمة والأثر" });
  kicker(s, "سؤال المسار الرابع: كيف يمكن تقليل تكلفة التوصيل ووقته؟");
  s.addText("كيف تقلل مرساة تكلفة التوصيل ووقته؟", { placeholder: "title" });
  const tr = [
    [lu.LuReceipt, "تكلفة واضحة قبل القرار", "يقارن الشاحن قسط كل مسار وميناء في مكان واحد، فيعرف التكلفة الكلية قبل أن يشحن، لا أجرة الشحن وحدها."],
    [lu.LuTimer, "وقت أقل في البحث", "مقارنة موحدة بين المؤمّنين بدل التواصل مع كل شركة تأمين على حدة."],
    [lu.LuBellRing, "توقيت أذكى", "تنبيهات تتيح التخطيط لموعد التأمين، بدل قرار تحت ضغط مراجعة الأقساط كل 24 ساعة."],
    [lu.LuHandCoins, "أموال تبقى في المملكة", "تمكين المؤمّنين السعوديين يقلل الاعتماد على المؤمّنين العالميين، فتبقى أقساط التأمين داخل الاقتصاد الوطني."],
  ];
  const tw2 = (CW - 0.4) / 2;
  for (let i = 0; i < 4; i++) {
    const col = i % 2, row = Math.floor(i / 2);
    const x = R - tw2 - col * (tw2 + 0.4), y = 1.85 + row * 2.4;
    const hl = i === 3;
    card(s, x, y, tw2, 2.15, hl ? NAVY : WHITE, hl ? NAVY : LINE);
    await bubble(s, tr[i][0], x + tw2 - 1.15, y + 0.35, 0.85, hl ? TEAL : ICE_BG, hl ? WHITE : TEAL_D);
    s.addText(tr[i][1], ar({ x: x + 0.35, y: y + 0.35, w: tw2 - 1.7, h: 0.5, fontSize: 20, bold: true, color: hl ? WHITE : NAVY }));
    s.addText(tr[i][2], ar({ x: x + 0.35, y: y + 0.95, w: tw2 - 1.7, h: 1.05, fontSize: 14.5, color: hl ? SOFT : MUTED, lineSpacingMultiple: 1.15 }));
  }
  // ================= 10. VISION 2030 + POOL =================
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "القيمة والأثر" });
  kicker(s, "رؤية المملكة");
  s.addText("مرساة تخدم رؤية 2030، وتكمل قرار الوعاء الوطني", { placeholder: "title" });
  {
    const cw2 = (CW - 0.4) / 2, y0 = 1.8, hh = 4.75;
    let x = R - cw2;
    card(s, x, y0, cw2, hh);
    await bubble(s, lu.LuFlag, x + cw2 - 1.15, y0 + 0.35, 0.8, TEAL, WHITE);
    s.addText("رؤية 2030", ar({ x: x + 0.35, y: y0 + 0.35, w: cw2 - 1.65, h: 0.8, fontSize: 26, bold: true, color: NAVY, valign: "middle" }));
    s.addText("المملكة مركز لوجستي عالمي", ar({ x: x + 0.35, y: y0 + 1.3, w: cw2 - 0.7, h: 0.45, fontSize: 17, color: MUTED }));
    s.addText("40 مليون", ar({ x: x + 0.35, y: y0 + 1.8, w: cw2 - 0.7, h: 0.95, fontSize: 48, bold: true, color: TEAL }));
    s.addText("حاوية نمطية سنوياً عبر الموانئ السعودية بحلول 2030، مقابل 8.3 مليون في 2025.", ar({ x: x + 0.35, y: y0 + 2.75, w: cw2 - 0.7, h: 0.7, fontSize: 14, color: MUTED }));
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.25, y: y0 + 3.5, w: cw2 - 0.5, h: 1.05, rectRadius: 0.1, fill: { color: ICE_BG }, line: { color: ICE_BG } });
    s.addText([{ text: "دور مرساة: ", options: { bold: true, color: TEAL_D } }, { text: "وضوح كلفة التأمين يشجع الشاحنين على اختيار الموانئ والمياه السعودية.", options: { color: INK } }], ar({ x: x + 0.45, y: y0 + 3.5, w: cw2 - 0.9, h: 1.05, fontSize: 15, valign: "middle" }));
    x = M;
    card(s, x, y0, cw2, hh);
    await bubble(s, lu.LuLandmark, x + cw2 - 1.15, y0 + 0.35, 0.8, TEAL, WHITE);
    s.addText("الوعاء الوطني لتأمين مخاطر الحرب", ar({ x: x + 0.35, y: y0 + 0.35, w: cw2 - 1.65, h: 0.8, fontSize: 22, bold: true, color: NAVY, valign: "middle" }));
    s.addText("أُقرّ في سبتمبر 2026 بقيادة الشركة السعودية لإعادة التأمين، لدعم استمرارية التجارة وسلاسل الإمداد. ويستفيد منه المؤمَّن لهم عبر شركات التأمين المشاركة.", ar({ x: x + 0.35, y: y0 + 1.3, w: cw2 - 0.7, h: 1.6, fontSize: 15, color: MUTED, lineSpacingMultiple: 1.2 }));
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.25, y: y0 + 3.5, w: cw2 - 0.5, h: 1.05, rectRadius: 0.1, fill: { color: ICE_BG }, line: { color: ICE_BG } });
    s.addText([{ text: "دور مرساة: ", options: { bold: true, color: TEAL_D } }, { text: "واجهة رقمية تقود الشاحن إلى المؤمّنين المشاركين في الوعاء، وتكمله ولا تنافسه.", options: { color: INK } }], ar({ x: x + 0.45, y: y0 + 3.5, w: cw2 - 0.9, h: 1.05, fontSize: 15, valign: "middle" }));
  }
  s.addText("المصادر: عرب نيوز، Container News، ذا سعودي غازيت عبر Hellenic Shipping News", ar({ x: M, y: 6.65, w: CW, h: 0.3, fontSize: 10, color: "8A97A8" }));
  // ================= 14. COMPARISON =================
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "القيمة والأثر" });
  kicker(s, "المقارنة");
  s.addText("ماذا تضيف مرساة على الحلول الموجودة في السوق؟", { placeholder: "title" });
  s.addText("درسنا الحلول المتاحة: كلها قيّمة، لكنها عالمية أو تعالج جانباً واحداً، ولم نجد منصة سعودية تجمع خطر المسار وتكلفته ومقارنة المؤمّنين المحليين.", ar({ x: M, y: 1.6, w: CW, h: 0.6, fontSize: 14.5, color: MUTED }));
  {
    const hdr = (t) => ({ text: t, options: { bold: true, color: WHITE, fill: { color: NAVY }, fontSize: 13 } });
    const rows = [
      ["ShipWRITE من Canopius", "عرض وإصدار فوري لتأمين الحرب مع تتبع عالمي للسفن", "مقارنة بين عدة مؤمّنين سعوديين بدل مؤمّن واحد"],
      ["WarRates.com", "أسعار استرشادية لتأمين حرب الهياكل من عروض الأعضاء", "ربط السعر بمسار الشحنة وميناءيها ونوعها"],
      ["Insurwave مع Ambrey", "بيانات التهديدات البحرية لرؤية عالمية للمخاطر", "تركيز وطني وبيانات محلية للموانئ والمؤمّنين"],
      ["DP World لتأمين حرب البضائع", "وثيقة واحدة للبضائع من الباب إلى الباب", "مقارنة محايدة بين عدة مؤمّنين، لا منتج واحد"],
      ["منصات مقارنة التأمين المحلية", "وسطاء إلكترونيون مرخصون من هيئة التأمين", "مؤشر خطر سعودي وتقدير الأقساط بناءً عليه"],
      ["الوعاء الوطني لمخاطر الحرب", "آلية وطنية بقيادة الشركة السعودية لإعادة التأمين", "واجهة رقمية مكملة تقود الشاحن إلى المؤمّنين المشاركين"],
    ];
    const data = [[hdr("ما تضيفه مرساة"), hdr("ما يقدمه"), hdr("الحل القائم")]];
    rows.forEach((r, i) => {
      const fill = { color: i % 2 ? "F7F9FC" : WHITE };
      data.push([{ text: r[2], options: { bold: true, color: TEAL_D, fill } }, { text: r[1], options: { color: MUTED, fill } }, { text: r[0], options: { bold: true, color: NAVY, fill } }]);
    });
    s.addTable(data, { x: M, y: 2.35, w: CW, colW: [4.55, 4.55, CW - 9.1], fontFace: F, fontSize: 13, align: "right", valign: "middle", rtlMode: true, border: { type: "solid", pt: 0.75, color: LINE }, margin: [0.06, 0.15, 0.06, 0.15], rowH: 0.6 });
  }
  // ================= 17. IMPACT =================
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "القيمة والأثر" });
  kicker(s, "الأثر المتوقع");
  s.addText("تمكين المؤمّن السعودي، وقرار أوضح للشاحن", { placeholder: "title" });
  const imp = [
    [lu.LuShieldCheck, "المؤمّنون السعوديون", "قناة رقمية تصلهم بعملاء جدد، وتقلل اعتماد السوق على المؤمّنين العالميين."],
    [lu.LuContainer, "الشاحنون وملاك السفن", "مقارنة المؤمّنين وتكلفة كل مسار للبضاعة والسفينة، وتنبيه بأفضل وقت."],
    [lu.LuLandmark, "الوعاء الوطني", "واجهة رقمية تسهّل وصول المستفيدين إلى المؤمّنين المشاركين."],
    [lu.LuDatabase, "الدولة", "بيانات وطنية عن المسارات والمخاطر تدعم قرارات الجهات المنظمة."],
  ];
  for (let i = 0; i < 4; i++) {
    const x = R - pw - i * (pw + 0.3), y = 1.8, hl = i === 0;
    card(s, x, y, pw, 2.55, hl ? NAVY : WHITE, hl ? NAVY : LINE);
    await bubble(s, imp[i][0], x + pw - 1.0, y + 0.3, 0.7, hl ? TEAL : ICE_BG, hl ? WHITE : TEAL_D);
    s.addText(imp[i][1], ar({ x: x + 0.2, y: y + 1.1, w: pw - 0.4, h: 0.45, fontSize: 15, bold: true, color: hl ? WHITE : NAVY }));
    s.addText(imp[i][2], ar({ x: x + 0.3, y: y + 1.55, w: pw - 0.6, h: 0.9, fontSize: 13, color: hl ? SOFT : MUTED }));
  }
  // opportunity
  const bw2 = (CW - 0.3) / 2;
  card(s, R - bw2, 4.6, bw2, 1.95, ICE_BG, ICE_BG, false);
  s.addText("فرصة نمو كبيرة", ar({ x: R - bw2 + 0.3, y: 4.75, w: bw2 - 0.6, h: 0.4, fontSize: 14, bold: true, color: TEAL_D }));
  s.addText("29.8 مليار ريال", ar({ x: R - bw2 + 0.3, y: 5.12, w: bw2 - 0.6, h: 0.7, fontSize: 32, bold: true, color: NAVY }));
  s.addText("أقساط سوق التأمين السعودي في الربع الأول 2026 بنمو 14.4%، بينما التأمين البحري والنقل حصته صغيرة: مساحة لتمكين المؤمّن السعودي.", ar({ x: R - bw2 + 0.3, y: 5.82, w: bw2 - 0.6, h: 0.65, fontSize: 12.5, color: INK }));
  // KPIs
  card(s, M, 4.6, bw2, 1.95);
  s.addText("نقيس النجاح بـ", ar({ x: M + 0.3, y: 4.75, w: bw2 - 0.6, h: 0.4, fontSize: 14, bold: true, color: TEAL_D }));
  const kp = ["عدد شركات التأمين المشاركة", "الموانئ والمسارات المغطاة", "زمن الوصول لمقارنة موحدة", "المستخدمون النشطون", "طلبات العروض للمؤمّنين المحليين"];
  const chk = await icon(lu.LuCircleCheck, TEAL);
  kp.forEach((k, i) => {
    const col = i < 3 ? 0 : 1, row = i < 3 ? i : i - 3, cwk = (bw2 - 0.6) / 2;
    const x = M + 0.3 + (1 - col) * cwk, y = 5.2 + row * 0.42;
    s.addImage({ data: chk, x: x + cwk - 0.28, y: y + 0.06, w: 0.22, h: 0.22 });
    s.addText(k, ar({ x, y, w: cwk - 0.38, h: 0.36, fontSize: 12.5, valign: "middle" }));
  });
  s.addText("المصادر: هيئة التأمين (تقرير الربع الأول 2026)، GlobalData", ar({ x: M, y: 6.66, w: CW, h: 0.28, fontSize: 10, color: "8A97A8" }));
  pres.addSection({ title: "التنفيذ" });
  // ================= T1. TIMELINE (overlapping Gantt) =================
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "التنفيذ" });
  kicker(s, "الجدول الزمني");
  s.addText("خمس سنوات بمراحل متداخلة، ولكل مرحلة بوابة مراجعة", { placeholder: "title" });
  {
    const lw = 3.3, gx1 = R - lw - 0.15, gx0 = M + 0.15, gw = gx1 - gx0, mx = (m) => gx1 - (m / 60) * gw;
    const top = 1.65;
    card(s, M, top, CW, 4.95);
    // year bands + month ticks
    for (let yv = 1; yv <= 5; yv++) {
      const x = mx(yv * 12), w = gw / 5;
      if (yv % 2) s.addShape(pres.shapes.RECTANGLE, { x, y: top + 0.55, w, h: 4.25, fill: { color: "F7F9FC" }, line: { color: "F7F9FC" } });
      s.addText(`السنة ${yv}`, ctr({ x, y: top + 0.12, w, h: 0.3, fontSize: 12, bold: true, color: NAVY }));
    }
    for (let m = 6; m < 60; m += 6) s.addText(String(m), ctr({ x: mx(m) - 0.25, y: top + 0.38, w: 0.5, h: 0.2, fontSize: 8.5, color: "8A97A8" }));
    s.addText("الشهر", ar({ x: gx1 + 0.15, y: top + 0.36, w: lw - 0.3, h: 0.22, fontSize: 9, color: "8A97A8" }));
    const rows = [
      ["صنع مقياس الخطر", 0, 15, TEAL, [[12, "الإصدار الأول للمؤشر"]]],
      ["الاتفاق مع المؤمّنين", 8, 18, TEAL_D, [[18, "اتفاقيتان على الأقل"]]],
      ["بناء المنصة وتأمينها", 12, 30, NAVY2, [[30, "اجتياز اختبار الاختراق"]]],
      ["إطلاق تجريبي", 27, 36, ORANGE, [[36, "تقييم نتائج التجربة"]]],
      ["الإطلاق الرسمي والتوسع", 36, 60, NAVY, [[37, "إطلاق رسمي"], [54, "بدء التوسع الإقليمي"]]],
      ["تحديث المؤشر ومراجعته", 12, 60, "7FCFD7", [[24, "مراجعة سنوية"], [36, ""], [48, ""]]],
    ];
    rows.forEach((r, i) => {
      const y = top + 0.68 + i * 0.69;
      badge(s, i < 5 ? i + 1 : "∞", R - 0.55, y + 0.06, 0.34, r[3]);
      s.addText(r[0], ar({ x: gx1 + 0.2, y: y + 0.02, w: lw - 0.75, h: 0.4, fontSize: 13, bold: true, color: NAVY, valign: "middle" }));
      s.addText(`الأشهر ${r[1] + 1}–${r[2]}`, ar({ x: gx1 + 0.2, y: y + 0.38, w: lw - 0.75, h: 0.22, fontSize: 9.5, color: MUTED }));
      const x1 = mx(r[2]), x2 = mx(r[1]);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x1, y: y + 0.08, w: x2 - x1, h: 0.3, rectRadius: 0.15, fill: { color: r[3] }, line: { color: r[3] } });
      r[4].forEach((ms) => {
        const dx = mx(ms[0]);
        s.addShape(pres.shapes.DIAMOND, { x: dx - 0.13, y: y + 0.03, w: 0.26, h: 0.4, fill: { color: AMBER }, line: { color: WHITE, width: 1 } });
        if (ms[1]) {
          const lx = Math.max(gx0, Math.min(dx - 0.9, gx1 - 1.8));
          s.addText(ms[1], ctr({ x: lx, y: y + 0.42, w: 1.8, h: 0.22, fontSize: 9.5, bold: true, color: "8A5A00" }));
        }
      });
    });
  }
  s.addText("◆ بوابة مراجعة لكل مرحلة. المراحل تتداخل، فنبدأ العمل على المرحلة التالية دون انتظار اكتمال السابقة.", ar({ x: M, y: 6.68, w: CW, h: 0.28, fontSize: 11, color: MUTED }));

  // ================= T2. OWNING ENTITIES =================
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "التنفيذ" });
  kicker(s, "الجهات المالكة");
  s.addText("مَن يتولى منصة مرساة؟", { placeholder: "title" });
  s.addText("نتوقع أن تتولى المنصةَ أربع جهات، لكل منها دور واضح:", ar({ x: M, y: 1.55, w: CW, h: 0.4, fontSize: 15, color: MUTED }));
  {
    const ents = [
      [lu.LuLandmark, "وزارة النقل والخدمات اللوجستية", "المظلة والإشراف", "تتبع لها المنصة ضمن الاستراتيجية الوطنية للنقل والخدمات اللوجستية، وتعتمد مؤشر مرساة مرجعاً وطنياً."],
      [lu.LuShip, "الهيئة العامة للنقل", "مالك المنصة ومشغّلها", "الجهة المنظِّمة للنقل البحري، تملك المنصة وتشغّلها مع شريك تقني وطني."],
      [lu.LuShieldCheck, "هيئة التأمين", "منظِّم الجانب التأميني", "تنظّم شركات التأمين وتشرف على الوعاء الوطني لمخاطر الحرب، وتعتمد عرض الأسعار والمقارنة في المنصة."],
      [lu.LuAnchor, "الهيئة العامة للموانئ (موانئ)", "شريك البيانات", "تزوّد المؤشر ببيانات الموانئ وحركة السفن، وتستفيد من جذب الشاحنين إلى الموانئ السعودية."],
    ];
    const cw4 = (CW - 3 * 0.3) / 4, y0 = 2.1, hh = 4.15;
    for (let i = 0; i < 4; i++) {
      const x = R - cw4 - i * (cw4 + 0.3), hl = i === 0, e = ents[i];
      card(s, x, y0, cw4, hh, hl ? NAVY : WHITE, hl ? NAVY : LINE);
      await bubble(s, e[0], x + cw4 / 2 - 0.55, y0 + 0.35, 1.1, hl ? TEAL : ICE_BG, hl ? WHITE : TEAL_D);
      s.addText(e[1], ctr({ x: x + 0.2, y: y0 + 1.6, w: cw4 - 0.4, h: 0.65, fontSize: 16, bold: true, color: hl ? WHITE : NAVY, valign: "middle" }));
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.45, y: y0 + 2.32, w: cw4 - 0.9, h: 0.38, rectRadius: 0.19, fill: { color: hl ? NAVY2 : ICE_BG }, line: { color: hl ? NAVY2 : ICE_BG } });
      s.addText(e[2], ctr({ x: x + 0.45, y: y0 + 2.32, w: cw4 - 0.9, h: 0.38, fontSize: 12.5, bold: true, color: hl ? ICE : TEAL_D, valign: "middle" }));
      s.addText(e[3], ctr({ x: x + 0.25, y: y0 + 2.85, w: cw4 - 0.5, h: 1.15, fontSize: 12.5, color: hl ? SOFT : MUTED }));
    }
  }
  s.addText([{ text: "وبالتعاون مع: ", options: { bold: true, color: NAVY } }, { text: "الشركة السعودية لإعادة التأمين (الوعاء الوطني)، وحرس الحدود (بيانات الحوادث البحرية)، وشركات التأمين السعودية.", options: { color: MUTED } }],
    ar({ x: M, y: 6.45, w: CW, h: 0.35, fontSize: 13 }));

  // ================= 13b. COST =================
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "التنفيذ" });
  kicker(s, "التكلفة التقديرية");
  s.addText("كم تكلّف مرساة؟ نحو 27 مليون ريال على خمس سنوات", { placeholder: "title" });
  {
    const items = [
      ["بناء مؤشر الخطر", "بيانات ومنهجية ومراجعة مستقلة", 5.0, TEAL],
      ["تطوير المنصة", "تصميم وبرمجة وربط المؤمّنين", 6.0, TEAL_D],
      ["الأمن السيبراني وحماية البيانات", "تفاصيله في البطاقة المجاورة", 5.0, RED],
      ["البنية السحابية داخل المملكة", "استضافة وتخزين ونسخ احتياطي", 2.5, NAVY2],
      ["تراخيص البيانات والشراكات", "مزود تتبع مؤقت، وقانوني وتعاقدي", 2.5, "6E8DB3"],
      ["التشغيل والدعم", "سنتان بعد الإطلاق", 3.5, NAVY],
      ["احتياطي للطوارئ", "نحو 10% من الإجمالي", 2.5, "A0AEC0"],
    ];
    const lw = 7.35, lx = R - lw, y0 = 1.65;
    card(s, lx, y0, lw, 5.05);
    s.addText("27 مليون ريال", ar({ x: lx + 0.3, y: y0 + 0.15, w: 3.6, h: 0.65, fontSize: 30, bold: true, color: NAVY, valign: "middle" }));
    s.addText("إجمالي تقديري لبناء المنصة وتشغيلها خلال 5 سنوات", ar({ x: lx + 3.9, y: y0 + 0.15, w: lw - 4.2, h: 0.65, fontSize: 12, color: MUTED, valign: "middle" }));
    const maxV = 6.0, barMax = 2.2;
    items.forEach((it, i) => {
      const y = y0 + 0.98 + i * 0.57, hl = i === 2;
      if (hl) s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: lx + 0.12, y: y - 0.04, w: lw - 0.24, h: 0.54, rectRadius: 0.06, fill: { color: "FDECEC" }, line: { color: "FDECEC" } });
      s.addText(it[0], ar({ x: lx + 3.6, y, w: lw - 3.9, h: 0.27, fontSize: 13, bold: true, color: NAVY }));
      s.addText(it[1], ar({ x: lx + 3.6, y: y + 0.25, w: lw - 3.9, h: 0.22, fontSize: 10, color: MUTED }));
      const bw = (it[2] / maxV) * barMax;
      s.addShape(pres.shapes.RECTANGLE, { x: lx + 3.45 - bw, y: y + 0.12, w: bw, h: 0.24, fill: { color: it[3] }, line: { color: it[3] } });
      s.addText(`${it[2].toFixed(1)} م`, ar({ x: lx + 0.25, y: y + 0.05, w: 0.85, h: 0.38, fontSize: 12.5, bold: true, color: INK, valign: "middle" }));
    });
    // security card
    const sw2 = lx - M - 0.3;
    card(s, M, y0, sw2, 5.05, NAVY, NAVY);
    await bubble(s, lu.LuShieldCheck, M + sw2 - 1.0, y0 + 0.25, 0.75, TEAL, WHITE);
    s.addText("الأمن أولاً", ar({ x: M + 0.3, y: y0 + 0.25, w: sw2 - 1.4, h: 0.4, fontSize: 19, bold: true, color: WHITE }));
    s.addText("5 ملايين ريال، نحو 19% من الميزانية", ar({ x: M + 0.3, y: y0 + 0.65, w: sw2 - 1.4, h: 0.35, fontSize: 12.5, bold: true, color: ICE }));
    s.addText("لأن بيانات المخاطر والأسعار والمستخدمين حساسة، خصصنا جزءاً كبيراً من الميزانية لحمايتها:", ar({ x: M + 0.3, y: y0 + 1.15, w: sw2 - 0.6, h: 0.62, fontSize: 12.5, color: SOFT }));
    const sec = ["برمجة آمنة (Secure SDLC) ومراجعة الشيفرة", "اختبارات اختراق دورية قبل الإطلاق وبعده", "تشفير البيانات، وإدارة الهويات والصلاحيات", "مركز عمليات أمنية (SOC) على مدار الساعة", "استضافة داخل المملكة، والامتثال لضوابط الهيئة الوطنية للأمن السيبراني ونظام حماية البيانات الشخصية"];
    const ck2 = await icon(lu.LuCircleCheck, ICE);
    let yy = y0 + 1.85;
    sec.forEach((t2, i) => {
      const hh2 = i === 4 ? 0.82 : 0.52;
      s.addImage({ data: ck2, x: M + sw2 - 0.55, y: yy + 0.08, w: 0.26, h: 0.26 });
      s.addText(t2, ar({ x: M + 0.3, y: yy, w: sw2 - 0.95, h: hh2, fontSize: 12.5, color: WHITE }));
      yy += hh2 + 0.06;
    });
  }
  s.addText("تقدير مبدئي للتخطيط يُراجع عند كل بوابة. الافتراضات: فريق محلي، واستضافة سحابية داخل المملكة، وتشغيل سنتين بعد الإطلاق.", ar({ x: M, y: 6.78, w: CW, h: 0.25, fontSize: 9.5, color: "8A97A8" }));
  // ================= 15. CHALLENGES =================
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "التنفيذ" });
  kicker(s, "التحديات");
  s.addText("التحديات المتوقعة، وكيف نعالجها", { placeholder: "title" });
  const shieldI = await icon(lu.LuShieldCheck, TEAL);
  const warnI = await icon(lu.LuTriangleAlert, "C2410C");
  // biggest challenge: JWC
  card(s, M, 1.75, CW, 1.85, "FFF4EF", "F6C9B6");
  s.addText("التحدي الأكبر", ar({ x: R - 2.3, y: 1.92, w: 2.0, h: 0.35, fontSize: 12, bold: true, color: "C2410C" }));
  s.addImage({ data: warnI, x: R - 0.62, y: 2.33, w: 0.34, h: 0.34 });
  s.addText("احتكار لجنة الحرب المشتركة (JWC) لتقييم الخطر", ar({ x: M + 0.3, y: 2.27, w: CW - 1.05, h: 0.45, fontSize: 18, bold: true, color: NAVY }));
  s.addText("هي المنافس الأكبر: مرجع لندني تعتمد عليه أغلب الدول وشركات التأمين، وأي مؤشر جديد سيُقارَن به.", ar({ x: M + 0.3, y: 2.72, w: CW - 1.05, h: 0.35, fontSize: 13.5, color: MUTED }));
  s.addImage({ data: shieldI, x: R - 0.62, y: 3.12, w: 0.32, h: 0.32 });
  s.addText("نبدأ بسوقنا ومساراتنا، ونكسب الثقة بمنهجية شفافة ومراجعة مستقلة، ثم نتوسع مستندين إلى مكانة المملكة اللوجستية.", ar({ x: M + 0.3, y: 3.08, w: CW - 1.05, h: 0.45, fontSize: 13.5, bold: true, color: TEAL_D }));
  const rk = [
    ["تردد شركات التأمين", "المنصة قناة لعملاء جدد، ونبدأ باتفاقيات تجريبية محدودة."],
    ["تغيّر الأسعار بسرعة", "يُعرض السعر تقديرياً مع وقته وصلاحيته، والتأكيد النهائي من المؤمّن."],
    ["حساسية البيانات وخصوصيتها", "صلاحيات مقيدة، وحوكمة وفق نظام حماية البيانات الشخصية."],
    ["تحيّز التوصيات", "نعرض سبب كل ترتيب للمستخدم، ولا يوجد ترتيب مدفوع."],
  ];
  const rw = (CW - 0.35) / 2;
  rk.forEach((r, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = R - rw - col * (rw + 0.35), y = 3.82 + row * 1.42;
    card(s, x, y, rw, 1.25);
    s.addImage({ data: warnI, x: x + rw - 0.62, y: y + 0.2, w: 0.3, h: 0.3 });
    s.addText(r[0], ar({ x: x + 0.3, y: y + 0.14, w: rw - 1.05, h: 0.42, fontSize: 16, bold: true, color: NAVY }));
    s.addImage({ data: shieldI, x: x + rw - 0.62, y: y + 0.72, w: 0.3, h: 0.3 });
    s.addText(r[1], ar({ x: x + 0.3, y: y + 0.62, w: rw - 1.05, h: 0.5, fontSize: 13.5, color: MUTED, valign: "middle" }));
  });
  pres.addSection({ title: "الختام" });
  // ================= 18. TEAM =================
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "الختام" });
  kicker(s, "الفريق");
  s.addText("فريق مرساة", { placeholder: "title" });
  const team = [[lu.LuUserRound, "قائد الفريق", "", "سارة أحمد العرادي"], [lu.LuBriefcase, "المجال والأعمال", "", "يارا يحيى عثمان"], [lu.LuCode, "التقنية والبيانات", "", "البتول باقيس"], [lu.LuPalette, "التصميم وتجربة المستخدم", "", "رازان باقيس"]];
  for (let i = 0; i < 4; i++) {
    const x = R - pw - i * (pw + 0.3), y = 1.85;
    card(s, x, y, pw, 3.9);
    await bubble(s, team[i][0], x + pw / 2 - 0.75, y + 0.45, 1.5, i === 0 ? NAVY : ICE_BG, i === 0 ? ICE : TEAL_D);
    s.addText(team[i][3], ctr({ x: x + 0.2, y: y + 2.3, w: pw - 0.4, h: 0.55, fontSize: 18, bold: true, color: NAVY }));
    s.addText(team[i][1], ctr({ x: x + 0.2, y: y + 2.9, w: pw - 0.4, h: 0.45, fontSize: 14, bold: true, color: TEAL }));
  }
  // ================= 19. CLOSING =================
  s = pres.addSlide({ masterName: "COVER", sectionTitle: "الختام" });
  shot(s, "map.png", M + 0.06, 0.95, 5.3 * 1736 / 1460, 5.3, "خريطة رادار الخطر");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: R - 1.1, y: 1.0, w: 1.1, h: 1.1, rectRadius: 0.2, fill: { color: NAVY }, line: { color: NAVY } });
  s.addImage({ data: ANCHOR_ICE, x: R - 0.88, y: 1.22, w: 0.66, h: 0.66 });
  s.addText("مرساة", ar({ x: 7.3, y: 2.25, w: R - 7.3, h: 1.2, fontSize: 64, bold: true, color: NAVY }));
  s.addText("قراءة سعودية للخطر البحري، وقرار تأمين أوضح", ar({ x: 7.3, y: 3.5, w: R - 7.3, h: 0.9, fontSize: 22, bold: true, color: TEAL }));
  const close = ["مؤشر خطر وطني نصنعه ونملكه", "مؤمّن سعودي أقوى، وأقساط تبقى في الاقتصاد الوطني", "خطوة نحو مركز لوجستي عالمي، انسجاماً مع رؤية 2030"];
  const ck = await icon(lu.LuCircleCheck, TEAL);
  close.forEach((c, i) => {
    const y = 4.55 + i * 0.5;
    s.addImage({ data: ck, x: R - 0.3, y: y + 0.07, w: 0.28, h: 0.28 });
    s.addText(c, ar({ x: 7.3, y, w: R - 7.3 - 0.45, h: 0.42, fontSize: 15, color: INK, valign: "middle" }));
  });
  s.addText("شكراً لكم", ar({ x: 7.3, y: 6.2, w: R - 7.3, h: 0.6, fontSize: 24, bold: true, color: NAVY }));

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
