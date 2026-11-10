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
  // ================= 1. COVER =================
  pres.addSection({ title: "البداية" });
  s = pres.addSlide({ masterName: "COVER", sectionTitle: "البداية" });
  shot(s, "map.png", M + 0.06, 0.95, 5.3 * 1736 / 1460, 5.3, "خريطة رادار الخطر من النموذج الأولي");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: R - 1.1, y: 1.0, w: 1.1, h: 1.1, rectRadius: 0.2, fill: { color: NAVY }, line: { color: NAVY } });
  s.addImage({ data: ANCHOR_ICE, x: R - 0.88, y: 1.22, w: 0.66, h: 0.66 });
  s.addText("مرساة", ar({ x: 7.3, y: 2.25, w: R - 7.3, h: 1.3, fontSize: 72, bold: true, color: NAVY }));
  s.addText("منصة وطنية للتأمين البحري وتقييم مخاطر المسارات", ar({ x: 7.3, y: 3.6, w: R - 7.3, h: 0.9, fontSize: 22, bold: true, color: TEAL }));
  s.addText("نصنع مؤشر خطر سعودياً للمسارات البحرية، ونجمع للشاحن الخطر والتكلفة وعروض المؤمّنين السعوديين في مكان واحد.", ar({ x: 7.3, y: 4.55, w: R - 7.3, h: 0.95, fontSize: 15, color: MUTED, lineSpacingMultiple: 1.15 }));
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.3, y: 5.75, w: R - 7.3, h: 0.5, rectRadius: 0.25, fill: { color: ICE_BG }, line: { color: ICE_BG } });
  s.addText("نقل ثون 5  ·  المسار الرابع: الخدمات اللوجستية وسلاسل الإمداد", ctr({ x: 7.35, y: 5.75, w: R - 7.4, h: 0.5, fontSize: 11.5, bold: true, color: TEAL_D, valign: "middle" }));

  // ================= 2. PROBLEM: geopolitics + numbers =================
  pres.addSection({ title: "المشكلة" });
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "المشكلة" });
  kicker(s, "المشكلة");
  s.addText("التوترات الجيوسياسية رفعت تأمين الشحن البحري بشكل حاد", { placeholder: "title" });
  s.addText("حين تعبر السفينة منطقة توتر أمني مثل البحر الأحمر أو مضيق هرمز، تدفع قسطاً إضافياً اسمه «تأمين مخاطر الحرب». ومع تصاعد التوترات في المنطقة، وسّعت لجنة الحرب المشتركة في لندن المنطقة عالية الخطر في البحر الأحمر (يوليو 2026)، فقفزت الأقساط.",
    ar({ x: M, y: 1.6, w: CW, h: 0.95, fontSize: 15.5, color: MUTED, lineSpacingMultiple: 1.2 }));
  {
    const gapA = 0.75, gapB = 0.3, nw = (CW - gapA - gapB) / 3, y = 3.25, ch = 2.35;
    const xs = [R - nw, R - 2 * nw - gapA, R - 3 * nw - gapA - gapB];
    // before
    s.addText("قبل الحرب", ctr({ x: xs[0], y: 2.68, w: nw, h: 0.42, fontSize: 15, bold: true, color: "2E7D4F", valign: "middle" }));
    // after (spans two cards)
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: xs[2], y: 2.68, w: 2 * nw + gapB, h: 0.42, rectRadius: 0.21, fill: { color: "FDE8E2" }, line: { color: "FDE8E2" } });
    s.addText("بعد الحرب", ctr({ x: xs[2], y: 2.68, w: 2 * nw + gapB, h: 0.42, fontSize: 15, bold: true, color: "B42318", valign: "middle" }));
    const nums = [["100 ألف $", "رحلة ناقلة واحدة", "تكلفة تأمين مخاطر الحرب لرحلة واحدة قبل التوترات", "2E7D4F"],
      ["3 ملايين $", "رحلة من ينبع", "قسط الناقلات: من أقل من 1% إلى 3% من قيمة السفينة", ORANGE],
      ["7 ملايين $", "موانئ جنوبية أو عبر هرمز", "وقد يصل القسط إلى 7% لموانئ جنوب ينبع مثل جازان", RED]];
    nums.forEach((n, i) => {
      const x = xs[i];
      card(s, x, y, nw, ch);
      s.addText(n[1], ar({ x: x + 0.3, y: y + 0.22, w: nw - 0.6, h: 0.4, fontSize: 14, bold: true, color: MUTED }));
      s.addText(n[0], ar({ x: x + 0.3, y: y + 0.65, w: nw - 0.6, h: 0.85, fontSize: 40, bold: true, color: n[3] }));
      s.addText(n[2], ar({ x: x + 0.3, y: y + 1.55, w: nw - 0.6, h: 0.7, fontSize: 13, color: MUTED }));
    });
    s.addShape(pres.shapes.LINE, { x: xs[0] - gapA + 0.12, y: y + ch / 2, w: gapA - 0.24, h: 0, line: { color: RED, width: 3, beginArrowType: "triangle" } });
  }
  s.addText("والأقساط لا تثبت: تغطي عادةً 7 أيام فقط، وتُراجع كل 24 ساعة.", ar({ x: M, y: 5.85, w: CW, h: 0.45, fontSize: 17, bold: true, color: NAVY }));
  s.addText("المصادر: رويترز عبر Insurance Journal (سبتمبر 2026)، Insurance Journal (يوليو 2026). الأرقام عروض سوقية وليست أسعاراً متفقاً عليها.", ar({ x: M, y: 6.4, w: CW, h: 0.3, fontSize: 10, color: "8A97A8" }));

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

  // ================= 4. SOLUTION =================
  pres.addSection({ title: "الحل والاستراتيجية" });
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "الحل والاستراتيجية" });
  kicker(s, "الحل");
  s.addText("لذلك نقترح «مرساة»", { placeholder: "title" });
  s.addText("منصة وطنية تقوم على فكرة واحدة: أن نصنع مؤشر خطر سعودياً خاصاً بنا للمسارات البحرية، ببياناتنا ومنهجيتنا، لا بالاعتماد على تقييم لجنة الحرب المشتركة، ونضعه في موقع واحد يجمع للشاحن الخطر والتكلفة والمؤمّنين السعوديين.",
    ar({ x: M, y: 1.6, w: CW, h: 0.95, fontSize: 15.5, color: MUTED, lineSpacingMultiple: 1.2 }));
  const ins = ["الحوادث الأمنية البحرية", "تتبع السفن (AIS) الوطني", "بيانات الموانئ السعودية", "أقساط ومطالبات المؤمّنين السعوديين"];
  const outs = ["درجة خطر لكل منطقة ومسار", "قسط تقديري لكل مسار وميناء", "مقارنة محايدة بين المؤمّنين", "تنبيه بأفضل وقت للتأمين"];
  s.addText("بياناتنا الوطنية", ar({ x: 9.2, y: 2.65, w: R - 9.2, h: 0.4, fontSize: 14, bold: true, color: TEAL }));
  s.addText("ما يحصل عليه الشاحن", ar({ x: M, y: 2.65, w: 3.7, h: 0.4, fontSize: 14, bold: true, color: TEAL }));
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

  // ================= 5. STRATEGY =================
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "الحل والاستراتيجية" });
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

  // ================= 6. HOW THE WEBSITE WORKS =================
  pres.addSection({ title: "الموقع" });
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "الموقع" });
  kicker(s, "موقع مرساة");
  s.addText("كيف يستخدم الشاحن مرساة؟ أربع خطوات", { placeholder: "title" });
  const steps = [
    [lu.LuRadar, "الرادار", "خريطة حية لمناطق الخطر وحركة السفن، ومسارات مقترحة لكل منها درجة خطر ومدة وقسط متوقع."],
    [lu.LuShieldCheck, "شركات التأمين", "كل مؤمّن سعودي: الشحنات التي يغطيها، وأنواع وثائقه، ومشاركته في الوعاء الوطني."],
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
    s.addText("لقطة من النموذج الأولي، والأسماء والأرقام توضيحية", ar({ x: ix, y: iy + h + 0.14, w, h: 0.3, fontSize: 10, color: "8A97A8" }));
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
    ["القسط التقديري", "يكتب الشاحن قيمة الشحنة، فيرى أقل قسط ومتوسطه وتغيّره خلال 14 يوماً."],
    ["توصية التوقيت", "الذكاء الاصطناعي يقترح متى يؤمّن الشاحن، مع درجة ثقة واضحة."],
    ["تنبيه السعر", "يحفظ الشحنة ويحدد سعراً، فيصله إشعار عند نزول القسط."],
  ]);
  await screen("الشاشة 3 · شركات التأمين", "المؤمّنون السعوديون والمشاركون في الوعاء الوطني", "insurers_top.png", 2880 / 1340, [
    ["مؤمّنو الوعاء الوطني", "علامة واضحة على المؤمّنين المشاركين في الوعاء الوطني لتأمين مخاطر الحرب."],
    ["ما يغطيه كل مؤمّن", "أنواع الشحنات (حاويات، زيوت، سيارات، مبردة) والوثائق ونطاق استرشادي للقسط."],
    ["اطلب عرضاً رسمياً", "ينتقل الشاحن مباشرة إلى المؤمّن السعودي، ولا يوجد ترتيب مدفوع."],
  ]);

  // ================= 10. VISION 2030 + POOL =================
  pres.addSection({ title: "رؤية المملكة" });
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "رؤية المملكة" });
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

  // ================= 11. WHY A SAUDI INDEX (+ knowledge pyramid) =================
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "رؤية المملكة" });
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
      [lu.LuDatabase, "نملك بياناتنا", "البيانات أساس الهرم، ومنها تُبنى المعلومات والمعرفة والقرار. ومع سنوات التشغيل تتراكم لدى المملكة بيانات متميزة عن مناطقها."],
      [lu.LuFlag, "استقلال لوجستي", "قدرة وطنية تقلل اعتماد خطة المملكة اللوجستية على المراجع الخارجية."],
      [lu.LuEarth, "مرجع عالمي مستقبلاً", "يبدأ بالمناطق المحيطة بالمملكة ويتوسع، وقد تعتمد عليه دول أخرى."],
      [lu.LuTrainFront, "يكمل منظومة النقل", "جزء مما تبنيه المملكة من موانئ وقطارات وبنية لوجستية نحو المركز العالمي."],
    ];
    const hs = [1.45, 1.1, 1.1, 1.1];
    let y = 1.8;
    for (let i = 0; i < 4; i++) {
      const x = 6.2;
      await bubble(s, why[i][0], R - 0.8, y + 0.05, 0.8);
      s.addText(why[i][1], ar({ x, y, w: R - x - 1.0, h: 0.42, fontSize: 17, bold: true, color: NAVY }));
      s.addText(why[i][2], ar({ x, y: y + 0.42, w: R - x - 1.0, h: hs[i] - 0.5, fontSize: 13.5, color: MUTED }));
      y += hs[i] + 0.1;
    }
  }

  // ================= 12. TRACK 4 ANSWER =================
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "رؤية المملكة" });
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

  // ================= 13. TIMELINE =================
  pres.addSection({ title: "التنفيذ والتفاصيل" });
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "التنفيذ والتفاصيل" });
  kicker(s, "خطة التنفيذ");
  s.addText("خمس مراحل خلال خمس سنوات", { placeholder: "title" });
  {
    const top = 1.75, labelW = 4.6, gx1 = R - labelW, gx0 = M, yw = (gx1 - gx0) / 5, rowH = 0.8;
    card(s, M, top, CW, 4.7);
    // year columns (year 1 on the right)
    for (let y = 1; y <= 5; y++) {
      const x = gx1 - y * yw;
      if (y % 2) s.addShape(pres.shapes.RECTANGLE, { x, y: top + 0.6, w: yw, h: 4.0 - 0.05, fill: { color: "F7F9FC" }, line: { color: "F7F9FC" } }) ;
      s.addText(`السنة ${y}`, ctr({ x, y: top + 0.15, w: yw, h: 0.4, fontSize: 14, bold: true, color: NAVY }));
    }
    s.addShape(pres.shapes.LINE, { x: M + 0.2, y: top + 0.6, w: CW - 0.4, h: 0, line: { color: LINE, width: 1 } });
    const ph = [
      ["صنع مقياس الخطر السعودي", "الموانئ والحوادث وحركة السفن (AIS)", 0, 2, TEAL],
      ["الاتفاق مع شركات التأمين", "مؤمّنان سعوديان على الأقل لعرض التغطيات والأسعار", 1, 2, TEAL_D],
      ["بناء الموقع وتأمينه", "بناء المنصة وحمايتها من الاختراق", 1, 3, NAVY2],
      ["إطلاق تجريبي", "ربط الخطر بالسعر والتنبيهات مع شركات محدودة", 2, 3, ORANGE],
      ["الإطلاق الرسمي", "اعتماد المؤشر محلياً ثم إقليمياً", 3, 5, NAVY],
    ];
    ph.forEach((p, i) => {
      const y = top + 0.68 + i * rowH;
      badge(s, i + 1, R - 0.55, y + 0.15, 0.42, p[4]);
      s.addText(p[0], ar({ x: gx1 + 0.15, y: y + 0.0, w: labelW - 0.85, h: 0.38, fontSize: 14, bold: true, color: NAVY }));
      s.addText(p[1], ar({ x: gx1 + 0.15, y: y + 0.38, w: labelW - 0.85, h: 0.34, fontSize: 11, color: MUTED }));
      const bx = gx1 - p[3] * yw + 0.08, bw = (p[3] - p[2]) * yw - 0.16;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: bx, y: y + 0.12, w: bw, h: 0.5, rectRadius: 0.25, fill: { color: p[4] }, line: { color: p[4] } });
      const span = p[3] - p[2] > 1 ? `السنة ${p[2] + 1}–${p[3]}` : `السنة ${p[3]}`;
      s.addText(span, ctr({ x: bx, y: y + 0.12, w: bw, h: 0.5, fontSize: 12, bold: true, color: WHITE, valign: "middle" }));
    });
  }
  s.addText([{ text: "الجهة المالكة المقترحة: ", options: { bold: true, color: NAVY } }, { text: "مشروع مشترك بين هيئة التأمين والهيئة العامة للنقل تحت وزارة النقل والخدمات اللوجستية، بالتنسيق مع موانئ والشركة السعودية لإعادة التأمين.", options: { color: MUTED } }],
    ar({ x: M, y: 6.62, w: CW, h: 0.35, fontSize: 12 }));

  // ================= 14. COMPARISON =================
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "التنفيذ والتفاصيل" });
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

  // ================= 15. CHALLENGES =================
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "التنفيذ والتفاصيل" });
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

  // ================= 16. DATA & TECH =================
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "التنفيذ والتفاصيل" });
  kicker(s, "البيانات والتقنيات");
  s.addText("من أين تأتي البيانات؟ وما التقنيات المستخدمة؟", { placeholder: "title" });
  const src = [
    [lu.LuFileWarning, "مناطق الخطر", "نحددها بأنفسنا بمنهجية سعودية: الحوادث الأمنية وتنبيهات الجهات الوطنية"],
    [lu.LuShip, "تتبع السفن (AIS)", "شبكة تتبع وطنية بالشراكة مع موانئ، مع مزود مرخص في البداية"],
    [lu.LuBuilding2, "شركات التأمين", "اتفاقية مع كل شركة، وتحديث الأسعار عبر واجهة برمجية"],
    [lu.LuLandmark, "الوعاء الوطني", "الشركة السعودية لإعادة التأمين وهيئة التأمين وفق الأنظمة"],
  ];
  for (let i = 0; i < 4; i++) {
    const x = R - pw - i * (pw + 0.3), y = 1.8;
    card(s, x, y, pw, 2.75);
    await bubble(s, src[i][0], x + pw - 1.05, y + 0.3, 0.75);
    s.addText(src[i][1], ar({ x: x + 0.3, y: y + 1.2, w: pw - 0.6, h: 0.45, fontSize: 17, bold: true, color: NAVY }));
    s.addText(src[i][2], ar({ x: x + 0.3, y: y + 1.68, w: pw - 0.6, h: 0.95, fontSize: 13, color: MUTED }));
  }
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y: 4.8, w: CW, h: 1.55, rectRadius: 0.12, fill: { color: NAVY }, line: { color: NAVY } });
  s.addText("التقنيات", ar({ x: R - 1.9, y: 4.8, w: 1.6, h: 1.55, fontSize: 18, bold: true, color: ICE, valign: "middle" }));
  const tech = [[lu.LuMap, "خرائط تفاعلية حية"], [lu.LuPlug, "واجهات برمجية لربط المؤمّنين"], [lu.LuChartLine, "نموذج تنبؤ بالأسعار (سلاسل زمنية)"], [lu.LuMessageCircle, "مساعد ذكي بنموذج لغوي"]];
  const tw4 = (CW - 2.2) / 4;
  for (let i = 0; i < 4; i++) {
    const x = R - 2.1 - tw4 * (i + 1);
    s.addImage({ data: await icon(tech[i][0], ICE), x: x + tw4 - 0.55, y: 5.33, w: 0.42, h: 0.42 });
    s.addText(tech[i][1], ar({ x: x + 0.1, y: 5.05, w: tw4 - 0.75, h: 1.0, fontSize: 13.5, color: WHITE, valign: "middle" }));
  }
  s.addText("المصادر مقترحة وتتطلب اتفاقيات في مرحلة التنفيذ، ولا يُستخدم الذكاء الاصطناعي في تحديد مناطق الخطر.", ar({ x: M, y: 6.5, w: CW, h: 0.3, fontSize: 11.5, color: MUTED }));

  // ================= 17. IMPACT =================
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "التنفيذ والتفاصيل" });
  kicker(s, "الأثر المتوقع");
  s.addText("تمكين المؤمّن السعودي، وقرار أوضح للشاحن", { placeholder: "title" });
  const imp = [
    [lu.LuShieldCheck, "المؤمّنون السعوديون", "قناة رقمية تصلهم بعملاء جدد، وتقلل اعتماد السوق على المؤمّنين العالميين."],
    [lu.LuContainer, "الشاحنون", "مقارنة المؤمّنين وتكلفة كل مسار في مكان واحد، وتنبيه بأفضل وقت للتأمين."],
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

  // ================= 18. TEAM =================
  pres.addSection({ title: "الختام" });
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "الختام" });
  kicker(s, "الفريق");
  s.addText("فريق مرساة", { placeholder: "title" });
  const team = [[lu.LuUserRound, "قائد الفريق", "التنسيق والتسليم وتمثيل الفريق"], [lu.LuBriefcase, "المجال والأعمال", "القيمة والجدوى وعلاقات المؤمّنين"], [lu.LuCode, "التقنية والبيانات", "الحل والنموذج الأولي ومصادر البيانات"], [lu.LuPalette, "التصميم وتجربة المستخدم", "الرحلة والواجهات"]];
  for (let i = 0; i < 4; i++) {
    const x = R - pw - i * (pw + 0.3), y = 1.85;
    card(s, x, y, pw, 4.5);
    await bubble(s, team[i][0], x + pw / 2 - 0.75, y + 0.45, 1.5, i === 0 ? NAVY : ICE_BG, i === 0 ? ICE : TEAL_D);
    s.addText("الاسم", ctr({ x: x + 0.2, y: y + 2.2, w: pw - 0.4, h: 0.5, fontSize: 20, bold: true, color: NAVY }));
    s.addText(team[i][1], ctr({ x: x + 0.2, y: y + 2.75, w: pw - 0.4, h: 0.45, fontSize: 13.5, bold: true, color: TEAL }));
    s.addText(team[i][2], ctr({ x: x + 0.25, y: y + 3.4, w: pw - 0.5, h: 0.8, fontSize: 13, color: MUTED }));
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
