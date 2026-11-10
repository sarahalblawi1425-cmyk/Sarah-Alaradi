// Mirsah (مرساة) pitch deck — Naqlthon 5, Track 4
const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const path = require("path");
const lu = require("react-icons/lu");
const { applyTheme } = require(process.env.SKILL + "/scripts/apply_theme.js");

const SHOTS = process.argv[2];
const OUT = process.argv[3];

// ---------- palette (from the prototype) ----------
const NAVY = "0B2545", DEEP = "0B1D33", INK = "0F1B2D", TEAL = "0E7C86", TEAL_D = "0B5C66",
  ICE = "8FD3DA", ICE_BG = "DDF2F3", BG = "F4F6F9", LINE = "DDE3EB", MUTED = "4A5568",
  RED = "B42318", ORANGE = "C2410C", AMBER = "A86A0C", GREEN = "2E7D4F", WHITE = "FFFFFF", SOFT = "C9D6E8";
const F = "Arial";

const THEME = {
  name: "Mirsah",
  headFontFace: F, bodyFontFace: F,
  colors: { dk1: INK, lt1: WHITE, dk2: NAVY, lt2: BG, accent1: TEAL, accent2: ICE, accent3: "E0A526",
    accent4: ORANGE, accent5: RED, accent6: GREEN, hlink: TEAL_D, folHlink: "13315C" },
};

const W = 13.333, H = 7.5, M = 0.6, R = W - M; // right edge for RTL layout

// ---------- helpers ----------
const ar = (o = {}) => Object.assign({ fontFace: F, rtlMode: true, lang: "ar-SA", align: "right", isTextBox: true, margin: 0, valign: "top", color: INK }, o);

async function icon(Comp, color, size = 256) {
  const svg = RDS.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).resize(size, size).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.rtlMode = true;
pres.title = "مرساة — منصة وطنية للتأمين البحري وتقييم مخاطر المسارات";
pres.theme = { headFontFace: F, bodyFontFace: F };

// ---------- layouts ----------
pres.defineSlideMaster({
  title: "LIGHT",
  background: { color: BG },
  objects: [
    { text: { text: "مرساة · نقل ثون 5 · المسار الرابع", options: ar({ x: 4.5, y: 7.0, w: 4.33, h: 0.3, fontSize: 10, color: "8A97A8", align: "center" }) } },
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.45, w: W - 2 * M, h: 0.8, fontFace: F, fontSize: 32, bold: true, color: NAVY, align: "right", valign: "middle", rtlMode: true, margin: 0 }, text: "" } },
  ],
  slideNumber: { x: 0.5, y: 7.0, w: 0.5, h: 0.3, fontFace: F, fontSize: 10, color: "8A97A8" },
});
pres.defineSlideMaster({
  title: "DARK",
  background: { color: NAVY },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.45, w: W - 2 * M, h: 0.8, fontFace: F, fontSize: 32, bold: true, color: WHITE, align: "right", valign: "middle", rtlMode: true, margin: 0 }, text: "" } },
  ],
  slideNumber: { x: 0.5, y: 7.0, w: 0.5, h: 0.3, fontFace: F, fontSize: 10, color: "6E8DB3" },
});
pres.defineSlideMaster({ title: "COVER", background: { color: DEEP }, objects: [] });

function subtitle(slide, text, dark = false, y = 1.22) {
  slide.addText(text, ar({ x: M, y, w: W - 2 * M, h: 0.45, fontSize: 16, color: dark ? SOFT : MUTED }));
}
function card(slide, x, y, w, h, fill = WHITE, line = LINE) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: fill }, line: { color: line, width: 1 } });
}
async function iconCircle(slide, Comp, x, y, d, bg, fg) {
  slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: bg }, line: { color: bg } });
  const pad = d * 0.24;
  slide.addImage({ data: await icon(Comp, fg), x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad });
}
function numBadge(slide, n, x, y, d, bg = TEAL, fg = WHITE) {
  slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: bg }, line: { color: bg } });
  slide.addText(String(n), { x, y, w: d, h: d, fontFace: F, fontSize: d > 0.45 ? 16 : 12, bold: true, color: fg, align: "center", valign: "middle", margin: 0, isTextBox: true });
}

(async () => {
  // =============== 1. Title ===============
  pres.addSection({ title: "البداية" });
  let s = pres.addSlide({ masterName: "COVER", sectionTitle: "البداية" });
  s.background = { color: DEEP };
  s.addImage({ path: path.join(SHOTS, "map.png"), x: 0, y: 0, w: 6.0, h: 7.5, sizing: { type: "cover", w: 6.0, h: 7.5 } });
  // anchor + name
  await iconCircle(s, lu.LuAnchor, R - 1.0, 1.15, 1.0, "13315C", ICE);
  s.addText("مرساة", ar({ x: 6.6, y: 2.3, w: R - 6.6, h: 1.2, fontSize: 66, bold: true, color: WHITE }));
  s.addText("منصة وطنية للتأمين البحري وتقييم مخاطر المسارات", ar({ x: 6.6, y: 3.55, w: R - 6.6, h: 0.9, fontSize: 24, color: ICE }));
  s.addText("مؤشر خطر سعودي يجمع خطر المسار وتكلفة التأمين ومقارنة المؤمّنين المحليين في مكان واحد", ar({ x: 6.6, y: 4.5, w: R - 6.6, h: 0.9, fontSize: 16, color: SOFT }));
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.4, y: 5.75, w: R - 6.4, h: 0.55, rectRadius: 0.27, fill: { color: "13315C" }, line: { color: "1E3A5F" } });
  s.addText("نقل ثون 5 · المسار الرابع: الابتكار في الخدمات اللوجستية وسلاسل الإمداد", ar({ x: 6.5, y: 5.75, w: R - 6.6, h: 0.55, fontSize: 12.5, color: WHITE, valign: "middle", align: "center" }));
  s.addNotes("أهلاً بكم. نقدم لكم مرساة: منصة وطنية مقترحة تبني مؤشر خطر سعودي للمسارات البحرية، وتجمع للمصدّر والمستورد الخطر والتكلفة وعروض المؤمّنين المحليين في مكان واحد.");

  // =============== 2. The problem — numbers ===============
  pres.addSection({ title: "التحدي" });
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "التحدي" });
  s.addText("التحدي: تأمين مخاطر الحرب أصبح مكلفاً ومتقلباً", { placeholder: "title" });
  subtitle(s, "عندما تعبر السفينة منطقة توتر أمني، تُؤمَّن الشحنة والسفينة بقسط إضافي يتغير بسرعة");
  // native chart: cost of a trip (million USD)
  card(s, M, 1.95, 6.4, 4.75);
  s.addText("تكلفة تأمين الحرب لرحلة ناقلة واحدة (مليون دولار)", ar({ x: M + 0.3, y: 2.1, w: 5.8, h: 0.4, fontSize: 14, bold: true, color: NAVY }));
  s.addChart(pres.charts.BAR, [{ name: "مليون دولار", labels: ["قبل الحرب", "من ينبع", "موانئ جنوبية أو عبر هرمز"], values: [0.1, 3, 7] }], {
    x: M + 0.2, y: 2.55, w: 6.0, h: 3.7, barDir: "col", chartColors: ["6E8DB3", ORANGE, RED],
    showValue: true, dataLabelPosition: "outEnd", dataLabelColor: INK, dataLabelFontSize: 14, dataLabelFontBold: true, dataLabelFontFace: "+mn-lt",
    dataLabelFormatCode: "0.0", catAxisLabelColor: MUTED, catAxisLabelFontSize: 12, catAxisLabelFontFace: "+mn-lt",
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false, barGapWidthPct: 60,
    catAxisOrientation: "maxMin", catAxisLineShow: false,
  });
  s.addText("المصدر: رويترز عبر Insurance Journal، سبتمبر 2026", ar({ x: M + 0.3, y: 6.3, w: 5.8, h: 0.3, fontSize: 10, color: "8A97A8" }));
  // stat callouts on the right (RTL reading starts right)
  const stats = [
    ["من أقل من 1% إلى 3%", "قسط ناقلات ميناء ينبع من قيمة السفينة منذ يوليو 2026", ORANGE],
    ["حتى 7%", "للموانئ جنوب ينبع مثل جازان", RED],
    ["كل 24 ساعة", "تُراجع أقساط الحرب، وتغطي عادةً 7 أيام فقط", NAVY],
  ];
  stats.forEach((st, i) => {
    const y = 1.95 + i * 1.62;
    card(s, 7.35, y, R - 7.35, 1.45);
    s.addText(st[0], ar({ x: 7.6, y: y + 0.18, w: R - 7.85, h: 0.6, fontSize: 28, bold: true, color: st[2] }));
    s.addText(st[1], ar({ x: 7.6, y: y + 0.8, w: R - 7.85, h: 0.5, fontSize: 14, color: MUTED }));
  });
  s.addNotes("التأمين من أهم ما تحرص عليه شركات النقل البحري. ومع التوترات، قفز قسط ناقلات ينبع من أقل من 1% إلى نحو 3% من قيمة السفينة، وقد يصل إلى 7% لموانئ جنوب ينبع. رحلة كانت تكلف 100 ألف دولار أصبحت 3 إلى 7 ملايين. والأقساط تُراجع كل 24 ساعة.");

  // =============== 3. Problem — user pain ===============
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "التحدي" });
  s.addText("لماذا يصعب على الشاحن أن يقرر؟", { placeholder: "title" });
  subtitle(s, "المصدّر والمستورد لا يجد مكاناً واحداً يرى فيه الخطر والتكلفة والخيارات المحلية");
  const pains = [
    [lu.LuGlobe, "مرجعية دولية واحدة", "يُسعَّر الخطر على قوائم لجنة الحرب المشتركة في لندن (JWC)، ولا تملك المملكة مؤشراً وطنياً خاصاً بمسارات تجارتها."],
    [lu.LuScale, "صعوبة المقارنة", "لا توجد بيانات عامة بالأسعار المتفق عليها فعلياً، والأرقام المتداولة عروض قد تختلف عن السعر الفعلي."],
    [lu.LuPhoneCall, "قرار دون صورة كاملة", "يتواصل الشاحن مع كل شركة تأمين على حدة، ولا يرى كيف يتغير القسط إذا غيّر الميناء أو المسار."],
  ];
  const cw = (W - 2 * M - 2 * 0.35) / 3;
  for (let i = 0; i < 3; i++) {
    const x = R - cw - i * (cw + 0.35), y = 2.05;
    card(s, x, y, cw, 4.5);
    await iconCircle(s, pains[i][0], x + cw - 1.2, y + 0.4, 0.9, ICE_BG, TEAL_D);
    s.addText(pains[i][1], ar({ x: x + 0.3, y: y + 1.55, w: cw - 0.6, h: 0.55, fontSize: 20, bold: true, color: NAVY }));
    s.addText(pains[i][2], ar({ x: x + 0.3, y: y + 2.2, w: cw - 0.6, h: 2.0, fontSize: 15, color: MUTED, lineSpacingMultiple: 1.25 }));
  }
  s.addNotes("المشكلة ليست السعر فقط: المرجع دولي واحد، والمقارنة صعبة لأن الأسعار غير معلنة، والشاحن يتصل بكل شركة على حدة فيقرر دون صورة كاملة.");

  // =============== 4. Solution ===============
  pres.addSection({ title: "الحل" });
  s = pres.addSlide({ masterName: "DARK", sectionTitle: "الحل" });
  s.addText("الحل: مرساة", { placeholder: "title" });
  subtitle(s, "مؤشر خطر سعودي للمسارات البحرية، يضعه موقع واحد في متناول المستخدم", true);
  // inputs (right) -> index (center) -> outputs (left)
  const ins = ["قوائم لجنة الحرب المشتركة المنشورة", "تتبع السفن (AIS) والاختناقات", "بيانات الموانئ والحوادث الأمنية", "أسعار وشروط المؤمّنين السعوديين"];
  const outs = ["درجة خطر لكل منطقة ومسار", "قسط تقديري لكل مسار وميناء", "مقارنة محايدة بين المؤمّنين", "تنبيهات بتوقيت التأمين"];
  s.addText("المدخلات", ar({ x: 9.3, y: 2.0, w: 3.43, h: 0.4, fontSize: 14, bold: true, color: ICE }));
  s.addText("ما يحصل عليه المستخدم", ar({ x: M, y: 2.0, w: 3.6, h: 0.4, fontSize: 14, bold: true, color: ICE }));
  ins.forEach((t, i) => {
    const y = 2.5 + i * 0.95;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 9.3, y, w: R - 9.3, h: 0.78, rectRadius: 0.1, fill: { color: "13315C" }, line: { color: "1E3A5F" } });
    s.addText(t, ar({ x: 9.45, y, w: R - 9.6, h: 0.78, fontSize: 14, color: WHITE, valign: "middle" }));
  });
  outs.forEach((t, i) => {
    const y = 2.5 + i * 0.95;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y, w: 3.6, h: 0.78, rectRadius: 0.1, fill: { color: WHITE }, line: { color: WHITE } });
    s.addText(t, ar({ x: M + 0.15, y, w: 3.3, h: 0.78, fontSize: 14, bold: true, color: NAVY, valign: "middle" }));
  });
  // center hub
  s.addShape(pres.shapes.OVAL, { x: 5.17, y: 2.45, w: 3.0, h: 3.0, fill: { color: TEAL }, line: { color: ICE, width: 3 } });
  s.addImage({ data: await icon(lu.LuAnchor, WHITE), x: 6.27, y: 2.8, w: 0.8, h: 0.8 });
  s.addText("مؤشر مرساة", { x: 5.17, y: 3.65, w: 3.0, h: 0.5, fontFace: F, fontSize: 22, bold: true, color: WHITE, align: "center", rtlMode: true, margin: 0, isTextBox: true });
  s.addText("للخطر البحري", { x: 5.17, y: 4.1, w: 3.0, h: 0.4, fontFace: F, fontSize: 15, color: ICE_BG, align: "center", rtlMode: true, margin: 0, isTextBox: true });
  // arrows
  s.addShape(pres.shapes.LINE, { x: 8.3, y: 3.95, w: 0.85, h: 0, line: { color: ICE, width: 2.5, beginArrowType: "triangle" } });
  s.addShape(pres.shapes.LINE, { x: 4.3, y: 3.95, w: 0.75, h: 0, line: { color: ICE, width: 2.5, beginArrowType: "triangle" } });
  s.addText("المؤشر لا يحل محل القوائم الدولية؛ يعمل بجوارها حتى يثبت اعتماده، ويبقى التأكيد النهائي للسعر من شركة التأمين.", ar({ x: M, y: 6.4, w: W - 2 * M, h: 0.4, fontSize: 13, color: SOFT }));
  s.addNotes("فكرتنا واحدة: أن تبني المملكة مؤشر خطر سعودي خاص بمساراتها البحرية. يأخذ مدخلات دولية مرخصة، ويضيف بيانات الموانئ والمؤمّنين المحليين، ويعطي المستخدم درجة خطر وقسطاً تقديرياً ومقارنة وتنبيهات.");

  // =============== 5. Prototype — radar & routes ===============
  pres.addSection({ title: "النموذج الأولي" });
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "النموذج الأولي" });
  s.addText("النموذج الأولي (1): البحث عن المسار ورادار الخطر", { placeholder: "title" });
  {
    const iw = 8.75, ih = iw * 1460 / 2544, ix = M, iy = 1.55;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: ix - 0.08, y: iy - 0.08, w: iw + 0.16, h: ih + 0.16, rectRadius: 0.1, fill: { color: WHITE }, line: { color: LINE }, shadow: { type: "outer", color: "0B2545", opacity: 0.18, blur: 10, offset: 3, angle: 90 } });
    s.addImage({ path: path.join(SHOTS, "search_row.png"), x: ix, y: iy, w: iw, h: ih, altText: "لقطة من النموذج الأولي: قائمة المسارات وخريطة رادار الخطر" });
    s.addText("لقطة من النموذج الأولي — الأرقام والأسماء توضيحية", ar({ x: ix, y: iy + ih + 0.12, w: iw, h: 0.3, fontSize: 10, color: "8A97A8" }));
  }
  const p1 = [
    ["يختار المستخدم ميناء الانطلاق والوجهة", "شنغهاي إلى جدة أو الدمام أو ينبع أو جازان، ونوع الشحنة وتاريخ الإبحار."],
    ["تقترح المنصة أكثر من مسار", "لكل مسار: أعلى درجة خطر، ومدة الإبحار، والمسافة، وعدد مناطق الخطر."],
    ["رادار حي للخطر وحركة السفن", "ألوان المناطق من منخفض إلى حرج، مثل باب المندب 7.4 وهرمز 5.8."],
  ];
  p1.forEach((p, i) => {
    const y = 1.6 + i * 1.6, x = 9.75;
    numBadge(s, i + 1, R - 0.5, y, 0.5);
    s.addText(p[0], ar({ x, y: y - 0.02, w: R - x - 0.65, h: 0.6, fontSize: 15, bold: true, color: NAVY }));
    s.addText(p[1], ar({ x, y: y + 0.6, w: R - x - 0.65, h: 0.85, fontSize: 13, color: MUTED }));
  });
  s.addNotes("هذه الشاشة الأولى من النموذج الأولي. يدخل المستورد الميناءين، فتظهر المسارات المتاحة، ولكل مسار درجة خطر ومدة ومسافة. وعلى اليمين رادار يوضح مناطق الخطر وحركة السفن. مثلاً المسار إلى الدمام عبر هرمز يوازن بين المدة والخطر.");

  // =============== 6. Prototype — prices ===============
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "النموذج الأولي" });
  s.addText("النموذج الأولي (2): تفقد أسعار التأمين", { placeholder: "title" });
  {
    const ih = 4.95, iw = ih * 2880 / 1880, ix = M, iy = 1.55;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: ix - 0.08, y: iy - 0.08, w: iw + 0.16, h: ih + 0.16, rectRadius: 0.1, fill: { color: WHITE }, line: { color: LINE }, shadow: { type: "outer", color: "0B2545", opacity: 0.18, blur: 10, offset: 3, angle: 90 } });
    s.addImage({ path: path.join(SHOTS, "prices_top.png"), x: ix, y: iy, w: iw, h: ih, altText: "لقطة من النموذج الأولي: صفحة تفقد أسعار التأمين" });
    s.addText("لقطة من النموذج الأولي — الأسعار تقديرية للتوضيح", ar({ x: ix, y: iy + ih + 0.12, w: iw, h: 0.3, fontSize: 10, color: "8A97A8" }));
  }
  const p2 = [
    ["مقارنة المؤمّنين السعوديين", "القسط التقديري ونوع الوثيقة وصلاحية العرض، دون ترتيب مدفوع."],
    ["اتجاه السعر خلال 14 يوماً", "يرى المستخدم هل القسط في ارتفاع أو انخفاض قبل أن يقرر."],
    ["توصية التوقيت وتنبيه السعر", "ذكاء اصطناعي يقترح وقت التأمين مع درجة ثقة، وتنبيه عند نزول القسط."],
  ];
  p2.forEach((p, i) => {
    const y = 1.6 + i * 1.6, x = 9.4;
    numBadge(s, i + 1, R - 0.5, y, 0.5);
    s.addText(p[0], ar({ x, y: y - 0.02, w: R - x - 0.65, h: 0.6, fontSize: 15, bold: true, color: NAVY }));
    s.addText(p[1], ar({ x, y: y + 0.6, w: R - x - 0.65, h: 0.85, fontSize: 13, color: MUTED }));
  });
  s.addNotes("الشاشة الثانية: يختار المستخدم المسار ويكتب قيمة الشحنة، فيرى أقل قسط متاح، وتغيّر السعر خلال أسبوعين، وتوصية بتوقيت التأمين، ثم جدول عروض المؤمّنين السعوديين ويطلب عرضاً رسمياً بضغطة.");

  // =============== 7. User journey ===============
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "النموذج الأولي" });
  s.addText("رحلة المستخدم في أربع محطات", { placeholder: "title" });
  subtitle(s, "مستورد يريد شحن بضاعة من ميناء إلى آخر يفتح مرساة فيمر بالمحطات التالية");
  const steps = [
    [lu.LuRadar, "الرادار", "خريطة حية لمناطق الخطر وحركة السفن، ومسارات مقترحة لكل منها نسبة خطر ومدة وقسط متوقع."],
    [lu.LuShieldCheck, "شركات التأمين", "أنواع الشحنات والوثائق لكل مؤمّن سعودي، والقسط كنسبة من قيمة الشحنة حسب خطر المسار."],
    [lu.LuSparkles, "الذكاء الاصطناعي", "تنبيه بوقت التأمين مع درجة ثقة، وتوصية بالمؤمّن الأنسب مع سبب الترتيب، ومساعد يجيب الأسئلة."],
    [lu.LuBellRing, "مشاريعي", "يحفظ الشحنة فيصله إشعار عند تغير القسط، أو عند وصوله إلى السعر الذي حدده."],
  ];
  const sw = (W - 2 * M - 3 * 0.3) / 4;
  for (let i = 0; i < 4; i++) {
    const x = R - sw - i * (sw + 0.3), y = 2.1;
    card(s, x, y, sw, 4.45);
    await iconCircle(s, steps[i][0], x + sw / 2 - 0.5, y + 0.4, 1.0, i === 2 ? TEAL : ICE_BG, i === 2 ? WHITE : TEAL_D);
    s.addText(`${i + 1}. ${steps[i][1]}`, { x: x + 0.2, y: y + 1.6, w: sw - 0.4, h: 0.5, fontFace: F, fontSize: 18, bold: true, color: NAVY, align: "center", rtlMode: true, margin: 0, isTextBox: true });
    s.addText(steps[i][2], ar({ x: x + 0.25, y: y + 2.2, w: sw - 0.5, h: 2.0, fontSize: 14, color: MUTED, lineSpacingMultiple: 1.2 }));
    if (i < 3) s.addShape(pres.shapes.LINE, { x: x - 0.27, y: y + 0.9, w: 0.24, h: 0, line: { color: TEAL, width: 2, endArrowType: "triangle" }, flipH: true });
  }
  s.addText("تحديد مناطق الخطر يعتمد على مصادر رسمية، وليس على الذكاء الاصطناعي.", ar({ x: M, y: 6.62, w: W - 2 * M, h: 0.3, fontSize: 12, color: TEAL_D, bold: true }));
  s.addNotes("الرحلة كاملة في أربع محطات: الرادار، ثم شركات التأمين، ثم الذكاء الاصطناعي في ثلاث وظائف، ثم مشاريعي للتنبيهات. ونؤكد أن تحديد مناطق الخطر من مصادر رسمية.");

  // =============== 8. Why a Saudi index ===============
  s = pres.addSlide({ masterName: "DARK", sectionTitle: "الحل" });
  s.addText("لماذا مؤشر خطر سعودي؟", { placeholder: "title" });
  subtitle(s, "مرساة لا تنافس المراجع الدولية بل تعمل بجوارها: المناطق الخطرة تبقى خطرة، والآمنة تبقى آمنة", true);
  const why = [
    [lu.LuDatabase, "قوة المعلومات", "من يملك البيانات يحللها ويبني الحلول عليها، ومع سنوات التشغيل تتراكم بيانات متميزة عن مناطقنا."],
    [lu.LuFlag, "الاستقلال اللوجستي", "قدرة وطنية تقلل اعتماد الخطة اللوجستية للمملكة على المراجع الخارجية وحدها."],
    [lu.LuEarth, "مرجعية مستقبلية", "يبدأ بالمناطق المحيطة بالمملكة ويتوسع تدريجياً، وقد تستفيد منه دول أخرى مرجعاً عالمياً."],
    [lu.LuTrainFront, "تكامل مع جهود المملكة", "جزء من منظومة القطارات والموانئ والبنية اللوجستية نحو المركز اللوجستي العالمي."],
  ];
  const ww = (W - 2 * M - 0.4) / 2;
  for (let i = 0; i < 4; i++) {
    const col = i % 2, row = Math.floor(i / 2);
    const x = R - ww - col * (ww + 0.4), y = 2.1 + row * 2.3;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: ww, h: 2.0, rectRadius: 0.12, fill: { color: "13315C" }, line: { color: "1E3A5F" } });
    await iconCircle(s, why[i][0], x + ww - 1.1, y + 0.3, 0.8, TEAL, WHITE);
    s.addText(why[i][1], ar({ x: x + 0.3, y: y + 0.35, w: ww - 1.6, h: 0.5, fontSize: 19, bold: true, color: WHITE }));
    s.addText(why[i][2], ar({ x: x + 0.3, y: y + 0.95, w: ww - 1.6, h: 0.9, fontSize: 14, color: SOFT }));
  }
  s.addNotes("لماذا مؤشر سعودي؟ لأن المملكة وهي تبني مكانتها مركزاً لوجستياً عالمياً تحتاج أن تملك قراءتها الوطنية للخطر. أربعة أسباب: قوة المعلومات، والاستقلال اللوجستي، ومرجعية مستقبلية، والتكامل مع جهود المملكة.");

  // =============== 9. Differentiation ===============
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "الحل" });
  s.addText("ما يميز مرساة عن الحلول القائمة", { placeholder: "title" });
  subtitle(s, "لم نجد منصة سعودية تجمع خطر المسار وتكلفته ومقارنة المؤمّنين المحليين في تجربة واحدة (بحث نوفمبر 2026)");
  {
    const hdr = (t) => ({ text: t, options: { bold: true, color: WHITE, fill: { color: NAVY }, fontSize: 13 } });
    const rows = [
      ["ShipWRITE من Canopius", "عرض وإصدار فوري لتأمين الحرب مع تتبع عالمي للسفن", "مقارنة بين عدة مؤمّنين سعوديين بدل مؤمّن واحد"],
      ["WarRates.com", "أسعار استرشادية لتأمين حرب الهياكل من عروض الأعضاء", "ربط السعر بمسار الشحنة وميناءيها ونوع الشحنة"],
      ["Insurwave مع Ambrey", "بيانات التهديدات البحرية لرؤية عالمية للمخاطر", "تركيز وطني وبيانات محلية للموانئ والمؤمّنين"],
      ["تأمين حرب البضائع من DP World", "وثيقة واحدة للبضائع من الباب إلى الباب", "مقارنة محايدة بين عدة مؤمّنين، لا منتج واحد"],
      ["الوعاء السعودي لتأمين مخاطر الحرب", "آلية وطنية بقيادة الشركة السعودية لإعادة التأمين", "واجهة رقمية مكملة تقود المستفيد إلى المؤمّنين المشاركين"],
    ];
    // RTL: first logical column on the right -> reverse arrays
    const data = [[hdr("ما تضيفه مرساة"), hdr("ما يقدمه"), hdr("الحل القائم")]];
    rows.forEach((r, i) => {
      const fill = { color: i % 2 ? "F7F9FC" : WHITE };
      data.push([
        { text: r[2], options: { bold: true, color: TEAL_D, fill } },
        { text: r[1], options: { color: MUTED, fill } },
        { text: r[0], options: { bold: true, color: NAVY, fill } },
      ]);
    });
    s.addTable(data, { x: M, y: 2.0, w: W - 2 * M, colW: [4.4, 4.4, W - 2 * M - 8.8], fontFace: F, fontSize: 13, align: "right", valign: "middle", rtlMode: true, border: { type: "solid", pt: 0.75, color: LINE }, margin: [0.08, 0.15, 0.08, 0.15], rowH: 0.72 });
  }
  s.addNotes("الحلول القائمة قيّمة لكنها عالمية أو تعالج جانباً واحداً. مرساة تضيف مقارنة محايدة بين مؤمّنين سعوديين، وتربط السعر بالمسار والميناء، وتكمل الوعاء الوطني لتأمين مخاطر الحرب ولا تنافسه.");

  // =============== 10. Track fit ===============
  pres.addSection({ title: "صلة المسار والأثر" });
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "صلة المسار والأثر" });
  s.addText("كيف تجيب مرساة عن سؤال المسار الرابع؟", { placeholder: "title" });
  // question banner
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y: 1.5, w: W - 2 * M, h: 0.85, rectRadius: 0.12, fill: { color: NAVY }, line: { color: NAVY } });
  s.addText([{ text: "سؤال التحدي:  ", options: { bold: true, color: ICE } }, { text: "كيف يمكن تقليل تكلفة التوصيل ووقته؟", options: { bold: true, color: WHITE } }], ar({ x: M + 0.3, y: 1.5, w: W - 2 * M - 0.6, h: 0.85, fontSize: 20, valign: "middle" }));
  const fit = [
    [lu.LuTimer, "كفاءة سلسلة الإمداد", "تنبيهات الأسعار تتيح التخطيط لتوقيت التأمين بدل قرار تحت ضغط المراجعة كل 24 ساعة."],
    [lu.LuRoute, "التكلفة والوقت", "يوازن بين الأسرع والأقل تكلفة والأكثر أماناً بالتكلفة الكلية للمسار."],
    [lu.LuShip, "تطوير الخدمات البحرية", "منصة واحدة للمسار والخطر والتأمين، وبيانات وطنية تتراكم عن الموانئ والمسارات."],
    [lu.LuLandmark, "دعم الاقتصاد الوطني", "تقليل الغموض في كلفة التأمين يشجع الشاحنين على اختيار الموانئ السعودية."],
  ];
  for (let i = 0; i < 4; i++) {
    const y = 2.65 + i * 1.0, x = 5.0;
    await iconCircle(s, fit[i][0], R - 0.7, y, 0.7, ICE_BG, TEAL_D);
    s.addText(fit[i][1], ar({ x, y: y - 0.02, w: R - x - 0.95, h: 0.4, fontSize: 16, bold: true, color: NAVY }));
    s.addText(fit[i][2], ar({ x, y: y + 0.38, w: R - x - 0.95, h: 0.5, fontSize: 13, color: MUTED }));
  }
  // national targets
  card(s, M, 2.65, 3.95, 3.85);
  s.addText("الأهداف الوطنية", ar({ x: M + 0.3, y: 2.85, w: 3.35, h: 0.4, fontSize: 14, bold: true, color: MUTED }));
  s.addText("40 مليون", ar({ x: M + 0.3, y: 3.3, w: 3.35, h: 0.75, fontSize: 36, bold: true, color: TEAL }));
  s.addText("حاوية نمطية سنوياً عبر الموانئ السعودية بحلول 2030", ar({ x: M + 0.3, y: 4.05, w: 3.35, h: 0.6, fontSize: 13, color: MUTED }));
  s.addText("8.3 مليون", ar({ x: M + 0.3, y: 4.8, w: 3.35, h: 0.75, fontSize: 36, bold: true, color: NAVY }));
  s.addText("حاوية ناولتها موانئ في 2025 بنمو 10.58%", ar({ x: M + 0.3, y: 5.55, w: 3.35, h: 0.6, fontSize: 13, color: MUTED }));
  s.addNotes("سؤال المسار الرابع: كيف نقلل تكلفة التوصيل ووقته؟ مرساة تجيب مباشرة: يعرف المستخدم التكلفة الكلية قبل أن يقرر، بدل التواصل مع كل شركة على حدة. وهذا يخدم هدف 40 مليون حاوية بحلول 2030.");

  // =============== 11. Data & tech ===============
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "صلة المسار والأثر" });
  s.addText("البيانات والتقنيات ومصادرها", { placeholder: "title" });
  subtitle(s, "المصادر مقترحة وتتطلب اتفاقيات في مرحلة التنفيذ");
  {
    const hdr = (t) => ({ text: t, options: { bold: true, color: WHITE, fill: { color: NAVY }, fontSize: 13 } });
    const rows = [
      ["المناطق المدرجة وعالية الخطر", "قوائم لجنة الحرب المشتركة المنشورة، وتنبيهات الجهات الوطنية", "متابعة المذكرات بعد التحقق من شروط استخدامها"],
      ["حركة السفن والاختناقات", "رادار تتبع (AIS) سعودي هدفاً للخطة، مع بيانات الهيئة العامة للموانئ", "شبكة تتبع وطنية بالشراكة مع موانئ، وترخيص مؤقت"],
      ["أسعار شركات التأمين وشروطها", "شركات التأمين البحري السعودية", "اتفاقية مع كل شركة، وتحديث عبر واجهة برمجية"],
      ["تفاصيل الوعاء الوطني", "الشركة السعودية لإعادة التأمين وهيئة التأمين", "شراكة وفق الأنظمة المعتمدة"],
    ];
    const data = [[hdr("آلية الحصول"), hdr("المصدر المقترح"), hdr("البيانات")]];
    rows.forEach((r, i) => {
      const fill = { color: i % 2 ? "F7F9FC" : WHITE };
      data.push([{ text: r[2], options: { color: MUTED, fill } }, { text: r[1], options: { color: MUTED, fill } }, { text: r[0], options: { bold: true, color: NAVY, fill } }]);
    });
    s.addTable(data, { x: M, y: 1.9, w: 8.3, colW: [2.9, 3.1, 2.3], fontFace: F, fontSize: 12, align: "right", valign: "middle", rtlMode: true, border: { type: "solid", pt: 0.75, color: LINE }, margin: [0.07, 0.12, 0.07, 0.12], rowH: 0.8 });
  }
  card(s, 9.2, 1.9, R - 9.2, 4.65, NAVY, NAVY);
  s.addText("التقنيات", ar({ x: 9.45, y: 2.1, w: R - 9.7, h: 0.45, fontSize: 18, bold: true, color: ICE }));
  const tech = [[lu.LuMap, "خرائط تفاعلية حية"], [lu.LuPlug, "واجهات برمجية لربط المؤمّنين"], [lu.LuChartLine, "نموذج تنبؤ بالسلاسل الزمنية لتنبيهات الأسعار"], [lu.LuMessageCircle, "مساعد بنموذج لغوي يفسر سبب الترتيب"]];
  for (let i = 0; i < 4; i++) {
    const y = 2.75 + i * 0.92;
    s.addImage({ data: await icon(tech[i][0], ICE), x: R - 0.65, y: y + 0.05, w: 0.4, h: 0.4 });
    s.addText(tech[i][1], ar({ x: 9.45, y, w: R - 9.45 - 0.8, h: 0.75, fontSize: 14, color: WHITE }));
  }
  s.addNotes("هذه مصادر البيانات المقترحة وآلية الحصول على كل منها، والتقنيات المستخدمة: خرائط حية، وواجهات برمجية، ونموذج تنبؤ للأسعار، ومساعد ذكي.");

  // =============== 12. Implementation plan (Gantt) ===============
  pres.addSection({ title: "التنفيذ" });
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "التنفيذ" });
  s.addText("خطة التنفيذ: خمس مراحل خلال خمس سنوات", { placeholder: "title" });
  subtitle(s, "نبدأ بصنع مقياس الخطر السعودي، ثم نبني عليه المنصة");
  {
    const labelW = 3.7, gx0 = M, gx1 = R - labelW - 0.2, gw = gx1 - gx0, yw = gw / 5, top = 2.05;
    card(s, M, top - 0.1, W - 2 * M, 4.55);
    // year headers (RTL: year 1 on the right)
    for (let yv = 1; yv <= 5; yv++) {
      const x = gx1 - yv * yw;
      s.addText(`السنة ${yv}`, { x, y: top + 0.05, w: yw, h: 0.4, fontFace: F, fontSize: 13, bold: true, color: MUTED, align: "center", rtlMode: true, margin: 0, isTextBox: true });
      if (yv < 5) s.addShape(pres.shapes.LINE, { x, y: top + 0.5, w: 0, h: 4.0, line: { color: LINE, width: 1, dashType: "dash" } });
    }
    const phases = [
      ["جمع المعلومات وصنع مقياس الخطر", "بيانات الموانئ والحوادث وJWC وAIS", 0, 2, TEAL],
      ["الاتفاق مع شركات التأمين", "مؤمّنان على الأقل", 1, 2, "13315C"],
      ["بناء الموقع وتأمينه", "بناء المنصة وحمايتها من الاختراق", 1, 3, "13315C"],
      ["إطلاق تجريبي", "مع شركات محدودة", 2, 3, ORANGE],
      ["الإطلاق الرسمي", "اعتماد تدريجي بجانب القوائم الدولية", 3, 5, GREEN],
    ];
    phases.forEach((p, i) => {
      const y = top + 0.6 + i * 0.8;
      numBadge(s, i + 1, R - 0.55, y + 0.08, 0.4, p[4]);
      s.addText(p[0], ar({ x: gx1 + 0.15, y: y + 0.02, w: labelW - 0.75, h: 0.55, fontSize: 14, bold: true, color: NAVY, valign: "middle" }));
      const bx = gx1 - p[3] * yw + 0.06, bw = (p[3] - p[2]) * yw - 0.12;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: bx, y: y + 0.05, w: bw, h: 0.5, rectRadius: 0.08, fill: { color: p[4] }, line: { color: p[4] } });
      s.addText(p[1], ar({ x: bx + 0.1, y: y + 0.05, w: bw - 0.2, h: 0.5, fontSize: 12, color: WHITE, valign: "middle", align: "center" }));
    });
  }
  s.addText([{ text: "الجهة المالكة المقترحة: ", options: { bold: true, color: NAVY } }, { text: "مشروع مشترك بين هيئة التأمين والهيئة العامة للنقل تحت وزارة النقل والخدمات اللوجستية، بالتنسيق مع موانئ والشركة السعودية لإعادة التأمين.", options: { color: MUTED } }], ar({ x: M, y: 6.6, w: W - 2 * M, h: 0.35, fontSize: 12 }));
  s.addNotes("الخطة خمس سنوات: أولاً نجمع المعلومات ونصنع المقياس، ثم نتفق مع شركات التأمين، ونبني الموقع، ثم إطلاق تجريبي في السنة الثالثة، والإطلاق الرسمي في السنتين الرابعة والخامسة. والمؤشر يُستخدم بالتوازي مع القوائم الدولية حتى يثبت اعتماده.");

  // =============== 13. Risks & mitigations ===============
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "التنفيذ" });
  s.addText("التحديات المتوقعة وكيف نعالجها", { placeholder: "title" });
  const risks = [
    ["تردد المؤمّنين في المشاركة", "المنصة قناة لعملاء جدد، والبدء باتفاقيات تجريبية محدودة"],
    ["تغير الأسعار وانتهاء صلاحية العرض", "السعر تقديري مع وقته وصلاحيته، والتأكيد النهائي من المؤمّن"],
    ["حساسية بيانات الخطر", "الاعتماد على قوائم منشورة ومصادر رسمية، وتقييد صلاحيات الوصول"],
    ["حماية بيانات المستخدمين", "سياسة حوكمة وفق نظام حماية البيانات الشخصية"],
    ["تحيز التوصيات", "عرض أسباب الترتيب للمستخدم، ومنع الترتيب المدفوع"],
    ["ترسيخ مرساة مرجعاً عالمياً", "بناء الثقة محلياً ثم إقليمياً، ومراجعة مستقلة للمنهجية"],
  ];
  const rw = (W - 2 * M - 0.35) / 2;
  const shieldG = await icon(lu.LuShieldCheck, GREEN);
  risks.forEach((r, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = R - rw - col * (rw + 0.35), y = 1.6 + row * 1.7;
    card(s, x, y, rw, 1.5);
    s.addText([{ text: "الخطر", options: { color: RED, bold: true, fontSize: 11, breakLine: true } }, { text: r[0], options: { color: NAVY, bold: true, fontSize: 16 } }], ar({ x: x + 0.3, y: y + 0.15, w: rw - 0.6, h: 0.65 }));
    s.addImage({ data: shieldG, x: x + rw - 0.6, y: y + 0.93, w: 0.3, h: 0.3 });
    s.addText(r[1], ar({ x: x + 0.3, y: y + 0.85, w: rw - 1.0, h: 0.5, fontSize: 13, color: MUTED, valign: "middle" }));
  });
  s.addNotes("توقعنا أبرز المخاطر وعالجناها: تردد المؤمّنين، وتغير الأسعار، وحساسية البيانات، وحماية بيانات المستخدمين، وتحيز التوصيات، وبناء الثقة في المؤشر.");

  // =============== 14. Impact & KPIs ===============
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "صلة المسار والأثر" });
  s.addText("الأثر المتوقع ومؤشرات القياس", { placeholder: "title" });
  subtitle(s, "قرارات تأمين أسرع وأوضح، وسوق تأمين بحري محلي أكثر ظهوراً، وبيانات وطنية عن المسارات");
  const who = [
    ["المصدّرون والمستوردون", "مقارنة المؤمّنين وتكلفة كل مسار في مكان واحد"],
    ["شركات التأمين المحلية", "قناة رقمية للتعريف بخدماتها والوصول لعملاء جدد"],
    ["الوعاء الوطني", "واجهة تسهّل وصول المستفيدين إلى المؤمّنين المشاركين"],
    ["الدولة", "معرفة وطنية بالمخاطر البحرية تزداد قيمتها بالتراكم"],
  ];
  who.forEach((w2, i) => {
    const y = 1.95 + i * 0.95;
    card(s, 6.95, y, R - 6.95, 0.82);
    s.addText(w2[0], ar({ x: 9.9, y, w: R - 10.1, h: 0.82, fontSize: 14, bold: true, color: TEAL_D, valign: "middle" }));
    s.addText(w2[1], ar({ x: 7.1, y, w: 2.7, h: 0.82, fontSize: 12, color: MUTED, valign: "middle" }));
  });
  // market stat
  card(s, 6.95, 5.8, R - 6.95, 0.85, ICE_BG, ICE_BG);
  s.addText("29.8 مليار ريال", ar({ x: 10.3, y: 5.8, w: R - 10.45, h: 0.85, fontSize: 18, bold: true, color: TEAL_D, valign: "middle" }));
  s.addText("أقساط سوق التأمين في الربع الأول 2026 بنمو 14.4%، وحصة التأمين البحري والنقل نحو 4.5% فقط: مساحة للنمو", ar({ x: 7.1, y: 5.8, w: 3.1, h: 0.85, fontSize: 11, color: INK, valign: "middle" }));
  // KPIs
  card(s, M, 1.95, 6.05, 4.7, NAVY, NAVY);
  s.addText("مؤشرات القياس", ar({ x: M + 0.3, y: 2.1, w: 5.45, h: 0.45, fontSize: 18, bold: true, color: ICE }));
  const kpi = ["عدد شركات التأمين المشاركة", "عدد الموانئ والمسارات المغطاة", "زمن الحصول على مقارنة موحدة", "المستخدمون النشطون والمشاريع المسجلة", "طلبات العروض الموجهة للمؤمّنين المحليين"];
  const chk = await icon(lu.LuCircleCheck, ICE);
  kpi.forEach((k, i) => {
    const y = 2.7 + i * 0.76;
    s.addImage({ data: chk, x: M + 5.4, y: y + 0.12, w: 0.35, h: 0.35 });
    s.addText(k, ar({ x: M + 0.3, y, w: 4.95, h: 0.6, fontSize: 15, color: WHITE, valign: "middle" }));
  });
  s.addNotes("المستفيد الأول من يشحن عبر الموانئ السعودية، ثم شركات التأمين والجهات المنظمة. ونقيس النجاح بخمسة مؤشرات. ولا نقدّر وفورات مالية قبل المرحلة التجريبية.");

  // =============== 15. Team ===============
  pres.addSection({ title: "الختام" });
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "الختام" });
  s.addText("الفريق وأدواره", { placeholder: "title" });
  const team = [["قائد الفريق", "التنسيق والتسليم وتمثيل الفريق", lu.LuUserRound], ["المجال والأعمال", "القيمة والجدوى وعلاقات المؤمّنين", lu.LuBriefcase], ["التقنية والبيانات", "الحل والنموذج الأولي ومصادر البيانات", lu.LuCode], ["التصميم وتجربة المستخدم", "الرحلة والواجهات", lu.LuPalette]];
  const tw = (W - 2 * M - 3 * 0.3) / 4;
  for (let i = 0; i < 4; i++) {
    const x = R - tw - i * (tw + 0.3), y = 1.8;
    card(s, x, y, tw, 4.6);
    await iconCircle(s, team[i][2], x + tw / 2 - 0.7, y + 0.45, 1.4, i === 0 ? NAVY : ICE_BG, i === 0 ? ICE : TEAL_D);
    s.addText("الاسم", { x: x + 0.2, y: y + 2.1, w: tw - 0.4, h: 0.5, fontFace: F, fontSize: 18, bold: true, color: NAVY, align: "center", rtlMode: true, margin: 0, isTextBox: true });
    s.addText(team[i][0], { x: x + 0.2, y: y + 2.65, w: tw - 0.4, h: 0.4, fontFace: F, fontSize: 15, bold: true, color: TEAL, align: "center", rtlMode: true, margin: 0, isTextBox: true });
    s.addText(team[i][1], { x: x + 0.25, y: y + 3.15, w: tw - 0.5, h: 0.8, fontFace: F, fontSize: 13, color: MUTED, align: "center", rtlMode: true, margin: 0, isTextBox: true });
  }
  s.addNotes("نعرّف بالفريق وأدوار كل عضو.");

  // =============== 16. Closing ===============
  s = pres.addSlide({ masterName: "COVER", sectionTitle: "الختام" });
  s.background = { color: DEEP };
  s.addImage({ path: path.join(SHOTS, "map.png"), x: 0, y: 0, w: 6.0, h: 7.5, sizing: { type: "cover", w: 6.0, h: 7.5 } });
  await iconCircle(s, lu.LuAnchor, R - 1.0, 1.3, 1.0, "13315C", ICE);
  s.addText("مرساة", ar({ x: 6.6, y: 2.45, w: R - 6.6, h: 1.0, fontSize: 54, bold: true, color: WHITE }));
  s.addText("قراءة سعودية للخطر البحري، وقرار تأمين أوضح وأسرع", ar({ x: 6.6, y: 3.5, w: R - 6.6, h: 0.9, fontSize: 24, color: ICE }));
  s.addText("نحو مركز لوجستي عالمي يملك معلوماته، انسجاماً مع رؤية 2030", ar({ x: 6.6, y: 4.45, w: R - 6.6, h: 0.6, fontSize: 16, color: SOFT }));
  s.addText("شكراً لكم", ar({ x: 6.6, y: 5.6, w: R - 6.6, h: 0.7, fontSize: 30, bold: true, color: WHITE }));
  s.addNotes("مرساة: قراءة سعودية للخطر البحري وقرار تأمين أوضح. شكراً لكم، ويسعدنا الإجابة عن أسئلتكم.");

  // remove the doubled title-slide image (first cropped call kept only as fallback)
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
