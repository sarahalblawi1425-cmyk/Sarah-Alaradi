// Mirsah (مرساة) pitch deck v2 — light & minimal
const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const path = require("path");
const lu = require("react-icons/lu");
const { applyTheme } = require(process.env.SKILL + "/scripts/apply_theme.js");

const SHOTS = process.argv[2], OUT = process.argv[3];

const INK = "1C1C1E", GRAY = "6B7280", LIGHT = "F3F4F6", HAIR = "E5E7EB", WHITE = "FFFFFF",
  ACC = "D9480F", ACC_BG = "FDEDE4", SEA = "0F4C5C";
const F = "Arial";
const THEME = { name: "Mirsah Light", headFontFace: F, bodyFontFace: F,
  colors: { dk1: INK, lt1: WHITE, dk2: SEA, lt2: LIGHT, accent1: ACC, accent2: SEA, accent3: "E0A526", accent4: "2E7D4F", accent5: "B42318", accent6: GRAY, hlink: SEA, folHlink: GRAY } };

const W = 13.333, H = 7.5, M = 0.75, R = W - M, CW = W - 2 * M;
const ar = (o = {}) => Object.assign({ fontFace: F, rtlMode: true, lang: "ar-SA", align: "right", isTextBox: true, margin: 0, valign: "top", color: INK }, o);
const ctr = (o = {}) => ar(Object.assign({ align: "center" }, o));

async function icon(Comp, color, size = 256) {
  const svg = RDS.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: String(size), strokeWidth: 1.6 }));
  return "image/png;base64," + (await sharp(Buffer.from(svg)).resize(size, size).png().toBuffer()).toString("base64");
}

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.rtlMode = true;
pres.title = "مرساة — منصة وطنية للتأمين البحري وتقييم مخاطر المسارات";
pres.theme = { headFontFace: F, bodyFontFace: F };

pres.defineSlideMaster({
  title: "CONTENT", background: { color: WHITE },
  objects: [
    { text: { text: "مرساة", options: ar({ x: R - 1.5, y: 6.95, w: 1.5, h: 0.3, fontSize: 11, bold: true, color: ACC }) } },
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.55, w: CW, h: 0.9, fontFace: F, fontSize: 36, bold: true, color: INK, align: "right", valign: "middle", rtlMode: true, margin: 0 }, text: "" } },
  ],
  slideNumber: { x: M, y: 6.95, w: 0.5, h: 0.3, fontFace: F, fontSize: 11, color: GRAY },
});
pres.defineSlideMaster({ title: "BLANK", background: { color: WHITE }, objects: [] });

async function bubble(s, Comp, x, y, d, bg = ACC_BG, fg = ACC) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: bg }, line: { color: bg } });
  const p = d * 0.26;
  s.addImage({ data: await icon(Comp, fg), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
}
function shot(s, file, x, y, w, h, alt) {
  s.addImage({ path: path.join(SHOTS, file), x, y, w, h, altText: alt, rounding: false,
    shadow: { type: "outer", color: "000000", opacity: 0.15, blur: 18, offset: 4, angle: 90 } });
}
const lede = (s, t, y = 1.45) => s.addText(t, ar({ x: M, y, w: CW, h: 0.5, fontSize: 20, color: GRAY }));

(async () => {
  // 1 — Cover
  pres.addSection({ title: "البداية" });
  let s = pres.addSlide({ masterName: "BLANK", sectionTitle: "البداية" });
  shot(s, "map.png", M, 1.0, 5.5 * 1736 / 1460 * 0.92, 5.5 * 0.92, "خريطة رادار الخطر من النموذج الأولي");
  await bubble(s, lu.LuAnchor, R - 0.95, 1.35, 0.95);
  s.addText("مرساة", ar({ x: 6.9, y: 2.45, w: R - 6.9, h: 1.4, fontSize: 80, bold: true }));
  s.addText("منصة وطنية للتأمين البحري\nوتقييم مخاطر المسارات", ar({ x: 6.9, y: 3.95, w: R - 6.9, h: 1.2, fontSize: 24, color: GRAY }));
  s.addText("نقل ثون 5  ·  المسار الرابع", ar({ x: 6.9, y: 5.65, w: R - 6.9, h: 0.4, fontSize: 16, bold: true, color: ACC }));
  s.addNotes("أهلاً بكم. نقدم لكم مرساة: منصة وطنية مقترحة تبني مؤشر خطر سعودي للمسارات البحرية، وتجمع للمصدّر والمستورد الخطر والتكلفة وعروض المؤمّنين المحليين في مكان واحد. المسار الرابع: الابتكار في الخدمات اللوجستية وسلاسل الإمداد.");

  // 2 — One big number
  pres.addSection({ title: "التحدي" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "التحدي" });
  s.addText("تأمين الحرب البحري قفز بشكل حاد", { placeholder: "title" });
  const big = [["100 ألف", "دولار · قبل الحرب", GRAY], ["3 ملايين", "دولار · رحلة من ينبع", INK], ["7 ملايين", "دولار · من موانئ جنوبية", ACC]];
  big.forEach((b, i) => {
    const w = CW / 3, x = R - w * (i + 1);
    s.addText(b[0], ctr({ x, y: 2.5, w, h: 1.3, fontSize: i === 2 ? 60 : 48, bold: true, color: b[2], valign: "bottom" }));
    s.addText(b[1], ctr({ x, y: 3.95, w, h: 0.5, fontSize: 20, color: GRAY }));
  });
  s.addText("تكلفة تأمين مخاطر الحرب لرحلة ناقلة واحدة، والأقساط تُراجع كل 24 ساعة", ctr({ x: M, y: 5.3, w: CW, h: 0.5, fontSize: 18, color: INK }));
  s.addText("المصدر: رويترز عبر Insurance Journal، سبتمبر 2026", ctr({ x: M, y: 5.85, w: CW, h: 0.35, fontSize: 12, color: GRAY }));
  s.addNotes("عندما تعبر السفينة منطقة توتر أمني تُؤمَّن بقسط إضافي. قسط ناقلات ينبع ارتفع من أقل من 1% إلى نحو 3% من قيمة السفينة، وقد يصل إلى 7% لموانئ جنوب ينبع مثل جازان. والرحلة التي كانت تكلف 100 ألف دولار أصبحت 3 إلى 7 ملايين.");

  // 3 — The pain
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "التحدي" });
  s.addText("والشاحن يقرر وهو لا يرى الصورة كاملة", { placeholder: "title" });
  const pains = [[lu.LuGlobe, "مرجع دولي واحد", "لا يوجد مؤشر سعودي للخطر"], [lu.LuScale, "لا مقارنة", "الأسعار غير معلنة"], [lu.LuPhoneCall, "شركة شركة", "يتصل بكل مؤمّن على حدة"]];
  for (let i = 0; i < 3; i++) {
    const w = CW / 3, x = R - w * (i + 1);
    await bubble(s, pains[i][0], x + w / 2 - 0.75, 2.2, 1.5);
    s.addText(pains[i][1], ctr({ x, y: 4.0, w, h: 0.6, fontSize: 26, bold: true }));
    s.addText(pains[i][2], ctr({ x, y: 4.65, w, h: 0.5, fontSize: 18, color: GRAY }));
  }
  s.addNotes("المشكلة ليست السعر فقط: السوق يرجع إلى قوائم لجنة الحرب المشتركة في لندن، ولا تملك المملكة مؤشراً وطنياً. الأسعار المتفق عليها غير معلنة فيصعب المقارنة. والشاحن يتصل بكل شركة على حدة ولا يرى كيف يتغير القسط إذا غيّر الميناء أو المسار.");

  // 4 — Solution
  pres.addSection({ title: "الحل" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "الحل" });
  s.addText("مرساة: مؤشر خطر سعودي في منصة واحدة", { placeholder: "title" });
  const flow = [[lu.LuRadar, "الخطر", "درجة لكل منطقة ومسار"], [lu.LuReceipt, "التكلفة", "قسط تقديري لكل مسار"], [lu.LuCircleCheck, "القرار", "مقارنة وتنبيه بالتوقيت"]];
  for (let i = 0; i < 3; i++) {
    const w = CW / 3, x = R - w * (i + 1);
    await bubble(s, flow[i][0], x + w / 2 - 0.85, 2.15, 1.7, i === 2 ? ACC : ACC_BG, i === 2 ? WHITE : ACC);
    s.addText(flow[i][1], ctr({ x, y: 4.1, w, h: 0.6, fontSize: 28, bold: true }));
    s.addText(flow[i][2], ctr({ x, y: 4.75, w, h: 0.5, fontSize: 18, color: GRAY }));
    if (i < 2) s.addShape(pres.shapes.LINE, { x: x - 0.55, y: 3.0, w: 0.9, h: 0, line: { color: HAIR, width: 2, endArrowType: "triangle" }, flipH: true });
  }
  s.addText("يعمل بجوار القوائم الدولية، والتأكيد النهائي للسعر من شركة التأمين", ctr({ x: M, y: 5.8, w: CW, h: 0.4, fontSize: 16, color: GRAY }));
  s.addNotes("فكرتنا واحدة: أن تبني المملكة مؤشر خطر سعودي خاص بمساراتها البحرية، يعتمد على قوائم الخطر المنشورة وتتبع السفن وبيانات الموانئ والمؤمّنين. ويحوّله الموقع إلى ثلاث نتائج: درجة خطر، وقسط تقديري، وقرار واضح بالمقارنة والتنبيه. المؤشر لا ينافس المراجع الدولية بل يعمل بجوارها.");

  // 5 — Prototype 1
  pres.addSection({ title: "النموذج الأولي" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "النموذج الأولي" });
  s.addText("اختر ميناءيك، وشاهد خطر كل مسار", { placeholder: "title" });
  { const h = 5.0, w = h * 2544 / 1460; shot(s, "search_row.png", (W - w) / 2, 1.65, w, h, "لقطة من النموذج الأولي: المسارات ورادار الخطر"); }
  s.addNotes("هذه الشاشة الأولى من النموذج الأولي. يدخل المستورد ميناء الانطلاق والوصول، فتظهر المسارات المتاحة، ولكل مسار درجة خطر ومدة ومسافة. وعلى اليسار رادار حي لمناطق الخطر وحركة السفن: باب المندب 7.4 حرج، وهرمز 5.8 مرتفع. الأرقام توضيحية.");

  // 6 — Prototype 2
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "النموذج الأولي" });
  s.addText("قارن عروض المؤمّنين السعوديين، واعرف متى تؤمّن", { placeholder: "title" });
  { const h = 5.0, w = h * 2880 / 1880; shot(s, "prices_top.png", (W - w) / 2, 1.65, w, h, "لقطة من النموذج الأولي: صفحة تفقد الأسعار"); }
  s.addNotes("الشاشة الثانية: يختار المسار ويكتب قيمة الشحنة، فيرى أقل قسط متاح، وتغيّر السعر خلال 14 يوماً، وتوصية بتوقيت التأمين مع درجة ثقة، وتنبيه عند نزول القسط. وأسفل الصفحة جدول عروض المؤمّنين السعوديين دون ترتيب مدفوع.");

  // 7 — Journey
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "النموذج الأولي" });
  s.addText("رحلة المستخدم", { placeholder: "title" });
  const steps = [[lu.LuRadar, "الرادار", "المسار والخطر"], [lu.LuShieldCheck, "المؤمّنون", "الأسعار والوثائق"], [lu.LuSparkles, "الذكاء الاصطناعي", "توقيت وتوصية"], [lu.LuBellRing, "مشاريعي", "تنبيه بالسعر"]];
  for (let i = 0; i < 4; i++) {
    const w = CW / 4, x = R - w * (i + 1);
    await bubble(s, steps[i][0], x + w / 2 - 0.7, 2.3, 1.4);
    s.addText(String(i + 1), ctr({ x, y: 1.75, w, h: 0.4, fontSize: 16, bold: true, color: ACC }));
    s.addText(steps[i][1], ctr({ x, y: 4.0, w, h: 0.6, fontSize: 24, bold: true }));
    s.addText(steps[i][2], ctr({ x, y: 4.6, w, h: 0.5, fontSize: 18, color: GRAY }));
  }
  s.addText("مناطق الخطر تُحدَّد من مصادر رسمية، لا بالذكاء الاصطناعي", ctr({ x: M, y: 5.8, w: CW, h: 0.4, fontSize: 16, color: GRAY }));
  s.addNotes("أربع محطات: الرادار يقترح المسار ودرجة خطره، ثم صفحة شركات التأمين بأنواع الشحنات والوثائق، ثم الذكاء الاصطناعي في ثلاث وظائف: تنبيه بالتوقيت، وتوصية بالمؤمّن مع السبب، ومساعد يجيب الأسئلة. وأخيراً مشاريعي: يحفظ الشحنة ويصله إشعار عند تغير القسط.");

  // 8 — Why Saudi index
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "الحل" });
  s.addText("لماذا مؤشر سعودي؟", { placeholder: "title" });
  const why = [[lu.LuDatabase, "نملك بياناتنا"], [lu.LuFlag, "استقلال لوجستي"], [lu.LuEarth, "مرجع عالمي مستقبلاً"], [lu.LuTrainFront, "يكمل منظومة النقل"]];
  for (let i = 0; i < 4; i++) {
    const w = CW / 4, x = R - w * (i + 1);
    await bubble(s, why[i][0], x + w / 2 - 0.7, 2.4, 1.4, LIGHT, SEA);
    s.addText(why[i][1], ctr({ x: x + 0.1, y: 4.1, w: w - 0.2, h: 0.9, fontSize: 24, bold: true }));
  }
  s.addNotes("لأن المملكة وهي تبني مكانتها مركزاً لوجستياً عالمياً تحتاج أن تملك قراءتها الوطنية للخطر. من يملك البيانات يحللها ويبني الحلول عليها. ويقلل الاعتماد على المراجع الخارجية وحدها. ويبدأ بالمناطق المحيطة بالمملكة ثم يتوسع ليصبح مرجعاً تستفيد منه دول أخرى. ويتكامل مع القطارات والموانئ والبنية اللوجستية.");

  // 9 — Differentiation (contrast)
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "الحل" });
  s.addText("ما الذي يميز مرساة؟", { placeholder: "title" });
  const colW = (CW - 0.5) / 2;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y: 1.9, w: colW, h: 4.3, rectRadius: 0.15, fill: { color: LIGHT }, line: { color: LIGHT } });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: R - colW, y: 1.9, w: colW, h: 4.3, rectRadius: 0.15, fill: { color: ACC_BG }, line: { color: ACC_BG } });
  s.addText("الحلول القائمة", ar({ x: M + 0.5, y: 2.3, w: colW - 1, h: 0.6, fontSize: 26, bold: true, color: GRAY }));
  s.addText("مرساة", ar({ x: R - colW + 0.5, y: 2.3, w: colW - 1, h: 0.6, fontSize: 26, bold: true, color: ACC }));
  const them = ["عالمية", "مؤمّن واحد أو منتج واحد", "تعالج جانباً واحداً"];
  const us = ["سعودية، ببيانات محلية", "مقارنة محايدة بين عدة مؤمّنين", "الخطر والتكلفة والمقارنة معاً"];
  s.addText(them.map((t, i) => ({ text: t, options: { breakLine: i < 2 } })), ar({ x: M + 0.5, y: 3.15, w: colW - 1, h: 2.6, fontSize: 22, color: INK, paraSpaceAfter: 16 }));
  s.addText(us.map((t, i) => ({ text: t, options: { breakLine: i < 2, bold: true } })), ar({ x: R - colW + 0.5, y: 3.15, w: colW - 1, h: 2.6, fontSize: 22, color: INK, paraSpaceAfter: 16 }));
  s.addNotes("درسنا الحلول القائمة: ShipWRITE من Canopius، وWarRates، وInsurwave مع Ambrey، وتأمين حرب البضائع من DP World. كلها قيّمة لكنها عالمية أو تعالج جانباً واحداً. ولم نجد منصة سعودية تجمع خطر المسار وتكلفته ومقارنة المؤمّنين المحليين. ومرساة تكمل الوعاء الوطني لتأمين مخاطر الحرب ولا تنافسه.");

  // 10 — Track fit
  pres.addSection({ title: "المسار والأثر" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "المسار والأثر" });
  s.addText("سؤال المسار الرابع", { placeholder: "title" });
  s.addText("«كيف يمكن تقليل تكلفة التوصيل ووقته؟»", ctr({ x: M, y: 1.9, w: CW, h: 0.9, fontSize: 34, bold: true, color: SEA }));
  s.addText("مرساة تُظهر التكلفة الكلية لكل مسار قبل القرار، وتنبّه بأفضل وقت للتأمين", ctr({ x: M, y: 2.95, w: CW, h: 0.6, fontSize: 20, color: GRAY }));
  s.addShape(pres.shapes.LINE, { x: W / 2 - 1.5, y: 3.95, w: 3, h: 0, line: { color: HAIR, width: 1 } });
  s.addText("40 مليون", ctr({ x: M, y: 4.25, w: CW, h: 1.2, fontSize: 64, bold: true, color: ACC }));
  s.addText("حاوية سنوياً هدف المملكة بحلول 2030، مقابل 8.3 مليون في 2025", ctr({ x: M, y: 5.5, w: CW, h: 0.5, fontSize: 18, color: GRAY }));
  s.addNotes("سؤال المسار: كيف نقلل تكلفة التوصيل ووقته؟ مرساة تجيب مباشرة: يقارن المستخدم بين المؤمّنين والمسارات والموانئ في مكان واحد، فيعرف التكلفة الكلية قبل أن يقرر، ويصله تنبيه بالوقت المناسب للتأمين. وتقليل الغموض في كلفة التأمين يشجع الشاحنين على الموانئ السعودية، دعماً لهدف 40 مليون حاوية بحلول 2030.");

  // 11 — Data sources
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "المسار والأثر" });
  s.addText("من أين تأتي البيانات؟", { placeholder: "title" });
  const src = [[lu.LuFileWarning, "قوائم الخطر", "لجنة الحرب المشتركة والجهات الوطنية"], [lu.LuShip, "تتبع السفن AIS", "بالشراكة مع موانئ"], [lu.LuBuilding2, "شركات التأمين", "اتفاقيات وواجهات برمجية"], [lu.LuLandmark, "الوعاء الوطني", "إعادة التأمين وهيئة التأمين"]];
  for (let i = 0; i < 4; i++) {
    const w = CW / 4, x = R - w * (i + 1);
    await bubble(s, src[i][0], x + w / 2 - 0.65, 2.2, 1.3, LIGHT, SEA);
    s.addText(src[i][1], ctr({ x: x + 0.1, y: 3.8, w: w - 0.2, h: 0.6, fontSize: 22, bold: true }));
    s.addText(src[i][2], ctr({ x: x + 0.2, y: 4.45, w: w - 0.4, h: 0.8, fontSize: 16, color: GRAY }));
  }
  s.addNotes("مصادر البيانات مقترحة وتتطلب اتفاقيات: قوائم لجنة الحرب المشتركة المنشورة وتنبيهات الجهات الوطنية، وتتبع السفن AIS بهدف بناء شبكة وطنية مع موانئ، وأسعار المؤمّنين عبر اتفاقيات وواجهات برمجية، وتفاصيل الوعاء الوطني من الشركة السعودية لإعادة التأمين وهيئة التأمين. والتقنيات: خرائط حية، ونموذج تنبؤ بالسلاسل الزمنية للأسعار، ومساعد بنموذج لغوي.");

  // 12 — Plan (simple timeline)
  pres.addSection({ title: "التنفيذ" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "التنفيذ" });
  s.addText("خطة خمس سنوات", { placeholder: "title" });
  const ph = [["السنة 1–2", "صنع مقياس الخطر"], ["السنة 2", "الاتفاق مع المؤمّنين"], ["السنة 2–3", "بناء المنصة وتأمينها"], ["السنة 3", "إطلاق تجريبي"], ["السنة 4–5", "الإطلاق الرسمي"]];
  const lw = CW / 5, ly = 3.35;
  s.addShape(pres.shapes.LINE, { x: M + lw / 2, y: ly, w: CW - lw, h: 0, line: { color: HAIR, width: 3 } });
  ph.forEach((p, i) => {
    const x = R - lw * (i + 1), last = i === 4;
    s.addShape(pres.shapes.OVAL, { x: x + lw / 2 - 0.22, y: ly - 0.22, w: 0.44, h: 0.44, fill: { color: last ? ACC : WHITE }, line: { color: ACC, width: 3 } });
    s.addText(p[0], ctr({ x, y: 2.3, w: lw, h: 0.5, fontSize: 18, bold: true, color: ACC }));
    s.addText(p[1], ctr({ x: x + 0.1, y: 3.85, w: lw - 0.2, h: 0.9, fontSize: 20, bold: true }));
  });
  s.addText("المالك المقترح: هيئة التأمين والهيئة العامة للنقل، تحت وزارة النقل والخدمات اللوجستية", ctr({ x: M, y: 5.6, w: CW, h: 0.5, fontSize: 16, color: GRAY }));
  s.addNotes("الخطة خمس مراحل خلال خمس سنوات: نجمع البيانات ونبني المقياس بمنهجية شفافة، ثم اتفاقيات مع مؤمّنين اثنين على الأقل، ثم بناء المنصة وحمايتها، ثم إطلاق تجريبي مع عدد محدود من الشركات، ثم الإطلاق الرسمي. المؤشر يُستخدم بالتوازي مع القوائم الدولية حتى يثبت اعتماده. ونقترح أن تكون مشروعاً مشتركاً بالتنسيق مع موانئ والشركة السعودية لإعادة التأمين.");

  // 13 — Risks
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "التنفيذ" });
  s.addText("التحديات وكيف نعالجها", { placeholder: "title" });
  const rk = [["تردد المؤمّنين", "قناة لعملاء جدد، وبداية تجريبية"], ["تغيّر الأسعار بسرعة", "سعر تقديري بوقته، والتأكيد من المؤمّن"], ["حساسية البيانات", "مصادر رسمية وصلاحيات مقيدة"], ["تحيّز التوصيات", "نعرض السبب، ولا ترتيب مدفوع"]];
  rk.forEach((r, i) => {
    const y = 1.95 + i * 1.08;
    if (i) s.addShape(pres.shapes.LINE, { x: M, y: y - 0.12, w: CW, h: 0, line: { color: HAIR, width: 1 } });
    s.addText(r[0], ar({ x: 7.4, y, w: R - 7.4, h: 0.8, fontSize: 24, bold: true, valign: "middle" }));
    s.addText(r[1], ar({ x: M, y, w: 6.2, h: 0.8, fontSize: 20, color: GRAY, valign: "middle" }));
  });
  s.addNotes("توقعنا المخاطر: تردد المؤمّنين نعالجه بتقديم المنصة قناة لعملاء جدد واتفاقيات تجريبية. تغير الأسعار بعرض السعر تقديرياً مع وقته وصلاحيته. حساسية البيانات بالاعتماد على مصادر رسمية وتقييد الوصول. حماية بيانات المستخدمين بسياسة حوكمة. وتحيز التوصيات بعرض أسباب الترتيب ومنع الترتيب المدفوع.");

  // 14 — Impact
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "المسار والأثر" });
  s.addText("الأثر المتوقع", { placeholder: "title" });
  const imp = [["الشاحن", "قرار أسرع وأوضح"], ["المؤمّن المحلي", "عملاء جدد"], ["الدولة", "معرفة وطنية بالمخاطر"]];
  imp.forEach((p, i) => {
    const w = CW / 3, x = R - w * (i + 1);
    s.addText(p[0], ctr({ x, y: 1.95, w, h: 0.5, fontSize: 18, color: GRAY }));
    s.addText(p[1], ctr({ x, y: 2.45, w, h: 0.7, fontSize: 26, bold: true }));
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y: 3.85, w: CW, h: 2.3, rectRadius: 0.15, fill: { color: LIGHT }, line: { color: LIGHT } });
  s.addText("29.8 مليار ريال", ar({ x: 7.0, y: 4.15, w: R - 7.4, h: 1.0, fontSize: 44, bold: true, color: ACC }));
  s.addText("سوق التأمين في الربع الأول 2026، والبحري والنقل نحو 4.5% فقط", ar({ x: 7.0, y: 5.15, w: R - 7.4, h: 0.8, fontSize: 16, color: GRAY }));
  s.addText("نقيس النجاح بـ", ar({ x: M + 0.4, y: 4.15, w: 5.6, h: 0.45, fontSize: 16, bold: true, color: SEA }));
  s.addText(["عدد المؤمّنين المشاركين", "الموانئ والمسارات المغطاة", "زمن الوصول لمقارنة موحدة"].map((t, i) => ({ text: t, options: { breakLine: i < 2 } })), ar({ x: M + 0.4, y: 4.65, w: 5.6, h: 1.35, fontSize: 17, paraSpaceAfter: 6 }));
  s.addNotes("المستفيد الأول من يشحن عبر الموانئ السعودية، ثم شركات التأمين المحلية التي تكسب قناة رقمية، والدولة التي تتراكم لديها بيانات وطنية عن المخاطر. سوق التأمين بلغ 29.8 مليار ريال في الربع الأول 2026 بنمو 14.4%، والتأمين البحري والنقل حصته صغيرة، ما يعني مساحة للنمو. ونقيس النجاح بعدد المؤمّنين والمسارات المغطاة وزمن المقارنة والمستخدمين النشطين وطلبات العروض. ولا نقدّر وفورات مالية قبل المرحلة التجريبية.");

  // 15 — Team
  pres.addSection({ title: "الختام" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "الختام" });
  s.addText("الفريق", { placeholder: "title" });
  const team = [[lu.LuUserRound, "قائد الفريق"], [lu.LuBriefcase, "المجال والأعمال"], [lu.LuCode, "التقنية والبيانات"], [lu.LuPalette, "التصميم وتجربة المستخدم"]];
  for (let i = 0; i < 4; i++) {
    const w = CW / 4, x = R - w * (i + 1);
    await bubble(s, team[i][0], x + w / 2 - 0.8, 2.2, 1.6, i === 0 ? ACC : LIGHT, i === 0 ? WHITE : SEA);
    s.addText("الاسم", ctr({ x, y: 4.1, w, h: 0.6, fontSize: 24, bold: true }));
    s.addText(team[i][1], ctr({ x: x + 0.1, y: 4.7, w: w - 0.2, h: 0.5, fontSize: 18, color: GRAY }));
  }
  s.addNotes("نعرّف بالفريق وأدوار كل عضو.");

  // 16 — Thanks
  s = pres.addSlide({ masterName: "BLANK", sectionTitle: "الختام" });
  await bubble(s, lu.LuAnchor, W / 2 - 0.6, 1.6, 1.2);
  s.addText("مرساة", ctr({ x: M, y: 3.0, w: CW, h: 1.3, fontSize: 72, bold: true }));
  s.addText("قراءة سعودية للخطر البحري، وقرار تأمين أوضح", ctr({ x: M, y: 4.35, w: CW, h: 0.6, fontSize: 24, color: GRAY }));
  s.addText("شكراً لكم", ctr({ x: M, y: 5.55, w: CW, h: 0.6, fontSize: 22, bold: true, color: ACC }));
  s.addNotes("مرساة: قراءة سعودية للخطر البحري، وقرار تأمين أوضح، انسجاماً مع رؤية 2030. شكراً لكم، ويسعدنا الإجابة عن أسئلتكم.");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
