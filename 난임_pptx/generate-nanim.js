"use strict";
const pptxgen = require("../node_modules/pptxgenjs");
const path = require("path");

const pptx = new pptxgen();
pptx.layout = "LAYOUT_16x9";
pptx.author = "의학 블로그";
pptx.title = "난임(不妊) 완전 가이드";

const FONT = "Pretendard";
const COLOR = {
  white: "FFFFFF",
  teal: "0A7E8C",
  navy: "1B2A4A",
  darkGray: "2D2D2D",
  medGray: "6B7280",
  lightTeal: "E6F3F5",
  lineGray: "D1D5DB",
  softRed: "DC3545",
  placeholder: "F0F0F0",
  placeholderText: "9CA3AF",
};

const IMG = {
  ivfProcess: path.resolve(__dirname, "img_ivf_process.png"),
  ivfTube: path.resolve(__dirname, "img_ivf_tube.jpg"),
  embryoGrades: path.resolve(__dirname, "img_embryo_grades.jpg"),
  blastocyst: path.resolve(__dirname, "img_blastocyst.png"),
  femaleOverview: path.resolve(__dirname, "img_female_overview.jpg"),
  uterusFallopian: path.resolve(__dirname, "img_uterus_fallopian.jpg"),
  uterus: path.resolve(__dirname, "img_uterus.jpg"),
  ovaryCross: path.resolve(__dirname, "img_ovary_crosssection.jpg"),
  femaleTract: path.resolve(__dirname, "img_female_tract.jpg"),
};

function addSlideTitle(slide, titleText) {
  slide.addText(titleText, {
    x: 0.5, y: 0.2, w: 9.0, h: 0.7,
    fontSize: 30,
    fontFace: FONT,
    bold: true,
    color: COLOR.navy,
    align: "left",
    valign: "middle",
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.5, y: 0.88, w: 1.4, h: 0.045,
    fill: { color: COLOR.teal },
    line: { type: "none" },
  });
}

function addPlaceholder(slide, x, y, w, h, label) {
  slide.addShape(pptx.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: COLOR.placeholder },
    line: { color: COLOR.lineGray, width: 1, dashType: "dash" },
  });
  slide.addText(label || "[이미지]", {
    x, y, w, h,
    align: "center", valign: "middle",
    fontSize: 12, color: COLOR.placeholderText, fontFace: FONT,
  });
}

// ============================================================
// 슬라이드 1: 표지
// ============================================================
{
  const slide = pptx.addSlide();
  slide.background = { fill: COLOR.navy };

  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.0, y: 0.0, w: 0.09, h: "100%",
    fill: { color: COLOR.teal }, line: { type: "none" },
  });

  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.6, y: 3.7, w: 9.0, h: 0.03,
    fill: { color: COLOR.teal }, line: { type: "none" },
  });

  slide.addText("난임(不妊) 완전 가이드", {
    x: 0.6, y: 1.0, w: 8.8, h: 1.4,
    fontSize: 44, fontFace: FONT, bold: true,
    color: COLOR.white, align: "left", valign: "bottom",
  });

  slide.addText("검사 · 인공수정 · 시험관 · 호르몬 · 한의 진료까지", {
    x: 0.6, y: 2.5, w: 8.8, h: 0.7,
    fontSize: 20, fontFace: FONT,
    color: COLOR.teal, align: "left",
  });

  slide.addText("AMH · 정액검사 · IUI · IVF · GnRH · PGT · 프로게스테론", {
    x: 0.6, y: 3.2, w: 8.8, h: 0.4,
    fontSize: 14, fontFace: FONT,
    color: COLOR.white, align: "left", transparency: 40,
  });

  slide.addText("2026.02.23  |  의학 블로그", {
    x: 0.6, y: 4.5, w: 8.8, h: 0.4,
    fontSize: 13, fontFace: FONT,
    color: COLOR.white, transparency: 50, align: "left",
  });
}

// ============================================================
// 슬라이드 2: 목차
// ============================================================
{
  const slide = pptx.addSlide();
  addSlideTitle(slide, "목차");

  const items = [
    "AMH 정상범위 & 정상 정액 검사 기준",
    "여성 난임 — 난소기능 평가(AMH) & 난자의 질",
    "남성 난임 — 정액 검사 해석",
    "난임 진료 흐름 — 난관조영술 → 인공수정 → 시험관",
    "시험관(IVF) 과배란 유도 — 단기 / 장기 / 저자극 요법",
    "과배란 유도 약제 상세 — FSH 단일 vs hMG(FSH+LH)",
    "GnRH 길항제 vs 작용제 비교",
    "hCG 주사 & 자궁내막 두께 · 임신성공률",
    "배아 배양 전략 & 배아 등급(포배기 · 눈사람 · 감자)",
    "PGT 검사 & 이식 결정 기준(E2 · P4)",
    "이식 & 프로게스테론 보충 & NK 수치",
    "한의 진료 접근 — 뜸 · 좌훈 · 한약",
  ];

  const col1 = items.slice(0, 6);
  const col2 = items.slice(6);

  col1.forEach((item, idx) => {
    const y = 1.1 + idx * 0.52;
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: 0.4, y: y, w: 0.42, h: 0.38,
      fill: { color: COLOR.teal }, line: { type: "none" }, rectRadius: 0.04,
    });
    slide.addText(String(idx + 1).padStart(2, "0"), {
      x: 0.4, y: y, w: 0.42, h: 0.38,
      fontSize: 13, fontFace: FONT, bold: true, color: COLOR.white,
      align: "center", valign: "middle",
    });
    slide.addText(item, {
      x: 0.9, y: y + 0.04, w: 3.9, h: 0.33,
      fontSize: 11.5, fontFace: FONT, color: COLOR.darkGray, valign: "middle",
    });
  });

  col2.forEach((item, idx) => {
    const y = 1.1 + idx * 0.52;
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: 5.1, y: y, w: 0.42, h: 0.38,
      fill: { color: COLOR.navy }, line: { type: "none" }, rectRadius: 0.04,
    });
    slide.addText(String(idx + 7).padStart(2, "0"), {
      x: 5.1, y: y, w: 0.42, h: 0.38,
      fontSize: 13, fontFace: FONT, bold: true, color: COLOR.white,
      align: "center", valign: "middle",
    });
    slide.addText(item, {
      x: 5.6, y: y + 0.04, w: 3.9, h: 0.33,
      fontSize: 11.5, fontFace: FONT, color: COLOR.darkGray, valign: "middle",
    });
  });
}

// ============================================================
// 슬라이드 3: AMH 정상범위 & 정상 정액 검사 기준
// ============================================================
{
  const slide = pptx.addSlide();
  addSlideTitle(slide, "AMH 정상범위 & 정상 정액 검사 기준");

  // AMH 표
  slide.addText("여성 — AMH(항뮬러관호르몬, Anti-Müllerian Hormone) 정상범위", {
    x: 0.4, y: 1.05, w: 4.4, h: 0.35,
    fontSize: 12, fontFace: FONT, bold: true, color: COLOR.teal,
  });

  const amhRows = [
    [{ text: "연령", options: { fontSize: 12, fontFace: FONT, bold: true, color: COLOR.white, fill: { color: COLOR.teal }, align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "정상범위 (ng/mL)", options: { fontSize: 12, fontFace: FONT, bold: true, color: COLOR.white, fill: { color: COLOR.teal }, align: "center", valign: "middle", margin: [4,6,4,6] }}],
    [{ text: "20~25세", options: { fontSize: 12, fontFace: FONT, color: COLOR.darkGray, align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "4.0 – 6.8", options: { fontSize: 12, fontFace: FONT, color: COLOR.darkGray, align: "center", valign: "middle", margin: [4,6,4,6] }}],
    [{ text: "26~30세", options: { fontSize: 12, fontFace: FONT, color: COLOR.darkGray, fill: { color: COLOR.lightTeal }, align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "3.0 – 6.3", options: { fontSize: 12, fontFace: FONT, color: COLOR.darkGray, fill: { color: COLOR.lightTeal }, align: "center", valign: "middle", margin: [4,6,4,6] }}],
    [{ text: "31~35세", options: { fontSize: 12, fontFace: FONT, color: COLOR.darkGray, align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "2.1 – 5.4", options: { fontSize: 12, fontFace: FONT, color: COLOR.darkGray, align: "center", valign: "middle", margin: [4,6,4,6] }}],
    [{ text: "36~40세", options: { fontSize: 12, fontFace: FONT, color: COLOR.darkGray, fill: { color: COLOR.lightTeal }, align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "1.1 – 3.8", options: { fontSize: 12, fontFace: FONT, color: COLOR.darkGray, fill: { color: COLOR.lightTeal }, align: "center", valign: "middle", margin: [4,6,4,6] }}],
    [{ text: "41~45세", options: { fontSize: 12, fontFace: FONT, color: COLOR.darkGray, align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "0.6 – 2.1", options: { fontSize: 12, fontFace: FONT, color: COLOR.darkGray, align: "center", valign: "middle", margin: [4,6,4,6] }}],
    [{ text: "46세 이상", options: { fontSize: 12, fontFace: FONT, color: COLOR.darkGray, fill: { color: COLOR.lightTeal }, align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "0.1 – 0.5", options: { fontSize: 12, fontFace: FONT, color: COLOR.darkGray, fill: { color: COLOR.lightTeal }, align: "center", valign: "middle", margin: [4,6,4,6] }}],
  ];
  slide.addTable(amhRows, {
    x: 0.4, y: 1.45, w: 4.4,
    border: { type: "solid", color: COLOR.lineGray, pt: 0.5 },
    colW: [2.0, 2.4], rowH: 0.42,
  });

  // 정액 검사 기준 표
  slide.addText("남성 — 정상 정액 검사 기준 (WHO 기준)", {
    x: 5.1, y: 1.05, w: 4.5, h: 0.35,
    fontSize: 12, fontFace: FONT, bold: true, color: COLOR.teal,
  });

  const semenRows = [
    [{ text: "검사 항목", options: { fontSize: 11.5, fontFace: FONT, bold: true, color: COLOR.white, fill: { color: COLOR.navy }, align: "center", valign: "middle", margin: [4,4,4,4] }},
     { text: "정상 기준치", options: { fontSize: 11.5, fontFace: FONT, bold: true, color: COLOR.white, fill: { color: COLOR.navy }, align: "center", valign: "middle", margin: [4,4,4,4] }}],
    [{ text: "정자 농도", options: { fontSize: 11, fontFace: FONT, bold: true, color: COLOR.darkGray, fill: { color: COLOR.lightTeal }, align: "left", valign: "middle", margin: [4,6,4,6] }},
     { text: "1mL당 1,500만 개 이상", options: { fontSize: 11, fontFace: FONT, color: COLOR.darkGray, fill: { color: COLOR.lightTeal }, align: "left", valign: "middle", margin: [4,6,4,6] }}],
    [{ text: "정자 총수", options: { fontSize: 11, fontFace: FONT, bold: true, color: COLOR.darkGray, align: "left", valign: "middle", margin: [4,6,4,6] }},
     { text: "3,900만 개 이상 (전체 정액)", options: { fontSize: 11, fontFace: FONT, color: COLOR.darkGray, align: "left", valign: "middle", margin: [4,6,4,6] }}],
    [{ text: "전진 운동성", options: { fontSize: 11, fontFace: FONT, bold: true, color: COLOR.darkGray, fill: { color: COLOR.lightTeal }, align: "left", valign: "middle", margin: [4,6,4,6] }},
     { text: "32% 이상", options: { fontSize: 11, fontFace: FONT, color: COLOR.darkGray, fill: { color: COLOR.lightTeal }, align: "left", valign: "middle", margin: [4,6,4,6] }}],
    [{ text: "전체 운동성", options: { fontSize: 11, fontFace: FONT, bold: true, color: COLOR.darkGray, align: "left", valign: "middle", margin: [4,6,4,6] }},
     { text: "42% 이상", options: { fontSize: 11, fontFace: FONT, color: COLOR.darkGray, align: "left", valign: "middle", margin: [4,6,4,6] }}],
    [{ text: "정자 형태", options: { fontSize: 11, fontFace: FONT, bold: true, color: COLOR.darkGray, fill: { color: COLOR.lightTeal }, align: "left", valign: "middle", margin: [4,6,4,6] }},
     { text: "정상 형태 4% 이상", options: { fontSize: 11, fontFace: FONT, color: COLOR.darkGray, fill: { color: COLOR.lightTeal }, align: "left", valign: "middle", margin: [4,6,4,6] }}],
    [{ text: "정액 부피", options: { fontSize: 11, fontFace: FONT, bold: true, color: COLOR.darkGray, align: "left", valign: "middle", margin: [4,6,4,6] }},
     { text: "1.4mL 이상", options: { fontSize: 11, fontFace: FONT, color: COLOR.darkGray, align: "left", valign: "middle", margin: [4,6,4,6] }}],
    [{ text: "액화시간 / pH", options: { fontSize: 11, fontFace: FONT, bold: true, color: COLOR.darkGray, fill: { color: COLOR.lightTeal }, align: "left", valign: "middle", margin: [4,6,4,6] }},
     { text: "60분 이내 / pH ≥7.2", options: { fontSize: 11, fontFace: FONT, color: COLOR.darkGray, fill: { color: COLOR.lightTeal }, align: "left", valign: "middle", margin: [4,6,4,6] }}],
  ];
  slide.addTable(semenRows, {
    x: 5.1, y: 1.45, w: 4.5,
    border: { type: "solid", color: COLOR.lineGray, pt: 0.5 },
    colW: [1.7, 2.8], rowH: 0.42,
  });

  // 하단 주의 박스
  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 0.4, y: 4.65, w: 9.2, h: 0.52,
    fill: { color: "FFF3E0" }, line: { color: "FF9800", width: 1 }, rectRadius: 0.06,
  });
  slide.addText("⚠  기준치는 임신을 위한 최소치입니다. 수치가 높을수록 임신 성공률이 높아집니다.", {
    x: 0.5, y: 4.65, w: 9.0, h: 0.52,
    fontSize: 12, fontFace: FONT, bold: true, color: "E65100",
    align: "center", valign: "middle",
  });
}

// ============================================================
// 슬라이드 4: 여성 난임 — AMH 해석 & 난자의 질
// ============================================================
{
  const slide = pptx.addSlide();
  addSlideTitle(slide, "여성 난임 — 난소기능 평가(AMH) & 난자의 질");

  // 좌측 이미지
  try {
    slide.addImage({ path: IMG.ovaryCross, x: 0.4, y: 1.05, w: 3.8, h: 3.7, sizing: { type: "contain", w: 3.8, h: 3.7 } });
  } catch(e) {
    addPlaceholder(slide, 0.4, 1.05, 3.8, 3.7, "난소 단면(Ovary Cross-Section)");
  }
  slide.addText("난소 단면(Ovary Cross-Section) | TeachMeAnatomy", {
    x: 0.4, y: 4.78, w: 3.8, h: 0.25,
    fontSize: 8, fontFace: FONT, color: COLOR.medGray, italic: true, align: "center",
  });

  // 우측 정보
  const items = [
    { label: "난저(Diminished Ovarian Reserve)", text: "AMH ≤ 1.0 ng/mL — 난소기능저하", color: "E65100" },
    { label: "극난저(Severely Diminished)", text: "AMH ≤ 0.4 ng/mL — 심한 난소기능저하", color: COLOR.softRed },
    { label: "AMH vs 난자의 질", text: "AMH는 자라나는 난포 수의 지표 ≠ 난자의 질. 임신에는 AMH보다 난자의 질이 더 중요합니다.", color: COLOR.teal },
    { label: "난자의 질 결정 요인", text: "나이의 영향이 가장 큼. 생활습관 · 몸 컨디션으로 개선 가능.", color: COLOR.navy },
    { label: "AMH 낮아도 임신 가능", text: "우성 난포(Dominant Follicle)가 잘 자란다면 배란 및 임신 가능", color: COLOR.darkGray },
    { label: "PCOS 주의", text: "AMH가 높아도 우성 난포가 없으면 배란 불가 → 임신 불가", color: COLOR.darkGray },
  ];

  items.forEach((item, idx) => {
    const y = 1.05 + idx * 0.62;
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 4.5, y: y, w: 0.08, h: 0.5,
      fill: { color: item.color === COLOR.darkGray ? COLOR.teal : item.color }, line: { type: "none" },
    });
    slide.addText(item.label, {
      x: 4.7, y: y, w: 4.8, h: 0.28,
      fontSize: 12, fontFace: FONT, bold: true, color: item.color === "E65100" ? "E65100" : COLOR.navy, valign: "bottom",
    });
    slide.addText(item.text, {
      x: 4.7, y: y + 0.28, w: 4.8, h: 0.3,
      fontSize: 11, fontFace: FONT, color: COLOR.darkGray, valign: "top",
    });
  });
}

// ============================================================
// 슬라이드 5: 남성 난임 — 정액 검사 해석
// ============================================================
{
  const slide = pptx.addSlide();
  addSlideTitle(slide, "남성 난임 — 정액 검사 해석");

  // 좌측 이미지 (placeholder for sperm)
  addPlaceholder(slide, 0.4, 1.05, 3.8, 3.7, "정자 구조\n(Complete diagram of a\nhuman spermatozoa)\n\nWikimedia Commons,\nLadyofHats, Public Domain");

  // 우측 정보
  const rows = [
    { title: "기준치의 의미", content: "WHO 기준치는 '임신 가능 최소치'입니다.\n수치가 높을수록 임신 성공률이 높아집니다." },
    { title: "무정자증(Azoospermia)이어도", content: "정액 분석 시 정상 정자가 단 1마리라도 있으면\n시험관 시술(IVF) 가능합니다." },
    { title: "정자 농도(Concentration)", content: "1mL당 1,500만 개 이상 필요.\n농도가 높을수록 수정 성공률 상승." },
    { title: "운동성(Motility)", content: "전진 운동성 ≥ 32% / 전체 운동성 ≥ 42%.\n운동성 불량 시 ICSI(세포질 내 정자 주입) 고려." },
    { title: "형태(Morphology)", content: "정상 형태 ≥ 4% (Kruger 기준).\n형태 이상이 심할수록 수정률 저하." },
  ];

  rows.forEach((row, idx) => {
    const y = 1.05 + idx * 0.72;
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 4.5, y: y, w: 5.1, h: 0.65,
      fill: { color: idx % 2 === 0 ? COLOR.lightTeal : COLOR.white },
      line: { color: COLOR.lineGray, width: 0.5 },
    });
    slide.addText(row.title, {
      x: 4.65, y: y + 0.04, w: 4.8, h: 0.28,
      fontSize: 12, fontFace: FONT, bold: true, color: COLOR.teal, valign: "middle",
    });
    slide.addText(row.content, {
      x: 4.65, y: y + 0.30, w: 4.8, h: 0.32,
      fontSize: 10.5, fontFace: FONT, color: COLOR.darkGray, valign: "top",
    });
  });
}

// ============================================================
// 슬라이드 6: 난임 진료 과정 흐름도
// ============================================================
{
  const slide = pptx.addSlide();
  addSlideTitle(slide, "난임 진료 과정 — 단계별 흐름");

  const steps = [
    { num: "01", title: "초기 검사", items: ["AMH 검사 (난소기능)", "정액 검사", "기저 호르몬 검사", "초음파 검사"] },
    { num: "02", title: "난관조영술\n(HSG) + 자궁경", items: ["난관 개통 여부 확인", "자궁경으로 자궁강 관찰", "시술 후 임신율 소폭 상승"] },
    { num: "03", title: "인공수정\n(IUI)", items: ["정자 전처리 후 난관 근처 주입", "과배란 유도 병행", "자연/호르몬 주기 선택"] },
    { num: "04", title: "시험관\n(IVF)", items: ["체외 수정 + 배양", "단기 / 장기 / 저자극 요법", "배아 이식 (신선/동결 주기)"] },
  ];

  const colors = [COLOR.teal, COLOR.navy, COLOR.teal, COLOR.navy];

  steps.forEach((step, idx) => {
    const x = 0.4 + idx * 2.43;
    const y = 1.15;

    // Header box
    slide.addShape(pptx.shapes.RECTANGLE, {
      x, y, w: 2.25, h: 0.65,
      fill: { color: colors[idx] }, line: { type: "none" },
    });
    slide.addText(`STEP ${step.num}`, {
      x, y, w: 2.25, h: 0.28,
      fontSize: 10, fontFace: FONT, bold: true, color: COLOR.white,
      align: "center", valign: "bottom",
    });
    slide.addText(step.title, {
      x, y: y + 0.28, w: 2.25, h: 0.37,
      fontSize: 12.5, fontFace: FONT, bold: true, color: COLOR.white,
      align: "center", valign: "top",
    });

    // Body box
    slide.addShape(pptx.shapes.RECTANGLE, {
      x, y: y + 0.65, w: 2.25, h: 2.6,
      fill: { color: COLOR.lightTeal },
      line: { color: idx % 2 === 0 ? COLOR.teal : COLOR.navy, width: 1 },
    });

    const itemTexts = step.items.map((item) => [
      { text: "▪  ", options: { fontSize: 11, fontFace: FONT, bold: true, color: colors[idx] } },
      { text: item, options: { fontSize: 11, fontFace: FONT, color: COLOR.darkGray, breakLine: true } },
    ]).flat();

    slide.addText(itemTexts, {
      x: x + 0.12, y: y + 0.75, w: 2.05, h: 2.35,
      valign: "top", lineSpacingMultiple: 1.7,
    });

    // Arrow
    if (idx < 3) {
      slide.addText("▶", {
        x: x + 2.25, y: y + 1.1, w: 0.18, h: 0.5,
        fontSize: 14, color: COLOR.teal, align: "center", valign: "middle", fontFace: FONT,
      });
    }
  });

  // 하단 설명 박스
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.4, y: 4.62, w: 9.2, h: 0.58,
    fill: { color: COLOR.navy }, line: { type: "none" },
  });
  slide.addText("인공수정(IUI) 3회 이상 실패 또는 난관 폐쇄, 남성 요인, 고령 등의 경우 시험관(IVF)으로 바로 진행하기도 합니다.", {
    x: 0.5, y: 4.62, w: 9.0, h: 0.58,
    fontSize: 12, fontFace: FONT, color: COLOR.white, align: "center", valign: "middle",
  });
}

// ============================================================
// 슬라이드 7: 난관조영술(HSG) + 자궁경
// ============================================================
{
  const slide = pptx.addSlide();
  addSlideTitle(slide, "난관조영술(Hysterosalpingography, HSG) + 자궁경(Hysteroscopy)");

  try {
    slide.addImage({ path: IMG.uterusFallopian, x: 0.4, y: 1.05, w: 4.3, h: 3.5, sizing: { type: "contain", w: 4.3, h: 3.5 } });
  } catch(e) {
    addPlaceholder(slide, 0.4, 1.05, 4.3, 3.5, "난관 구조\n(Fallopian Tube Parts)\nTeachMeAnatomy");
  }
  slide.addText("출처: TeachMeAnatomy, Fallopian Tubes Anatomy", {
    x: 0.4, y: 4.58, w: 4.3, h: 0.25, fontSize: 8, fontFace: FONT, color: COLOR.medGray, italic: true, align: "center",
  });

  const tableData = [
    [{ text: "항목", options: { fontSize: 12, fontFace: FONT, bold: true, color: COLOR.white, fill: { color: COLOR.teal }, align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "내용", options: { fontSize: 12, fontFace: FONT, bold: true, color: COLOR.white, fill: { color: COLOR.teal }, align: "center", valign: "middle", margin: [4,6,4,6] }}],
    [{ text: "검사 목적", options: { fontSize: 11, fontFace: FONT, bold: true, color: COLOR.navy, fill: { color: COLOR.lightTeal }, valign: "middle", margin: [4,6,4,6] }},
     { text: "난관 개통 여부 확인 + 자궁강 이상 여부 관찰", options: { fontSize: 11, fontFace: FONT, color: COLOR.darkGray, fill: { color: COLOR.lightTeal }, valign: "middle", margin: [4,6,4,6] }}],
    [{ text: "자궁경 효과", options: { fontSize: 11, fontFace: FONT, bold: true, color: COLOR.navy, valign: "middle", margin: [4,6,4,6] }},
     { text: "시술 후 일시적 관 확장 → 임신 확률 소폭 상승", options: { fontSize: 11, fontFace: FONT, color: COLOR.darkGray, valign: "middle", margin: [4,6,4,6] }}],
    [{ text: "검사 한계", options: { fontSize: 11, fontFace: FONT, bold: true, color: COLOR.softRed, fill: { color: COLOR.lightTeal }, valign: "middle", margin: [4,6,4,6] }},
     { text: "난관 내 액체 흐름 · 편모 운동 방향은 확인 불가\n→ 개통 확인 ≠ 기능 정상", options: { fontSize: 11, fontFace: FONT, color: COLOR.darkGray, fill: { color: COLOR.lightTeal }, valign: "middle", margin: [4,6,4,6] }}],
    [{ text: "적용 시점", options: { fontSize: 11, fontFace: FONT, bold: true, color: COLOR.navy, valign: "middle", margin: [4,6,4,6] }},
     { text: "난임 초기 검사 단계에서 시행 (인공수정 전 권장)", options: { fontSize: 11, fontFace: FONT, color: COLOR.darkGray, valign: "middle", margin: [4,6,4,6] }}],
  ];
  slide.addTable(tableData, {
    x: 5.0, y: 1.05, w: 4.6,
    border: { type: "solid", color: COLOR.lineGray, pt: 0.5 },
    colW: [1.6, 3.0], rowH: [0.45, 0.55, 0.55, 0.72, 0.55],
  });
}

// ============================================================
// 슬라이드 8: 인공수정(IUI)
// ============================================================
{
  const slide = pptx.addSlide();
  addSlideTitle(slide, "인공수정(Intrauterine Insemination, IUI)");

  try {
    slide.addImage({ path: IMG.femaleOverview, x: 0.4, y: 1.05, w: 4.3, h: 3.5, sizing: { type: "contain", w: 4.3, h: 3.5 } });
  } catch(e) {
    addPlaceholder(slide, 0.4, 1.05, 4.3, 3.5, "자궁 · 난관 · 난소 해부도\nTeachMeAnatomy");
  }
  slide.addText("출처: TeachMeAnatomy, Female Reproductive Overview", {
    x: 0.4, y: 4.58, w: 4.3, h: 0.25, fontSize: 8, fontFace: FONT, color: COLOR.medGray, italic: true, align: "center",
  });

  const rows = [
    { title: "정의", content: "정자에 화학적 전처리(세척·농축) 후 난관 근처(자궁강) 직접 주입" },
    { title: "수정 여부 확인 불가", content: "정자 주입 후 수정 성공 여부 · 난관 이동 · 착상 여부는 확인 불가능" },
    { title: "과배란 병행", content: "자연주기 또는 과배란 유도(보통 2~3개 난자) 후 시행" },
    { title: "주기 선택", content: "자연주기 vs 호르몬 제제를 이용한 조절 주기 중 선택" },
    { title: "적응증", content: "경증 남성인자 / 경부 인자 / 원인불명 난임 / 3회 이상 반복 시 IVF 고려" },
  ];

  rows.forEach((row, idx) => {
    const y = 1.05 + idx * 0.72;
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 5.0, y, w: 4.6, h: 0.64,
      fill: { color: idx % 2 === 0 ? COLOR.lightTeal : COLOR.white },
      line: { color: COLOR.lineGray, width: 0.5 },
    });
    slide.addText(row.title, {
      x: 5.12, y: y + 0.03, w: 4.35, h: 0.26,
      fontSize: 12, fontFace: FONT, bold: true, color: COLOR.teal, valign: "middle",
    });
    slide.addText(row.content, {
      x: 5.12, y: y + 0.30, w: 4.35, h: 0.30,
      fontSize: 10.5, fontFace: FONT, color: COLOR.darkGray, valign: "top",
    });
  });
}

// ============================================================
// 슬라이드 9: 시험관(IVF) 과배란 유도 — 개요
// ============================================================
{
  const slide = pptx.addSlide();
  addSlideTitle(slide, "시험관(IVF) — 과배란 유도 3가지 요법");

  try {
    slide.addImage({ path: IMG.ivfProcess, x: 0.4, y: 1.05, w: 4.2, h: 3.65, sizing: { type: "contain", w: 4.2, h: 3.65 } });
  } catch(e) {
    addPlaceholder(slide, 0.4, 1.05, 4.2, 3.65, "IVF 전체 과정 인포그래픽\nWikimedia Commons, CC BY-SA 4.0");
  }
  slide.addText("출처: Wikimedia Commons — In Vitro Fertilization (IVF) - English.png, CC BY-SA 4.0", {
    x: 0.4, y: 4.73, w: 4.2, h: 0.25, fontSize: 8, fontFace: FONT, color: COLOR.medGray, italic: true, align: "center",
  });

  const methods = [
    {
      title: "단기요법(Short Protocol)",
      color: COLOR.teal,
      items: [
        "생리 시작 3일차부터 FSH 또는 hMG 주사/복약",
        "조기배란 방지 필요시 GnRH 길항제(antagonist) 추가",
        "FSH 단일 제제 또는 FSH+LH 혼합 제제 사용",
      ],
    },
    {
      title: "장기요법(Long Protocol)",
      color: COLOR.navy,
      items: [
        "단기요법 반응 불량 시 선택",
        "생리 시작 7일 전부터 GnRH 작용제 투여 → FSH/LH 억제",
        "생리 2~3일차부터 과배란 유도 시작 → 난포 균일 성장",
        "변형법: 생리 시작 후 1~2주 작용제 → 단기요법 방식으로 전환",
      ],
    },
    {
      title: "저자극 요법(Minimal Stimulation)",
      color: "607D8B",
      items: [
        "난소기능 저하(Diminished Ovarian Reserve) 환자에 적용",
        "적은 용량 자극으로 소수의 난자 채취",
        "신체 부담 최소화, 난소과자극증후군 위험 감소",
      ],
    },
  ];

  methods.forEach((method, idx) => {
    const y = 1.05 + idx * 1.2;
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 4.8, y, w: 4.8, h: 0.38,
      fill: { color: method.color }, line: { type: "none" },
    });
    slide.addText(method.title, {
      x: 4.9, y, w: 4.6, h: 0.38,
      fontSize: 13, fontFace: FONT, bold: true, color: COLOR.white,
      valign: "middle",
    });
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 4.8, y: y + 0.38, w: 4.8, h: 0.76,
      fill: { color: COLOR.lightTeal }, line: { color: method.color, width: 0.8 },
    });
    const bullets = method.items.map((item) => [
      { text: "•  ", options: { fontSize: 10.5, fontFace: FONT, bold: true, color: method.color } },
      { text: item, options: { fontSize: 10.5, fontFace: FONT, color: COLOR.darkGray, breakLine: true } },
    ]).flat();
    slide.addText(bullets, {
      x: 4.95, y: y + 0.42, w: 4.55, h: 0.68,
      valign: "top", lineSpacingMultiple: 1.5,
    });
  });
}

// ============================================================
// 슬라이드 10: 과배란 유도 약제 상세 비교
// ============================================================
{
  const slide = pptx.addSlide();
  addSlideTitle(slide, "과배란 유도 약제 상세 — FSH 단일 vs hMG(FSH+LH 혼합)");

  // FSH 단일 요법
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.4, y: 1.05, w: 9.2, h: 0.38,
    fill: { color: COLOR.teal }, line: { type: "none" },
  });
  slide.addText("FSH 단일 요법", {
    x: 0.5, y: 1.05, w: 9.0, h: 0.38,
    fontSize: 14, fontFace: FONT, bold: true, color: COLOR.white, valign: "middle",
  });

  const fshRows = [
    [{ text: "성분명", options: { fontSize: 11.5, fontFace: FONT, bold: true, color: COLOR.white, fill: { color: COLOR.navy }, align: "center", valign: "middle", margin: [3,4,3,4] }},
     { text: "Follitropin-α", options: { fontSize: 11.5, fontFace: FONT, bold: true, color: COLOR.navy, fill: { color: COLOR.lightTeal }, align: "center", valign: "middle", margin: [3,4,3,4] }},
     { text: "Follitropin-β", options: { fontSize: 11.5, fontFace: FONT, bold: true, color: COLOR.navy, fill: { color: COLOR.lightTeal }, align: "center", valign: "middle", margin: [3,4,3,4] }},
     { text: "Follitropin (r-hFSH)", options: { fontSize: 11.5, fontFace: FONT, bold: true, color: COLOR.navy, fill: { color: COLOR.lightTeal }, align: "center", valign: "middle", margin: [3,4,3,4] }},
     { text: "Follitropin-δ", options: { fontSize: 11.5, fontFace: FONT, bold: true, color: COLOR.navy, fill: { color: COLOR.lightTeal }, align: "center", valign: "middle", margin: [3,4,3,4] }}],
    [{ text: "제품명", options: { fontSize: 11.5, fontFace: FONT, bold: true, color: COLOR.white, fill: { color: COLOR.navy }, align: "center", valign: "middle", margin: [3,4,3,4] }},
     { text: "고날에프", options: { fontSize: 13, fontFace: FONT, bold: true, color: COLOR.teal, align: "center", valign: "middle", margin: [3,4,3,4] }},
     { text: "퓨레곤", options: { fontSize: 13, fontFace: FONT, bold: true, color: COLOR.teal, align: "center", valign: "middle", margin: [3,4,3,4] }},
     { text: "폴리트롭\n고나도핀", options: { fontSize: 12, fontFace: FONT, bold: true, color: COLOR.teal, align: "center", valign: "middle", margin: [3,4,3,4] }},
     { text: "레코벨", options: { fontSize: 13, fontFace: FONT, bold: true, color: COLOR.teal, align: "center", valign: "middle", margin: [3,4,3,4] }}],
  ];
  slide.addTable(fshRows, {
    x: 0.4, y: 1.43, w: 9.2,
    border: { type: "solid", color: COLOR.lineGray, pt: 0.5 },
    colW: [1.3, 1.98, 1.98, 1.98, 1.98], rowH: [0.42, 0.52],
  });

  // hMG 혼합 요법
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.4, y: 2.65, w: 9.2, h: 0.38,
    fill: { color: COLOR.navy }, line: { type: "none" },
  });
  slide.addText("hMG(FSH+LH 혼합) 요법", {
    x: 0.5, y: 2.65, w: 9.0, h: 0.38,
    fontSize: 14, fontFace: FONT, bold: true, color: COLOR.white, valign: "middle",
  });

  const hmgRows = [
    [{ text: "성분명", options: { fontSize: 11.5, fontFace: FONT, bold: true, color: COLOR.white, fill: { color: COLOR.teal }, align: "center", valign: "middle", margin: [3,4,3,4] }},
     { text: "Follitropin-α + r-hLH (2:1)", options: { fontSize: 11.5, fontFace: FONT, bold: true, color: COLOR.navy, fill: { color: COLOR.lightTeal }, align: "center", valign: "middle", margin: [3,4,3,4] }},
     { text: "FSH+LH (2:1) Menotrophin", options: { fontSize: 11.5, fontFace: FONT, bold: true, color: COLOR.navy, fill: { color: COLOR.lightTeal }, align: "center", valign: "middle", margin: [3,4,3,4] }}],
    [{ text: "제품명", options: { fontSize: 11.5, fontFace: FONT, bold: true, color: COLOR.white, fill: { color: COLOR.teal }, align: "center", valign: "middle", margin: [3,4,3,4] }},
     { text: "퍼고베리스", options: { fontSize: 14, fontFace: FONT, bold: true, color: COLOR.navy, align: "center", valign: "middle", margin: [3,4,3,4] }},
     { text: "메노푸어 / IVF-M / 고나도핀", options: { fontSize: 13, fontFace: FONT, bold: true, color: COLOR.navy, align: "center", valign: "middle", margin: [3,4,3,4] }}],
    [{ text: "특징", options: { fontSize: 11.5, fontFace: FONT, bold: true, color: COLOR.white, fill: { color: COLOR.teal }, align: "center", valign: "middle", margin: [3,4,3,4] }},
     { text: "FSH+LH 동시 보충 → 황체화 지원\n특히 LH 불충분한 환자에게 유리", options: { fontSize: 11, fontFace: FONT, color: COLOR.darkGray, valign: "middle", margin: [4,6,4,6] }},
     { text: "소변 유래 hMG / 재조합 FSH+LH 혼합\n다양한 농도 옵션 존재", options: { fontSize: 11, fontFace: FONT, color: COLOR.darkGray, valign: "middle", margin: [4,6,4,6] }}],
  ];
  slide.addTable(hmgRows, {
    x: 0.4, y: 3.03, w: 9.2,
    border: { type: "solid", color: COLOR.lineGray, pt: 0.5 },
    colW: [1.3, 3.95, 3.95], rowH: [0.42, 0.52, 0.65],
  });

  slide.addText("생리 시작 3일차부터 복약/주사 시작 | 각 제제 내에서도 작용 유닛이 다름", {
    x: 0.4, y: 4.82, w: 9.2, h: 0.3,
    fontSize: 10, fontFace: FONT, color: COLOR.medGray, align: "center", italic: true,
  });
}

// ============================================================
// 슬라이드 11: GnRH 길항제 vs 작용제 비교
// ============================================================
{
  const slide = pptx.addSlide();
  addSlideTitle(slide, "GnRH 길항제(Antagonist) vs 작용제(Agonist) 비교");

  const cols = [
    {
      title: "GnRH 길항제(Antagonist)",
      color: COLOR.teal,
      items: [
        { label: "작용 원리", text: "수용체에 GnRH 대신 결합 → GnRH 작용 즉시 차단" },
        { label: "결과", text: "FSH · LH 즉시 분비 억제 → 배란 억제" },
        { label: "사용 시점", text: "단기요법 중 조기배란 방지 필요 시" },
        { label: "장점", text: "즉각 효과 / 투여 기간 짧음" },
        { label: "약제", text: "오가루트란 · 가니레버\n세트로타이드 · 유레릭스" },
      ],
    },
    {
      title: "GnRH 작용제(Agonist)",
      color: COLOR.navy,
      items: [
        { label: "작용 원리", text: "GnRH 분비 지속 신호 → 초기 FSH/LH 일시 상승 → 이후 억제(Down-regulation)" },
        { label: "왜 억제되나?", text: "GnRH는 박동성 분비해야 FSH/LH 자극 가능\n→ 지속 투여 시 수용체 둔감화 → FSH/LH 억제" },
        { label: "사용 시점", text: "장기요법: 생리 7일 전부터 선행 투여" },
        { label: "장점", text: "난포 균일 성장 / 조기 황체화 방지" },
        { label: "약제", text: "로렐린 · 데카펩틸\n루프론 · 루프린 · 슈퍼팍트" },
      ],
    },
  ];

  cols.forEach((col, idx) => {
    const x = 0.4 + idx * 4.8;
    slide.addShape(pptx.shapes.RECTANGLE, {
      x, y: 1.05, w: 4.5, h: 0.5,
      fill: { color: col.color }, line: { type: "none" },
    });
    slide.addText(col.title, {
      x, y: 1.05, w: 4.5, h: 0.5,
      fontSize: 14, fontFace: FONT, bold: true, color: COLOR.white,
      align: "center", valign: "middle",
    });

    col.items.forEach((item, iIdx) => {
      const iy = 1.62 + iIdx * 0.68;
      slide.addShape(pptx.shapes.RECTANGLE, {
        x, y: iy, w: 4.5, h: 0.62,
        fill: { color: iIdx % 2 === 0 ? COLOR.lightTeal : COLOR.white },
        line: { color: col.color, width: 0.5 },
      });
      slide.addText(item.label, {
        x: x + 0.12, y: iy + 0.03, w: 4.25, h: 0.25,
        fontSize: 11, fontFace: FONT, bold: true, color: col.color, valign: "middle",
      });
      slide.addText(item.text, {
        x: x + 0.12, y: iy + 0.28, w: 4.25, h: 0.3,
        fontSize: 10, fontFace: FONT, color: COLOR.darkGray, valign: "top",
      });
    });
  });
}

// ============================================================
// 슬라이드 12: hCG 주사 & 자궁내막 두께
// ============================================================
{
  const slide = pptx.addSlide();
  addSlideTitle(slide, "hCG 주사(배란 유도) & 자궁내막 두께와 임신성공률");

  // hCG 설명
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.4, y: 1.05, w: 9.2, h: 0.38,
    fill: { color: COLOR.teal }, line: { type: "none" },
  });
  slide.addText("hCG 주사제 — 난포 최종 성숙 & 배란 유도", {
    x: 0.5, y: 1.05, w: 9.0, h: 0.38,
    fontSize: 13, fontFace: FONT, bold: true, color: COLOR.white, valign: "middle",
  });

  const hcgRows = [
    [{ text: "특징", options: { fontSize: 11.5, fontFace: FONT, bold: true, color: COLOR.white, fill: { color: COLOR.navy }, align: "center", valign: "middle", margin: [3,5,3,5] }},
     { text: "오비드렐", options: { fontSize: 12, fontFace: FONT, bold: true, color: COLOR.teal, align: "center", valign: "middle", margin: [3,5,3,5] }},
     { text: "IVF-C", options: { fontSize: 12, fontFace: FONT, bold: true, color: COLOR.teal, align: "center", valign: "middle", margin: [3,5,3,5] }},
     { text: "프레그닐", options: { fontSize: 12, fontFace: FONT, bold: true, color: COLOR.teal, align: "center", valign: "middle", margin: [3,5,3,5] }},
     { text: "코리오몬", options: { fontSize: 12, fontFace: FONT, bold: true, color: COLOR.teal, align: "center", valign: "middle", margin: [3,5,3,5] }}],
    [{ text: "36시간 후 배란\nhCG → 임테기 양성 가능(10일 영향)", options: { fontSize: 10.5, fontFace: FONT, color: COLOR.darkGray, fill: { color: COLOR.lightTeal }, align: "center", valign: "middle", margin: [3,5,3,5] }},
     { text: "재조합 hCG\n(r-hCG)", options: { fontSize: 11, fontFace: FONT, color: COLOR.darkGray, fill: { color: COLOR.lightTeal }, align: "center", valign: "middle", margin: [3,5,3,5] }},
     { text: "재조합 hCG\n복부주사", options: { fontSize: 11, fontFace: FONT, color: COLOR.darkGray, fill: { color: COLOR.lightTeal }, align: "center", valign: "middle", margin: [3,5,3,5] }},
     { text: "소변 유래\nhCG", options: { fontSize: 11, fontFace: FONT, color: COLOR.darkGray, fill: { color: COLOR.lightTeal }, align: "center", valign: "middle", margin: [3,5,3,5] }},
     { text: "소변 유래\nhCG (고농도)", options: { fontSize: 11, fontFace: FONT, color: COLOR.darkGray, fill: { color: COLOR.lightTeal }, align: "center", valign: "middle", margin: [3,5,3,5] }}],
  ];
  slide.addTable(hcgRows, {
    x: 0.4, y: 1.43, w: 9.2,
    border: { type: "solid", color: COLOR.lineGray, pt: 0.5 },
    colW: [2.4, 1.7, 1.7, 1.7, 1.6], rowH: [0.45, 0.6],
  });

  // 자궁내막 두께 표
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.4, y: 2.72, w: 9.2, h: 0.38,
    fill: { color: COLOR.navy }, line: { type: "none" },
  });
  slide.addText("배란기 자궁내막(Endometrium) 두께와 임신성공률", {
    x: 0.5, y: 2.72, w: 9.0, h: 0.38,
    fontSize: 13, fontFace: FONT, bold: true, color: COLOR.white, valign: "middle",
  });

  const endoRows = [
    [{ text: "자궁내막 두께", options: { fontSize: 12, fontFace: FONT, bold: true, color: COLOR.white, fill: { color: COLOR.teal }, align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "< 6mm", options: { fontSize: 12, fontFace: FONT, bold: true, color: COLOR.softRed, align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "6~8mm", options: { fontSize: 12, fontFace: FONT, bold: true, color: "E65100", align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "8~10mm", options: { fontSize: 12, fontFace: FONT, bold: true, color: "2E7D32", align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "10~12mm", options: { fontSize: 12, fontFace: FONT, bold: true, color: "1565C0", align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "> 12mm", options: { fontSize: 12, fontFace: FONT, bold: true, color: COLOR.medGray, align: "center", valign: "middle", margin: [4,6,4,6] }}],
    [{ text: "임신성공률", options: { fontSize: 12, fontFace: FONT, bold: true, color: COLOR.white, fill: { color: COLOR.teal }, align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "3.7%", options: { fontSize: 15, fontFace: FONT, bold: true, color: COLOR.softRed, align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "14.8%", options: { fontSize: 15, fontFace: FONT, bold: true, color: "E65100", align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "55.5% ★", options: { fontSize: 15, fontFace: FONT, bold: true, color: "2E7D32", align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "35.3%", options: { fontSize: 15, fontFace: FONT, bold: true, color: "1565C0", align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "3.7%", options: { fontSize: 15, fontFace: FONT, bold: true, color: COLOR.medGray, align: "center", valign: "middle", margin: [4,6,4,6] }}],
  ];
  slide.addTable(endoRows, {
    x: 0.4, y: 3.10, w: 9.2,
    border: { type: "solid", color: COLOR.lineGray, pt: 0.5 },
    colW: [1.7, 1.5, 1.5, 1.5, 1.5, 1.5], rowH: [0.45, 0.6],
  });

  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.4, y: 4.35, w: 9.2, h: 0.5,
    fill: { color: COLOR.lightTeal }, line: { color: COLOR.teal, width: 0.8 },
  });
  slide.addText("★  최적 내막 두께: 8~10mm (임신성공률 55.5%)  |  내막이 너무 얇으면 비아그라 질정으로 내막 두께 증가 유도 시도", {
    x: 0.5, y: 4.35, w: 9.0, h: 0.5,
    fontSize: 11, fontFace: FONT, bold: true, color: COLOR.teal, align: "center", valign: "middle",
  });
}

// ============================================================
// 슬라이드 13: 배아 배양 전략
// ============================================================
{
  const slide = pptx.addSlide();
  addSlideTitle(slide, "배아 배양(Embryo Culture) 전략 — 3일 vs 5일 배양");

  try {
    slide.addImage({ path: IMG.embryoGrades, x: 0.4, y: 1.05, w: 4.2, h: 3.6, sizing: { type: "contain", w: 4.2, h: 3.6 } });
  } catch(e) {
    addPlaceholder(slide, 0.4, 1.05, 4.2, 3.6, "배아 등급 다이어그램\nWikimedia Commons, CC BY-SA 4.0");
  }
  slide.addText("출처: Wikimedia Commons — Human embryo grades for IVF, CC BY-SA 4.0", {
    x: 0.4, y: 4.68, w: 4.2, h: 0.25, fontSize: 8, fontFace: FONT, color: COLOR.medGray, italic: true, align: "center",
  });

  const rows = [
    { title: "3일 배양", color: COLOR.teal, text: "난자 채취 수가 적을 때 선택. 추가 배양 리스크 없이 이식.\n하급 배아도 임신 가능 → 채취 수 적으면 무리해서 키우지 않음." },
    { title: "5일 배양", color: COLOR.navy, text: "난자 채취 수가 많을 때 최대한 키워봄.\n포배기(Blastocyst) 단계까지 배양 → 등급 상·중·하 분류.\n5일 배양 과정 중 분열 멈추는 경우 많음." },
    { title: "배양 환경", color: "607D8B", text: "엄마 몸속(진동·온도·자궁 환경)이 수정란에 최적.\n배양 기술이 발전했지만 체내보다 못함.\n→ 모양 안 좋으면 조기(2~3일차) 이식이 유리." },
    { title: "수정 vs 분화", color: COLOR.teal, text: "수정 단계: 난자 + 정자의 상태가 중요\n분화(배양) 단계: 난자의 상태가 핵심" },
    { title: "수정률 기준", color: COLOR.navy, text: "5일 배양 (상>중>하)  >  3일 배양 (상>중>하)\n모양·분열 속도로 등급 결정" },
  ];

  rows.forEach((row, idx) => {
    const y = 1.05 + idx * 0.72;
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 4.8, y, w: 4.8, h: 0.64,
      fill: { color: idx % 2 === 0 ? COLOR.lightTeal : COLOR.white },
      line: { color: row.color, width: 0.8 },
    });
    slide.addText(row.title, {
      x: 4.92, y: y + 0.03, w: 4.55, h: 0.25,
      fontSize: 12, fontFace: FONT, bold: true, color: row.color, valign: "middle",
    });
    slide.addText(row.text, {
      x: 4.92, y: y + 0.27, w: 4.55, h: 0.34,
      fontSize: 10, fontFace: FONT, color: COLOR.darkGray, valign: "top",
    });
  });
}

// ============================================================
// 슬라이드 14: 배아 단계 설명 (포배기·눈사람·감자)
// ============================================================
{
  const slide = pptx.addSlide();
  addSlideTitle(slide, "5일 배양 배아 단계 — 포배기 · 눈사람 · 감자");

  try {
    slide.addImage({ path: IMG.blastocyst, x: 0.4, y: 1.05, w: 4.0, h: 3.6, sizing: { type: "contain", w: 4.0, h: 3.6 } });
  } catch(e) {
    addPlaceholder(slide, 0.4, 1.05, 4.0, 3.6, "포배기 배아(Blastocyst)\nWikimedia Commons");
  }
  slide.addText("출처: Wikimedia Commons — Blastocyst embryo.png (DBCLS)", {
    x: 0.4, y: 4.68, w: 4.0, h: 0.25, fontSize: 8, fontFace: FONT, color: COLOR.medGray, italic: true, align: "center",
  });

  const stages = [
    {
      name: "포배기(Blastocyst)",
      color: COLOR.teal,
      desc: "얇은 막(투명대, Zona Pellucida)으로 쌓인 배아.\n내부에 내세포괴(ICM)와 영양배엽(Trophectoderm) 분화 시작.",
    },
    {
      name: "눈사람(Hatching Blastocyst)",
      color: COLOR.navy,
      desc: "분열된 세포가 얇아진 막을 뚫고 나오는 과정 = '부화(Hatching)'.\n막을 빠져나오는 모습이 눈사람처럼 보임.",
    },
    {
      name: "감자(Hatched Blastocyst)",
      color: "607D8B",
      desc: "완전히 부화가 끝난 배아.\n투명대(껍질) 없이 동그란 모양 → 감자처럼 보임.\n이 상태에서 자궁내막에 착상(Implantation) 시도.",
    },
  ];

  stages.forEach((stage, idx) => {
    const y = 1.05 + idx * 1.22;
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 4.6, y, w: 5.0, h: 0.4,
      fill: { color: stage.color }, line: { type: "none" },
    });
    slide.addText(stage.name, {
      x: 4.7, y, w: 4.8, h: 0.4,
      fontSize: 14, fontFace: FONT, bold: true, color: COLOR.white, valign: "middle",
    });
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 4.6, y: y + 0.4, w: 5.0, h: 0.75,
      fill: { color: COLOR.lightTeal }, line: { color: stage.color, width: 0.8 },
    });
    slide.addText(stage.desc, {
      x: 4.72, y: y + 0.44, w: 4.75, h: 0.67,
      fontSize: 11, fontFace: FONT, color: COLOR.darkGray, valign: "top", lineSpacingMultiple: 1.4,
    });
  });
}

// ============================================================
// 슬라이드 15: PGT 검사 & E2/P4 이식 결정
// ============================================================
{
  const slide = pptx.addSlide();
  addSlideTitle(slide, "PGT 검사 & 이식 결정 기준 — E2 · P4 수치");

  // PGT 설명
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.4, y: 1.05, w: 9.2, h: 0.38,
    fill: { color: COLOR.teal }, line: { type: "none" },
  });
  slide.addText("착상 전 유전자 검사 (PGT, Preimplantation Genetic Testing)", {
    x: 0.5, y: 1.05, w: 9.0, h: 0.38,
    fontSize: 13, fontFace: FONT, bold: true, color: COLOR.white, valign: "middle",
  });

  const pgtItems = [
    "습관성 유산(Recurrent Pregnancy Loss) 환자를 위해 시행",
    "5일 배양 배아의 일부 세포 채취 → 유전자 분석 → 정상 배아만 이식 선별",
    "검사 정확도: 100%가 아님 (일부 세포만 채취하는 특성상)",
    "비용: 1회당 약 100만 원 수준 (고가)",
    "병원별 차이: 모든 환자에게 시행(차병원 등) vs 선택적 시행",
    "PGT 불통과 배아 → 폐기 처리",
    "유산 종결 시 MTX(항암주사) 사용 → 2~3개월 이식 불가",
  ];
  const pgtTexts = pgtItems.map(item => [
    { text: "•  ", options: { fontSize: 11.5, fontFace: FONT, bold: true, color: COLOR.teal } },
    { text: item, options: { fontSize: 11.5, fontFace: FONT, color: COLOR.darkGray, breakLine: true } },
  ]).flat();
  slide.addText(pgtTexts, {
    x: 0.5, y: 1.48, w: 9.0, h: 1.55,
    valign: "top", lineSpacingMultiple: 1.55,
  });

  // E2/P4 표
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.4, y: 3.15, w: 9.2, h: 0.38,
    fill: { color: COLOR.navy }, line: { type: "none" },
  });
  slide.addText("이식 결정 기준 — E2(에스트라디올) & P4(프로게스테론) 수치", {
    x: 0.5, y: 3.15, w: 9.0, h: 0.38,
    fontSize: 13, fontFace: FONT, bold: true, color: COLOR.white, valign: "middle",
  });

  const e2p4Rows = [
    [{ text: "시점", options: { fontSize: 12, fontFace: FONT, bold: true, color: COLOR.white, fill: { color: COLOR.teal }, align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "E2(에스트라디올)", options: { fontSize: 12, fontFace: FONT, bold: true, color: COLOR.white, fill: { color: COLOR.teal }, align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "P4(프로게스테론)", options: { fontSize: 12, fontFace: FONT, bold: true, color: COLOR.white, fill: { color: COLOR.teal }, align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "임상적 의미", options: { fontSize: 12, fontFace: FONT, bold: true, color: COLOR.white, fill: { color: COLOR.teal }, align: "center", valign: "middle", margin: [4,6,4,6] }}],
    [{ text: "채취 2~3일 전", options: { fontSize: 12, fontFace: FONT, bold: true, color: COLOR.navy, fill: { color: COLOR.lightTeal }, align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "< 3,000 pg/mL", options: { fontSize: 12, fontFace: FONT, color: COLOR.darkGray, fill: { color: COLOR.lightTeal }, align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "< 1.5 ng/mL", options: { fontSize: 12, fontFace: FONT, color: COLOR.darkGray, fill: { color: COLOR.lightTeal }, align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "기준 초과 시 동결 주기로 전환", options: { fontSize: 11, fontFace: FONT, color: COLOR.darkGray, fill: { color: COLOR.lightTeal }, align: "center", valign: "middle", margin: [4,6,4,6] }}],
    [{ text: "이식 날", options: { fontSize: 12, fontFace: FONT, bold: true, color: COLOR.navy, align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "임신율에 영향 적음\n(3,000 이상도 이식 가능)", options: { fontSize: 11, fontFace: FONT, color: COLOR.darkGray, align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "> 20~25 ng/mL 권장\n(병원마다 20 이상 기준 상이)", options: { fontSize: 11, fontFace: FONT, color: COLOR.darkGray, align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "P4 수치가 충분해야\n착상 환경 최적화", options: { fontSize: 11, fontFace: FONT, color: COLOR.darkGray, align: "center", valign: "middle", margin: [4,6,4,6] }}],
  ];
  slide.addTable(e2p4Rows, {
    x: 0.4, y: 3.53, w: 9.2,
    border: { type: "solid", color: COLOR.lineGray, pt: 0.5 },
    colW: [1.8, 2.4, 2.4, 2.6], rowH: [0.42, 0.5, 0.65],
  });
}

// ============================================================
// 슬라이드 16: 이식 & 프로게스테론 보충
// ============================================================
{
  const slide = pptx.addSlide();
  addSlideTitle(slide, "이식(Embryo Transfer) & 프로게스테론 보충 & NK 수치");

  // 프로게스테론 보충
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.4, y: 1.05, w: 9.2, h: 0.38,
    fill: { color: COLOR.teal }, line: { type: "none" },
  });
  slide.addText("프로게스테론(Progesterone) 보충 — 이유 & 제제", {
    x: 0.5, y: 1.05, w: 9.0, h: 0.38,
    fontSize: 13, fontFace: FONT, bold: true, color: COLOR.white, valign: "middle",
  });

  slide.addText("과배란 유도 시 자연 배란보다 LH 분비 감소 → 황체기 단축 → 프로게스테론 불충분 → 보충 필요", {
    x: 0.5, y: 1.48, w: 9.0, h: 0.38,
    fontSize: 11.5, fontFace: FONT, color: COLOR.darkGray,
  });

  const progRows = [
    [{ text: "투여 경로", options: { fontSize: 11.5, fontFace: FONT, bold: true, color: COLOR.white, fill: { color: COLOR.navy }, align: "center", valign: "middle", margin: [3,5,3,5] }},
     { text: "질좌제 / 질정", options: { fontSize: 12, fontFace: FONT, bold: true, color: COLOR.teal, fill: { color: COLOR.lightTeal }, align: "center", valign: "middle", margin: [3,5,3,5] }},
     { text: "경구약", options: { fontSize: 12, fontFace: FONT, bold: true, color: COLOR.navy, align: "center", valign: "middle", margin: [3,5,3,5] }},
     { text: "피하주사", options: { fontSize: 12, fontFace: FONT, bold: true, color: COLOR.teal, fill: { color: COLOR.lightTeal }, align: "center", valign: "middle", margin: [3,5,3,5] }},
     { text: "근육주사", options: { fontSize: 12, fontFace: FONT, bold: true, color: COLOR.navy, align: "center", valign: "middle", margin: [3,5,3,5] }}],
    [{ text: "제품명", options: { fontSize: 11.5, fontFace: FONT, bold: true, color: COLOR.white, fill: { color: COLOR.navy }, align: "center", valign: "middle", margin: [3,5,3,5] }},
     { text: "크리논\n유트로게스탄\n예나트론\n사이클로제스트\n루티너스", options: { fontSize: 11, fontFace: FONT, color: COLOR.darkGray, fill: { color: COLOR.lightTeal }, align: "center", valign: "middle", margin: [3,5,3,5] }},
     { text: "유트로게스탄\n듀파스톤", options: { fontSize: 11, fontFace: FONT, color: COLOR.darkGray, align: "center", valign: "middle", margin: [3,5,3,5] }},
     { text: "프롤루텍스", options: { fontSize: 11, fontFace: FONT, color: COLOR.darkGray, fill: { color: COLOR.lightTeal }, align: "center", valign: "middle", margin: [3,5,3,5] }},
     { text: "타이유", options: { fontSize: 11, fontFace: FONT, color: COLOR.darkGray, align: "center", valign: "middle", margin: [3,5,3,5] }}],
  ];
  slide.addTable(progRows, {
    x: 0.4, y: 1.88, w: 9.2,
    border: { type: "solid", color: COLOR.lineGray, pt: 0.5 },
    colW: [1.4, 2.7, 1.7, 1.7, 1.7], rowH: [0.42, 0.82],
  });

  // NK 수치 + 이식 개수
  slide.addText("NK 수치(Natural Killer Cell) > 12% 시: 면역글로불린(아이비글로불린·리브감마) 또는 인트라리피드(콩주사) 투여", {
    x: 0.4, y: 3.32, w: 9.2, h: 0.42,
    fontSize: 11.5, fontFace: FONT, color: COLOR.navy,
  });

  // 이식 개수 표
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.4, y: 3.75, w: 9.2, h: 0.35,
    fill: { color: COLOR.navy }, line: { type: "none" },
  });
  slide.addText("이식 배아 개수 기준", {
    x: 0.5, y: 3.75, w: 9.0, h: 0.35,
    fontSize: 13, fontFace: FONT, bold: true, color: COLOR.white, valign: "middle",
  });

  const transferRows = [
    [{ text: "연령", options: { fontSize: 12, fontFace: FONT, bold: true, color: COLOR.white, fill: { color: COLOR.teal }, align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "2~4일 배양", options: { fontSize: 12, fontFace: FONT, bold: true, color: COLOR.white, fill: { color: COLOR.teal }, align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "5~6일 배양", options: { fontSize: 12, fontFace: FONT, bold: true, color: COLOR.white, fill: { color: COLOR.teal }, align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "참고", options: { fontSize: 12, fontFace: FONT, bold: true, color: COLOR.white, fill: { color: COLOR.teal }, align: "center", valign: "middle", margin: [4,6,4,6] }}],
    [{ text: "35세 미만", options: { fontSize: 13, fontFace: FONT, bold: true, color: COLOR.navy, fill: { color: COLOR.lightTeal }, align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "2개", options: { fontSize: 16, fontFace: FONT, bold: true, color: COLOR.teal, fill: { color: COLOR.lightTeal }, align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "3개", options: { fontSize: 16, fontFace: FONT, bold: true, color: COLOR.teal, fill: { color: COLOR.lightTeal }, align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "기준대로 안 하는 경우 많음", options: { fontSize: 11, fontFace: FONT, color: COLOR.medGray, fill: { color: COLOR.lightTeal }, align: "center", valign: "middle", margin: [4,6,4,6] }}],
    [{ text: "35세 이상", options: { fontSize: 13, fontFace: FONT, bold: true, color: COLOR.navy, align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "1개", options: { fontSize: 16, fontFace: FONT, bold: true, color: COLOR.navy, align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "2개", options: { fontSize: 16, fontFace: FONT, bold: true, color: COLOR.navy, align: "center", valign: "middle", margin: [4,6,4,6] }},
     { text: "임상적 판단에 따라 조정", options: { fontSize: 11, fontFace: FONT, color: COLOR.medGray, align: "center", valign: "middle", margin: [4,6,4,6] }}],
  ];
  slide.addTable(transferRows, {
    x: 0.4, y: 4.1, w: 9.2,
    border: { type: "solid", color: COLOR.lineGray, pt: 0.5 },
    colW: [1.8, 2.2, 2.2, 3.0], rowH: [0.4, 0.52, 0.52],
  });
}

// ============================================================
// 슬라이드 17: 한의 진료 접근
// ============================================================
{
  const slide = pptx.addSlide();
  addSlideTitle(slide, "한의 진료 접근 — 뜸 · 좌훈 · 한약");

  try {
    slide.addImage({ path: IMG.femaleTract, x: 0.4, y: 1.05, w: 3.6, h: 3.6, sizing: { type: "contain", w: 3.6, h: 3.6 } });
  } catch(e) {
    addPlaceholder(slide, 0.4, 1.05, 3.6, 3.6, "여성 생식기 해부도\nTeachMeAnatomy");
  }
  slide.addText("출처: TeachMeAnatomy, Female Reproductive Tract Overview", {
    x: 0.4, y: 4.68, w: 3.6, h: 0.25, fontSize: 8, fontFace: FONT, color: COLOR.medGray, italic: true, align: "center",
  });

  const sections = [
    {
      title: "뜸(Moxibustion) + 좌훈(Smoke Fumigation)",
      color: COLOR.teal,
      items: [
        "배란 전 하복부 온열 자극 → 난자 성숙 · 자궁 환경 개선",
        "배란 이후 온열 자극은 배아에 부정적 영향 → 배란 전까지만",
        "좌훈 약재: 사상자·애엽·고삼·황백·오배자·백반·산초",
        "(경희의료원 한방부인과 좌훈요법실 기준 약재 구성)",
      ],
    },
    {
      title: "질염 관리",
      color: COLOR.navy,
      items: [
        "프로게스테론 질정 사용 빈도가 높아 질염 발생 多",
        "좌훈요법으로 질염 접근 (환자들이 쑥 사용 여부 확인 선호)",
      ],
    },
    {
      title: "한약",
      color: "607D8B",
      items: [
        "시험관 환자들은 의사 말을 전적으로 신뢰 → 한약 복용 소극적",
        "마케팅 업체의 흑염소 광고가 시험관 카페에 많이 노출됨",
        "원하는 환자에게만 선택적으로 시행 권장",
        "목표: 뜸+좌훈을 배란 전까지 시행하여 난자 성숙·자궁 환경 개선 포인트로 접근",
      ],
    },
  ];

  sections.forEach((sec, idx) => {
    const y = 1.05 + idx * 1.22;
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 4.2, y, w: 5.4, h: 0.38,
      fill: { color: sec.color }, line: { type: "none" },
    });
    slide.addText(sec.title, {
      x: 4.3, y, w: 5.2, h: 0.38,
      fontSize: 12, fontFace: FONT, bold: true, color: COLOR.white, valign: "middle",
    });
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 4.2, y: y + 0.38, w: 5.4, h: 0.78,
      fill: { color: idx % 2 === 0 ? COLOR.lightTeal : COLOR.white },
      line: { color: sec.color, width: 0.8 },
    });
    const bullets = sec.items.map(item => [
      { text: "▪  ", options: { fontSize: 10.5, fontFace: FONT, bold: true, color: sec.color } },
      { text: item, options: { fontSize: 10.5, fontFace: FONT, color: COLOR.darkGray, breakLine: true } },
    ]).flat();
    slide.addText(bullets, {
      x: 4.32, y: y + 0.42, w: 5.18, h: 0.7,
      valign: "top", lineSpacingMultiple: 1.4,
    });
  });
}

// ============================================================
// 슬라이드 18: 요약
// ============================================================
{
  const slide = pptx.addSlide();
  addSlideTitle(slide, "요약 — 핵심 포인트");

  const points = [
    { num: "01", title: "AMH는 난포 수 지표, 임신에는 난자의 '질'이 더 중요", text: "AMH ≤1.0: 난소기능저하 / ≤0.4: 심한 기능저하. 나이·생활습관이 난자 질 결정." },
    { num: "02", title: "정액 기준치는 최소치, 무정자증도 IVF 가능", text: "수치가 높을수록 성공률 상승. 정상 정자 1마리라도 있으면 시험관 가능." },
    { num: "03", title: "진료 단계: 난관조영술 → IUI → IVF 순서", text: "난관 개통 확인 후 인공수정, 3회 이상 실패 시 시험관 시술 진행 권장." },
    { num: "04", title: "IVF 과배란 유도: 단기/장기/저자극 중 선택", text: "난소반응 따라 맞춤 요법 선택. GnRH 길항제(즉효)·작용제(예비억제) 역할 구분." },
    { num: "05", title: "배아 배양: 채취 수 적으면 3일, 많으면 5일", text: "모양 안 좋으면 일찍 이식이 유리. 엄마 몸이 최적 환경." },
    { num: "06", title: "이식 결정 기준: E2 < 3,000 / P4 < 1.5 / 이식 날 P4 ≥ 20", text: "기준 초과 시 동결 주기로 전환. 내막 두께 8~10mm가 최적." },
    { num: "07", title: "한의 접근: 배란 전까지 뜸·좌훈으로 자궁 환경 개선", text: "배란 후 온열자극은 배아에 불리. 한약은 원하는 환자에게만 선택적 시행." },
  ];

  points.forEach((pt, idx) => {
    const col = idx < 4 ? 0 : 1;
    const row = col === 0 ? idx : idx - 4;
    const x = col === 0 ? 0.4 : 5.1;
    const y = 1.05 + row * 0.9;
    const w = 4.5;

    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x, y, w, h: 0.82,
      fill: { color: idx % 2 === 0 ? COLOR.lightTeal : COLOR.white },
      line: { color: COLOR.teal, width: 0.6 }, rectRadius: 0.05,
    });
    slide.addShape(pptx.shapes.RECTANGLE, {
      x, y, w: 0.55, h: 0.82,
      fill: { color: idx < 4 ? COLOR.teal : COLOR.navy }, line: { type: "none" },
    });
    slide.addText(pt.num, {
      x, y, w: 0.55, h: 0.82,
      fontSize: 14, fontFace: FONT, bold: true, color: COLOR.white,
      align: "center", valign: "middle",
    });
    slide.addText(pt.title, {
      x: x + 0.62, y: y + 0.03, w: w - 0.7, h: 0.3,
      fontSize: 11, fontFace: FONT, bold: true, color: COLOR.navy, valign: "middle",
    });
    slide.addText(pt.text, {
      x: x + 0.62, y: y + 0.34, w: w - 0.7, h: 0.44,
      fontSize: 9.5, fontFace: FONT, color: COLOR.darkGray, valign: "top",
    });
  });

  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.4, y: 4.65, w: 9.2, h: 0.5,
    fill: { color: COLOR.teal }, line: { type: "none" },
  });
  slide.addText("난임 치료는 개인의 상태에 따라 맞춤형으로 진행됩니다. 담당 의료진과 충분한 상담이 필수입니다.", {
    x: 0.5, y: 4.65, w: 9.0, h: 0.5,
    fontSize: 12, fontFace: FONT, bold: true, color: COLOR.white,
    align: "center", valign: "middle",
  });
}

// ============================================================
// 슬라이드 19: 참고문헌
// ============================================================
{
  const slide = pptx.addSlide();

  slide.addText("참고문헌(References)", {
    x: 0.5, y: 0.2, w: 9.0, h: 0.7,
    fontSize: 28, fontFace: FONT, bold: true, color: COLOR.navy, align: "left",
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.5, y: 0.88, w: 1.4, h: 0.045,
    fill: { color: COLOR.teal }, line: { type: "none" },
  });

  const refs = [
    "김준호. 난임 진료 정리 자료(미출판). 2026.",
    "World Health Organization. WHO laboratory manual for the examination and processing of human semen. 6th ed. WHO Press; 2021.",
    "Macklon NS, Stouffer RL, Giudice LC, Fauser BC. The science behind 25 years of ovarian stimulation for in vitro fertilization. Endocr Rev. 2006;27(2):170-207.",
    "Wu Z, Li R, Ma Y, et al. Effect of HCG-day serum progesterone and oestradiol concentrations on pregnancy outcomes in GnRH agonist cycles. Reprod Biomed Online. 2012;24(5):511-20.",
    "Progesterone elevation on the day of hCG trigger has detrimental effect on live birth rate in low and intermediate ovarian responders, but not in high responders. Hum Reprod. 2019.",
    "Image: In Vitro Fertilization (IVF) - English.png, Wikimedia Commons, CC BY-SA 4.0\nhttps://commons.wikimedia.org/wiki/File:In_Vitro_Fertilization_(IVF)_-_English.png",
    "Image: 2902 IVF-02.jpg (IVF Procedure), Wikimedia Commons, CC BY-SA 3.0\nhttps://commons.wikimedia.org/wiki/File:2902_IVF-02.jpg",
    "Image: Diagram showing human embryo grades for in vitro fertilisation (IVF).jpg, Wikimedia Commons, CC BY-SA 4.0\nhttps://commons.wikimedia.org/wiki/File:Diagram_showing_human_embryo_grades_for_in_vitro_fertilisation_(IVF).jpg",
    "Image: Blastocyst embryo.png, Wikimedia Commons / DBCLS Togo Picture Gallery\nhttps://commons.wikimedia.org/wiki/File:Blastocyst_embryo.png",
    "Image: Cross-Section of an Ovary, Fallopian Tube Parts, Uterus, Female Reproductive Tract Overview, TeachMeAnatomy.info\nhttps://teachmeanatomy.info/pelvis/female-reproductive-tract/",
  ];

  const refTexts = refs.map((ref, idx) => [
    { text: `${idx + 1}.  `, options: { fontSize: 10, fontFace: FONT, bold: true, color: COLOR.teal } },
    { text: ref, options: { fontSize: 10, fontFace: FONT, color: COLOR.medGray, breakLine: true } },
  ]).flat();

  slide.addText(refTexts, {
    x: 0.5, y: 1.1, w: 9.1, h: 4.0,
    valign: "top", lineSpacingMultiple: 1.55,
  });
}

// ============================================================
// 저장
// ============================================================
pptx.writeFile({ fileName: path.resolve(__dirname, "..", "난임_완전가이드.pptx") })
  .then(() => console.log("✅ PPTX 생성 완료: 난임_완전가이드.pptx"))
  .catch((err) => console.error("❌ 오류:", err));
