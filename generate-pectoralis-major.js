const pptxgen = require("pptxgenjs");
const pptx = new pptxgen();

// === 기본 설정 ===
pptx.layout = "LAYOUT_16x9";
pptx.author = "의학 블로그";
pptx.title = "대흉근(Pectoralis Major)의 해부학";

// === 컬러 상수 ===
const C = {
  white: "FFFFFF",
  teal: "0A7E8C",
  navy: "1B2A4A",
  dark: "2D2D2D",
  med: "6B7280",
  ltTeal: "E6F3F5",
  line: "D1D5DB",
  ph: "F0F0F0",
  phTxt: "9CA3AF",
};
const FONT = "Pretendard";

// ============================================================
// 슬라이드 1: 표지
// ============================================================
const s1 = pptx.addSlide();
s1.background = { fill: C.navy };

// 좌측 틸블루 악센트 바
s1.addShape(pptx.shapes.RECTANGLE, {
  x: 0.0, y: 0.0, w: 0.08, h: "100%",
  fill: { color: C.teal },
});

// 제목
s1.addText("대흉근(Pectoralis Major)의\n해부학", {
  x: 0.6, y: 1.5, w: 8.5, h: 1.6,
  fontSize: 42,
  fontFace: FONT,
  bold: true,
  color: C.white,
  align: "left",
  valign: "middle",
  lineSpacingMultiple: 1.2,
});

// 부제
s1.addText("움직임해부학 시리즈 — 상지 근육편", {
  x: 0.6, y: 3.3, w: 8.5, h: 0.5,
  fontSize: 18,
  fontFace: FONT,
  color: C.teal,
  align: "left",
});

// 하단 날짜
s1.addText("2026.02.12  |  의학 블로그", {
  x: 0.6, y: 4.8, w: 8.5, h: 0.4,
  fontSize: 13,
  fontFace: FONT,
  color: C.white,
  transparency: 50,
  align: "left",
});

// ============================================================
// 슬라이드 2: 목차
// ============================================================
const s2 = pptx.addSlide();

s2.addText("목차", {
  x: 0.5, y: 0.4, w: 9.0, h: 0.8,
  fontSize: 36,
  fontFace: FONT,
  bold: true,
  color: C.navy,
});

s2.addShape(pptx.shapes.RECTANGLE, {
  x: 0.5, y: 1.15, w: 1.5, h: 0.04,
  fill: { color: C.teal },
});

const tocItems = [
  "개요 및 해부학적 위치",
  "기시(Origin)와 정지(Insertion)",
  "작용(Action)",
  "신경지배(Innervation)",
  "임상적 의의",
];

const tocTexts = [];
tocItems.forEach((item, idx) => {
  tocTexts.push(
    {
      text: `${String(idx + 1).padStart(2, "0")}`,
      options: { fontSize: 22, fontFace: FONT, bold: true, color: C.teal },
    },
    {
      text: `   ${item}`,
      options: { fontSize: 20, fontFace: FONT, color: C.dark, breakLine: true },
    }
  );
});

s2.addText(tocTexts, {
  x: 0.5, y: 1.5, w: 8.5, h: 3.8,
  valign: "top",
  lineSpacingMultiple: 1.9,
  margin: 0,
});

// ============================================================
// 슬라이드 3: 개요 + 해부학 이미지 (레이아웃 A)
// ============================================================
const s3 = pptx.addSlide();

s3.addText("개요 및 해부학적 위치", {
  x: 0.5, y: 0.3, w: 9.0, h: 0.7,
  fontSize: 32,
  fontFace: FONT,
  bold: true,
  color: C.navy,
});

s3.addShape(pptx.shapes.RECTANGLE, {
  x: 0.5, y: 0.95, w: 1.2, h: 0.04,
  fill: { color: C.teal },
});

// 좌측: Gray's Anatomy 이미지
s3.addImage({
  path: "./gray410.png",
  x: 0.5, y: 1.3, w: 4.2, h: 3.6,
  sizing: { type: "contain", w: 4.2, h: 3.6 },
});

// 이미지 캡션
s3.addText("Gray's Anatomy, Plate 410 (Public Domain)", {
  x: 0.5, y: 4.9, w: 4.2, h: 0.3,
  fontSize: 9,
  fontFace: FONT,
  italic: true,
  color: C.med,
  align: "center",
});

// 우측: 개요 텍스트
const overviewTexts = [
  {
    text: "대흉근(Pectoralis Major)",
    options: { fontSize: 18, fontFace: FONT, bold: true, color: C.teal, breakLine: true },
  },
  {
    text: "\n",
    options: { fontSize: 8, breakLine: true },
  },
  {
    text: "흉벽(Chest Wall) 전면에 위치한 크고 부채꼴 모양의 근육으로, 상지(Upper Limb)의 움직임에 핵심적인 역할을 한다.",
    options: { fontSize: 14, fontFace: FONT, color: C.dark, breakLine: true },
  },
  {
    text: "\n",
    options: { fontSize: 8, breakLine: true },
  },
  {
    text: "구성",
    options: { fontSize: 15, fontFace: FONT, bold: true, color: C.navy, breakLine: true },
  },
  {
    text: "\n",
    options: { fontSize: 4, breakLine: true },
  },
  {
    text: "\u2022  쇄골부(Clavicular Head)",
    options: { fontSize: 13, fontFace: FONT, color: C.dark, breakLine: true },
  },
  {
    text: "\u2022  흉늑부(Sternocostal Head)",
    options: { fontSize: 13, fontFace: FONT, color: C.dark, breakLine: true },
  },
  {
    text: "\u2022  복부(Abdominal Part)",
    options: { fontSize: 13, fontFace: FONT, color: C.dark, breakLine: true },
  },
  {
    text: "\n",
    options: { fontSize: 8, breakLine: true },
  },
  {
    text: "위치: 전흉벽(Anterior Thoracic Wall)의 표층",
    options: { fontSize: 13, fontFace: FONT, color: C.med, italic: true },
  },
];

s3.addText(overviewTexts, {
  x: 5.1, y: 1.3, w: 4.4, h: 3.8,
  valign: "top",
  margin: [8, 8, 8, 8],
});

// ============================================================
// 슬라이드 4: 기시와 정지 — 정보표 (레이아웃 B)
// ============================================================
const s4 = pptx.addSlide();

s4.addText("기시(Origin)와 정지(Insertion)", {
  x: 0.5, y: 0.3, w: 9.0, h: 0.7,
  fontSize: 32,
  fontFace: FONT,
  bold: true,
  color: C.navy,
});

s4.addShape(pptx.shapes.RECTANGLE, {
  x: 0.5, y: 0.95, w: 1.2, h: 0.04,
  fill: { color: C.teal },
});

const cellStyle = (isBg) => ({
  fontSize: 12,
  fontFace: FONT,
  color: isBg ? C.white : C.dark,
  valign: "middle",
  margin: [5, 8, 5, 8],
  ...(isBg ? { fill: { color: C.teal }, bold: true } : {}),
});

const tableRows = [
  // 헤더
  [
    { text: "구분", options: { ...cellStyle(true), align: "center" } },
    { text: "부위", options: { ...cellStyle(true), align: "center" } },
    { text: "부착 부위", options: { ...cellStyle(true), align: "center" } },
  ],
  // 기시 - 쇄골부
  [
    { text: "기시\n(Origin)", options: { ...cellStyle(false), bold: true, color: C.teal, align: "center", rowspan: 3 } },
    { text: "쇄골부\n(Clavicular Head)", options: cellStyle(false) },
    { text: "쇄골(Clavicle) 내측 1/2의 전면", options: cellStyle(false) },
  ],
  // 기시 - 흉늑부
  [
    { text: "", options: cellStyle(false) },
    { text: "흉늑부\n(Sternocostal Head)", options: cellStyle(false) },
    { text: "흉골(Sternum) 전면, 제1~6늑연골(Costal Cartilage)", options: cellStyle(false) },
  ],
  // 기시 - 복부
  [
    { text: "", options: cellStyle(false) },
    { text: "복부\n(Abdominal Part)", options: cellStyle(false) },
    { text: "외복사근건막(External Oblique Aponeurosis)", options: cellStyle(false) },
  ],
  // 정지
  [
    { text: "정지\n(Insertion)", options: { ...cellStyle(false), bold: true, color: C.teal, align: "center" } },
    { text: "공통", options: cellStyle(false) },
    { text: "상완골(Humerus) 대결절릉\n(Crest of Greater Tubercle)\n= 결절간고랑(Intertubercular Groove) 외측연", options: cellStyle(false) },
  ],
];

s4.addTable(tableRows, {
  x: 0.5, y: 1.2, w: 9.0,
  border: { type: "solid", color: C.line, pt: 0.5 },
  colW: [1.5, 2.5, 5.0],
  rowH: [0.45, 0.7, 0.7, 0.7, 0.85],
  autoPage: false,
});

// 보충 설명
s4.addShape(pptx.shapes.RECTANGLE, {
  x: 0.5, y: 4.65, w: 9.0, h: 0.55,
  fill: { color: C.ltTeal },
});

s4.addText("\u26A0  세 부분의 섬유가 정지 부위에서 꼬이며 합쳐진다 — 복부 섬유가 가장 위쪽, 쇄골부 섬유가 가장 아래쪽에 부착", {
  x: 0.7, y: 4.65, w: 8.6, h: 0.55,
  fontSize: 12,
  fontFace: FONT,
  color: C.navy,
  valign: "middle",
  margin: 0,
});

// ============================================================
// 슬라이드 5: 작용 — 비교 다이어그램 (레이아웃 D)
// ============================================================
const s5 = pptx.addSlide();

s5.addText("작용(Action)", {
  x: 0.5, y: 0.3, w: 9.0, h: 0.7,
  fontSize: 32,
  fontFace: FONT,
  bold: true,
  color: C.navy,
});

s5.addShape(pptx.shapes.RECTANGLE, {
  x: 0.5, y: 0.95, w: 1.2, h: 0.04,
  fill: { color: C.teal },
});

// 3컬럼: 쇄골부 / 흉늑부 / 전체
const actionCols = [
  {
    title: "쇄골부(Clavicular Head)",
    color: C.teal,
    items: [
      "어깨관절 굴곡\n(Shoulder Flexion)",
      "수평내전\n(Horizontal Adduction)",
      "내회전\n(Internal Rotation)",
    ],
  },
  {
    title: "흉늑부(Sternocostal Head)",
    color: C.navy,
    items: [
      "어깨관절 신전\n(Shoulder Extension)\n— 굴곡 위치에서",
      "내전\n(Adduction)",
      "내회전\n(Internal Rotation)",
    ],
  },
  {
    title: "전체(Combined)",
    color: C.teal,
    items: [
      "강력한 내전\n(Powerful Adduction)",
      "강력한 내회전\n(Internal Rotation)",
      "보조 호흡근\n(Accessory Respiratory M.)\n— 상지 고정 시",
    ],
  },
];

const colW = 2.8;
const gap = 0.3;
const startX = (10 - (colW * 3 + gap * 2)) / 2;

actionCols.forEach((col, idx) => {
  const x = startX + idx * (colW + gap);

  // 헤더
  s5.addShape(pptx.shapes.RECTANGLE, {
    x: x, y: 1.3, w: colW, h: 0.55,
    fill: { color: col.color },
  });

  s5.addText(col.title, {
    x: x, y: 1.3, w: colW, h: 0.55,
    fontSize: 13,
    fontFace: FONT,
    bold: true,
    color: C.white,
    align: "center",
    valign: "middle",
    margin: 0,
  });

  // 본문 배경
  s5.addShape(pptx.shapes.RECTANGLE, {
    x: x, y: 1.85, w: colW, h: 3.2,
    fill: { color: C.ltTeal },
    line: { color: C.line, width: 0.5 },
  });

  // 항목들
  col.items.forEach((item, iIdx) => {
    const itemY = 1.95 + iIdx * 1.0;

    // 번호 원
    s5.addShape(pptx.shapes.OVAL, {
      x: x + 0.15, y: itemY + 0.05, w: 0.3, h: 0.3,
      fill: { color: col.color },
    });

    s5.addText(String(iIdx + 1), {
      x: x + 0.15, y: itemY + 0.05, w: 0.3, h: 0.3,
      fontSize: 11,
      fontFace: FONT,
      bold: true,
      color: C.white,
      align: "center",
      valign: "middle",
      margin: 0,
    });

    // 텍스트
    s5.addText(item, {
      x: x + 0.55, y: itemY, w: colW - 0.7, h: 0.9,
      fontSize: 11,
      fontFace: FONT,
      color: C.dark,
      valign: "top",
      margin: 0,
    });
  });
});

// ============================================================
// 슬라이드 6: 신경지배 — 흐름도 (레이아웃 C)
// ============================================================
const s6 = pptx.addSlide();

s6.addText("신경지배(Innervation)", {
  x: 0.5, y: 0.3, w: 9.0, h: 0.7,
  fontSize: 32,
  fontFace: FONT,
  bold: true,
  color: C.navy,
});

s6.addShape(pptx.shapes.RECTANGLE, {
  x: 0.5, y: 0.95, w: 1.2, h: 0.04,
  fill: { color: C.teal },
});

// 흐름도 상단
const flowSteps = [
  { label: "척수\n(Spinal Cord)\nC5-T1", bg: C.navy },
  { label: "전지\n(Ventral\nRami)", bg: C.dark },
  { label: "상완신경총\n(Brachial\nPlexus)", bg: C.dark },
  { label: "외측속\n(Lateral\nCord)", bg: C.teal },
  { label: "내측속\n(Medial\nCord)", bg: C.teal },
];

const fBoxW = 1.5;
const fBoxH = 0.9;
const fArrowW = 0.35;
const fStartX = 0.5;
const fY = 1.4;

flowSteps.forEach((step, idx) => {
  const x = fStartX + idx * (fBoxW + fArrowW);

  s6.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: fY, w: fBoxW, h: fBoxH,
    fill: { color: step.bg },
    rectRadius: 0.05,
  });

  s6.addText(step.label, {
    x: x, y: fY, w: fBoxW, h: fBoxH,
    fontSize: 10,
    fontFace: FONT,
    bold: true,
    color: C.white,
    align: "center",
    valign: "middle",
    margin: 0,
  });

  if (idx < flowSteps.length - 1) {
    s6.addText("\u25B6", {
      x: x + fBoxW, y: fY, w: fArrowW, h: fBoxH,
      fontSize: 14,
      color: C.teal,
      align: "center",
      valign: "middle",
      margin: 0,
    });
  }
});

// 하단: 두 신경으로 분기
// 외측 흉근신경
s6.addShape(pptx.shapes.LINE, {
  x: fStartX + 3 * (fBoxW + fArrowW) + fBoxW / 2,
  y: fY + fBoxH,
  w: 0, h: 0.5,
  line: { color: C.teal, width: 2 },
});

s6.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 1.0, y: 3.0, w: 3.5, h: 0.7,
  fill: { color: C.teal },
  rectRadius: 0.05,
});

s6.addText("외측 흉근신경(Lateral Pectoral N.)\nC5, C6, C7", {
  x: 1.0, y: 3.0, w: 3.5, h: 0.7,
  fontSize: 12,
  fontFace: FONT,
  bold: true,
  color: C.white,
  align: "center",
  valign: "middle",
  margin: 0,
});

// 내측 흉근신경
s6.addShape(pptx.shapes.LINE, {
  x: fStartX + 4 * (fBoxW + fArrowW) + fBoxW / 2,
  y: fY + fBoxH,
  w: 0, h: 0.5,
  line: { color: C.teal, width: 2 },
});

s6.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 5.5, y: 3.0, w: 3.5, h: 0.7,
  fill: { color: C.navy },
  rectRadius: 0.05,
});

s6.addText("내측 흉근신경(Medial Pectoral N.)\nC8, T1", {
  x: 5.5, y: 3.0, w: 3.5, h: 0.7,
  fontSize: 12,
  fontFace: FONT,
  bold: true,
  color: C.white,
  align: "center",
  valign: "middle",
  margin: 0,
});

// 화살표 → 대흉근
s6.addShape(pptx.shapes.LINE, {
  x: 2.75, y: 3.7, w: 0, h: 0.4,
  line: { color: C.teal, width: 2 },
});
s6.addShape(pptx.shapes.LINE, {
  x: 7.25, y: 3.7, w: 0, h: 0.4,
  line: { color: C.navy, width: 2 },
});

// 대흉근 합류 박스
s6.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 2.5, y: 4.1, w: 5.0, h: 0.65,
  fill: { color: C.ltTeal },
  line: { color: C.teal, width: 1.5 },
  rectRadius: 0.05,
});

s6.addText("대흉근(Pectoralis Major)", {
  x: 2.5, y: 4.1, w: 5.0, h: 0.65,
  fontSize: 16,
  fontFace: FONT,
  bold: true,
  color: C.navy,
  align: "center",
  valign: "middle",
  margin: 0,
});

// ============================================================
// 슬라이드 7: 임상적 의의 + 이미지 (레이아웃 A)
// ============================================================
const s7 = pptx.addSlide();

s7.addText("임상적 의의", {
  x: 0.5, y: 0.3, w: 9.0, h: 0.7,
  fontSize: 32,
  fontFace: FONT,
  bold: true,
  color: C.navy,
});

s7.addShape(pptx.shapes.RECTANGLE, {
  x: 0.5, y: 0.95, w: 1.2, h: 0.04,
  fill: { color: C.teal },
});

// 좌측: Anatomography 이미지
s7.addImage({
  path: "./pectoralis_major.png",
  x: 0.5, y: 1.3, w: 4.2, h: 3.5,
  sizing: { type: "contain", w: 4.2, h: 3.5 },
});

s7.addText("Anatomography, Wikimedia Commons (CC BY-SA 3.0)", {
  x: 0.5, y: 4.85, w: 4.2, h: 0.3,
  fontSize: 9,
  fontFace: FONT,
  italic: true,
  color: C.med,
  align: "center",
});

// 우측: 임상 내용
const clinicalRows = [
  [
    { text: "대흉근 파열\n(Rupture)", options: { fontSize: 12, fontFace: FONT, bold: true, color: C.white, fill: { color: C.teal }, valign: "middle", margin: [4, 6, 4, 6] } },
    { text: "벤치프레스 등 과부하 시 발생\n건-골 접합부(Tendon-Bone Junction)에서 호발\n전면 액와주름(Anterior Axillary Fold) 소실이 특징적 소견", options: { fontSize: 11, fontFace: FONT, color: C.dark, valign: "middle", margin: [4, 6, 4, 6] } },
  ],
  [
    { text: "Poland 증후군\n(Poland Syndrome)", options: { fontSize: 12, fontFace: FONT, bold: true, color: C.white, fill: { color: C.teal }, valign: "middle", margin: [4, 6, 4, 6] } },
    { text: "대흉근의 선천적 결손\n흉늑부(Sternocostal Head) 부재가 가장 흔함\n동측 상지 기형 동반 가능", options: { fontSize: 11, fontFace: FONT, color: C.dark, valign: "middle", margin: [4, 6, 4, 6] } },
  ],
  [
    { text: "단축/긴장\n(Tightness)", options: { fontSize: 12, fontFace: FONT, bold: true, color: C.white, fill: { color: C.teal }, valign: "middle", margin: [4, 6, 4, 6] } },
    { text: "둥근어깨(Rounded Shoulder) 자세 유발\n상위교차증후군(Upper Crossed Syndrome)의 핵심 요소\n도수치료 시 스트레칭 대상 근육", options: { fontSize: 11, fontFace: FONT, color: C.dark, valign: "middle", margin: [4, 6, 4, 6] } },
  ],
];

s7.addTable(clinicalRows, {
  x: 5.0, y: 1.3, w: 4.5,
  border: { type: "solid", color: C.line, pt: 0.5 },
  colW: [1.3, 3.2],
  rowH: [1.1, 1.1, 1.1],
  autoPage: false,
});

// ============================================================
// 슬라이드 8: 요약
// ============================================================
const s8 = pptx.addSlide();

s8.addText("요약", {
  x: 0.5, y: 0.3, w: 9.0, h: 0.7,
  fontSize: 36,
  fontFace: FONT,
  bold: true,
  color: C.navy,
});

s8.addShape(pptx.shapes.RECTANGLE, {
  x: 0.5, y: 0.95, w: 1.2, h: 0.04,
  fill: { color: C.teal },
});

const summaryPoints = [
  "대흉근은 쇄골부, 흉늑부, 복부로 구성된 대근육이다",
  "상완골 대결절릉에 부착하며, 섬유가 꼬이며 합쳐지는 독특한 구조를 갖는다",
  "어깨관절의 내전(Adduction), 굴곡(Flexion), 내회전(Internal Rotation)을 담당한다",
  "외측 흉근신경(C5-C7)과 내측 흉근신경(C8-T1)의 이중 신경지배를 받는다",
  "임상적으로 파열, Poland 증후군, 상위교차증후군과 관련이 깊다",
];

const sumTexts = [];
summaryPoints.forEach((pt, idx) => {
  sumTexts.push(
    {
      text: `${idx + 1}`,
      options: { fontSize: 17, fontFace: FONT, bold: true, color: C.teal },
    },
    {
      text: `   ${pt}`,
      options: { fontSize: 15, fontFace: FONT, color: C.dark, breakLine: true },
    }
  );
});

s8.addText(sumTexts, {
  x: 0.5, y: 1.3, w: 9.0, h: 3.0,
  valign: "top",
  lineSpacingMultiple: 1.8,
  margin: 0,
});

// 하단 강조 박스
s8.addShape(pptx.shapes.RECTANGLE, {
  x: 0.5, y: 4.5, w: 9.0, h: 0.65,
  fill: { color: C.teal },
});

s8.addText("핵심: 대흉근은 상지 내전과 내회전의 주동근이며, 자세 교정과 재활에서 핵심적으로 다루는 근육이다", {
  x: 0.7, y: 4.5, w: 8.6, h: 0.65,
  fontSize: 13,
  fontFace: FONT,
  bold: true,
  color: C.white,
  align: "center",
  valign: "middle",
  margin: 0,
});

// ============================================================
// 슬라이드 9: 참고문헌
// ============================================================
const s9 = pptx.addSlide();

s9.addText("참고문헌(References)", {
  x: 0.5, y: 0.3, w: 9.0, h: 0.7,
  fontSize: 28,
  fontFace: FONT,
  bold: true,
  color: C.navy,
});

s9.addShape(pptx.shapes.RECTANGLE, {
  x: 0.5, y: 0.95, w: 1.2, h: 0.04,
  fill: { color: C.teal },
});

const refs = [
  "Neumann DA. Kinesiology of the Musculoskeletal System: Foundations for Rehabilitation. 3rd ed. Elsevier; 2017.",
  "Moore KL, Dalley AF, Agur AMR. Clinically Oriented Anatomy. 8th ed. Wolters Kluwer; 2018.",
  "Standring S. Gray's Anatomy: The Anatomical Basis of Clinical Practice. 42nd ed. Elsevier; 2020.",
  "Image: Gray410.png, Wikimedia Commons, Public Domain, https://commons.wikimedia.org/wiki/File:Gray410.png",
  "Image: Pectoralis_major.png, by Anatomography (BodyParts3D), Wikimedia Commons, CC BY-SA 3.0, https://commons.wikimedia.org/wiki/File:Pectoralis_major.png",
];

const refTexts = refs.map((ref, idx) => ({
  text: `${idx + 1}. ${ref}`,
  options: {
    fontSize: 11,
    fontFace: FONT,
    color: C.med,
    breakLine: true,
  },
}));

s9.addText(refTexts, {
  x: 0.5, y: 1.2, w: 9.0, h: 4.0,
  valign: "top",
  lineSpacingMultiple: 1.7,
  margin: 0,
});

// === 파일 저장 ===
pptx.writeFile({ fileName: "대흉근_해부학.pptx" })
  .then(() => console.log("PPTX 생성 완료: 대흉근_해부학.pptx"))
  .catch((err) => console.error("Error:", err));
