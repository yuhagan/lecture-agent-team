const pptxgen = require("pptxgenjs");
const pptx = new pptxgen();

pptx.layout = "LAYOUT_16x9";
pptx.author = "의학 블로그";
pptx.title = "승모근(Trapezius)의 움직임해부학";

// 컬러 상수
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

const FONT = "Pretendard";

// ============================================================
// 슬라이드 1: 표지
// ============================================================
const slide1 = pptx.addSlide();
slide1.background = { fill: COLOR.navy };

slide1.addShape(pptx.shapes.RECTANGLE, {
  x: 0.0, y: 0.0, w: 0.08, h: "100%",
  fill: { color: COLOR.teal },
});

slide1.addText("승모근(Trapezius)의\n움직임해부학", {
  x: 0.6, y: 1.5, w: 8.5, h: 1.5,
  fontSize: 44,
  fontFace: FONT,
  bold: true,
  color: COLOR.white,
  align: "left",
  valign: "middle",
  lineSpacingMultiple: 1.2,
});

slide1.addText("'옷걸이 근육'이 목·어깨 통증의 열쇠입니다", {
  x: 0.6, y: 3.2, w: 8.5, h: 0.6,
  fontSize: 18,
  fontFace: FONT,
  color: COLOR.teal,
  align: "left",
});

slide1.addText("2026.02.12 | 의학 블로그", {
  x: 0.6, y: 4.8, w: 8.5, h: 0.4,
  fontSize: 14,
  fontFace: FONT,
  color: COLOR.white,
  transparency: 50,
  align: "left",
});

// ============================================================
// 슬라이드 2: 목차
// ============================================================
const slide2 = pptx.addSlide();

slide2.addText("목차", {
  x: 0.5, y: 0.4, w: 9.0, h: 0.8,
  fontSize: 36,
  fontFace: FONT,
  bold: true,
  color: COLOR.navy,
  align: "left",
});

slide2.addShape(pptx.shapes.RECTANGLE, {
  x: 0.5, y: 1.2, w: 1.5, h: 0.04,
  fill: { color: COLOR.teal },
});

const tocItems = [
  "승모근 개요",
  "구조와 부착 (기시/정지)",
  "기능 (상부/중부/하부)",
  "Scapulohumeral Rhythm",
  "임상 포인트: TrP와 방사통",
  "활성화 요인과 교정",
];

const tocTexts = tocItems.map((item, idx) => [
  {
    text: `${String(idx + 1).padStart(2, "0")}  `,
    options: { fontSize: 20, fontFace: FONT, bold: true, color: COLOR.teal },
  },
  {
    text: item,
    options: { fontSize: 20, fontFace: FONT, color: COLOR.darkGray, breakLine: true },
  },
]).flat();

slide2.addText(tocTexts, {
  x: 0.5, y: 1.6, w: 8.5, h: 3.5,
  valign: "top",
  lineSpacingMultiple: 1.8,
  margin: 0,
});

// ============================================================
// 슬라이드 3: 승모근 개요 (레이아웃 A: 이미지 + 정보표)
// ============================================================
const slide3 = pptx.addSlide();

slide3.addText("승모근(Trapezius) — 개요", {
  x: 0.5, y: 0.3, w: 9.0, h: 0.7,
  fontSize: 32,
  fontFace: FONT,
  bold: true,
  color: COLOR.navy,
  align: "left",
});

slide3.addShape(pptx.shapes.RECTANGLE, {
  x: 0.5, y: 1.0, w: 1.2, h: 0.04,
  fill: { color: COLOR.teal },
});

// 좌측: Gray's Anatomy 이미지
slide3.addImage({
  path: "./trapezius_gray409.png",
  x: 0.5, y: 1.3, w: 4.2, h: 3.8,
  sizing: { type: "contain", w: 4.2, h: 3.8 },
});

// 이미지 캡션
slide3.addText("Gray's Anatomy, 1918 (Public Domain)", {
  x: 0.5, y: 5.1, w: 4.2, h: 0.3,
  fontSize: 9,
  fontFace: FONT,
  italic: true,
  color: COLOR.medGray,
  align: "center",
});

// 우측: 개요 정보
const overviewData = [
  ["별명(Alias)", "'옷걸이 근육(The Coat Hanger)'\n양쪽 상부 섬유가 옷걸이 형태"],
  ["형태(Shape)", "다이아몬드(Diamond) 형태\n후두골~T12 / 쇄골~견갑극"],
  ["구분(Division)", "상부(Upper)\n중부(Middle)\n하부(Lower)"],
  ["신경지배\n(Innervation)", "운동: 부신경(CN XI)\n감각: C2~C4 경추신경"],
];

const overviewRows = overviewData.map((row) => [
  {
    text: row[0],
    options: {
      fontSize: 12, fontFace: FONT, bold: true, color: COLOR.white,
      fill: { color: COLOR.teal }, valign: "middle", margin: [4, 6, 4, 6],
    },
  },
  {
    text: row[1],
    options: {
      fontSize: 11, fontFace: FONT, color: COLOR.darkGray,
      valign: "middle", margin: [4, 6, 4, 6],
    },
  },
]);

slide3.addTable(overviewRows, {
  x: 5.0, y: 1.3, w: 4.5,
  border: { type: "solid", color: COLOR.lineGray, pt: 0.5 },
  colW: [1.3, 3.2],
  rowH: 0.9,
  autoPage: false,
});

// ============================================================
// 슬라이드 4: 기시와 정지 (레이아웃 B: 전폭 3행 정보표)
// ============================================================
const slide4 = pptx.addSlide();

slide4.addText("승모근(Trapezius) — 기시(Origin)와 정지(Insertion)", {
  x: 0.5, y: 0.3, w: 9.0, h: 0.7,
  fontSize: 28,
  fontFace: FONT,
  bold: true,
  color: COLOR.navy,
  align: "left",
});

slide4.addShape(pptx.shapes.RECTANGLE, {
  x: 0.5, y: 1.0, w: 1.2, h: 0.04,
  fill: { color: COLOR.teal },
});

// 3행 표: 상부/중부/하부
const headerRow4 = [
  { text: "구분", options: { fontSize: 13, fontFace: FONT, bold: true, color: COLOR.white, fill: { color: COLOR.teal }, align: "center", valign: "middle" } },
  { text: "기시(Origin)", options: { fontSize: 13, fontFace: FONT, bold: true, color: COLOR.white, fill: { color: COLOR.teal }, align: "center", valign: "middle" } },
  { text: "정지(Insertion)", options: { fontSize: 13, fontFace: FONT, bold: true, color: COLOR.white, fill: { color: COLOR.teal }, align: "center", valign: "middle" } },
];

const dataRows4 = [
  [
    { text: "상부 섬유\n(Upper)", options: { fontSize: 12, fontFace: FONT, bold: true, color: COLOR.teal, valign: "middle", align: "center", margin: [4, 6, 4, 6] } },
    { text: "상항선(Superior Nuchal Line) 내측 1/3\n항인대(Ligamentum Nuchae)\nC1~C5 극돌기(Spinous Process)", options: { fontSize: 11, fontFace: FONT, color: COLOR.darkGray, valign: "middle", margin: [4, 6, 4, 6] } },
    { text: "쇄골(Clavicle) 외측 1/3", options: { fontSize: 11, fontFace: FONT, color: COLOR.darkGray, valign: "middle", margin: [4, 6, 4, 6] } },
  ],
  [
    { text: "중부 섬유\n(Middle)", options: { fontSize: 12, fontFace: FONT, bold: true, color: COLOR.teal, valign: "middle", align: "center", margin: [4, 6, 4, 6] } },
    { text: "C5~T3 극돌기(Spinous Process)\n극간인대(Interspinous Ligament)", options: { fontSize: 11, fontFace: FONT, color: COLOR.darkGray, valign: "middle", margin: [4, 6, 4, 6] } },
    { text: "견봉(Acromion)\n견갑극(Spine of Scapula) 상순", options: { fontSize: 11, fontFace: FONT, color: COLOR.darkGray, valign: "middle", margin: [4, 6, 4, 6] } },
  ],
  [
    { text: "하부 섬유\n(Lower)", options: { fontSize: 12, fontFace: FONT, bold: true, color: COLOR.teal, valign: "middle", align: "center", margin: [4, 6, 4, 6] } },
    { text: "T4~T12 극돌기(Spinous Process)\n극간인대(Interspinous Ligament)", options: { fontSize: 11, fontFace: FONT, color: COLOR.darkGray, valign: "middle", margin: [4, 6, 4, 6] } },
    { text: "견갑극(Spine of Scapula)\n내측 끝 결절", options: { fontSize: 11, fontFace: FONT, color: COLOR.darkGray, valign: "middle", margin: [4, 6, 4, 6] } },
  ],
];

slide4.addTable([headerRow4, ...dataRows4], {
  x: 0.5, y: 1.3, w: 9.0,
  border: { type: "solid", color: COLOR.lineGray, pt: 0.5 },
  colW: [1.5, 4.0, 3.5],
  rowH: [0.5, 1.0, 1.0, 1.0],
  autoPage: false,
});

// 하단 핵심 노트
slide4.addShape(pptx.shapes.RECTANGLE, {
  x: 0.5, y: 4.6, w: 9.0, h: 0.6,
  fill: { color: COLOR.lightTeal },
  line: { color: COLOR.teal, width: 1 },
});

slide4.addText("핵심: 운동은 부신경(Spinal Accessory N., CN XI), 감각은 C2~C4 경추신경이 담당한다", {
  x: 0.5, y: 4.6, w: 9.0, h: 0.6,
  fontSize: 14,
  fontFace: FONT,
  bold: true,
  color: COLOR.navy,
  align: "center",
  valign: "middle",
  margin: 0,
});

// ============================================================
// 슬라이드 5: 기능 — 비교/대조 (레이아웃 D: 3컬럼)
// ============================================================
const slide5 = pptx.addSlide();

slide5.addText("승모근(Trapezius) — 기능(Action)", {
  x: 0.5, y: 0.3, w: 9.0, h: 0.7,
  fontSize: 32,
  fontFace: FONT,
  bold: true,
  color: COLOR.navy,
  align: "left",
});

slide5.addShape(pptx.shapes.RECTANGLE, {
  x: 0.5, y: 1.0, w: 1.2, h: 0.04,
  fill: { color: COLOR.teal },
});

// 3컬럼: 상부 / 중부 / 하부
const columns5 = [
  {
    title: "상부(Upper)",
    color: COLOR.teal,
    items: [
      "어깨 거상(Elevation)",
      "견갑골 상방 회전\n(Upward Rotation)",
      "반대측 머리 회전 보조",
      "타이핑 시 EMG 증가",
    ],
  },
  {
    title: "중부(Middle)",
    color: COLOR.navy,
    items: [
      "견갑골 후인\n(Retraction)",
      "팔 굴곡·외전 시\n견갑골 고정(Stabilization)",
    ],
  },
  {
    title: "하부(Lower)",
    color: COLOR.teal,
    items: [
      "견갑골 내측연 하강",
      "견갑골 후인 보조",
      "관절와 상방 회전 보조\n(Upward Rotation)",
    ],
  },
];

const colCount = columns5.length;
const colW5 = (9.0 - (colCount - 1) * 0.3) / colCount;
const startX5 = 0.5;

columns5.forEach((col, idx) => {
  const x = startX5 + idx * (colW5 + 0.3);

  slide5.addShape(pptx.shapes.RECTANGLE, {
    x: x, y: 1.4, w: colW5, h: 0.55,
    fill: { color: col.color },
  });

  slide5.addText(col.title, {
    x: x, y: 1.4, w: colW5, h: 0.55,
    fontSize: 16, fontFace: FONT, bold: true, color: COLOR.white,
    align: "center", valign: "middle", margin: 0,
  });

  slide5.addShape(pptx.shapes.RECTANGLE, {
    x: x, y: 1.95, w: colW5, h: 2.8,
    fill: { color: COLOR.lightTeal },
    line: { color: COLOR.lineGray, width: 0.5 },
  });

  const itemTexts = col.items.map((item) => ({
    text: `\u2022  ${item}`,
    options: { fontSize: 13, fontFace: FONT, color: COLOR.darkGray, breakLine: true },
  }));

  slide5.addText(itemTexts, {
    x: x + 0.15, y: 2.1, w: colW5 - 0.3, h: 2.5,
    valign: "top",
    lineSpacingMultiple: 1.5,
    margin: 0,
  });
});

// ============================================================
// 슬라이드 6: Scapulohumeral Rhythm (흐름도)
// ============================================================
const slide6 = pptx.addSlide();

slide6.addText("견갑상완 리듬(Scapulohumeral Rhythm)", {
  x: 0.5, y: 0.3, w: 9.0, h: 0.7,
  fontSize: 30,
  fontFace: FONT,
  bold: true,
  color: COLOR.navy,
  align: "left",
});

slide6.addShape(pptx.shapes.RECTANGLE, {
  x: 0.5, y: 1.0, w: 1.2, h: 0.04,
  fill: { color: COLOR.teal },
});

// 설명 텍스트
slide6.addText("팔을 외전(Abduction)시킬 때, 견갑골(Scapula) 회전과 상완골(Humerus) 움직임이 함께 작용합니다.\n견갑골 움직임 없이는 팔을 완전히 올릴 수 없습니다.", {
  x: 0.5, y: 1.3, w: 9.0, h: 0.8,
  fontSize: 14,
  fontFace: FONT,
  color: COLOR.darkGray,
  align: "left",
  valign: "top",
  margin: 0,
});

// 흐름도
const steps6 = [
  "승모근\n(Trapezius)",
  "견갑골 회전\n(Scapular\nRotation)",
  "극상근·삼각근\n(Supraspinatus\n·Deltoid)",
  "상완골 움직임\n(Glenohumeral\nMovement)",
  "완전 외전\n(Full\nAbduction)",
];

const boxW6 = 1.6;
const boxH6 = 0.9;
const arrowW6 = 0.3;
const totalW6 = steps6.length * boxW6 + (steps6.length - 1) * arrowW6;
const startX6 = (10.0 - totalW6) / 2;
const centerY6 = 2.8;

steps6.forEach((step, idx) => {
  const x = startX6 + idx * (boxW6 + arrowW6);

  slide6.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: centerY6, w: boxW6, h: boxH6,
    fill: { color: idx === 0 ? COLOR.navy : idx === steps6.length - 1 ? COLOR.teal : COLOR.lightTeal },
    rectRadius: 0.05,
  });

  slide6.addText(step, {
    x: x, y: centerY6, w: boxW6, h: boxH6,
    fontSize: 10, fontFace: FONT,
    bold: idx === 0 || idx === steps6.length - 1,
    color: idx === 0 || idx === steps6.length - 1 ? COLOR.white : COLOR.darkGray,
    align: "center", valign: "middle", margin: 0,
  });

  if (idx < steps6.length - 1) {
    slide6.addText("\u25B6", {
      x: x + boxW6, y: centerY6, w: arrowW6, h: boxH6,
      fontSize: 14, color: COLOR.teal,
      align: "center", valign: "middle", margin: 0,
    });
  }
});

// 하단 강조 박스
slide6.addShape(pptx.shapes.RECTANGLE, {
  x: 0.5, y: 4.3, w: 9.0, h: 0.7,
  fill: { color: COLOR.lightTeal },
  line: { color: COLOR.teal, width: 1 },
});

slide6.addText("타이핑 시에도 승모근 상부 EMG가 상승 — 긴장된 자세, 높은 키보드가 원인", {
  x: 0.5, y: 4.3, w: 9.0, h: 0.7,
  fontSize: 13, fontFace: FONT, bold: true, color: COLOR.navy,
  align: "center", valign: "middle", margin: 0,
});

// ============================================================
// 슬라이드 7: TrP와 방사통 (표)
// ============================================================
const slide7 = pptx.addSlide();

slide7.addText("임상 포인트 — TrP(Trigger Point)와 방사통", {
  x: 0.5, y: 0.3, w: 9.0, h: 0.7,
  fontSize: 28,
  fontFace: FONT,
  bold: true,
  color: COLOR.navy,
  align: "left",
});

slide7.addShape(pptx.shapes.RECTANGLE, {
  x: 0.5, y: 1.0, w: 1.2, h: 0.04,
  fill: { color: COLOR.teal },
});

const trpHeader = [
  { text: "TrP", options: { fontSize: 12, fontFace: FONT, bold: true, color: COLOR.white, fill: { color: COLOR.teal }, align: "center", valign: "middle" } },
  { text: "위치(Location)", options: { fontSize: 12, fontFace: FONT, bold: true, color: COLOR.white, fill: { color: COLOR.teal }, align: "center", valign: "middle" } },
  { text: "주요 방사통(Referred Pain)", options: { fontSize: 12, fontFace: FONT, bold: true, color: COLOR.white, fill: { color: COLOR.teal }, align: "center", valign: "middle" } },
];

const trpData = [
  ["TrP1", "상부 (가장 흔함!)", "목 후외측 → 유양돌기 → 관자놀이·안와 뒤"],
  ["TrP2", "상부 후방", "TrP1과 유사, 귀 뒤 중심"],
  ["TrP3", "하부", "경추 상부·유양돌기·견봉 + 견갑상부 deep ache"],
  ["TrP4", "하부", "견갑골 내측연 작열통"],
  ["TrP5", "중부", "C7~T1 사이 천층 작열통"],
  ["TrP6", "중부 견봉 근처", "어깨 꼭지점·견봉돌기 쑤시는 통증"],
  ["TrP7", "중부 천층", "통증 아닌 소름(Gooseflesh) 자율신경 반응"],
];

const trpRows = trpData.map((row, idx) => [
  { text: row[0], options: { fontSize: 11, fontFace: FONT, bold: true, color: idx === 0 ? COLOR.softRed : COLOR.darkGray, align: "center", valign: "middle", margin: [2, 4, 2, 4], fill: { color: idx === 0 ? "FFF0F0" : COLOR.white } } },
  { text: row[1], options: { fontSize: 11, fontFace: FONT, color: COLOR.darkGray, valign: "middle", margin: [2, 4, 2, 4], fill: { color: idx === 0 ? "FFF0F0" : COLOR.white } } },
  { text: row[2], options: { fontSize: 11, fontFace: FONT, color: COLOR.darkGray, valign: "middle", margin: [2, 4, 2, 4], fill: { color: idx === 0 ? "FFF0F0" : COLOR.white } } },
]);

slide7.addTable([trpHeader, ...trpRows], {
  x: 0.3, y: 1.2, w: 9.4,
  border: { type: "solid", color: COLOR.lineGray, pt: 0.5 },
  colW: [0.8, 2.0, 6.6],
  rowH: [0.45, 0.45, 0.45, 0.45, 0.45, 0.45, 0.45, 0.45],
  autoPage: false,
});

// TrP1 강조 박스
slide7.addShape(pptx.shapes.RECTANGLE, {
  x: 0.3, y: 4.9, w: 9.4, h: 0.55,
  fill: { color: COLOR.softRed },
});

slide7.addText("TrP1은 인체 모든 Myofascial TrP 중 가장 높은 빈도 — 편두통으로 오진되기 쉬움!", {
  x: 0.3, y: 4.9, w: 9.4, h: 0.55,
  fontSize: 13, fontFace: FONT, bold: true, color: COLOR.white,
  align: "center", valign: "middle", margin: 0,
});

// ============================================================
// 슬라이드 8: TrP3 & 활성화 요인 (2컬럼)
// ============================================================
const slide8 = pptx.addSlide();

slide8.addText("TrP3 '조커(Joker)' & 활성화 요인", {
  x: 0.5, y: 0.3, w: 9.0, h: 0.7,
  fontSize: 30,
  fontFace: FONT,
  bold: true,
  color: COLOR.navy,
  align: "left",
});

slide8.addShape(pptx.shapes.RECTANGLE, {
  x: 0.5, y: 1.0, w: 1.2, h: 0.04,
  fill: { color: COLOR.teal },
});

// 좌측: TrP3 설명
slide8.addShape(pptx.shapes.RECTANGLE, {
  x: 0.5, y: 1.4, w: 4.2, h: 0.55,
  fill: { color: COLOR.teal },
});

slide8.addText("TrP3 — 숨은 원인 'Joker'", {
  x: 0.5, y: 1.4, w: 4.2, h: 0.55,
  fontSize: 16, fontFace: FONT, bold: true, color: COLOR.white,
  align: "center", valign: "middle", margin: 0,
});

slide8.addShape(pptx.shapes.RECTANGLE, {
  x: 0.5, y: 1.95, w: 4.2, h: 2.8,
  fill: { color: COLOR.lightTeal },
  line: { color: COLOR.lineGray, width: 0.5 },
});

const trp3Items = [
  "하부 승모근(Lower Trapezius)에 위치",
  "다른 TrP를 모두 제거한 후에도\n끈질기게 남아 있는 통증의 원인",
  "경추 상부·유양돌기·견봉 방사통",
  "견갑상부 깊은 통증(Deep Ache)",
  "치료에서 마지막으로 확인해야 할 포인트",
];

slide8.addText(trp3Items.map((item) => ({
  text: `\u2022  ${item}`,
  options: { fontSize: 12, fontFace: FONT, color: COLOR.darkGray, breakLine: true },
})), {
  x: 0.65, y: 2.1, w: 3.9, h: 2.5,
  valign: "top", lineSpacingMultiple: 1.5, margin: 0,
});

// 우측: 활성화 요인
slide8.addShape(pptx.shapes.RECTANGLE, {
  x: 5.3, y: 1.4, w: 4.2, h: 0.55,
  fill: { color: COLOR.navy },
});

slide8.addText("활성화 요인(Activating Factors)", {
  x: 5.3, y: 1.4, w: 4.2, h: 0.55,
  fontSize: 16, fontFace: FONT, bold: true, color: COLOR.white,
  align: "center", valign: "middle", margin: 0,
});

slide8.addShape(pptx.shapes.RECTANGLE, {
  x: 5.3, y: 1.95, w: 4.2, h: 2.8,
  fill: { color: COLOR.lightTeal },
  line: { color: COLOR.lineGray, width: 0.5 },
});

const factorItems = [
  "하지길이 불일치(LLLI)\nSmall Hemipelvis",
  "팔꿈치 지지 없이 수화기 잡기",
  "높은 키보드 / 요부 지지 부족",
  "타이트한 브래지어 끈",
  "무거운 코트",
  "편타성 손상(Whiplash)",
];

slide8.addText(factorItems.map((item) => ({
  text: `\u2022  ${item}`,
  options: { fontSize: 12, fontFace: FONT, color: COLOR.darkGray, breakLine: true },
})), {
  x: 5.45, y: 2.1, w: 3.9, h: 2.5,
  valign: "top", lineSpacingMultiple: 1.5, margin: 0,
});

// ============================================================
// 슬라이드 9: 요약
// ============================================================
const slide9 = pptx.addSlide();

slide9.addText("요약", {
  x: 0.5, y: 0.3, w: 9.0, h: 0.7,
  fontSize: 36,
  fontFace: FONT,
  bold: true,
  color: COLOR.navy,
  align: "left",
});

slide9.addShape(pptx.shapes.RECTANGLE, {
  x: 0.5, y: 1.0, w: 1.2, h: 0.04,
  fill: { color: COLOR.teal },
});

const keyPoints = [
  "승모근은 상·중·하부가 독립적으로 기능하는 다이아몬드형 근육이다",
  "상부는 거상·회전, 중부는 후인, 하부는 후인+회전 보조 — 세 가지를 반드시 구분한다",
  "신경지배: 운동은 부신경(CN XI), 감각은 C2~C4 경추신경",
  "TrP1은 인체에서 가장 흔한 Myofascial TrP이며, 편두통·긴장성 항통의 주범이다",
  "교정의 핵심은 자세 개선(팔꿈치 지지, 키보드 높이 조정)과 스트레치이다",
];

const pointTexts = keyPoints.map((point, idx) => [
  {
    text: `${idx + 1}. `,
    options: { fontSize: 16, fontFace: FONT, bold: true, color: COLOR.teal },
  },
  {
    text: point,
    options: { fontSize: 15, fontFace: FONT, color: COLOR.darkGray, breakLine: true },
  },
]).flat();

slide9.addText(pointTexts, {
  x: 0.5, y: 1.4, w: 9.0, h: 2.8,
  valign: "top",
  lineSpacingMultiple: 1.8,
  margin: 0,
});

slide9.addShape(pptx.shapes.RECTANGLE, {
  x: 0.5, y: 4.5, w: 9.0, h: 0.7,
  fill: { color: COLOR.teal },
});

slide9.addText("핵심: 승모근 TrP를 이해하면 목·어깨 통증의 원인과 해결책이 보입니다", {
  x: 0.5, y: 4.5, w: 9.0, h: 0.7,
  fontSize: 14, fontFace: FONT, bold: true, color: COLOR.white,
  align: "center", valign: "middle", margin: 0,
});

// ============================================================
// 슬라이드 10: 참고문헌
// ============================================================
const slide10 = pptx.addSlide();

slide10.addText("참고문헌(References)", {
  x: 0.5, y: 0.3, w: 9.0, h: 0.7,
  fontSize: 28,
  fontFace: FONT,
  bold: true,
  color: COLOR.navy,
  align: "left",
});

slide10.addShape(pptx.shapes.RECTANGLE, {
  x: 0.5, y: 1.0, w: 1.2, h: 0.04,
  fill: { color: COLOR.teal },
});

const references = [
  "Simons DG, Travell JG, Simons LS. Myofascial Pain and Dysfunction: The Trigger Point Manual. Vol 1. 2nd ed. Williams & Wilkins; 1999.",
  "Neumann DA. Kinesiology of the Musculoskeletal System: Foundations for Rehabilitation. 3rd ed. Elsevier; 2017.",
  "Moore KL, Dalley AF, Agur AMR. Clinically Oriented Anatomy. 8th ed. Wolters Kluwer; 2018.",
  "Image: Trapezius_Gray409.PNG, by Mikael Häggström (based on Gray's Anatomy), Wikimedia Commons, Public Domain, https://commons.wikimedia.org/wiki/File:Trapezius_Gray409.PNG",
  "Image: Trapezius_pain.svg, by InjuryMap, Wikimedia Commons, CC BY-SA 4.0, https://commons.wikimedia.org/wiki/File:Trapezius_pain.svg",
];

const refTexts = references.map((ref, idx) => ({
  text: `${idx + 1}. ${ref}`,
  options: { fontSize: 11, fontFace: FONT, color: COLOR.medGray, breakLine: true },
}));

slide10.addText(refTexts, {
  x: 0.5, y: 1.3, w: 9.0, h: 3.8,
  valign: "top",
  lineSpacingMultiple: 1.6,
  margin: 0,
});

// ============================================================
// 파일 저장
// ============================================================
pptx.writeFile({ fileName: "승모근_움직임해부학.pptx" })
  .then(() => console.log("PPTX 생성 완료: 승모근_움직임해부학.pptx"))
  .catch((err) => console.error("Error:", err));
