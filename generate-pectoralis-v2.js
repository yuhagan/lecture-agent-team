const pptxgen = require("pptxgenjs");
const pptx = new pptxgen();

pptx.layout = "LAYOUT_16x9";
pptx.author = "의학 블로그";
pptx.title = "대흉근(Pectoralis Major)의 해부학";

// === 레퍼런스 스타일 컬러 ===
const C = {
  white: "FFFFFF",
  bg: "F5F5F5",       // 라이트그레이 배경
  blue: "2B3FE0",     // 비비드 블루 (포인트 컬러 유일)
  black: "1A1A1A",    // 타이틀용 블랙
  dark: "2D2D2D",     // 본문 텍스트
  med: "7A7A7A",      // 보조/캡션 텍스트
  light: "C0C0C0",    // 라인/구분선
  ltBg: "EEEEF2",     // 표 배경 (연한 그레이)
};
const FONT = "Pretendard";

// === 공통 함수: 푸터 ===
function addFooter(slide, pageNum, totalPages) {
  // 좌측 푸터 텍스트
  slide.addText("의학 블로그", {
    x: 0.5, y: 4.95, w: 2.0, h: 0.3,
    fontSize: 8, fontFace: FONT, color: C.med,
  });
  // 우측 페이지 번호
  slide.addText(`${String(pageNum).padStart(2, "0")} / ${String(totalPages).padStart(2, "0")}`, {
    x: 7.5, y: 4.95, w: 2.0, h: 0.3,
    fontSize: 8, fontFace: FONT, color: C.med, align: "right",
  });
  // 상단 얇은 구분선
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.5, y: 4.9, w: 9.0, h: 0.005,
    fill: { color: C.light },
  });
}

// === 공통 함수: 파란 스크리블 밑줄 (hand-drawn 느낌) ===
function addScribbleLine(slide, x, y, w) {
  // 메인 라인
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: x, y: y, w: w, h: 0.04,
    fill: { color: C.blue },
    rotate: -0.5,
  });
  // 약간 오프셋된 두 번째 라인 (hand-drawn 느낌)
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: x + 0.05, y: y + 0.06, w: w * 0.7, h: 0.03,
    fill: { color: C.blue },
    rotate: 0.8,
  });
}

// === 공통 함수: 파란 데코 원 ===
function addDecoCircle(slide, x, y, size) {
  slide.addShape(pptx.shapes.OVAL, {
    x: x, y: y, w: size, h: size,
    line: { color: C.blue, width: 1.5 },
    fill: { type: "none" },
  });
}

const TOTAL = 9;

// ============================================================
// 슬라이드 1: 표지 — 에디토리얼 스타일
// ============================================================
const s1 = pptx.addSlide();
s1.background = { fill: C.bg };

// 좌상단 작은 텍스트
s1.addText("움직임해부학 시리즈", {
  x: 0.5, y: 0.4, w: 3.0, h: 0.3,
  fontSize: 9, fontFace: FONT, color: C.med,
});

s1.addText("2026", {
  x: 0.5, y: 0.8, w: 1.5, h: 0.6,
  fontSize: 28, fontFace: FONT, bold: true, color: C.black,
});

// 메인 타이틀 — 대형, 좌측 하단 배치
s1.addText("대흉근의\n해부학", {
  x: 0.5, y: 1.5, w: 5.0, h: 2.0,
  fontSize: 48, fontFace: FONT, bold: true, color: C.blue,
  lineSpacingMultiple: 1.1,
});

// 스크리블 밑줄
addScribbleLine(s1, 0.5, 3.5, 4.5);

// 우측 상단 설명 텍스트
s1.addText("대흉근(Pectoralis Major)은\n전흉벽에 위치한 부채꼴 근육으로,\n상지 내전과 내회전의 주동근이다.", {
  x: 5.5, y: 0.5, w: 3.8, h: 1.2,
  fontSize: 11, fontFace: FONT, color: C.dark,
  lineSpacingMultiple: 1.6,
});

// 우측 하단 이미지 (그레이스케일 분위기)
s1.addImage({
  path: "./gray410.png",
  x: 5.8, y: 2.0, w: 3.5, h: 2.7,
  sizing: { type: "contain", w: 3.5, h: 2.7 },
});

// 하단 작은 본문
s1.addText("흉벽 전면의 표층 근육으로, 쇄골부 · 흉늑부 · 복부의\n세 부분으로 구성된다. 의학 블로그 시리즈.", {
  x: 0.5, y: 3.8, w: 4.5, h: 0.8,
  fontSize: 9, fontFace: FONT, color: C.med,
  lineSpacingMultiple: 1.5,
});

addFooter(s1, 1, TOTAL);

// ============================================================
// 슬라이드 2: 목차 — about us 스타일 (비대칭)
// ============================================================
const s2 = pptx.addSlide();
s2.background = { fill: C.bg };

// 우측 상단 대형 타이틀
s2.addText("목차", {
  x: 5.5, y: 0.3, w: 4.0, h: 1.2,
  fontSize: 52, fontFace: FONT, bold: true, color: C.blue,
  align: "right",
});

addScribbleLine(s2, 6.5, 1.4, 3.0);

// 좌측 설명
s2.addText("이 슬라이드에서 다루는\n대흉근의 핵심 주제들을\n순서대로 정리하였습니다.", {
  x: 0.5, y: 0.5, w: 4.0, h: 1.0,
  fontSize: 12, fontFace: FONT, color: C.dark,
  lineSpacingMultiple: 1.6,
});

// 목차 항목 — 큰 번호 + 제목
const tocItems = [
  "개요 및 해부학적 위치",
  "기시(Origin)와 정지(Insertion)",
  "작용(Action)",
  "신경지배(Innervation)",
  "임상적 의의",
];

tocItems.forEach((item, idx) => {
  const y = 2.0 + idx * 0.55;

  // 큰 번호
  s2.addText(String(idx + 1).padStart(2, "0"), {
    x: 0.5, y: y, w: 0.8, h: 0.45,
    fontSize: 22, fontFace: FONT, bold: true, color: C.blue,
    valign: "middle",
  });

  // 구분 점
  s2.addShape(pptx.shapes.OVAL, {
    x: 1.4, y: y + 0.17, w: 0.08, h: 0.08,
    fill: { color: C.blue },
  });

  // 항목명
  s2.addText(item, {
    x: 1.7, y: y, w: 5.0, h: 0.45,
    fontSize: 16, fontFace: FONT, color: C.dark,
    valign: "middle",
  });

  // 우측 작은 설명
  if (idx === 0) {
    s2.addText("위치와 구조", {
      x: 7.0, y: y, w: 2.5, h: 0.45,
      fontSize: 10, fontFace: FONT, color: C.med, valign: "middle", align: "right",
    });
  }
});

addFooter(s2, 2, TOTAL);

// ============================================================
// 슬라이드 3: 개요 — project concept 스타일
// ============================================================
const s3 = pptx.addSlide();
s3.background = { fill: C.bg };

// 좌상단 작은 라벨
s3.addText("01  개요", {
  x: 0.5, y: 0.4, w: 3.0, h: 0.3,
  fontSize: 9, fontFace: FONT, color: C.med,
});

// 대형 타이틀 — 좌측
s3.addText("개요 및\n해부학적 위치", {
  x: 0.5, y: 0.8, w: 4.5, h: 1.4,
  fontSize: 38, fontFace: FONT, bold: true, color: C.blue,
  lineSpacingMultiple: 1.1,
});

addScribbleLine(s3, 0.5, 2.2, 3.8);

// 좌측 하단 본문 텍스트
s3.addText(
  "대흉근(Pectoralis Major)은 전흉벽(Anterior Thoracic Wall)의 표층에 위치한 크고 부채꼴 모양의 근육이다.\n\n" +
  "상지(Upper Limb)의 움직임, 특히 내전, 굴곡, 내회전에 핵심적인 역할을 수행한다.", {
  x: 0.5, y: 2.6, w: 4.2, h: 1.8,
  fontSize: 12, fontFace: FONT, color: C.dark,
  lineSpacingMultiple: 1.6,
});

// 우측 상단: 구성 정보 블록
s3.addText("구성(Components)", {
  x: 5.5, y: 0.5, w: 4.0, h: 0.4,
  fontSize: 11, fontFace: FONT, bold: true, color: C.black,
});

const components = [
  ["쇄골부(Clavicular Head)", "쇄골 내측 1/2의 전면에서 기시"],
  ["흉늑부(Sternocostal Head)", "흉골 전면 및 제1~6늑연골에서 기시"],
  ["복부(Abdominal Part)", "외복사근건막에서 기시"],
];

components.forEach((comp, idx) => {
  const y = 1.1 + idx * 0.9;

  s3.addShape(pptx.shapes.OVAL, {
    x: 5.5, y: y + 0.05, w: 0.08, h: 0.08,
    fill: { color: C.blue },
  });

  s3.addText(comp[0], {
    x: 5.8, y: y - 0.05, w: 3.7, h: 0.35,
    fontSize: 12, fontFace: FONT, bold: true, color: C.dark,
  });

  s3.addText(comp[1], {
    x: 5.8, y: y + 0.3, w: 3.7, h: 0.35,
    fontSize: 10, fontFace: FONT, color: C.med,
  });
});

// 우측 하단: 이미지
s3.addImage({
  path: "./gray410.png",
  x: 5.8, y: 3.3, w: 3.2, h: 1.5,
  sizing: { type: "contain", w: 3.2, h: 1.5 },
});

s3.addText("Gray's Anatomy, Plate 410  |  Public Domain", {
  x: 5.8, y: 4.7, w: 3.2, h: 0.2,
  fontSize: 7, fontFace: FONT, italic: true, color: C.light, align: "right",
});

addFooter(s3, 3, TOTAL);

// ============================================================
// 슬라이드 4: 기시와 정지 — project name 스타일 (표 중심)
// ============================================================
const s4 = pptx.addSlide();
s4.background = { fill: C.bg };

s4.addText("02  기시와 정지", {
  x: 0.5, y: 0.4, w: 3.0, h: 0.3,
  fontSize: 9, fontFace: FONT, color: C.med,
});

// 우측 대형 타이틀
s4.addText("기시와\n정지", {
  x: 6.0, y: 0.3, w: 3.5, h: 1.4,
  fontSize: 44, fontFace: FONT, bold: true, color: C.blue,
  align: "right", lineSpacingMultiple: 1.0,
});

addScribbleLine(s4, 7.0, 1.65, 2.5);

// 좌측 상단 부연 설명
s4.addText("세 부분의 근섬유가 정지 부위에서 꼬이며 합쳐지는 독특한 구조를 가진다. 복부 섬유가 가장 위쪽, 쇄골부 섬유가 가장 아래쪽에 부착된다.", {
  x: 0.5, y: 0.5, w: 5.0, h: 0.9,
  fontSize: 11, fontFace: FONT, color: C.dark, lineSpacingMultiple: 1.5,
});

// 기시 테이블
const mkCell = (txt, isHead, isLabel) => ({
  text: txt,
  options: {
    fontSize: isHead ? 11 : 11,
    fontFace: FONT,
    bold: isHead || isLabel,
    color: isHead ? C.white : isLabel ? C.blue : C.dark,
    fill: { color: isHead ? C.blue : C.white },
    valign: "middle",
    margin: [5, 8, 5, 8],
    align: isHead ? "center" : "left",
  },
});

const t4Rows = [
  [mkCell("구분", true), mkCell("부위", true), mkCell("부착 부위", true)],
  [mkCell("기시\n(Origin)", false, true), mkCell("쇄골부(Clavicular Head)", false), mkCell("쇄골(Clavicle) 내측 1/2 전면", false)],
  [mkCell("", false), mkCell("흉늑부(Sternocostal Head)", false), mkCell("흉골(Sternum) 전면, 제1~6늑연골(Costal Cartilage)", false)],
  [mkCell("", false), mkCell("복부(Abdominal Part)", false), mkCell("외복사근건막(External Oblique Aponeurosis)", false)],
  [mkCell("정지\n(Insertion)", false, true), mkCell("공통", false), mkCell("상완골(Humerus) 대결절릉(Crest of Greater Tubercle)", false)],
];

s4.addTable(t4Rows, {
  x: 0.5, y: 1.8, w: 9.0,
  border: { type: "solid", color: C.light, pt: 0.5 },
  colW: [1.3, 2.7, 5.0],
  rowH: [0.4, 0.55, 0.55, 0.55, 0.55],
  autoPage: false,
});

// 하단 노트
addDecoCircle(s4, 0.5, 4.45, 0.25);
s4.addText("정지 구조가 독특 — 섬유가 180° 회전하며 합류", {
  x: 0.9, y: 4.45, w: 5.0, h: 0.3,
  fontSize: 10, fontFace: FONT, italic: true, color: C.med,
  valign: "middle",
});

addFooter(s4, 4, TOTAL);

// ============================================================
// 슬라이드 5: 작용 — project stages 스타일 (3컬럼)
// ============================================================
const s5 = pptx.addSlide();
s5.background = { fill: C.bg };

s5.addText("03  작용", {
  x: 0.5, y: 0.4, w: 3.0, h: 0.3,
  fontSize: 9, fontFace: FONT, color: C.med,
});

// 좌측 대형 타이틀
s5.addText("작용", {
  x: 0.5, y: 0.7, w: 3.0, h: 1.0,
  fontSize: 52, fontFace: FONT, bold: true, color: C.blue,
});

addScribbleLine(s5, 0.5, 1.65, 2.0);

// 좌측 본문 설명
s5.addText("대흉근의 세 부분은 각기 다른\n작용을 수행하며, 전체적으로\n상지 내전과 내회전의 주동근이다.", {
  x: 0.5, y: 1.9, w: 3.0, h: 1.0,
  fontSize: 11, fontFace: FONT, color: C.dark, lineSpacingMultiple: 1.6,
});

// 3컬럼: 우측에 배치
const actionData = [
  {
    head: "쇄골부",
    sub: "Clavicular Head",
    items: ["어깨관절 굴곡(Flexion)", "수평내전(Horizontal Adduction)", "내회전(Internal Rotation)"],
  },
  {
    head: "흉늑부",
    sub: "Sternocostal Head",
    items: ["어깨관절 신전(Extension)\n— 굴곡 위치에서", "내전(Adduction)", "내회전(Internal Rotation)"],
  },
  {
    head: "전체",
    sub: "Combined",
    items: ["강력한 내전(Adduction)", "강력한 내회전(Int. Rotation)", "보조 호흡근\n— 상지 고정 시"],
  },
];

actionData.forEach((col, idx) => {
  const x = 3.8 + idx * 2.1;
  const colW2 = 1.9;

  // 컬럼 헤더
  s5.addText(col.head, {
    x: x, y: 0.5, w: colW2, h: 0.45,
    fontSize: 18, fontFace: FONT, bold: true, color: C.blue,
  });

  s5.addText(col.sub, {
    x: x, y: 0.95, w: colW2, h: 0.25,
    fontSize: 9, fontFace: FONT, color: C.med,
  });

  // 구분선
  s5.addShape(pptx.shapes.RECTANGLE, {
    x: x, y: 1.25, w: colW2, h: 0.015,
    fill: { color: C.blue },
  });

  // 항목
  col.items.forEach((item, iIdx) => {
    const iy = 1.5 + iIdx * 1.05;

    // 번호
    s5.addText(String(iIdx + 1), {
      x: x, y: iy, w: 0.3, h: 0.3,
      fontSize: 14, fontFace: FONT, bold: true, color: C.blue,
    });

    // 텍스트
    s5.addText(item, {
      x: x + 0.35, y: iy, w: colW2 - 0.4, h: 0.85,
      fontSize: 10, fontFace: FONT, color: C.dark,
      lineSpacingMultiple: 1.4, valign: "top",
    });
  });
});

addFooter(s5, 5, TOTAL);

// ============================================================
// 슬라이드 6: 신경지배 — timeline 스타일 (수평 흐름도)
// ============================================================
const s6 = pptx.addSlide();
s6.background = { fill: C.bg };

s6.addText("04  신경지배", {
  x: 0.5, y: 0.4, w: 3.0, h: 0.3,
  fontSize: 9, fontFace: FONT, color: C.med,
});

// 우측 대형 타이틀
s6.addText("신경지배", {
  x: 5.5, y: 0.3, w: 4.0, h: 1.0,
  fontSize: 48, fontFace: FONT, bold: true, color: C.blue,
  align: "right",
});

addScribbleLine(s6, 6.5, 1.25, 3.0);

// 좌측 설명
s6.addText("대흉근은 상완신경총(Brachial Plexus)의 외측속(Lateral Cord)과 내측속(Medial Cord)에서 각각 기원하는 두 개의 흉근신경에 의해 이중 지배를 받는다.", {
  x: 0.5, y: 0.5, w: 4.5, h: 1.0,
  fontSize: 11, fontFace: FONT, color: C.dark, lineSpacingMultiple: 1.6,
});

// 수평 타임라인 흐름도
const timelineY = 2.2;
const steps = [
  { label: "척수\nSpinal Cord", sub: "C5 — T1" },
  { label: "전지\nVentral Rami", sub: "" },
  { label: "상완신경총\nBrachial Plexus", sub: "" },
];

steps.forEach((step, idx) => {
  const x = 0.5 + idx * 3.0;

  // 번호 원
  s6.addShape(pptx.shapes.OVAL, {
    x: x, y: timelineY, w: 0.35, h: 0.35,
    fill: { color: C.blue },
  });
  s6.addText(String(idx + 1), {
    x: x, y: timelineY, w: 0.35, h: 0.35,
    fontSize: 12, fontFace: FONT, bold: true, color: C.white,
    align: "center", valign: "middle", margin: 0,
  });

  // 연결선
  if (idx < steps.length - 1) {
    s6.addShape(pptx.shapes.RECTANGLE, {
      x: x + 0.45, y: timelineY + 0.16, w: 2.45, h: 0.02,
      fill: { color: C.blue },
    });
  }

  // 레이블
  s6.addText(step.label, {
    x: x + 0.5, y: timelineY - 0.05, w: 2.2, h: 0.5,
    fontSize: 12, fontFace: FONT, bold: true, color: C.dark,
    valign: "middle",
  });

  if (step.sub) {
    s6.addText(step.sub, {
      x: x + 0.5, y: timelineY + 0.4, w: 2.2, h: 0.3,
      fontSize: 10, fontFace: FONT, color: C.med,
    });
  }
});

// 분기 화살표: 두 신경
const branchY = 3.2;

// 외측 흉근신경
s6.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: branchY, w: 4.2, h: 0.75,
  fill: { color: C.white },
  line: { color: C.blue, width: 1.5 },
  rectRadius: 0.03,
});

s6.addText([
  { text: "외측 흉근신경(Lateral Pectoral N.)", options: { fontSize: 13, fontFace: FONT, bold: true, color: C.blue, breakLine: true } },
  { text: "외측속(Lateral Cord) 기원  |  C5, C6, C7  |  쇄골부 지배", options: { fontSize: 10, fontFace: FONT, color: C.med } },
], {
  x: 0.7, y: branchY, w: 3.8, h: 0.75,
  valign: "middle", margin: 0,
});

// 내측 흉근신경
s6.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 5.3, y: branchY, w: 4.2, h: 0.75,
  fill: { color: C.white },
  line: { color: C.blue, width: 1.5 },
  rectRadius: 0.03,
});

s6.addText([
  { text: "내측 흉근신경(Medial Pectoral N.)", options: { fontSize: 13, fontFace: FONT, bold: true, color: C.blue, breakLine: true } },
  { text: "내측속(Medial Cord) 기원  |  C8, T1  |  흉늑부·복부 지배", options: { fontSize: 10, fontFace: FONT, color: C.med } },
], {
  x: 5.5, y: branchY, w: 3.8, h: 0.75,
  valign: "middle", margin: 0,
});

// 하단 합류 → 대흉근
s6.addShape(pptx.shapes.RECTANGLE, {
  x: 2.6, y: branchY + 0.75, w: 0.01, h: 0.3,
  fill: { color: C.blue },
});
s6.addShape(pptx.shapes.RECTANGLE, {
  x: 7.4, y: branchY + 0.75, w: 0.01, h: 0.3,
  fill: { color: C.blue },
});

// 합류 가로선
s6.addShape(pptx.shapes.RECTANGLE, {
  x: 2.6, y: branchY + 1.05, w: 4.8, h: 0.015,
  fill: { color: C.blue },
});
// 중앙 세로선
s6.addShape(pptx.shapes.RECTANGLE, {
  x: 5.0, y: branchY + 1.05, w: 0.01, h: 0.25,
  fill: { color: C.blue },
});

s6.addText("대흉근(Pectoralis Major)", {
  x: 3.0, y: branchY + 1.3, w: 4.0, h: 0.45,
  fontSize: 16, fontFace: FONT, bold: true, color: C.blue,
  align: "center", valign: "middle",
});

addFooter(s6, 6, TOTAL);

// ============================================================
// 슬라이드 7: 임상적 의의 — project stages 스타일 (이미지 + 카드)
// ============================================================
const s7 = pptx.addSlide();
s7.background = { fill: C.bg };

s7.addText("05  임상적 의의", {
  x: 0.5, y: 0.4, w: 3.0, h: 0.3,
  fontSize: 9, fontFace: FONT, color: C.med,
});

// 대형 타이틀
s7.addText("임상적\n의의", {
  x: 0.5, y: 0.7, w: 4.0, h: 1.4,
  fontSize: 44, fontFace: FONT, bold: true, color: C.blue,
  lineSpacingMultiple: 1.0,
});

addScribbleLine(s7, 0.5, 2.1, 3.0);

// 좌측 하단 이미지
s7.addImage({
  path: "./pectoralis_major.png",
  x: 0.5, y: 2.5, w: 3.0, h: 2.2,
  sizing: { type: "contain", w: 3.0, h: 2.2 },
});

s7.addText("Anatomography  |  CC BY-SA 3.0", {
  x: 0.5, y: 4.65, w: 3.0, h: 0.2,
  fontSize: 7, fontFace: FONT, italic: true, color: C.light,
});

// 우측: 3개 임상 카드
const clinicals = [
  {
    title: "대흉근 파열(Rupture)",
    body: "벤치프레스 등 과부하 시 건-골 접합부에서 호발. 전면 액와주름 소실이 특징적 소견.",
  },
  {
    title: "Poland 증후군(Poland Syndrome)",
    body: "대흉근의 선천적 결손. 흉늑부 부재가 가장 흔하며, 동측 상지 기형 동반 가능.",
  },
  {
    title: "단축/긴장(Tightness)",
    body: "둥근어깨(Rounded Shoulder) 자세 유발. 상위교차증후군의 핵심 요소로 스트레칭 대상.",
  },
];

clinicals.forEach((item, idx) => {
  const y = 0.7 + idx * 1.45;
  const x = 4.2;
  const cardW = 5.3;

  // 카드 배경
  s7.addShape(pptx.shapes.RECTANGLE, {
    x: x, y: y, w: cardW, h: 1.25,
    fill: { color: C.white },
    line: { color: C.light, width: 0.5 },
  });

  // 좌측 블루 바
  s7.addShape(pptx.shapes.RECTANGLE, {
    x: x, y: y, w: 0.05, h: 1.25,
    fill: { color: C.blue },
  });

  // 번호
  s7.addText(`Stage ${idx + 1}.`, {
    x: x + 0.2, y: y + 0.1, w: 1.0, h: 0.25,
    fontSize: 9, fontFace: FONT, bold: true, color: C.blue,
  });

  // 타이틀
  s7.addText(item.title, {
    x: x + 0.2, y: y + 0.3, w: cardW - 0.5, h: 0.35,
    fontSize: 13, fontFace: FONT, bold: true, color: C.black,
  });

  // 본문
  s7.addText(item.body, {
    x: x + 0.2, y: y + 0.65, w: cardW - 0.5, h: 0.5,
    fontSize: 10, fontFace: FONT, color: C.med,
    lineSpacingMultiple: 1.4,
  });
});

addFooter(s7, 7, TOTAL);

// ============================================================
// 슬라이드 8: 요약 — outcomes 스타일
// ============================================================
const s8 = pptx.addSlide();
s8.background = { fill: C.bg };

// 하단 좌측에 대형 타이틀
s8.addText("요약", {
  x: 0.5, y: 3.6, w: 3.0, h: 1.0,
  fontSize: 56, fontFace: FONT, bold: true, color: C.blue,
});

addScribbleLine(s8, 0.5, 4.55, 2.2);

// 상단 상세 내용: 5개 포인트를 2열로
const summaryPoints = [
  { num: "01", text: "대흉근은 쇄골부, 흉늑부, 복부로\n구성된 대근육이다" },
  { num: "02", text: "상완골 대결절릉에 부착하며,\n섬유가 꼬이며 합쳐진다" },
  { num: "03", text: "어깨관절의 내전, 굴곡,\n내회전을 담당한다" },
  { num: "04", text: "외측/내측 흉근신경의\n이중 신경지배를 받는다" },
  { num: "05", text: "파열, Poland 증후군,\n상위교차증후군과 관련이 깊다" },
];

summaryPoints.forEach((pt, idx) => {
  const col = idx < 3 ? 0 : 1;
  const row = idx < 3 ? idx : idx - 3;
  const x = 0.5 + col * 4.8;
  const y = 0.5 + row * 1.1;

  s8.addText(pt.num, {
    x: x, y: y, w: 0.6, h: 0.4,
    fontSize: 20, fontFace: FONT, bold: true, color: C.blue,
  });

  s8.addText(pt.text, {
    x: x + 0.65, y: y, w: 3.8, h: 0.8,
    fontSize: 11, fontFace: FONT, color: C.dark,
    lineSpacingMultiple: 1.5, valign: "top",
  });
});

// 우측 하단 핵심 강조
s8.addShape(pptx.shapes.RECTANGLE, {
  x: 5.3, y: 3.7, w: 4.2, h: 0.8,
  fill: { color: C.white },
  line: { color: C.blue, width: 1.5 },
});

s8.addText("핵심: 대흉근은 상지 내전과 내회전의\n주동근이며, 자세 교정과 재활의 핵심 근육이다", {
  x: 5.5, y: 3.7, w: 3.8, h: 0.8,
  fontSize: 11, fontFace: FONT, bold: true, color: C.blue,
  valign: "middle", lineSpacingMultiple: 1.4,
});

addFooter(s8, 8, TOTAL);

// ============================================================
// 슬라이드 9: 참고문헌 — thank you 스타일
// ============================================================
const s9 = pptx.addSlide();
s9.background = { fill: C.bg };

// 좌측 연도 + 감사
s9.addText("2026", {
  x: 0.5, y: 0.5, w: 2.0, h: 0.6,
  fontSize: 28, fontFace: FONT, bold: true, color: C.black,
});

s9.addText("감사합니다", {
  x: 0.5, y: 1.0, w: 4.0, h: 1.0,
  fontSize: 44, fontFace: FONT, bold: true, color: C.blue,
});

addScribbleLine(s9, 0.5, 1.95, 3.8);

// 좌측 하단 이미지
s9.addImage({
  path: "./gray410.png",
  x: 0.5, y: 2.5, w: 2.5, h: 2.2,
  sizing: { type: "contain", w: 2.5, h: 2.2 },
});

// 우측 참고문헌 리스트
s9.addText("참고문헌(References)", {
  x: 5.0, y: 0.5, w: 4.5, h: 0.4,
  fontSize: 13, fontFace: FONT, bold: true, color: C.black,
});

s9.addShape(pptx.shapes.RECTANGLE, {
  x: 5.0, y: 0.9, w: 4.5, h: 0.01,
  fill: { color: C.light },
});

const refs = [
  "Neumann DA. Kinesiology of the Musculoskeletal System: Foundations for Rehabilitation. 3rd ed. Elsevier; 2017.",
  "Moore KL, Dalley AF, Agur AMR. Clinically Oriented Anatomy. 8th ed. Wolters Kluwer; 2018.",
  "Standring S. Gray's Anatomy: The Anatomical Basis of Clinical Practice. 42nd ed. Elsevier; 2020.",
  "Image: Gray410.png, Wikimedia Commons, Public Domain.",
  "Image: Pectoralis_major.png, by Anatomography (BodyParts3D), Wikimedia Commons, CC BY-SA 3.0.",
];

refs.forEach((ref, idx) => {
  const y = 1.1 + idx * 0.7;

  s9.addText(String(idx + 1), {
    x: 5.0, y: y, w: 0.3, h: 0.25,
    fontSize: 11, fontFace: FONT, bold: true, color: C.blue,
  });

  s9.addText(ref, {
    x: 5.35, y: y, w: 4.15, h: 0.6,
    fontSize: 9, fontFace: FONT, color: C.med,
    lineSpacingMultiple: 1.4, valign: "top",
  });
});

addFooter(s9, 9, TOTAL);

// === 저장 ===
pptx.writeFile({ fileName: "대흉근_해부학_v2.pptx" })
  .then(() => console.log("PPTX 생성 완료: 대흉근_해부학_v2.pptx"))
  .catch((err) => console.error("Error:", err));
