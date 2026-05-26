const pptxgen = require("pptxgenjs");
const path = require("path");

const pptx = new pptxgen();
pptx.layout = "LAYOUT_16x9";
pptx.author = "메디스트림 실손 강의";
pptx.title = "한번에 끝내는 실손 강의";

// === 컬러 (보험/비즈니스 톤, 의학 아님) ===
const C = {
  white: "FFFFFF",
  bg: "F8FAFC",
  primary: "1B3A5C",    // 네이비 (신뢰)
  secondary: "0D9488",  // 틸 (성장/활용)
  accent: "3B82F6",     // 블루
  dark: "1E293B",
  med: "64748B",
  light: "E2E8F0",
  success: "059669",
  warn: "DC2626",
  ltBg: "F1F5F9",
};
const FONT = "Pretendard";

function addFooter(slide, pageNum, totalPages) {
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.5, y: 4.92, w: 9.0, h: 0.008,
    fill: { color: C.light },
  });
  slide.addText("한번에 끝내는 실손 강의", {
    x: 0.5, y: 4.95, w: 3.0, h: 0.25,
    fontSize: 8, fontFace: FONT, color: C.med,
  });
  slide.addText(`${pageNum} / ${totalPages}`, {
    x: 8.2, y: 4.95, w: 1.2, h: 0.25,
    fontSize: 8, fontFace: FONT, color: C.med, align: "right",
  });
}

const TOTAL = 26;

// ========== 1. 표지 ==========
const s1 = pptx.addSlide();
s1.background = { fill: C.primary };
s1.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 0.12, h: 5.25,
  fill: { color: C.secondary },
});
s1.addText("5,000건 보험 데이터 분석", {
  x: 0.6, y: 0.8, w: 8.0, h: 0.5,
  fontSize: 18, fontFace: FONT, color: C.secondary,
});
s1.addText("한번에 끝내는\n실손 강의", {
  x: 0.6, y: 1.6, w: 7.0, h: 1.8,
  fontSize: 44, fontFace: FONT, bold: true, color: C.white,
  lineSpacingMultiple: 1.15,
});
s1.addText("메디스트림 실손 강의", {
  x: 0.6, y: 4.6, w: 4.0, h: 0.35,
  fontSize: 12, fontFace: FONT, color: "FFFFFF",
  valign: "bottom", opacity: 0.85,
});
addFooter(s1, 1, TOTAL);

// ========== 2. 강사 소개 ==========
const s2 = pptx.addSlide();
s2.background = { fill: C.bg };
s2.addText("강사 소개", {
  x: 0.5, y: 0.35, w: 4.0, h: 0.6,
  fontSize: 28, fontFace: FONT, bold: true, color: C.primary,
});
s2.addShape(pptx.shapes.RECTANGLE, {
  x: 0.5, y: 0.9, w: 0.08, h: 0.35,
  fill: { color: C.secondary },
});
const intro = [
  "김종연 · 메디스트림 소프트웨어 서비스 팀장, 생명/손해보험 설계사, 존스홉킨스 의대 포닥, 한방재활의학과 전문의",
  "정희범 · 메디스트림 설립, 경영세션 4년 운영, 경영 지원 한의원 280개소",
  "실비왕 원장님 · 실손보험 활용 모범 한의원, 실손 활용 후 월 매출 1천만원 이상 상승",
];
intro.forEach((t, i) => {
  s2.addText(t, {
    x: 0.7, y: 1.0 + i * 0.95, w: 8.6, h: 0.85,
    fontSize: 11, fontFace: FONT, color: C.dark,
    lineSpacingMultiple: 1.4,
  });
});
addFooter(s2, 2, TOTAL);

// ========== 3. 실손으로 매출 천만원 가능 (인포그래픽) ==========
const s3 = pptx.addSlide();
s3.background = { fill: C.bg };
s3.addText("실손으로 매출 천만원\n가능합니다.", {
  x: 0.5, y: 0.3, w: 5.0, h: 1.0,
  fontSize: 24, fontFace: FONT, bold: true, color: C.primary,
  lineSpacingMultiple: 1.2,
});
s3.addShape(pptx.shapes.RECTANGLE, {
  x: 0.5, y: 1.25, w: 0.06, h: 0.5,
  fill: { color: C.secondary },
});
s3.addText("환자 부담은 그대로 · 한의원 추가 매출 37,730원", {
  x: 0.65, y: 1.3, w: 5.0, h: 0.25,
  fontSize: 11, fontFace: FONT, color: C.dark,
});
s3.addText("원장님 1분이 월 800명 감당 시, 1/3만 추나 전환해도 달성 가능", {
  x: 0.65, y: 1.6, w: 5.5, h: 0.35,
  fontSize: 10, fontFace: FONT, color: C.med,
});
// 인포그래픽: 진료비 구성 카드 4개
const cards = [
  { label: "본부금", val: "35,360원", x: 0.6 },
  { label: "총진료비", val: "73,090원", x: 2.5 },
  { label: "전침/IR/유관법", val: "일반치료", x: 4.4 },
  { label: "복잡추나 80%", val: "실손 청구", x: 6.3 },
];
cards.forEach((c, i) => {
  s3.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: c.x, y: 2.2, w: 1.65, h: 1.0,
    fill: { color: i === 3 ? C.secondary : C.ltBg },
    line: { color: C.light, width: 0.5 },
    rectRadius: 0.08,
  });
  s3.addText(c.val, {
    x: c.x + 0.1, y: 2.35, w: 1.45, h: 0.4,
    fontSize: 12, fontFace: FONT, bold: true, color: i === 3 ? C.white : C.primary,
    align: "center",
  });
  s3.addText(c.label, {
    x: c.x + 0.1, y: 2.8, w: 1.45, h: 0.3,
    fontSize: 9, fontFace: FONT, color: i === 3 ? C.white : C.med,
    align: "center",
  });
});
addFooter(s3, 3, TOTAL);

// ========== 4. 연령대별 실손 가입율 (인포그래픽 막대) ==========
const s4 = pptx.addSlide();
s4.background = { fill: C.bg };
s4.addText("연령대별 실손 가입율", {
  x: 0.5, y: 0.3, w: 5.0, h: 0.5,
  fontSize: 22, fontFace: FONT, bold: true, color: C.primary,
});
s4.addText("실손, 거의 다 있습니다.", {
  x: 0.5, y: 0.75, w: 5.0, h: 0.3,
  fontSize: 12, fontFace: FONT, color: C.med,
});
const ages = [
  { label: "20~29세", pct: 80.8 },
  { label: "30~39세", pct: 84.0 },
  { label: "40~49세", pct: 82.3 },
  { label: "50~59세", pct: 77.9 },
  { label: "60~69세", pct: 66.3 },
  { label: "70~79세", pct: 26.5 },
  { label: "80세 이상", pct: 1.1 },
];
const barStartX = 3.2;
const barW = 4.5;
ages.forEach((a, i) => {
  const y = 1.35 + i * 0.52;
  s4.addText(a.label, {
    x: 0.5, y: y - 0.05, w: 1.4, h: 0.3,
    fontSize: 10, fontFace: FONT, color: C.dark,
  });
  s4.addShape(pptx.shapes.RECTANGLE, {
    x: barStartX, y: y, w: (a.pct / 100) * barW, h: 0.38,
    fill: { color: a.pct >= 70 ? C.secondary : a.pct >= 50 ? C.accent : C.med },
  });
  s4.addText(`${a.pct}%`, {
    x: barStartX + (a.pct / 100) * barW + 0.08, y: y - 0.02, w: 0.6, h: 0.35,
    fontSize: 10, fontFace: FONT, bold: true, color: C.dark,
  });
});
s4.addText("출처: 김경선, 실손의료보험 현황과 과제. 보험연구원(2023)", {
  x: 0.5, y: 5.0, w: 6.0, h: 0.25,
  fontSize: 8, fontFace: FONT, color: C.med,
});
addFooter(s4, 4, TOTAL);

// ========== 5. 가격 저항·객단가·재진 (3기둥 인포그래픽) ==========
const s5 = pptx.addSlide();
s5.background = { fill: C.bg };
s5.addText("객단가, 그 이상의 가치", {
  x: 0.5, y: 0.35, w: 5.0, h: 0.5,
  fontSize: 22, fontFace: FONT, bold: true, color: C.primary,
});
s5.addText("가격 저항을 낮추고 · 객단가를 높이고 · 재진율을 높입니다.", {
  x: 0.5, y: 0.85, w: 7.0, h: 0.3,
  fontSize: 12, fontFace: FONT, color: C.dark,
});
// 3개 기둥
const pillars = [
  { title: "낮은 가격 저항", sub: "20~30% 이벤트만 해도 구매율 변화" },
  { title: "재진 유도", sub: "자주 오시니 추가 비급여 매출 기회" },
  { title: "비급여 매출", sub: "추나 실손은 80% 할인 효과" },
];
pillars.forEach((p, i) => {
  const x = 0.7 + i * 3.1;
  s5.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: 1.5, w: 2.7, h: 2.2,
    fill: { color: C.white },
    line: { color: C.secondary, width: 1 },
    rectRadius: 0.1,
  });
  s5.addShape(pptx.shapes.OVAL, {
    x: x + 0.9, y: 1.7, w: 0.9, h: 0.9,
    fill: { color: C.secondary },
  });
  s5.addText(String(i + 1), {
    x: x + 0.9, y: 1.85, w: 0.9, h: 0.6,
    fontSize: 28, fontFace: FONT, bold: true, color: C.white,
    align: "center", valign: "middle",
  });
  s5.addText(p.title, {
    x: x + 0.15, y: 2.75, w: 2.4, h: 0.4,
    fontSize: 14, fontFace: FONT, bold: true, color: C.primary,
    align: "center",
  });
  s5.addText(p.sub, {
    x: x + 0.15, y: 3.2, w: 2.4, h: 0.45,
    fontSize: 9, fontFace: FONT, color: C.med,
    align: "center", wrap: true,
  });
});
addFooter(s5, 5, TOTAL);

// ========== 6. 왜 원장님이 알아야 하나 ==========
const s6 = pptx.addSlide();
s6.background = { fill: C.bg };
s6.addText("요즘 실비 있는지 다 알던데요?", {
  x: 0.5, y: 0.35, w: 6.0, h: 0.45,
  fontSize: 20, fontFace: FONT, bold: true, color: C.primary,
});
s6.addText("환자도 사실 정확히 모릅니다. 실비는 간단해 보여도 예외가 많습니다.", {
  x: 0.5, y: 0.85, w: 8.0, h: 0.35,
  fontSize: 11, fontFace: FONT, color: C.dark,
});
const bullets6 = [
  "환자가 설계사 말은 더 믿습니다 · 설계사는 우리 편이 아닙니다.",
  "원장님께서 확실히 아셔야 상담이 가능합니다.",
  "설계사를 이길 설득력이 생깁니다.",
];
bullets6.forEach((b, i) => {
  s6.addShape(pptx.shapes.OVAL, {
    x: 0.55, y: 1.35 + i * 0.5, w: 0.12, h: 0.12,
    fill: { color: C.secondary },
  });
  s6.addText(b, {
    x: 0.75, y: 1.32 + i * 0.5, w: 8.0, h: 0.45,
    fontSize: 11, fontFace: FONT, color: C.dark,
    lineSpacingMultiple: 1.35,
  });
});
addFooter(s6, 6, TOTAL);

// ========== 7. 장모님 사례 ==========
const s7 = pptx.addSlide();
s7.background = { fill: C.bg };
s7.addText("장모님 사례", {
  x: 0.5, y: 0.35, w: 3.0, h: 0.5,
  fontSize: 22, fontFace: FONT, bold: true, color: C.primary,
});
s7.addText("실손이 있다고 알고 있었고, 청구했는데 지급 거절 문자를 받음.", {
  x: 0.5, y: 0.9, w: 8.5, h: 0.3,
  fontSize: 11, fontFace: FONT, color: C.dark,
});
s7.addText("알고 보니 한의원에선 안 되는 특약으로 심사한 것이었습니다. 조회해 보니 한의원 상해 치료 비급여까지 보장하는 특약이 있는 분이었습니다. 일반상해의료비 특약을 적용하여 보험금을 돌려받으셨습니다.", {
  x: 0.5, y: 1.3, w: 8.5, h: 1.2,
  fontSize: 11, fontFace: FONT, color: C.dark,
  lineSpacingMultiple: 1.5,
});
s7.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 2.6, w: 8.5, h: 1.0,
  fill: { color: C.ltBg },
  line: { color: C.secondary, width: 0.5 },
  rectRadius: 0.06,
});
s7.addText("확인해 보지 않았다면, 평생 한의원 보상 못 받는 줄 아셨을 것입니다. 전문가의 관리가 필요한 이유입니다.", {
  x: 0.6, y: 2.75, w: 8.3, h: 0.7,
  fontSize: 11, fontFace: FONT, bold: true, color: C.primary,
  align: "center", valign: "middle",
});
addFooter(s7, 7, TOTAL);

// ========== 8. 본인 보험을 아는가 ==========
const s8 = pptx.addSlide();
s8.background = { fill: C.bg };
s8.addText("원장님께서는 본인 보험에 대해 얼마나 알고 계신가요?", {
  x: 0.5, y: 0.4, w: 8.0, h: 0.6,
  fontSize: 20, fontFace: FONT, bold: true, color: C.primary,
});
s8.addText("보험을 모르니 찾는 전문가가 주변 보험설계사, 보험사 콜센터. 설계사는 보험 파는 사람이라 팔 상품만 공부합니다. 우리 스스로 잘 알고 있어야 대응할 수 있습니다.", {
  x: 0.5, y: 1.1, w: 8.5, h: 0.9,
  fontSize: 11, fontFace: FONT, color: C.dark,
  lineSpacingMultiple: 1.5,
});
s8.addText("\"한의원은 안 되는 보험이래요\" → 보험 담당자도 한의원 치료(보약? 비급여?)를 잘 모릅니다. 청구했는데 못 받은 경험, 실수였는지도 모릅니다.", {
  x: 0.5, y: 2.15, w: 8.5, h: 0.85,
  fontSize: 10, fontFace: FONT, color: C.med,
  lineSpacingMultiple: 1.45,
});
addFooter(s8, 8, TOTAL);

// ========== 9. 2009년 8월, 가입일만으로는 부족 ==========
const s9 = pptx.addSlide();
s9.background = { fill: C.bg };
s9.addText("2009년 8월을 1세대 기준으로 알고 계십니다.", {
  x: 0.5, y: 0.35, w: 7.0, h: 0.4,
  fontSize: 18, fontFace: FONT, bold: true, color: C.primary,
});
s9.addText("이 시기 보험은 급여/비급여 모두 되는 것처럼 알려져 있으나, 생명보험사에서는 급여만 보상하는 상품을 팔았습니다. 가입일만 봐선 확실히 알 수 없습니다.", {
  x: 0.5, y: 0.85, w: 8.5, h: 0.6,
  fontSize: 11, fontFace: FONT, color: C.dark,
});
s9.addText("예: 2009년 8월 이전에도 급여만 되는 실손 존재 (입원 80% 보상, 통원 5,000원 공제 후 80% 보상)", {
  x: 0.5, y: 1.55, w: 8.5, h: 0.4,
  fontSize: 10, fontFace: FONT, color: C.med,
});
s9.addTable([
  ["구분", "질병통원", "상해통원", "비고"],
  ["한화손해(1세대)", "X", "O", "질병통원 한의과 제외"],
  ["KB손해(3세대)", "X", "O", "상해 한의과 급여만 가능"],
], {
  x: 0.5, y: 2.2, w: 8.5, colW: [2, 2, 2, 2.5],
  fontSize: 10, fontFace: FONT,
  fill: { color: C.ltBg },
  align: "center",
  valign: "middle",
  border: { type: "solid", pt: 0.5, color: C.light },
});
s9.addText("상품명만으로는 알 수 없습니다. 주계약만 확실하고, 딸린 실손 특약 가입 여부는 전혀 알 수 없으며, 특약만 121개인 상품도 있습니다.", {
  x: 0.5, y: 3.5, w: 8.5, h: 0.7,
  fontSize: 10, fontFace: FONT, color: C.dark,
  lineSpacingMultiple: 1.4,
});
addFooter(s9, 9, TOTAL);

// ========== 10. 실비 진료 거부감 / 도수 1조원 ==========
const s10 = pptx.addSlide();
s10.background = { fill: C.bg };
s10.addText("실비 진료에 거부감이 드신다면", {
  x: 0.5, y: 0.35, w: 5.0, h: 0.45,
  fontSize: 20, fontFace: FONT, bold: true, color: C.primary,
});
s10.addText("양방의 부도덕한 행태를 답습하자는 것이 아닙니다. 한의원은 급여 항목만 보장하므로 불필요하게 의료비를 올리지 않습니다(추나 수가 통제). 추나 연 한도가 있어 건보 재정을 악화시키지도 않습니다.", {
  x: 0.5, y: 0.9, w: 8.5, h: 0.9,
  fontSize: 11, fontFace: FONT, color: C.dark,
  lineSpacingMultiple: 1.45,
});
s10.addText("도수 실비 1조원 시스템, 우리도 할 수 있습니다. 2019년 추나 급여화는 한의계의 큰 기회입니다.", {
  x: 0.5, y: 2.0, w: 8.5, h: 0.5,
  fontSize: 12, fontFace: FONT, bold: true, color: C.secondary,
});
addFooter(s10, 10, TOTAL);

// ========== 11. 제재 사례 ==========
const s11 = pptx.addSlide();
s11.background = { fill: C.bg };
s11.addText("제도를 악용하면 주목을 받고 제재를 받습니다.", {
  x: 0.5, y: 0.35, w: 7.0, h: 0.4,
  fontSize: 18, fontFace: FONT, bold: true, color: C.primary,
});
s11.addText("한의사는 필요한 치료를 하고, 환자는 합리적인 가격에 훌륭한 의료 서비스를 받게 하는 것이 목표입니다.", {
  x: 0.5, y: 0.85, w: 8.5, h: 0.35,
  fontSize: 11, fontFace: FONT, color: C.dark,
});
s11.addText("사례: 하지정맥류 시술 비용 부풀려 50억 보험사기 – 병원장·브로커 실형. 안과 백내장 수술 실손보험금 1조 넘을 듯, ‘생내장 수술’ 논란 등.", {
  x: 0.5, y: 1.35, w: 8.5, h: 1.0,
  fontSize: 10, fontFace: FONT, color: C.med,
  lineSpacingMultiple: 1.4,
});
addFooter(s11, 11, TOTAL);

// ========== 12. 목차 ==========
const s12 = pptx.addSlide();
s12.background = { fill: C.bg };
s12.addText("목차", {
  x: 0.5, y: 0.35, w: 3.0, h: 0.55,
  fontSize: 28, fontFace: FONT, bold: true, color: C.primary,
});
s12.addShape(pptx.shapes.RECTANGLE, {
  x: 0.5, y: 0.88, w: 0.06, h: 0.4,
  fill: { color: C.secondary },
});
const toc = [
  "1. 실손보험 세대별 특징",
  "2. 추나요법과 실손보험",
  "3. 첩약 건보와 실손보험",
  "4. 실손보험 FAQ",
  "5. 메디스트림 실손 서비스 활용",
];
toc.forEach((t, i) => {
  s12.addText(t, {
    x: 0.65, y: 1.0 + i * 0.58, w: 6.0, h: 0.5,
    fontSize: 16, fontFace: FONT, color: C.dark,
  });
});
addFooter(s12, 12, TOTAL);

// ========== 13. 4가지 타입 (인포그래픽 매트릭스) ==========
const s13 = pptx.addSlide();
s13.background = { fill: C.bg };
s13.addText("4가지 타입만 기억하세요", {
  x: 0.5, y: 0.3, w: 5.0, h: 0.5,
  fontSize: 20, fontFace: FONT, bold: true, color: C.primary,
});
const types = [
  { name: "급여 타입 (2~4세대)", desc: "1만원 제외 보상", fill: C.secondary },
  { name: "상해 타입 (1세대)", desc: "상해 100% 보상", fill: C.accent },
  { name: "입원 타입 (1세대)", desc: "입원 100% 보상", fill: "7C3AED" },
  { name: "급여 타입 (1세대)", desc: "5천원 공제 후 80% 보상", fill: "EA580C" },
];
types.forEach((t, i) => {
  const col = i % 2;
  const row = Math.floor(i / 2);
  const x = 0.5 + col * 4.6;
  const y = 1.0 + row * 2.0;
  s13.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: y, w: 4.3, h: 1.75,
    fill: { color: t.fill },
    rectRadius: 0.08,
  });
  s13.addText(t.name, {
    x: x + 0.15, y: y + 0.25, w: 4.0, h: 0.45,
    fontSize: 12, fontFace: FONT, bold: true, color: C.white,
  });
  s13.addText(t.desc, {
    x: x + 0.15, y: y + 0.75, w: 4.0, h: 0.5,
    fontSize: 11, fontFace: FONT, color: C.white,
  });
});
addFooter(s13, 13, TOTAL);

// ========== 14. 1세대 요약 표 ==========
const s14 = pptx.addSlide();
s14.background = { fill: C.bg };
s14.addText("1세대 요약", {
  x: 0.5, y: 0.35, w: 3.0, h: 0.5,
  fontSize: 20, fontFace: FONT, bold: true, color: C.primary,
});
s14.addTable([
  ["특약 이름", "급여 보장", "비급여 보장"],
  ["일반상해의료비 / 상해입원의료비", "O", "O"],
  ["상해통원의료비", "X", "X"],
  ["질병입원의료비", "O", "O"],
  ["질병통원의료비", "X", "X"],
  ["재해·질병입원의료비", "O", "X"],
  ["재해·질병통원의료비", "X", "X"],
], {
  x: 0.5, y: 0.95, w: 8.5, colW: [3.5, 2.5, 2.5],
  fontSize: 10, fontFace: FONT,
  fill: { color: C.ltBg },
  align: "center",
  valign: "middle",
  border: { type: "solid", pt: 0.5, color: C.light },
});
s14.addText("한의원에서 쓸 수 있는 ‘일반상해의료비’ 보험은 일 한도가 없는 경우가 많습니다. (양방은 보통 1회 10~30만원 한도)", {
  x: 0.5, y: 3.5, w: 8.5, h: 0.5,
  fontSize: 9, fontFace: FONT, color: C.med,
});
addFooter(s14, 14, TOTAL);

// ========== 15. 일반상해의료비 유의사항 ==========
const s15 = pptx.addSlide();
s15.background = { fill: C.bg };
s15.addText("일반상해의료비 보험 유의사항 (초진)", {
  x: 0.5, y: 0.35, w: 6.0, h: 0.45,
  fontSize: 18, fontFace: FONT, bold: true, color: C.primary,
});
s15.addTable([
  ["구분", "유의사항"],
  ["상해 3요소", "외래성·우연성·급격성. 반복 작업 통증은 상해 X. 장소·상황 기재."],
  ["상해 발생일", "상해일로부터 180일 이내만 보상. 지나면 보상 안 함."],
  ["인과성", "상해 사고와 증상 인과성 기록. S, T 코드 입력. 건보 접수."],
], {
  x: 0.5, y: 0.9, w: 8.5, colW: [1.8, 6.7],
  fontSize: 9, fontFace: FONT,
  fill: { color: C.ltBg },
  align: "left",
  valign: "middle",
  border: { type: "solid", pt: 0.5, color: C.light },
});
s15.addText("재진: 한약 처방 시 상해 치료와 의학적 연관 소명 자료 요청될 수 있음. 보상 기간 180일. 보상 한도는 메디 서비스로 조회 가능(보통 100만~1,000만원).", {
  x: 0.5, y: 2.8, w: 8.5, h: 0.7,
  fontSize: 9, fontFace: FONT, color: C.dark,
  lineSpacingMultiple: 1.4,
});
addFooter(s15, 15, TOTAL);

// ========== 16. 2~4세대 초간단 정리 ==========
const s16 = pptx.addSlide();
s16.background = { fill: C.bg };
s16.addText("2~4세대 초간단 정리", {
  x: 0.5, y: 0.3, w: 4.0, h: 0.5,
  fontSize: 20, fontFace: FONT, bold: true, color: C.primary,
});
s16.addText("무조건 1만원 실손: 2013년 3월까지 판매. 20% 실손: 1만원 또는 20% 중 큰 금액 제외 보상.", {
  x: 0.5, y: 0.85, w: 8.5, h: 0.4,
  fontSize: 10, fontFace: FONT, color: C.dark,
});
s16.addTable([
  ["세대", "판매 기간", "통원 자기부담금", "입원 자기부담금"],
  ["2-1", "2009.08~2013.03", "한의원 1만/한방병원 1.5만", "10% (90% 보상)"],
  ["2-2", "2013.04~2015.08", "표준 1만 vs 20% 큰 금액 / 선택형 1만 vs 10%", "표준 20% / 선택 10%, 연 200만 한도"],
  ["2-3", "2015.09~2017.03", "표준/선택2 유사", "표준 20% / 선택2 10%(비급여 80%만)"],
  ["3", "2017.04~2021.06", "3대 비급여 분리, 2만 vs 30%", "동일"],
  ["4", "2021.07~현재", "1만 vs 20% / 비급여 3만 vs 30%", "재가입 5년, 비급여 할증 최대 300%"],
], {
  x: 0.5, y: 1.35, w: 8.5, colW: [1.0, 2.2, 2.4, 2.9],
  fontSize: 8, fontFace: FONT,
  fill: { color: C.ltBg },
  align: "left",
  valign: "middle",
  border: { type: "solid", pt: 0.5, color: C.light },
});
addFooter(s16, 16, TOTAL);

// ========== 17. 추나 진료비 시나리오 ==========
const s17 = pptx.addSlide();
s17.background = { fill: C.bg };
s17.addText("추나 진료비 시나리오", {
  x: 0.5, y: 0.35, w: 4.0, h: 0.5,
  fontSize: 20, fontFace: FONT, bold: true, color: C.primary,
});
s17.addTable([
  ["조합", "구분", "총 진료비", "무조건 1만원", "20% 실손"],
  ["침/유관/IR+복잡80", "초진", "78,630원", "10,000원", "10,000원"],
  ["침/유관/IR+복잡80", "재진", "73,090원", "10,000원", "10,000원"],
  ["침/유관/IR+단순", "초진", "61,460원", "10,000원", "10,000원"],
  ["침/유관/IR+단순", "재진", "55,920원", "10,000원", "10,000원"],
], {
  x: 0.5, y: 0.95, w: 8.5, colW: [2.2, 1.2, 2.0, 1.6, 1.5],
  fontSize: 9, fontFace: FONT,
  fill: { color: C.ltBg },
  align: "center",
  valign: "middle",
  border: { type: "solid", pt: 0.5, color: C.light },
});
s17.addText("단순이나 복잡이나 환자는 1만원. 초진 당일 청구하게 만들어야 합니다. 1~2일 내 보상 받으면 재진율이 올라갑니다.", {
  x: 0.5, y: 2.9, w: 8.5, h: 0.5,
  fontSize: 10, fontFace: FONT, color: C.dark,
});
addFooter(s17, 17, TOTAL);

// ========== 18. 초진 당일 청구 (플로우 인포그래픽) ==========
const s18 = pptx.addSlide();
s18.background = { fill: C.bg };
s18.addText("초진은 수납 후 바로 청구하게 하세요!", {
  x: 0.5, y: 0.35, w: 6.0, h: 0.5,
  fontSize: 20, fontFace: FONT, bold: true, color: C.primary,
});
const flowSteps = ["수납", "당일 청구", "1~2일 내 보상", "재진 유도"];
const fw = 2.0;
flowSteps.forEach((step, i) => {
  const x = 0.6 + i * 2.35;
  s18.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: 1.2, w: 2.0, h: 0.9,
    fill: { color: i === 1 ? C.secondary : C.ltBg },
    line: { color: C.light, width: 0.5 },
    rectRadius: 0.06,
  });
  s18.addText(step, {
    x: x + 0.1, y: 1.5, w: 1.8, h: 0.35,
    fontSize: 12, fontFace: FONT, bold: true, color: i === 1 ? C.white : C.primary,
    align: "center", valign: "middle",
  });
  if (i < flowSteps.length - 1) {
    s18.addShape(pptx.shapes.RECTANGLE, {
      x: x + 2.05, y: 1.64, w: 0.25, h: 0.02,
      fill: { color: C.med },
    });
    s18.addShape(pptx.shapes.ISOSCELES_TRIANGLE, {
      x: x + 2.2, y: 1.58, w: 0.15, h: 0.14,
      fill: { color: C.med },
      rotate: 270,
    });
  }
});
s18.addText("혹시 모를 예외 상황(부담보 등) 대비·총 진료비 커지기 전 문제 파악·보상 담당자 실수 여부 확인 가능.", {
  x: 0.5, y: 2.4, w: 8.5, h: 0.6,
  fontSize: 10, fontFace: FONT, color: C.dark,
  lineSpacingMultiple: 1.4,
});
addFooter(s18, 18, TOTAL);

// ========== 19. 첩약 건보 ==========
const s19 = pptx.addSlide();
s19.background = { fill: C.bg };
s19.addText("첩약 건보 (4월 시행 예정)", {
  x: 0.5, y: 0.35, w: 5.0, h: 0.5,
  fontSize: 20, fontFace: FONT, bold: true, color: C.primary,
});
s19.addText("대상 질환: 월경통, 안면신경마비, 뇌혈관질환 후유증, 요추추간판탈출증, 알레르기 비염, 기능성 소화불량 등. 질환당 20일씩 처방(연간 2개 질환까지), 이후 100/100 급여.", {
  x: 0.5, y: 0.9, w: 8.5, h: 0.7,
  fontSize: 10, fontFace: FONT, color: C.dark,
  lineSpacingMultiple: 1.4,
});
s19.addText("수가: 심층변증방제기술료 47,140원 / 탕전료(자체) 46,980원, (원외) 34,380원. 약재비 변증 41,450원~월경통 74,450원 등.", {
  x: 0.5, y: 1.7, w: 8.5, h: 0.6,
  fontSize: 10, fontFace: FONT, color: C.med,
  lineSpacingMultiple: 1.4,
});
addFooter(s19, 19, TOTAL);

// ========== 20. 청구 절차 (플로우 인포그래픽) ==========
const s20 = pptx.addSlide();
s20.background = { fill: C.bg };
s20.addText("청구 절차의 이해", {
  x: 0.5, y: 0.35, w: 4.0, h: 0.5,
  fontSize: 20, fontFace: FONT, bold: true, color: C.primary,
});
const steps = ["청구서 접수", "보상 담당자 지정", "심사", "보상 지급"];
steps.forEach((step, i) => {
  const x = 0.5 + i * 2.35;
  s20.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: 1.1, w: 2.1, h: 0.85,
    fill: { color: C.ltBg },
    line: { color: C.secondary, width: 0.5 },
    rectRadius: 0.06,
  });
  s20.addText(step, {
    x: x + 0.1, y: 1.35, w: 1.9, h: 0.35,
    fontSize: 11, fontFace: FONT, bold: true, color: C.primary,
    align: "center", valign: "middle",
  });
  if (i < steps.length - 1) {
    s20.addShape(pptx.shapes.RECTANGLE, {
      x: x + 2.15, y: 1.51, w: 0.2, h: 0.02,
      fill: { color: C.secondary },
    });
    s20.addShape(pptx.shapes.ISOSCELES_TRIANGLE, {
      x: x + 2.28, y: 1.46, w: 0.12, h: 0.12,
      fill: { color: C.secondary },
      rotate: 270,
    });
  }
});
s20.addText("담당자 지정 시점에 환자가 보험사 문자 수신. 법적으로 3영업일 이내 보험금 지급 의무(서류 접수일 제외). 실무적으로 생보 평균 0.9일, 손보 평균 1.1일.", {
  x: 0.5, y: 2.2, w: 8.5, h: 0.65,
  fontSize: 10, fontFace: FONT, color: C.dark,
  lineSpacingMultiple: 1.4,
});
addFooter(s20, 20, TOTAL);

// ========== 21. 비례보상 ==========
const s21 = pptx.addSlide();
s21.background = { fill: C.bg };
s21.addText("보험을 여러 개 갖고 있는 경우", {
  x: 0.5, y: 0.35, w: 5.0, h: 0.45,
  fontSize: 18, fontFace: FONT, bold: true, color: C.primary,
});
s21.addText("실손보험은 중복 보상하지 않습니다. 보험사별 책임 비율에 따라 비례 보상합니다. 보험금이 덜 들어왔다면 먼저 모든 보험사에 청구했는지 확인하세요.", {
  x: 0.5, y: 0.9, w: 8.5, h: 0.6,
  fontSize: 11, fontFace: FONT, color: C.dark,
  lineSpacingMultiple: 1.4,
});
s21.addText("예: 총 진료비 20만원(급여 5만, 비급여 15만). 1세대 상해 100% + 2세대 1만원 제외 → 각 보험사에 청구해야 전액 수령.", {
  x: 0.5, y: 1.6, w: 8.5, h: 0.5,
  fontSize: 10, fontFace: FONT, color: C.med,
});
addFooter(s21, 21, TOTAL);

// ========== 22. FAQ – 청구하면 보험료 / 한의원 안된대요 ==========
const s22 = pptx.addSlide();
s22.background = { fill: C.bg };
s22.addText("실손보험 FAQ", {
  x: 0.5, y: 0.35, w: 3.0, h: 0.5,
  fontSize: 20, fontFace: FONT, bold: true, color: C.primary,
});
s22.addText("• \"청구하면 보험료 오른대요\" → 나이 때문에 오르는 것이지, 많이 청구해서 오르지 않습니다. 다만 4세대부터 비급여 청구액에 따라 할증 최대 300%. 추나는 건보라 할증 없음.", {
  x: 0.5, y: 0.95, w: 8.5, h: 0.85,
  fontSize: 10, fontFace: FONT, color: C.dark,
  lineSpacingMultiple: 1.4,
});
s22.addText("• \"내 보험으론 한의원 치료 못한대요\" → 잘못된 보험사에 청구(손보/생보 혼동), 일부만 청구, 심사 담당자 실수, 상해/질병 구분 오류 등이 있을 수 있습니다. 조회 결과에 따라 치료·청구하세요.", {
  x: 0.5, y: 1.95, w: 8.5, h: 0.85,
  fontSize: 10, fontFace: FONT, color: C.dark,
  lineSpacingMultiple: 1.4,
});
addFooter(s22, 22, TOTAL);

// ========== 23. 실손 조회 권유 (플로우) ==========
const s23 = pptx.addSlide();
s23.background = { fill: C.bg };
s23.addText("실손 조회, 어떻게 권하나요?", {
  x: 0.5, y: 0.35, w: 5.0, h: 0.5,
  fontSize: 18, fontFace: FONT, bold: true, color: C.primary,
});
s23.addText("초진 설문지: 실비 보험 있으신가요? (네/아니요/모름) → 있으시면 한의원에서 어떤 치료 보장받는지 확인해 드릴 수 있습니다. 원장님 상담 시 추나 필요성 강조하면 \"아니요\" 답한 분도 조회하게 만들 수 있습니다.", {
  x: 0.5, y: 0.95, w: 8.5, h: 1.0,
  fontSize: 10, fontFace: FONT, color: C.dark,
  lineSpacingMultiple: 1.4,
});
s23.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 2.2, w: 2.2, h: 0.7,
  fill: { color: C.ltBg },
  line: { color: C.secondary, width: 0.5 },
  rectRadius: 0.06,
});
s23.addText("조회하세요", { x: 0.6, y: 2.4, w: 2.0, h: 0.3, fontSize: 12, fontFace: FONT, bold: true, color: C.primary, align: "center" });
s23.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 2.9, y: 2.2, w: 2.2, h: 0.7,
  fill: { color: C.secondary },
  rectRadius: 0.06,
});
s23.addText("자신 있게 상담", { x: 3.0, y: 2.4, w: 2.0, h: 0.3, fontSize: 12, fontFace: FONT, bold: true, color: C.white, align: "center" });
s23.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 5.3, y: 2.2, w: 2.2, h: 0.7,
  fill: { color: C.ltBg },
  line: { color: C.secondary, width: 0.5 },
  rectRadius: 0.06,
});
s23.addText("청구하게 하세요", { x: 5.4, y: 2.4, w: 2.0, h: 0.3, fontSize: 12, fontFace: FONT, bold: true, color: C.primary, align: "center" });
s23.addText("→ 객단가 상승, 재진율 상승 = 월 매출 천만원을 위한 단순한 원리", {
  x: 0.5, y: 3.05, w: 8.5, h: 0.35,
  fontSize: 11, fontFace: FONT, bold: true, color: C.secondary,
});
addFooter(s23, 23, TOTAL);

// ========== 24. 메디스트림 서비스 ==========
const s24 = pptx.addSlide();
s24.background = { fill: C.bg };
s24.addText("메디스트림 실손 서비스 활용", {
  x: 0.5, y: 0.35, w: 5.0, h: 0.5,
  fontSize: 20, fontFace: FONT, bold: true, color: C.primary,
});
const services = [
  "마이데이터 조회: 여러 보험사 정보 한곳에, 100% 자동화 24/365, 실손 특약 상세·보상 한도 조회",
  "보장분석 콜센터(3월 말 예정): 환자 전화 → 이름·주민번호로 가입 내역 조회 → 한의원 보장 전화 상담 → 원장님께 결과 공유",
  "주계약 조회: 휴대전화 본인인증으로 30초 내 확인. 2018년 4월 이후 실손은 무조건 단독 상품.",
  "환자 관리(insudesk): 환자별 보장 분석, 조회/청구 링크 전송. 청구: 40개 보험사·공제조합, 24/365, 팩스 무제한.",
];
services.forEach((sv, i) => {
  s24.addShape(pptx.shapes.OVAL, {
    x: 0.52, y: 1.0 + i * 0.95, w: 0.15, h: 0.15,
    fill: { color: C.secondary },
  });
  s24.addText(sv, {
    x: 0.75, y: 0.98 + i * 0.95, w: 8.2, h: 0.85,
    fontSize: 9, fontFace: FONT, color: C.dark,
    lineSpacingMultiple: 1.35,
  });
});
addFooter(s24, 24, TOTAL);

// ========== 25. 플랜 요금 (3컬럼 인포그래픽) ==========
const s25 = pptx.addSlide();
s25.background = { fill: C.bg };
s25.addText("월 구독 요금", {
  x: 0.5, y: 0.35, w: 3.0, h: 0.5,
  fontSize: 20, fontFace: FONT, bold: true, color: C.primary,
});
const plans = [
  { name: "베이직 플랜", price: "무료", sub: "주계약 리스트만 빠르게 조회" },
  { name: "실손 에센셜", price: "월 30만원↓23만원", sub: "마이데이터 조회, 지급거절 1명당 월 최대 30만원 보상, 콜센터, 정기 교육, 청구 팩스 무제한" },
  { name: "실손 컨시어지", price: "월 150만원↓120만원", sub: "에센셜 전체 + 무제한 보장분석 콜센터, 지급거절 무제한 보상" },
];
plans.forEach((p, i) => {
  const x = 0.5 + i * 3.1;
  s25.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: 0.95, w: 2.9, h: 2.5,
    fill: { color: i === 1 ? C.secondary : C.ltBg },
    line: { color: C.light, width: 0.5 },
    rectRadius: 0.08,
  });
  s25.addText(p.name, {
    x: x + 0.1, y: 1.05, w: 2.7, h: 0.4,
    fontSize: 14, fontFace: FONT, bold: true, color: i === 1 ? C.white : C.primary,
    align: "center",
  });
  s25.addText(p.price, {
    x: x + 0.1, y: 1.55, w: 2.7, h: 0.4,
    fontSize: 12, fontFace: FONT, bold: true, color: i === 1 ? C.white : C.secondary,
    align: "center",
  });
  s25.addText(p.sub, {
    x: x + 0.12, y: 2.05, w: 2.66, h: 1.3,
    fontSize: 8, fontFace: FONT, color: i === 1 ? C.white : C.med,
    wrap: true,
    lineSpacingMultiple: 1.3,
  });
});
s25.addText("지급 거절 보상: 메디스트림 안내로 한의원 보장 가능하다고 했으나 실제 보장 불가인 경우 환불·한의원 보상 (심평원 청구 제외).", {
  x: 0.5, y: 3.6, w: 8.5, h: 0.5,
  fontSize: 8, fontFace: FONT, color: C.med,
});
addFooter(s25, 25, TOTAL);

// ========== 26. 마무리 ==========
const s26 = pptx.addSlide();
s26.background = { fill: C.primary };
s26.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 0.1, h: 5.25,
  fill: { color: C.secondary },
});
s26.addText("한번에 끝내는 실손 강의", {
  x: 0.5, y: 1.8, w: 8.0, h: 0.6,
  fontSize: 32, fontFace: FONT, bold: true, color: C.white,
  align: "center",
});
s26.addText("5,000건 보험 데이터 분석 · 메디스트림", {
  x: 0.5, y: 2.5, w: 8.0, h: 0.4,
  fontSize: 14, fontFace: FONT, color: C.secondary,
  align: "center",
});
addFooter(s26, 26, TOTAL);

// ========== 저장 ==========
const outPath = path.join(__dirname, "실손강의_메디.pptx");
pptx.writeFile({ fileName: outPath }).then(() => {
  console.log("저장 완료:", outPath);
}).catch((err) => {
  console.error("저장 실패:", err);
});
