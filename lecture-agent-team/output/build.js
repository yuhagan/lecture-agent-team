"use strict";
const pptxgen = require("pptxgenjs");
const pptx = new pptxgen();

pptx.layout = "LAYOUT_16x9";
pptx.author = "강의 자동화 시스템";

// ── 디자인 토큰 ──────────────────────────────────────────────────────────────
const C = {
  bgLight:       "F5F0EB",
  bgDark:        "1A1A1A",
  cardDark:      "2A2A2A",
  cardLight:     "FFFFFF",
  textPrimary:   "1A1A1A",
  textInverse:   "FFFFFF",
  textMuted:     "9CA3AF",
  accentDanger:  "E53E3E",
  accentSuccess: "48BB78",
  accentBrand:   "1E3A8A",
  badgeBg:       "374151",
  badgeBgLight:  "E5E7EB",
};
const F = "Pretendard";

// ── 공통 헬퍼 ────────────────────────────────────────────────────────────────
function bg(slide, isDark) {
  slide.background = { color: isDark ? C.bgDark : C.bgLight };
}

function addBadge(slide, text, isDark) {
  const bgColor = isDark ? C.badgeBgLight : C.badgeBg;
  const txColor = isDark ? C.textPrimary : C.textInverse;
  const w = Math.max(text.length * 0.13 + 0.6, 1.4);
  const x = (10 - w) / 2;
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y: 0.28, w, h: 0.32,
    fill: { color: bgColor }, line: { color: bgColor }, rectRadius: 0.16,
  });
  slide.addText(text, {
    x, y: 0.28, w, h: 0.32,
    fontSize: 11, bold: true, color: txColor, fontFace: F,
    align: "center", valign: "middle",
  });
}

function addSectionTitle(slide, text, isDark) {
  slide.addText(text, {
    x: 0.6, y: 0.75, w: 8.8, h: 0.65,
    fontSize: 34, bold: true,
    color: isDark ? C.textInverse : C.textPrimary,
    fontFace: F, valign: "middle",
  });
}

function addCard(slide, opts) {
  const { x, y, w, h, isDark, title, body, bodyColor, titleSize, bodySize } = opts;
  const bgCol = isDark ? C.cardDark : C.cardLight;
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h, fill: { color: bgCol }, line: { color: bgCol }, rectRadius: 0.22,
  });
  if (title) {
    slide.addText(title, {
      x: x + 0.22, y: y + 0.15, w: w - 0.44, h: 0.42,
      fontSize: titleSize || 18, bold: true,
      color: isDark ? C.textInverse : C.textPrimary, fontFace: F,
    });
  }
  if (body) {
    slide.addText(body, {
      x: x + 0.22, y: y + 0.6, w: w - 0.44, h: h - 0.78,
      fontSize: bodySize || 13, color: bodyColor || C.textMuted,
      fontFace: F, wrap: true,
    });
  }
}

function hline(slide, x, y, w, color) {
  slide.addShape(pptx.ShapeType.line, {
    x, y, w, h: 0,
    line: { color: color || C.textMuted, width: 1 },
  });
}

function vline(slide, x, y, h, color, widthPx) {
  slide.addShape(pptx.ShapeType.rect, {
    x, y, w: (widthPx || 3) / 72, h,
    fill: { color: color || C.accentBrand },
    line: { color: color || C.accentBrand },
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// 슬라이드 1 — 표지 (TYPE-A, Light)
// ═══════════════════════════════════════════════════════════════════════════
(function slide01() {
  const s = pptx.addSlide();
  bg(s, false);

  s.addShape(pptx.ShapeType.rect, {
    x: 0.5, y: 0.7, w: 0.07, h: 3.2,
    fill: { color: C.accentBrand }, line: { color: C.accentBrand },
  });

  s.addText("손저림의 한의학적 이해와 치료", {
    x: 0.75, y: 0.8, w: 8.8, h: 1.5,
    fontSize: 46, bold: true, color: C.textPrimary, fontFace: F,
    valign: "middle",
  });

  s.addText("원인 감별부터 임상 적용까지", {
    x: 0.75, y: 2.35, w: 8.8, h: 0.6,
    fontSize: 22, color: C.accentBrand, fontFace: F, valign: "middle",
  });

  hline(s, 0.75, 3.1, 6, C.textMuted);

  s.addText("한의학 임상 강의 시리즈  |  총 20장  |  40분", {
    x: 0.75, y: 3.3, w: 8.0, h: 0.4,
    fontSize: 14, color: C.textMuted, fontFace: F,
  });

  s.addNotes("【발표 멘트】\n안녕하십니까. 오늘 강의에서는 외래에서 자주 마주치는 '손저림'이라는 증상을 깊이 있게 다뤄보겠습니다. 환자분들이 \"손이 저려요\"라고 호소할 때, 우리가 어떤 감별 과정을 거쳐야 하는지, 그리고 한의학적으로 어떤 근거 위에서 치료할 수 있는지를 체계적으로 정리해 드리겠습니다. 임상에서 바로 활용하실 수 있는 내용 중심으로 구성했습니다.\n\n【전환 멘트】\n먼저 오늘 강의의 전체 흐름을 간단히 살펴보겠습니다.");
})();

// ═══════════════════════════════════════════════════════════════════════════
// 슬라이드 2 — 목차 (TYPE-B, Dark)
// ═══════════════════════════════════════════════════════════════════════════
(function slide02() {
  const s = pptx.addSlide();
  bg(s, true);
  addBadge(s, "CONTENTS", true);
  addSectionTitle(s, "오늘 강의 목차", true);

  const items = [
    { num: "01", text: "왜 손저림인가? — 역학과 현황" },
    { num: "02", text: "원인 질환별 병태생리와 해부학" },
    { num: "03", text: "감별 진단 — 이학적 검사와 감별 포인트" },
    { num: "04", text: "한의학 치료 근거 — 침, 한약, 추나" },
    { num: "05", text: "임상 적용 가이드 — 변증, 알고리즘, Red Flags" },
  ];

  items.forEach(function(item, i) {
    const y = 1.6 + i * 0.72;
    s.addShape(pptx.ShapeType.roundRect, {
      x: 0.6, y: y, w: 0.55, h: 0.52,
      fill: { color: C.accentBrand }, line: { color: C.accentBrand }, rectRadius: 0.1,
    });
    s.addText(item.num, {
      x: 0.6, y: y, w: 0.55, h: 0.52,
      fontSize: 16, bold: true, color: C.textInverse, fontFace: F,
      align: "center", valign: "middle",
    });
    s.addText(item.text, {
      x: 1.3, y: y + 0.04, w: 8.0, h: 0.45,
      fontSize: 18, color: C.textInverse, fontFace: F, valign: "middle",
    });
  });

  s.addNotes("【발표 멘트】\n오늘 강의는 크게 다섯 파트로 구성됩니다. 먼저 손저림이 왜 중요한 주제인지 역학 데이터로 확인하고, 원인 질환별 병태생리를 해부학적으로 정리합니다. 이어서 외래에서 실제로 쓸 수 있는 감별 진단 포인트를 정리한 뒤, 침치료와 한약의 RCT 근거를 살펴보겠습니다. 마지막으로 변증별 치료 알고리즘과 위험신호까지 다루겠습니다.\n\n【전환 멘트】\n그럼 첫 번째 파트, 왜 지금 손저림에 주목해야 하는지부터 시작하겠습니다.");
})();

// ═══════════════════════════════════════════════════════════════════════════
// 슬라이드 3 — 도입 질문 (TYPE-C, Light)
// ═══════════════════════════════════════════════════════════════════════════
(function slide03() {
  const s = pptx.addSlide();
  bg(s, false);
  addBadge(s, "INTRODUCTION", false);

  s.addText("\"선생님, 손이 저려요\"", {
    x: 0.6, y: 0.8, w: 8.8, h: 1.1,
    fontSize: 44, bold: true, color: C.textPrimary, fontFace: F,
    align: "center", valign: "middle",
  });

  hline(s, 2.0, 2.0, 6.0, C.textMuted);

  s.addText("손저림(Hand Numbness/Tingling) = 감각 이상의 총칭", {
    x: 0.6, y: 2.15, w: 8.8, h: 0.45,
    fontSize: 18, bold: true, color: C.textPrimary, fontFace: F, align: "center",
  });

  var terms = [
    { title: "지각이상", sub: "Paresthesia" },
    { title: "감각저하", sub: "Hypoesthesia" },
    { title: "이감각증", sub: "Dysesthesia" },
  ];
  terms.forEach(function(t, i) {
    var x = 0.75 + i * 2.95;
    s.addShape(pptx.ShapeType.roundRect, {
      x: x, y: 2.8, w: 2.6, h: 1.0,
      fill: { color: C.cardLight }, line: { color: C.cardLight }, rectRadius: 0.22,
    });
    s.addText(t.title, {
      x: x + 0.1, y: 2.88, w: 2.4, h: 0.42,
      fontSize: 20, bold: true, color: C.textPrimary, fontFace: F, align: "center",
    });
    s.addText(t.sub, {
      x: x + 0.1, y: 3.32, w: 2.4, h: 0.38,
      fontSize: 14, color: C.textMuted, fontFace: F, align: "center",
    });
  });

  s.addText("말초신경 포착부터 중추성 병변까지 — 스펙트럼이 넓다", {
    x: 0.6, y: 4.05, w: 8.8, h: 0.45,
    fontSize: 16, color: C.accentBrand, fontFace: F, align: "center", bold: true,
  });

  s.addNotes("【발표 멘트】\n여러분, 한 주에 \"손이 저려요\"라는 호소를 몇 번이나 들으십니까? 아마 손에 꼽기 어려우실 겁니다. 손저림은 의학적으로 지각이상(Paresthesia), 감각저하(Hypoesthesia), 이감각증(Dysesthesia)으로 세분되는데, 환자분들은 이 모든 걸 '저림'이라는 한 단어로 표현합니다. 문제는 이 단순한 증상 뒤에 수근관증후군부터 뇌졸중까지 매우 다양한 원인이 숨어 있다는 점입니다.\n\n【전환 멘트】\n그렇다면 실제로 어떤 질환들이 얼마나 흔한지, 역학 데이터를 살펴보겠습니다.");
})();

// ═══════════════════════════════════════════════════════════════════════════
// 슬라이드 4 — 역학: CTS (TYPE-G, Dark)
// ═══════════════════════════════════════════════════════════════════════════
(function slide04() {
  const s = pptx.addSlide();
  bg(s, true);
  addBadge(s, "EPIDEMIOLOGY", true);
  addSectionTitle(s, "수근관증후군(CTS) — 역학 데이터", true);

  s.addText("3.8%", {
    x: 0.6, y: 1.55, w: 3.0, h: 1.1,
    fontSize: 90, bold: true, color: C.accentSuccess, fontFace: F,
    align: "center", valign: "middle",
  });
  s.addText("일반 성인 유병률", {
    x: 0.6, y: 2.65, w: 3.0, h: 0.35,
    fontSize: 13, color: C.textMuted, fontFace: F, align: "center",
  });
  s.addText("Atroshi et al. (1999, JAMA)", {
    x: 0.6, y: 2.95, w: 3.0, h: 0.3,
    fontSize: 11, color: C.textMuted, fontFace: F, align: "center",
  });

  var stats = [
    { title: "여성 3배", body: "남성 대비 발생률\n40~60대 집중" },
    { title: "20만 명+", body: "한국 HIRA 2022년\n연간 진료 인원" },
    { title: "10~15%", body: "반복 작업 직군\n유병률 (임신 중 43%까지)" },
  ];
  stats.forEach(function(st, i) {
    addCard(s, {
      x: 4.2 + i * 1.95, y: 1.5, w: 1.75, h: 2.5,
      isDark: true, title: st.title, body: st.body,
    });
  });

  s.addShape(pptx.ShapeType.roundRect, {
    x: 0.6, y: 4.15, w: 8.8, h: 0.85,
    fill: { color: C.cardDark }, line: { color: C.cardDark }, rectRadius: 0.22,
  });
  s.addText("경추 신경근병증(Cervical Radiculopathy)  |  연간 인구 10만 명당 83.2명  |  C6·C7 신경근 70% 이상", {
    x: 0.8, y: 4.25, w: 8.4, h: 0.6,
    fontSize: 14, color: C.textInverse, fontFace: F, align: "center", valign: "middle",
  });

  s.addNotes("【발표 멘트】\n역학 데이터를 보면, 수근관증후군은 Atroshi 등의 JAMA 연구에서 일반 성인 유병률이 약 3.8%로 보고되었습니다. 여성이 남성보다 3배 높고, 건강보험심사평가원 자료에서도 2022년 기준 약 20만 명 이상이 진료를 받았습니다. 반복적인 손목 사용 직군에서는 유병률이 10~15%까지 올라갑니다. 경추 신경근병증도 인구 10만 명당 연간 83.2명으로 적지 않은 빈도이며, C6와 C7 신경근이 전체의 70% 이상을 차지합니다.\n\n【전환 멘트】\nCTS와 경추 신경근병증 외에도 반드시 알아야 할 원인 질환들이 더 있습니다.");
})();

// ═══════════════════════════════════════════════════════════════════════════
// 슬라이드 5 — 역학: 기타 원인 (TYPE-D, Light)
// ═══════════════════════════════════════════════════════════════════════════
(function slide05() {
  const s = pptx.addSlide();
  bg(s, false);
  addBadge(s, "EPIDEMIOLOGY", false);
  addSectionTitle(s, "그 밖의 주요 원인 질환", false);

  var cards = [
    {
      title: "흉곽출구증후군\n(TOS)",
      body: "유병률 0.3~2%\n신경성 TOS 95% 이상\n20~40대 여성·사무직 호발",
    },
    {
      title: "당뇨병성\n말초신경병증(DPN)",
      body: "당뇨 환자 약 50%에서 발생\n(Boulton et al., 2005, Lancet)\n양측 대칭 장갑·양말형 분포",
    },
    {
      title: "레이노 현상\n(Raynaud Phenomenon)",
      body: "여성 ~20%, 남성 ~11%\n원발성이 이차성보다 4배\n한랭 자극 → 혈관 과수축",
    },
  ];

  cards.forEach(function(c, i) {
    addCard(s, {
      x: 0.6 + i * 3.15, y: 1.65, w: 2.9, h: 2.8,
      isDark: false, title: c.title, body: c.body,
    });
  });

  s.addShape(pptx.ShapeType.roundRect, {
    x: 0.6, y: 4.65, w: 8.8, h: 0.68,
    fill: { color: C.cardLight },
    line: { color: C.accentBrand, pt: 2 }, rectRadius: 0.22,
  });
  s.addText("임상 5대 원인: CTS · 경추 신경근병증 · TOS · DPN · 레이노 현상", {
    x: 0.8, y: 4.72, w: 8.4, h: 0.52,
    fontSize: 15, bold: true, color: C.accentBrand, fontFace: F, align: "center", valign: "middle",
  });

  s.addNotes("【발표 멘트】\nCTS와 경추 질환 외에 세 가지 더 기억하셔야 합니다. 흉곽출구증후군은 유병률 0.3~2%로 낮아 보이지만, 20~40대 사무직 여성에서 호발하므로 놓치기 쉽습니다. 당뇨병성 말초신경병증은 당뇨 환자의 절반에서 발생하는데, Boulton 등의 Lancet 논문에서 확인된 수치입니다. 양측 대칭성의 장갑·양말형 분포가 특징입니다. 레이노 현상은 여성에서 약 20%까지 보고될 정도로 흔하며, 한랭 자극에 의한 혈관 수축이 핵심 기전입니다.\n\n【전환 멘트】\n이제 각 질환의 병태생리를 해부학적으로 좀 더 깊이 들어가 보겠습니다.");
})();

// ═══════════════════════════════════════════════════════════════════════════
// 슬라이드 6 — CTS 해부학 다이어그램 (TYPE-H, Dark)
// ═══════════════════════════════════════════════════════════════════════════
(function slide06() {
  const s = pptx.addSlide();
  bg(s, true);
  addBadge(s, "ANATOMY", true);
  addSectionTitle(s, "수근관증후군(CTS) — 해부학과 병태생리", true);

  // 수근관 단면 다이어그램
  s.addShape(pptx.ShapeType.roundRect, {
    x: 0.6, y: 1.6, w: 4.0, h: 2.5,
    fill: { color: C.cardDark }, line: { color: "4B5563" }, rectRadius: 0.3,
  });
  s.addText("수근관(Carpal Tunnel) 단면도", {
    x: 0.7, y: 1.65, w: 3.8, h: 0.32,
    fontSize: 11, color: C.textMuted, fontFace: F, align: "center",
  });
  // 횡수근인대
  s.addShape(pptx.ShapeType.rect, {
    x: 0.85, y: 2.0, w: 3.5, h: 0.28,
    fill: { color: C.accentBrand }, line: { color: C.accentBrand },
  });
  s.addText("횡수근인대 (Flexor Retinaculum)", {
    x: 0.85, y: 2.0, w: 3.5, h: 0.28,
    fontSize: 10, color: C.textInverse, fontFace: F, align: "center", valign: "middle",
  });
  // 정중신경
  s.addShape(pptx.ShapeType.ellipse, {
    x: 1.5, y: 2.4, w: 0.85, h: 0.55,
    fill: { color: C.accentDanger }, line: { color: C.accentDanger },
  });
  s.addText("정중\n신경", {
    x: 1.5, y: 2.4, w: 0.85, h: 0.55,
    fontSize: 9, bold: true, color: C.textInverse, fontFace: F,
    align: "center", valign: "middle",
  });
  // 굴곡건들
  var tendonPos = [
    [2.5, 2.42], [3.0, 2.42], [3.5, 2.42],
    [2.5, 2.92], [3.0, 2.92], [3.5, 2.92],
    [2.5, 3.42], [3.0, 3.42], [3.5, 3.42],
  ];
  tendonPos.forEach(function(pos) {
    s.addShape(pptx.ShapeType.ellipse, {
      x: pos[0], y: pos[1], w: 0.35, h: 0.35,
      fill: { color: "6B7280" }, line: { color: "9CA3AF" },
    });
  });
  s.addText("굴곡건 9개", {
    x: 2.35, y: 3.8, w: 1.3, h: 0.22,
    fontSize: 10, color: C.textMuted, fontFace: F, align: "center",
  });

  // 압력 데이터
  var pressureData = [
    { label: "정상 내압", value: "2.5 mmHg", color: C.accentSuccess },
    { label: "CTS 안정 시", value: "32 mmHg", color: C.accentDanger },
    { label: "손목 굴곡 시", value: "90+ mmHg", color: C.accentDanger },
  ];
  pressureData.forEach(function(pd, i) {
    var y = 1.65 + i * 0.9;
    s.addShape(pptx.ShapeType.roundRect, {
      x: 4.85, y: y, w: 4.75, h: 0.78,
      fill: { color: C.cardDark }, line: { color: C.cardDark }, rectRadius: 0.22,
    });
    s.addText(pd.label, {
      x: 5.05, y: y + 0.08, w: 2.0, h: 0.3,
      fontSize: 12, color: C.textMuted, fontFace: F, valign: "middle",
    });
    s.addText(pd.value, {
      x: 7.1, y: y + 0.08, w: 2.3, h: 0.55,
      fontSize: 28, bold: true, color: pd.color, fontFace: F,
      align: "right", valign: "middle",
    });
  });

  s.addShape(pptx.ShapeType.roundRect, {
    x: 4.85, y: 4.35, w: 4.75, h: 0.72,
    fill: { color: "2A1A1A" }, line: { color: C.accentDanger, pt: 1.5 }, rectRadius: 0.22,
  });
  s.addText("압박 → 탈수초(Demyelination) → 1~3번째 손가락 + 4번째 내측 반쪽 저림", {
    x: 5.0, y: 4.42, w: 4.45, h: 0.55,
    fontSize: 12, color: C.accentDanger, fontFace: F, valign: "middle", wrap: true,
  });

  s.addNotes("【발표 멘트】\n수근관증후군의 해부학을 정리해 보겠습니다. 수근관은 손목의 수근골과 횡수근인대가 만드는 폐쇄된 터널입니다. 이 좁은 공간 안에 정중신경 하나와 굴곡건 9개가 함께 들어있습니다. Szabo와 Chidgey의 연구에 따르면, 정상인의 수근관 내압은 겨우 2.5 mmHg인데 CTS 환자에서는 안정 시에도 32 mmHg, 손목을 굴곡하면 90 mmHg 이상으로 급등합니다. 이 압력이 정중신경의 탈수초를 일으켜 엄지부터 넷째 손가락 내측까지의 저림이 발생합니다.\n\n【전환 멘트】\n다음으로, 경추에서 오는 손저림의 해부학적 배경을 보겠습니다.");
})();

// ═══════════════════════════════════════════════════════════════════════════
// 슬라이드 7 — 피부분절 (TYPE-H, Light)
// ═══════════════════════════════════════════════════════════════════════════
(function slide07() {
  const s = pptx.addSlide();
  bg(s, false);
  addBadge(s, "ANATOMY", false);
  addSectionTitle(s, "경추 신경근병증 — 피부분절(Dermatome)", false);

  // 피부분절 시각화
  s.addShape(pptx.ShapeType.roundRect, {
    x: 0.6, y: 1.55, w: 3.5, h: 3.4,
    fill: { color: "FFF5E0" }, line: { color: C.textMuted }, rectRadius: 0.4,
  });
  s.addText("상지 피부분절 (Dermatome)", {
    x: 0.7, y: 1.65, w: 3.3, h: 0.4,
    fontSize: 11, color: C.textMuted, fontFace: F, align: "center",
  });

  var dermZones = [
    { label: "C5", sub: "상완 외측", color: "D1FAE5" },
    { label: "C6", sub: "엄지·검지", color: "6EE7B7" },
    { label: "C7", sub: "중지", color: C.accentBrand },
    { label: "C8", sub: "약지·소지", color: C.accentDanger },
    { label: "T1", sub: "상완 내측", color: "9CA3AF" },
  ];
  dermZones.forEach(function(z, i) {
    var y = 2.18 + i * 0.55;
    s.addShape(pptx.ShapeType.roundRect, {
      x: 0.75, y: y, w: 3.2, h: 0.44,
      fill: { color: z.color }, line: { color: z.color }, rectRadius: 0.1,
    });
    var txtColor = (z.label === "C7" || z.label === "C8") ? C.textInverse : C.textPrimary;
    s.addText(z.label + "  —  " + z.sub, {
      x: 0.85, y: y + 0.05, w: 3.0, h: 0.34,
      fontSize: 13, bold: true, color: txtColor, fontFace: F, valign: "middle",
    });
  });

  // 우측 메커니즘 카드
  var mechanisms = [
    { num: "01", title: "추간판 탈출(HNP)", body: "수핵이 후방 돌출 → 신경근 직접 압박" },
    { num: "02", title: "골극·후관절 압박", body: "퇴행성 변화 → 신경공(Foramen) 협착" },
    { num: "03", title: "염증성 사이토카인", body: "IL-1β, TNF-α → 신경근 과흥분 → 통증·저림" },
  ];
  mechanisms.forEach(function(m, i) {
    var y = 1.55 + i * 1.3;
    s.addShape(pptx.ShapeType.roundRect, {
      x: 4.35, y: y, w: 5.25, h: 1.15,
      fill: { color: C.cardLight }, line: { color: C.cardLight }, rectRadius: 0.22,
    });
    s.addText(m.num, {
      x: 4.55, y: y + 0.1, w: 0.5, h: 0.4,
      fontSize: 18, bold: true, color: C.accentBrand, fontFace: F,
    });
    s.addText(m.title, {
      x: 5.1, y: y + 0.1, w: 4.3, h: 0.4,
      fontSize: 16, bold: true, color: C.textPrimary, fontFace: F,
    });
    s.addText(m.body, {
      x: 4.55, y: y + 0.55, w: 4.85, h: 0.48,
      fontSize: 13, color: C.textMuted, fontFace: F, wrap: true,
    });
  });

  s.addNotes("【발표 멘트】\n경추에서 오는 손저림을 감별하려면 피부분절 지도를 반드시 숙지하셔야 합니다. 엄지와 검지 저림이면 C6, 중지 저림이면 C7, 약지와 소지 저림이면 C8을 먼저 의심합니다. 기전은 추간판 탈출이나 골극에 의한 신경근 직접 압박이 가장 흔하고, 여기에 IL-1β나 TNF-α 같은 염증성 사이토카인이 신경근 과흥분을 유발합니다.\n\n【전환 멘트】\n같은 내측 손가락 저림이라도 경추가 아닌 흉곽출구에서 문제가 생기는 경우도 있습니다.");
})();

// ═══════════════════════════════════════════════════════════════════════════
// 슬라이드 8 — TOS·레이노·DPN (TYPE-B, Dark)
// ═══════════════════════════════════════════════════════════════════════════
(function slide08() {
  const s = pptx.addSlide();
  bg(s, true);
  addBadge(s, "PATHOPHYSIOLOGY", true);
  addSectionTitle(s, "흉곽출구증후군·레이노·DPN 기전", true);

  var cards = [
    {
      title: "흉곽출구증후군(TOS)",
      body: "사각근 삼각 → 완신경총 압박\nC8·T1 → 내측 전완·소지 저림\n경늑골·사각근 비대·불량 자세 유발",
    },
    {
      title: "레이노 현상(Raynaud)",
      body: "한랭·스트레스 → 교감신경 과활성\n지단 혈관 과수축\n창백(허혈) → 청색(정체) → 충혈(재관류)",
    },
    {
      title: "당뇨병성 신경병증(DPN)",
      body: "고혈당 → 소르비톨 축적 + 산화 스트레스\n신경 혈류 장애 → 탈수초\n장갑·양말형(Glove-Stocking) 양측 대칭",
    },
  ];

  cards.forEach(function(c, i) {
    addCard(s, {
      x: 0.6 + i * 3.15, y: 1.6, w: 2.9, h: 3.3,
      isDark: true, title: c.title, body: c.body,
    });
  });

  s.addNotes("【발표 멘트】\n흉곽출구증후군은 사각근 삼각이라는 좁은 통로에서 완신경총과 쇄골하동맥이 압박되는 질환입니다. 특히 신경성 TOS에서는 하완신경총의 C8·T1 분지가 눌리면서 내측 전완과 소지 쪽 저림이 나타납니다. 레이노 현상은 한랭 자극에 의해 손가락 혈관이 과수축되면서 창백-청색-충혈의 3단계 색 변화를 보이는 것이 특징입니다. 당뇨병성 말초신경병증은 고혈당에 의한 대사성 신경 손상으로, 장갑·양말형의 양측 대칭 분포가 감별 포인트입니다.\n\n【전환 멘트】\n원인 질환을 이해했으니, 이제 외래에서 실제로 어떻게 감별하는지 이학적 검사를 정리하겠습니다.");
})();

// ═══════════════════════════════════════════════════════════════════════════
// 슬라이드 9 — 병력 청취 (TYPE-F, Light)
// ═══════════════════════════════════════════════════════════════════════════
(function slide09() {
  const s = pptx.addSlide();
  bg(s, false);
  addBadge(s, "DIAGNOSIS", false);
  addSectionTitle(s, "병력 청취 — 6가지 핵심 항목", false);

  vline(s, 0.6, 1.55, 0.75, C.accentBrand);
  s.addText("체계적 병력 청취만으로 원인의 범위를 크게 좁힐 수 있다", {
    x: 0.85, y: 1.55, w: 8.55, h: 0.75,
    fontSize: 17, color: C.textPrimary, fontFace: F,
    italic: true, valign: "middle", wrap: true,
  });

  var items = [
    { num: "1", q: "저림 분포", ans: "손 전체 vs 특정 손가락 → 신경지배 영역 확인" },
    { num: "2", q: "야간 악화", ans: "CTS의 특징적 소견 (손목 굴곡위 수면)" },
    { num: "3", q: "경부 통증 동반", ans: "경추 병변 시사" },
    { num: "4", q: "양측성 여부", ans: "중추성, DPN, TOS 가능성" },
    { num: "5", q: "색 변화", ans: "레이노 현상 시사 (창백→청색→충혈)" },
    { num: "6", q: "직업·취미", ans: "반복 손목 사용, 진동 노출" },
  ];

  items.forEach(function(item, i) {
    var col = i < 3 ? 0 : 1;
    var row = i % 3;
    var x = 0.6 + col * 4.75;
    var y = 2.55 + row * 0.82;
    s.addShape(pptx.ShapeType.roundRect, {
      x: x, y: y, w: 4.55, h: 0.72,
      fill: { color: C.cardLight }, line: { color: C.cardLight }, rectRadius: 0.18,
    });
    s.addText(item.num + ". " + item.q, {
      x: x + 0.15, y: y + 0.07, w: 1.7, h: 0.28,
      fontSize: 13, bold: true, color: C.textPrimary, fontFace: F,
    });
    s.addText(item.ans, {
      x: x + 0.15, y: y + 0.36, w: 4.2, h: 0.28,
      fontSize: 12, color: C.textMuted, fontFace: F,
    });
  });

  s.addNotes("【발표 멘트】\n감별 진단의 첫 단계는 역시 병력 청취입니다. 여섯 가지만 확인하면 원인의 범위를 상당히 좁힐 수 있습니다. 첫째, 어떤 손가락이 저린지. 엄지~넷째 내측이면 정중신경, 소지 쪽이면 척골신경이나 C8 문제를 생각합니다. 둘째, 야간에 악화되는지. 야간 저림은 CTS의 매우 특징적인 소견입니다. 셋째, 목 통증이 동반되는지. 넷째, 양쪽 다 저린지. 다섯째, 색 변화가 있는지. 여섯째, 반복적인 손목 사용 직업인지.\n\n【전환 멘트】\n병력 청취 다음에는 이학적 검사로 확인해야 합니다.");
})();

// ═══════════════════════════════════════════════════════════════════════════
// 슬라이드 10 — 이학적 검사 (TYPE-H, Dark)
// ═══════════════════════════════════════════════════════════════════════════
(function slide10() {
  const s = pptx.addSlide();
  bg(s, true);
  addBadge(s, "PHYSICAL EXAM", true);
  addSectionTitle(s, "이학적 검사 — Phalen · Tinel · Spurling", true);

  var tests = [
    {
      name: "Phalen Test", target: "CTS",
      method: "손목 최대 굴곡 90°\n1분 유지",
      positive: "정중신경 영역 저림 재현",
      sens: "68~80%", spec: "59~86%", color: C.accentSuccess,
    },
    {
      name: "Tinel Sign", target: "CTS",
      method: "수근관 위 정중신경 타진",
      positive: "전기 자극감 방사",
      sens: "50~72%", spec: "55~87%", color: C.accentSuccess,
    },
    {
      name: "Spurling Test", target: "경추 신경근병증",
      method: "환측 측방 굴곡 + 신전\n+ 축성 압박",
      positive: "동측 상지 방사통·저림 재현",
      sens: "40~60%", spec: "92~100%", color: C.accentDanger,
    },
  ];

  tests.forEach(function(t, i) {
    var x = 0.6 + i * 3.15;
    s.addShape(pptx.ShapeType.roundRect, {
      x: x, y: 1.55, w: 2.95, h: 3.6,
      fill: { color: C.cardDark }, line: { color: C.cardDark }, rectRadius: 0.22,
    });
    s.addText(t.name, {
      x: x + 0.15, y: 1.68, w: 2.65, h: 0.42,
      fontSize: 18, bold: true, color: t.color, fontFace: F,
    });
    s.addShape(pptx.ShapeType.roundRect, {
      x: x + 0.15, y: 2.12, w: 2.65, h: 0.28,
      fill: { color: C.badgeBg }, line: { color: C.badgeBg }, rectRadius: 0.1,
    });
    s.addText(t.target, {
      x: x + 0.15, y: 2.12, w: 2.65, h: 0.28,
      fontSize: 11, color: C.textInverse, fontFace: F, align: "center", valign: "middle",
    });
    s.addText("방법", {
      x: x + 0.15, y: 2.52, w: 2.65, h: 0.25,
      fontSize: 11, color: C.textMuted, fontFace: F,
    });
    s.addText(t.method, {
      x: x + 0.15, y: 2.77, w: 2.65, h: 0.45,
      fontSize: 13, color: C.textInverse, fontFace: F, wrap: true,
    });
    s.addText("양성 소견", {
      x: x + 0.15, y: 3.25, w: 2.65, h: 0.25,
      fontSize: 11, color: C.textMuted, fontFace: F,
    });
    s.addText(t.positive, {
      x: x + 0.15, y: 3.5, w: 2.65, h: 0.38,
      fontSize: 13, color: C.textInverse, fontFace: F, wrap: true,
    });
    s.addText("민감도 " + t.sens + "  |  특이도 " + t.spec, {
      x: x + 0.1, y: 4.6, w: 2.75, h: 0.3,
      fontSize: 11, color: t.color, fontFace: F, align: "center",
    });
  });

  s.addNotes("【발표 멘트】\n외래에서 바로 시행할 수 있는 이학적 검사 세 가지입니다. 먼저 Phalen Test, 손목을 최대 굴곡해서 1분간 유지하면 정중신경 영역에 저림이 유발됩니다. 민감도가 68~80%로 상당히 쓸 만합니다. Tinel Sign은 수근관 위를 타진해서 전기 자극감이 방사되는지 확인합니다. 경추 쪽을 볼 때는 Spurling Test가 핵심입니다. 특이도가 92~100%로 매우 높아, 양성이면 경추 신경근병증일 가능성이 매우 높습니다.\n\n【전환 멘트】\n이제 이 검사 결과들을 종합해서 감별 포인트를 한 눈에 정리해 보겠습니다.");
})();

// ═══════════════════════════════════════════════════════════════════════════
// 슬라이드 11 — 감별 진단 요약표 (TYPE-F, Light)
// ═══════════════════════════════════════════════════════════════════════════
(function slide11() {
  const s = pptx.addSlide();
  bg(s, false);
  addBadge(s, "DIFFERENTIAL DX", false);

  vline(s, 0.6, 0.72, 0.7, C.accentBrand);
  s.addText("감별 진단 요약 — 6가지 질환 체크리스트", {
    x: 0.85, y: 0.72, w: 8.55, h: 0.7,
    fontSize: 30, bold: true, color: C.textPrimary, fontFace: F, valign: "middle",
  });

  // 헤더 행
  s.addShape(pptx.ShapeType.rect, {
    x: 0.6, y: 1.55, w: 8.8, h: 0.3,
    fill: { color: C.accentBrand }, line: { color: C.accentBrand },
  });
  ["질환", "특징적 소견", "핵심 검사"].forEach(function(h, i) {
    var xs = [0.7, 3.4, 7.5];
    s.addText(h, {
      x: xs[i], y: 1.59, w: 2.5, h: 0.22,
      fontSize: 11, bold: true, color: C.textInverse, fontFace: F,
    });
  });

  var rows = [
    { disease: "수근관증후군(CTS)", feature: "야간 저림, 엄지~4번째, Phalen (+)", test: "NCS/EMG", color: C.accentSuccess },
    { disease: "경추 신경근병증", feature: "경부통, 분절별 분포, Spurling (+)", test: "MRI 경추", color: C.accentBrand },
    { disease: "흉곽출구증후군(TOS)", feature: "팔 거상 시 악화, Adson (+)", test: "혈관초음파", color: C.accentBrand },
    { disease: "당뇨병성 신경병증(DPN)", feature: "양측 대칭, 당뇨 병력, 장갑형", test: "NCS + HbA1c", color: C.textMuted },
    { disease: "레이노 현상", feature: "한랭 유발, 3상 색 변화", test: "냉각 유발 검사", color: C.textMuted },
    { disease: "뇌졸중 (즉시 의뢰)", feature: "편측 전체 저림 + 신경학적 징후", test: "뇌 MRI/CT", color: C.accentDanger },
  ];

  rows.forEach(function(r, i) {
    var y = 1.87 + i * 0.6;
    var bgCol = i % 2 === 0 ? C.cardLight : "F0EBE0";
    s.addShape(pptx.ShapeType.rect, {
      x: 0.6, y: y, w: 8.8, h: 0.55,
      fill: { color: bgCol }, line: { color: bgCol },
    });
    s.addText(r.disease, {
      x: 0.7, y: y + 0.09, w: 2.6, h: 0.36,
      fontSize: 12, bold: true, color: r.color, fontFace: F,
    });
    s.addText(r.feature, {
      x: 3.4, y: y + 0.09, w: 4.0, h: 0.36,
      fontSize: 12, color: C.textPrimary, fontFace: F,
    });
    s.addText(r.test, {
      x: 7.5, y: y + 0.09, w: 1.8, h: 0.36,
      fontSize: 12, color: C.textMuted, fontFace: F,
    });
  });

  s.addNotes("【발표 멘트】\n이 표를 한 번 머릿속에 넣어두시면 외래에서 큰 도움이 됩니다. CTS는 야간 저림이 키워드이고, 경추 신경근병증은 경부통과 분절별 분포가 핵심입니다. TOS는 팔을 올렸을 때 악화되는 패턴, DPN은 양측 대칭의 장갑형 분포, 레이노는 한랭에 의한 3상 색 변화가 감별 포인트입니다. 가장 중요한 것은 편측 전체 저림에 다른 신경학적 징후가 동반되면 뇌졸중을 의심해야 한다는 점입니다.\n\n【전환 멘트】\n감별이 끝났으면, 이제 우리 한의사가 가장 궁금한 부분, 치료 근거로 넘어가겠습니다.");
})();

// ═══════════════════════════════════════════════════════════════════════════
// 슬라이드 12 — 침치료 근거: CTS (TYPE-G, Dark)
// ═══════════════════════════════════════════════════════════════════════════
(function slide12() {
  const s = pptx.addSlide();
  bg(s, true);
  addBadge(s, "EVIDENCE", true);
  addSectionTitle(s, "침치료(Acupuncture) — CTS 메타분석 근거", true);

  s.addText("622", {
    x: 0.6, y: 1.55, w: 3.0, h: 1.1,
    fontSize: 90, bold: true, color: C.accentSuccess, fontFace: F,
    align: "center", valign: "middle",
  });
  s.addText("명 분석 (12개 RCT)", {
    x: 0.6, y: 2.65, w: 3.0, h: 0.35,
    fontSize: 13, color: C.textMuted, fontFace: F, align: "center",
  });
  s.addText("Yao et al. (2020)", {
    x: 0.6, y: 2.95, w: 3.0, h: 0.3,
    fontSize: 11, color: C.textMuted, fontFace: F, align: "center",
  });

  var evidCards = [
    { title: "VAS 통증 유의 개선", body: "MD −1.36\n(95% CI −2.12 ~ −0.61)\n침치료 vs 위약 비교" },
    { title: "전기침 병합 효과 우수", body: "단독 침치료보다\nElectroacupuncture 병합 시\n더 우수한 결과" },
    { title: "근거 수준 I~II", body: "경증~중등도 CTS 유효\n부목·스테로이드 주사와\n동등 수준" },
  ];
  evidCards.forEach(function(ec, i) {
    addCard(s, {
      x: 3.8 + i * 2.1, y: 1.55, w: 1.95, h: 2.8,
      isDark: true, title: ec.title, body: ec.body,
    });
  });

  s.addShape(pptx.ShapeType.roundRect, {
    x: 0.6, y: 4.55, w: 8.8, h: 0.72,
    fill: { color: "0A3020" }, line: { color: C.accentSuccess, pt: 1.5 }, rectRadius: 0.22,
  });
  s.addText("임상 적용: 경증~중등도 CTS 환자에게 침치료를 1차 옵션으로 제안 가능", {
    x: 0.8, y: 4.62, w: 8.4, h: 0.55,
    fontSize: 15, bold: true, color: C.accentSuccess, fontFace: F,
    align: "center", valign: "middle",
  });

  s.addNotes("【발표 멘트】\n한의학 치료 근거, 먼저 수근관증후군에 대한 침치료입니다. Yao 등이 2020년에 발표한 메타분석에서 12개 RCT, 총 622명을 분석한 결과, 침치료는 위약 대비 VAS 통증 점수를 유의하게 개선했습니다. 평균 차이가 -1.36점이고, 신뢰구간이 영점을 포함하지 않습니다. 전기침을 병합하면 효과가 더 좋았습니다. 근거 수준으로 보면 부목이나 스테로이드 주사와 비슷한 I~II 등급입니다.\n\n【전환 멘트】\n침치료가 왜 효과가 있는지, 신경과학적 기전도 밝혀져 있습니다.");
})();

// ═══════════════════════════════════════════════════════════════════════════
// 슬라이드 13 — 침치료 기전 (TYPE-F, Light)
// ═══════════════════════════════════════════════════════════════════════════
(function slide13() {
  const s = pptx.addSlide();
  bg(s, false);
  addBadge(s, "MECHANISM", false);
  addSectionTitle(s, "침치료 기전 — 중추 가소성과 혈위 근거", false);

  vline(s, 0.6, 1.58, 0.85, C.accentBrand);
  s.addText("Napadow et al. (2007, Pain) — 하버드 의대 fMRI\n침치료 후 체성감각피질(Somatosensory Cortex) 재조직화 확인 → 중추 가소성(Central Plasticity)", {
    x: 0.85, y: 1.58, w: 8.55, h: 0.85,
    fontSize: 14, color: C.textPrimary, fontFace: F, italic: true, wrap: true, valign: "middle",
  });

  var acupoints = [
    { name: "大陵 (PC7)", anatomy: "수근관 직상부 정중신경 경로", effect: "국소 자극·염증 억제" },
    { name: "內關 (PC6)", anatomy: "전완 굴근 사이 정중신경", effect: "통증 조절·Gate Control" },
    { name: "合谷 (LI4)", anatomy: "대측 감각 피질 활성화", effect: "Gate Control 기반 진통" },
    { name: "八邪 (EX-UE9)", anatomy: "손가락 사이 말초신경", effect: "혈류 개선·저림 완화" },
  ];
  acupoints.forEach(function(ap, i) {
    var x = 0.6 + i * 2.4;
    s.addShape(pptx.ShapeType.roundRect, {
      x: x, y: 2.65, w: 2.2, h: 2.5,
      fill: { color: C.cardLight }, line: { color: C.cardLight }, rectRadius: 0.22,
    });
    s.addShape(pptx.ShapeType.roundRect, {
      x: x, y: 2.65, w: 2.2, h: 0.55,
      fill: { color: C.accentBrand }, line: { color: C.accentBrand }, rectRadius: 0.22,
    });
    s.addText(ap.name, {
      x: x + 0.1, y: 2.7, w: 2.0, h: 0.44,
      fontSize: 15, bold: true, color: C.textInverse, fontFace: F, align: "center", valign: "middle",
    });
    s.addText(ap.anatomy, {
      x: x + 0.1, y: 3.28, w: 2.0, h: 0.65,
      fontSize: 12, color: C.textPrimary, fontFace: F, wrap: true,
    });
    s.addText(ap.effect, {
      x: x + 0.1, y: 3.98, w: 2.0, h: 0.55,
      fontSize: 12, color: C.textMuted, fontFace: F, wrap: true,
    });
  });

  s.addText("Herman et al. (2018, Brain): 로컬 침 + 원위 침 모두 피질 재지도화 유도 확인", {
    x: 0.6, y: 5.32, w: 8.8, h: 0.25,
    fontSize: 11, color: C.textMuted, fontFace: F, italic: true,
  });

  s.addNotes("【발표 멘트】\n침치료의 기전에 대해 가장 흥미로운 연구는 하버드 의대 Napadow 팀의 fMRI 연구입니다. CTS 환자에게 침치료 후 체성감각피질의 재조직화, 즉 중추 가소성이 확인되었습니다. 침이 단순히 말초에서만 작용하는 게 아니라 뇌의 감각 지도를 바꾼다는 뜻입니다. 혈위별로 보면, 大陵(PC7)은 수근관 직상부에서 정중신경을 직접 자극하고, 合谷(LI4)은 대측 감각 피질을 활성화합니다.\n\n【전환 멘트】\n경추 신경근병증에 대한 침치료 근거도 살펴보겠습니다.");
})();

// ═══════════════════════════════════════════════════════════════════════════
// 슬라이드 14 — 경추 침치료 + 추나 (TYPE-E, Dark)
// ═══════════════════════════════════════════════════════════════════════════
(function slide14() {
  const s = pptx.addSlide();
  bg(s, true);
  addBadge(s, "EVIDENCE", true);
  addSectionTitle(s, "경추 신경근병증 — 침치료 + 추나 근거", true);

  s.addShape(pptx.ShapeType.roundRect, {
    x: 0.6, y: 1.6, w: 4.5, h: 3.3,
    fill: { color: C.cardDark }, line: { color: C.cardDark }, rectRadius: 0.22,
  });
  s.addText("침치료 근거", {
    x: 0.8, y: 1.72, w: 4.1, h: 0.42,
    fontSize: 20, bold: true, color: C.accentSuccess, fontFace: F,
  });
  s.addText("Liu et al. (2019, Spine)\n14개 RCT 분석\n\n• 침치료 + 추나/물리치료 병합 시 VAS 유의 개선\n• 단독 침치료도 4~8주 통증·저림 개선에 유효\n• 주요 혈위: 頸夾脊(Ex-B2), 風池(GB20),\n  後溪(SI3), 懸鍾(GB39)", {
    x: 0.8, y: 2.2, w: 4.1, h: 2.55,
    fontSize: 14, color: C.textInverse, fontFace: F, wrap: true,
  });

  s.addText("+", {
    x: 5.15, y: 2.85, w: 0.5, h: 0.55,
    fontSize: 36, bold: true, color: C.accentSuccess, fontFace: F,
    align: "center", valign: "middle",
  });

  s.addShape(pptx.ShapeType.roundRect, {
    x: 5.75, y: 1.6, w: 3.85, h: 3.3,
    fill: { color: C.cardDark }, line: { color: C.cardDark }, rectRadius: 0.22,
  });
  s.addText("추나(Chuna) 근거", {
    x: 5.95, y: 1.72, w: 3.45, h: 0.42,
    fontSize: 20, bold: true, color: C.accentSuccess, fontFace: F,
  });
  s.addText("Gross et al. (2015, Cochrane)\n도수치료 > 단독 운동요법\n(단기 통증 개선)\n\n대한추나의학회 임상진료지침\n(2021): 급성기 이후 B등급 권고", {
    x: 5.95, y: 2.2, w: 3.45, h: 2.55,
    fontSize: 14, color: C.textInverse, fontFace: F, wrap: true,
  });

  s.addNotes("【발표 멘트】\n경추 신경근병증에 대해서는 Liu 등의 Spine 논문에서 14개 RCT를 분석했는데, 침치료와 추나 또는 물리치료를 병합하면 VAS 점수가 유의하게 개선되었습니다. 단독 침치료도 4~8주의 단기간 치료에서 효과가 확인되었습니다. 주요 혈위는 頸夾脊, 風池, 後溪, 懸鍾입니다. 추나의 경우, Cochrane Review에서 경추 도수치료가 단독 운동요법보다 우수한 것으로 나왔고, 대한추나의학회 진료지침에서도 급성기 이후 B등급으로 권고하고 있습니다.\n\n【전환 멘트】\n다음으로, 한약 치료의 근거를 살펴보겠습니다.");
})();

// ═══════════════════════════════════════════════════════════════════════════
// 슬라이드 15 — 한약 치료 근거 (TYPE-F, Light)
// ═══════════════════════════════════════════════════════════════════════════
(function slide15() {
  const s = pptx.addSlide();
  bg(s, false);
  addBadge(s, "HERBAL MEDICINE", false);

  vline(s, 0.6, 0.72, 0.7, C.accentBrand);
  s.addText("한약 치료 근거 — 당귀사역탕 · 황기계지오물탕 · 우차신기환", {
    x: 0.85, y: 0.72, w: 8.55, h: 0.7,
    fontSize: 26, bold: true, color: C.textPrimary, fontFace: F, valign: "middle",
  });

  var herbCards = [
    {
      name: "당귀사역탕\n(當歸四逆湯)",
      composition: "당귀·계지·작약·세신·통초·대조·감초",
      indication: "氣虛血滯, 사지 냉증·저림",
      evidence: "Kim et al. (2018, J Ethnopharmacol)\n말초 혈류 개선·혈관 수축 억제",
      apply: "레이노 현상, 한증 손저림",
      color: C.accentBrand,
    },
    {
      name: "황기계지오물탕\n(黃芪桂枝五物湯)",
      composition: "황기·계지·작약·생강·대조",
      indication: "血痹(혈비) — 혈허·영위불화",
      evidence: "Liu et al. (2020, Front Pharmacol)\nNGF 상향 조절 → 말초신경 재생",
      apply: "기허혈체 상지 저림",
      color: C.accentSuccess,
    },
    {
      name: "우차신기환\n(牛車腎氣丸)",
      composition: "육미지황 + 우슬·차전자·부자·육계",
      indication: "신양허 → 당뇨병성 신경병증",
      evidence: "Watanabe et al. (2014)\nNGF 증가, 산화 스트레스 경감",
      apply: "일본 가이드라인 1차 선택",
      color: C.accentDanger,
    },
  ];

  herbCards.forEach(function(hc, i) {
    var x = 0.6 + i * 3.15;
    s.addShape(pptx.ShapeType.roundRect, {
      x: x, y: 1.6, w: 2.95, h: 3.5,
      fill: { color: C.cardLight }, line: { color: hc.color, pt: 1.5 }, rectRadius: 0.22,
    });
    s.addText(hc.name, {
      x: x + 0.15, y: 1.72, w: 2.65, h: 0.65,
      fontSize: 16, bold: true, color: hc.color, fontFace: F,
    });
    s.addText("구성: " + hc.composition, {
      x: x + 0.15, y: 2.42, w: 2.65, h: 0.5,
      fontSize: 11, color: C.textMuted, fontFace: F, wrap: true,
    });
    s.addText("적응: " + hc.indication, {
      x: x + 0.15, y: 2.96, w: 2.65, h: 0.42,
      fontSize: 12, color: C.textPrimary, fontFace: F, wrap: true,
    });
    s.addText(hc.evidence, {
      x: x + 0.15, y: 3.42, w: 2.65, h: 0.75,
      fontSize: 11, color: C.textMuted, fontFace: F, wrap: true, italic: true,
    });
    s.addText("→ " + hc.apply, {
      x: x + 0.15, y: 4.22, w: 2.65, h: 0.3,
      fontSize: 12, bold: true, color: hc.color, fontFace: F,
    });
  });

  s.addNotes("【발표 멘트】\n한약 치료는 병태에 따라 세 가지 처방을 기억하시면 됩니다. 첫째, 당귀사역탕. 傷寒論 351조에 근거한 처방으로, Kim 등의 J Ethnopharmacol 논문에서 말초 혈류 개선과 혈관 수축 억제 효과가 확인되었습니다. 레이노 현상이나 한증 손저림에 적합합니다. 둘째, 황기계지오물탕. 금궤요략 혈비편의 처방인데, NGF를 상향 조절하여 말초신경 재생을 촉진합니다. 셋째, 우차신기환. 일본 임상 가이드라인에서 당뇨병성 말초신경병증의 1차 선택 한약으로 수록되어 있습니다.\n\n【전환 멘트】\n이제 이러한 치료 근거들을 실제 외래 진료에 어떻게 적용할 것인지, 변증 체계를 정리해 보겠습니다.");
})();

// ═══════════════════════════════════════════════════════════════════════════
// 슬라이드 16 — 변증별 치료 가이드 (TYPE-B, Dark)
// ═══════════════════════════════════════════════════════════════════════════
(function slide16() {
  const s = pptx.addSlide();
  bg(s, true);
  addBadge(s, "CLINICAL GUIDE", true);
  addSectionTitle(s, "변증별(辨證別) 치료 가이드", true);

  var syndromes = [
    { code: "01", name: "氣虛血滯", feature: "무기력·창백·둔한 저림", method: "益氣活血", herb: "황기계지오물탕" },
    { code: "02", name: "寒凝血脈", feature: "냉각 시 악화·색 변화", method: "溫陽散寒", herb: "당귀사역탕·부자탕" },
    { code: "03", name: "瘀血阻絡", feature: "고정 통증·야간 악화·자반", method: "活血化瘀", herb: "혈부축어탕" },
    { code: "04", name: "痰濕阻絡", feature: "무겁고 부은 저림", method: "化痰除濕", herb: "이진탕 합 도담탕" },
    { code: "05", name: "肝腎虧虛", feature: "고령·근위축·반복 재발", method: "補益肝腎", herb: "독활기생탕·우차신기환" },
  ];

  syndromes.forEach(function(syn, i) {
    var y = 1.6 + i * 0.78;
    s.addShape(pptx.ShapeType.roundRect, {
      x: 0.6, y: y, w: 9.0, h: 0.68,
      fill: { color: C.cardDark }, line: { color: C.cardDark }, rectRadius: 0.18,
    });
    s.addText(syn.code, {
      x: 0.75, y: y + 0.1, w: 0.4, h: 0.45,
      fontSize: 18, bold: true, color: C.accentSuccess, fontFace: F, align: "center",
    });
    s.addText(syn.name, {
      x: 1.25, y: y + 0.1, w: 2.0, h: 0.45,
      fontSize: 16, bold: true, color: C.textInverse, fontFace: F,
    });
    s.addText(syn.feature, {
      x: 3.35, y: y + 0.12, w: 2.5, h: 0.42,
      fontSize: 13, color: C.textMuted, fontFace: F,
    });
    s.addText(syn.method, {
      x: 5.95, y: y + 0.12, w: 1.6, h: 0.42,
      fontSize: 13, color: C.accentSuccess, fontFace: F, bold: true,
    });
    s.addText(syn.herb, {
      x: 7.65, y: y + 0.12, w: 1.85, h: 0.42,
      fontSize: 12, color: C.textInverse, fontFace: F,
    });
  });

  s.addNotes("【발표 멘트】\n실제 외래에서 손저림 환자를 변증할 때, 이 다섯 가지 패턴으로 분류하시면 치료 방향이 명확해집니다. 기허혈체는 무기력하고 안색이 창백한 환자에서 둔한 저림이 특징이고, 황기계지오물탕으로 익기활혈합니다. 한응혈맥은 추울 때 악화되고 색 변화가 동반되면 당귀사역탕으로 온양산한합니다. 어혈조락은 고정 통증에 야간 악화가 있을 때, 혈부축어탕으로 활혈화어합니다. 담습조락은 부은 느낌과 저림이 함께 오면 이진탕 합 도담탕을 고려하고, 간신휴허는 고령에서 반복 재발하는 패턴이면 독활기생탕이나 우차신기환을 씁니다.\n\n【전환 멘트】\n변증을 정했으면, 이제 구체적인 치료 알고리즘을 따라가 보겠습니다.");
})();

// ═══════════════════════════════════════════════════════════════════════════
// 슬라이드 17 — 치료 알고리즘 + Red Flags (TYPE-H, Light)
// ═══════════════════════════════════════════════════════════════════════════
(function slide17() {
  const s = pptx.addSlide();
  bg(s, false);
  addBadge(s, "ALGORITHM", false);
  addSectionTitle(s, "치료 알고리즘 + Red Flags", false);

  var steps = [
    { num: "1", text: "문진·이학적 검사 — Phalen·Tinel·Spurling" },
    { num: "2", text: "변증 결정 + 영상 판독 (필요 시 X-ray, MRI 의뢰)" },
    { num: "3", text: "경증 CTS: PC7·PC6·LI4·八邪 침 + 황기계지오물탕 4~6주" },
    { num: "4", text: "경추 유래: 頸夾脊·風池·後溪 + 추나 병합 6~8주" },
    { num: "5", text: "4~6주 재평가 → 미흡 시 양방 협진" },
  ];

  steps.forEach(function(step, i) {
    var y = 1.55 + i * 0.72;
    s.addShape(pptx.ShapeType.roundRect, {
      x: 0.6, y: y, w: 5.4, h: 0.62,
      fill: { color: C.cardLight }, line: { color: C.cardLight }, rectRadius: 0.18,
    });
    s.addShape(pptx.ShapeType.roundRect, {
      x: 0.6, y: y, w: 0.42, h: 0.62,
      fill: { color: C.accentBrand }, line: { color: C.accentBrand }, rectRadius: 0.18,
    });
    s.addText(step.num, {
      x: 0.6, y: y, w: 0.42, h: 0.62,
      fontSize: 14, bold: true, color: C.textInverse, fontFace: F,
      align: "center", valign: "middle",
    });
    s.addText(step.text, {
      x: 1.1, y: y + 0.05, w: 4.8, h: 0.52,
      fontSize: 12, color: C.textPrimary, fontFace: F, wrap: true,
    });
    if (i < steps.length - 1) {
      s.addText("v", {
        x: 1.2, y: y + 0.62, w: 0.3, h: 0.1,
        fontSize: 10, color: C.textMuted, fontFace: F, align: "center",
      });
    }
  });

  // Red Flags 박스
  s.addShape(pptx.ShapeType.roundRect, {
    x: 6.2, y: 1.55, w: 3.4, h: 3.5,
    fill: { color: "FFF0F0" }, line: { color: C.accentDanger, pt: 2 }, rectRadius: 0.22,
  });
  s.addText("Red Flags — 즉시 의뢰", {
    x: 6.35, y: 1.68, w: 3.1, h: 0.38,
    fontSize: 14, bold: true, color: C.accentDanger, fontFace: F,
  });
  var redFlags = [
    "편측 상하지 동시 저림\n→ 뇌졸중",
    "양측 저림 + 보행 장애\n→ 척수병증",
    "진행성 근위축·근력 저하\n→ ALS 의심",
    "발열 + 경부 강직\n→ 감염성 병변",
  ];
  redFlags.forEach(function(rf, i) {
    s.addText("• " + rf, {
      x: 6.35, y: 2.15 + i * 0.72, w: 3.1, h: 0.62,
      fontSize: 12, color: C.accentDanger, fontFace: F, wrap: true,
    });
  });

  s.addNotes("【발표 멘트】\n실제 외래 알고리즘입니다. 첫 단계는 문진과 이학적 검사로 원인을 감별하고, 필요하면 영상 의뢰를 합니다. 변증을 결정한 뒤 경증 CTS면 침치료에 황기계지오물탕을 4~6주 처방합니다. 경추 유래면 頸夾脊 침에 추나를 병합해서 6~8주 치료합니다. 4~6주 후 반드시 재평가하고, 개선이 미흡하면 양방 협진을 고려합니다. Red Flags: 갑작스러운 편측 상하지 동시 저림은 뇌졸중, 양측 저림에 보행 장애가 동반되면 척수병증, 진행성 근위축은 ALS를 의심해야 합니다.\n\n【전환 멘트】\n마지막으로 환자 교육과 예방 지도도 빠뜨릴 수 없습니다.");
})();

// ═══════════════════════════════════════════════════════════════════════════
// 슬라이드 18 — 예방 및 생활 지도 (TYPE-D, Light)
// ═══════════════════════════════════════════════════════════════════════════
(function slide18() {
  const s = pptx.addSlide();
  bg(s, false);
  addBadge(s, "PREVENTION", false);
  addSectionTitle(s, "예방 및 생활 지도 — 원인별 맞춤 가이드", false);

  var guides = [
    {
      disease: "CTS 예방",
      items: ["손목 중립 자세 유지\n(컴퓨터 사용 시 손목 받침대)", "야간 손목 부목 착용"],
      color: C.accentBrand,
    },
    {
      disease: "경추 신경근병증",
      items: ["경부 스트레칭·강화 운동", "자세 교정\n(스마트폰·모니터 높이)"],
      color: C.accentSuccess,
    },
    {
      disease: "레이노 현상",
      items: ["한랭 회피", "방한 장갑 착용"],
      color: C.accentBrand,
    },
    {
      disease: "당뇨병성 신경병증",
      items: ["혈당 조절\n(HbA1c 목표치 관리)", "정기적 신경전도 검사"],
      color: C.accentDanger,
    },
  ];

  guides.forEach(function(g, i) {
    var x = 0.6 + i * 2.4;
    s.addShape(pptx.ShapeType.roundRect, {
      x: x, y: 1.58, w: 2.2, h: 3.3,
      fill: { color: C.cardLight }, line: { color: C.cardLight }, rectRadius: 0.22,
    });
    s.addShape(pptx.ShapeType.roundRect, {
      x: x, y: 1.58, w: 2.2, h: 0.5,
      fill: { color: g.color }, line: { color: g.color }, rectRadius: 0.22,
    });
    s.addText(g.disease, {
      x: x + 0.08, y: 1.62, w: 2.04, h: 0.42,
      fontSize: 13, bold: true, color: C.textInverse, fontFace: F, align: "center", valign: "middle",
    });
    g.items.forEach(function(item, j) {
      s.addText("• " + item, {
        x: x + 0.12, y: 2.22 + j * 1.1, w: 1.96, h: 0.95,
        fontSize: 13, color: C.textPrimary, fontFace: F, wrap: true, valign: "top",
      });
    });
  });

  s.addShape(pptx.ShapeType.roundRect, {
    x: 0.6, y: 5.07, w: 8.8, h: 0.42,
    fill: { color: C.cardLight },
    line: { color: C.accentBrand, pt: 1.5 }, rectRadius: 0.18,
  });
  s.addText("치료와 함께 생활 지도 병행 — 재발 방지의 핵심", {
    x: 0.8, y: 5.1, w: 8.4, h: 0.35,
    fontSize: 14, bold: true, color: C.accentBrand, fontFace: F, align: "center", valign: "middle",
  });

  s.addNotes("【발표 멘트】\n치료만큼 중요한 것이 예방과 생활 지도입니다. CTS 환자에게는 컴퓨터 사용 시 손목 중립 자세를 유지하도록 교육하고, 야간 손목 부목 착용을 권합니다. 경추 문제가 있는 환자에게는 경부 스트레칭과 강화 운동을 알려드리고, 스마트폰이나 모니터 높이를 교정하도록 안내합니다. 레이노 현상 환자에게는 한랭 회피와 방한 장갑 착용이 필수입니다. 당뇨병성 신경병증 환자에게는 무엇보다 혈당 조절의 중요성을 강조해야 합니다.\n\n【전환 멘트】\n이제 오늘 강의의 핵심을 정리하겠습니다.");
})();

// ═══════════════════════════════════════════════════════════════════════════
// 슬라이드 19 — 핵심 요약 (TYPE-B, Dark)
// ═══════════════════════════════════════════════════════════════════════════
(function slide19() {
  const s = pptx.addSlide();
  bg(s, true);
  addBadge(s, "SUMMARY", true);
  addSectionTitle(s, "손저림 진료 — 5단계 핵심 정리", true);

  var summaryItems = [
    { num: "01", title: "분류", body: "말초신경 포착 / 경추 신경근병증\n/ TOS / 혈관성 / 대사성 / 중추성" },
    { num: "02", title: "감별", body: "병력 청취 6가지\n+ 이학적 검사 3가지\n(Phalen, Tinel, Spurling)" },
    { num: "03", title: "치료 근거", body: "침: CTS 메타분석 Level I~II\n한약: 당귀사역탕·황기계지오물탕\n·우차신기환 / 추나: 경추 B등급" },
    { num: "04", title: "변증 5가지", body: "기허혈체·한응혈맥·어혈조락\n담습조락·간신휴허" },
    { num: "05", title: "Red Flags", body: "편측 동시 저림 → 뇌졸중\n보행 장애 → 척수병증\n근위축 → ALS" },
  ];

  // 상단 3개 + 하단 2개 배치
  summaryItems.slice(0, 3).forEach(function(item, i) {
    var x = 0.6 + i * 3.05;
    s.addShape(pptx.ShapeType.roundRect, {
      x: x, y: 1.65, w: 2.85, h: 1.6,
      fill: { color: C.cardDark }, line: { color: C.cardDark }, rectRadius: 0.22,
    });
    s.addText(item.num, {
      x: x + 0.15, y: 1.75, w: 0.55, h: 0.45,
      fontSize: 22, bold: true, color: C.accentSuccess, fontFace: F,
    });
    s.addText(item.title, {
      x: x + 0.75, y: 1.75, w: 1.95, h: 0.45,
      fontSize: 18, bold: true, color: C.textInverse, fontFace: F,
    });
    s.addText(item.body, {
      x: x + 0.15, y: 2.25, w: 2.55, h: 0.85,
      fontSize: 12, color: C.textMuted, fontFace: F, wrap: true,
    });
  });

  summaryItems.slice(3).forEach(function(item, i) {
    var x = 0.6 + i * 4.6;
    s.addShape(pptx.ShapeType.roundRect, {
      x: x, y: 3.55, w: 4.35, h: 1.6,
      fill: { color: C.cardDark }, line: { color: C.cardDark }, rectRadius: 0.22,
    });
    s.addText(item.num, {
      x: x + 0.15, y: 3.65, w: 0.55, h: 0.45,
      fontSize: 22, bold: true, color: C.accentSuccess, fontFace: F,
    });
    s.addText(item.title, {
      x: x + 0.75, y: 3.65, w: 3.45, h: 0.45,
      fontSize: 18, bold: true, color: C.textInverse, fontFace: F,
    });
    s.addText(item.body, {
      x: x + 0.15, y: 4.15, w: 4.05, h: 0.85,
      fontSize: 12, color: C.textMuted, fontFace: F, wrap: true,
    });
  });

  s.addNotes("【발표 멘트】\n오늘 강의의 핵심을 다섯 가지로 정리합니다. 첫째, 손저림은 여섯 가지 카테고리로 분류합니다. 둘째, 병력 청취 여섯 가지 항목과 이학적 검사 세 가지로 감별합니다. 셋째, 침치료는 CTS에 메타분석 수준의 근거가 있고, 한약은 병태에 따라 당귀사역탕, 황기계지오물탕, 우차신기환을 선별 적용합니다. 넷째, 변증은 다섯 가지 패턴으로 나누어 접근합니다. 다섯째, 편측 동시 저림이나 보행 장애 같은 Red Flags는 절대 놓치지 마십시오.\n\n【전환 멘트】\n마지막으로 참고문헌을 안내드리고 마무리하겠습니다.");
})();

// ═══════════════════════════════════════════════════════════════════════════
// 슬라이드 20 — Q&A / 참고문헌 (TYPE-A, Light)
// ═══════════════════════════════════════════════════════════════════════════
(function slide20() {
  const s = pptx.addSlide();
  bg(s, false);

  s.addText("Q & A", {
    x: 0.6, y: 0.35, w: 5.0, h: 1.0,
    fontSize: 56, bold: true, color: C.accentBrand, fontFace: F,
  });
  s.addText("감사합니다", {
    x: 0.6, y: 1.25, w: 5.0, h: 0.55,
    fontSize: 24, color: C.textMuted, fontFace: F,
  });

  hline(s, 0.6, 1.9, 8.8, C.textMuted);

  s.addText("참고문헌", {
    x: 0.6, y: 2.05, w: 2.0, h: 0.3,
    fontSize: 13, bold: true, color: C.textPrimary, fontFace: F,
  });

  var refs = [
    "1. Atroshi et al. (1999) JAMA — CTS 유병률",
    "2. Szabo & Chidgey (1989) J Hand Surg Am — 수근관 내압",
    "3. Radhakrishnan et al. (1994) Brain — 경추 신경근병증 역학",
    "4. Boulton et al. (2005) Lancet — DPN 역학",
    "5. Yao et al. (2020) Evid Based Complement Alternat Med — 침치료 CTS 메타분석",
    "6. Napadow et al. (2007) Pain — 침치료 fMRI 연구",
    "7. Herman et al. (2018) Brain — 침치료 피질 재조직화",
    "8. Liu et al. (2019) Spine — 경추 침치료 체계적 고찰",
    "9. Kim et al. (2018) J Ethnopharmacol — 당귀사역탕",
    "10. Liu et al. (2020) Front Pharmacol — 황기계지오물탕 NGF",
    "11. Watanabe et al. (2014) J Diabetes Investig — 우차신기환 DPN",
    "12. Gross et al. (2015) Cochrane Review — 경추 도수치료",
    "13. 대한추나의학회 (2021) — 추나 임상진료지침",
  ];

  refs.slice(0, 7).forEach(function(ref, i) {
    s.addText(ref, {
      x: 0.6, y: 2.42 + i * 0.38, w: 4.6, h: 0.34,
      fontSize: 10, color: C.textMuted, fontFace: F,
    });
  });
  refs.slice(7).forEach(function(ref, i) {
    s.addText(ref, {
      x: 5.3, y: 2.42 + i * 0.38, w: 4.3, h: 0.34,
      fontSize: 10, color: C.textMuted, fontFace: F,
    });
  });

  s.addNotes("【발표 멘트】\n이상으로 손저림의 한의학적 이해와 치료에 대한 강의를 마칩니다. 오늘 다룬 내용 중 임상에서 궁금하신 점이나, 실제 환자 사례에 대해 논의하고 싶으신 부분이 있으시면 편하게 질문해 주십시오. 참고문헌은 슬라이드에 정리해 두었으니 추후 원문 확인 시 활용하시기 바랍니다. 경청해 주셔서 감사합니다.\n\n【전환 멘트】\n(강의 종료)");
})();

// ── 파일 저장 ────────────────────────────────────────────────────────────────
pptx.writeFile({ fileName: "lecture-agent-team/output/lecture.pptx" })
  .then(function() { console.log("lecture.pptx 생성 완료"); })
  .catch(function(err) { console.error("오류:", err); process.exit(1); });
