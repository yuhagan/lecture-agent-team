const pptxgen = require("pptxgenjs");
const pptx = new pptxgen();

pptx.layout = "LAYOUT_16x9";
pptx.author = "능골한의원";

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
  // 디자이너 명세 navy 테마
  bgDarkNav:     "1B2A4A",
  cardDarkNav:   "253750",
  accentTeal:    "0A7E8C",
};
const F = "Pretendard";

// ─────────────────────────────────────────────
// 헬퍼 함수
// ─────────────────────────────────────────────

function addBadge(slide, text, isDark) {
  const bgColor = isDark ? C.badgeBgLight : C.badgeBg;
  const txColor = isDark ? C.textPrimary  : C.textInverse;
  const w = Math.max(text.length * 0.13 + 0.6, 1.2);
  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 0.28, w, h: 0.32,
    fill: { color: bgColor }, line: { color: bgColor }, rectRadius: 0.16,
  });
  slide.addText(text, {
    x: 0.6, y: 0.28, w, h: 0.32,
    fontSize: 11, bold: true, color: txColor, fontFace: F,
    align: "center", valign: "middle",
  });
}

function addBadgeCustom(slide, text, bgColor, txColor) {
  const w = Math.max(text.length * 0.13 + 0.6, 1.2);
  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 0.28, w, h: 0.32,
    fill: { color: bgColor }, line: { color: bgColor }, rectRadius: 0.16,
  });
  slide.addText(text, {
    x: 0.6, y: 0.28, w, h: 0.32,
    fontSize: 11, bold: true, color: txColor, fontFace: F,
    align: "center", valign: "middle",
  });
}

function addNavCard(slide, { x, y, w, h, title, body, titleColor, bodyColor, titleSize, bodySize }) {
  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h,
    fill: { color: C.cardDarkNav }, line: { color: C.cardDarkNav }, rectRadius: 0.22,
  });
  if (title) slide.addText(title, {
    x: x + 0.22, y: y + 0.13, w: w - 0.44, h: 0.4,
    fontSize: titleSize || 16, bold: true,
    color: titleColor || C.textInverse, fontFace: F,
  });
  if (body) slide.addText(body, {
    x: x + 0.22, y: y + 0.55, w: w - 0.44, h: h - 0.7,
    fontSize: bodySize || 13, color: bodyColor || C.textInverse, fontFace: F, wrap: true,
  });
}

function addLightCard(slide, { x, y, w, h, title, body, titleColor, bodyColor, titleSize, bodySize }) {
  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h,
    fill: { color: C.cardLight }, line: { color: C.cardLight }, rectRadius: 0.22,
  });
  if (title) slide.addText(title, {
    x: x + 0.2, y: y + 0.13, w: w - 0.4, h: 0.4,
    fontSize: titleSize || 16, bold: true,
    color: titleColor || C.textPrimary, fontFace: F,
  });
  if (body) slide.addText(body, {
    x: x + 0.2, y: y + 0.55, w: w - 0.4, h: h - 0.7,
    fontSize: bodySize || 13, color: bodyColor || C.textPrimary, fontFace: F, wrap: true,
  });
}

function addVline(slide, x, y, h, color) {
  slide.addShape(pptx.shapes.RECTANGLE, {
    x, y, w: 0.05, h,
    fill: { color }, line: { color },
  });
}

function addHline(slide, x, y, w, color) {
  slide.addShape(pptx.shapes.RECTANGLE, {
    x, y, w, h: 0.03,
    fill: { color }, line: { color },
  });
}

// ─────────────────────────────────────────────
// 슬라이드 1 — 표지 (TYPE-A, Light)
// ─────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bgLight };

  // 우측 수직 accentTeal 띠
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 9.5, y: 0, w: 0.5, h: 5.63,
    fill: { color: C.accentTeal }, line: { color: C.accentTeal },
  });

  slide.addText("뭐 들다가 삔\n급성 허리 통증\n셀프 케어 비법", {
    x: 0.6, y: 1.0, w: 6.5, h: 2.2,
    fontSize: 40, bold: true, color: C.textPrimary, fontFace: F,
    align: "left", valign: "top", wrap: true,
  });

  slide.addText("집에서 바로 실천하는 허리 회복 가이드", {
    x: 0.6, y: 3.3, w: 6.0, h: 0.5,
    fontSize: 20, bold: false, color: C.textMuted, fontFace: F, align: "left",
  });

  slide.addText("능골한의원", {
    x: 0.6, y: 4.8, w: 3.0, h: 0.4,
    fontSize: 16, bold: true, color: C.accentTeal, fontFace: F, align: "left",
  });

  slide.addNotes(
    "【발표 멘트】\n" +
    "안녕하세요, 능골한의원입니다. 오늘은 물건을 들다가 허리를 삐었을 때, 집에서 바로 할 수 있는 셀프 케어 방법에 대해 말씀드리려고 해요. 누구나 한 번쯤은 무거운 짐을 들다가 억! 하고 허리가 뻐근해진 경험이 있으실 거예요. 이럴 때 어떻게 해야 하는지, 오늘 쉽고 실용적으로 알려드릴게요.\n\n" +
    "【전환 멘트】\n" +
    "먼저, 이런 경험 혹시 있으신가요?"
  );
}

// ─────────────────────────────────────────────
// 슬라이드 2 — 목차 (TYPE-B, Dark)
// ─────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bgDarkNav };

  addBadge(slide, "INTRO", false);

  slide.addText("오늘의 강의 목차", {
    x: 0.6, y: 0.75, w: 8.0, h: 0.5,
    fontSize: 30, bold: true, color: C.textInverse, fontFace: F,
  });

  const items = [
    "이런 경험 있으신가요? — 허리 통증의 흔함",
    "허리를 삔다는 것? / 왜 들다가 삐는 걸까?",
    "안심하세요 — 대부분 좋아집니다 + 급성기 대처",
    "냉·온찜질 / 편한 자세 / 금지 동작 / 올바른 들기",
    "가볍게 움직이기 + 스트레칭 3가지 + 한의원 치료 + 재발 방지",
  ];

  items.forEach((text, i) => {
    const y = 1.5 + i * 0.75;
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: 0.6, y, w: 8.8, h: 0.62,
      fill: { color: C.cardDarkNav }, line: { color: C.cardDarkNav }, rectRadius: 0.22,
    });
    slide.addText("0" + (i + 1), {
      x: 0.7, y: y + 0.06, w: 0.6, h: 0.5,
      fontSize: 22, bold: true, color: C.accentTeal, fontFace: F, align: "center",
    });
    addVline(slide, 1.3, y + 0.1, 0.42, C.accentTeal);
    slide.addText(text, {
      x: 1.45, y: y + 0.1, w: 7.8, h: 0.42,
      fontSize: 16, color: C.textInverse, fontFace: F, valign: "middle",
    });
  });

  slide.addNotes(
    "【발표 멘트】\n" +
    "오늘 강의 순서입니다. 허리가 왜 삐는지부터, 급할 때 바로 해야 할 것, 스트레칭, 한의원 치료, 그리고 재발 방지까지 쭉 알려드릴게요. 끝까지 들으시면 집에서 바로 실천할 수 있는 방법들을 가져가실 수 있습니다.\n\n" +
    "【전환 멘트】\n" +
    "자, 그럼 첫 번째 이야기부터 시작해볼까요?"
  );
}

// ─────────────────────────────────────────────
// 슬라이드 3 — 이런 경험 있으신가요? (TYPE-C, Light)
// ─────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bgLight };

  addVline(slide, 1.5, 0.9, 1.8, C.accentTeal);

  slide.addText("허리 통증은\n정말 흔한 일입니다", {
    x: 2.0, y: 0.9, w: 6.5, h: 1.1,
    fontSize: 36, bold: true, color: C.textPrimary, fontFace: F, wrap: true,
  });

  addHline(slide, 2.0, 2.2, 5.0, C.bgDarkNav);

  const experiences = [
    "무거운 택배를 들다가 갑자기 허리가 뻐근...",
    "아이를 번쩍 안아올리다가 허리에서 뚝 소리가...",
    "허리를 구부려 바닥의 물건을 집다가 꼼짝도 못하게...",
  ];

  experiences.forEach((text, i) => {
    const y = 2.45 + i * 0.68;
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: 0.6, y, w: 5.8, h: 0.6,
      fill: { color: C.cardLight }, line: { color: C.cardLight }, rectRadius: 0.22,
    });
    slide.addText('"' + text + '"', {
      x: 0.8, y: y + 0.03, w: 5.5, h: 0.54,
      fontSize: 14, color: C.textPrimary, fontFace: F, italic: true, valign: "middle",
    });
  });

  // 통계 카드
  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 6.7, y: 2.45, w: 2.9, h: 2.7,
    fill: { color: C.cardLight }, line: { color: C.cardLight }, rectRadius: 0.22,
  });
  addVline(slide, 6.75, 2.5, 2.6, C.accentTeal);
  slide.addText("전 세계 인구\n60~80%\n일생에 1회 이상\n허리 통증 경험", {
    x: 6.95, y: 2.55, w: 2.5, h: 1.2,
    fontSize: 13, color: C.textPrimary, fontFace: F, wrap: true, bold: true,
  });
  slide.addText("한국 성인\n약 16%\n현재 허리 통증\n호소 중", {
    x: 6.95, y: 3.85, w: 2.5, h: 1.1,
    fontSize: 13, color: C.textPrimary, fontFace: F, wrap: true, bold: true,
  });

  slide.addText("나만 이런 게 아닙니다 — 안심하세요!", {
    x: 0.6, y: 5.18, w: 6.0, h: 0.32,
    fontSize: 15, italic: true, color: C.accentTeal, fontFace: F,
  });

  slide.addNotes(
    "【발표 멘트】\n" +
    "혹시 이런 경험 있으신가요? 택배를 들다가, 아이를 안다가, 바닥에 있는 물건을 집다가 갑자기 허리가 뻐근해진 적요. 사실 이건 정말 흔한 일이에요. 전 세계적으로 10명 중 6~8명은 살면서 한 번은 허리 통증을 겪는다고 합니다. 한국 성인만 봐도 약 16%가 허리 통증을 갖고 계세요. 그러니까 나만 이런가 하고 걱정하실 필요 전혀 없습니다.\n\n" +
    "【전환 멘트】\n" +
    "그런데 허리를 삔다는 게 정확히 어떤 건지, 쉽게 설명해드릴게요."
  );
}

// ─────────────────────────────────────────────
// 슬라이드 4 — 허리를 삔다는 것? (TYPE-F, Light)
// ─────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bgLight };

  addBadgeCustom(slide, "ANATOMY", C.badgeBgLight, C.accentTeal);

  slide.addText("허리를 삔다는 것?", {
    x: 0.6, y: 0.75, w: 8.0, h: 0.5,
    fontSize: 30, bold: true, color: C.textPrimary, fontFace: F,
  });

  addVline(slide, 0.6, 1.45, 0.9, C.accentTeal);
  slide.addText("요추(허리뼈) = 5개의 뼈 + 추간판(디스크) + 근육·인대", {
    x: 0.85, y: 1.45, w: 8.8, h: 0.4,
    fontSize: 17, bold: true, color: C.textPrimary, fontFace: F,
  });
  slide.addText("뼈 사이 충격흡수 물렁뼈와 주변 근육·인대가 허리를 지탱합니다", {
    x: 0.85, y: 1.9, w: 8.8, h: 0.35,
    fontSize: 14, color: C.textMuted, fontFace: F,
  });

  addLightCard(slide, {
    x: 0.6, y: 2.5, w: 4.2, h: 2.6,
    title: "근육 좌상 (Strain)",
    titleColor: C.accentTeal,
    body: "근육이 과하게 늘어나\n미세하게 찢어진 것\n\n통증, 뻐근함, 움직임 제한",
    bodyColor: C.textPrimary,
    titleSize: 18, bodySize: 14,
  });
  addLightCard(slide, {
    x: 5.2, y: 2.5, w: 4.2, h: 2.6,
    title: "인대 염좌 (Sprain)",
    titleColor: C.accentTeal,
    body: "인대가 과하게 늘어나\n손상된 것\n\n통증, 불안정감",
    bodyColor: C.textPrimary,
    titleSize: 18, bodySize: 14,
  });

  slide.addText("약 90%의 급성 허리 통증 = 근육·인대 손상 (뼈나 디스크 문제가 아닌 경우 대부분!)", {
    x: 0.6, y: 5.18, w: 9.0, h: 0.32,
    fontSize: 13, bold: true, color: C.accentTeal, fontFace: F, align: "center",
  });

  slide.addNotes(
    "【발표 멘트】\n" +
    "우리 허리는 요추라고 하는 5개의 뼈로 이루어져 있어요. 이 뼈 사이사이에는 충격을 흡수하는 물렁뼈, 전문용어로 추간판이라고 하는 게 있고요, 그 주변을 근육과 인대가 꽉 잡아주고 있습니다. 허리를 삔다는 건, 이 근육이나 인대가 과하게 늘어나면서 미세하게 찢어진 거예요. 의학용어로는 근육 좌상이라고 하고요. 중요한 건, 급성 허리 통증의 약 90%가 이런 근육이나 인대 손상이라는 거예요. 뼈가 부러지거나 디스크가 터진 게 아닌 경우가 대부분이라는 말씀입니다.\n\n" +
    "【전환 멘트】\n" +
    "그럼 왜 하필 물건을 들다가 이런 일이 생기는 걸까요?"
  );
}

// ─────────────────────────────────────────────
// 슬라이드 5 — 왜 들다가 삐는 걸까? (TYPE-D, Dark)
// ─────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bgDarkNav };

  addBadgeCustom(slide, "CAUSE", C.badgeBg, C.textInverse);

  slide.addText("왜 들다가 삐는 걸까?", {
    x: 0.6, y: 0.75, w: 8.0, h: 0.5,
    fontSize: 30, bold: true, color: C.textInverse, fontFace: F,
  });

  const causes = [
    "허리만 구부려서 들기 → 다리 아닌 허리 근육에 힘이 집중",
    "물건을 몸에서 멀리 두고 들기 → 지렛대 효과로 허리 부하 급증",
    "들면서 동시에 몸을 비트는 동작 → 인대·관절 손상 위험",
    "준비 없이 갑자기 훅 하고 들기 → 근육이 방어할 시간 부족",
  ];

  causes.forEach((text, i) => {
    const y = 1.55 + i * 0.82;
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: 0.6, y, w: 8.8, h: 0.72,
      fill: { color: C.cardDarkNav }, line: { color: C.cardDarkNav }, rectRadius: 0.22,
    });
    addVline(slide, 0.65, y + 0.1, 0.52, C.accentDanger);
    slide.addText(text, {
      x: 0.9, y: y + 0.1, w: 8.3, h: 0.52,
      fontSize: 15, color: C.textInverse, fontFace: F, valign: "middle",
    });
  });

  slide.addText("무릎은 안 굽히고 허리만 굽히는 것이 가장 흔한 원인!", {
    x: 0.6, y: 5.05, w: 8.8, h: 0.35,
    fontSize: 14, bold: true, color: C.accentTeal, fontFace: F, align: "center",
  });

  slide.addNotes(
    "【발표 멘트】\n" +
    "물건을 들다가 허리가 삐는 이유, 크게 네 가지입니다. 첫째, 무릎은 안 굽히고 허리만 구부려서 들 때. 이러면 다리 근육 대신 허리 근육에 힘이 다 몰려요. 둘째, 물건을 몸에서 멀리 두고 들 때. 물건이 몸에서 멀어질수록 허리에 가는 부담이 훨씬 커집니다. 셋째, 들면서 동시에 몸을 비틀 때. 이건 인대나 관절에 아주 안 좋아요. 넷째, 준비 없이 갑자기 확 들 때. 근육이 준비할 시간도 없이 힘이 가해지면 당연히 다칠 수밖에 없겠죠.\n\n" +
    "【전환 멘트】\n" +
    "이렇게 허리를 삐면 많이 걱정되시죠. 그런데 아주 중요한 사실을 먼저 알려드릴게요."
  );
}

// ─────────────────────────────────────────────
// 슬라이드 6 — 안심하세요, 대부분 좋아집니다 (TYPE-G, Light)
// ─────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bgLight };

  addBadgeCustom(slide, "RECOVERY", C.badgeBgLight, C.accentTeal);

  slide.addText("안심하세요 — 대부분 좋아집니다", {
    x: 0.6, y: 0.75, w: 8.8, h: 0.5,
    fontSize: 28, bold: true, color: C.textPrimary, fontFace: F,
  });

  const stats = [
    { num: "1개월", sub: "통증 58% 감소\n직장 복귀 82%" },
    { num: "6주",   sub: "대부분\n의미 있는 호전" },
    { num: "3개월", sub: "통증 안정화\n일상 복귀" },
    { num: "12개월", sub: "직장 복귀\n93%" },
  ];

  stats.forEach(function(s, i) {
    const x = 0.5 + i * 2.3;
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x, y: 1.5, w: 2.1, h: 2.1,
      fill: { color: C.cardLight }, line: { color: C.cardLight }, rectRadius: 0.22,
    });
    slide.addShape(pptx.shapes.RECTANGLE, {
      x, y: 1.5, w: 2.1, h: 0.04,
      fill: { color: C.accentTeal }, line: { color: C.accentTeal },
    });
    slide.addText(s.num, {
      x: x + 0.05, y: 1.65, w: 2.0, h: 0.55,
      fontSize: 28, bold: true, color: C.accentTeal, fontFace: F, align: "center",
    });
    slide.addText(s.sub, {
      x: x + 0.1, y: 2.25, w: 1.9, h: 1.2,
      fontSize: 13, color: C.textMuted, fontFace: F, align: "center", wrap: true,
    });
  });

  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 3.85, w: 8.8, h: 0.75,
    fill: { color: C.cardLight }, line: { color: C.accentTeal, width: 2 }, rectRadius: 0.22,
  });
  slide.addText("약 90%의 환자가 6주 이내에 의미 있게 좋아집니다", {
    x: 0.7, y: 3.9, w: 8.6, h: 0.65,
    fontSize: 18, bold: true, color: C.accentTeal, fontFace: F, align: "center", valign: "middle",
  });

  slide.addText("올바르게 관리하면 더 빨리 회복됩니다!", {
    x: 0.6, y: 4.75, w: 8.8, h: 0.35,
    fontSize: 14, italic: true, color: C.textMuted, fontFace: F, align: "center",
  });

  slide.addNotes(
    "【발표 멘트】\n" +
    "가장 먼저 드리고 싶은 말씀은, 안심하세요입니다. 연구 결과를 보면, 급성 허리 통증 환자의 약 90%가 6주 안에 의미 있게 좋아집니다. 한 달만 지나도 통증이 절반 이상 줄고, 82%가 일상으로 복귀해요. 지금은 많이 아프고 걱정되시겠지만, 대부분은 시간이 지나면서 자연스럽게 회복됩니다. 물론 올바르게 관리하면 더 빨리 좋아질 수 있어요.\n\n" +
    "【전환 멘트】\n" +
    "그럼 지금 당장 집에서 뭘 해야 하는지, 바로 알려드릴게요."
  );
}

// ─────────────────────────────────────────────
// 슬라이드 7 — 지금 당장 해야 할 것 (TYPE-D, Light)
// ─────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bgLight };

  addBadgeCustom(slide, "ACUTE CARE", C.accentDanger, C.textInverse);

  slide.addText("지금 당장 해야 할 것 (급성기 1~2일)", {
    x: 0.6, y: 0.75, w: 8.8, h: 0.5,
    fontSize: 28, bold: true, color: C.textPrimary, fontFace: F,
  });

  const actions = [
    { label: "보호", text: "무거운 것 들기 자제, 통증 유발 동작 피하기" },
    { label: "적절한 휴식", text: "통증이 심한 첫 1~2일은 활동 줄이기" },
    { label: "온찜질", text: "하루 2번, 1회 20분 (다음 슬라이드에서 자세히)" },
    { label: "가볍게 움직이기", text: "화장실 가기, 집안 걸어 다니기 — 통증 범위 안에서" },
  ];

  actions.forEach(function(a, i) {
    const y = 1.5 + i * 0.74;
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: 0.6, y, w: 8.8, h: 0.65,
      fill: { color: C.cardLight }, line: { color: C.cardLight }, rectRadius: 0.22,
    });
    addVline(slide, 0.65, y + 0.1, 0.45, C.accentSuccess);
    slide.addText(a.label, {
      x: 0.9, y: y + 0.1, w: 1.8, h: 0.45,
      fontSize: 14, bold: true, color: C.accentTeal, fontFace: F, valign: "middle",
    });
    slide.addText(a.text, {
      x: 2.75, y: y + 0.1, w: 6.5, h: 0.45,
      fontSize: 14, color: C.textPrimary, fontFace: F, valign: "middle",
    });
  });

  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 4.55, w: 8.8, h: 0.65,
    fill: { color: "FEF2F2" }, line: { color: C.accentDanger, width: 2 }, rectRadius: 0.22,
  });
  slide.addText("48시간 이상 침상 안정은 금지! — 오히려 회복이 느려집니다", {
    x: 0.8, y: 4.6, w: 8.4, h: 0.55,
    fontSize: 15, bold: true, color: C.accentDanger, fontFace: F, align: "center", valign: "middle",
  });

  slide.addNotes(
    "【발표 멘트】\n" +
    "허리를 삐었을 때, 첫 1~2일이 가장 중요해요. 이때는 무거운 것 들기를 피하시고, 통증을 악화시키는 동작은 하지 마세요. 그런데 여기서 아주 중요한 포인트! 이틀 넘게 꼼짝도 안 하고 누워만 있으면 오히려 회복이 느려진다는 연구 결과가 있어요. 여러 가이드라인에서 공통으로 권고하는 사항입니다. 통증이 허락하는 범위에서 화장실도 가고, 집안을 가볍게 걸어 다니세요. 그리고 온찜질을 해주시면 좋아요. 하루에 두 번, 한 번에 20분 정도요.\n\n" +
    "【전환 멘트】\n" +
    "그런데 냉찜질을 해야 할까요, 온찜질을 해야 할까요? 이것 참 헷갈리시죠?"
  );
}

// ─────────────────────────────────────────────
// 슬라이드 8 — 냉찜질 vs 온찜질 (TYPE-C, Dark)
// ─────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bgDarkNav };

  addBadgeCustom(slide, "HEAT THERAPY", C.badgeBg, C.textInverse);

  addVline(slide, 1.5, 0.85, 1.3, C.accentTeal);
  slide.addText("허리 통증엔\n온찜질이 먼저!", {
    x: 2.0, y: 0.85, w: 7.0, h: 1.0,
    fontSize: 34, bold: true, color: C.textInverse, fontFace: F, wrap: true,
  });

  addHline(slide, 2.0, 2.1, 5.0, C.accentTeal);

  // 온찜질 카드
  addNavCard(slide, {
    x: 0.6, y: 2.38, w: 4.3, h: 2.85,
    title: "온찜질 (Heat) — 1순위 추천",
    titleColor: C.accentSuccess,
    body: "371명 대상 연구:\n온열 패치가 진통제보다\n통증 완화 33~52% 우수\n\n혈류 증가 -> 근육 이완 -> 통증 차단\n\n1회 20분, 하루 2회, 1주일간",
    bodyColor: C.textInverse,
    titleSize: 15, bodySize: 13,
  });
  addVline(slide, 0.65, 2.43, 2.75, C.accentSuccess);

  // 냉찜질 카드
  addNavCard(slide, {
    x: 5.2, y: 2.38, w: 4.3, h: 2.85,
    title: "냉찜질 (Cold) — 근거 제한적",
    titleColor: C.textMuted,
    body: "허리 통증 냉찜질 연구 매우 부족\n\n타박상 등 외상 직후에만\n제한적 사용 가능\n\n허리 통증엔 권고 미흡",
    bodyColor: C.textInverse,
    titleSize: 15, bodySize: 13,
  });
  addVline(slide, 5.25, 2.43, 2.75, C.textMuted);

  slide.addText("결론: 허리 통증엔 온찜질 먼저! 본인이 편한 쪽을 선택하셔도 됩니다", {
    x: 0.6, y: 5.3, w: 9.0, h: 0.25,
    fontSize: 12, bold: true, color: C.accentTeal, fontFace: F, align: "center",
  });

  slide.addNotes(
    "【발표 멘트】\n" +
    "다치면 냉찜질부터 해야 한다고 많이 들으셨죠? 그런데 허리 통증에서는 좀 달라요. 371명을 대상으로 한 연구를 보면, 온찜질이 진통제보다 통증 완화에 33~52%나 더 효과적이었습니다. 온찜질이 혈류를 증가시키고 근육을 이완시켜서 통증 신호를 줄여주거든요. 반면에 냉찜질은 허리 통증에 대한 근거가 아직 충분하지 않아요. 그래서 전문가들은 허리 통증에는 온찜질을 먼저 추천합니다. 핫팩이나 따뜻한 수건으로 하루에 두 번, 한 번에 20분씩 해보세요.\n\n" +
    "【전환 멘트】\n" +
    "자, 이번에는 아플 때 어떤 자세가 편한지 알아볼게요."
  );
}

// ─────────────────────────────────────────────
// 슬라이드 9 — 이 자세가 편해요 (TYPE-E, Light)
// ─────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bgLight };

  addBadgeCustom(slide, "POSTURE", C.badgeBgLight, C.accentTeal);

  slide.addText("이 자세가 편해요 (통증 완화 자세)", {
    x: 0.6, y: 0.75, w: 8.8, h: 0.5,
    fontSize: 28, bold: true, color: C.textPrimary, fontFace: F,
  });

  const postures = [
    {
      title: "누울 때",
      body: "바로 누워 무릎 아래 베개 1~2개\n옆으로 누워 무릎 사이 베개\n엎드려 자는 것은 피하세요",
    },
    {
      title: "앉을 때",
      body: "수건 말아 허리에 받치기\n10~15분마다 자세 바꾸기\n푹신한 소파·낮은 의자 피하기",
    },
    {
      title: "서 있을 때",
      body: "한 발을 낮은 발판에 올리기\n체중 분산으로 허리 부담 감소",
    },
    {
      title: "피할 자세",
      body: "엎드려 자기\n깊이 기대는 소파\n장시간 같은 자세 유지",
    },
  ];

  postures.forEach(function(p, i) {
    const x = 0.55 + i * 2.3;
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x, y: 1.55, w: 2.15, h: 3.55,
      fill: { color: C.cardLight }, line: { color: C.cardLight }, rectRadius: 0.22,
    });
    slide.addShape(pptx.shapes.RECTANGLE, {
      x, y: 1.55, w: 2.15, h: 0.04,
      fill: { color: C.accentTeal }, line: { color: C.accentTeal },
    });
    slide.addText(p.title, {
      x: x + 0.15, y: 1.65, w: 1.85, h: 0.4,
      fontSize: 15, bold: true, color: C.accentTeal, fontFace: F,
    });
    slide.addText(p.body, {
      x: x + 0.12, y: 2.1, w: 1.9, h: 2.9,
      fontSize: 12, color: C.textPrimary, fontFace: F, wrap: true,
    });
  });

  slide.addText("자세만 바꿔도 통증이 줄어듭니다", {
    x: 0.6, y: 5.22, w: 9.0, h: 0.28,
    fontSize: 14, italic: true, color: C.accentTeal, fontFace: F, align: "center",
  });

  slide.addNotes(
    "【발표 멘트】\n" +
    "허리가 아플 때 자세가 정말 중요해요. 누울 때는 바로 누워서 무릎 아래에 베개를 한두 개 받쳐보세요. 이것만으로도 허리에 가는 부담이 확 줄어듭니다. 옆으로 누울 때는 무릎 사이에 베개를 끼우면 허리가 비틀어지는 걸 막아줘요. 엎드려 자는 건 급성기에 통증이 악화될 수 있으니 피하시고요. 앉을 때는 수건을 말아서 허리 뒤에 받쳐주세요. 그리고 한 자세로 10~15분 이상 앉아있지 마시고 자주 바꿔주세요. 푹신한 소파나 낮은 의자는 허리 곡선을 망가뜨려서 안 좋습니다.\n\n" +
    "【전환 멘트】\n" +
    "편한 자세를 알았으니, 이번엔 절대 하면 안 되는 동작을 짚어볼게요."
  );
}

// ─────────────────────────────────────────────
// 슬라이드 10 — 이건 절대 하지 마세요 (TYPE-E, Dark)
// ─────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bgDarkNav };

  addBadgeCustom(slide, "WARNING", C.accentDanger, C.textInverse);

  slide.addText("이건 절대 하지 마세요 (금지 동작)", {
    x: 0.6, y: 0.75, w: 8.8, h: 0.5,
    fontSize: 28, bold: true, color: C.textInverse, fontFace: F,
  });

  const donts = [
    "허리 굽혀 발가락 만지기 (신경 자극 + 디스크 압박)",
    "앉은 자세에서 갑자기 앞으로 몸 기울이기",
    "무거운 것 들기 (급성기 5kg 이상 자제)",
    "30분 이상 같은 자세로 있기",
    "달리기·줄넘기 등 충격이 큰 운동",
  ];

  donts.forEach(function(text, i) {
    const y = 1.55 + i * 0.7;
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: 0.6, y, w: 8.8, h: 0.62,
      fill: { color: C.cardDarkNav }, line: { color: C.cardDarkNav }, rectRadius: 0.22,
    });
    addVline(slide, 0.65, y + 0.1, 0.42, C.accentDanger);
    slide.addText("X", {
      x: 0.8, y: y + 0.1, w: 0.4, h: 0.42,
      fontSize: 14, bold: true, color: C.accentDanger, fontFace: F, align: "center", valign: "middle",
    });
    slide.addText(text, {
      x: 1.25, y: y + 0.1, w: 8.0, h: 0.42,
      fontSize: 14, color: C.textInverse, fontFace: F, valign: "middle",
    });
  });

  slide.addText("급성기에 이 동작들은 통증을 악화시키고 회복을 늦춥니다", {
    x: 0.6, y: 5.18, w: 8.8, h: 0.32,
    fontSize: 13, bold: true, color: C.accentDanger, fontFace: F, align: "center",
  });

  slide.addNotes(
    "【발표 멘트】\n" +
    "허리가 아플 때 절대 하시면 안 되는 동작들이 있어요. 첫째, 허리를 굽혀서 발끝을 만지려는 스트레칭. 이게 좋을 것 같지만, 급성기에는 신경을 자극하고 디스크에 압력을 줄 수 있어요. 둘째, 앉아 있다가 갑자기 앞으로 확 기울이는 거. 셋째, 무거운 것 들기. 급성기에는 5kg 이상은 피해주세요. 넷째, 30분 이상 같은 자세로 가만히 있는 것. 이건 정말 안 좋습니다. 다섯째, 달리기나 줄넘기 같은 충격이 큰 운동도 급성기에는 피하셔야 해요.\n\n" +
    "【전환 멘트】\n" +
    "그럼 나중에 물건을 들 때는 어떻게 해야 안전할까요?"
  );
}

// ─────────────────────────────────────────────
// 슬라이드 11 — 올바른 물건 들기 (TYPE-F, Light)
// ─────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bgLight };

  addBadgeCustom(slide, "TECHNIQUE", C.badgeBgLight, C.accentTeal);

  slide.addText("올바른 물건 들기", {
    x: 0.6, y: 0.75, w: 8.0, h: 0.5,
    fontSize: 30, bold: true, color: C.textPrimary, fontFace: F,
  });

  addVline(slide, 0.6, 1.45, 0.6, C.accentTeal);
  slide.addText("물건은 몸에 가까이, 무릎을 굽혀서, 천천히 들어 올리세요", {
    x: 0.85, y: 1.45, w: 8.8, h: 0.55,
    fontSize: 16, bold: true, color: C.textPrimary, fontFace: F,
  });

  const steps = [
    { num: "1", text: "물건에 최대한 가까이 다가가기" },
    { num: "2", text: "발을 어깨 너비로 벌려 안정감 확보" },
    { num: "3", text: "무릎을 굽혀 쪼그려 앉기 (허리 아닌 다리 힘!)" },
    { num: "4", text: "배에 힘주기 (코어 안정화)" },
    { num: "5", text: "등을 곧게 유지하며 서서히 일어서기" },
  ];

  steps.forEach(function(s, i) {
    const y = 2.2 + i * 0.56;
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: 0.6, y, w: 8.8, h: 0.49,
      fill: { color: C.cardLight }, line: { color: C.cardLight }, rectRadius: 0.22,
    });
    slide.addText(s.num, {
      x: 0.72, y: y + 0.04, w: 0.45, h: 0.41,
      fontSize: 18, bold: true, color: C.accentTeal, fontFace: F, align: "center", valign: "middle",
    });
    slide.addText(s.text, {
      x: 1.25, y: y + 0.04, w: 8.0, h: 0.41,
      fontSize: 14, color: C.textPrimary, fontFace: F, valign: "middle",
    });
  });

  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 5.03, w: 8.8, h: 0.42,
    fill: { color: "FEF2F2" }, line: { color: C.accentDanger, width: 1 }, rectRadius: 0.18,
  });
  slide.addText("절대 금지: 허리 구부린 채 몸 비틀기", {
    x: 0.7, y: 5.06, w: 8.6, h: 0.36,
    fontSize: 14, bold: true, color: C.accentDanger, fontFace: F, align: "center", valign: "middle",
  });

  slide.addNotes(
    "【발표 멘트】\n" +
    "물건을 안전하게 드는 방법, 꼭 기억해주세요. 가장 중요한 건, 물건에 최대한 가까이 다가가는 거예요. 물건이 몸에서 멀어질수록 허리에 가는 부담이 크게 늘어납니다. 그 다음, 발을 어깨 너비로 벌리고, 무릎을 굽혀서 쪼그려 앉으세요. 허리가 아니라 다리 근육으로 드는 겁니다. 그리고 배에 살짝 힘을 주면서 등을 곧게 유지한 채 천천히 일어나세요. 한 가지 절대 하지 말아야 할 게 있는데요, 허리를 구부린 채로 몸을 비트는 동작이에요. 이건 인대와 관절에 정말 무리를 줍니다.\n\n" +
    "【전환 멘트】\n" +
    "자, 급성기 첫 이틀이 지났다면, 이제 조금씩 움직이기 시작해야 해요."
  );
}

// ─────────────────────────────────────────────
// 슬라이드 12 — 2~3일 후, 가볍게 움직이세요 (TYPE-F, Dark)
// ─────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bgDarkNav };

  addBadgeCustom(slide, "MOVEMENT", C.badgeBg, C.textInverse);

  slide.addText("2~3일 후, 가볍게 움직이세요", {
    x: 0.6, y: 0.75, w: 8.0, h: 0.5,
    fontSize: 28, bold: true, color: C.textInverse, fontFace: F,
  });

  addVline(slide, 0.6, 1.45, 0.6, C.accentTeal);
  slide.addText("통증이 허락하는 범위에서 움직이기가 최선의 전략", {
    x: 0.85, y: 1.45, w: 8.8, h: 0.55,
    fontSize: 15, color: C.textMuted, fontFace: F,
  });

  const items = [
    { step: "STEP 1", text: "급성기(첫 1~2일) 지나면 가볍게 걷기 시작", color: C.accentSuccess },
    { step: "STEP 2", text: "하루 15~30분 걷기부터 (무리하지 않는 범위에서)", color: C.accentSuccess },
    { step: "STEP 3", text: "통증이 50% 이상 줄었을 때 스트레칭 시작", color: C.accentTeal },
  ];

  items.forEach(function(item, i) {
    const y = 2.2 + i * 0.85;
    addNavCard(slide, { x: 0.6, y, w: 8.8, h: 0.75 });
    addVline(slide, 0.65, y + 0.1, 0.55, item.color);
    slide.addText(item.step, {
      x: 0.85, y: y + 0.1, w: 1.1, h: 0.55,
      fontSize: 12, bold: true, color: item.color, fontFace: F, valign: "middle",
    });
    slide.addText(item.text, {
      x: 2.0, y: y + 0.1, w: 7.2, h: 0.55,
      fontSize: 15, color: C.textInverse, fontFace: F, valign: "middle",
    });
  });

  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 4.9, w: 8.8, h: 0.55,
    fill: { color: C.cardDarkNav }, line: { color: C.accentDanger, width: 1 }, rectRadius: 0.22,
  });
  slide.addText("침대에만 누워있지 마세요! -> 근육 약화 + 회복 지연", {
    x: 0.7, y: 4.93, w: 8.6, h: 0.49,
    fontSize: 14, bold: true, color: C.accentDanger, fontFace: F, align: "center", valign: "middle",
  });

  slide.addNotes(
    "【발표 멘트】\n" +
    "첫 이틀이 지나면, 이제 조금씩 움직여야 합니다. 아직 아픈데 움직여도 되나요? 걱정되시죠? 여러 연구에서 확인된 건, 계속 누워만 있으면 오히려 근육이 약해지고 회복이 느려진다는 거예요. 통증이 허락하는 범위에서 하루 15~30분 정도 가볍게 걸어보세요. 그리고 통증이 처음의 절반 이하로 줄었을 때, 다음에 알려드릴 스트레칭을 시작하시면 됩니다. 아프지 않은 범위에서 움직이기가 가장 좋은 전략이에요.\n\n" +
    "【전환 멘트】\n" +
    "통증이 좀 가라앉았다면, 이 세 가지 스트레칭을 해보세요."
  );
}

// ─────────────────────────────────────────────
// 슬라이드 13 — 스트레칭 ① 무릎 당기기 (TYPE-F, Light)
// ─────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bgLight };

  addBadgeCustom(slide, "STRETCH 1", C.badgeBgLight, C.accentTeal);

  slide.addText("스트레칭 1  무릎 당기기 (Knee-to-Chest)", {
    x: 0.6, y: 0.75, w: 8.8, h: 0.5,
    fontSize: 26, bold: true, color: C.textPrimary, fontFace: F,
  });

  addVline(slide, 0.6, 1.45, 0.6, C.accentTeal);
  slide.addText("허리 뒤쪽 근육과 관절을 부드럽게 풀어줌 — 효과가 검증된 클래식 요통 스트레칭", {
    x: 0.85, y: 1.45, w: 8.8, h: 0.55,
    fontSize: 14, color: C.textMuted, fontFace: F,
  });

  const steps = [
    { num: "1", text: "바닥에 편하게 바로 눕기" },
    { num: "2", text: "양 무릎을 가슴 쪽으로 천천히 당기기" },
    { num: "3", text: "20~30초 유지" },
    { num: "4", text: "3회 반복" },
  ];

  steps.forEach(function(s, i) {
    const y = 2.2 + i * 0.65;
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: 0.6, y, w: 6.0, h: 0.57,
      fill: { color: C.cardLight }, line: { color: C.cardLight }, rectRadius: 0.22,
    });
    slide.addText(s.num, {
      x: 0.75, y: y + 0.06, w: 0.4, h: 0.45,
      fontSize: 16, bold: true, color: C.accentTeal, fontFace: F, align: "center", valign: "middle",
    });
    slide.addText(s.text, {
      x: 1.25, y: y + 0.06, w: 5.2, h: 0.45,
      fontSize: 14, color: C.textPrimary, fontFace: F, valign: "middle",
    });
  });

  // 오른쪽 정보 카드
  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 6.9, y: 2.2, w: 2.7, h: 2.6,
    fill: { color: C.cardLight }, line: { color: C.cardLight }, rectRadius: 0.22,
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 6.9, y: 2.2, w: 2.7, h: 0.04,
    fill: { color: C.accentTeal }, line: { color: C.accentTeal },
  });
  slide.addText("효과", {
    x: 7.1, y: 2.28, w: 2.3, h: 0.4,
    fontSize: 15, bold: true, color: C.accentTeal, fontFace: F,
  });
  slide.addText("허리 뒤쪽 근육과\n관절을 부드럽게\n풀어줌", {
    x: 7.1, y: 2.72, w: 2.3, h: 1.0,
    fontSize: 13, color: C.textPrimary, fontFace: F, wrap: true,
  });
  slide.addText("20~30초 x 3회", {
    x: 7.1, y: 3.8, w: 2.3, h: 0.5,
    fontSize: 18, bold: true, color: C.accentTeal, fontFace: F, align: "center",
  });

  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 4.9, w: 6.0, h: 0.45,
    fill: { color: "FEF2F2" }, line: { color: C.accentDanger, width: 1 }, rectRadius: 0.18,
  });
  slide.addText("통증이 심해지면 즉시 멈추세요!", {
    x: 0.7, y: 4.93, w: 5.8, h: 0.39,
    fontSize: 13, bold: true, color: C.accentDanger, fontFace: F, align: "center", valign: "middle",
  });

  slide.addNotes(
    "【발표 멘트】\n" +
    "첫 번째 스트레칭은 무릎 당기기예요. 바닥에 편하게 바로 누워서, 양 무릎을 가슴 쪽으로 천천히 당겨보세요. 이 자세를 20~30초 정도 유지하시고, 3번 반복하시면 됩니다. 이 동작은 허리 뒤쪽 근육과 관절을 부드럽게 풀어주는 효과가 있어요. 아주 클래식한 요통 스트레칭인데, 효과가 검증된 동작입니다. 다만 이 동작을 하다가 오히려 통증이 심해진다면, 아직 시기가 이른 거예요. 바로 멈추시고 며칠 뒤에 다시 시도해보세요.\n\n" +
    "【전환 멘트】\n" +
    "두 번째 스트레칭도 아주 쉽습니다."
  );
}

// ─────────────────────────────────────────────
// 슬라이드 14 — 스트레칭 ② 고양이-소 자세 (TYPE-F, Dark)
// ─────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bgDarkNav };

  addBadgeCustom(slide, "STRETCH 2", C.badgeBg, C.textInverse);

  slide.addText("스트레칭 2  고양이-소 자세 (Cat-Cow)", {
    x: 0.6, y: 0.75, w: 8.8, h: 0.5,
    fontSize: 26, bold: true, color: C.textInverse, fontFace: F,
  });

  addVline(slide, 0.6, 1.45, 0.6, C.accentTeal);
  slide.addText("허리 전체의 유연성 회복, 주변 근육 혈액 순환 개선 — 메이요 클리닉 권고 동작", {
    x: 0.85, y: 1.45, w: 8.8, h: 0.55,
    fontSize: 14, color: C.textMuted, fontFace: F,
  });

  const cards = [
    { title: "자세",  body: "네발기기 자세\n(손과 무릎을 바닥에)" },
    { title: "고양이 동작", body: "등을 위로 올려\n고양이처럼 둥글게 굽히기\n5초 유지" },
    { title: "소 동작", body: "등을 아래로 내려\n소처럼 배 내밀기\n5초 유지" },
    { title: "횟수",  body: "10회 반복\n천천히, 부드럽게\n통증 없는 범위에서" },
  ];

  cards.forEach(function(c, i) {
    const x = 0.55 + i * 2.3;
    addNavCard(slide, { x, y: 2.2, w: 2.1, h: 2.85, title: c.title, body: c.body, titleColor: C.accentTeal, titleSize: 14, bodySize: 13 });
  });

  slide.addText("천천히, 부드럽게, 아프지 않은 범위에서만 하세요!", {
    x: 0.6, y: 5.2, w: 9.0, h: 0.3,
    fontSize: 13, bold: true, color: C.accentSuccess, fontFace: F, align: "center",
  });

  slide.addNotes(
    "【발표 멘트】\n" +
    "두 번째는 고양이-소 자세입니다. 이름이 재밌죠? 네발기기 자세에서, 등을 위로 올려서 고양이처럼 둥글게 만들었다가, 다시 아래로 내려서 소처럼 배를 늘어뜨리는 거예요. 고양이 5초, 소 5초, 이렇게 10번 반복하시면 됩니다. 이 동작은 허리 전체의 유연성을 회복시키고, 주변 근육에 혈액 순환을 좋게 해줘요. 메이요 클리닉, 클리블랜드 클리닉 같은 유명 병원에서도 추천하는 동작입니다. 천천히, 부드럽게, 아프지 않은 범위에서만 하세요.\n\n" +
    "【전환 멘트】\n" +
    "마지막 세 번째 스트레칭이에요."
  );
}

// ─────────────────────────────────────────────
// 슬라이드 15 — 스트레칭 ③ 골반 기울이기 (TYPE-E, Light)
// ─────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bgLight };

  addBadgeCustom(slide, "STRETCH 3", C.badgeBgLight, C.accentTeal);

  slide.addText("스트레칭 3  골반 기울이기 (Pelvic Tilt)", {
    x: 0.6, y: 0.75, w: 8.8, h: 0.5,
    fontSize: 26, bold: true, color: C.textPrimary, fontFace: F,
  });

  addVline(slide, 0.6, 1.45, 0.6, C.accentTeal);
  slide.addText("복횡근(배 속 깊은 근육) 활성화 -> 허리 안정화의 기초 — 가장 안전한 코어 운동 입문!", {
    x: 0.85, y: 1.45, w: 8.8, h: 0.55,
    fontSize: 14, color: C.textMuted, fontFace: F,
  });

  const cards = [
    { title: "자세",  body: "바닥에 바로 누워\n무릎 굽히기" },
    { title: "동작",  body: "배에 힘 주어\n허리를 바닥에\n납작하게 붙이기" },
    { title: "유지",  body: "5초 유지\nx 10회 반복" },
    { title: "효과",  body: "복횡근 활성화\n허리 안정화 근육\n가장 안전한 코어 운동" },
  ];

  cards.forEach(function(c, i) {
    const x = 0.55 + i * 2.3;
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x, y: 2.2, w: 2.1, h: 2.85,
      fill: { color: C.cardLight }, line: { color: C.cardLight }, rectRadius: 0.22,
    });
    slide.addShape(pptx.shapes.RECTANGLE, {
      x, y: 2.2, w: 2.1, h: 0.04,
      fill: { color: C.accentSuccess }, line: { color: C.accentSuccess },
    });
    slide.addText(c.title, {
      x: x + 0.15, y: 2.28, w: 1.8, h: 0.4,
      fontSize: 14, bold: true, color: C.accentTeal, fontFace: F,
    });
    slide.addText(c.body, {
      x: x + 0.12, y: 2.72, w: 1.86, h: 2.2,
      fontSize: 13, color: C.textPrimary, fontFace: F, wrap: true,
    });
  });

  slide.addText("별것 아닌 것 같지만, 허리를 안정적으로 잡아주는 핵심 근육을 깨워줍니다", {
    x: 0.6, y: 5.2, w: 9.0, h: 0.3,
    fontSize: 13, italic: true, color: C.accentTeal, fontFace: F, align: "center",
  });

  slide.addNotes(
    "【발표 멘트】\n" +
    "세 번째는 골반 기울이기입니다. 바닥에 바로 누워서 무릎을 굽힌 다음, 배에 힘을 딱 주면서 허리를 바닥에 납작하게 붙여보세요. 5초 유지하고, 10번 반복합니다. 이게 별것 아닌 것 같지만, 우리 배 깊은 곳에 있는 복횡근이라는 근육을 깨워주는 아주 중요한 운동이에요. 이 근육이 허리를 안정적으로 잡아주는 코어 근육의 핵심입니다. 가장 안전한 코어 운동 입문 동작이니까, 허리가 아픈 분들도 부담 없이 시작할 수 있어요.\n\n" +
    "【전환 멘트】\n" +
    "여기까지가 집에서 하실 수 있는 셀프 케어였고요, 한의원에서는 어떤 치료를 해드리는지 궁금하시죠?"
  );
}

// ─────────────────────────────────────────────
// 슬라이드 16 — 한의원에서는 뭘 해드리나요? (TYPE-H, Dark)
// ─────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bgDarkNav };

  addBadgeCustom(slide, "TREATMENT", C.badgeBg, C.textInverse);

  slide.addText("한의원에서는 뭘 해드리나요?", {
    x: 0.6, y: 0.75, w: 8.0, h: 0.5,
    fontSize: 28, bold: true, color: C.textInverse, fontFace: F,
  });

  const treatments = [
    {
      title: "침 치료 (Acupuncture)",
      titleColor: C.accentTeal,
      body: "13개 연구(707명):\n통증 의미 있게 감소\n진통제 복용량도 줄어듦\n급성 허리 통증 중등도 근거",
      lineColor: C.accentTeal,
    },
    {
      title: "부항 치료 (Cupping)",
      titleColor: "60A5FA",
      body: "통증 및 기능장애\n의미 있게 개선\n(고~중등도 근거)\n경혈 부위 부항이 더 효과적",
      lineColor: "60A5FA",
    },
    {
      title: "추나 요법 (Chuna)",
      titleColor: C.accentSuccess,
      body: "194명 대상 국내 연구:\n통증 감소 효과 크기 0.96\n(큰 효과)\n일반 치료보다 2배 이상 감소",
      lineColor: C.accentSuccess,
    },
  ];

  treatments.forEach(function(t, i) {
    const x = 0.5 + i * 3.1;
    addNavCard(slide, { x, y: 1.55, w: 2.85, h: 3.5 });
    addVline(slide, x + 0.05, 1.6, 3.4, t.lineColor);
    slide.addText(t.title, {
      x: x + 0.22, y: 1.65, w: 2.5, h: 0.45,
      fontSize: 13, bold: true, color: t.titleColor, fontFace: F, wrap: true,
    });
    slide.addText(t.body, {
      x: x + 0.22, y: 2.15, w: 2.5, h: 2.7,
      fontSize: 12, color: C.textInverse, fontFace: F, wrap: true,
    });
  });

  slide.addText("급성기: 침·부항 치료   |   회복기: 추나 요법 병행", {
    x: 0.6, y: 5.2, w: 9.0, h: 0.3,
    fontSize: 13, bold: true, color: C.accentTeal, fontFace: F, align: "center",
  });

  slide.addNotes(
    "【발표 멘트】\n" +
    "한의원에서는 세 가지 치료를 주로 해드립니다. 첫째, 침 치료. 13개 연구를 종합해보면, 침 치료가 허리 통증을 의미 있게 줄여주고, 진통제 복용량까지 줄여준다고 합니다. 둘째, 부항 치료. 피부에 컵을 붙여서 음압을 만드는 건데요, 통증과 기능장애를 의미 있게 개선해준다는 연구 결과가 있어요. 셋째, 추나 요법. 한의사가 직접 손으로 허리를 교정하는 치료인데요, 국내 194명 대상 연구에서 일반 치료만 할 때보다 통증이 2배 이상 감소하는 큰 효과가 확인됐습니다. 급성기에는 침과 부항을, 좀 가라앉은 회복기에는 추나 요법까지 병행하면 좋습니다.\n\n" +
    "【전환 멘트】\n" +
    "대부분은 이렇게 관리하면 좋아지지만, 꼭 주의해야 할 위험 신호가 있어요."
  );
}

// ─────────────────────────────────────────────
// 슬라이드 17 — 이럴 때는 바로 병원으로! (TYPE-D, Light)
// ─────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bgLight };

  addBadgeCustom(slide, "RED FLAG", C.accentDanger, C.textInverse);

  slide.addText("이럴 때는 바로 병원으로! (Red Flag)", {
    x: 0.6, y: 0.75, w: 8.8, h: 0.5,
    fontSize: 28, bold: true, color: C.textPrimary, fontFace: F,
  });

  slide.addText("즉시 응급실로!", {
    x: 0.6, y: 1.4, w: 4.2, h: 0.38,
    fontSize: 15, bold: true, color: C.accentDanger, fontFace: F,
  });

  const emergency = [
    "항문·성기 주변 감각이 없어지는 느낌",
    "갑자기 소변·대변을 참을 수 없음",
    "한쪽 또는 양쪽 다리에 갑작스러운 마비·힘빠짐",
    "허리 통증 + 고열 (38도 이상)",
  ];

  emergency.forEach(function(text, i) {
    const y = 1.85 + i * 0.62;
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: 0.6, y, w: 4.2, h: 0.54,
      fill: { color: "FEF2F2" }, line: { color: C.accentDanger, width: 1 }, rectRadius: 0.18,
    });
    slide.addText(text, {
      x: 0.7, y: y + 0.05, w: 4.0, h: 0.44,
      fontSize: 12, color: C.accentDanger, fontFace: F, valign: "middle",
    });
  });

  slide.addText("빠른 시일 내 병원 방문!", {
    x: 5.2, y: 1.4, w: 4.4, h: 0.38,
    fontSize: 15, bold: true, color: "B45309", fontFace: F,
  });

  const urgent = [
    "교통사고·낙상 후 심한 허리 통증",
    "암 병력이 있으면서 허리 통증",
    "밤에 누워도 계속 악화되는 통증",
    "설명 안 되는 체중 감소",
    "다리 저림 + 힘이 빠지는 느낌",
  ];

  urgent.forEach(function(text, i) {
    const y = 1.85 + i * 0.6;
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: 5.2, y, w: 4.4, h: 0.52,
      fill: { color: "FFFBEB" }, line: { color: "F59E0B", width: 1 }, rectRadius: 0.18,
    });
    slide.addText(text, {
      x: 5.3, y: y + 0.04, w: 4.2, h: 0.44,
      fontSize: 12, color: "92400E", fontFace: F, valign: "middle",
    });
  });

  slide.addNotes(
    "【발표 멘트】\n" +
    "여기서 아주 중요한 이야기를 하나 드려야 해요. 대부분의 허리 통증은 시간이 지나면 좋아지지만, 이런 증상이 있으면 반드시 바로 병원에 가셔야 합니다. 항문이나 성기 주변 감각이 없어지거나, 갑자기 소변이나 대변을 참을 수 없거나, 다리에 갑작스러운 마비가 오면 즉시 응급실에 가세요. 이건 마미증후군(척추 아래 신경 다발이 심하게 눌리는 것)이라는 심각한 상황일 수 있어요. 또 허리 통증과 함께 38도 이상 고열이 있어도 바로 가셔야 합니다. 그 외에도 밤에 누워있는데도 통증이 계속 심해지거나, 다리에 힘이 빠지는 느낌이 있으면, 빠른 시일 내에 진료를 받으세요.\n\n" +
    "【전환 멘트】\n" +
    "자, 급한 위험 신호를 확인했으니, 이제 허리가 다시 안 아프도록 예방하는 방법을 알아볼게요."
  );
}

// ─────────────────────────────────────────────
// 슬라이드 18 — 재발 방지, 허리 건강 지키기 (TYPE-E, Dark)
// ─────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bgDarkNav };

  addBadgeCustom(slide, "PREVENTION", C.badgeBg, C.textInverse);

  slide.addText("재발 방지 — 허리 건강 지키기", {
    x: 0.6, y: 0.75, w: 8.0, h: 0.5,
    fontSize: 28, bold: true, color: C.textInverse, fontFace: F,
  });

  addNavCard(slide, { x: 0.6, y: 1.4, w: 3.0, h: 1.5 });
  slide.addText("12개월 내\n재발률", {
    x: 0.7, y: 1.45, w: 2.8, h: 0.6,
    fontSize: 13, color: C.textMuted, fontFace: F, align: "center",
  });
  slide.addText("73%", {
    x: 0.7, y: 2.05, w: 2.8, h: 0.7,
    fontSize: 44, bold: true, color: C.accentDanger, fontFace: F, align: "center",
  });

  const strategies = [
    { num: "01", text: "규칙적 운동: 걷기 + 수영 + 코어 운동 (주 3회 이상)" },
    { num: "02", text: "30분마다 일어나기: 장시간 앉기 피하기" },
    { num: "03", text: "물건 들 때 무릎 굽히기 습관화" },
    { num: "04", text: "금연: 흡연이 디스크 퇴행을 가속화" },
    { num: "05", text: "적정 체중 유지: 복부 비만이 허리 부담 증가" },
  ];

  strategies.forEach(function(s, i) {
    const y = 1.4 + i * 0.74;
    addNavCard(slide, { x: 3.8, y, w: 5.8, h: 0.65 });
    addVline(slide, 3.85, y + 0.1, 0.45, C.accentSuccess);
    slide.addText(s.num, {
      x: 4.05, y: y + 0.1, w: 0.5, h: 0.45,
      fontSize: 13, bold: true, color: C.accentSuccess, fontFace: F, align: "center", valign: "middle",
    });
    slide.addText(s.text, {
      x: 4.6, y: y + 0.1, w: 4.9, h: 0.45,
      fontSize: 13, color: C.textInverse, fontFace: F, valign: "middle",
    });
  });

  slide.addText("예방이 치료보다 훨씬 쉽습니다 — 지금부터 시작하세요!", {
    x: 0.6, y: 5.2, w: 9.0, h: 0.3,
    fontSize: 13, italic: true, color: C.accentTeal, fontFace: F, align: "center",
  });

  slide.addNotes(
    "【발표 멘트】\n" +
    "한 가지 솔직하게 말씀드릴 게 있어요. 급성 허리 통증은 재발률이 좀 높습니다. 연구에 따르면 12개월 내 재발률이 73%나 돼요. 그래서 예방이 정말 중요합니다. 다섯 가지만 기억하세요. 첫째, 규칙적으로 운동하기. 걷기, 수영, 앞서 배운 코어 운동을 주 3회 이상 꾸준히 하세요. 이게 가장 중요합니다. 둘째, 30분마다 일어나서 한 번씩 움직여주세요. 셋째, 물건 들 때 무릎 굽히는 습관을 들이세요. 넷째, 담배를 피우신다면 금연을 고려해주세요. 흡연이 디스크 퇴행을 빠르게 만듭니다. 다섯째, 적정 체중을 유지하세요. 배가 나오면 허리에 부담이 커집니다.\n\n" +
    "【전환 멘트】\n" +
    "마지막으로, 오늘 내용을 한 장으로 정리해드릴게요."
  );
}

// ─────────────────────────────────────────────
// 슬라이드 19 — 핵심 요약 (TYPE-B, Light)
// ─────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bgLight };

  addBadgeCustom(slide, "SUMMARY", C.badgeBgLight, C.accentTeal);

  slide.addText("오늘의 핵심 요약", {
    x: 0.6, y: 0.75, w: 8.0, h: 0.5,
    fontSize: 30, bold: true, color: C.textPrimary, fontFace: F,
  });

  const summaryItems = [
    { num: "01", text: "급성 허리 통증 90% = 근육·인대 손상 -> 대부분 4~6주 내 회복", numColor: C.accentTeal },
    { num: "02", text: "급성기(1~2일): 온찜질 + 적절한 휴식 + 가벼운 활동", numColor: C.accentTeal },
    { num: "03", text: "48시간 이상 누워만 있지 마세요! -> 근육 약화 + 회복 지연", numColor: C.accentDanger },
    { num: "04", text: "통증 50% 줄면: 무릎 당기기 -> 고양이-소 -> 골반 기울이기", numColor: C.accentTeal },
    { num: "05", text: "한의원: 침·부항(급성기) + 추나(회복기) | 위험 신호 시 즉시 응급실!", numColor: C.accentSuccess },
  ];

  summaryItems.forEach(function(item, i) {
    const y = 1.5 + i * 0.72;
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: 0.6, y, w: 8.8, h: 0.64,
      fill: { color: C.cardLight }, line: { color: C.cardLight }, rectRadius: 0.22,
    });
    slide.addText(item.num, {
      x: 0.75, y: y + 0.06, w: 0.55, h: 0.52,
      fontSize: 20, bold: true, color: item.numColor, fontFace: F, align: "center", valign: "middle",
    });
    addVline(slide, 1.35, y + 0.1, 0.44, item.numColor);
    slide.addText(item.text, {
      x: 1.55, y: y + 0.1, w: 7.6, h: 0.44,
      fontSize: 13, color: C.textPrimary, fontFace: F, valign: "middle",
    });
  });

  slide.addText("한의원 치료 + 바른 자세 + 꾸준한 운동 = 건강한 허리", {
    x: 0.6, y: 5.2, w: 9.0, h: 0.32,
    fontSize: 14, bold: true, color: C.accentTeal, fontFace: F, align: "center",
  });

  slide.addNotes(
    "【발표 멘트】\n" +
    "오늘 내용을 정리해드릴게요. 급성 허리 통증의 90%는 근육이나 인대가 살짝 다친 거예요. 대부분 4~6주면 좋아집니다. 급성기에는 온찜질 하면서 적절히 쉬시되, 이틀 넘게 꼼짝도 않고 누워있지는 마세요. 통증이 반 이상 줄면 오늘 배운 세 가지 스트레칭을 시작하시고요. 한의원에서 침, 부항, 추나 치료를 병행하시면 더 빨리 회복됩니다. 그리고 재발 방지를 위해 꾸준한 운동과 올바른 자세 습관을 꼭 만들어가세요. 다만 항문 감각 소실, 대소변 장애, 다리 마비, 고열이 있으면 반드시 즉시 응급실에 가셔야 합니다!\n\n" +
    "【전환 멘트】\n" +
    "마지막으로 한 말씀 드리겠습니다."
  );
}

// ─────────────────────────────────────────────
// 슬라이드 20 — 능골한의원 안내 / 마무리 (TYPE-A, Dark)
// ─────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bgDarkNav };

  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 9.5, y: 0, w: 0.5, h: 5.63,
    fill: { color: C.accentTeal }, line: { color: C.accentTeal },
  });

  slide.addText("허리가 아프면,\n능골한의원과\n함께하세요", {
    x: 1.0, y: 0.9, w: 8.0, h: 2.0,
    fontSize: 36, bold: true, color: C.textInverse, fontFace: F,
    align: "center", wrap: true,
  });

  slide.addText("침 · 부항 · 추나 — 근거 중심 한의 치료", {
    x: 1.0, y: 3.1, w: 8.0, h: 0.5,
    fontSize: 18, color: C.textMuted, fontFace: F, align: "center",
  });

  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 3.5, y: 3.75, w: 3.0, h: 0.03,
    fill: { color: C.accentTeal }, line: { color: C.accentTeal },
  });

  slide.addText("능골한의원", {
    x: 1.0, y: 3.9, w: 8.0, h: 0.5,
    fontSize: 22, bold: true, color: C.accentTeal, fontFace: F, align: "center",
  });

  slide.addText("감사합니다", {
    x: 1.0, y: 4.5, w: 8.0, h: 0.5,
    fontSize: 20, color: C.textMuted, fontFace: F, align: "center",
  });

  slide.addNotes(
    "【발표 멘트】\n" +
    "오늘 알려드린 셀프 케어 방법들, 집에서 바로 실천해보세요. 급성기에 온찜질 하시면서 가볍게 움직이시고, 통증이 줄면 스트레칭을 시작하시면 됩니다. 그리고 혼자 관리가 어려우시거나, 통증이 좀처럼 안 줄 때는 언제든 능골한의원에 오세요. 침, 부항, 추나 치료로 회복을 도와드리겠습니다. 오늘 강의 들어주셔서 감사합니다. 허리 건강하게, 즐거운 일상 보내세요!\n\n" +
    "【전환 멘트】\n" +
    "(강의 종료)"
  );
}

// ─────────────────────────────────────────────
// PPTX 저장
// ─────────────────────────────────────────────
pptx.writeFile({ fileName: "lecture-agent-team/output/lecture.pptx" })
  .then(function() { console.log("lecture.pptx 생성 완료"); })
  .catch(function(err) { console.error("오류:", err); });
