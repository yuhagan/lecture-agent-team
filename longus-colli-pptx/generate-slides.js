const pptxgen = require("pptxgenjs");
const pptx = new pptxgen();

pptx.layout = "LAYOUT_16x9";
pptx.author = "의학 블로그";
pptx.title = "목긴근(Longus Colli)의 해부학";

// ─── 폰트 & 컬러 상수 ────────────────────────────────────────────────
const F = "Noto Sans KR";
const C = {
  navy:    "1B2A4A",
  teal:    "0A7E8C",
  white:   "FFFFFF",
  dark:    "2D2D2D",
  mid:     "6B7280",
  ltTeal:  "E6F3F5",
  ltGray:  "D1D5DB",
  bgGray:  "F0F0F0",
  red:     "DC3545",
};

// ─── 헬퍼: 좌측 틸블루 세로 바 ─────────────────────────────────────
function addAccentBar(slide, h = 5.63) {
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.08, h,
    fill: { color: C.teal },
    line: { color: C.teal },
  });
}

// ─── 헬퍼: 슬라이드 제목 ────────────────────────────────────────────
function addSlideTitle(slide, text) {
  addAccentBar(slide);
  slide.addText(text, {
    x: 0.22, y: 0.25, w: 9.3, h: 0.65,
    fontSize: 28, bold: true, color: C.navy,
    fontFace: F, valign: "middle",
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.22, y: 0.85, w: 9.3, h: 0.03,
    fill: { color: C.ltGray },
    line: { color: C.ltGray },
  });
}

// ════════════════════════════════════════════════════════════════════
// 슬라이드 1 — 표지
// ════════════════════════════════════════════════════════════════════
{
  const slide = pptx.addSlide();
  slide.background = { color: C.navy };

  // 좌측 틸블루 악센트 바
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.12, h: 5.63,
    fill: { color: C.teal },
    line: { color: C.teal },
  });

  // 메인 제목
  slide.addText("목긴근(Longus Colli)의 해부학", {
    x: 0.4, y: 1.5, w: 9.1, h: 1.3,
    fontSize: 40, bold: true, color: C.white,
    fontFace: F, valign: "middle", align: "center",
  });

  // 부제목
  slide.addText("전척주근(Prevertebral Muscle) | 경추 심부 굴곡근", {
    x: 0.4, y: 2.9, w: 9.1, h: 0.6,
    fontSize: 18, color: C.teal,
    fontFace: F, align: "center",
  });

  // 구분선
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 2.5, y: 3.6, w: 4.8, h: 0.04,
    fill: { color: C.teal },
    line: { color: C.teal },
  });

  // 날짜 / 블로그명
  slide.addText("2026  |  의학 블로그", {
    x: 0.4, y: 4.3, w: 9.1, h: 0.4,
    fontSize: 13, color: "99AABB",
    fontFace: F, align: "center",
  });
}

// ════════════════════════════════════════════════════════════════════
// 슬라이드 2 — 목차
// ════════════════════════════════════════════════════════════════════
{
  const slide = pptx.addSlide();
  slide.background = { color: C.white };
  addSlideTitle(slide, "목차");

  const items = [
    { num: "01", text: "목긴근 개요 및 위치" },
    { num: "02", text: "세 부분의 구조 (상사·수직·하사)" },
    { num: "03", text: "기시(Origin) & 정지(Insertion) 상세" },
    { num: "04", text: "작용(Action) 및 기능" },
    { num: "05", text: "신경지배(Innervation) 및 임상적 의의" },
  ];

  items.forEach((item, i) => {
    const y = 1.2 + i * 0.82;

    // 번호 배지
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 0.5, y, w: 0.55, h: 0.52,
      fill: { color: C.teal },
      line: { color: C.teal },
      rectRadius: 0.06,
    });
    slide.addText(item.num, {
      x: 0.5, y, w: 0.55, h: 0.52,
      fontSize: 16, bold: true, color: C.white,
      fontFace: F, align: "center", valign: "middle",
    });

    // 항목명
    slide.addText(item.text, {
      x: 1.2, y: y + 0.02, w: 8.2, h: 0.48,
      fontSize: 18, color: C.dark,
      fontFace: F, valign: "middle",
    });

    // 구분선 (마지막 제외)
    if (i < items.length - 1) {
      slide.addShape(pptx.shapes.RECTANGLE, {
        x: 0.5, y: y + 0.62, w: 8.9, h: 0.01,
        fill: { color: C.ltGray },
        line: { color: C.ltGray },
      });
    }
  });
}

// ════════════════════════════════════════════════════════════════════
// 슬라이드 3 — 개요 & 해부학 이미지 (레이아웃 A: 2컬럼)
// ════════════════════════════════════════════════════════════════════
{
  const slide = pptx.addSlide();
  slide.background = { color: C.white };
  addSlideTitle(slide, "01  목긴근(Longus Colli) 개요");

  // 좌측 — 해부학 이미지
  slide.addImage({
    path: "./longus_colli_highlighted.png",
    x: 0.35, y: 1.05, w: 4.5, h: 4.2,
    sizing: { type: "contain", w: 4.5, h: 4.2 },
  });
  slide.addText("Gray's Anatomy Plate 378 (수정본) | Public Domain", {
    x: 0.35, y: 5.15, w: 4.5, h: 0.3,
    fontSize: 9, color: C.mid, fontFace: F,
    italic: true, align: "center",
  });

  // 우측 — 개요 텍스트 블록
  const overviewItems = [
    { label: "분류", value: "전척주근(Prevertebral Muscle)\n심부 경추 굴곡근(Deep Cervical Flexor)" },
    { label: "위치", value: "경추(Cervical Spine) ~ 흉추 3번(T3)\n척추체 전면(Anterior Surface of Vertebral Bodies)" },
    { label: "형태", value: "3개 부분으로 구성된 복합근\n중간이 넓고 양 끝이 가늘어지는 방추형" },
    { label: "중요성", value: "경추 심부 안정화의 핵심 근육\n두부 전방 자세(FHP) 교정에 필수" },
  ];

  overviewItems.forEach((item, i) => {
    const y = 1.08 + i * 1.02;

    // 레이블 배지
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 5.1, y, w: 1.4, h: 0.38,
      fill: { color: C.teal },
      line: { color: C.teal },
    });
    slide.addText(item.label, {
      x: 5.1, y, w: 1.4, h: 0.38,
      fontSize: 12, bold: true, color: C.white,
      fontFace: F, align: "center", valign: "middle",
    });

    // 내용
    slide.addText(item.value, {
      x: 6.6, y: y - 0.02, w: 3.0, h: 0.88,
      fontSize: 12, color: C.dark,
      fontFace: F, valign: "middle",
    });

    // 구분선
    if (i < overviewItems.length - 1) {
      slide.addShape(pptx.shapes.RECTANGLE, {
        x: 5.1, y: y + 0.9, w: 4.5, h: 0.01,
        fill: { color: C.ltGray },
        line: { color: C.ltGray },
      });
    }
  });
}

// ════════════════════════════════════════════════════════════════════
// 슬라이드 4 — 세 부분의 구조 (레이아웃 C: 다이어그램)
// ════════════════════════════════════════════════════════════════════
{
  const slide = pptx.addSlide();
  slide.background = { color: C.white };
  addSlideTitle(slide, "02  세 부분의 구조(Three Portions)");

  // 중앙 척추 바
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 4.55, y: 1.1, w: 0.7, h: 4.2,
    fill: { color: C.ltTeal },
    line: { color: C.teal, width: 1.5 },
  });
  slide.addText("척추체\n(Vertebral\nBodies)", {
    x: 4.55, y: 2.6, w: 0.7, h: 0.9,
    fontSize: 9, color: C.teal, bold: true,
    fontFace: F, align: "center", valign: "middle",
  });

  // 척추 레벨 표시
  const levels = [
    { label: "C1", y: 1.15 },
    { label: "C3", y: 1.65 },
    { label: "C5", y: 2.15 },
    { label: "T1", y: 2.95 },
    { label: "T3", y: 3.55 },
  ];
  levels.forEach(lv => {
    slide.addText(lv.label, {
      x: 4.58, y: lv.y, w: 0.64, h: 0.28,
      fontSize: 10, bold: true, color: C.navy,
      fontFace: F, align: "center", valign: "middle",
    });
  });

  // ── 상사부 (Superior Oblique) ──
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.4, y: 1.1, w: 3.8, h: 1.2,
    fill: { color: "FFF8E7" },
    line: { color: "F59E0B", width: 1.5 },
  });
  slide.addText("상사부(Superior Oblique Part)", {
    x: 0.5, y: 1.12, w: 3.6, h: 0.32,
    fontSize: 13, bold: true, color: "92400E",
    fontFace: F,
  });
  slide.addText("기시: C3~C5 횡돌기 전결절\n정지: C1(환추) 전결절", {
    x: 0.5, y: 1.44, w: 3.6, h: 0.64,
    fontSize: 11, color: C.dark,
    fontFace: F,
  });
  // 연결선
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 4.2, y: 1.45, w: 0.38, h: 0.04,
    fill: { color: "F59E0B" },
    line: { color: "F59E0B" },
  });

  // ── 수직부 (Vertical Part) ──
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.4, y: 2.45, w: 3.8, h: 1.15,
    fill: { color: "F0FFF4" },
    line: { color: C.teal, width: 1.5 },
  });
  slide.addText("수직부(Vertical Part)", {
    x: 0.5, y: 2.47, w: 3.6, h: 0.32,
    fontSize: 13, bold: true, color: "065F46",
    fontFace: F,
  });
  slide.addText("기시: C5~T3 추체 전면\n정지: C2~C4 추체 전면", {
    x: 0.5, y: 2.79, w: 3.6, h: 0.6,
    fontSize: 11, color: C.dark,
    fontFace: F,
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 4.2, y: 2.92, w: 0.38, h: 0.04,
    fill: { color: C.teal },
    line: { color: C.teal },
  });

  // ── 하사부 (Inferior Oblique) ──
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.4, y: 3.75, w: 3.8, h: 1.2,
    fill: { color: "FFF0F0" },
    line: { color: C.red, width: 1.5 },
  });
  slide.addText("하사부(Inferior Oblique Part)", {
    x: 0.5, y: 3.77, w: 3.6, h: 0.32,
    fontSize: 13, bold: true, color: "9B1C1C",
    fontFace: F,
  });
  slide.addText("기시: T1~T3 추체 전면\n정지: C5~C6 횡돌기 전결절", {
    x: 0.5, y: 4.09, w: 3.6, h: 0.64,
    fontSize: 11, color: C.dark,
    fontFace: F,
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 4.2, y: 4.3, w: 0.38, h: 0.04,
    fill: { color: C.red },
    line: { color: C.red },
  });

  // ── 우측 이미지 ──
  slide.addImage({
    path: "./longus_colli.png",
    x: 5.6, y: 1.1, w: 3.9, h: 4.2,
    sizing: { type: "contain", w: 3.9, h: 4.2 },
  });
  slide.addText("Gray387 | Wikimedia Commons | Public Domain", {
    x: 5.6, y: 5.2, w: 3.9, h: 0.28,
    fontSize: 9, color: C.mid, fontFace: F,
    italic: true, align: "center",
  });
}

// ════════════════════════════════════════════════════════════════════
// 슬라이드 5 — 기시 & 정지 상세 표 (레이아웃 B)
// ════════════════════════════════════════════════════════════════════
{
  const slide = pptx.addSlide();
  slide.background = { color: C.white };
  addSlideTitle(slide, "03  기시(Origin) & 정지(Insertion) 상세");

  const tableData = [
    // 헤더
    [
      { text: "부분(Part)", options: { bold: true, color: C.white, fill: C.navy, fontFace: F, fontSize: 13, align: "center", valign: "middle" } },
      { text: "기시(Origin)", options: { bold: true, color: C.white, fill: C.teal, fontFace: F, fontSize: 13, align: "center", valign: "middle" } },
      { text: "정지(Insertion)", options: { bold: true, color: C.white, fill: C.teal, fontFace: F, fontSize: 13, align: "center", valign: "middle" } },
      { text: "주요 특징", options: { bold: true, color: C.white, fill: C.teal, fontFace: F, fontSize: 13, align: "center", valign: "middle" } },
    ],
    // 상사부
    [
      { text: "상사부\n(Superior\nOblique)", options: { bold: true, color: C.navy, fill: "FFF8E7", fontFace: F, fontSize: 12, align: "center", valign: "middle" } },
      { text: "C3~C5\n횡돌기(Transverse Process)\n전결절(Anterior Tubercle)", options: { color: C.dark, fill: C.white, fontFace: F, fontSize: 11, valign: "middle" } },
      { text: "C1(환추, Atlas)\n전결절(Anterior Tubercle)", options: { color: C.dark, fill: C.white, fontFace: F, fontSize: 11, valign: "middle" } },
      { text: "환추 전방 안정화\n두부 굴곡 시 주요 역할", options: { color: C.dark, fill: C.white, fontFace: F, fontSize: 11, valign: "middle" } },
    ],
    // 수직부
    [
      { text: "수직부\n(Vertical\nPart)", options: { bold: true, color: "065F46", fill: "F0FFF4", fontFace: F, fontSize: 12, align: "center", valign: "middle" } },
      { text: "C5~T3\n추체 전면\n(Anterior Body Surface)", options: { color: C.dark, fill: "FAFAFA", fontFace: F, fontSize: 11, valign: "middle" } },
      { text: "C2~C4\n추체 전면\n(Anterior Body Surface)", options: { color: C.dark, fill: "FAFAFA", fontFace: F, fontSize: 11, valign: "middle" } },
      { text: "경추-흉추 연결 핵심부\n중간 경추 안정화 역할", options: { color: C.dark, fill: "FAFAFA", fontFace: F, fontSize: 11, valign: "middle" } },
    ],
    // 하사부
    [
      { text: "하사부\n(Inferior\nOblique)", options: { bold: true, color: "9B1C1C", fill: "FFF0F0", fontFace: F, fontSize: 12, align: "center", valign: "middle" } },
      { text: "T1~T3\n추체 전면\n(Anterior Body Surface)", options: { color: C.dark, fill: C.white, fontFace: F, fontSize: 11, valign: "middle" } },
      { text: "C5~C6\n횡돌기(Transverse Process)\n전결절(Anterior Tubercle)", options: { color: C.dark, fill: C.white, fontFace: F, fontSize: 11, valign: "middle" } },
      { text: "흉추-경추 연결\n하부 경추 측굴 보조", options: { color: C.dark, fill: C.white, fontFace: F, fontSize: 11, valign: "middle" } },
    ],
  ];

  slide.addTable(tableData, {
    x: 0.3, y: 1.05, w: 9.2,
    rowH: [0.42, 1.0, 1.0, 1.0],
    border: { type: "solid", color: C.ltGray, pt: 0.5 },
    colW: [1.5, 2.4, 2.4, 2.9],
  });

  // 하단 참고 박스
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.3, y: 4.55, w: 9.2, h: 0.78,
    fill: { color: C.ltTeal },
    line: { color: C.teal, width: 1 },
  });
  slide.addText(
    "💡  기억법: 상사부(上斜部)→위에서 대각선, 수직부(垂直部)→일직선, 하사부(下斜部)→아래에서 대각선. 세 부분이 합쳐져 경추 전면을 V자 형태로 덮는다.",
    {
      x: 0.45, y: 4.6, w: 8.9, h: 0.65,
      fontSize: 11, color: C.navy, fontFace: F, valign: "middle",
    }
  );
}

// ════════════════════════════════════════════════════════════════════
// 슬라이드 6 — 작용(Action) 텍스트 도해
// ════════════════════════════════════════════════════════════════════
{
  const slide = pptx.addSlide();
  slide.background = { color: C.white };
  addSlideTitle(slide, "04  작용(Action) 및 기능");

  // 3개 주요 작용 카드
  const actions = [
    {
      icon: "①",
      title: "경추 굴곡\n(Cervical Flexion)",
      desc: "목을 앞으로 숙이는 동작\n수직부가 주로 담당\n심부 경추 굴곡근 중 가장 중요",
      color: C.teal,
      fill: C.ltTeal,
      x: 0.3,
    },
    {
      icon: "②",
      title: "경추 측굴\n(Lateral Flexion)",
      desc: "목을 좌우로 기울이는 동작\n편측 수축 시 동측 측굴\n상사부·하사부가 주로 기여",
      color: "065F46",
      fill: "F0FFF4",
      x: 3.6,
    },
    {
      icon: "③",
      title: "경추 회전\n(Cervical Rotation)",
      desc: "목을 회전시키는 보조 동작\n하사부·상사부가 기여\n단독 회전근보다 보조 역할",
      color: "9B1C1C",
      fill: "FFF0F0",
      x: 6.9,
    },
  ];

  actions.forEach(a => {
    // 카드 배경
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: a.x, y: 1.1, w: 3.0, h: 3.5,
      fill: { color: a.fill },
      line: { color: a.color, width: 1.5 },
    });
    // 아이콘 원
    slide.addShape(pptx.shapes.OVAL, {
      x: a.x + 1.1, y: 1.2, w: 0.8, h: 0.8,
      fill: { color: a.color },
      line: { color: a.color },
    });
    slide.addText(a.icon, {
      x: a.x + 1.1, y: 1.2, w: 0.8, h: 0.8,
      fontSize: 18, bold: true, color: C.white,
      fontFace: F, align: "center", valign: "middle",
    });
    // 제목
    slide.addText(a.title, {
      x: a.x + 0.1, y: 2.12, w: 2.8, h: 0.72,
      fontSize: 13, bold: true, color: a.color,
      fontFace: F, align: "center", valign: "middle",
    });
    // 구분선
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: a.x + 0.2, y: 2.88, w: 2.6, h: 0.03,
      fill: { color: a.color },
      line: { color: a.color },
    });
    // 설명
    slide.addText(a.desc, {
      x: a.x + 0.1, y: 2.95, w: 2.8, h: 1.5,
      fontSize: 11, color: C.dark,
      fontFace: F, align: "center", valign: "top",
    });
  });

  // 하단 — 심부 경추 굴곡근 그룹
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.3, y: 4.75, w: 9.2, h: 0.72,
    fill: { color: C.navy },
    line: { color: C.navy },
  });
  slide.addText(
    "심부 경추 굴곡근(Deep Cervical Flexors) 그룹  |  목긴근(Longus Colli) · 머리긴근(Longus Capitis) · 앞머리곧은근(Rectus Capitis Anterior) · 가쪽머리곧은근(Rectus Capitis Lateralis)",
    {
      x: 0.45, y: 4.78, w: 8.9, h: 0.62,
      fontSize: 11, color: C.white,
      fontFace: F, valign: "middle", align: "center",
    }
  );
}

// ════════════════════════════════════════════════════════════════════
// 슬라이드 7 — 신경지배 & 임상적 의의
// ════════════════════════════════════════════════════════════════════
{
  const slide = pptx.addSlide();
  slide.background = { color: C.white };
  addSlideTitle(slide, "05  신경지배(Innervation) 및 임상적 의의");

  // ── 상단: 신경지배 흐름도 ──
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.3, y: 1.05, w: 9.2, h: 0.3,
    fill: { color: C.ltTeal },
    line: { color: C.teal },
  });
  slide.addText("신경지배(Innervation)", {
    x: 0.35, y: 1.05, w: 9.1, h: 0.3,
    fontSize: 13, bold: true, color: C.teal,
    fontFace: F, valign: "middle",
  });

  // 흐름도 박스들
  const flowItems = [
    { text: "경수(Cervical\nSpinal Cord)", w: 1.6 },
    { text: "→", w: 0.3, isArrow: true },
    { text: "경신경 전지\n(Anterior Rami\nC2 ~ C6)", w: 1.9 },
    { text: "→", w: 0.3, isArrow: true },
    { text: "직접 분지\n(Direct Branches)", w: 1.9 },
    { text: "→", w: 0.3, isArrow: true },
    { text: "목긴근\n(Longus Colli)", w: 1.9 },
  ];

  let fx = 0.3;
  flowItems.forEach(item => {
    if (item.isArrow) {
      slide.addText(item.text, {
        x: fx, y: 1.38, w: item.w, h: 0.7,
        fontSize: 20, bold: true, color: C.teal,
        fontFace: F, align: "center", valign: "middle",
      });
    } else {
      slide.addShape(pptx.shapes.RECTANGLE, {
        x: fx, y: 1.38, w: item.w, h: 0.7,
        fill: { color: C.ltTeal },
        line: { color: C.teal, width: 1 },
      });
      slide.addText(item.text, {
        x: fx, y: 1.38, w: item.w, h: 0.7,
        fontSize: 11, bold: true, color: C.navy,
        fontFace: F, align: "center", valign: "middle",
      });
    }
    fx += item.w + 0.02;
  });

  slide.addText("※ 경추 신경근(C2~C6) 손상 시 목긴근 약화 → 경추 불안정 유발 가능", {
    x: 0.3, y: 2.1, w: 9.2, h: 0.3,
    fontSize: 10, color: C.mid, fontFace: F, italic: true,
  });

  // ── 하단: 임상적 의의 ──
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.3, y: 2.5, w: 9.2, h: 0.3,
    fill: { color: C.navy },
    line: { color: C.navy },
  });
  slide.addText("임상적 의의(Clinical Significance)", {
    x: 0.35, y: 2.5, w: 9.1, h: 0.3,
    fontSize: 13, bold: true, color: C.white,
    fontFace: F, valign: "middle",
  });

  const clinicalItems = [
    {
      title: "전방 경추 수술(ACDF) 시 표지",
      desc: "전방 경추 추간판 제거 및 유합술(Anterior Cervical Discectomy and Fusion) 시\n목긴근을 박리하고 견인하여 추체 전면에 접근한다. 수술 중 핵심 해부학 표지.",
      color: C.teal,
    },
    {
      title: "두부 전방 자세(Forward Head Posture, FHP) 약화",
      desc: "장시간 스마트폰·컴퓨터 사용 시 목긴근이 약해지고 길어진 상태로 유지.\n심부 굴곡근 강화 운동(Deep Neck Flexor Training)으로 교정이 필요하다.",
      color: "065F46",
    },
    {
      title: "경추 안정화 및 목 통증",
      desc: "목긴근 약화는 경추 분절 불안정(Segmental Instability)을 초래하고\n목 통증(Neck Pain), 두통, 경추 디스크 질환과 밀접히 연관된다.",
      color: "9B1C1C",
    },
  ];

  clinicalItems.forEach((item, i) => {
    const y = 2.85 + i * 0.9;
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 0.3, y, w: 0.12, h: 0.75,
      fill: { color: item.color },
      line: { color: item.color },
    });
    slide.addText(item.title, {
      x: 0.52, y, w: 9.0, h: 0.3,
      fontSize: 12, bold: true, color: item.color,
      fontFace: F, valign: "middle",
    });
    slide.addText(item.desc, {
      x: 0.52, y: y + 0.3, w: 9.0, h: 0.45,
      fontSize: 11, color: C.dark,
      fontFace: F, valign: "top",
    });
  });
}

// ════════════════════════════════════════════════════════════════════
// 슬라이드 8 — 요약
// ════════════════════════════════════════════════════════════════════
{
  const slide = pptx.addSlide();
  slide.background = { color: C.white };
  addSlideTitle(slide, "요약(Summary)");

  const bullets = [
    { num: "1", text: "목긴근(Longus Colli)은 경추~T3를 잇는 전척주 심부근으로, 상사·수직·하사 3개 부분으로 구성된다." },
    { num: "2", text: "신경지배는 경신경 전지(C2~C6)이며, 경추 굴곡(Flexion) · 측굴(Lateral Flexion) · 회전(Rotation)을 담당한다." },
    { num: "3", text: "심부 경추 굴곡근(Deep Cervical Flexors) 그룹의 핵심 근육으로 경추 분절 안정성을 유지한다." },
    { num: "4", text: "두부 전방 자세(FHP) 개선과 경추 재활에서 목긴근 강화 운동이 필수적이다." },
    { num: "5", text: "전방 경추 수술(ACDF) 시 가장 중요한 해부학적 표지 근육 중 하나이다." },
  ];

  bullets.forEach((b, i) => {
    const y = 1.1 + i * 0.75;
    slide.addShape(pptx.shapes.OVAL, {
      x: 0.3, y: y + 0.05, w: 0.48, h: 0.48,
      fill: { color: C.teal },
      line: { color: C.teal },
    });
    slide.addText(b.num, {
      x: 0.3, y: y + 0.05, w: 0.48, h: 0.48,
      fontSize: 14, bold: true, color: C.white,
      fontFace: F, align: "center", valign: "middle",
    });
    slide.addText(b.text, {
      x: 0.9, y, w: 8.6, h: 0.62,
      fontSize: 13, color: C.dark,
      fontFace: F, valign: "middle",
    });
  });

  // 핵심 포인트 강조 박스
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.3, y: 4.82, w: 9.2, h: 0.65,
    fill: { color: C.teal },
    line: { color: C.teal },
  });
  slide.addText(
    "핵심 포인트  |  목긴근은 경추 건강의 '숨겨진 핵심근'  — 약해지면 목 통증·두통·자세 불량이 연쇄적으로 발생한다.",
    {
      x: 0.45, y: 4.85, w: 8.9, h: 0.55,
      fontSize: 12, bold: true, color: C.white,
      fontFace: F, valign: "middle", align: "center",
    }
  );
}

// ════════════════════════════════════════════════════════════════════
// 슬라이드 9 — 참고문헌
// ════════════════════════════════════════════════════════════════════
{
  const slide = pptx.addSlide();
  slide.background = { color: C.white };
  addSlideTitle(slide, "참고문헌(References)");

  const refs = [
    "Gray H. Anatomy of the Human Body. Philadelphia: Lea & Febiger; 1918. Plate 378.",
    "Drake RL, Vogl AW, Mitchell AWM. Gray's Anatomy for Students. 4th ed. Elsevier; 2019. pp.982-985.",
    "Moore KL, Dalley AF, Agur AMR. Clinically Oriented Anatomy. 8th ed. Wolters Kluwer; 2018. pp.476-478.",
    "Netter FH. Atlas of Human Anatomy. 7th ed. Elsevier; 2019. Plate 25.",
    "Falla DL, Jull GA, Hodges PW. Patients with neck pain demonstrate reduced electromyographic activity of the deep cervical flexor muscles during performance of the craniocervical flexion test. Spine. 2004;29(19):2108-2114.",
    "Image: Gray — musculus longus colli.png (Gray's Anatomy Plate 378, 목긴근 강조본)\n       Wikimedia Commons | Public Domain | https://commons.wikimedia.org/wiki/File:Gray_%E2%80%94_musculus_longus_colli.png",
    "Image: Longus colli.png (Gray387 기반 목긴근 단독 이미지)\n       Wikimedia Commons | Public Domain | https://commons.wikimedia.org/wiki/File:Longus_colli.png",
  ];

  refs.forEach((ref, i) => {
    const isImage = ref.startsWith("Image:");
    slide.addText(`${i + 1}.  ${ref}`, {
      x: 0.4, y: 1.05 + i * 0.6, w: 9.1, h: 0.52,
      fontSize: isImage ? 10 : 11,
      color: isImage ? C.teal : C.mid,
      fontFace: F, valign: "top",
      italic: isImage,
    });
  });
}

// ════════════════════════════════════════════════════════════════════
// 저장
// ════════════════════════════════════════════════════════════════════
pptx.writeFile({ fileName: "목긴근_Longus_Colli_해부학.pptx" })
  .then(() => console.log("✅  목긴근_Longus_Colli_해부학.pptx 생성 완료!"))
  .catch(err => console.error("❌ 오류:", err));
