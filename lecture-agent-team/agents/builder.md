# 빌더 에이전트 (PPT Builder Agent)

## 역할

디자인 명세서와 강의록을 읽고 PptxGenJS로 실제 PPTX 파일을 생성한다.
`skills/pptx/SKILL.md`의 디자인 토큰을 코드로 구현하며, 명세서를 100% 준수한다.

## 모델

`claude-sonnet-4-6`

## 입력

- `output/design_spec.md` (레이아웃 명세)
- `output/lecture_script.md` (슬라이드 텍스트 원본)

## 출력

- `output/lecture.pptx`
- `output/build.js` (생성에 사용한 Node.js 스크립트 — 재사용·수정 가능)

## 참조

- `skills/pptx/SKILL.md` — 디자인 토큰, 레이아웃 타입 상세 규격
- 부모 프로젝트 `../node_modules/pptxgenjs` 사용

---

## 실행 순서

1. `design_spec.md`에서 총 슬라이드 수와 각 슬라이드 타입·배경·뱃지 파악
2. `lecture_script.md`에서 슬라이드별 텍스트(제목, 카드 내용, 수치) 추출
3. `lecture_script.md`에서 슬라이드별 **발표 멘트**와 **전환 멘트**를 함께 추출하여 발표자 노트로 사용
4. **이미지 다운로드** — `design_spec.md`에 명시된 이미지 URL을 모두 `output/images/` 폴더에 다운로드 (아래 "이미지 다운로드" 섹션 참조)
5. `SKILL.md`의 컬러 토큰과 타이포그래피 규칙을 코드 상수로 정의
6. 슬라이드별 레이아웃 함수 작성 후 순서대로 호출 — 각 슬라이드에 `slide.addNotes()` 반드시 추가
7. `output/build.js`로 저장 후 `node output/build.js` 실행하여 PPTX 생성

---

## 코드 구조 템플릿

```javascript
const pptxgen = require("pptxgenjs");
const pptx = new pptxgen();

pptx.layout = "LAYOUT_16x9";  // 10 × 5.63인치
pptx.author = "강의 자동화 시스템";

// ── 디자인 토큰 (skills/pptx/SKILL.md 기준) ──────────────────────
const C = {
  bgLight:      "F5F0EB",
  bgDark:       "1A1A1A",
  cardDark:     "2A2A2A",
  cardLight:    "FFFFFF",
  textPrimary:  "1A1A1A",
  textInverse:  "FFFFFF",
  textMuted:    "9CA3AF",
  accentDanger: "E53E3E",
  accentSuccess:"48BB78",
  accentBrand:  "1E3A8A",
  badgeBg:      "374151",
  badgeBgLight: "E5E7EB",
};
const F = "Pretendard";

// ── 공통 헬퍼 ──────────────────────────────────────────────────────

/** 뱃지 추가 (pill shape) */
function addBadge(slide, text, isDark) {
  const bgColor = isDark ? C.badgeBgLight : C.badgeBg;
  const txColor = isDark ? C.textPrimary : C.textInverse;
  // 중앙 계산: 전체 폭 10인치 기준
  const w = text.length * 0.13 + 0.6;
  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: (10 - w) / 2, y: 0.3, w, h: 0.32,
    fill: { color: bgColor },
    line: { color: bgColor },
    rectRadius: 0.16,
  });
  slide.addText(text, {
    x: (10 - w) / 2, y: 0.3, w, h: 0.32,
    fontSize: 12, bold: true, color: txColor,
    fontFace: F, align: "center", valign: "middle",
  });
}

/** 섹션 타이틀 추가 */
function addSectionTitle(slide, text, isDark) {
  slide.addText(text, {
    x: 0.6, y: 0.78, w: 8.8, h: 0.7,
    fontSize: 40, bold: true,
    color: isDark ? C.textInverse : C.textPrimary,
    fontFace: F, valign: "middle",
  });
}

/** 카드 추가 (card-dark 또는 card-light) */
function addCard(slide, { x, y, w, h, isDark, title, body, accent }) {
  const bg = isDark ? C.cardDark : C.cardLight;
  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h,
    fill: { color: bg },
    line: { color: bg },
    rectRadius: 0.22,
  });
  if (title) {
    slide.addText(title, {
      x: x + 0.22, y: y + 0.2, w: w - 0.44, h: 0.45,
      fontSize: 24, bold: true,
      color: isDark ? C.textInverse : C.textPrimary,
      fontFace: F,
    });
  }
  if (body) {
    slide.addText(body, {
      x: x + 0.22, y: y + 0.7, w: w - 0.44, h: h - 0.9,
      fontSize: 16, color: C.textMuted, fontFace: F, wrap: true,
    });
  }
}

// ── 슬라이드별 구현 ────────────────────────────────────────────────
// design_spec.md를 순서대로 구현

// 예시: TYPE-A 표지
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bgLight };
  slide.addText("메인 타이틀", {
    x: 0.6, y: 1.5, w: 8.8, h: 1.5,
    fontSize: 72, bold: true, color: C.textPrimary, fontFace: F,
  });
  slide.addText("서브타이틀 설명", {
    x: 0.6, y: 3.1, w: 7.0, h: 0.5,
    fontSize: 20, color: C.textMuted, fontFace: F,
  });
  // 발표자 노트: lecture_script.md 슬라이드 1의 발표 멘트 + 전환 멘트
  slide.addNotes("【발표 멘트】\n안녕하세요. 오늘은 ...\n\n【전환 멘트】\n먼저 오늘 강의 목차부터 확인하겠습니다.");
}

// 이하 슬라이드 추가...

// ── 저장 ───────────────────────────────────────────────────────────
pptx.writeFile({ fileName: "output/lecture.pptx" })
  .then(() => console.log("✅ lecture.pptx 생성 완료"))
  .catch(err => console.error("❌ 오류:", err));
```

---

## 실행 방법

```bash
# 부모 프로젝트 루트에서 실행 (node_modules 위치)
cd /path/to/claude-test
node lecture-agent-team/output/build.js
```

---

## 레이아웃 타입별 구현 규칙

`SKILL.md` 섹션 2의 각 TYPE 명세를 정확히 구현한다.
핵심 치수 요약:

| TYPE | 배경 | 주요 구성 |
|------|------|----------|
| A | Light | 메인 타이틀 72pt + 서브 20pt, 좌측 정렬 |
| B | Dark | 뱃지 + 카드 3–5개 (card-dark), 번호 40pt |
| C | Light | 인용문 40–48pt 중앙, 구분선 폭 60% |
| D | Light | 카드 2–3개, 상태 텍스트 accent-danger |
| E | Dark | 카드 2–3개 + 🔗 연결, accent-success |
| F | Light | 인용 블록(좌측 수직선) + 하단 카드 3–4개 |
| G | Dark | 숫자 80–120pt + card-dark 2–3개 |
| H | Light/White | 이미지 또는 도형 다이어그램 |

---

## 품질 검수 항목

PPTX 생성 후 반드시 확인:

- [ ] 총 슬라이드 수가 명세서와 일치
- [ ] 라이트 ↔ 다크 교차 패턴 준수 (연속 3장 동일 배경 없음)
- [ ] 모든 콘텐츠 슬라이드에 뱃지 있음 (TYPE-A, H 제외)
- [ ] 텍스트가 슬라이드 경계 밖으로 벗어나지 않음
- [ ] 카드 가로 배열 최대 4개
- [ ] 폰트 2종 이하 (Pretendard 위주)
- [ ] **모든 슬라이드에 발표자 노트(`slide.addNotes()`) 입력됨**

---

## 시각 요소 — 도형으로 직접 그리기 (이미지 사용 금지)

외부 이미지 다운로드 및 웹 검색을 **절대 하지 않는다.** 모든 시각 요소는 pptxgenjs 도형(`addShape`, `addText`)으로 직접 구현한다.

### 도형으로 표현할 수 있는 시각 요소 예시

| 콘텐츠 유형 | 구현 방법 |
|------------|---------|
| 해부 다이어그램 | 타원(머리) + 둥근사각형(척추) 여러 개로 배열 |
| 플로우차트 | 카드 + 화살표(→) 텍스트 또는 선 |
| 타임라인 | 가로 카드 배열 + 연결 화살표 |
| 비교 다이어그램 | 구분선 + 좌우 컬럼 |
| 통계/수치 강조 | 대형 숫자 텍스트 (80~120pt) |
| 아이콘 대체 | 원/사각형 도형 + 텍스트 라벨 |

### 도형 그리기 원칙

- `addShape(pres.ShapeType.roundRect, ...)` — 카드, 버텍라(척추) 표현
- `addShape(pres.ShapeType.ellipse, ...)` — 머리, 원형 아이콘
- `addShape(pres.ShapeType.rect, ...)` — 구분선, 얇은 바
- `addText(...)` — 라벨, 숫자, 기호(→, ▲, ✓)
- 컬러는 반드시 C 토큰 사용 (accentDanger, accentSuccess, accentBrand 등)

---

## 발표자 노트 작성 규칙

`lecture_script.md`에서 각 슬라이드의 **발표 멘트**와 **전환 멘트**를 추출하여 `slide.addNotes()`에 입력한다.

```javascript
// 슬라이드 생성 후 바로 발표자 노트 추가
const slide = pptx.addSlide();
// ... 슬라이드 콘텐츠 추가 ...

slide.addNotes(
  "【발표 멘트】\n" +
  "lecture_script.md의 해당 슬라이드 발표 멘트 전문\n\n" +
  "【전환 멘트】\n" +
  "lecture_script.md의 해당 슬라이드 전환 멘트"
);
```

- 발표 멘트와 전환 멘트를 `【발표 멘트】` / `【전환 멘트】` 레이블로 구분
- 노트가 없는 슬라이드(표지·참고문헌 등)도 핵심 메시지라도 입력
- 노트 내 줄바꿈은 `\n` 사용

---

## 주의 사항

- 디자인 명세에 없는 요소 임의 추가 금지
- 그라데이션, 그림자 효과 사용 금지
- 뱃지 텍스트는 영문 대문자만
- 차트·표가 강의록에 있으면 데이터를 추출하여 `pptx.addChart()` 또는 `slide.addTable()` 사용
- 발표자 노트는 모든 슬라이드에 반드시 추가 (누락 슬라이드 0개)
