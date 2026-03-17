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

## 이미지 다운로드 (build.js 작성 전 반드시 실행)

`design_spec.md`에 이미지 URL이 있는 슬라이드는 **build.js 코드 작성 전에** 반드시 로컬에 다운로드한다.

### 다운로드 절차

```bash
# 1. 이미지 폴더 생성
mkdir -p lecture-agent-team/output/images

# 2. 각 이미지 다운로드 (URL당 한 번씩 실행)
curl -L -o "lecture-agent-team/output/images/slide_N.png" "https://upload.wikimedia.org/..."

# 3. 다운로드 성공 여부 확인 (파일 크기가 0이면 실패)
ls -lh lecture-agent-team/output/images/
```

- 파일명은 슬라이드 번호 기준으로 통일: `slide_4.png`, `slide_7.jpg` 등
- 다운로드 실패 시 PLACEHOLDER로 대체 (아래 이미지 삽입 섹션 참조)
- SVG 형식은 pptxgenjs 호환 불가 → PNG 썸네일 URL로 대체

### build.js에서 로컬 이미지 삽입

```javascript
// 로컬 다운로드한 이미지 삽입 (URL 직접 사용 금지)
slide.addImage({
  path: "lecture-agent-team/output/images/slide_4.png",
  x: 5.5, y: 1.2, w: 4.0, h: 3.2,
  sizing: { type: "contain", w: 4.0, h: 3.2 },
});
```

---

## 이미지 삽입

`design_spec.md`에 이미지 명세가 있는 슬라이드는 `slide.addImage()`로 삽입한다.

```javascript
// URL 직접 삽입 (로컬 다운로드 불필요)
slide.addImage({
  path: "https://upload.wikimedia.org/wikipedia/commons/...",
  x: 5.5, y: 1.2, w: 4.0, h: 3.2,
});

// 캡션 (이미지 바로 아래)
slide.addText("그림 설명 — CC BY-SA 4.0", {
  x: 5.5, y: 4.5, w: 4.0, h: 0.25,
  fontSize: 9, color: C.textMuted, fontFace: F, align: "center",
});
```

### 이미지 공란 (PLACEHOLDER)

명세의 URL이 `PLACEHOLDER`이거나 다운로드에 실패한 경우 회색 사각형 + 안내 텍스트로 공란을 렌더링한다.

```javascript
// URL이 PLACEHOLDER인 경우
slide.addShape(pptx.shapes.RECTANGLE, {
  x: 5.5, y: 1.2, w: 4.0, h: 3.2,
  fill: { color: "E5E7EB" },
  line: { color: "9CA3AF", width: 1 },
});
slide.addText("🖼 이미지 삽입 필요\n[어떤 이미지가 필요한지 설명]", {
  x: 5.5, y: 1.2, w: 4.0, h: 3.2,
  fontSize: 13, color: "6B7280", fontFace: F,
  align: "center", valign: "middle", wrap: true,
});
```

### 주의
- **반드시 로컬 파일 경로 사용** — 원격 URL 직접 삽입 금지
- 이미지는 build.js 실행 전에 `output/images/` 폴더에 미리 다운로드해야 함
- 명세에 없는 이미지 임의 추가 금지
- SVG 파일은 pptxgenjs 호환 불가 → PNG 썸네일 URL로 대체 다운로드

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
