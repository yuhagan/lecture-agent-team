# 디자인 명세: 손저림의 한의학적 이해와 치료

작성일: 2026-03-24
디자인 시스템: skills/pptx/SKILL.md v1.0
총 슬라이드: 20장

---

## 슬라이드 1

- **타입**: TYPE-A
- **배경**: Light (bg-light #F5F0EB)
- **뱃지**: 없음 (TYPE-A 예외)
- **레이아웃**:
  - 메인 타이틀: "손저림의 한의학적 이해와 치료" — 64pt, Black(900), text-primary, 좌측 정렬
    - x=0.6, y=1.6, w=8.8, h=1.2
  - 서브타이틀: "원인 감별부터 임상 적용까지" — 22pt, Regular(400), text-muted
    - x=0.6, y=2.9, w=8.8, h=0.5
  - 강사명/소속: "강사명 | 소속" — 16pt, Regular(400), text-muted
    - x=0.6, y=3.6, w=8.8, h=0.4
  - 하단 장식선: accent-brand 컬러 바
    - x=0.6, y=4.8, w=2.0, h=0.06, fill=accentBrand
- **특수 요소**: 없음

---

## 슬라이드 2

- **타입**: TYPE-B
- **배경**: Dark (bg-dark #1A1A1A)
- **뱃지**: "ROADMAP" — pill shape, 상단 중앙, badgeBgLight + text-primary
- **레이아웃**:
  - 섹션 헤딩: "오늘의 강의 흐름" — 42pt, Bold(700), text-inverse, 좌측 정렬
    - x=0.6, y=0.9, w=8.8, h=0.7
  - 목차 카드 5개 (세로 나열):
    - 카드 영역: x=0.6, y=1.9, w=8.8, h=0.55 (각 카드, 간격 0.12)
    - 카드 배경: card-dark #2A2A2A, rectRadius=0.22
    - 넘버링: "01" ~ "05" — 20pt, Bold(700), accent-brand
    - 텍스트: 각 목차 항목 — 18pt, Regular(400), text-inverse
    - 항목:
      1. "왜 손저림인가? — 역학과 현황"
      2. "원인 질환별 병태생리와 해부학"
      3. "감별 진단 — 이학적 검사와 감별 포인트"
      4. "한의학 치료 근거 — 침, 한약, 추나"
      5. "임상 적용 가이드 — 변증, 알고리즘, Red Flags"
- **특수 요소**: 없음

---

## 슬라이드 3

- **타입**: TYPE-C
- **배경**: Light (bg-light #F5F0EB)
- **뱃지**: "KEY QUESTION" — pill shape, 상단 중앙, badgeBg #374151 + text-inverse
- **레이아웃**:
  - 인용 텍스트: "\"선생님, 손이 저려요\"" — 48pt, Bold(700), text-primary, 중앙 정렬
    - x=0.6, y=1.0, w=8.8, h=0.9
  - 서브 설명: "— 외래에서 가장 많이 듣는 말" — 20pt, Regular(400), text-muted, 중앙 정렬
    - x=0.6, y=2.0, w=8.8, h=0.5
  - 정의 카드 3개 (가로 배열):
    - 카드 영역: x=0.6, y=3.0, w=2.73, h=1.8 (간격 0.2)
    - 카드 배경: card-light #FFFFFF, rectRadius=0.22
    - 카드 제목: 각 유형명 — 20pt, Bold(700), accent-brand
    - 카드 본문: 설명 — 14pt, Regular(400), text-primary
    - 항목:
      1. "지각이상(Paresthesia)" — "따끔거림, 쥐어짜는 느낌"
      2. "감각저하(Hypoesthesia)" — "감각이 둔해지는 느낌"
      3. "이감각증(Dysesthesia)" — "비정상적 불쾌 감각"
  - 하단 요약: "원인: 말초신경 포착 ~ 중추성 병변까지 넓은 스펙트럼" — 16pt, Regular(400), text-muted
    - x=0.6, y=5.0, w=8.8, h=0.4
- **특수 요소**: 없음

---

## 슬라이드 4

- **타입**: TYPE-G
- **배경**: Dark (bg-dark #1A1A1A)
- **뱃지**: "PERFORMANCE" — pill shape, 상단 중앙, badgeBgLight + text-primary
- **레이아웃**:
  - 섹션 헤딩: "수근관증후군(CTS) & 경추 신경근병증 역학" — 36pt, Bold(700), text-inverse
    - x=0.6, y=0.9, w=8.8, h=0.7
  - 좌측 카드 — CTS 통계:
    - 카드 영역: x=0.6, y=1.9, w=4.2, h=3.2
    - 카드 배경: card-dark #2A2A2A, rectRadius=0.22
    - 대형 숫자: "3.8%" — 80pt, Black(900), accent-danger
    - 라벨: "일반 성인 유병률" — 16pt, Regular(400), text-muted
    - 하위 불릿 4개 — 16pt, Regular(400), text-inverse:
      - "여성이 남성의 3배"
      - "HIRA 2022: 약 20만 명+"
      - "직업 위험군: 10~15%"
      - "임신 중: 7~43%"
  - 우측 카드 — 경추 신경근병증 통계:
    - 카드 영역: x=5.2, y=1.9, w=4.2, h=3.2
    - 카드 배경: card-dark #2A2A2A, rectRadius=0.22
    - 대형 숫자: "83.2" — 80pt, Black(900), accent-success
    - 라벨: "인구 10만 명당 연간 발생" — 16pt, Regular(400), text-muted
    - 하위 불릿 1개 — 16pt, Regular(400), text-inverse:
      - "C6·C7 신경근이 70% 이상"
- **특수 요소**: 없음

---

## 슬라이드 5

- **타입**: TYPE-D
- **배경**: Light (bg-light #F5F0EB)
- **뱃지**: "PROBLEM" — pill shape, 상단 중앙, badgeBg #374151 + text-inverse
- **레이아웃**:
  - 섹션 헤딩: "그 밖의 주요 원인 질환" — 40pt, Bold(700), text-primary
    - x=0.6, y=0.9, w=8.8, h=0.7
  - 카드 3개 (가로 배열):
    - 카드 영역: x=0.6, y=1.9, w=2.73, h=3.2 (간격 0.2)
    - 카드 배경: card-light #FFFFFF, rectRadius=0.22
    - 카드 1 — TOS:
      - 카드 타이틀: "흉곽출구증후군(TOS)" — 20pt, Bold(700), accent-brand
      - 유병률 숫자: "0.3~2%" — 32pt, Black(900), accent-danger
      - 불릿: "신경성 TOS 95%+", "20~40대 여성·사무직" — 14pt, Regular(400), text-primary
    - 카드 2 — DPN:
      - 카드 타이틀: "당뇨병성 말초신경병증(DPN)" — 20pt, Bold(700), accent-brand
      - 유병률 숫자: "~50%" — 32pt, Black(900), accent-danger
      - 불릿: "당뇨 환자의 약 절반", "장갑·양말형 분포" — 14pt, Regular(400), text-primary
    - 카드 3 — 레이노:
      - 카드 타이틀: "레이노 현상(Raynaud)" — 20pt, Bold(700), accent-brand
      - 유병률 숫자: "20% / 11%" — 32pt, Black(900), accent-danger
      - 불릿: "여성 / 남성", "원발성이 이차성의 4배" — 14pt, Regular(400), text-primary
- **특수 요소**: 없음

---

## 슬라이드 6

- **타입**: TYPE-H
- **배경**: Light/White (#FFFFFF)
- **뱃지**: 없음 (TYPE-H 예외)
- **레이아웃**:
  - 섹션 헤딩: "수근관(Carpal Tunnel) 해부학" — 36pt, Bold(700), text-primary
    - x=0.6, y=0.5, w=4.5, h=0.6
  - 좌측 텍스트 영역:
    - x=0.6, y=1.3, w=4.5, h=3.8
    - 불릿 텍스트 — 16pt, Regular(400), text-primary:
      - "수근골 + 횡수근인대로 형성된 폐쇄 공간"
      - "내용물: 정중신경 + 굴곡건 9개"
      - "정상 내압: 2.5 mmHg"
      - "CTS 안정 시: 32 mmHg"
      - "손목 굴곡 시: 90 mmHg 이상"
      - "압박 → 탈수초 → 1~3번째, 4번째 내측 저림"
  - **시각 요소** (우측 — 수근관 단면 다이어그램):
    - 영역: x=5.5, y=1.0, w=4.0, h=3.8 (인치)
    - 유형: 해부 다이어그램 (단면도)
    - 내용: 수근관 단면. 반원형 수근골(하부 아치)과 횡수근인대(상부 수평 바)로 둘러싸인 폐쇄 공간. 내부에 원형으로 정중신경(1개, 크게 강조) + 굴곡건(9개, 작은 원 배열). 정중신경은 accent-danger로 강조, 굴곡건은 text-muted 컬러, 뼈/인대는 accentBrand 컬러
    - 컬러: accentDanger(정중신경), textMuted(굴곡건), accentBrand(뼈·인대), bgLight(배경)
  - 캡션: "수근관 단면 모식도 (도형 일러스트)" — 12pt, Regular(400), text-muted
    - x=5.5, y=4.9, w=4.0, h=0.3
- **특수 요소**: 없음

---

## 슬라이드 7

- **타입**: TYPE-E
- **배경**: Dark (bg-dark #1A1A1A)
- **뱃지**: "DEFINITION" — pill shape, 상단 중앙, badgeBgLight + text-primary
- **레이아웃**:
  - 섹션 헤딩: "경추 신경근별 피부분절(Dermatome)" — 36pt, Bold(700), text-inverse
    - x=0.6, y=0.9, w=8.8, h=0.7
  - **시각 요소** (좌측 — 상지 피부분절 다이어그램):
    - 영역: x=0.6, y=1.8, w=4.2, h=3.3 (인치)
    - 유형: 해부 다이어그램 (상지 피부분절)
    - 내용: 팔 윤곽 도형(직사각형/타원 조합). 5개 영역을 색 구분하여 C5(상완 외측), C6(엄지·검지, 전완 외측), C7(중지, 전완 배측), C8(약지·소지, 전완 내측), T1(상완 내측) 표시. 각 영역에 라벨 텍스트 배치
    - 컬러: accentBrand(C5), accentSuccess(C6), accentDanger(C7), badgeBg(C8), textMuted(T1)
  - 우측 텍스트 영역:
    - x=5.2, y=1.8, w=4.2, h=3.3
    - 신경근 목록 — 18pt, Regular(400), text-inverse:
      - "C5 → 상완 외측"
      - "C6 → 엄지·검지, 전완 외측"
      - "C7 → 중지, 전완 배측"
      - "C8 → 약지·소지, 전완 내측"
      - "T1 → 상완 내측"
    - 기전 설명: "추간판 탈출(HNP), 골극 → 신경근 압박" — 16pt, Regular(400), text-muted
    - 추가: "염증 사이토카인(IL-1β, TNF-α) → 과흥분" — 16pt, Regular(400), text-muted
- **특수 요소**: 없음

---

## 슬라이드 8

- **타입**: TYPE-D
- **배경**: Light (bg-light #F5F0EB)
- **뱃지**: "PROBLEM" — pill shape, 상단 중앙, badgeBg #374151 + text-inverse
- **레이아웃**:
  - 섹션 헤딩: "TOS · 레이노 · DPN 병태생리" — 36pt, Bold(700), text-primary
    - x=0.6, y=0.9, w=8.8, h=0.7
  - 카드 3개 (가로 배열):
    - 카드 영역: x=0.6, y=1.8, w=2.73, h=3.3 (간격 0.2)
    - 카드 배경: card-light #FFFFFF, rectRadius=0.22
    - 카드 1 — TOS:
      - 카드 타이틀: "흉곽출구증후군" — 20pt, Bold(700), accent-brand
      - 불릿 — 14pt, Regular(400), text-primary:
        - "사각근 삼각: 전·중사각근 + 첫 번째 늑골"
        - "완신경총 + 쇄골하동맥 통과"
        - "하완신경총(C8·T1) 압박"
        - "→ 내측 전완·소지 저림"
    - 카드 2 — 레이노:
      - 카드 타이틀: "레이노 현상" — 20pt, Bold(700), accent-brand
      - 불릿 — 14pt, Regular(400), text-primary:
        - "한랭·스트레스 → 교감신경 과활성"
        - "지단 혈관 과수축"
        - "3단계 색 변화:"
        - "창백 → 청색 → 충혈"
    - 카드 3 — DPN:
      - 카드 타이틀: "당뇨병성 말초신경병증" — 20pt, Bold(700), accent-brand
      - 불릿 — 14pt, Regular(400), text-primary:
        - "고혈당 → 소르비톨 축적"
        - "산화 스트레스 → 신경 혈류 장애"
        - "장갑·양말형 분포"
- **특수 요소**: 없음

---

## 슬라이드 9

- **타입**: TYPE-E
- **배경**: Dark (bg-dark #1A1A1A)
- **뱃지**: "PRACTICE" — pill shape, 상단 중앙, badgeBgLight + text-primary
- **레이아웃**:
  - 섹션 헤딩: "병력 청취 6대 핵심 항목" — 40pt, Bold(700), text-inverse
    - x=0.6, y=0.9, w=8.8, h=0.7
  - 카드 6개 (2행 x 3열 배열):
    - 1행: x=0.6, y=1.8, w=2.73, h=1.5 (간격 0.2)
    - 2행: x=0.6, y=3.5, w=2.73, h=1.5 (간격 0.2)
    - 카드 배경: card-dark #2A2A2A, rectRadius=0.22
    - 각 카드:
      - 넘버 뱃지: "01"~"06" — 14pt, Bold(700), accent-brand
      - 카드 타이틀 — 18pt, Bold(700), text-inverse
      - 카드 설명 — 13pt, Regular(400), text-muted
    - 항목:
      1. "저림 분포" — "손 전체 vs 특정 손가락"
      2. "야간 악화" — "CTS의 특징적 소견"
      3. "경부 통증 동반" — "경추 병변 시사"
      4. "양측성 여부" — "중추성, DPN, TOS 가능성"
      5. "색 변화" — "레이노 현상 시사"
      6. "직업·취미" — "반복 손목 사용, 진동 노출"
- **특수 요소**: 없음

---

## 슬라이드 10

- **타입**: TYPE-F
- **배경**: Light (bg-light #F5F0EB)
- **뱃지**: "PRACTICE" — pill shape, 상단 중앙, badgeBg #374151 + text-inverse
- **레이아웃**:
  - 섹션 헤딩: "이학적 검사 3종" — 40pt, Bold(700), text-primary
    - x=0.6, y=0.9, w=8.8, h=0.7
  - 카드 3개 (가로 배열):
    - 카드 영역: x=0.6, y=1.8, w=2.73, h=3.3 (간격 0.2)
    - 카드 배경: card-light #FFFFFF, rectRadius=0.22
    - 카드 1 — Phalen Test:
      - 카드 타이틀: "Phalen Test" — 22pt, Bold(700), accent-brand
      - 적응증: "CTS" — 14pt, Bold(700), accent-danger
      - 불릿 — 14pt, Regular(400), text-primary:
        - "손목 최대 굴곡 90°, 1분 유지"
        - "양성: 정중신경 영역 저림"
        - "민감도 68~80%"
        - "특이도 59~86%"
    - 카드 2 — Tinel Sign:
      - 카드 타이틀: "Tinel Sign" — 22pt, Bold(700), accent-brand
      - 적응증: "CTS" — 14pt, Bold(700), accent-danger
      - 불릿 — 14pt, Regular(400), text-primary:
        - "수근관 위 정중신경 타진"
        - "양성: 전기 자극감 방사"
        - "민감도 50~72%"
        - "특이도 55~87%"
    - 카드 3 — Spurling Test:
      - 카드 타이틀: "Spurling Test" — 22pt, Bold(700), accent-brand
      - 적응증: "경추 신경근병증" — 14pt, Bold(700), accent-success
      - 불릿 — 14pt, Regular(400), text-primary:
        - "환측 측방 굴곡+신전+축성 압박"
        - "양성: 동측 상지 방사통"
        - "민감도 40~60%"
        - "특이도 92~100%"
- **특수 요소**: 없음

---

## 슬라이드 11

- **타입**: TYPE-G
- **배경**: Dark (bg-dark #1A1A1A)
- **뱃지**: "SUMMARY" — pill shape, 상단 중앙, badgeBgLight + text-primary
- **레이아웃**:
  - 섹션 헤딩: "감별 진단 요약표" — 40pt, Bold(700), text-inverse
    - x=0.6, y=0.9, w=8.8, h=0.7
  - 표(Table):
    - 영역: x=0.6, y=1.8, w=8.8, h=3.3
    - 헤더 행: "질환 | 특징적 소견 | 핵심 검사" — 16pt, Bold(700), text-inverse, 배경 accentBrand
    - 데이터 행 (6행) — 14pt, Regular(400), text-inverse, 배경 교차 card-dark / bg-dark:
      1. "CTS | 야간 저림, Phalen(+), 엄지~4번째 | NCS/EMG"
      2. "경추 신경근병증 | 경부통, Spurling(+), 분절별 | MRI 경추"
      3. "TOS | 팔 거상 시 악화, Adson(+) | 영상, 혈관초음파"
      4. "DPN | 양측 대칭, 당뇨 병력, 장갑형 | NCS + HbA1c"
      5. "레이노 현상 | 한랭 유발, 3상 색 변화 | 냉각 유발 검사"
      6. "뇌졸중 | 편측 전체 저림+신경학적 징후 | 뇌 MRI/CT"
- **특수 요소**: 없음

---

## 슬라이드 12

- **타입**: TYPE-E
- **배경**: Dark (bg-dark #1A1A1A)
- **뱃지**: "PERFORMANCE" — pill shape, 상단 중앙, badgeBgLight + text-primary
- **레이아웃**:
  - 섹션 헤딩: "침치료 근거: CTS 메타분석" — 40pt, Bold(700), text-inverse
    - x=0.6, y=0.9, w=8.8, h=0.7
  - 좌측 — 핵심 수치 카드:
    - 카드 영역: x=0.6, y=1.8, w=4.2, h=3.3
    - 카드 배경: card-dark #2A2A2A, rectRadius=0.22
    - 연구명: "Yao et al. (2020)" — 18pt, Bold(700), accent-success
    - 서브: "체계적 고찰 & 메타분석" — 14pt, Regular(400), text-muted
    - 대형 숫자: "12 RCT" — 48pt, Black(900), text-inverse
    - 라벨: "622명 분석" — 16pt, Regular(400), text-muted
    - 핵심 결과: "VAS MD -1.36" — 32pt, Bold(700), accent-danger
    - CI 표기: "95% CI: -2.12 to -0.61" — 14pt, Regular(400), text-muted
  - 우측 — 근거 수준 비교:
    - 카드 영역: x=5.2, y=1.8, w=4.2, h=3.3
    - 카드 배경: card-dark #2A2A2A, rectRadius=0.22
    - 카드 타이틀: "근거 수준 비교" — 20pt, Bold(700), text-inverse
    - 비교 항목 3개 (수평 바 형태):
      - "침치료: Level I~II" — accent-success
      - "부목(Splint): Level I" — text-muted
      - "스테로이드 주사: Level I" — text-muted
    - 하단 텍스트: "전기침(EA) 병합 시 효과 우수" — 14pt, Regular(400), accent-success
- **특수 요소**: 없음

---

## 슬라이드 13

- **타입**: TYPE-F
- **배경**: Light (bg-light #F5F0EB)
- **뱃지**: "DEFINITION" — pill shape, 상단 중앙, badgeBg #374151 + text-inverse
- **레이아웃**:
  - 섹션 헤딩: "침치료 기전: 중추 가소성과 혈위" — 36pt, Bold(700), text-primary
    - x=0.6, y=0.9, w=8.8, h=0.7
  - 좌측 — fMRI 연구 카드:
    - 카드 영역: x=0.6, y=1.8, w=4.2, h=1.5
    - 카드 배경: card-light #FFFFFF, rectRadius=0.22
    - 연구 1: "Napadow et al. (2007, Pain)" — 16pt, Bold(700), accent-brand
    - 설명: "침치료 후 체성감각피질 재조직화 확인" — 14pt, Regular(400), text-primary
    - 연구 2: "Herman et al. (2018, Brain)" — 16pt, Bold(700), accent-brand
    - 설명: "로컬 침·원위 침 모두 피질 재지도화" — 14pt, Regular(400), text-primary
  - **시각 요소** (우측 상단 — 기전 플로우):
    - 영역: x=5.2, y=1.8, w=4.2, h=1.5 (인치)
    - 유형: 플로우차트 (가로 3단계)
    - 내용: "침 자극" → "말초 감각 입력 변화" → "중추 가소성(피질 재지도화)"
    - 각 단계를 둥근 직사각형 도형으로, 화살표로 연결
    - 컬러: accentBrand(도형 배경), cardLight(도형 fill), textPrimary(텍스트), accentSuccess(화살표)
  - 하단 — 혈위 카드 4개 (가로 배열):
    - 카드 영역: x=0.6, y=3.6, w=2.1, h=1.6 (간격 0.13)
    - 카드 배경: card-light #FFFFFF, rectRadius=0.22
    - 혈위명 — 18pt, Bold(700), accent-brand
    - 설명 — 12pt, Regular(400), text-primary
    - 항목:
      1. "大陵(PC7)" — "수근관 직상부, 정중신경 국소 자극"
      2. "內關(PC6)" — "전완 굴근 사이, 통증 조절"
      3. "合谷(LI4)" — "대측 감각 피질 활성화, Gate Control"
      4. "八邪(EX-UE9)" — "손가락 사이, 말초 혈류 개선"
- **특수 요소**: 없음

---

## 슬라이드 14

- **타입**: TYPE-E
- **배경**: Dark (bg-dark #1A1A1A)
- **뱃지**: "SOLUTION" — pill shape, 상단 중앙, badgeBgLight + text-primary
- **레이아웃**:
  - 섹션 헤딩: "경추 신경근병증: 침 + 추나 근거" — 36pt, Bold(700), text-inverse
    - x=0.6, y=0.9, w=8.8, h=0.7
  - 좌측 카드 — 침치료:
    - 카드 영역: x=0.6, y=1.8, w=4.2, h=3.3
    - 카드 배경: card-dark #2A2A2A, rectRadius=0.22
    - 카드 타이틀: "침치료" — 24pt, Bold(700), accent-success
    - 연구: "Liu et al. (2019, Spine)" — 14pt, Regular(400), text-muted
    - 불릿 — 16pt, Regular(400), text-inverse:
      - "14개 RCT 분석"
      - "침 + 추나/물리치료 → VAS 유의 개선"
      - "단독 침치료: 4~8주 유효"
    - 주요 혈위 라벨: "頸夾脊 · 風池 · 後溪 · 懸鍾" — 16pt, Bold(700), accent-brand
  - 우측 카드 — 추나:
    - 카드 영역: x=5.2, y=1.8, w=4.2, h=3.3
    - 카드 배경: card-dark #2A2A2A, rectRadius=0.22
    - 카드 타이틀: "추나(Chuna)" — 24pt, Bold(700), accent-success
    - 불릿 — 16pt, Regular(400), text-inverse:
      - "Cochrane Review (Gross 2015):"
      - "도수치료 > 단독 운동요법"
      - "(단기 통증 개선)"
    - 하단 강조: "대한추나의학회 (2021)" — 16pt, Bold(700), accent-brand
    - 권고: "급성기 이후 B등급 권고" — 16pt, Regular(400), text-inverse
- **특수 요소**: 없음

---

## 슬라이드 15

- **타입**: TYPE-F
- **배경**: Light (bg-light #F5F0EB)
- **뱃지**: "SOLUTION" — pill shape, 상단 중앙, badgeBg #374151 + text-inverse
- **레이아웃**:
  - 섹션 헤딩: "한약 치료 근거" — 40pt, Bold(700), text-primary
    - x=0.6, y=0.9, w=8.8, h=0.7
  - 카드 3개 (가로 배열):
    - 카드 영역: x=0.6, y=1.8, w=2.73, h=3.3 (간격 0.2)
    - 카드 배경: card-light #FFFFFF, rectRadius=0.22
    - 카드 1 — 당귀사역탕:
      - 카드 타이틀: "당귀사역탕" — 20pt, Bold(700), accent-brand
      - 한자: "當歸四逆湯" — 14pt, Regular(400), text-muted
      - 적응증 뱃지: "한랭성 저림" — 12pt, Bold(700), accent-danger, pill 배경
      - 불릿 — 13pt, Regular(400), text-primary:
        - "당귀, 계지, 작약, 세신 등"
        - "말초 혈류 개선"
        - "혈관 수축 억제"
      - 출처: "Kim et al. 2018" — 12pt, Regular(400), text-muted
    - 카드 2 — 황기계지오물탕:
      - 카드 타이틀: "황기계지오물탕" — 20pt, Bold(700), accent-brand
      - 한자: "黃芪桂枝五物湯" — 14pt, Regular(400), text-muted
      - 적응증 뱃지: "혈비(血痹)" — 12pt, Bold(700), accent-success, pill 배경
      - 불릿 — 13pt, Regular(400), text-primary:
        - "황기, 계지, 작약, 생강 등"
        - "말초신경 재생 촉진"
        - "NGF 상향 조절"
      - 출처: "Liu et al. 2020" — 12pt, Regular(400), text-muted
    - 카드 3 — 우차신기환:
      - 카드 타이틀: "우차신기환" — 20pt, Bold(700), accent-brand
      - 한자: "牛車腎氣丸" — 14pt, Regular(400), text-muted
      - 적응증 뱃지: "DPN" — 12pt, Bold(700), accent-danger, pill 배경
      - 불릿 — 13pt, Regular(400), text-primary:
        - "일본 가이드라인 1차 선택"
        - "NGF 증가"
        - "산화 스트레스 경감"
      - 출처: "Watanabe et al. 2014" — 12pt, Regular(400), text-muted
- **특수 요소**: 없음

---

## 슬라이드 16

- **타입**: TYPE-G
- **배경**: Dark (bg-dark #1A1A1A)
- **뱃지**: "PRACTICE" — pill shape, 상단 중앙, badgeBgLight + text-primary
- **레이아웃**:
  - 섹션 헤딩: "변증별 치료 가이드" — 40pt, Bold(700), text-inverse
    - x=0.6, y=0.9, w=8.8, h=0.7
  - 표(Table):
    - 영역: x=0.6, y=1.8, w=8.8, h=3.3
    - 헤더 행: "변증(辨證) | 증상 특징 | 치법 | 대표 처방" — 14pt, Bold(700), text-inverse, 배경 accentBrand
    - 데이터 행 (5행) — 13pt, Regular(400), text-inverse, 배경 교차 card-dark / bg-dark:
      1. "氣虛血滯 | 무기력, 창백, 둔한 저림 | 益氣活血 | 황기계지오물탕"
      2. "寒凝血脈 | 냉각 시 악화, 색 변화 | 溫陽散寒 | 당귀사역탕"
      3. "瘀血阻絡 | 고정 통증, 야간 악화 | 活血化瘀 | 혈부축어탕"
      4. "痰濕阻絡 | 무겁고 부은 느낌 | 化痰除濕 | 이진탕 합 도담탕"
      5. "肝腎虧虛 | 고령, 근위축, 재발 | 補益肝腎 | 독활기생탕"
- **특수 요소**: 없음

---

## 슬라이드 17

- **타입**: TYPE-H
- **배경**: Light/White (#FFFFFF)
- **뱃지**: 없음 (TYPE-H 예외)
- **레이아웃**:
  - 섹션 헤딩: "치료 알고리즘 & Red Flags" — 36pt, Bold(700), text-primary
    - x=0.6, y=0.5, w=8.8, h=0.6
  - **시각 요소** (좌측 — 치료 알고리즘 플로우차트):
    - 영역: x=0.6, y=1.3, w=5.0, h=3.8 (인치)
    - 유형: 플로우차트 (세로 흐름, 분기 포함)
    - 내용:
      1. "문진 · 이학적 검사" (직사각형, accentBrand 테두리)
      2. "영상 판독 (필요 시)" (직사각형, 점선 테두리)
      3. "변증 결정" (직사각형, accentBrand fill)
      4. 분기 3갈래:
         - "경증 CTS → 침(PC7·PC6·LI4) + 한약 4~6주"
         - "경추 신경근병증 → 夾脊침 + 추나 6~8주"
         - "한랭성 → 당귀사역탕 + 온침"
      5. "4~6주 재평가" (직사각형)
      6. "개선 미흡 → 양방 협진" (직사각형, 점선 테두리)
    - 컬러: accentBrand(단계 도형 테두리), accentSuccess(분기 화살표), cardLight(도형 fill), textPrimary(텍스트)
  - 우측 — Red Flags 카드:
    - 카드 영역: x=6.0, y=1.3, w=3.4, h=3.8
    - 카드 배경: accentDanger #E53E3E
    - 카드 타이틀: "Red Flags — 즉시 의뢰" — 20pt, Bold(700), text-inverse
    - 불릿 — 14pt, Regular(400), text-inverse:
      - "편측 상·하지 동시 저림 → 뇌졸중"
      - "양측 저림 + 보행 장애 → 척수병증"
      - "진행성 근위축 → ALS"
      - "발열 + 경부 강직 → 감염"
- **특수 요소**: 없음

---

## 슬라이드 18

- **타입**: TYPE-F
- **배경**: Light (bg-light #F5F0EB)
- **뱃지**: "PRACTICE" — pill shape, 상단 중앙, badgeBg #374151 + text-inverse
- **레이아웃**:
  - 섹션 헤딩: "예방 및 생활 지도" — 40pt, Bold(700), text-primary
    - x=0.6, y=0.9, w=8.8, h=0.7
  - 카드 4개 (2행 x 2열 배열):
    - 1행: x=0.6, y=1.8, w=4.2, h=1.5 / x=5.2, y=1.8, w=4.2, h=1.5
    - 2행: x=0.6, y=3.5, w=4.2, h=1.5 / x=5.2, y=3.5, w=4.2, h=1.5
    - 카드 배경: card-light #FFFFFF, rectRadius=0.22
    - 카드 1 — CTS 예방:
      - 카드 타이틀: "CTS 예방" — 20pt, Bold(700), accent-brand
      - 불릿 — 14pt, Regular(400), text-primary:
        - "손목 중립 자세 유지"
        - "야간 손목 부목 착용"
    - 카드 2 — 경추 예방:
      - 카드 타이틀: "경추 신경근병증 예방" — 20pt, Bold(700), accent-brand
      - 불릿 — 14pt, Regular(400), text-primary:
        - "경부 스트레칭·강화 운동"
        - "자세 교정 (모니터 높이)"
    - 카드 3 — 레이노:
      - 카드 타이틀: "레이노 현상" — 20pt, Bold(700), accent-brand
      - 불릿 — 14pt, Regular(400), text-primary:
        - "한랭 회피"
        - "방한 장갑 착용"
    - 카드 4 — DPN:
      - 카드 타이틀: "당뇨병성 말초신경병증" — 20pt, Bold(700), accent-brand
      - 불릿 — 14pt, Regular(400), text-primary:
        - "혈당 조절 (HbA1c 관리)"
- **특수 요소**: 없음

---

## 슬라이드 19

- **타입**: TYPE-B
- **배경**: Dark (bg-dark #1A1A1A)
- **뱃지**: "SUMMARY" — pill shape, 상단 중앙, badgeBgLight + text-primary
- **레이아웃**:
  - 섹션 헤딩: "핵심 요약 5가지" — 42pt, Bold(700), text-inverse
    - x=0.6, y=0.9, w=8.8, h=0.7
  - 카드 5개 (세로 나열):
    - 카드 영역: x=0.6, y=1.8, w=8.8, h=0.6 (각 카드, 간격 0.1)
    - 카드 배경: card-dark #2A2A2A, rectRadius=0.22
    - 넘버링: "01"~"05" — 20pt, Bold(700), accent-danger
    - 텍스트 — 16pt, Regular(400), text-inverse
    - 항목:
      1. "분류 — 말초신경 포착 / 경추 / TOS / 혈관성 / 대사성 / 중추성"
      2. "감별 — 병력 청취 6항목 + 이학적 검사 3종"
      3. "치료 근거 — 침(Level I~II), 한약(변증별), 추나(B등급)"
      4. "변증 5가지 — 기허혈체, 한응혈맥, 어혈조락, 담습조락, 간신휴허"
      5. "Red Flags — 편측 동시 저림, 보행 장애, 진행성 근위축"
- **특수 요소**: 없음

---

## 슬라이드 20

- **타입**: TYPE-A
- **배경**: Light (bg-light #F5F0EB)
- **뱃지**: 없음 (TYPE-A 예외)
- **레이아웃**:
  - 메인 타이틀: "Q&A" — 72pt, Black(900), text-primary, 중앙 정렬
    - x=0.6, y=1.2, w=8.8, h=1.0
  - 서브타이틀: "감사합니다" — 28pt, Bold(700), text-muted, 중앙 정렬
    - x=0.6, y=2.4, w=8.8, h=0.6
  - 참고문헌 영역:
    - x=0.6, y=3.3, w=8.8, h=2.0
    - 제목: "참고문헌" — 16pt, Bold(700), text-primary
    - 목록 — 10pt, Regular(400), text-muted (2열 배치):
      - 좌측 열 (1~7):
        1. "Atroshi et al. (1999) JAMA"
        2. "Szabo & Chidgey (1989) J Hand Surg Am"
        3. "Radhakrishnan et al. (1994) Brain"
        4. "Boulton et al. (2005) Lancet"
        5. "Yao et al. (2020) EBCAM"
        6. "Napadow et al. (2007) Pain"
        7. "Herman et al. (2018) Brain"
      - 우측 열 (8~13):
        8. "Liu et al. (2019) Spine"
        9. "Kim et al. (2018) J Ethnopharmacol"
        10. "Liu et al. (2020) Front Pharmacol"
        11. "Watanabe et al. (2014) J Diabetes Investig"
        12. "Gross et al. (2015) Cochrane Review"
        13. "대한추나의학회 (2021) 추나 임상진료지침"
- **특수 요소**: 없음

---

## 라이트/다크 교차 검증

| 슬라이드 | 타입 | 배경 | 연속 동일 배경 |
|---------|------|------|--------------|
| 1 | TYPE-A | Light | - |
| 2 | TYPE-B | Dark | - |
| 3 | TYPE-C | Light | L1 |
| 4 | TYPE-G | Dark | - |
| 5 | TYPE-D | Light | L1 |
| 6 | TYPE-H | Light/White | L2 |
| 7 | TYPE-E | Dark | - |
| 8 | TYPE-D | Light | L1 |
| 9 | TYPE-E | Dark | - |
| 10 | TYPE-F | Light | L1 |
| 11 | TYPE-G | Dark | - |
| 12 | TYPE-E | Dark | D2 |
| 13 | TYPE-F | Light | - |
| 14 | TYPE-E | Dark | - |
| 15 | TYPE-F | Light | L1 |
| 16 | TYPE-G | Dark | - |
| 17 | TYPE-H | Light/White | L1 |
| 18 | TYPE-F | Light | L2 |
| 19 | TYPE-B | Dark | - |
| 20 | TYPE-A | Light | - |

밝은 배경 3장 연속 없음 -- 교차 원칙 준수 확인 완료
