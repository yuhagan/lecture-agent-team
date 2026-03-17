# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 프로젝트 개요

한의학·의학 블로그용 PowerPoint 슬라이드를 Node.js(pptxgenjs)로 생성하는 프로젝트이다. 루트의 독립 실행형 스크립트들과 `lecture-agent-team/` 멀티에이전트 파이프라인으로 구성된다.

## 실행 명령어

```bash
# 의존성 설치
npm install

# PPTX 생성 스크립트 실행 (각 스크립트가 독립적으로 .pptx 파일을 출력)
node generate-trapezius.js          # 승모근_움직임해부학.pptx
node generate-pectoralis-major.js   # 대흉근_해부학.pptx
node generate-pectoralis-v2.js      # 대흉근_해부학_v2.pptx
node generate-silson-lecture.js     # 실손강의_메디.pptx

# 멀티에이전트 강의 생성 파이프라인
cd lecture-agent-team
node output/build.js                # 빌더가 생성한 스크립트 직접 실행
```

## 아키텍처

### 루트 레벨 — 독립 실행형 스크립트

각 `generate-*.js` 파일은 완결된 단일 스크립트로, `pptxgenjs`를 CommonJS(`require`)로 불러와 슬라이드를 직접 조립하고 `pptx.writeFile()`로 저장한다. 이미지는 로컬 경로(`./파일명.png`)로 참조한다.

### `lecture-agent-team/` — 멀티에이전트 파이프라인

Claude Code 오케스트레이터가 4개의 서브에이전트를 순차·병렬 조합하여 강의를 자동 제작한다:

```
researcher (Sonnet) → research_output.md
writer (Opus)       → lecture_script.md
                    ↙              ↘  (병렬)
designer (Opus)              builder (Sonnet)
→ design_spec.md             → output/build.js → lecture.pptx
```

각 에이전트 지시서는 `lecture-agent-team/agents/*.md`에 있다. 빌더는 `lecture-agent-team/skills/pptx/SKILL.md`의 디자인 토큰을 참조하여 `output/build.js`를 작성한 뒤 실행한다. 멀티에이전트 시스템 상세는 `lecture-agent-team/CLAUDE.md` 참조.

## 디자인 시스템 규칙

### 한의학 해부학/근육학 (anatomy scripts)

- **폰트**: `Pretendard`
- **색상**: `navy: "1B2A4A"`, `teal: "0A7E8C"`, `ltTeal: "E6F3F5"`
- **레이아웃**: 16:9 (`LAYOUT_16x9`, 10×5.63인치)
- **슬라이드 구조**: 표지 → 목차 → 개요+이미지 → 기시/정지 표 → 작용 → 신경지배 흐름도 → 임상적 의의 → 요약 → 참고문헌

### 실손보험/경영 강의 (silson scripts)

- **폰트**: `Pretendard`
- **색상**: `primary: "1B3A5C"`, `secondary: "0D9488"`, `accent: "3B82F6"`
- **하단 푸터**: 강의명(좌) + 페이지 번호(우)

### 공통 규칙

- 한국어 의학 용어: `한글(English)` 병기 (예: `족저근막염(Plantar Fasciitis)`)
- 이미지 캡션에 출처 및 라이선스 명시 (Public Domain, CC BY-SA 등)
- 슬라이드 패딩: 좌우 0.5~0.6인치
- 카드 모서리: `rectRadius: 0.22` (일반 카드), `rectRadius: 0.28` (뱃지/pill)

## 외부 리서치 → 강의 자동화 워크플로우

### 트리거 패턴

사용자가 제미나이 딥리서치 등 외부 보고서를 직접 제공하면 researcher를 건너뛰고 writer부터 실행한다:

```
리서치 줄게 강의 만들어줘: [주제]
[보고서 내용 붙여넣기 또는 파일 경로]
```

### 실행 순서

```
1. 사용자가 제공한 내용을 research_output.md로 저장
   └─ lecture-agent-team/output/research_output.md 에 그대로 저장

2. writer → designer + builder 파이프라인 실행 (researcher 생략)
   ├─ writer Agent     → lecture-agent-team/output/lecture_script.md
   └─ designer + builder Agent (병렬)
       ├─ designer → lecture-agent-team/output/design_spec.md
       └─ builder  → lecture-agent-team/output/build.js + output/lecture.pptx
```

### 주의 사항

- 사용자가 제공한 보고서 내용을 researcher.md 형식으로 변환하지 말고 그대로 저장
- 보고서에 없는 내용을 임의로 추가하지 않음 (writer가 있는 내용만 활용)

---

## 노션 → 강의 자동화 워크플로우

### 트리거 패턴

사용자가 아래 형태로 요청하면 이 워크플로우를 자동 실행한다:

```
노션 강의 만들어줘: [페이지 제목 또는 URL]
노션 강의 만들어줘: [페이지 제목] → [저장할 노션 페이지 URL]  (저장 위치 지정 시)
```

### 실행 순서

```
1. 노션 페이지 읽기
   └─ notion-search 또는 notion-fetch로 소스 페이지 내용 추출
   └─ 소스 페이지의 parent URL 기록 (저장 위치 기본값으로 사용)

2. lecture-agent-team 파이프라인 실행
   ├─ researcher Agent → lecture-agent-team/output/research_output.md
   ├─ writer Agent     → lecture-agent-team/output/lecture_script.md
   └─ designer + builder Agent (병렬)
       ├─ designer → lecture-agent-team/output/design_spec.md
       └─ builder  → lecture-agent-team/output/build.js + output/lecture.pptx

3. 노션 새 페이지 저장
   └─ notion-create-pages로 강의 스크립트(lecture_script.md)를 노션 페이지로 저장
   └─ 저장 위치: 사용자가 지정한 페이지 > 미지정 시 소스 페이지의 parent 페이지 아래
   └─ 페이지 제목: "[강의] [원본 페이지 제목]"

4. 결과 보고
   └─ 로컬 PPTX 경로: lecture-agent-team/output/lecture.pptx
   └─ 노션 저장 URL 제공
```

### 노션 페이지 저장 형식

notion-create-pages로 저장 시 아래 구조로 작성한다:

```
제목: [강의] [원본 페이지 제목]
내용:
  # 강의록: [강의 제목]
  > PPTX 파일 위치: lecture-agent-team/output/lecture.pptx (로컬)
  ---
  [lecture_script.md 전체 내용]
```

### 주의 사항

- PPTX 바이너리 파일은 노션 MCP로 업로드 불가 → 로컬 경로만 안내
- 소스 페이지를 특정할 수 없으면 notion-search로 검색 후 확인
- 저장 위치를 특정할 수 없으면 사용자에게 확인 후 진행
- 파이프라인 에이전트 호출 방식은 `lecture-agent-team/CLAUDE.md` 참조

## 의존성

- `pptxgenjs ^4.0.1` — PPTX 생성 라이브러리 (CommonJS `require("pptxgenjs")`)
- `@types/node` — Node.js 타입
- `image-size` — 이미지 크기 감지 (pptxgenjs 내부 의존성)
