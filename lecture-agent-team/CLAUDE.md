# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

---

## 시스템 개요

한의학·의학 강의를 자동으로 기획·제작하는 멀티에이전트 시스템이다.
Claude Code(오케스트레이터)가 Task tool로 4개의 서브에이전트를 순차·병렬 조합하여 실행한다.

**청중**: 한의사, 의료인 (임상 사례·근거 중심 선호)
**출력 도메인**: 한의학 근육학/해부학, 임상 질환, 실손보험, 의학 일반(양방 포함)

---

## 에이전트 구성 및 파이프라인

```
[오케스트레이터: Claude Code]
         │
         ▼
 1. researcher  →  research_output.md
         │
         ▼
 2. writer      →  lecture_script.md
         │
    ┌────┴────┐
    ▼         ▼
 3. designer  4. builder  →  (병렬 실행)
    │              │
    ▼              ▼
design_spec.md  lecture.pptx
```

### 에이전트별 역할

| 에이전트 | 모델 | 입력 | 출력 |
|---------|------|------|------|
| researcher | Sonnet | 강의 주제 | `research_output.md` |
| writer | Opus | `research_output.md` | `lecture_script.md` |
| designer | Opus | `lecture_script.md` | `design_spec.md` |
| builder | Sonnet + medical-pptx 스킬 | `lecture_script.md` + `design_spec.md` | `lecture.pptx` |

각 에이전트의 상세 지시서는 `agents/` 폴더에 위치한다.

---

## 파일 구조

```
lecture-agent-team/
├── CLAUDE.md                  # 이 파일 (오케스트레이터 지침)
├── agents/
│   ├── researcher.md          # 리서처 에이전트 역할 지시서
│   ├── writer.md              # 라이터 에이전트 역할 지시서
│   ├── designer.md            # 디자이너 에이전트 역할 지시서
│   └── builder.md             # 빌더 에이전트 역할 지시서
├── output/
│   ├── research_output.md     # 리서처 결과물
│   ├── lecture_script.md      # 라이터 결과물
│   ├── design_spec.md         # 디자이너 결과물
│   └── lecture.pptx           # 최종 산출물
└── templates/                 # 참조용 디자인 템플릿 (선택)
```

---

## 오케스트레이터 라우팅 규칙

사용자 요청에 따라 실행 범위를 결정한다:

| 요청 패턴 | 실행 에이전트 |
|----------|--------------|
| "강의 만들어줘 / 전체 파이프라인" | researcher → writer → designer + builder (병렬) |
| "자료만 조사해줘" | researcher |
| "강의록만 작성해줘" | writer (research_output.md 기존 파일 활용) |
| "PPT만 수정/다시 만들어줘" | designer + builder (lecture_script.md 기존 파일 활용) |
| "PPT 바로 만들어줘" | builder (design_spec.md 기존 파일 활용) |

---

## 에이전트 호출 방식

오케스트레이터(Claude Code)는 Task tool을 사용하여 서브에이전트를 호출한다.
각 에이전트를 호출할 때 해당 `agents/*.md` 파일의 내용을 프롬프트에 포함시킨다.

**순차 실행**: researcher → writer (앞 결과가 다음 입력이므로 반드시 완료 후 다음 호출)
**병렬 실행**: designer + builder는 동시에 Task tool 두 개를 한 메시지에서 호출

---

## 빌더 에이전트와 PPTX 스킬

빌더는 Claude Code의 `medical-pptx` 스킬을 활용하여 PPTX를 생성한다.
`agents/builder.md`에는 스킬 호출 방법과 디자인 명세 해석 규칙이 명시된다.

빌더가 생성하는 PPTX의 공통 규칙 (디자이너 명세와 충돌 시 명세 우선):
- 레이아웃: 16:9 (10 × 5.63인치)
- 한국어 의학 용어: `한글(English)` 병기
- 출처 슬라이드: 마지막 슬라이드에 참고문헌 목록

---

## 리서처 데이터 소스 우선순위

1. PubMed, Cochrane Library (영문 논문)
2. KISTI, OASIS, RISS (국내 한의학 논문)
3. 국가기관: 건강보험심사평가원, 질병관리청, 보건복지부
4. 대학병원 건강정보 (서울대, 연세, 삼성서울 등)
5. 실손보험: 금융감독원, 보험개발원

`research_output.md`는 반드시 출처(URL, 저자, 연도)를 포함해야 한다.

---

## 디자인 시스템 (도메인별)

### 한의학 해부학/근육학
- 폰트: `Noto Sans KR`
- 색상: navy `1B2A4A` + teal `0A7E8C` + white background
- 좌측 teal 세로 악센트 바 (모든 콘텐츠 슬라이드)

### 임상 질환 / 의학 일반
- 폰트: `Noto Sans KR`
- 색상: 도메인 특성에 따라 디자이너가 팔레트 결정

### 실손보험 / 경영
- 폰트: `Pretendard`
- 색상: navy `1B3A5C` + teal `0D9488` + blue accent `3B82F6`
- 하단 푸터: 강의명(좌) + 페이지 번호(우)

---

## 산출물 품질 기준

- **라이터**: 슬라이드 1개 = 핵심 메시지 1개, 슬라이드별 발표 멘트 + 전환 멘트 포함
- **디자이너**: 슬라이드별 레이아웃 타입 명시 (표지/목차/2분할/도표/이미지+텍스트/참고문헌)
- **빌더**: `design_spec.md` 명세를 100% 준수, 차트·표는 스크립트 데이터에서 자동 생성
