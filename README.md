# AI-native PowerPoint Creator 
written by: Gukhwan Hyun

---
**Key points**: System Design · Agent Integration · Model Context Protocol

---

### Scope

Google NotebookLM은 텍스트를 기반으로 Microsoft PowerPoint를 만들어주는 기능을 제공합니다.

본 프로젝트는 상용서비스에서 ideation을 얻고 product까지의 하나의 소프트웨어 라이프 사이클을 구현해보는 것을 목적으로 합니다.


### TL;DR

HTML→PPTX 직접 파싱의 구조적 한계(flexbox·grid 등 동적 레이아웃 비호환성)를 분석하고, Headless Chromium 스크린샷 기반 변환 파이프라인을 설계하여 LLM 생성 슬라이드의 레이아웃을 손실 없이 재현

Slide Canvas 이탈 여부를 검사하는 Overflow detection 알고리즘을 구현하고, LLM Feedback Loop 자동화 환경 구성

LLM 생성과 렌더링으로 구성된 2-Phase로 분리된 아키텍처를 설계하여 LLM과 Framwork agnostic한 라이브러리 구조 설계

라이브러리 코어 기능을 MCP Tool, Resources, Prompt를 활용하여 LLM과 End user에게 모두 친화적인 구조로 설계하고 **remote MCP 서버**로 구현

Sync 기반 Playwright을 asyncio.to_thread()로 래핑해 비동기 서버 환경에서의 Event Loop Block 문제를 줄이고, 외부 네트워크 요청 차단, 임시 파일 자동 정리, 환경변수 기반 시스템 제어를 적용하여 Production Ready 수준 운영 환경 구성


---

### Objective
사용자는 원하는 내용을 자연어로 설명하고, 시스템은 슬라이드를 한 장씩 생성하며, 사용자는 결과를 검토한 뒤 수정 요청을 할 수 있다. 

모든 슬라이드가 확정되면 최종 결과물을 `.pptx` 형식으로 내보내 PowerPoint, Keynote, Google Slides에서 바로 열 수 있어야 한다.


### System Design

HTML 슬라이드를 실제 브라우저 엔진으로 렌더링한 뒤 이미지를 캡처하고, 그 이미지를 PowerPoint 슬라이드 전체에 삽입하는 파이프라인을 설계하였습니다.

```text
HTML string
    │
    ▼
Headless Chromium (Playwright)
    │  1280 × 720 px로 렌더링
    │  overflow 검사 수행
    ▼
PNG screenshot
    │
    ▼
python-pptx
    │  빈 슬라이드에 전체 크기 이미지 삽입
    │  + 숨겨진 제목 메타데이터
    │  + 발표자 노트
    ▼
output.pptx
```

### Architecture:

시스템은 두 단계로 구성되어 있습니다.

**1단계 — Interactive (Application-controlled).** LLM이 슬라이드별 HTML을 생성하고, 사용자는 이를 채팅 UI에서 미리 본 뒤 수정 요청을 하며, 각 슬라이드 내용과 구조를 순차적으로 생성합니다.

**2단계 — Conversion(Library-controlled).** 모든 슬라이드가 확정되면 라이브러리가 HTML 문자열을 입력받아 렌더링하고 `.pptx`를 생성합니다. 

(Conversion 단계를 stateless하게 설계함으로써, 특정 애플리케이션 구조나 LLM 프레임워크에 종속되지 않고 폭넓게 통합할 수 있도록 했습니다.)

---

### Components

#### `Converter`

시스템 메모리 상에서 슬라이드 목록을 관리하며, 슬라이드를 추가, 삭제, 재정렬할 수 있는 method를 포함합니다.

| 메서드 | 목적 |
|---|---|
| `add_slide_from_string(html, title, notes)` | LLM이 생성한 HTML 문자열 추가 |
| `remove_slide(index)` | 사용자가 내보내기 전 슬라이드를 제거할 수 있도록 지원 |
| `reorder_slides(new_order)` | 사용자가 슬라이드 순서를 재정렬할 수 있도록 지원 |
| `check_slide(index)` | 슬라이드에 대한 사전 overflow 검사 |
| `save(output_path)` | 모든 슬라이드를 렌더링하고 `.pptx` 파일 생성 |

#### `Renderer`

각 슬라이드마다 렌더러는 HTML을 로드하고, 레이아웃이 로드 되도록 일정 시간동안 대기한 뒤, 1280 × 720 픽셀 영역을 캡처합니다.

- **외부 네트워크 체크**: LLM이 생성한 슬라이드에 Google Fonts 링크나 외부 이미지 URL이 포함되어있는지 확인한다.
- **overflow 검사**: 브라우저 렌더링이 완료된 직후 안전성 검사를 실행한다.

---

### General Workflow

End User의 관점에서 보면 전체 흐름은 다음과 같습니다.

```text
사용자: "우리 회사 3분기 실적에 대한 4장짜리 파워포인트를 만들어줘"
        │
        ▼
애플리케이션
  → 슬라이드 주제 + HTML 가이드라인을 포함해 LLM 프롬프트 구성
  → LLM이 1번 슬라이드용 HTML 생성
  → 채팅 UI에서 HTML 렌더링 결과 미리보기
        │
  사용자: "배경색을 진한 파란색으로 바꿔줘"
        │
  → LLM이 1번 슬라이드 HTML 재생성
  → 사용자: "좋아요" ✓  ← 1번 슬라이드 확정
        │
  2, 3, 4번 슬라이드에 대해 반복...
        │
  사용자: "PowerPoint로 저장해줘"
        │
        ▼
    본 라이브러리
        │
        ▼
사용자가 output.pptx 수신
```

---

### Limitations

- **Google NotebookLM과 동일하게 제공되는 슬라이드는 이미지형태여서 PowerPoint 내부의 수정하기가 어렵다는 한계점이 있습니다.** 

---

### Applications
- [Financial Researcher](https://github.com/serener91/OpenAgent/tree/main/financial_research)


