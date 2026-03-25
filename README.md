# AI-native PowerPoint Creator 
written by: Gukhwan Hyun

---
**Key points**: System Design · Agent Integration · Model Context Protocol

---

## Scope

Google NotebookLM은 텍스트를 기반으로 Microsoft PowerPoint를 만들어주는 기능을 제공한다.

본 프로젝트는 상용서비스에서 ideation을 얻고 product까지의 하나의 소프트웨어 라이프사이클을 구현해보는 것을 목적으로 한다.


## TL;DR

HTML→PPTX 직접 파싱의 구조적 한계(flexbox·grid 등 동적 레이아웃 비호환성)를 분석하고, Headless Chromium 스크린샷 기반 변환 파이프라인을 설계하여 LLM 생성 슬라이드의 레이아웃을 손실 없이 재현

Slide Canvas 이탈 여부를 검사하는 Overflow detection 알고리즘을 구현하고, LLM Feedback Loop 자동화 기반 마련

LLM 생성과 렌더링으로 구성된 2-Phase로 분리된 아키텍처를 설계하여 LLM과 Framwork agnostic한 라이브러리 구조 설계

라이브러리 코어 기능을 MCP Tool, Resources, Prompt를 활용하여 LLM과 End user에게 모두 친화적인 구조로 설계하고 remote MCP 서버로 구현

동기 기반 Playwright을 asyncio.to_thread()로 래핑해 비동기 서버 환경에서의 Event Loop Block 문제를 줄이고, 외부 네트워크 요청 차단, 임시 파일 자동 정리, 환경변수 기반 시스템 제어를 적용하여 Production Ready 수준 운영 환경 구성


---

## Objective
사용자는 원하는 내용을 자연어로 설명하고, 시스템은 슬라이드를 한 장씩 생성하며, 사용자는 결과를 검토한 뒤 수정 요청을 할 수 있다. 

모든 슬라이드가 확정되면 최종 결과물을 `.pptx` 형식으로 내보내 PowerPoint, Keynote, Google Slides에서 바로 열 수 있어야 한다.


## System Design

HTML 슬라이드를 실제 브라우저 엔진으로 렌더링한 뒤 이미지를 캡처하고, 그 이미지를 PowerPoint 슬라이드 전체에 삽입하는 파이프라인을 설계했다. 

그 결과 최종 `.pptx`의 각 슬라이드는 브라우저에서 보이던 모습과 동일하게 유지되는데 이 방식은 Puppeteer 기반 PDF 생성기나 스크린샷 서비스가 채택하는 원리와 동일하다.

브라우저는 HTML을 가장 정확하게 해석할 수 있는 렌더러이므로, 그 강점을 그대로 활용하는 것이 가장 합리적이었다.

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


### Approach:

아래 세 가지 대안을 검토했다.

**옵션 A — HTML을 직접 PPTX로 파싱하는 방식.** 

일부 라이브러리는 HTML 요소를 PPTX 도형으로 매핑하려고 시도한다. 그러나 이런 방식은 제목, 문단 수준의 단순 콘텐츠에서는 동작하더라도, flexbox, grid, absolute positioning, gradient 등 현대적인 CSS 레이아웃이 포함되는 순간 안정성을 잃는다. LLM이 시각적 완성도를 위해 이러한 기능을 자연스럽게 사용할 것이라는 점을 고려하면, 이 옵션은 초기 단계에서 제외하는 것이 타당했다.

**옵션 B — HTML을 PDF로 변환한 뒤 PDF를 PPTX로 변환하는 방식.** 

이 방식은 불필요한 중간 단계를 추가한다. PDF 렌더링 자체가 별도의 문제를 만들고, PDF 페이지에서 이미지를 다시 추출하는 과정에서 구현 복잡도와 품질 저하 가능성도 커진다. PNG로 바로 가는 편이 더 단순하고 명확하다.

**옵션 C (채택) — HTML을 PNG로 렌더링하고, PNG를 PPTX에 삽입하는 방식.** 

직접적이고, 손실이 없으며, 구현과 운영 측면 모두에서 신뢰성이 높다. 유일한 트레이드오프는 슬라이드 내용이 PowerPoint 내부의 수정 가능한 텍스트가 아니라 이미지라는 점이다. 그러나 이 프로젝트의 사용 시나리오에서는 슬라이드 편집이 채팅 UI에서 이뤄지므로, 이 제약은 충분히 수용 가능하다고 판단했다.



### Architecture:

시스템은 명확한 경계를 가진 두 단계로 의도적으로 분리했다.

**1단계 — Interactive (Application-controlled).** LLM이 슬라이드별 HTML을 생성하고, 사용자는 이를 채팅 UI에서 미리 본 뒤 수정 요청을 하며, 각 슬라이드를 순차적으로 확정한다.

**2단계 — Conversion(Library-controlled).** 모든 슬라이드가 확정되면 라이브러리가 HTML 문자열을 입력받아 렌더링하고 `.pptx`를 생성한다. 이 단계는 1단계가 어떤 LLM, 어떤 UI, 어떤 프레임워크와도 얽매이지 않게한다.

Conversion 단계를 stateless하게 설계함으로써, 특정 애플리케이션 구조나 LLM 프레임워크에 종속되지 않고 폭넓게 통합할 수 있도록 했다.

---

## Components

### `Converter`

시스템 메모리 상에서 슬라이드 목록을 관리하며, 슬라이드를 추가, 삭제, 재정렬할 수 있는 method를 포함한다.

주요 메서드는 다음과 같다.

| 메서드 | 목적 |
|---|---|
| `add_slide_from_string(html, title, notes)` | LLM이 생성한 HTML 문자열 추가 |
| `remove_slide(index)` | 사용자가 내보내기 전 슬라이드를 제거할 수 있도록 지원 |
| `reorder_slides(new_order)` | 사용자가 슬라이드 순서를 재정렬할 수 있도록 지원 |
| `check_slide(index)` | 슬라이드에 대한 사전 overflow 검사 |
| `save(output_path)` | 모든 슬라이드를 렌더링하고 `.pptx` 파일 생성 |

### `Renderer`

각 슬라이드마다 렌더러는 HTML을 로드하고, 레이아웃이 로드 되도록 일정 시간동안 대기한 뒤, 1280 × 720 픽셀 영역을 캡처한다.

렌더러는 두 가지 중요한 동작을 수행한다.

- **외부 네트워크 체크**: LLM이 생성한 슬라이드에 Google Fonts 링크나 외부 이미지 URL이 포함되어있는지 확인한다.
- **overflow 검사**: 브라우저 렌더링이 완료된 직후 안전성 검사를 실행한다.

---

## General Workflow

최종 사용자의 관점에서 보면 전체 흐름은 다음과 같다.

```text
사용자: "우리 회사 3분기 실적에 대한 4장짜리 덱을 만들어줘"
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

## Evaluation


### Limitations

- **Google NotebookLM과 동일하게 제공되는 슬라이드는 이미지형태여서 PowerPoint 내부의 수정하기가 어렵다는 한계점이 있다.** 

### Improvements

- **슬라이드 템플릿 시스템.**: 각 슬라이드가 완전히 독립적인 HTML이다. 기본 테마를 정의하는 베이스 HTML과, 개별 슬라이드가 콘텐츠만 주입하는 구조를 도입하면 시각적 일관성을 높이고 LLM이 매번 생성해야 하는 HTML 양도 줄일 수 있다고 생각된다.

---



