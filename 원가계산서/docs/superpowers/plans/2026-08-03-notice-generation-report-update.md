# 공고문 초안 생성 계획 HTML 업데이트 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 기존 원가계산서 검증 현황 HTML에 공고문 생성 사용자 흐름, 단계별 동작 상태, 입력항목, 구현 범위와 로드맵을 한눈에 볼 수 있는 독립 섹션으로 추가한다.

**Architecture:** 기존 단일 HTML의 디자인 토큰과 초록·주황·빨강 상태 체계를 재사용한다. 새 섹션은 상태 태그가 붙은 10단계 사용자 흐름, 단계별 상세 표, 입력 범주, 현재·후속 범위와 기술 마일스톤으로 구성하며 외부 자산이나 빌드 의존성을 추가하지 않는다.

**Tech Stack:** 정적 HTML5, 인라인 CSS, 기존 인라인 JavaScript, Node.js 정적 구조 검사, 인앱 Browser 시각 QA

## Global Constraints

- 수정 대상은 `docs/원가계산서-검증-현황과-개선안.html` 한 파일뿐이다.
- 기존 7단계 검증 흐름, 33개 기능 집계, 23개 개선 체크리스트와 체크박스 저장 동작을 보존한다.
- 공고문 생성은 `설계 완료 · 구현 예정`인 독립 후속 기능으로 표시한다.
- 사용자 흐름의 모든 단계에는 `동작`, `일부 동작`, `동작 안함` 중 하나의 태그를 표시한다.
- 현재 조건과 관계없는 입력항목은 숨기지 않고 회색 비활성화한다는 방침을 설명한다.
- 활성화된 입력항목은 모두 필수라는 방침을 설명한다.
- 1차는 공사입찰공고 HWPX 생성, 후속은 용역·구매 템플릿과 날짜·금액 검증으로 구분한다.
- 날짜 후속 규칙은 초일 불산입 7일, 7일째 공휴일이면 다음 영업일까지 연장으로 명시한다.
- 외부 폰트·CDN·네트워크 요청을 추가하지 않는다.

---

### Task 1: 공고문 생성 현황 섹션 구조와 내용 추가

**Files:**
- Modify: `docs/원가계산서-검증-현황과-개선안.html`

**Interfaces:**
- Consumes: 기존 CSS 변수 `--ok-*`, `--half-*`, `--no-*`, 기존 내비게이션과 `.wrap` 레이아웃
- Produces: `#notice-plan` 앵커, `.notice-flow`, `.notice-step`, `.status-tag`, `.notice-detail`, `.notice-input-grid`, `.notice-roadmap` 마크업과 스타일

- [ ] **Step 1: 수정 전 구조 검사를 실행해 새 섹션이 아직 없음을 확인**

```powershell
$html = Get-Content -Raw -Encoding UTF8 'docs/원가계산서-검증-현황과-개선안.html'
if ($html.Contains('id="notice-plan"')) { throw '수정 전인데 notice-plan 섹션이 이미 존재함' }
```

Expected: 성공 종료하며 `notice-plan` 섹션이 없음을 확인한다.

- [ ] **Step 2: 기존 CSS에 공고문 흐름·상태·상세표 스타일 추가**

다음 책임을 가진 CSS를 `@media print` 앞에 추가한다.

```css
.notice-hero { border:1px solid var(--line); border-radius:14px; padding:22px; background:var(--panel); }
.notice-flow { display:grid; grid-template-columns:repeat(5,minmax(150px,1fr)); gap:10px; }
.notice-step { position:relative; border:1px solid var(--line); border-radius:10px; padding:14px; background:#fff; }
.status-tag { display:inline-block; border-radius:999px; padding:2px 8px; font-size:12px; font-weight:600; }
.status-tag.ok { background:var(--ok-bg); color:var(--ok-hd); }
.status-tag.half { background:var(--half-bg); color:var(--half-hd); }
.status-tag.no { background:var(--no-bg); color:var(--no-hd); }
.notice-input-grid { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:10px; }
@media (max-width:760px) { .notice-flow,.notice-input-grid { grid-template-columns:1fr; } }
```

상세표의 `동작하는 것`과 `부족한 것` 열, 단계 번호, 비활성 입력 설명, 인쇄 시 페이지 넘침을 위한 스타일도 같은 접두사로 추가한다. 기존 전역 선택자와 이름 충돌을 만들지 않는다.

- [ ] **Step 3: 상단 요약과 목차에 후속 기능 진입점 추가**

`.stats`에는 아래 배지를 추가한다.

```html
<span class="pill half">후속 기능 1개 계획</span>
```

`nav`에는 아래 링크를 추가한다.

```html
<a href="#notice-plan">공고문 초안 생성</a>
```

- [ ] **Step 4: 상태 태그가 붙은 10단계 사용자 흐름 추가**

`#checklist` 섹션 뒤, `#criteria` 앞에 `<h2 id="notice-plan">공고문 초안 생성 · 사용자 흐름과 개발 현황</h2>`를 추가한다. 다이어그램 단계와 상태는 아래 순서를 그대로 사용한다.

```text
1. 공고문 생성 진입 — 동작
2. 원가계산서 입력 — 동작
3. 원가계산서 자동 채움 — 일부 동작
4. 유형별 정보 입력 — 일부 동작
5. 필수값 확인 — 일부 동작
6. 금액 확인 — 동작 안함
7. 날짜·공휴일 확인 — 동작 안함
8. 유형별 템플릿 선택 — 일부 동작
9. HWPX 생성·다운로드 — 일부 동작
10. 한글에서 최종 확인 — 동작 안함
```

각 `.notice-step`에는 단계 번호, 제목, 정확한 상태 태그와 한 줄 요약을 넣는다. 흐름 앞에는 `설계 완료 · 구현 예정` 배지와 `원가계산서 + 사용자 입력 → HWPX 공고문 초안` 목표를 표시한다.

- [ ] **Step 5: 각 단계의 동작·미동작 상세내역 추가**

다이어그램 아래 상세표는 설계 명세의 10개 판정 행을 그대로 반영한다. 특히 아래 사실을 생략하지 않는다.

```text
- 진입·기존 Excel 업로드는 동작한다.
- 자동 채움은 파일명 기반 사업명 후보와 합계 후보만 일부 동작한다.
- 캡처의 발주계획번호·발주년월·세부계약종류·협정·사급자재비·수의계약·품목·부서별 담당자 입력은 없다.
- 필수값 검사는 공사명·발주기관·추정가격 3개뿐이다.
- 금액 대조와 날짜·공휴일 규칙은 동작하지 않는다.
- 유형별 템플릿 분기는 없고 현재 템플릿은 14개 후보 중 입찰종류 1개만 치환한다.
- HWPX 엔진과 fill-test는 동작하지만 실제 입력 대부분이 문서에 반영되지 않는다.
- 한글에서의 최종 인수는 이루어지지 않았다.
```

- [ ] **Step 6: 입력항목과 현재·후속 범위 추가**

입력항목은 기본정보, 계약정보, 금액정보, 수의계약정보, 품목정보, 담당자정보, 공고정보의 7개 그룹으로 표시한다. 조건과 무관한 항목은 숨기지 않고 회색 비활성화하며, 활성 항목은 모두 필수라는 원칙을 별도 안내 상자로 강조한다.

현재 범위에는 공사 템플릿, 사용자 날짜·금액 직접 입력, HWPX 치환·다운로드를 넣는다. 후속 범위에는 용역·구매 템플릿, 원가계산서 금액 대조, 초일 불산입 7일, 공휴일 시 다음 영업일 연장, 공휴일 데이터 연동을 넣는다.

- [ ] **Step 7: 정적 구조 검사를 실행해 필수 내용과 기존 내용을 함께 확인**

```powershell
$html = Get-Content -Raw -Encoding UTF8 'docs/원가계산서-검증-현황과-개선안.html'
$required = @('id="notice-plan"','공고문 생성 진입','원가계산서 자동 채움','날짜·공휴일 확인','유형별 템플릿 선택','HWPX 생성·다운로드','동작 안함','초일 불산입','다음 영업일','23개')
$missing = $required | Where-Object { -not $html.Contains($_) }
if ($missing.Count) { throw "누락: $($missing -join ', ')" }
if (($html | Select-String -Pattern 'class="notice-step"' -AllMatches).Matches.Count -ne 10) { throw '사용자 흐름 단계가 10개가 아님' }
```

Expected: 누락 없이 성공하고 `.notice-step`이 정확히 10개다.

- [ ] **Step 8: 변경 내용을 별도 커밋**

```powershell
git add -- 'docs/원가계산서-검증-현황과-개선안.html'
git diff --cached --check
git commit -m "docs: add notice generation flow status"
```

### Task 2: 반응형·인쇄·상호작용 시각 검증

**Files:**
- Modify if QA finds issues: `docs/원가계산서-검증-현황과-개선안.html`

**Interfaces:**
- Consumes: Task 1의 `#notice-plan`과 `.notice-*` 마크업·스타일
- Produces: 데스크톱·모바일·인쇄에서 읽을 수 있고 목차 이동이 확인된 최종 HTML

- [ ] **Step 1: 로컬 정적 서버로 보고서를 제공**

```powershell
Start-Process -FilePath python -ArgumentList '-m','http.server','8765','--bind','127.0.0.1' -WorkingDirectory 'docs' -WindowStyle Hidden
```

Expected: `http://127.0.0.1:8765/원가계산서-검증-현황과-개선안.html`이 HTTP 200으로 열린다.

- [ ] **Step 2: 인앱 Browser에서 데스크톱 화면 검증**

브라우저에서 1440×900 뷰포트로 페이지를 열고 다음을 확인한다.

```text
- 제목과 기존 요약이 보인다.
- 공고문 초안 생성 목차 링크가 #notice-plan으로 이동한다.
- 10단계가 순서대로 보이고 각 단계에 상태 태그가 있다.
- 상세표의 동작·부족 열이 겹치거나 잘리지 않는다.
- 콘솔에 관련 error 또는 warning이 없다.
```

- [ ] **Step 3: 인앱 Browser에서 모바일 화면 검증**

390×844 뷰포트에서 흐름과 입력 그룹이 한 열로 쌓이고 가로 페이지 넘침, 글자 겹침, 잘린 상태 태그가 없는지 확인한다. 기존 SVG는 자체 가로 스크롤 영역 안에만 머물러야 한다.

- [ ] **Step 4: 상호작용과 기존 체크리스트 회귀 검증**

`공고문 초안 생성` 목차 링크를 클릭해 `#notice-plan`으로 이동한다. 기존 체크리스트 하나를 선택한 뒤 완료 숫자가 1 증가하고 새로고침 후 유지되는지 확인한다.

- [ ] **Step 5: 인쇄 스타일 정적 검사와 화면 점검**

인쇄 CSS가 새 `.notice-flow`, 상세표와 입력 그룹을 페이지 폭 안에 배치하도록 확인한다. 인쇄 미리보기 또는 인쇄 미디어 에뮬레이션에서 목차와 진행문구는 숨고 새 섹션의 텍스트가 잘리지 않아야 한다.

- [ ] **Step 6: QA에서 발견된 문제를 최소 수정하고 동일 검사를 재실행**

수정이 필요하면 같은 HTML의 `.notice-*` 스타일 또는 마크업만 조정하고 Task 1 Step 7의 구조 검사, 데스크톱·모바일 화면, 목차 클릭과 체크리스트 회귀 검사를 다시 수행한다.

- [ ] **Step 7: 최종 검증과 커밋 상태 확인**

```powershell
git diff --check
git status --short
```

Expected: HTML에 공백 오류가 없고, 사용자 소유의 기존 변경은 그대로 남으며 이번 HTML 변경은 커밋되어 있다. QA 보정 커밋이 필요하면 HTML만 스테이징해 `docs: refine notice report layout`으로 커밋한다.
