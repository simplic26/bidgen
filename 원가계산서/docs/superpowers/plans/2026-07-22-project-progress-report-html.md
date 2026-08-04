# 6조 프로젝트 진행상황 HTML Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 승인된 시안과 실제 앱 화면을 바탕으로 팀원이 파싱 방식, 검증 기준, 구현 범위와 한계를 이해할 수 있는 독립형 프로젝트 진행상황 HTML을 만든다.

**Architecture:** 승인된 `progress-report-layout-v5.html`을 단일 파일 산출물의 기반으로 사용하고, 이미지 경로를 저장소 상대경로로 바꾼다. 콘텐츠·스타일·상호작용은 모두 최종 HTML 내부에 두되 실제 앱 화면 PNG는 `qa/`의 기존 파일을 참조하며, Playwright 기반 독립 QA 스크립트로 데스크톱·모바일·인쇄·키보드 동작을 검증한다.

**Tech Stack:** 시맨틱 HTML5, 내장 CSS, 바닐라 JavaScript, Playwright 1.61.1, Node.js ESM

## Global Constraints

- 산출물은 `docs_0722_status/6조_프로젝트_진행상황_260722.html` 한 개의 독립 HTML 문서로 만든다.
- 외부 패키지, 빌드 과정, CDN, 폰트 서버, 네트워크 요청을 사용하지 않는다.
- 기존 React 앱 코드와 스타일은 수정하지 않는다.
- 실제 화면은 `../qa/upload-desktop.png`, `../qa/recognition-desktop.png`, `../qa/dashboard-desktop.png`, `../qa/detail-desktop.png`를 사용한다.
- 숫자 판정은 AI 없는 규칙 기반 방식이며, AI API는 후속 보조 기능 후보로만 표현한다.
- 공고문 생성과 관련 법령 수집·해석은 미구현 상태로 표시한다.
- 제비율과 노임단가는 사용자가 별도로 첨부한 외부 기준자료를 판정 기준으로 사용한다.
- 원가계산서 내부의 제비율·노임 시트는 구조와 적용값 확인용이며 외부 기준자료를 대체하지 않는다.
- `prefers-reduced-motion`을 존중하고 키보드로 모든 버튼과 이미지 확대 기능을 조작할 수 있게 한다.

---

### Task 1: 독립 QA 계약 작성

**Files:**
- Create: `qa/project-report.test.mjs`
- Test: `qa/project-report.test.mjs`

**Interfaces:**
- Consumes: 최종 HTML 경로와 `qa/` 이미지 네 개
- Produces: `node qa/project-report.test.mjs`로 실행할 수 있는 콘텐츠·레이아웃·상호작용 회귀 검사

- [ ] **Step 1: 최종 문서의 필수 계약을 검사하는 실패 테스트 작성**

```js
import assert from 'node:assert/strict';
import { existsSync } from 'node:fs';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { chromium } from 'playwright';

const reportPath = resolve('docs_0722_status/6조_프로젝트_진행상황_260722.html');
assert.ok(existsSync(reportPath), `보고서가 없습니다: ${reportPath}`);

const browser = await chromium.launch({ headless: true });
try {
  const page = await browser.newPage({ viewport: { width: 1440, height: 900 } });
  await page.goto(pathToFileURL(reportPath).href, { waitUntil: 'domcontentloaded' });

  const bodyText = await page.locator('body').innerText();
  for (const phrase of [
    '원가계산서를 어떻게 파싱했는가',
    'Workbook IR',
    '사용자가 별도로 첨부한 외부 제비율표',
    '사용자가 별도로 첨부한 외부 노임단가표',
    'AI 없이',
    '공고문',
    '관련 법령',
  ]) assert.ok(bodyText.includes(phrase), `필수 문구 누락: ${phrase}`);

  assert.equal(await page.locator('.sidebar .nav a').count(), 10);
  assert.equal(await page.locator('.shot img').count(), 4);
  const images = await page.locator('.shot img').evaluateAll((nodes) =>
    nodes.map((image) => ({ complete: image.complete, width: image.naturalWidth })),
  );
  assert.ok(images.every((image) => image.complete && image.width > 0), JSON.stringify(images));

  await page.locator('.shot img').first().focus();
  await page.keyboard.press('Enter');
  assert.equal(await page.locator('#lightbox').evaluate((dialog) => dialog.open), true);
  await page.keyboard.press('Escape');
  assert.equal(await page.locator('#lightbox').evaluate((dialog) => dialog.open), false);

  await page.emulateMedia({ media: 'print' });
  assert.equal(await page.locator('.sidebar').evaluate((node) => getComputedStyle(node).display), 'none');

  await page.setViewportSize({ width: 390, height: 844 });
  await page.emulateMedia({ media: 'screen' });
  const overflow = await page.evaluate(() => document.documentElement.scrollWidth - innerWidth);
  assert.ok(overflow <= 1, `모바일 가로 넘침: ${overflow}px`);
  assert.equal(await page.locator('[data-mobile-menu]').count(), 1);

  const externalResources = await page.locator('link[href^="http"],script[src^="http"],img[src^="http"]').count();
  assert.equal(externalResources, 0);
} finally {
  await browser.close();
}

console.log('project report QA passed');
```

- [ ] **Step 2: 테스트를 실행해 산출물 부재로 실패하는지 확인**

Run: `node qa/project-report.test.mjs`

Expected: FAIL with `보고서가 없습니다: ...6조_프로젝트_진행상황_260722.html`

- [ ] **Step 3: QA 계약 커밋**

```bash
git add qa/project-report.test.mjs
git commit -m "test: define project report QA contract"
```

### Task 2: 승인된 콘텐츠와 시각 구조를 최종 HTML로 제작

**Files:**
- Create: `docs_0722_status/6조_프로젝트_진행상황_260722.html`
- Reference: `.superpowers/brainstorm/1333-1784704715/content/progress-report-layout-v5.html`
- Reference: `docs/superpowers/specs/2026-07-22-project-progress-report-html-design.md`
- Test: `qa/project-report.test.mjs`

**Interfaces:**
- Consumes: 승인된 v5 시안, 설계 문서, `qa/`의 실제 앱 화면 PNG 네 개
- Produces: 네이비 사이드바, 10개 목차, 파싱 6단계, 기준자료 우선순위, 입력·판단·출력, 화면 예시, 진행상황, 운영기준, 문제와 계획을 포함한 단일 HTML

- [ ] **Step 1: 승인된 시안을 최종 경로로 복사하고 임시 시안 기능을 제거**

Run:

```powershell
Copy-Item -LiteralPath '.superpowers/brainstorm/1333-1784704715/content/progress-report-layout-v5.html' -Destination 'docs_0722_status/6조_프로젝트_진행상황_260722.html'
```

Then remove the visual-companion-only choice buttons and `toggleSelect` calls. Replace the four image sources exactly:

```html
<img src="../qa/upload-desktop.png" alt="원가계산서와 기준자료 업로드 화면">
<img src="../qa/recognition-desktop.png" alt="문서 인식 결과 화면">
<img src="../qa/dashboard-desktop.png" alt="검증결과 대시보드">
<img src="../qa/detail-desktop.png" alt="항목별 상세 근거 화면">
```

- [ ] **Step 2: 파싱 성공 방식과 기준자료 우선순위가 설계 문서와 일치하는지 확정**

The final section order and headings must be exactly:

```html
<section id="overview">한눈에 보기</section>
<section id="intro">프로젝트 소개</section>
<section id="scope">현재 구현 경계</section>
<section id="parsing">원가계산서를 어떻게 파싱했는가</section>
<section id="io">입력·판단·출력</section>
<section id="flow">사용자 흐름</section>
<section id="progress">진행 상황</section>
<section id="guide">꼭 알아둘 점</section>
<section id="issues">현재 문제</section>
<section id="next">다음 계획</section>
```

The source-of-truth copy must state:

```html
<strong>별도로 첨부한 외부 기준자료</strong>
<p>사용자가 업로드한 외부 제비율표 Excel과 외부 노임단가표 Excel을 공식 비교 기준으로 사용합니다.</p>
<strong>구조·적용값 확인용</strong>
<p>원가계산서 통합문서 안의 제비율·노임 시트는 적용값과 참조 구조를 파악하는 자료일 뿐, 판정 기준으로 대체하지 않습니다.</p>
```

- [ ] **Step 3: 콘텐츠 계약을 실행해 상호작용·모바일 항목만 남아 실패하는지 확인**

Run: `node qa/project-report.test.mjs`

Expected: content phrases, ten navigation links, and four images pass; test fails because `[data-mobile-menu]` or print/mobile interaction is not finished.

- [ ] **Step 4: 콘텐츠 산출물 커밋**

```bash
git add docs_0722_status/6조_프로젝트_진행상황_260722.html
git commit -m "docs: add detailed project progress report"
```

### Task 3: 접근 가능한 탐색·확대·인쇄 동작 완성

**Files:**
- Modify: `docs_0722_status/6조_프로젝트_진행상황_260722.html`
- Test: `qa/project-report.test.mjs`

**Interfaces:**
- Consumes: Task 2의 10개 섹션 ID와 네 개 `.shot img`
- Produces: `toggleMobileMenu()`, `openLightbox(image)`, `openLightboxWithKey(event, image)`, `closeLightboxOnBackdrop(event)`, `window.print()` 동작

- [ ] **Step 1: 인쇄와 모바일 목차 제어 버튼 추가**

```html
<div class="report-tools" aria-label="보고서 도구">
  <button type="button" onclick="window.print()">인쇄/PDF</button>
  <button type="button" data-mobile-menu aria-expanded="false" aria-controls="report-nav" onclick="toggleMobileMenu(this)">목차</button>
</div>
```

Set `id="report-nav"` on the sidebar navigation and add:

```js
function toggleMobileMenu(button) {
  const sidebar = document.querySelector('.sidebar');
  const open = sidebar.classList.toggle('mobile-open');
  button.setAttribute('aria-expanded', String(open));
}
```

- [ ] **Step 2: 현재 목차 강조와 이미지 모달 접근성 완성**

```js
const links = [...document.querySelectorAll('.nav a')];
const sections = links.map((link) => document.querySelector(link.hash)).filter(Boolean);
const observer = new IntersectionObserver((entries) => {
  const visible = entries.filter((entry) => entry.isIntersecting).sort((a, b) => b.intersectionRatio - a.intersectionRatio)[0];
  if (!visible) return;
  links.forEach((link) => link.classList.toggle('active', link.hash === `#${visible.target.id}`));
}, { rootMargin: '-20% 0px -65% 0px', threshold: [0, 0.1, 0.5] });
sections.forEach((section) => observer.observe(section));
```

Keep `role="button"`, `tabindex="0"`, Enter/Space handling, native `<dialog>`, Esc close, backdrop close, and an explicit close button for each screenshot.

- [ ] **Step 3: 모바일·인쇄·감소된 모션 CSS 추가**

```css
.report-tools{position:fixed;right:20px;bottom:20px;z-index:20;display:flex;gap:8px}
.report-tools button{border:1px solid var(--line);border-radius:10px;background:#fff;padding:10px 13px;box-shadow:var(--shadow);cursor:pointer}
.nav a.active{background:rgba(255,255,255,.11);color:#fff}
@media(max-width:900px){
  .sidebar.mobile-open{display:block;width:min(82vw,300px);box-shadow:20px 0 60px rgba(0,0,0,.3)}
  [data-mobile-menu]{display:inline-block}
}
@media(min-width:901px){[data-mobile-menu]{display:none}}
@media(prefers-reduced-motion:reduce){html{scroll-behavior:auto}*,*::before,*::after{animation-duration:.01ms!important;transition-duration:.01ms!important}}
@media print{.report-tools{display:none!important}}
```

- [ ] **Step 4: QA 계약 전체를 실행해 통과 확인**

Run: `node qa/project-report.test.mjs`

Expected: `project report QA passed`

- [ ] **Step 5: 상호작용 완성 커밋**

```bash
git add docs_0722_status/6조_프로젝트_진행상황_260722.html qa/project-report.test.mjs
git commit -m "docs: finish report navigation and print behavior"
```

### Task 4: 최종 시각 QA와 저장소 상태 확인

**Files:**
- Verify: `docs_0722_status/6조_프로젝트_진행상황_260722.html`
- Verify: `qa/project-report.test.mjs`

**Interfaces:**
- Consumes: 완성된 HTML과 QA 스크립트
- Produces: 데스크톱·모바일·인쇄 환경에서 재현 가능한 검증 증거

- [ ] **Step 1: 앱 회귀 테스트와 보고서 QA 실행**

Run:

```powershell
pnpm test
node qa/project-report.test.mjs
```

Expected: TypeScript exits 0; report prints `project report QA passed`.

- [ ] **Step 2: HTML 정적 오류와 Git 변경 범위 확인**

Run:

```powershell
git diff --check
git status --short
```

Expected: no whitespace errors; only the planned report and QA/plan files are changed or newly committed. `.superpowers/` remains untracked and is not staged.

- [ ] **Step 3: 최종 브라우저 검토**

Use Playwright at 1440×900 and 390×844 to confirm: no horizontal overflow, all screenshots load, modal opens/closes, mobile menu opens, current navigation changes while scrolling, and print media hides navigation/tools.

- [ ] **Step 4: 완료 상태 기록**

If any verification changed tracked files, commit only those verified changes:

```bash
git add docs_0722_status/6조_프로젝트_진행상황_260722.html qa/project-report.test.mjs
git commit -m "test: verify standalone project report"
```
