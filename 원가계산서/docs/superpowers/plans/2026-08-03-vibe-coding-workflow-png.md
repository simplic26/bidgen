# Vibe Coding Workflow PNG Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 바이브코딩 웹앱 개발의 일반적인 10단계를 1920×1080px 흰 배경 PNG 한 장으로 제작한다.

**Architecture:** 정확한 한글 조판을 위해 HTML/CSS 기반의 고정 크기 벡터 도식을 만든 뒤 Playwright로 PNG를 캡처한다. 결과 PNG를 프로젝트 루트에 저장하고 이미지 크기와 시각적 완성도를 검증한다.

**Tech Stack:** HTML, CSS, Playwright, Chromium

## Global Constraints

- 최종 파일명은 `바이브코딩-웹앱-개발절차.png`이다.
- 최종 크기는 정확히 1920×1080px이다.
- 배경은 순백색이며 모든 문구는 한국어로 표시한다.
- 기획, 구현·반복, 검증·출시의 세 구간과 전체 10단계를 빠짐없이 표시한다.
- 구현 반복 고리와 `작게 요청 → 직접 실행 → 결과 전달 → 수정` 원칙을 강조한다.

---

### Task 1: 고정 크기 도식 제작 및 PNG 렌더링

**Files:**
- Create: `바이브코딩-웹앱-개발절차.html`
- Create: `scripts/render-vibe-coding-workflow.mjs`
- Create: `바이브코딩-웹앱-개발절차.png`

**Interfaces:**
- Consumes: 디자인 명세의 10단계 문구와 1920×1080px 규격
- Produces: 프로젝트 루트의 최종 PNG 파일

- [ ] **Step 1: HTML 도식 작성**

  세 구간 카드, 10개 단계, 반복 고리, 하단 원칙 띠를 1920×1080px 고정 캔버스에 배치한다.

- [ ] **Step 2: 렌더링 스크립트 작성**

```javascript
await page.setViewportSize({ width: 1920, height: 1080 });
await page.goto(pathToFileURL(source).href);
await page.screenshot({ path: output, fullPage: false });
```

- [ ] **Step 3: PNG 생성**

Run: `node scripts/render-vibe-coding-workflow.mjs`

Expected: 프로젝트 루트에 `바이브코딩-웹앱-개발절차.png`가 생성된다.

### Task 2: 결과 검증 및 정리

**Files:**
- Verify: `바이브코딩-웹앱-개발절차.png`
- Delete: `바이브코딩-웹앱-개발절차.html`
- Delete: `scripts/render-vibe-coding-workflow.mjs`

**Interfaces:**
- Consumes: Task 1의 PNG
- Produces: 검증을 통과한 PNG 단일 산출물

- [ ] **Step 1: 픽셀 크기 검증**

```javascript
const size = await page.evaluate(() => ({ width: innerWidth, height: innerHeight }));
if (size.width !== 1920 || size.height !== 1080) throw new Error('invalid size');
```

- [ ] **Step 2: 이미지를 직접 열어 시각 검증**

Expected: 한글 깨짐, 잘림, 겹침 없이 모든 단계와 연결 방향이 보인다.

- [ ] **Step 3: 임시 HTML과 렌더링 스크립트 제거**

Expected: 최종 PNG와 설계·계획 문서만 남는다.
