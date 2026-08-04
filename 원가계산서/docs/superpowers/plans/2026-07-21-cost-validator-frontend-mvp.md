# Cost Validator Frontend MVP Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a React/Vite frontend-only MVP that demonstrates upload, document recognition, validation progress, result dashboard, and evidence detail flows for the cost estimate validation system.

**Architecture:** The app is a client-side workflow with seeded validation data. `App` owns the current step, uploaded file metadata, selected result, and filter state; screens and components consume typed props. Validation data is shaped to match the PRD so a future FastAPI response can replace the seeded data.

**Tech Stack:** React, TypeScript, Vite, Vitest, React Testing Library, CSS modules via plain CSS files, lucide-react for icons.

## Global Constraints

- Frontend-only demo: do not implement real Excel parsing, PDF parsing, backend APIs, auth, persistence, or export.
- Uploaded file names must be reflected in the UI.
- Result data must follow the PRD-style validation result shape.
- The first screen must be the usable upload workflow, not a landing page.
- The dashboard must show summary counts, filters, result list, and evidence panel.
- Result selection and filters must work with local state.
- The UI must remain usable at desktop and mobile widths without text overlap.
- Keep the interface calm, dense, and work-focused.

---

## File Structure

- Create `package.json`: scripts and frontend dependencies.
- Create `index.html`: Vite root.
- Create `tsconfig.json`, `tsconfig.node.json`, `vite.config.ts`, `vitest.setup.ts`: TypeScript and test setup.
- Create `src/main.tsx`: React entry.
- Create `src/App.tsx`: workflow orchestration and top-level state.
- Create `src/types.ts`: shared domain types.
- Create `src/data/sampleResults.ts`: seeded recognition, progress, and validation data.
- Create `src/utils/results.ts`: count, sort, filter, and navigation helpers.
- Create `src/components/ui.tsx`: shared buttons, badges, upload panels, cards.
- Create `src/screens/UploadScreen.tsx`: upload and procurement selection.
- Create `src/screens/RecognitionScreen.tsx`: document recognition summary.
- Create `src/screens/ProgressScreen.tsx`: staged validation progress.
- Create `src/screens/DashboardScreen.tsx`: summary, filters, result list, evidence panel.
- Create `src/screens/DetailScreen.tsx`: full evidence details.
- Create `src/styles.css`: design tokens, layout, responsive behavior.
- Create `src/App.test.tsx`: user-flow tests.
- Create `src/utils/results.test.ts`: data helper tests.

## Task 1: Project Scaffold And Test Harness

**Files:**
- Create: `package.json`
- Create: `index.html`
- Create: `tsconfig.json`
- Create: `tsconfig.node.json`
- Create: `vite.config.ts`
- Create: `vitest.setup.ts`
- Create: `src/main.tsx`
- Create: `src/App.tsx`
- Create: `src/App.test.tsx`

**Interfaces:**
- Produces: `App(): JSX.Element`
- Consumes: none

- [ ] **Step 1: Write the failing smoke test**

```tsx
import { render, screen } from '@testing-library/react';
import App from './App';

test('renders the upload workflow first', () => {
  render(<App />);

  expect(screen.getByRole('heading', { name: '원가계산서 자동검증' })).toBeInTheDocument();
  expect(screen.getByText('검증할 원가계산서')).toBeInTheDocument();
  expect(screen.getByRole('button', { name: '자동검증 시작' })).toBeDisabled();
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pnpm test -- --run src/App.test.tsx`

Expected: fail because package setup and `App` do not exist yet.

- [ ] **Step 3: Add the minimal scaffold**

Create the Vite files and a minimal `App` rendering the heading, upload label, and disabled button.

- [ ] **Step 4: Run test to verify it passes**

Run: `pnpm test -- --run src/App.test.tsx`

Expected: one passing test.

## Task 2: Domain Types, Sample Data, And Result Helpers

**Files:**
- Create: `src/types.ts`
- Create: `src/data/sampleResults.ts`
- Create: `src/utils/results.ts`
- Create: `src/utils/results.test.ts`

**Interfaces:**
- Produces: `ValidationResult`, `RecognitionSummary`, `ProgressStep`
- Produces: `getStatusCounts(results)`, `filterResults(results, filters)`, `sortResultsByPriority(results)`, `getAdjacentResult(results, selectedId, direction)`
- Consumes: none

- [ ] **Step 1: Write failing helper tests**

```ts
import { filterResults, getAdjacentResult, getStatusCounts, sortResultsByPriority } from './results';
import { sampleResults } from '../data/sampleResults';

test('counts validation statuses', () => {
  expect(getStatusCounts(sampleResults)).toEqual({
    ERROR: 3,
    NEEDS_REVIEW: 7,
    WARNING: 2,
    OK: 126,
    UNAVAILABLE: 0,
  });
});

test('filters by status and validation type', () => {
  const filtered = filterResults(sampleResults, {
    status: 'ERROR',
    validationType: 'RATE',
    sheetName: 'ALL',
  });

  expect(filtered).toHaveLength(1);
  expect(filtered[0].item.canonicalName).toBe('산업안전보건관리비');
});

test('sorts review-priority results first', () => {
  const sorted = sortResultsByPriority(sampleResults);

  expect(sorted[0].status).toBe('ERROR');
  expect(sorted[0].item.canonicalName).toBe('산업안전보건관리비');
});

test('finds adjacent review result', () => {
  expect(getAdjacentResult(sampleResults, 'vr-001', 'next')?.resultId).toBe('vr-002');
});
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pnpm test -- --run src/utils/results.test.ts`

Expected: fail because data and helpers do not exist.

- [ ] **Step 3: Add types, seeded data, and helpers**

Implement PRD-shaped sample validation results with 3 errors, 7 needs-review items, 2 warnings, and 126 normal generated items. Add recognition data and progress steps for the screens.

- [ ] **Step 4: Run tests to verify they pass**

Run: `pnpm test -- --run src/utils/results.test.ts`

Expected: all helper tests pass.

## Task 3: Upload, Recognition, And Progress Screens

**Files:**
- Create: `src/components/ui.tsx`
- Create: `src/screens/UploadScreen.tsx`
- Create: `src/screens/RecognitionScreen.tsx`
- Create: `src/screens/ProgressScreen.tsx`
- Modify: `src/App.tsx`
- Modify: `src/App.test.tsx`
- Modify: `src/styles.css`

**Interfaces:**
- Produces: `UploadScreen`, `RecognitionScreen`, `ProgressScreen`
- Consumes: `RecognitionSummary`, `ProgressStep`

- [ ] **Step 1: Write failing workflow test**

```tsx
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import App from './App';

test('moves from upload to recognition and progress', async () => {
  render(<App />);

  const excelInput = screen.getByLabelText('Excel 파일 선택');
  const pdfInput = screen.getByLabelText('PDF 파일 추가');

  fireEvent.change(excelInput, {
    target: { files: [new File(['demo'], '추정가격내역서_최종.xlsx')] },
  });
  fireEvent.change(pdfInput, {
    target: { files: [new File(['pdf'], '2026년 건설공사 원가계산 제비율 적용기준.pdf')] },
  });

  fireEvent.click(screen.getByRole('button', { name: '자동검증 시작' }));

  expect(await screen.findByRole('heading', { name: '문서 인식 결과' })).toBeInTheDocument();
  expect(screen.getByText('추정가격내역서_최종.xlsx')).toBeInTheDocument();

  fireEvent.click(screen.getByRole('button', { name: '검증 실행' }));

  expect(await screen.findByRole('heading', { name: '원가계산서를 검증하고 있습니다' })).toBeInTheDocument();
  await waitFor(() => expect(screen.getByRole('heading', { name: '검증결과' })).toBeInTheDocument(), {
    timeout: 2500,
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pnpm test -- --run src/App.test.tsx`

Expected: fail because workflow screens do not exist.

- [ ] **Step 3: Implement screens and shared UI primitives**

Add upload panels, procurement segmented controls, recognition tables, progress stage list, and automatic transition from progress to dashboard.

- [ ] **Step 4: Run test to verify it passes**

Run: `pnpm test -- --run src/App.test.tsx`

Expected: upload-to-dashboard workflow test passes.

## Task 4: Dashboard, Filters, Evidence Panel, And Detail View

**Files:**
- Create: `src/screens/DashboardScreen.tsx`
- Create: `src/screens/DetailScreen.tsx`
- Modify: `src/App.tsx`
- Modify: `src/App.test.tsx`
- Modify: `src/styles.css`

**Interfaces:**
- Produces: `DashboardScreen`, `DetailScreen`
- Consumes: `ValidationResult[]`, filters, selected result id

- [ ] **Step 1: Write failing dashboard interaction tests**

```tsx
import { fireEvent, render, screen } from '@testing-library/react';
import App from './App';

test('dashboard filters results and updates the evidence panel', async () => {
  render(<App initialStep="dashboard" />);

  expect(screen.getByRole('heading', { name: '검증결과' })).toBeInTheDocument();
  expect(screen.getByText('산업안전보건관리비')).toBeInTheDocument();

  fireEvent.click(screen.getByRole('button', { name: '확인 필요 7' }));

  expect(screen.queryByText('산업안전보건관리비')).not.toBeInTheDocument();
  expect(screen.getByText('국민건강보험료')).toBeInTheDocument();

  fireEvent.click(screen.getByText('국민건강보험료'));

  expect(screen.getByText('끝수처리 확인이 필요합니다.')).toBeInTheDocument();
});

test('opens the detailed evidence view', () => {
  render(<App initialStep="dashboard" />);

  fireEvent.click(screen.getByRole('button', { name: '상세 근거 보기' }));

  expect(screen.getByRole('heading', { name: '산업안전보건관리비 상세' })).toBeInTheDocument();
  expect(screen.getByText('시스템 계산과정')).toBeInTheDocument();
  expect(screen.getByText('PDF 판단근거')).toBeInTheDocument();
});
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pnpm test -- --run src/App.test.tsx`

Expected: fail because dashboard and detail view do not exist.

- [ ] **Step 3: Implement dashboard, filtering, evidence, and detail screens**

Add result list ordering, status filter buttons, validation-type filters, sheet selector, evidence panel, detail layout, and previous/next review navigation.

- [ ] **Step 4: Run tests to verify they pass**

Run: `pnpm test -- --run src/App.test.tsx`

Expected: dashboard interaction tests pass.

## Task 5: Visual Polish, Responsive Layout, And Browser Verification

**Files:**
- Modify: `src/styles.css`
- Modify: `src/components/ui.tsx`
- Modify: `src/screens/*.tsx`

**Interfaces:**
- Consumes: all implemented screens.
- Produces: polished responsive UI.

- [ ] **Step 1: Run full test and build checks**

Run: `pnpm test -- --run`

Expected: all tests pass.

Run: `pnpm build`

Expected: production build succeeds.

- [ ] **Step 2: Start local app**

Run: `pnpm dev -- --host 127.0.0.1`

Expected: Vite serves the app on a local URL.

- [ ] **Step 3: Verify in browser**

Check the upload flow, dashboard filters, result selection, detail view, desktop width, and mobile width. Confirm no text overlap or clipped primary controls.

- [ ] **Step 4: Fix visual issues and rerun checks**

If any layout, interaction, or build issue appears, fix the implementation and rerun `pnpm test -- --run` and `pnpm build`.

## Self-Review Checklist

- Spec coverage: upload, recognition, progress, dashboard, evidence panel, detail view, filters, and file-name reflection are all covered.
- Placeholder scan: no task relies on unspecified behavior; seeded data replaces real parsing for this MVP.
- Type consistency: all screens consume `ValidationResult`, `RecognitionSummary`, and `ProgressStep` from `src/types.ts`.
