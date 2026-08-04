# 세 문서 정규화 검수실 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 원가계산서·제비율표·노임단가표의 전체 정규화 결과와 원본 셀 근거를 검증 실행 전에 탭형 검수실에서 확인하고, 원가계산서에서 추출한 공사조건으로 적용 제비율을 재현 가능하게 선택한다.

**Architecture:** 세 Excel을 공통 `WorkbookIR`로 읽은 뒤 문서별 정규화기로 분리한다. 제비율표는 병합범위를 인덱스화해 값 셀에서 조건 헤더를 역추적하고, 원가계산서 조건 추출기와 규칙 매처를 별도 순수 함수로 둔다. React는 정규화 결과를 상태로 보관하고 검색·필터·근거 패널만 렌더하며 파싱을 렌더 경로에서 반복하지 않는다.

**Tech Stack:** React 19, TypeScript 5.7, Vite 6, SheetJS `xlsx` 0.18.5, Node test runner, Playwright 1.61.

## Global Constraints

- 구현 전 `superpowers:using-git-worktrees`로 격리 작업공간을 만들고 기존 미커밋 변경을 보존한다.
- 모든 새 동작은 실패하는 테스트를 먼저 확인한 뒤 구현한다.
- 대표 제비율표는 저장소 이력의 `public/_ref_rate.xlsx`, 대표 노임단가표는 `public/_ref_labor.xlsx`를 QA fixture로 복원해 사용한다.
- 정규화 화면은 읽기 전용이며 사용자 보정 UI를 구현하지 않는다.
- 숫자·적용 규칙 판정에 AI/LLM을 사용하지 않는다.
- 제비율표의 모든 조건 조합, 적용 제외, 복합식, 관련규정, 원본 오류를 보존한다.
- 비어 있지 않은 원본 셀은 규칙·조건·값·주석·규정·장식·미분류 중 하나로 집계한다.
- 파일 읽기 실패 또는 필수 문서의 정규화 행 0개만 검증 실행을 차단한다.
- 새 런타임 의존성을 추가하지 않는다.
- React 필터링은 원시 배열과 원시 필터 값을 기준으로 `useMemo`에서 수행하고 컴포넌트 내부 컴포넌트 선언을 피한다.
- 설계 기준은 `docs/superpowers/specs/2026-07-22-normalization-audit-design.md`다.

---

## File Structure

- `src/types.ts` — 공통 정규화 계약, 공사조건, 제비율 규칙, 노임단가 행 타입.
- `src/utils/normalization/evidence.ts` — 셀 좌표, 병합 인덱스, 근거, 커버리지 공통 함수.
- `src/utils/normalization/projectContext.ts` — 원가계산서 전체 시트의 공사종류·금액·기간 후보 추출.
- `src/utils/normalization/laborWorkbook.ts` — 노임단가표 정규화.
- `src/utils/normalization/rateParsing.ts` — 금액·기간 범위, 요율, 복합식 파싱.
- `src/utils/normalization/rateStructure.ts` — 제비율표 병합 헤더와 섹션 구조 인덱싱.
- `src/utils/normalization/rateWorkbook.ts` — 모든 제비율 섹션 규칙 생성, 규정 연결, 커버리지.
- `src/utils/normalization/ruleMatcher.ts` — 공사조건에 적용되는 규칙 선택.
- `src/utils/normalization/viewModel.ts` — 탭 요약, 치명 오류, 검색·필터 순수 함수.
- `src/components/normalization/ConditionSummary.tsx` — 자동 추출 공사조건과 보완 입력.
- `src/components/normalization/NormalizationTabs.tsx` — 세 문서 탭과 상태 배지.
- `src/components/normalization/NormalizationTable.tsx` — 문서별 읽기 전용 표.
- `src/components/normalization/EvidencePanel.tsx` — 선택 행의 셀·병합범위·규정 근거.
- `src/screens/NormalizationAuditScreen.tsx` — 검수실 화면 조립과 검증 실행 게이트.
- `src/screens/UploadScreen.tsx` — 문서별 `NormalizedDocument` 반환.
- `src/App.tsx` — `normalization` 단계와 정규화 상태, 검증 엔진 배선.
- `src/utils/criteria.ts` — 기존 공개 함수의 호환 래퍼만 유지.
- `src/styles.css` — 탭, 요약, 필터, 표, 근거 패널, 반응형 레이아웃.
- `qa/helpers/vite-module.mjs` — Vite SSR 모듈 로더와 fixture `File` 헬퍼.
- `qa/helpers/create-cost-fixture.mjs` — 공사조건과 기존 원가 정규화 UI용 재현 가능한 원가계산서 fixture 생성기.
- `qa/fixtures/cost-reference.xlsx` — 테스트용 원가계산서.
- `qa/fixtures/rate-reference.xlsx` — 대표 제비율표.
- `qa/fixtures/labor-reference.xlsx` — 대표 노임단가표.
- `qa/normalization-common.test.mjs` — 병합·근거·커버리지.
- `qa/project-context.test.mjs` — 공사조건 후보.
- `qa/labor-normalization.test.mjs` — 노임단가 전체행.
- `qa/rate-parsing.test.mjs` — 범위·요율·복합식.
- `qa/rate-normalization.test.mjs` — 대표 제비율표 회귀.
- `qa/rule-matcher.test.mjs` — 적용 규칙 선택.
- `qa/normalization-view-model.test.mjs` — 화면 필터·게이트.
- `qa/normalization-screen.test.mjs` — 검수실 SSR 구조와 차단 버튼.

---

### Task 1: 대표 fixture와 Node 테스트 하네스

**Files:**
- Create: `qa/fixtures/rate-reference.xlsx`
- Create: `qa/fixtures/labor-reference.xlsx`
- Create: `qa/fixtures/cost-reference.xlsx`
- Create: `qa/helpers/vite-module.mjs`
- Create: `qa/helpers/create-cost-fixture.mjs`
- Create: `qa/fixtures.test.mjs`
- Modify: `package.json`

**Interfaces:**
- Consumes: Git commit `2a77f55`의 `public/_ref_rate.xlsx`, `public/_ref_labor.xlsx` blob.
- Produces: `withViteModule(modulePath, callback)`, `fixtureFile(name)`, 재현 가능한 세 Excel fixture.

- [ ] **Step 1: fixture가 없어서 실패하는 테스트를 작성한다**

```js
// qa/fixtures.test.mjs
import assert from 'node:assert/strict';
import test from 'node:test';
import XLSX from 'xlsx';
import { fixtureFile } from './helpers/vite-module.mjs';

test('ships the representative rate and labor workbooks', async () => {
  const rate = await fixtureFile('rate-reference.xlsx');
  const labor = await fixtureFile('labor-reference.xlsx');
  const cost = await fixtureFile('cost-reference.xlsx');
  const rateBook = XLSX.read(await rate.arrayBuffer(), { type: 'array' });
  const laborBook = XLSX.read(await labor.arrayBuffer(), { type: 'array' });
  const costBook = XLSX.read(await cost.arrayBuffer(), { type: 'array' });

  assert.deepEqual(rateBook.SheetNames, ['건축제비율(1.1)', '관련규정 링크']);
  assert.deepEqual(laborBook.SheetNames, ['Sheet1']);
  assert.deepEqual(costBook.SheetNames, ['표지', '원가계산서', '공사개요']);
});
```

- [ ] **Step 2: 테스트를 실행해 fixture 누락 실패를 확인한다**

Run: `node --test qa/fixtures.test.mjs`

Expected: `ENOENT` for `qa/fixtures/rate-reference.xlsx`.

- [ ] **Step 3: fixture를 저장소 이력에서 복원하고 헬퍼를 구현한다**

PowerShell에서 다음 순서로 정확한 두 blob을 복원한다.

```powershell
New-Item -ItemType Directory -Force qa\fixtures | Out-Null
git archive --format=zip --output qa\fixtures\refs.zip 2a77f55 public/_ref_rate.xlsx public/_ref_labor.xlsx
Expand-Archive -LiteralPath qa\fixtures\refs.zip -DestinationPath qa\fixtures\restored -Force
Move-Item -LiteralPath qa\fixtures\restored\public\_ref_rate.xlsx -Destination qa\fixtures\rate-reference.xlsx
Move-Item -LiteralPath qa\fixtures\restored\public\_ref_labor.xlsx -Destination qa\fixtures\labor-reference.xlsx
Remove-Item -LiteralPath qa\fixtures\refs.zip
Remove-Item -LiteralPath qa\fixtures\restored -Recurse
node qa\helpers\create-cost-fixture.mjs
```

```js
// qa/helpers/create-cost-fixture.mjs
import XLSX from 'xlsx';

const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
  ['공사명', '건축공사 정규화 테스트'],
  ['공사종류', '건축공사'],
  ['추정가격', 500000000],
]), '표지');
XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
  ['항목', '산출기초', '요율', '금액'],
  ['직접재료비', '', '', 120000000],
  ['직접노무비', '', '', 80000000],
  ['간접노무비', '직접노무비', 0.175, 14000000],
]), '원가계산서');
XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
  ['착공일', new Date('2026-01-01T00:00:00Z')],
  ['준공일', new Date('2026-06-30T00:00:00Z')],
]), '공사개요');
XLSX.writeFile(workbook, 'qa/fixtures/cost-reference.xlsx', { cellDates: true });
```

```js
// qa/helpers/vite-module.mjs
import fs from 'node:fs/promises';
import path from 'node:path';
import { createServer } from 'vite';

const root = process.cwd();

export async function fixtureFile(name) {
  const bytes = await fs.readFile(path.join(root, 'qa', 'fixtures', name));
  return new File([bytes], name, {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}

export async function withViteModule(modulePath, callback) {
  const server = await createServer({
    root,
    configFile: path.join(root, 'vite.config.mjs'),
    configLoader: 'native',
    server: { middlewareMode: true },
  });
  try {
    return await callback(await server.ssrLoadModule(modulePath));
  } finally {
    await server.close();
  }
}
```

`package.json` scripts를 다음 값으로 교체한다.

```json
{
  "test:unit": "node --test qa/fixtures.test.mjs qa/normalization-common.test.mjs qa/project-context.test.mjs qa/labor-normalization.test.mjs qa/rate-parsing.test.mjs qa/rate-normalization.test.mjs qa/rule-matcher.test.mjs qa/normalization-view-model.test.mjs qa/normalization-screen.test.mjs",
  "test": "tsc --noEmit && pnpm run test:unit"
}
```

- [ ] **Step 4: fixture 테스트가 통과하는지 확인한다**

Run: `node --test qa/fixtures.test.mjs`

Expected: `1 pass, 0 fail`.

- [ ] **Step 5: 커밋한다**

```bash
git add package.json qa/fixtures qa/helpers/vite-module.mjs qa/fixtures.test.mjs
git commit -m "test: add normalization workbook fixtures"
```

---

### Task 2: 공통 정규화 타입, 병합 인덱스, 근거와 커버리지

**Files:**
- Modify: `src/types.ts`
- Create: `src/utils/normalization/evidence.ts`
- Create: `qa/normalization-common.test.mjs`

**Interfaces:**
- Consumes: `WorkbookIR`, `IRSheet`, `IRCell`, `FileMeta`.
- Produces: `CellEvidence`, `NormalizationIssue`, `NormalizationCoverage`, `NormalizedDocument<Row>`, `buildMergeIndex`, `evidenceForCell`, `buildCoverage`.

- [ ] **Step 1: 병합범위 경계와 커버리지 실패 테스트를 작성한다**

```js
// qa/normalization-common.test.mjs
import assert from 'node:assert/strict';
import test from 'node:test';
import { withViteModule } from './helpers/vite-module.mjs';

test('limits merged header inheritance to the declared range', async () => {
  await withViteModule('/src/utils/normalization/evidence.ts', ({ buildMergeIndex }) => {
    const index = buildMergeIndex([
      { address: 'A1', mergedRange: 'A1:C2', rawValue: '조건', displayValue: '조건', dataType: 'STRING', cachedValue: '조건', numberFormat: null, hidden: false, references: [] },
    ]);
    assert.equal(index.get('C2')?.anchor.address, 'A1');
    assert.equal(index.get('D2'), undefined);
  });
});

test('reports every unclassified non-empty source cell', async () => {
  await withViteModule('/src/utils/normalization/evidence.ts', ({ buildCoverage }) => {
    const coverage = buildCoverage({ nonEmpty: 10, classified: new Set(['A1', 'B1']), normalizedRows: 1, issues: [] });
    assert.equal(coverage.unclassifiedSourceCells, 8);
    assert.equal(coverage.classifiedSourceCells, 2);
  });
});
```

- [ ] **Step 2: 테스트가 export 누락으로 실패하는지 확인한다**

Run: `node --test qa/normalization-common.test.mjs`

Expected: module/export not found failure for `evidence.ts`.

- [ ] **Step 3: 설계 문서 §4.2의 공통 타입을 `src/types.ts`에 추가한다**

다음 타입 이름과 필드를 그대로 사용한다.

```ts
export type NormalizationKind = 'COST_STATEMENT' | 'RATE_TABLE' | 'LABOR_TABLE';
export type NormalizationSeverity = 'INFO' | 'WARNING' | 'FATAL';
export type NormalizationRowStatus = 'NORMAL' | 'NEEDS_REVIEW' | 'SOURCE_ERROR';

export interface CellEvidence {
  sheetName: string;
  cell: string;
  mergedRange: string | null;
  rawValue: string;
  displayValue: string;
  numberFormat: string | null;
}

export interface NormalizationIssue {
  issueId: string;
  code: 'UNCLASSIFIED_SOURCE_CELL' | 'LOW_CONFIDENCE' | 'DUPLICATE_RULE' | 'CONDITION_CONFLICT' | 'SOURCE_FORMULA_ERROR' | 'MISSING_REQUIRED_VALUE' | 'NO_NORMALIZED_ROWS';
  severity: NormalizationSeverity;
  message: string;
  evidence: CellEvidence[];
}

export interface NormalizationCoverage {
  nonEmptySourceCells: number;
  classifiedSourceCells: number;
  unclassifiedSourceCells: number;
  normalizedRows: number;
  warningCount: number;
  fatalCount: number;
}

export interface NormalizedDocument<Row> {
  kind: NormalizationKind;
  file: FileMeta;
  ir: WorkbookIR;
  rows: Row[];
  issues: NormalizationIssue[];
  coverage: NormalizationCoverage;
}
```

- [ ] **Step 4: 병합 인덱스와 커버리지 함수를 구현한다**

`buildMergeIndex`는 `mergedRange`가 있는 anchor를 `decodeCell`/`encodeCell`로 확장하고 범위 안 좌표만 `Map`에 등록한다. `evidenceForCell(sheetName, cell)`은 `CellEvidence` 필드를 그대로 복사한다. `buildCoverage`는 `nonEmpty - classified.size`를 음수가 되지 않게 계산하고 이슈 severity를 집계한다.

```ts
export interface MergeHit { anchor: IRCell; range: string }

export function buildMergeIndex(cells: IRCell[]): Map<string, MergeHit>;
export function evidenceForCell(sheetName: string, cell: IRCell): CellEvidence;
export function buildCoverage(input: {
  nonEmpty: number;
  classified: Set<string>;
  normalizedRows: number;
  issues: NormalizationIssue[];
}): NormalizationCoverage;
```

- [ ] **Step 5: 테스트와 타입 검사를 실행한다**

Run: `node --test qa/normalization-common.test.mjs`

Expected: `2 pass, 0 fail`.

Run: `pnpm exec tsc --noEmit`

Expected: exit code 0.

- [ ] **Step 6: 커밋한다**

```bash
git add src/types.ts src/utils/normalization/evidence.ts qa/normalization-common.test.mjs
git commit -m "feat: add normalization evidence contracts"
```

---

### Task 3: 원가계산서 전체 시트 공사조건 추출

**Files:**
- Modify: `src/types.ts`
- Create: `src/utils/normalization/projectContext.ts`
- Create: `qa/project-context.test.mjs`

**Interfaces:**
- Consumes: `WorkbookIR`.
- Produces: `extractProjectContext(ir): ProjectContext`.

- [ ] **Step 1: 다른 시트 값, 날짜 계산, 충돌 후보 테스트를 작성한다**

테스트 fixture IR에는 `표지!B3=건축공사`, `요약!C8=추정가격`, `요약!D8=500000000`, `공사개요!B5=착공일`, `C5=2026-01-01`, `B6=준공일`, `C6=2026-06-30`을 넣는다. 두 번째 테스트에는 서로 다른 `추정가격` 후보 두 개를 넣는다.

```js
assert.equal(context.constructionType.value, '건축');
assert.equal(context.amount.value.basis, 'ESTIMATED_PRICE');
assert.equal(context.amount.value.value, 500000000);
assert.equal(context.durationDays.value, 181);
assert.equal(conflicted.amount, null);
assert.equal(conflicted.issues[0].code, 'CONDITION_CONFLICT');
```

- [ ] **Step 2: 테스트가 함수 누락으로 실패하는지 확인한다**

Run: `node --test qa/project-context.test.mjs`

Expected: `extractProjectContext is not a function`.

- [ ] **Step 3: 공사조건 타입과 추출기를 구현한다**

`src/types.ts`에는 설계 문서 §5.2의 `ProjectAmountBasis`, `ContextCandidate<T>`, `ProjectContext`를 추가한다. 추출기는 모든 `ir.sheets`의 문자열 셀을 라벨로 검사하고 같은 행 오른쪽 1~3셀과 다음 행 같은 열을 후보로 본다. 날짜 두 개는 UTC 날짜 차이에 1을 더해 계산한다. 후보 점수는 정확한 라벨 0.45, `COVER_SUMMARY`/`COST_SUMMARY` 역할 0.2, 수식 참조 또는 같은 행 인접 0.2, 유효 범위 0.15로 계산한다. 최고 두 후보가 서로 다른 값이고 점수 차가 0.1 미만이면 해당 필드를 `null`로 두고 `CONDITION_CONFLICT`를 만든다.

```ts
export function extractProjectContext(ir: WorkbookIR): ProjectContext;
```

- [ ] **Step 4: RED 사례와 전체 타입 검사를 GREEN으로 만든다**

Run: `node --test qa/project-context.test.mjs`

Expected: all project context tests pass.

Run: `pnpm exec tsc --noEmit`

Expected: exit code 0.

- [ ] **Step 5: 커밋한다**

```bash
git add src/types.ts src/utils/normalization/projectContext.ts qa/project-context.test.mjs
git commit -m "feat: extract project conditions from every sheet"
```

---

### Task 4: 노임단가표 전체행 정규화

**Files:**
- Modify: `src/types.ts`
- Create: `src/utils/normalization/laborWorkbook.ts`
- Modify: `src/utils/criteria.ts`
- Create: `qa/labor-normalization.test.mjs`

**Interfaces:**
- Consumes: `WorkbookIR`, `FileMeta`.
- Produces: `normalizeLaborWorkbook(ir, file): NormalizedDocument<NormalizedLaborRate>`.

- [ ] **Step 1: 실제 fixture의 유효·표식·미공표 행 테스트를 작성한다**

```js
assert.equal(doc.rows.find((row) => row.occupationCode === '1001')?.unitPrice, 215907);
assert.deepEqual(doc.rows.find((row) => row.occupationCode === '1014')?.markers, ['*']);
assert.equal(doc.rows.find((row) => row.occupationCode === '1032')?.publicationStatus, 'NOT_PUBLISHED');
assert.equal(doc.rows.find((row) => row.occupationCode === '1032')?.unitPrice, null);
assert.equal(doc.rows.find((row) => row.occupationCode === '1001')?.evidence[0].cell, 'A2');
```

- [ ] **Step 2: 테스트가 새 normalizer 누락으로 실패하는지 확인한다**

Run: `node --test qa/labor-normalization.test.mjs`

Expected: module/export missing failure.

- [ ] **Step 3: 타입과 normalizer를 구현한다**

```ts
export interface NormalizedLaborRate {
  rowId: string;
  occupationCode: string;
  occupationName: string;
  unitPrice: number | null;
  publicationStatus: 'PUBLISHED' | 'NOT_PUBLISHED' | 'INVALID';
  markers: string[];
  notes: string[];
  confidence: number;
  evidence: CellEvidence[];
}

export function normalizeLaborWorkbook(
  ir: WorkbookIR,
  file: FileMeta,
): NormalizedDocument<NormalizedLaborRate>;
```

헤더는 같은 행에서 `직종번호`, `직종명`, `노임단가`를 찾고 열 좌표를 기억한다. 그 아래 행에서 코드는 선행 `*`를 `markers`로 분리하고 숫자부를 `occupationCode`로 쓴다. 단가 `-`는 `NOT_PUBLISHED`, 유효 숫자는 `PUBLISHED`, 그 외 값은 `INVALID`로 보존한다. 기존 `buildLaborRates`는 새 행 중 `PUBLISHED`만 기존 `LaborRate`로 투영하는 호환 래퍼로 바꾼다.

- [ ] **Step 4: 테스트와 기존 criteria 회귀 테스트를 통과시킨다**

Run: `node --test qa/labor-normalization.test.mjs qa/criteria.test.mjs`

Expected: all tests pass.

- [ ] **Step 5: 커밋한다**

```bash
git add src/types.ts src/utils/normalization/laborWorkbook.ts src/utils/criteria.ts qa/labor-normalization.test.mjs
git commit -m "feat: normalize every labor rate row"
```

---

### Task 5: 제비율 숫자범위와 계산식 파서

**Files:**
- Modify: `src/types.ts`
- Create: `src/utils/normalization/rateParsing.ts`
- Create: `qa/rate-parsing.test.mjs`

**Interfaces:**
- Consumes: 한국어 금액·기간 텍스트, 셀값, 숫자형식.
- Produces: `parseNumericRange`, `parseRateValue`, `parseRateFormula`, `normalizeConstructionType`.

- [ ] **Step 1: 경계·단위·복합식 실패 테스트를 작성한다**

```js
assert.deepEqual(parseNumericRange('10억 미만', 'KRW'), { min: null, minInclusive: false, max: 1_000_000_000, maxInclusive: false, unit: 'KRW', sourceText: '10억 미만' });
assert.deepEqual(parseNumericRange('50억 이상 - 100억 미만', 'KRW'), { min: 5_000_000_000, minInclusive: true, max: 10_000_000_000, maxInclusive: false, unit: 'KRW', sourceText: '50억 이상 - 100억 미만' });
assert.equal(parseNumericRange('36개월 초과 (1096일)', 'DAY').min, 1096);
assert.equal(parseRateValue(17.5, 'General').rate, 0.175);
assert.equal(parseRateValue(0.175, '0.00%').rate, 0.175);
assert.equal(parseRateFormula('[79만원+(직공비-75억원)x0.0070%]x공기(년)').kind, 'FORMULA');
```

- [ ] **Step 2: 테스트가 함수 누락으로 실패하는지 확인한다**

Run: `node --test qa/rate-parsing.test.mjs`

Expected: module/export missing failure.

- [ ] **Step 3: 설계 문서 §6.2 타입과 파서를 구현한다**

`NumericRange`, `RateValueKind`, `AmountBasis`, `RateConditions`, `RegulationReference`, `NormalizedRateRule`을 `src/types.ts`에 추가한다. `억`, `만원`, `천원`, `개월`, `일`, `년`을 각각 원·일·년 기준 숫자로 변환한다. `미만`, `이하`, `초과`, `이상`을 경계 플래그로 보존한다. `%` 형식 숫자는 이미 소수인 셀값을 다시 100으로 나누지 않는다. 일반 형식의 1 초과 100 미만 값만 100으로 나눈다.

```ts
export function parseNumericRange(text: string, unit: NumericRange['unit']): NumericRange | null;
export function parseRateValue(value: unknown, numberFormat: string | null): { kind: RateValueKind; rate: number | null; fixedAmount: number | null };
export function parseRateFormula(text: string): { kind: RateValueKind; formula: string; rate: number | null; fixedAmount: number | null };
export function normalizeConstructionType(text: string): string[];
```

- [ ] **Step 4: 범위와 복합식 테스트를 통과시킨다**

Run: `node --test qa/rate-parsing.test.mjs`

Expected: all parsing assertions pass.

- [ ] **Step 5: 커밋한다**

```bash
git add src/types.ts src/utils/normalization/rateParsing.ts qa/rate-parsing.test.mjs
git commit -m "feat: parse rate ranges and composite formulas"
```

---

### Task 6: 제비율표 구조 인덱스와 헤더 역추적

**Files:**
- Create: `src/utils/normalization/rateStructure.ts`
- Extend: `qa/rate-normalization.test.mjs`

**Interfaces:**
- Consumes: 대표 제비율표 `IRSheet`.
- Produces: `buildRateStructure(sheet)`, `contextForValueCell(structure, address)`.

- [ ] **Step 1: AC15와 ER19의 헤더가 섞이지 않는 실패 테스트를 작성한다**

```js
assert.equal(structure.sections.find((section) => section.title === '간접노무비')?.range, 'AC3:AK56');
assert.deepEqual(contextForValueCell(structure, 'AC15').labels, ['간접노무비', '(직노) x 율', '건축', '10억 미만', '6개월 이하 (183일)']);
assert.equal(contextForValueCell(structure, 'ER19').labels.includes('비계·구조물해체공사'), false);
assert.equal(contextForValueCell(structure, 'ER19').labels.includes('건축공사'), true);
assert.equal(contextForValueCell(structure, 'ER19').labels.includes('5억 미만'), true);
```

- [ ] **Step 2: 테스트가 구조 함수 누락으로 실패하는지 확인한다**

Run: `node --test qa/rate-normalization.test.mjs --test-name-pattern="header context"`

Expected: module/export missing failure.

- [ ] **Step 3: 병합 인덱스를 이용한 구조 분석기를 구현한다**

대분류는 대괄호 제목 병합셀과 다음 대분류 시작 열을 기준으로 나눈다. 표 상단 대분류와 47행 이후 새 대분류가 같은 열 블록에 다시 등장할 수 있으므로 행 범위도 함께 사용한다. 값 셀의 조건은 현재 섹션 내부에서 위쪽 병합 헤더와 왼쪽 행 헤더만 수집하고, 다른 섹션 열의 텍스트는 절대 포함하지 않는다.

```ts
export interface RateSection { title: string; range: string; anchor: IRCell }
export interface RateCellContext { section: RateSection; labels: string[]; evidence: CellEvidence[] }
export interface RateStructure { sheet: IRSheet; sections: RateSection[]; mergeIndex: Map<string, MergeHit> }

export function buildRateStructure(sheet: IRSheet): RateStructure;
export function contextForValueCell(structure: RateStructure, address: string): RateCellContext;
```

- [ ] **Step 4: 헤더 문맥 테스트를 통과시킨다**

Run: `node --test qa/rate-normalization.test.mjs --test-name-pattern="header context"`

Expected: AC15 and ER19 assertions pass.

- [ ] **Step 5: 커밋한다**

```bash
git add src/utils/normalization/rateStructure.ts qa/rate-normalization.test.mjs
git commit -m "feat: index merged rate table headers"
```

---

### Task 7: 제비율 전체 섹션 정규화, 규정 연결, 커버리지

**Files:**
- Create: `src/utils/normalization/rateWorkbook.ts`
- Modify: `src/utils/criteria.ts`
- Extend: `qa/rate-normalization.test.mjs`

**Interfaces:**
- Consumes: `WorkbookIR`, `FileMeta`, Task 5/6 파서.
- Produces: `normalizeRateWorkbook(ir, file): NormalizedDocument<NormalizedRateRule>`.

- [ ] **Step 1: 대표 셀, 복합식, 오류, 규정, 커버리지 테스트를 먼저 작성한다**

```js
const byCell = new Map(doc.rows.flatMap((row) => row.evidence.map((evidence) => [evidence.cell, row])));
assert.equal(byCell.get('AC15').rate, 0.175);
assert.equal(byCell.get('ER19').rate, 0.0311);
assert.equal(byCell.get('P101').valueKind, 'FORMULA');
assert.match(byCell.get('P101').formula, /79만원/);
assert.equal(byCell.get('AH81').valueKind, 'SOURCE_ERROR');
assert.ok(byCell.get('AH81').status === 'SOURCE_ERROR');
assert.ok(doc.rows.some((row) => row.regulations.some((regulation) => regulation.ministry === '고용노동부')));
assert.ok(doc.coverage.nonEmptySourceCells > 0);
assert.equal(doc.coverage.classifiedSourceCells + doc.coverage.unclassifiedSourceCells, doc.coverage.nonEmptySourceCells);
```

섹션 존재 검증은 다음 표준항목 목록 전체를 사용한다.

```js
const required = ['간접노무비', '기타경비', '일반관리비', '이윤', '건설기계대여대금 지급보증액 발급금액', '산업안전보건관리비', '건강보험료', '노인장기요양보험료', '연금보험료', '산재보험료', '고용보험료', '환경보전비', '건설하도급대금지급보증서발급수수료', '공사이행보증수수료', '법정부담금', '석면분담금', '임금채권부담금', '퇴직공제부금비'];
for (const name of required) assert.ok(doc.rows.some((row) => row.canonicalName === name), name);
```

- [ ] **Step 2: 테스트가 normalizer 누락 또는 불완전으로 실패하는지 확인한다**

Run: `node --test qa/rate-normalization.test.mjs`

Expected: representative rule assertions fail before implementation.

- [ ] **Step 3: 섹션별 규칙 생성기를 구현한다**

`SECTION_TITLES`를 위 `required` 목록과 동일하게 유지한다. 각 섹션에서 값 셀을 찾고 `contextForValueCell`의 문맥을 `RateConditions`로 변환한다. `P98:P110`은 계산식 규칙, `AH81`은 원본 오류, `ER`/`EZ` 계열은 요율과 기초액을 같은 규칙으로 결합한다. 빈 셀과 `-`는 `NOT_APPLICABLE` 행으로 보존한다. 모든 규칙의 `evidence`에는 값 셀을 첫 번째로, 조건 병합 anchor를 그 뒤에 중복 없이 넣는다.

```ts
export function normalizeRateWorkbook(
  ir: WorkbookIR,
  file: FileMeta,
): NormalizedDocument<NormalizedRateRule>;
```

- [ ] **Step 4: 관련규정 시트를 연결하고 원본 셀을 분류한다**

`관련규정 링크` 시트의 `구분·소관부처·관련규정 링크` 헤더를 찾고 빈 구분은 직전 구분을 상속한다. `matchCanonicalLabel`과 공백 제거 비교로 `RegulationReference`를 만든다. 규칙에 사용된 셀, 헤더, 주석, 규정은 classified set에 추가하고 나머지 비어 있지 않은 셀은 `UNCLASSIFIED_SOURCE_CELL` 경고로 남긴다.

- [ ] **Step 5: 기존 `parseRateWorkbook`을 호환 래퍼로 바꾼다**

새 반환값에 `normalized`를 추가하되 기존 호출을 깨지 않게 `criteria`도 제공한다. `criteria`는 `NORMAL`, `PERCENT_RATE`, 조건 충돌 없음인 규칙만 기존 `RateCriterion`으로 투영한다.

```ts
export async function parseRateWorkbook(file: File, procurementType: ProcurementType) {
  const ir = await buildWorkbookIR(file, procurementType);
  const meta = { name: file.name, sizeLabel: `${(file.size / 1024 / 1024).toFixed(1)}MB`, detail: '' };
  const normalized = normalizeRateWorkbook(ir, meta);
  const criteria = normalized.rows
    .filter((row) => row.status === 'NORMAL' && row.valueKind === 'PERCENT_RATE' && row.rate != null)
    .map((row) => ({ canonicalName: row.canonicalName, rate: row.rate, sheetName: row.evidence[0].sheetName, cell: row.evidence[0].cell, displayValue: row.evidence[0].displayValue, confidence: row.confidence, condition: row.conditions.requiredConditions.join(' / ') }));
  return { ir, normalized, criteria };
}
```

- [ ] **Step 6: 전체 제비율 회귀 테스트를 통과시킨다**

Run: `node --test qa/rate-normalization.test.mjs qa/criteria.test.mjs`

Expected: all representative cell, section, regulation, coverage, legacy assertions pass.

- [ ] **Step 7: 커밋한다**

```bash
git add src/utils/normalization/rateWorkbook.ts src/utils/criteria.ts qa/rate-normalization.test.mjs
git commit -m "feat: normalize complete rate workbook rules"
```

---

### Task 8: 공사조건과 제비율 규칙 매칭

**Files:**
- Create: `src/utils/normalization/ruleMatcher.ts`
- Create: `qa/rule-matcher.test.mjs`

**Interfaces:**
- Consumes: `ProjectContext`, `NormalizedRateRule[]`.
- Produces: `matchApplicableRateRules(context, rules): RateRuleSelection[]`.

- [ ] **Step 1: 단일 매치, 충돌, 없음, 낮은 신뢰도 테스트를 작성한다**

```js
assert.equal(matchApplicableRateRules(context, [matchingRule])[0].status, 'APPLIED');
assert.equal(matchApplicableRateRules(context, [matchingRule, duplicateRule])[0].status, 'CONFLICT');
assert.equal(matchApplicableRateRules(context, [nonMatchingRule])[0].status, 'UNAVAILABLE');
assert.equal(matchApplicableRateRules(context, [{ ...matchingRule, confidence: 0.5 }])[0].status, 'NEEDS_REVIEW');
```

- [ ] **Step 2: 테스트가 matcher 누락으로 실패하는지 확인한다**

Run: `node --test qa/rule-matcher.test.mjs`

Expected: module/export missing failure.

- [ ] **Step 3: 매칭 타입과 순수 함수를 구현한다**

```ts
export interface RateRuleSelection {
  canonicalName: string;
  status: 'APPLIED' | 'CONFLICT' | 'UNAVAILABLE' | 'NEEDS_REVIEW';
  selected: NormalizedRateRule | null;
  candidates: NormalizedRateRule[];
  reason: string;
}

export function matchApplicableRateRules(
  context: ProjectContext,
  rules: NormalizedRateRule[],
): RateRuleSelection[];
```

적용시기, 공사종류, 해당 `amountBasis`, 금액 경계, 기간 경계, 업종, 등급, 계약방식, 도급자관급, 제외조건 순으로 필터링한다. 필수 context가 없으면 해당 항목을 `NEEDS_REVIEW`로 반환한다. 같은 `canonicalName`에서 후보 1개만 남고 confidence가 0.75 이상이면 `APPLIED`다.

- [ ] **Step 4: matcher 테스트를 통과시킨다**

Run: `node --test qa/rule-matcher.test.mjs`

Expected: all matcher states pass.

- [ ] **Step 5: 커밋한다**

```bash
git add src/utils/normalization/ruleMatcher.ts qa/rule-matcher.test.mjs
git commit -m "feat: match applicable rate rules"
```

---

### Task 9: 화면용 요약·필터·차단 순수 함수

**Files:**
- Create: `src/utils/normalization/viewModel.ts`
- Create: `qa/normalization-view-model.test.mjs`

**Interfaces:**
- Consumes: 세 `NormalizedDocument`, rate filter state.
- Produces: `normalizationTabSummary`, `filterRateRules`, `hasFatalNormalizationIssue`.

- [ ] **Step 1: 탭 수치, 복합필터, 치명 오류 테스트를 작성한다**

```js
assert.deepEqual(normalizationTabSummary(rateDoc), { total: 386, normal: 370, needsReview: 15, sourceError: 1, unclassified: 4 });
assert.deepEqual(filterRateRules(rows, { query: '안전', status: 'SOURCE_ERROR', constructionType: 'ALL', valueKind: 'ALL' }).map((row) => row.ruleId), ['rate-2']);
assert.equal(hasFatalNormalizationIssue([rateDoc, laborDoc, costDoc]), false);
assert.equal(hasFatalNormalizationIssue([{ ...rateDoc, rows: [] }]), true);
```

- [ ] **Step 2: 테스트가 함수 누락으로 실패하는지 확인한다**

Run: `node --test qa/normalization-view-model.test.mjs`

Expected: module/export missing failure.

- [ ] **Step 3: 순수 함수를 구현한다**

검색 대상은 표준항목, 원본 섹션명, 조건 원문, 셀 주소다. 필터는 단일 루프에서 조합하고 입력 배열을 변경하지 않는다. 치명 오류는 `issues.some(severity === 'FATAL') || rows.length === 0`만 사용한다.

```ts
export function normalizationTabSummary<Row>(doc: NormalizedDocument<Row>): NormalizationTabSummary;
export function filterRateRules(rows: NormalizedRateRule[], filters: RateRuleFilters): NormalizedRateRule[];
export function hasFatalNormalizationIssue(docs: Array<NormalizedDocument<unknown> | null>): boolean;
```

- [ ] **Step 4: view-model 테스트를 통과시킨다**

Run: `node --test qa/normalization-view-model.test.mjs`

Expected: all view-model assertions pass.

- [ ] **Step 5: 커밋한다**

```bash
git add src/utils/normalization/viewModel.ts qa/normalization-view-model.test.mjs
git commit -m "feat: add normalization audit view model"
```

---

### Task 10: 업로드 결과를 세 정규화 문서로 보존

**Files:**
- Modify: `src/screens/UploadScreen.tsx`
- Modify: `src/App.tsx`
- Modify: `src/utils/criteria.ts`
- Create: `src/screens/NormalizationAuditScreen.tsx`

**Interfaces:**
- Consumes: 문서별 `File`, 문서별 normalizer.
- Produces: App state `costDocument`, `rateDocument`, `laborDocument`, `projectContext`.

- [ ] **Step 1: App가 `normalization` 단계와 문서 상태를 요구하도록 타입 실패를 만든다**

`App.tsx`에서 먼저 다음 Step과 상태 타입을 선언하고 아직 없는 `NormalizationAuditScreen` import를 추가한다.

```ts
type Step = 'upload' | 'normalization' | 'progress' | 'dashboard' | 'detail';
const [costDocument, setCostDocument] = useState<NormalizedDocument<NormalizedCostRow> | null>(null);
const [rateDocument, setRateDocument] = useState<NormalizedDocument<NormalizedRateRule> | null>(null);
const [laborDocument, setLaborDocument] = useState<NormalizedDocument<NormalizedLaborRate> | null>(null);
```

- [ ] **Step 2: 타입 검사를 실행해 화면과 콜백 누락을 확인한다**

Run: `pnpm exec tsc --noEmit`

Expected: missing `NormalizationAuditScreen` and incompatible upload callback errors.

- [ ] **Step 3: 업로드 콜백 계약을 정규화 문서로 변경한다**

```ts
onCostExcel: (document: NormalizedDocument<NormalizedCostRow>) => void;
onRateExcel: (document: NormalizedDocument<NormalizedRateRule>) => void;
onLaborExcel: (document: NormalizedDocument<NormalizedLaborRate>) => void;
```

원가계산서는 `buildWorkbookIR` 뒤 `normalizeCostStatement`와 `extractProjectContext`를 실행하고 `CellEvidence`를 연결한다. 제비율과 노임단가는 각 normalizer 결과를 전달한다. 업로드 실패는 빈 행과 `FATAL/MISSING_REQUIRED_VALUE` issue가 있는 문서를 만들어 화면에서 설명 가능하게 한다.

- [ ] **Step 4: 컴파일 가능한 최소 검수 화면을 만든다**

```tsx
import type { NormalizedCostRow, NormalizedDocument, NormalizedLaborRate, NormalizedRateRule, ProjectContext, ValidationMode } from '../types';
import { Button } from '../components/ui';

export interface NormalizationAuditScreenProps {
  costDocument: NormalizedDocument<NormalizedCostRow>;
  rateDocument: NormalizedDocument<NormalizedRateRule>;
  laborDocument: NormalizedDocument<NormalizedLaborRate>;
  projectContext: ProjectContext;
  mode: ValidationMode;
  onModeChange: (mode: ValidationMode) => void;
  onBack: () => void;
  onRun: () => void;
}

export function NormalizationAuditScreen({ onBack, onRun }: NormalizationAuditScreenProps) {
  return (
    <main className="app-shell">
      <section className="workspace">
        <h1>정규화 검수</h1>
        <div className="footer-actions">
          <Button onClick={onBack}>이전</Button>
          <Button onClick={onRun} variant="primary">검증 실행</Button>
        </div>
      </section>
    </main>
  );
}
```

- [ ] **Step 5: App 상태와 다음 단계 전환을 배선한다**

세 필수 문서가 선택되면 `onStart`가 `setStep('normalization')`을 실행한다. 문서 교체 시 그 문서 상태만 갱신하고 기존 검증결과는 초기화한다. `projectContext`는 `costDocument.ir`이 바뀔 때 `useMemo`로 한 번 파생한다.

- [ ] **Step 6: 타입 검사와 단위 테스트를 통과시킨다**

Run: `pnpm test`

Expected: all unit tests and TypeScript checks pass.

- [ ] **Step 7: 독립적으로 커밋한다**

```bash
git add src/App.tsx src/screens/UploadScreen.tsx src/screens/NormalizationAuditScreen.tsx src/utils/criteria.ts
git commit -m "feat: retain normalized upload documents"
```

---

### Task 11: 탭형 정규화 검수실 UI

**Files:**
- Create: `src/components/normalization/ConditionSummary.tsx`
- Create: `src/components/normalization/NormalizationTabs.tsx`
- Create: `src/components/normalization/NormalizationTable.tsx`
- Create: `src/components/normalization/EvidencePanel.tsx`
- Modify: `src/screens/NormalizationAuditScreen.tsx`
- Delete: `src/screens/RecognitionScreen.tsx`
- Modify: `src/App.tsx`
- Modify: `src/screens/UploadScreen.tsx`
- Create: `qa/normalization-screen.test.mjs`

**Interfaces:**
- Consumes: 세 정규화 문서, `ProjectContext`, 검증 모드.
- Produces: 검증 전 읽기 전용 검수 화면과 `onRun` 게이트.

- [ ] **Step 1: 탭·근거·치명 오류 버튼을 요구하는 실패 테스트를 작성한다**

```js
import assert from 'node:assert/strict';
import test from 'node:test';
import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { withViteModule } from './helpers/vite-module.mjs';

const document = (kind, rows, fatal = false) => ({
  kind,
  file: { name: `${kind}.xlsx`, sizeLabel: '0.1MB', detail: '' },
  ir: { schemaVersion: '1.0', fileName: `${kind}.xlsx`, procurementType: 'CONSTRUCTION', generatedAt: '2026-07-22T00:00:00.000Z', sheets: [], totals: { sheetCount: 0, cellCount: 0, formulaCount: 0, mergeCount: 0 } },
  rows,
  issues: fatal ? [{ issueId: 'fatal-1', code: 'NO_NORMALIZED_ROWS', severity: 'FATAL', message: '행 없음', evidence: [] }] : [],
  coverage: { nonEmptySourceCells: 1, classifiedSourceCells: 1, unclassifiedSourceCells: 0, normalizedRows: rows.length, warningCount: 0, fatalCount: fatal ? 1 : 0 },
});

test('renders three audit tabs and blocks only fatal documents', async () => {
  await withViteModule('/src/screens/NormalizationAuditScreen.tsx', ({ NormalizationAuditScreen }) => {
    const markup = renderToStaticMarkup(React.createElement(NormalizationAuditScreen, {
      costDocument: document('COST_STATEMENT', [{ rowId: 'cost-1', evidence: [] }]),
      rateDocument: document('RATE_TABLE', [], true),
      laborDocument: document('LABOR_TABLE', [{ rowId: 'labor-1', evidence: [] }]),
      projectContext: { constructionType: null, amount: null, durationDays: null, bidDate: null, alternatives: [], issues: [] },
      mode: 'ARITHMETIC_AND_RATE',
      onModeChange() {}, onBack() {}, onRun() {},
    }));
    assert.match(markup, /원가계산서/);
    assert.match(markup, /제비율표/);
    assert.match(markup, /노임단가표/);
    assert.match(markup, /disabled/);
  });
});
```

- [ ] **Step 2: 테스트가 최소 화면에서 실패하는지 확인한다**

Run: `node --test qa/normalization-screen.test.mjs`

Expected: missing tab labels assertion failure.

- [ ] **Step 3: 화면을 A안 구조로 구현한다**

`NormalizationAuditScreen`은 active tab, selected row id, 각 필터만 소유한다. 탭 요약과 필터 결과는 Task 9 함수에 `useMemo`를 사용한다. 순서는 공사조건 카드, 탭, 필터, 표+근거 2열, 검증 설정, 하단 버튼이다. 검증 버튼은 `hasFatalNormalizationIssue`가 true일 때만 disabled다.

```tsx
const documents = [costDocument, rateDocument, laborDocument];
const blocked = hasFatalNormalizationIssue(documents);
const visibleRateRows = useMemo(
  () => filterRateRules(rateDocument.rows, rateFilters),
  [rateDocument.rows, rateFilters],
);
```

- [ ] **Step 4: 공사조건 카드와 부족한 값 입력을 구현한다**

자동 추출값은 값·신뢰도·`sheet!cell`을 표시한다. 값이 null이거나 `CONDITION_CONFLICT`일 때만 텍스트/숫자 입력을 노출하고, App가 세션용 context override를 상태로 보관한다. 읽기 전용 정규화 행은 변경하지 않는다.

- [ ] **Step 5: 세 문서 표와 근거 패널을 구현한다**

원가계산서 표는 항목·산출기초·요율·금액·계산셀, 제비율표는 항목·공사종류·금액·기간·값유형·요율/식·상태, 노임단가표는 코드·직종·단가·공표상태를 표시한다. 행 선택 시 `EvidencePanel`에 모든 `CellEvidence`, notes, regulations, issue를 표시한다. 표는 읽기 전용이고 입력 요소를 포함하지 않는다.

- [ ] **Step 6: App에서 이전 화면과 검증 실행을 연결한다**

`RecognitionScreen` import와 렌더 분기를 제거하고 `NormalizationAuditScreen`을 사용한다. `onBack`은 upload, `onRun`은 기존 `handleRun`을 호출한다. 검증 모드 선택은 화면 하단에 유지한다.

- [ ] **Step 7: SSR 화면 테스트와 전체 테스트를 실행한다**

Run: `node --test qa/normalization-screen.test.mjs`

Expected: screen SSR test passes.

Run: `pnpm test`

Expected: TypeScript and all unit tests pass.

- [ ] **Step 8: UI를 커밋한다**

```bash
git add src/App.tsx src/screens/UploadScreen.tsx src/screens/NormalizationAuditScreen.tsx src/components/normalization src/utils/criteria.ts
git rm src/screens/RecognitionScreen.tsx
git commit -m "feat: add pre-validation normalization audit"
```

---

### Task 12: 검증 엔진에 선택된 정규화 규칙 연결

**Files:**
- Modify: `src/App.tsx`
- Modify: `src/utils/validation.ts`
- Extend: `qa/rule-matcher.test.mjs`

**Interfaces:**
- Consumes: `RateRuleSelection[]`, `NormalizedLaborRate[]`.
- Produces: 기존 `ValidationConfig.referenceRates`, `laborRates`와 동일한 검증 입력.

- [ ] **Step 1: 선택 규칙만 reference rate로 투영하는 실패 테스트를 작성한다**

```js
const references = selectionsToReferenceRates([
  { canonicalName: '간접노무비', status: 'APPLIED', selected: matchingRule, candidates: [matchingRule], reason: '조건 일치' },
  { canonicalName: '기타경비', status: 'CONFLICT', selected: null, candidates: [ruleA, ruleB], reason: '2개 조건 일치' },
]);
assert.equal(references['간접노무비'].rate, matchingRule.rate);
assert.equal(references['기타경비'], undefined);
```

- [ ] **Step 2: 테스트 실패를 확인한다**

Run: `node --test qa/rule-matcher.test.mjs --test-name-pattern="reference rates"`

Expected: `selectionsToReferenceRates is not a function`.

- [ ] **Step 3: 투영 함수와 검증 결과 상태를 구현한다**

```ts
export function selectionsToReferenceRates(
  selections: RateRuleSelection[],
): Record<string, ReferenceRate>;
```

`APPLIED`만 셀 근거와 함께 `ReferenceRate`로 만든다. `CONFLICT`, `UNAVAILABLE`, `NEEDS_REVIEW`는 자동 기준요율에서 제외하고 `runValidation`이 해당 항목을 `NEEDS_REVIEW` 또는 `UNAVAILABLE`로 반환하도록 선택 상태를 config에 전달한다.

- [ ] **Step 4: App의 수기 rateInputs 우선경로를 제거한다**

`handleRun`은 `matchApplicableRateRules(effectiveProjectContext, rateDocument.rows)` 결과를 사용한다. 기존 `rateInputs` 상태와 정규화 화면의 수기 요율 입력은 삭제한다. 노임단가는 `publicationStatus === 'PUBLISHED'`인 행만 기존 `LaborRate`로 투영하고 미공표 직종은 검증 불가 근거로 유지한다.

- [ ] **Step 5: matcher·validation 회귀 테스트와 전체 테스트를 통과시킨다**

Run: `pnpm test`

Expected: all tests pass.

- [ ] **Step 6: 커밋한다**

```bash
git add src/App.tsx src/utils/validation.ts src/utils/normalization/ruleMatcher.ts qa/rule-matcher.test.mjs
git commit -m "feat: validate with matched normalized rate rules"
```

---

### Task 13: 검수실 스타일과 반응형 레이아웃

**Files:**
- Modify: `src/styles.css`

**Interfaces:**
- Consumes: Task 11 class names.
- Produces: 데스크톱 2열 검수실, 모바일 단일열과 표 가로 스크롤.

- [ ] **Step 1: 스타일 없이 앱을 빌드하고 브라우저에서 구조를 확인한다**

Run: `pnpm run build`

Expected: build passes; normalization markup renders with unstyled new classes.

- [ ] **Step 2: 다음 레이아웃 계약을 CSS로 구현한다**

```css
.normalization-tabs { display: flex; gap: 8px; overflow-x: auto; }
.normalization-tab { min-width: 180px; }
.normalization-audit-grid { display: grid; grid-template-columns: minmax(0, 1fr) 340px; gap: 18px; align-items: start; }
.normalization-table-scroll { overflow: auto; max-height: 620px; }
.normalization-table { width: 100%; min-width: 1080px; border-collapse: separate; border-spacing: 0; }
.normalization-table thead { position: sticky; top: 0; z-index: 2; }
.normalization-evidence { position: sticky; top: 18px; max-height: calc(100vh - 36px); overflow: auto; }
.normalization-row-source-error { background: #fff1f2; }
.normalization-row-needs-review { background: #fff7ed; }

@media (max-width: 980px) {
  .normalization-audit-grid { grid-template-columns: 1fr; }
  .normalization-evidence { position: static; max-height: none; }
}
```

기존 색상·타이포그래피·Panel·Button 스타일을 재사용하고 상태를 색상과 텍스트로 함께 표시한다.

- [ ] **Step 3: 데스크톱과 모바일 브라우저에서 레이아웃을 확인한다**

`browser:control-in-app-browser`로 실행 중인 로컬 앱을 먼저 검증한다. 해당 도구를 사용할 수 없을 때만 그 사유를 기록하고 일반 Playwright로 대체한다.

브라우저 검증 대상:

- 1440×900: 표와 340px 근거 패널이 겹치지 않는다.
- 390×844: 탭 가로 스크롤, 표 가로 스크롤, 근거 패널 단일열이 작동한다.
- 긴 규정 URL과 병합범위 텍스트가 컨테이너를 넘지 않는다.

- [ ] **Step 4: 커밋한다**

```bash
git add src/styles.css
git commit -m "style: polish normalization audit workspace"
```

---

### Task 14: 전체 검증, 브라우저 QA, 문서 정리

**Files:**
- Modify when evidence requires: files changed in Tasks 1–13 only.
- Verify: `docs/superpowers/specs/2026-07-22-normalization-audit-design.md` acceptance criteria.

**Interfaces:**
- Consumes: completed implementation.
- Produces: test/build/browser evidence and clean diff.

- [ ] **Step 1: 전체 단위 테스트와 타입 검사를 새로 실행한다**

Run: `pnpm test`

Expected: every listed Node test passes and `tsc --noEmit` exits 0.

- [ ] **Step 2: 프로덕션 빌드를 실행한다**

Run: `pnpm run build`

Expected: Vite exits 0. Existing 500kB chunk warning is allowed; new error or warning is not.

- [ ] **Step 3: 실제 대표 파일로 전체 사용자 흐름을 브라우저에서 검증한다**

`browser:control-in-app-browser`를 우선 사용하고, 사용할 수 없을 때만 그 사유를 기록한 뒤 일반 Playwright를 사용한다.

검증 순서:

1. `cost-reference.xlsx` 업로드.
2. `rate-reference.xlsx`, `labor-reference.xlsx` 업로드.
3. 정규화 검수실 진입.
4. 세 탭 행·상태 배지 확인.
5. 제비율표에서 `산업안전보건관리비` 검색.
6. `ER19` 행 선택 후 공사종류·5억 미만·3.11%·원본 셀 근거 확인.
7. `AH81` 원본 오류가 표시되는지 확인.
8. 노임단가표에서 `연마공` 검색 후 미공표 상태 확인.
9. 경고만 있는 상태에서 검증 실행 가능 확인.
10. 필수 문서 행을 0개로 만든 테스트 상태에서 버튼 차단 확인.

- [ ] **Step 4: 원본 커버리지와 git diff를 확인한다**

Run: `git diff --check`

Expected: no whitespace errors.

Run: `git status --short`

Expected: only intended feature files and pre-existing user changes are listed.

- [ ] **Step 5: 결함이 발견되면 해당 Task의 실패 테스트로 재현하고 그 Task 파일만 수정·재검증·커밋한다**

최종 QA 자체는 파일을 바꾸지 않는다. 결함 수정 커밋에는 재현 테스트와 직접 수정한 production 파일만 포함하며, 기존 사용자 변경이나 관계없는 파일을 스테이징하지 않는다.

---

## Self-Review

- 설계 §3 검증 전 흐름 → Tasks 10–11.
- 설계 §4 공통 계약·커버리지 → Task 2.
- 설계 §5 전체 시트 공사조건 추출 → Task 3.
- 설계 §6 전체 제비율 규칙·조건 축·관련규정·적용 선택 → Tasks 5–8.
- 설계 §7 노임단가 미공표·표식 보존 → Task 4.
- 설계 §8 탭형 A안 UI와 근거 패널 → Tasks 9–11, 13.
- 설계 §9 치명 오류만 차단 → Tasks 9–11.
- 설계 §10 향후 보정 레이어 → 원본과 세션 override 분리를 Task 11에서 유지하며 수정 UI는 제외.
- 설계 §11 성능·React 원칙 → Tasks 10–11.
- 설계 §12 테스트 → Tasks 1–9, 12–14.
- 설계 §13 수용 기준 → Task 14 사용자 흐름으로 전부 확인.
- 타입 이름은 설계와 Tasks 2–12에서 일치한다.
- 계획에 미정 구현, 빈 함수, 후속 결정을 요구하는 단계가 없다.
