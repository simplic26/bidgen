# Cost Parser Hardening Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Improve Excel cost-estimate parsing so uploaded workbooks are classified by sheet content, scanned without missing late important sheets, and validated without false formula errors from unsupported Excel functions.

**Architecture:** Keep the existing `WorkbookIR` pipeline as the stable ingestion layer, then add focused parser modules around it. `sheetRecognition` assigns roles and confidence from sheet names plus cell contents, `scanPolicy` decides which sheets/ranges must be scanned fully, and `formulaEvaluator` reports supported and unsupported formulas explicitly so validation can distinguish real arithmetic mismatches from formulas it cannot recalculate.

**Tech Stack:** React, TypeScript, Vite, SheetJS `xlsx`, Node 24 built-in test runner for parser unit tests.

## Global Constraints

- Do not add a backend for this improvement.
- Continue accepting `.xlsx` and `.xls` through the browser upload flow.
- Keep `WorkbookIR` compatible with the existing app screens.
- Do not treat unsupported Excel formulas as arithmetic errors.
- Large files such as 33-sheet 용역 설계내역서 must not skip `1.원가계산서` or `첨부7-*` sheets because of a global cell limit.
- Parser tests must use synthetic in-repo fixtures; manual reference-file checks may use the external sample folder.

---

## File Structure

- Modify `package.json`: add a parser unit test script using Node's built-in test runner.
- Modify `src/types.ts`: add sheet recognition and formula evaluation metadata types.
- Create `src/utils/parserTestFixtures.ts`: small helpers for building synthetic `WorkbookIR` fixtures in tests.
- Create `src/utils/sheetRecognition.ts`: content-based sheet role scoring and confidence reasons.
- Create `src/utils/sheetRecognition.test.ts`: tests for ambiguous sheet names and Korean cost-estimate layouts.
- Modify `src/utils/excel.ts`: replace sheet-name-only classification with content-aware classification while preserving the existing IR shape.
- Create `src/utils/scanPolicy.ts`: role-aware per-sheet scan planning.
- Create `src/utils/scanPolicy.test.ts`: tests proving late important sheets are still scanned.
- Modify `src/utils/validation.ts`: use scan plans and formula evaluation outcomes.
- Create `src/utils/formulaEvaluator.ts`: supported formula tokenizer/evaluator and unsupported-formula reporting.
- Create `src/utils/formulaEvaluator.test.ts`: tests for supported formulas, unsupported formulas, and cached-value behavior.
- Create `scripts/inspect-reference-workbooks.mjs`: optional local probe for the nine external reference files.

## Task 1: Parser Unit Test Harness

**Files:**
- Modify: `package.json`
- Create: `src/utils/parserTestFixtures.ts`

**Interfaces:**
- Produces: `makeCell(address, value, options): IRCell`
- Produces: `makeSheet(sheetName, rows): IRSheet`
- Produces: `makeWorkbook(fileName, sheets): WorkbookIR`

- [ ] **Step 1: Add the parser unit test script**

Modify `package.json` scripts:

```json
{
  "scripts": {
    "dev": "tsc --noEmit && vite build --configLoader native && vite preview --configLoader native --host 127.0.0.1 --port 4173",
    "build": "tsc --noEmit && vite build --configLoader native",
    "test": "tsc --noEmit",
    "test:parser": "tsc --noEmit && node --test \"src/utils/*.test.ts\"",
    "preview": "vite preview --configLoader native --host 127.0.0.1 --port 4173"
  }
}
```

- [ ] **Step 2: Create synthetic fixture helpers**

Create `src/utils/parserTestFixtures.ts`:

```ts
import type { IRCell, IRCellType, IRSheet, SheetRole, WorkbookIR } from '../types';

interface CellOptions {
  type?: IRCellType;
  formula?: string;
  numberFormat?: string | null;
}

export function makeCell(address: string, value: string | number | boolean | null, options: CellOptions = {}): IRCell {
  const dataType: IRCellType =
    options.type ?? (options.formula ? 'FORMULA' : typeof value === 'number' ? 'NUMBER' : value == null ? 'BLANK' : 'STRING');

  return {
    address,
    dataType,
    rawValue: options.formula ?? (value == null ? '' : String(value)),
    cachedValue: value,
    displayValue: value == null ? '' : String(value),
    numberFormat: options.numberFormat ?? null,
    mergedRange: null,
    hidden: false,
    references: [],
  };
}

export function makeSheet(sheetName: string, cells: IRCell[], role: SheetRole = 'OTHER'): IRSheet {
  return {
    sheetName,
    sheetRole: role,
    rowCount: 100,
    columnCount: 20,
    cellCount: cells.length,
    formulaCount: cells.filter((cell) => cell.dataType === 'FORMULA').length,
    mergeCount: 0,
    cells,
  };
}

export function makeWorkbook(fileName: string, sheets: IRSheet[]): WorkbookIR {
  return {
    schemaVersion: '1.0',
    fileName,
    procurementType: 'CONSTRUCTION',
    generatedAt: '2026-07-21T00:00:00.000Z',
    sheets,
    totals: {
      sheetCount: sheets.length,
      cellCount: sheets.reduce((sum, sheet) => sum + sheet.cellCount, 0),
      formulaCount: sheets.reduce((sum, sheet) => sum + sheet.formulaCount, 0),
      mergeCount: sheets.reduce((sum, sheet) => sum + sheet.mergeCount, 0),
    },
  };
}
```

- [ ] **Step 3: Verify the harness compiles**

Run: `pnpm test`

Expected: TypeScript passes with no parser implementation changes yet.

- [ ] **Step 4: Commit**

Run:

```bash
git add package.json src/utils/parserTestFixtures.ts
git commit -m "test: add parser unit test harness"
```

## Task 2: Content-Based Sheet Recognition

**Files:**
- Modify: `src/types.ts`
- Create: `src/utils/sheetRecognition.ts`
- Create: `src/utils/sheetRecognition.test.ts`
- Modify: `src/utils/excel.ts`

**Interfaces:**
- Produces: `recognizeSheetRole(sheet): SheetRecognition`
- Produces: `classifySheetRoleFromContent(sheetName, cells): SheetRole`
- Consumes: `IRSheet`, `IRCell`, `SheetRole`

- [ ] **Step 1: Add recognition metadata types**

Add to `src/types.ts`:

```ts
export interface SheetRoleScore {
  role: SheetRole;
  score: number;
  reasons: string[];
}

export interface SheetRecognition {
  role: SheetRole;
  confidence: 'HIGH' | 'MEDIUM' | 'LOW';
  scores: SheetRoleScore[];
}
```

- [ ] **Step 2: Write failing recognition tests**

Create `src/utils/sheetRecognition.test.ts`:

```ts
import test from 'node:test';
import assert from 'node:assert/strict';
import { makeCell, makeSheet } from './parserTestFixtures.ts';
import { recognizeSheetRole } from './sheetRecognition.ts';

test('recognizes a cost summary from content even when the sheet name is generic', () => {
  const sheet = makeSheet('Sheet1', [
    makeCell('A1', '공   사   원   가   계   산   서'),
    makeCell('A3', '비목'),
    makeCell('C4', '직접재료비'),
    makeCell('C8', '직접노무비'),
    makeCell('C13', '산재보험료'),
    makeCell('C20', '일반관리비'),
    makeCell('C21', '이윤'),
    makeCell('C25', '부가가치세'),
  ]);

  const result = recognizeSheetRole(sheet);

  assert.equal(result.role, 'COST_SUMMARY');
  assert.equal(result.confidence, 'HIGH');
});

test('recognizes rate standard sheets from attachment names and body terms', () => {
  const sheet = makeSheet('첨부7-3.산업안전보건관리비', [
    makeCell('A1', '산업안전보건관리비 적용기준'),
    makeCell('A5', '대상액'),
    makeCell('B5', '요율'),
    makeCell('C5', '금액구간'),
  ]);

  assert.equal(recognizeSheetRole(sheet).role, 'RATE_STANDARD');
});

test('recognizes item detail sheets from repeated table headers', () => {
  const sheet = makeSheet('2-1.설계(평일주간)', [
    makeCell('A4', '명칭'),
    makeCell('B4', '규격'),
    makeCell('C4', '수량'),
    makeCell('D4', '단위'),
    makeCell('E4', '재료비'),
    makeCell('G4', '노무비'),
    makeCell('K4', '합계'),
  ]);

  assert.equal(recognizeSheetRole(sheet).role, 'CONSTRUCTION_ITEMS');
});
```

- [ ] **Step 3: Implement the recognition scorer**

Create `src/utils/sheetRecognition.ts`:

```ts
import type { IRCell, IRSheet, SheetRecognition, SheetRole, SheetRoleScore } from '../types';

const RULES: Array<{ role: SheetRole; strong: string[]; medium: string[]; weak: string[] }> = [
  {
    role: 'COVER_SUMMARY',
    strong: ['설계서표지', '표지', '목차'],
    medium: ['공사명', '계약명', '추정가격내역서'],
    weak: ['부가가치세포함'],
  },
  {
    role: 'COST_SUMMARY',
    strong: ['원가계산서', '공사원가계산서', '용역원가계산서', '추정가격내역서'],
    medium: ['직접재료비', '직접노무비', '간접노무비', '산재보험료', '일반관리비', '이윤', '부가가치세'],
    weak: ['비목', '구성내역', '적용율', '금액'],
  },
  {
    role: 'CONSTRUCTION_ITEMS',
    strong: ['설계내역서', '공종별내역서', '내역서종합'],
    medium: ['품명', '명칭', '규격', '수량', '단위', '단가', '금액', '합계'],
    weak: ['비고', '공종코드'],
  },
  {
    role: 'UNIT_PRICE',
    strong: ['일위대가', '단가산출서', '일위대가목록'],
    medium: ['재료비', '노무비', '경비', '단가', '금액', '산출근거'],
    weak: ['공량', '품셈'],
  },
  {
    role: 'QUANTITY',
    strong: ['수량산출서', '물량산출서'],
    medium: ['수량', '산식', '합계', '구분'],
    weak: ['소계'],
  },
  {
    role: 'WAGE_RATE',
    strong: ['노임단가', '임금', '직종'],
    medium: ['직종명', '단가', '기준'],
    weak: ['상반기', '하반기'],
  },
  {
    role: 'RATE_STANDARD',
    strong: ['제비율', '요율', '기타경비요율', '산업안전보건관리비'],
    medium: ['보험료', '일반관리비', '이윤', '적용기준', '금액구간'],
    weak: ['대상액', '공사규모', '공사기간'],
  },
  {
    role: 'PRICE_SURVEY',
    strong: ['가격조사', '견적서', '단가대비표', '자재단가표'],
    medium: ['업체', '견적', '적용단가', '물가자료'],
    weak: ['비교'],
  },
];

function normalize(value: string): string {
  return value.replace(/\s+/g, '').toLowerCase();
}

function textFromCells(cells: IRCell[], limit = 160): string {
  return cells
    .filter((cell) => cell.dataType === 'STRING' || cell.dataType === 'FORMULA')
    .slice(0, limit)
    .map((cell) => cell.displayValue || cell.rawValue)
    .join('\n');
}

function scoreRule(haystack: string, role: SheetRole, strong: string[], medium: string[], weak: string[]): SheetRoleScore {
  let score = 0;
  const reasons: string[] = [];
  for (const keyword of strong) {
    if (haystack.includes(normalize(keyword))) {
      score += 12;
      reasons.push(keyword);
    }
  }
  for (const keyword of medium) {
    if (haystack.includes(normalize(keyword))) {
      score += 5;
      reasons.push(keyword);
    }
  }
  for (const keyword of weak) {
    if (haystack.includes(normalize(keyword))) {
      score += 2;
      reasons.push(keyword);
    }
  }
  return { role, score, reasons };
}

export function recognizeSheetRole(sheet: Pick<IRSheet, 'sheetName' | 'cells'>): SheetRecognition {
  const haystack = normalize(`${sheet.sheetName}\n${textFromCells(sheet.cells)}`);
  const scores = RULES.map((rule) => scoreRule(haystack, rule.role, rule.strong, rule.medium, rule.weak)).sort(
    (a, b) => b.score - a.score,
  );
  const best = scores[0];
  const runnerUp = scores[1];
  const role = best.score >= 10 ? best.role : 'OTHER';
  const margin = best.score - (runnerUp?.score ?? 0);
  const confidence = best.score >= 24 && margin >= 8 ? 'HIGH' : best.score >= 12 ? 'MEDIUM' : 'LOW';
  return { role, confidence, scores };
}

export function classifySheetRoleFromContent(sheetName: string, cells: IRCell[]): SheetRole {
  return recognizeSheetRole({ sheetName, cells }).role;
}
```

- [ ] **Step 4: Wire recognition into Excel IR creation**

Modify the sheet construction in `src/utils/excel.ts` so each sheet builds `cells` first, then sets:

```ts
sheetRole: classifySheetRoleFromContent(sheetName, cells),
```

Keep the existing exported `classifySheetRole(sheetName)` as a fallback helper for external callers.

- [ ] **Step 5: Run recognition tests**

Run: `pnpm test:parser`

Expected: all recognition tests pass.

- [ ] **Step 6: Commit**

Run:

```bash
git add src/types.ts src/utils/sheetRecognition.ts src/utils/sheetRecognition.test.ts src/utils/excel.ts
git commit -m "feat: classify sheets from workbook content"
```

## Task 3: Role-Aware Scan Policy For Large Workbooks

**Files:**
- Create: `src/utils/scanPolicy.ts`
- Create: `src/utils/scanPolicy.test.ts`
- Modify: `src/utils/validation.ts`

**Interfaces:**
- Produces: `buildScanPlan(ir, options): SheetScanPlan[]`
- Produces: `iterPlannedCells(ir, plan): Iterable<{ sheetName: string; cell: IRCell }>`
- Consumes: `WorkbookIR`, `IRSheet`, `IRCell`

- [ ] **Step 1: Write failing scan-policy tests**

Create `src/utils/scanPolicy.test.ts`:

```ts
import test from 'node:test';
import assert from 'node:assert/strict';
import { makeCell, makeSheet, makeWorkbook } from './parserTestFixtures.ts';
import { buildScanPlan, iterPlannedCells } from './scanPolicy.ts';

test('fully scans a late cost summary even after many earlier cells', () => {
  const dummyCells = Array.from({ length: 25000 }, (_, index) => makeCell(`A${index + 1}`, `dummy-${index}`));
  const ir = makeWorkbook('large.xlsx', [
    makeSheet('3-1.일위(평일주간)', dummyCells, 'UNIT_PRICE'),
    makeSheet(
      '1.원가계산서',
      [
        makeCell('A1', '용역 원가 계산서'),
        makeCell('C10', '산재보험료'),
        makeCell('F10', 12345),
      ],
      'COST_SUMMARY',
    ),
  ]);

  const plan = buildScanPlan(ir, { defaultCellLimit: 20000 });
  const scanned = Array.from(iterPlannedCells(ir, plan));

  assert.equal(scanned.some((entry) => entry.sheetName === '1.원가계산서' && entry.cell.address === 'C10'), true);
});

test('marks item sheets as capped while summary and standard sheets are full', () => {
  const ir = makeWorkbook('mixed.xlsx', [
    makeSheet('공종별내역서', [makeCell('A1', '품명')], 'CONSTRUCTION_ITEMS'),
    makeSheet('원가계산서', [makeCell('A1', '원가계산서')], 'COST_SUMMARY'),
    makeSheet('첨부7-1.제비율', [makeCell('A1', '제비율 적용기준')], 'RATE_STANDARD'),
  ]);

  const plan = buildScanPlan(ir, { defaultCellLimit: 20000 });

  assert.equal(plan.find((sheet) => sheet.sheetName === '공종별내역서')?.mode, 'CAPPED');
  assert.equal(plan.find((sheet) => sheet.sheetName === '원가계산서')?.mode, 'FULL');
  assert.equal(plan.find((sheet) => sheet.sheetName === '첨부7-1.제비율')?.mode, 'FULL');
});
```

- [ ] **Step 2: Implement scan policy**

Create `src/utils/scanPolicy.ts`:

```ts
import type { IRCell, SheetRole, WorkbookIR } from '../types';

export interface SheetScanPlan {
  sheetName: string;
  role: SheetRole;
  mode: 'FULL' | 'CAPPED' | 'SKIP';
  maxCells: number;
  reason: string;
}

const FULL_SCAN_ROLES = new Set<SheetRole>(['COST_SUMMARY', 'RATE_STANDARD', 'WAGE_RATE', 'PRICE_SURVEY']);
const CAPPED_SCAN_ROLES = new Set<SheetRole>(['CONSTRUCTION_ITEMS', 'UNIT_PRICE', 'QUANTITY']);

export function buildScanPlan(ir: WorkbookIR, options: { defaultCellLimit?: number } = {}): SheetScanPlan[] {
  const defaultCellLimit = options.defaultCellLimit ?? 20000;
  return ir.sheets.map((sheet) => {
    if (FULL_SCAN_ROLES.has(sheet.sheetRole)) {
      return {
        sheetName: sheet.sheetName,
        role: sheet.sheetRole,
        mode: 'FULL',
        maxCells: sheet.cellCount,
        reason: `${sheet.sheetRole} sheets contain totals, rates, or reference standards and must be scanned fully.`,
      };
    }
    if (CAPPED_SCAN_ROLES.has(sheet.sheetRole)) {
      return {
        sheetName: sheet.sheetName,
        role: sheet.sheetRole,
        mode: 'CAPPED',
        maxCells: Math.min(sheet.cellCount, defaultCellLimit),
        reason: `${sheet.sheetRole} sheets can be large; scan header and early calculation ranges first.`,
      };
    }
    return {
      sheetName: sheet.sheetName,
      role: sheet.sheetRole,
      mode: sheet.cellCount <= 5000 ? 'FULL' : 'CAPPED',
      maxCells: Math.min(sheet.cellCount, 5000),
      reason: 'Support sheets are sampled unless small.',
    };
  });
}

export function* iterPlannedCells(ir: WorkbookIR, plan: SheetScanPlan[]): Iterable<{ sheetName: string; cell: IRCell }> {
  const planBySheet = new Map(plan.map((sheetPlan) => [sheetPlan.sheetName, sheetPlan]));
  for (const sheet of ir.sheets) {
    const sheetPlan = planBySheet.get(sheet.sheetName);
    if (!sheetPlan || sheetPlan.mode === 'SKIP') continue;
    let count = 0;
    for (const cell of sheet.cells) {
      if (count >= sheetPlan.maxCells) break;
      count += 1;
      yield { sheetName: sheet.sheetName, cell };
    }
  }
}
```

- [ ] **Step 3: Replace global scan counters in validation**

Modify `src/utils/validation.ts`:

```ts
import { buildScanPlan, iterPlannedCells } from './scanPolicy';
```

Use `const scanPlan = buildScanPlan(ir, { defaultCellLimit: maxCells });` once in `runValidation`. Replace loops that stop after a global `scanned > maxCells` with planned iteration, and update `detectItems` to accept `scanPlan` or use `iterPlannedCells` internally.

- [ ] **Step 4: Add a scan coverage result**

In `runValidation`, add one warning only when any sheet plan is capped:

```ts
const capped = scanPlan.filter((sheet) => sheet.mode === 'CAPPED' && sheet.maxCells < (ctx.sheets.get(sheet.sheetName)?.byAddr.size ?? 0));
```

The result text should name the capped sheets and say totals/rate-standard sheets were scanned fully.

- [ ] **Step 5: Run scan-policy tests**

Run: `pnpm test:parser`

Expected: scan-policy tests pass and existing validation behavior still compiles.

- [ ] **Step 6: Commit**

Run:

```bash
git add src/utils/scanPolicy.ts src/utils/scanPolicy.test.ts src/utils/validation.ts
git commit -m "feat: scan important cost sheets in large workbooks"
```

## Task 4: Formula Evaluation Outcomes

**Files:**
- Create: `src/utils/formulaEvaluator.ts`
- Create: `src/utils/formulaEvaluator.test.ts`
- Modify: `src/utils/validation.ts`

**Interfaces:**
- Produces: `evaluateFormulaValue(rawValue, lookup): FormulaEvaluation`
- Consumes: formula string and numeric cell lookup callback.

- [ ] **Step 1: Add formula evaluation types**

Add to `src/types.ts`:

```ts
export type FormulaEvaluationStatus = 'SUPPORTED' | 'UNSUPPORTED' | 'MISSING_VALUE' | 'EVALUATION_ERROR';

export interface FormulaEvaluation {
  status: FormulaEvaluationStatus;
  value: number | null;
  reason: string;
}
```

- [ ] **Step 2: Write failing formula tests**

Create `src/utils/formulaEvaluator.test.ts`:

```ts
import test from 'node:test';
import assert from 'node:assert/strict';
import { evaluateFormulaValue } from './formulaEvaluator.ts';

test('evaluates supported arithmetic and SUM formulas', () => {
  const lookup = (addr: string) => ({ A1: 100, A2: 50, A3: 25 }[addr] ?? null);

  assert.deepEqual(evaluateFormulaValue('=SUM(A1:A3)*0.1', lookup), {
    status: 'SUPPORTED',
    value: 17.5,
    reason: 'Formula recalculated by supported evaluator.',
  });
});

test('reports unsupported formulas without producing a false mismatch', () => {
  const result = evaluateFormulaValue('=IF(A1>0,A1,0)', () => 100);

  assert.equal(result.status, 'UNSUPPORTED');
  assert.equal(result.value, null);
});

test('reports missing referenced values separately', () => {
  const result = evaluateFormulaValue('=A1*0.1', () => null);

  assert.equal(result.status, 'MISSING_VALUE');
  assert.equal(result.value, null);
});
```

- [ ] **Step 3: Move the current evaluator into the new module**

Create `src/utils/formulaEvaluator.ts` by moving the tokenizer and parser from `validation.ts`. Preserve support for:

```ts
SUM
ROUND
ROUNDDOWN
INT
+
-
*
/
parentheses
cell references
ranges inside SUM
```

Return `UNSUPPORTED` for formulas containing unsupported functions or comparison operators, and return `MISSING_VALUE` when a referenced cell is blank or non-numeric.

- [ ] **Step 4: Update validation to use explicit statuses**

Modify formula self-consistency validation:

```ts
const evaluation = evaluateFormulaValue(cell.rawValue, (addr) => ctx.lookup(addr, sheet.sheetName));
if (evaluation.status === 'UNSUPPORTED') {
  continue;
}
if (evaluation.status === 'MISSING_VALUE' || evaluation.status === 'EVALUATION_ERROR') {
  continue;
}
const value = evaluation.value;
```

Only compare cached values when `evaluation.status === 'SUPPORTED'` and `evaluation.value != null`.

- [ ] **Step 5: Run formula tests**

Run: `pnpm test:parser`

Expected: formula tests pass, and unsupported Excel functions do not become `ERROR` validation results.

- [ ] **Step 6: Commit**

Run:

```bash
git add src/types.ts src/utils/formulaEvaluator.ts src/utils/formulaEvaluator.test.ts src/utils/validation.ts
git commit -m "feat: separate unsupported formulas from arithmetic errors"
```

## Task 5: Manual Reference Workbook Probe

**Files:**
- Create: `scripts/inspect-reference-workbooks.mjs`
- Modify: `docs/superpowers/plans/2026-07-21-cost-parser-hardening.md`

**Interfaces:**
- Produces: console summary with file name, sheet roles, detected item count, capped scan sheets, and validation status counts.

- [ ] **Step 1: Create a probe script**

Create `scripts/inspect-reference-workbooks.mjs` that accepts file paths:

```js
import { createRequire } from 'node:module';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

const require = createRequire(import.meta.url);
const XLSX = require('xlsx');
const { workbookToIR } = await import(pathToFileURL(path.resolve('src/utils/excel.ts')).href);
const { detectItems, runValidation } = await import(pathToFileURL(path.resolve('src/utils/validation.ts')).href);
const { buildScanPlan } = await import(pathToFileURL(path.resolve('src/utils/scanPolicy.ts')).href);

for (const file of process.argv.slice(2)) {
  const workbook = XLSX.readFile(file, { cellNF: true, cellStyles: true, cellFormula: true, cellDates: true });
  const ir = workbookToIR(workbook, {
    fileName: path.basename(file),
    procurementType: file.includes('용역') ? 'SERVICE' : 'CONSTRUCTION',
  });
  const scanPlan = buildScanPlan(ir);
  const items = detectItems(ir, { maxCells: 250000 });
  const results = runValidation(ir, { mode: 'ARITHMETIC_ONLY', referenceRates: {}, maxCellsScanned: 20000 });
  const statuses = Object.groupBy(results, (result) => result.status);

  console.log(`\n${path.basename(file)}`);
  console.log(`  sheets=${ir.totals.sheetCount} cells=${ir.totals.cellCount} formulas=${ir.totals.formulaCount}`);
  console.log(`  roles=${ir.sheets.map((sheet) => `${sheet.sheetName}:${sheet.sheetRole}`).join(' / ')}`);
  console.log(`  detectedItems=${items.length}`);
  console.log(`  capped=${scanPlan.filter((sheet) => sheet.mode === 'CAPPED').map((sheet) => sheet.sheetName).join(', ') || '-'}`);
  console.log(`  statuses=${JSON.stringify(Object.fromEntries(Object.entries(statuses).map(([key, value]) => [key, value.length])))}`);
}
```

- [ ] **Step 2: Run the probe on the nine reference files**

Run the script with the nine file paths from `C:\Users\user\Documents\AI Pioneer\7주차\레퍼런스 엄선`.

Expected:
- All nine files read successfully.
- `1.원가계산서`, `원가계산서`, `원가계산`, or `추정가격내역서(설계)` classify as `COST_SUMMARY`.
- `첨부7-*`, `제비율 적용기준`, `기타경비요율`, and `건축제비율(8.9.)` classify as `RATE_STANDARD`.
- Large 용역 files report capped item/detail sheets but full summary/rate-standard sheets.
- Unsupported formula counts do not appear as arithmetic `ERROR` results.

- [ ] **Step 3: Record the final reference check summary**

Append a short "Reference Workbook Check" section to this plan with the date, command used, and one-line result for each workbook.

- [ ] **Step 4: Commit**

Run:

```bash
git add scripts/inspect-reference-workbooks.mjs docs/superpowers/plans/2026-07-21-cost-parser-hardening.md
git commit -m "test: add reference workbook parser probe"
```

## Self-Review Checklist

- Spec coverage: the three requested improvements are covered by Task 2, Task 3, and Task 4.
- Placeholder scan: each task names files, interfaces, test code, implementation shape, verification commands, and commit commands.
- Type consistency: `SheetRecognition`, `SheetRoleScore`, and `FormulaEvaluation` are defined before later tasks consume them.
- Risk control: tests use synthetic in-repo fixtures; external reference files are used only for manual probe verification.
