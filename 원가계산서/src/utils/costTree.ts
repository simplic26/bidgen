import type { DetectedItem, IRCell, IRSheet, WorkbookIR } from '../types';
import { getCanonicalDef, matchCanonicalLabel, matchStructuralLabel, parseKoreanNumber, parseRate } from './costItems';

function cellLabel(cell: IRCell): string {
  return typeof cell.cachedValue === 'string' ? cell.cachedValue : cell.displayValue;
}

// 원가계산서의 법정 골격 5항목 — 요약시트라면 거의 반드시 있다.
const CORE_LABELS = ['재료비', '노무비', '경비', '일반관리비', '이윤'] as const;

// 내역서(상세 수량산출표) 헤더 어휘. 한 행에 3개 이상 모이면 그 시트는 요약표가 아니라 내역서다.
const DETAIL_HEADER_WORDS = ['명칭', '품명', '규격', '수량', '단위', '단가', '공종', '적요', '비고'];

// 스코어 상한(최대 획득 가능): role 3 + core 10 + vat 3 + total 1 + detail 6 + density 2 = 25
const SCORE_DETAIL_CAP = 6; // 세부항목 풍부도 상한 — core 커버리지(10)를 넘지 못하게 묶는다
// 통과선: "법정 골격 4/5 + 검증된 부가세 앵커"(8+3=11) 수준. 골격만 갖춘 표는 통과하고,
// 라벨만 겹치는 요율 기준표(role −5, rateTableSignals −4)는 세부항목이 많아도 넘지 못한다.
const SUMMARY_SCORE_THRESHOLD = 11;

export interface SheetScore {
  sheetName: string;
  score: number;
  parts: Record<string, number>;
}

/**
 * 요약시트 스코어링. 시트명 하드코딩 없이 아래 신호만 쓴다.
 *
 * (+) `sheetRole === 'COST_SUMMARY'` — 엑셀 파서가 이미 계산해둔 역할. 이름 정규식은 약한 폴백.
 * (+) 법정 골격 5항목(재료비·노무비·경비·일반관리비·이윤) 커버리지 — 항목당 2점, 최대 10점.
 * (+) 검증된 부가세 앵커 — 라벨 부분일치가 아니라 findVatAnchor로 실제 10% 관계가 있는 금액
 *     셀까지 확인해야 인정한다. "총공사비(부가세포함) :" 같은 잡담성 문구는 점수를 못 얻는다.
 * (+) 세부 원가항목 풍부도("리치니스") — 단, 골격 5항목·부가가치세는 이미 위에서 셌으므로
 *     **제외**하고 세고 상한을 둔다(무제한 가산이면 항목 나열만 많은 요율표가 이길 수 있다).
 * (+) 라벨 밀도 — 요약시트는 내용 있는 행 대비 원가항목 행의 비율이 높다. 같은 항목을 그대로
 *     재참조만 하는 넓은 롤업표(ref-07 '총괄표')와 실제 계산이 일어나는 상세표를 가른다.
 * (−) 요율 기준표 신호(`항목별 비율`·`공사규모`·`공사기간`·`적용기준`), 역할 RATE_STANDARD/WAGE_RATE.
 * (−) 내역서 헤더 행(`명칭|규격|수량|단위` 등 3개 이상 동일 행) — 요약표가 아니라 수량산출 내역서.
 * (−) 셀 수 스케일 패널티 — 요약시트는 작다. 2000셀 단위로 누진(최대 −6).
 */
export function scoreSummarySheets(ir: WorkbookIR, ctx?: WorkbookContext): SheetScore[] {
  const { wb, indexes } = ctx ?? makeWorkbookContext(ir);
  return ir.sheets.map((sheet) => {
    const parts: Record<string, number> = {};
    if (sheet.sheetRole === 'COST_SUMMARY') parts.role = 3;
    else if (/원가|총괄|추정/.test(sheet.sheetName)) parts.role = 1; // 이름만 힌트인 경우는 약하게
    if (sheet.sheetRole === 'RATE_STANDARD' || sheet.sheetRole === 'WAGE_RATE') parts.role = (parts.role ?? 0) - 5;

    const idx = indexes.get(sheet.sheetName)!;
    const vatAnchor = findVatAnchor(idx, sheet.sheetName, wb);

    const structuralHits = new Set<string>();
    const canonicalHits = new Set<string>(); // CANONICAL_ITEMS 매칭 — 앵커 검증된 부가가치세만 포함
    const labeledRows = new Set<number>();
    let rateTableSignals = 0;
    let detailHeaderRows = 0;
    for (const [row, cells] of idx.byRow) {
      let headerWords = 0;
      for (const cell of cells) {
        if (cell.dataType !== 'STRING') continue;
        const raw = cellLabel(cell);
        const canon = matchCanonicalLabel(raw);
        if (canon) {
          if (canon !== '부가가치세' || vatAnchor) {
            canonicalHits.add(canon);
            labeledRows.add(row);
          }
        } else {
          const struct = matchStructuralLabel(raw);
          if (struct) {
            structuralHits.add(struct);
            labeledRows.add(row);
          }
        }
        if (/항목별\s*비율|적용기준|공사규모|공사기간/.test(raw)) rateTableSignals += 1;
        if (DETAIL_HEADER_WORDS.some((w) => raw.replace(/\s+/g, '') === w)) headerWords += 1;
      }
      if (headerWords >= 3) detailHeaderRows += 1;
    }

    parts.core = CORE_LABELS.filter((k) => structuralHits.has(k) || canonicalHits.has(k)).length * 2;
    if (vatAnchor) parts.vat = 3;
    if (structuralHits.has('총원가') || structuralHits.has('총계')) parts.total = 1;

    // 리치니스: 골격 5항목·부가가치세를 제외한 세부 원가항목 수 (상한 SCORE_DETAIL_CAP)
    const detailHits = [...canonicalHits].filter(
      (n) => n !== '부가가치세' && !CORE_LABELS.includes(n as (typeof CORE_LABELS)[number]),
    );
    parts.detail = Math.min(detailHits.length, SCORE_DETAIL_CAP);

    // 라벨 밀도: 내용 있는 행 중 원가항목 행의 비율. 상세 원가표는 높고, 넓은 롤업표·내역서는 낮다.
    const contentRows = idx.byRow.size || 1;
    const density = labeledRows.size / contentRows;
    if (density >= 0.3) parts.density = 2;
    else if (density >= 0.15) parts.density = 1;

    if (rateTableSignals) parts.rateTable = -Math.min(rateTableSignals, 4);
    if (detailHeaderRows) parts.detailTable = -3;
    const bulk = Math.min(Math.floor(sheet.cellCount / 2000) * 2, 6);
    if (bulk) parts.bulk = -bulk;

    const score = Object.values(parts).reduce((s, v) => s + v, 0);
    return { sheetName: sheet.sheetName, score, parts };
  });
}

export function findSummarySheet(ir: WorkbookIR, ctx?: WorkbookContext): string | null {
  let best: SheetScore | null = null;
  for (const scored of scoreSummarySheets(ir, ctx)) {
    if (!best || scored.score > best.score) best = scored;
  }
  return best && best.score >= SUMMARY_SCORE_THRESHOLD ? best.sheetName : null;
}

const RANGE_RE = /^(?:([^!]+)!)?([A-Z]{1,3})(\d+):([A-Z]{1,3})(\d+)$/;

function colToNum(col: string): number {
  let n = 0;
  for (const ch of col) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}
function numToCol(n: number): string {
  let s = '';
  while (n > 0) { s = String.fromCharCode(65 + ((n - 1) % 26)) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
function decodeAddr(addr: string): { row: number; col: number } {
  const m = addr.match(/^([A-Z]{1,3})(\d+)$/);
  if (!m) return { row: 0, col: 0 };
  return { row: Number(m[2]), col: colToNum(m[1]) };
}

/** 참조 정규화: '$'·따옴표 제거, 세로/가로 범위는 개별 셀로 전개(500셀 상한) */
export function expandReference(ref: string): string[] {
  const clean = ref.replace(/\$/g, '').replace(/'/g, '');
  const m = clean.match(RANGE_RE);
  if (!m) return [clean];
  const [, sheet, c1, r1s, c2, r2s] = m;
  const r1 = Number(r1s), r2 = Number(r2s);
  const out: string[] = [];
  for (let c = colToNum(c1); c <= colToNum(c2); c++) {
    for (let r = r1; r <= r2 && out.length < 500; r++) {
      out.push(`${sheet ? sheet + '!' : ''}${numToCol(c)}${r}`);
    }
  }
  return out;
}

export type RoundingKind = 'INT' | 'ROUND' | 'ROUNDDOWN' | 'ROUNDUP' | 'TRUNC' | null;

export interface FormulaShape {
  refs: string[];
  rateRefs: string[];
  baseRefs: string[];
  rounding: RoundingKind;
  roundDigits: number | null;
  literalRate: number | null;
}

// 수식 본문에서 범위 표현(A1:A10, 시트간 포함)을 다시 찾기 위한 패턴.
// IR의 references 배열은 REFERENCE_PATTERN이 콜론을 버리고 추출한 결과라 SUM(D13:D25) 같은
// 범위 수식에서도 양끝 셀(D13, D25)만 남아있다 — 이 패턴으로 본문을 재스캔해 원래 범위를 복원한다.
const RANGE_IN_FORMULA_RE = /(?:(?:'[^']+'|[A-Za-z0-9_가-힣]+)!)?\$?[A-Z]{1,3}\$?[0-9]+:\$?[A-Z]{1,3}\$?[0-9]+/g;

export function decomposeFormula(
  formula: string,
  references: string[],
  isRateRef: (ref: string) => boolean,
): FormulaShape {
  const body = formula.replace(/^=/, '');

  // 범위 복원: 본문에서 찾은 각 범위를 전개하고, references에 남아있는 그 범위의 양끝 bare
  // 참조(예: D13, D25)는 제거한다 — 전개된 범위로 대체되므로 중복 카운트를 막는다.
  const rangeMatches = body.match(RANGE_IN_FORMULA_RE) ?? [];
  let refList = [...references];
  const expandedRangeRefs: string[] = [];
  for (const rangeText of rangeMatches) {
    const clean = rangeText.replace(/\$/g, '').replace(/'/g, '');
    const m = clean.match(RANGE_RE);
    if (!m) continue;
    const [, sheet, c1, r1s, c2, r2s] = m;
    const startAddr = `${sheet ? sheet + '!' : ''}${c1}${r1s}`;
    const endAddr = `${sheet ? sheet + '!' : ''}${c2}${r2s}`;
    refList = refList.filter((r) => {
      const rc = r.replace(/\$/g, '').replace(/'/g, '');
      return rc !== startAddr && rc !== endAddr && rc !== clean;
    });
    expandedRangeRefs.push(...expandReference(rangeText));
  }
  const rawRefs = Array.from(new Set([...refList, ...expandedRangeRefs]));
  const refs = Array.from(new Set(rawRefs.flatMap(expandReference)));
  const rateRefs: string[] = [];
  const baseRefs: string[] = [];
  for (const ref of refs) (isRateRef(ref) ? rateRefs : baseRefs).push(ref);

  const roundMatch = body.match(/\b(INT|ROUNDDOWN|ROUNDUP|ROUND|TRUNC)\s*\(/i);
  const rounding = (roundMatch ? roundMatch[1].toUpperCase() : null) as RoundingKind;
  let roundDigits: number | null = null;
  if (rounding && rounding !== 'INT') {
    const digits = body.match(/,\s*(-?\d+)\s*\)/);
    if (digits) roundDigits = Number(digits[1]);
  }

  // 요율 리터럴: 셀참조·함수명 제거 후 남은 0<x<1 숫자 또는 N% (부가세 =X*0.1 대응)
  let literalRate: number | null = null;
  const stripped = body
    .replace(/'[^']*'!/g, '').replace(/[A-Za-z가-힣_]+!/g, '')
    .replace(/\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?/g, '')
    .replace(/\b[A-Z]+\b/gi, '');
  const pct = stripped.match(/(\d+(?:\.\d+)?)\s*%/);
  if (pct) literalRate = Number(pct[1]) / 100;
  else {
    const frac = stripped.match(/0\.\d+/);
    if (frac && Number(frac[0]) > 0 && Number(frac[0]) < 1) literalRate = Number(frac[0]);
  }

  return { refs, rateRefs, baseRefs, rounding, roundDigits, literalRate };
}

export interface NumCell { address: string; row: number; col: number; value: number }

/** 절사 모드별 일치 판정. expected = 원시 계산값(합·곱), actual = 시트에 적힌 값 */
export function matchesWithRounding(expected: number, actual: number): RoundingKind | 'EXACT' | null {
  if (expected === actual) return 'EXACT';
  const diff = expected - actual; // INT/ROUNDDOWN은 원시값 ≥ 절사값
  if (diff >= 0 && diff < 1) return 'INT';
  if (Math.abs(diff) <= 0.5) return 'ROUND';
  // 천단위 절사(ROUNDDOWN(x,-3) 및 ROUND(x,-3)). 창을 실제 함수의 정의역에 맞춘다:
  //  · 아래 방향은 0 이상 1000 미만(절사는 원시값이 절사값보다 크거나 같다),
  //  · 위 방향은 ROUND(x,-3)의 올림 몫만큼(−500까지)만 허용.
  //  · actual === 0 은 제외 — 0은 어떤 값과도 %1000===0을 만족해서, 금액 0 행에 아무 요율 간선이나
  //    붙어버린다(0 = base × rate 라는 거짓 근거). 금액 0 행은 관계 없음으로 두는 게 맞다.
  if (diff >= -500 && diff < 1000 && actual % 1000 === 0 && actual !== 0) return 'ROUNDDOWN';
  return null;
}

/** 두 금액이 "같은 값"인가 — 절사(INT/ROUND/천단위) 오차를 허용한다. null은 엄격 비교. */
export function isSameAmount(a: number | null, b: number | null): boolean {
  if (a == null || b == null) return a === b;
  return Boolean(matchesWithRounding(a, b));
}

/** target에 가까운 쪽부터 누적하는 연속 구간(2셀 이상) 합 탐색. 위쪽 방향 우선, 실패 시 아래쪽 */
export function inferSumRange(target: NumCell, column: NumCell[]): NumCell[] | null {
  const tryDir = (cells: NumCell[]): NumCell[] | null => {
    let sum = 0;
    const picked: NumCell[] = [];
    for (const c of cells) {
      picked.push(c);
      sum += c.value;
      if (picked.length >= 2 && matchesWithRounding(sum, target.value)) return [...picked];
    }
    return null;
  };
  const others = column.filter((c) => c.address !== target.address);
  const above = others.filter((c) => c.row < target.row).sort((a, b) => b.row - a.row); // 가까운 위쪽부터
  const below = others.filter((c) => c.row > target.row).sort((a, b) => a.row - b.row); // 가까운 아래쪽부터
  return tryDir(above) ?? tryDir(below);
}

/** target ≈ rounding(base × rate) 를 만족하는 (base, rate) 조합 탐색 */
export function inferRateEdge(
  target: NumCell,
  bases: NumCell[],
  rates: NumCell[],
): { base: NumCell; rate: NumCell; rounding: RoundingKind } | null {
  for (const rate of rates) {
    if (!(rate.value > 0 && rate.value < 1)) continue;
    for (const base of bases) {
      if (base.address === target.address || base.value <= target.value) continue;
      const rounding = matchesWithRounding(base.value * rate.value, target.value);
      if (rounding) return { base, rate, rounding: rounding === 'EXACT' ? null : rounding };
    }
  }
  return null;
}

// ---------------------------------------------------------------------------
// 트리 조립
// ---------------------------------------------------------------------------

export type CostNodeStatus = 'FORMULA_VERIFIED' | 'VALUE_INFERRED' | 'MISMATCH' | 'UNRESOLVED';

export interface CostNode {
  id: string;
  label: string;
  labelCell: string | null;
  canonicalName: string | null;
  depth: number;
  amountCell: string;
  amount: number | null;
  rateCell: string | null;
  rate: number | null;
  baseCells: string[];
  baseAmount: number | null;
  baseLabel: string | null;
  children: CostNode[];
  status: CostNodeStatus;
  rounding: RoundingKind;
  note: string;
}

export interface CostTree {
  summarySheet: string | null;
  amountColumn: string | null;
  roots: CostNode[];
  issues: string[];
}

// 법정 골격: 값-추론 시 "부모 = 이 자식들의 합" 후보 (앞선 변형 우선)
const KNOWN_CHILDREN: Record<string, string[][]> = {
  '총계': [['총원가', '부가가치세']],
  '총원가': [['순공사원가', '일반관리비', '이윤'], ['재료비', '노무비', '경비', '일반관리비', '이윤']],
  '순공사원가': [['재료비', '노무비', '경비']],
  '재료비': [['직접재료비', '간접재료비'], ['직접재료비']],
  '노무비': [['직접노무비', '간접노무비'], ['직접노무비']],
};

// 요율 간선의 조합 산출기초 후보 (실무 관행: 재+노, 재+직노, 순공사원가 등)
const COMPOSITE_BASES: string[][] = [
  ['재료비', '노무비'], ['재료비', '직접노무비'], ['재료비', '노무비', '경비'],
  ['노무비', '경비'], ['노무비', '경비', '일반관리비'], ['재료비', '노무비', '경비', '일반관리비'],
  ['순공사원가'], ['총원가'], ['노무비'], ['직접노무비'],
];

interface SheetIndex {
  byAddr: Map<string, IRCell>;
  byRow: Map<number, IRCell[]>;
  numByCol: Map<number, NumCell[]>; // 열별 숫자 셀(FORMULA 캐시 포함), row 오름차순
}

function indexSheet(sheet: IRSheet): SheetIndex {
  const byAddr = new Map<string, IRCell>();
  const byRow = new Map<number, IRCell[]>();
  const numByCol = new Map<number, NumCell[]>();
  for (const cell of sheet.cells) {
    byAddr.set(cell.address, cell);
    const { row, col } = decodeAddr(cell.address);
    const rowArr = byRow.get(row) ?? [];
    rowArr.push(cell);
    byRow.set(row, rowArr);
    const num = parseKoreanNumber(cell.cachedValue);
    if (num != null && (cell.dataType === 'NUMBER' || cell.dataType === 'FORMULA')) {
      const colArr = numByCol.get(col) ?? [];
      colArr.push({ address: cell.address, row, col, value: num });
      numByCol.set(col, colArr);
    }
  }
  for (const arr of byRow.values()) arr.sort((a, b) => decodeAddr(a.address).col - decodeAddr(b.address).col);
  for (const arr of numByCol.values()) arr.sort((a, b) => a.row - b.row);
  return { byAddr, byRow, numByCol };
}

function makeWorkbookIndex(ir: WorkbookIR): Map<string, Map<string, IRCell>> {
  return new Map(ir.sheets.map((s) => [s.sheetName, new Map(s.cells.map((c) => [c.address, c]))]));
}

/** 통합북 단위 인덱스 1회 계산 — 시트 스코어링과 트리 조립이 같은 인덱스를 공유한다. */
export interface WorkbookContext {
  wb: Map<string, Map<string, IRCell>>;
  indexes: Map<string, SheetIndex>;
}

function makeWorkbookContext(ir: WorkbookIR): WorkbookContext {
  return {
    wb: makeWorkbookIndex(ir),
    indexes: new Map(ir.sheets.map((s) => [s.sheetName, indexSheet(s)])),
  };
}

/** "시트!F26" | "F26" 참조에서 셀 조회 */
function lookupCell(ref: string, currentSheet: string, wb: Map<string, Map<string, IRCell>>): IRCell | null {
  let sheetName = currentSheet;
  let addr = ref;
  const bang = ref.indexOf('!');
  if (bang >= 0) {
    sheetName = ref.slice(0, bang);
    addr = ref.slice(bang + 1);
  }
  return wb.get(sheetName)?.get(addr) ?? null;
}

// 순수 단일참조 수식(예: "=O7", "=+O7", "='내역서'!O7", "=$O$7") — 다른 셀을 화면에 그대로
// 비추기만 하는 "거울 셀" 판별용. 연산자·함수가 섞이면(=O7+1, =SUM(O7)) 매칭되지 않는다.
const PURE_SINGLE_REF_RE = /^=\+?(?:'[^']+'!|[A-Za-z0-9_가-힣]+!)?\$?[A-Z]{1,3}\$?\d+$/;

/** rawValue가 순수 단일참조 수식이면 그 참조를('$'·따옴표 제거해) 돌려준다. 아니면 null. */
function extractPureRefTarget(rawValue: string): string | null {
  if (!PURE_SINGLE_REF_RE.test(rawValue)) return null;
  return rawValue.replace(/^=\+?/, '').replace(/\$/g, '').replace(/'/g, '');
}

/**
 * 같은 행의 "거울 셀"(다른 셀 값을 그대로 표시용으로 재참조하는 순수 단일참조 수식) 체인을
 * 최대 3홉까지 따라가 최종 타깃 주소를 돌려준다. 순환이거나 더 이상 순수 참조가 아니면 그 자리에서 멈춘다.
 * startAddr 자신은 이미 호출측에서 "순수 단일참조 수식"임이 확인된 상태로 들어온다.
 */
function resolveMirrorChain(
  startAddr: string,
  currentSheet: string,
  wb: Map<string, Map<string, IRCell>>,
): string {
  let sheet = currentSheet;
  let addr = startAddr;
  const visited = new Set<string>([`${sheet}!${addr}`]);
  for (let hop = 0; hop < 3; hop++) {
    const cell = wb.get(sheet)?.get(addr);
    if (!cell || cell.dataType !== 'FORMULA') break;
    const target = extractPureRefTarget(cell.rawValue);
    if (!target) break;
    let nextSheet = sheet;
    let nextAddr = target;
    const bang = target.indexOf('!');
    if (bang >= 0) {
      nextSheet = target.slice(0, bang);
      nextAddr = target.slice(bang + 1);
    }
    const key = `${nextSheet}!${nextAddr}`;
    if (visited.has(key)) break; // 순환 방지
    visited.add(key);
    sheet = nextSheet;
    addr = nextAddr;
  }
  return sheet === currentSheet ? addr : `${sheet}!${addr}`;
}

/** row 행에서 beforeCol보다 왼쪽에 있는 STRING 중 CANONICAL → STRUCTURAL → 가장 왼쪽 순으로 라벨 셀을 고른다. */
function pickRowLabelCell(idx: SheetIndex, row: number, beforeCol: number): IRCell | null {
  const rowCells = idx.byRow.get(row) ?? [];
  const strings = rowCells.filter(
    (c) => c.dataType === 'STRING' && decodeAddr(c.address).col < beforeCol && cellLabel(c).trim(),
  );
  const canonicalCandidates = strings.filter((c) => matchCanonicalLabel(cellLabel(c)));
  const structuralCandidates = strings.filter((c) => matchStructuralLabel(cellLabel(c)));
  return canonicalCandidates[0] ?? structuralCandidates[0] ?? strings[0] ?? null;
}

/** 부가세 앵커: 라벨=부가가치세 행에서 "수식 리터럴≈10%" 또는 "같은 열 위쪽×0.1≈값"인 금액 셀 */
function findVatAnchor(
  idx: SheetIndex,
  currentSheet: string,
  wb: Map<string, Map<string, IRCell>>,
): { amountCell: string; col: number } | null {
  for (const [row, cells] of idx.byRow) {
    const hasVatLabel = cells.some(
      (c) => c.dataType === 'STRING' && matchCanonicalLabel(cellLabel(c)) === '부가가치세',
    );
    if (!hasVatLabel) continue;
    for (const c of cells) {
      const v = parseKoreanNumber(c.cachedValue);
      if (v == null || v < 1) continue; // 요율(0<x<1)·빈 셀 제외
      const { col } = decodeAddr(c.address);
      if (c.dataType === 'FORMULA') {
        const shape = decomposeFormula(c.rawValue, c.references, () => false);
        if (shape.literalRate != null && Math.abs(shape.literalRate - 0.1) < 1e-9) {
          return { amountCell: c.address, col };
        }
        // 리터럴이 아닌 요율 셀 참조형(=D31*E32 등)도 값 검사로 흡수
      }
      const column = idx.numByCol.get(col) ?? [];
      if (column.some((nc) => nc.row < row && matchesWithRounding(nc.value * 0.1, v))) {
        return { amountCell: c.address, col };
      }
    }
  }
  return null;
}

/** 앵커 실패 시 폴백: 1000 이상 숫자 셀이 가장 많은 열 */
function dominantNumericColumn(idx: SheetIndex): number {
  let best = 1;
  let bestCount = -1;
  for (const [col, cells] of idx.numByCol) {
    const big = cells.filter((c) => c.value >= 1000).length;
    if (big > bestCount) {
      bestCount = big;
      best = col;
    }
  }
  return best;
}

/** 금액 열의 라벨된 행: canonical → NumCell. 총계·총원가는 최하단, 그 외 최상단 채택 */
function collectLabeledAmounts(idx: SheetIndex, amountColNum: number): Map<string, NumCell> {
  const map = new Map<string, NumCell>();
  const rows = [...idx.byRow.entries()].sort((a, b) => a[0] - b[0]);
  for (const [row, cells] of rows) {
    for (const cell of cells) {
      if (cell.dataType !== 'STRING') continue;
      const m = matchCanonicalLabel(cellLabel(cell)) ?? matchStructuralLabel(cellLabel(cell));
      if (!m) continue;
      const amtCell = idx.byAddr.get(`${numToCol(amountColNum)}${row}`);
      const v = amtCell ? parseKoreanNumber(amtCell.cachedValue) : null;
      if (v == null) continue;
      const nc: NumCell = { address: amtCell!.address, row, col: amountColNum, value: v };
      if (m === '총계' || m === '총원가') map.set(m, nc); // 아래쪽 행이 최종 합계 — 마지막 갱신 유지
      else if (!map.has(m)) map.set(m, nc);
    }
  }
  return map;
}

// 호출 간 파싱 캐시. parseCostTree는 IR 객체에 대한 순수 함수이고(입력 IR을 변형하지 않는다),
// 한 업로드에서 4~5회 호출된다(App useMemo · detectItems 경유 3곳 · normalizeCostStatement).
// 138k셀 워크북을 매번 다시 인덱싱하는 비용을 IR 객체 동일성 기준으로 1회로 줄인다.
// WeakMap이므로 IR이 버려지면 캐시도 함께 회수된다.
const treeCache = new WeakMap<WorkbookIR, CostTree>();

export function parseCostTree(ir: WorkbookIR): CostTree {
  const cached = treeCache.get(ir);
  if (cached) return cached;
  const tree = parseCostTreeUncached(ir);
  treeCache.set(ir, tree);
  return tree;
}

function parseCostTreeUncached(ir: WorkbookIR): CostTree {
  const issues: string[] = [];
  const ctx = makeWorkbookContext(ir); // 인덱스는 1회만 — 스코어링과 트리 조립이 공유
  const summarySheet = findSummarySheet(ir, ctx);
  if (!summarySheet) {
    return { summarySheet: null, amountColumn: null, roots: [], issues: ['요약시트를 찾지 못했습니다.'] };
  }
  const { wb } = ctx;
  const idx = ctx.indexes.get(summarySheet)!;

  // ② 부가세 앵커로 금액 열 확정
  const anchor = findVatAnchor(idx, summarySheet, wb);
  const amountColNum = anchor?.col ?? dominantNumericColumn(idx);
  if (!anchor) issues.push('부가가치세 앵커를 찾지 못해 숫자 최다 열로 금액 열을 추정했습니다.');
  const labeled = collectLabeledAmounts(idx, amountColNum);

  const isRateRef = (ref: string): boolean => {
    const rc = lookupCell(ref, summarySheet, wb);
    if (!rc) return false;
    const v = parseKoreanNumber(rc.cachedValue);
    return Boolean(rc.numberFormat?.includes('%')) || (v != null && v > 0 && v < 1);
  };

  const seen = new Set<string>();

  const buildNode = (amountAddr: string, depth: number): CostNode | null => {
    if (seen.has(amountAddr) || depth > 12) return null;
    seen.add(amountAddr);
    const cell = idx.byAddr.get(amountAddr);
    if (!cell) return null;
    const amount = parseKoreanNumber(cell.cachedValue);
    const { row, col } = decodeAddr(amountAddr);

    // 라벨: 금액 셀 왼쪽 STRING 중 CANONICAL 매칭되는 가장 왼쪽 → 없으면 STRUCTURAL 매칭 가장 왼쪽 → 없으면 가장 왼쪽 STRING.
    // 왼쪽부터 찾는 이유: 이 행의 "구분"(주 라벨)은 보통 맨 앞 열에 오고, 그 뒤(오른쪽)에는 산출근거 설명이나
    // 기초금액의 "기준 항목명"(예: 간접노무비 행의 "직접노무비" 기초 설명)이 이어진다 — 그 설명문이 우연히
    // 다른 canonical/structural 이름을 포함해도 주 라벨보다 먼저 매칭되면 안 된다.
    // CANONICAL을 STRUCTURAL보다 먼저 보는 이유: 같은 행에 "이윤"(주 라벨)과 "일반관리비"(기초 설명, 역시
    // CANONICAL)가 함께 나오는 경우에도 반드시 있는 canonical 매칭 중 leftmost가 주 라벨이다.
    const rowCells = idx.byRow.get(row) ?? [];
    const labelCellObj = pickRowLabelCell(idx, row, col);
    const label = labelCellObj ? cellLabel(labelCellObj).trim() : '';
    const canonicalName = label ? matchCanonicalLabel(label) ?? matchStructuralLabel(label) : null;

    const node: CostNode = {
      id: `${summarySheet}!${amountAddr}`,
      label, labelCell: labelCellObj?.address ?? null, canonicalName, depth,
      amountCell: amountAddr, amount,
      rateCell: null, rate: null, baseCells: [], baseAmount: null, baseLabel: null, children: [],
      status: 'UNRESOLVED', rounding: null, note: '',
    };

    if (cell.dataType === 'FORMULA') {
      // ── 1순위: 수식 간선
      const shape = decomposeFormula(cell.rawValue, cell.references, isRateRef);
      node.status = 'FORMULA_VERIFIED';
      node.rounding = shape.rounding;
      if (shape.rateRefs.length > 0) {
        node.rateCell = shape.rateRefs[0];
        const rc = lookupCell(node.rateCell, summarySheet, wb);
        if (rc) node.rate = parseRate(rc.cachedValue as string | number | boolean | null, rc.numberFormat);
      } else if (shape.literalRate != null) {
        node.rate = shape.literalRate;
      } else if (shape.baseRefs.length > 0 && shape.baseRefs.every((r) => r.includes('!'))) {
        // 금액 수식이 전부 시트-간 참조로만 구성된 경우(예: ='총괄표'!M73) 수식 자체에서는 rate를 못 뽑는다.
        // 이런 파일은 보통 요율 계산은 다른 시트에서 끝내고 이 시트엔 결과 금액 + 참고용 요율 셀만
        // 병기하는 패턴이므로, 같은 행의 요율 셀을 채택한다(ref-07 '원가계산' 시트 사례: F45 = 0.0% 서식).
        // "0<x<1 숫자"만으로는 비율·계수·소수 잔액 등 아무 숫자나 요율로 오인할 수 있어
        // %서식(numberFormat)을 필수 조건으로 두고, 값 해석은 로컬 요율 경로와 같은 parseRate를 쓴다.
        const rateCand = rowCells.find((c) => {
          if (c.address === amountAddr || (c.dataType !== 'NUMBER' && c.dataType !== 'FORMULA')) return false;
          if (!c.numberFormat?.includes('%')) return false;
          const v = parseKoreanNumber(c.cachedValue);
          return v != null && v > 0 && v < 1;
        });
        if (rateCand) {
          node.rate = parseRate(rateCand.cachedValue as string | number | boolean | null, rateCand.numberFormat);
          node.rateCell = rateCand.address;
        }
      }
      for (const ref of shape.baseRefs) {
        // 거울 셀: 같은 행 + 로컬 참조 + "다른 셀을 그대로 비추기만 하는" 순수 단일참조 수식
        // (예: 간접노무비 행의 E8=$O$7, 직접노무비 실금액을 표시용으로 재참조). 이런 셀은 별도
        // 자식 노드로 만들면 같은 라벨이 중복 노출되므로, 체인을 따라간 최종 타깃을 baseCells에
        // 기록하기만 하고 자식은 만들지 않는다. 숫자 셀(수식 아님)이나 복합 수식은 기존 경로 유지.
        if (!ref.includes('!') && decodeAddr(ref).row === row) {
          const mirrorCell = idx.byAddr.get(ref);
          if (mirrorCell?.dataType === 'FORMULA' && PURE_SINGLE_REF_RE.test(mirrorCell.rawValue)) {
            node.baseCells.push(resolveMirrorChain(ref, summarySheet, wb));
            continue;
          }
        }
        node.baseCells.push(ref);
        if (!ref.includes('!')) {
          const child = buildNode(ref, depth + 1); // 이미 방문한 셀은 null → baseCells 기록만 남음
          if (child) {
            // 같은 행에 있는 기초금액 셀(예: "노무비 [기초] × 요율 = 결과" 한 행에 라벨·기초·결과가 함께 있는
            // 레이아웃)은 라벨 탐색 시 이 행의 라벨을 그대로 재사용해 canonicalName을 얻지만, 그 값은
            // 이 행의 결과(부모 amount)가 아니라 참조된 다른 항목의 총액이므로 canonical 항목으로 잘못 노출된다.
            // 값이 부모와 다르면(=단순 passthrough가 아니라 별도 입력값) canonicalName을 제거해 오탐을 막는다.
            // 값이 같으면(예: ref-09 D열=SUM(E열) 단순 재참조) 정당한 중복 표기이므로 유지한다.
            // "같다" 판정은 절사 허용 비교(matchesWithRounding)로 한다 — 부모가 INT/ROUND로 적힌
            // passthrough(예: =INT(E12), 1원 차이)를 별도 입력값으로 오판해 정당한 항목을 죽이면 안 된다.
            if (decodeAddr(ref).row === row && !isSameAmount(child.amount, amount)) child.canonicalName = null;
            node.children.push(child);
          }
        }
      }
    } else if (amount != null) {
      // ── 2순위: 값-관계 추론
      const target: NumCell = { address: amountAddr, row, col, value: amount };
      let matched = false;

      // ①' 법정 골격 합 (라벨 기반)
      for (const variant of node.canonicalName ? KNOWN_CHILDREN[node.canonicalName] ?? [] : []) {
        const cells = variant.map((name) => labeled.get(name)).filter((x): x is NumCell => Boolean(x));
        if (cells.length !== variant.length) continue;
        const sum = cells.reduce((s, c) => s + c.value, 0);
        if (!matchesWithRounding(sum, amount)) continue;
        for (const nc of cells) {
          const child = buildNode(nc.address, depth + 1);
          if (child) node.children.push(child);
          else node.baseCells.push(nc.address);
        }
        node.status = 'VALUE_INFERRED';
        matched = true;
        break;
      }

      // ②' 부가세 특례: 총원가×10%
      if (!matched && node.canonicalName === '부가가치세') {
        const total = labeled.get('총원가');
        if (total && matchesWithRounding(total.value * 0.1, amount)) {
          node.rate = 0.1;
          node.baseCells.push(total.address);
          node.status = 'VALUE_INFERRED';
          matched = true;
        }
      }

      // ③' 요율 간선 (단일 셀 기초 + 조합 기초)
      if (!matched) {
        const columnAbove = (idx.numByCol.get(col) ?? []).filter((c) => c.row < row);
        const composites: NumCell[] = COMPOSITE_BASES.flatMap((combo) => {
          const cells = combo.map((n) => labeled.get(n));
          if (cells.some((x) => !x)) return [];
          return [{ address: combo.join('+'), row: -1, col, value: cells.reduce((s, c) => s + c!.value, 0) }];
        });
        const rowRates: NumCell[] = rowCells
          .map((c) => ({ c, v: parseKoreanNumber(c.cachedValue) }))
          .filter((x): x is { c: IRCell; v: number } => x.v != null && x.v > 0 && x.v < 1)
          .map((x) => ({ address: x.c.address, row, col: decodeAddr(x.c.address).col, value: x.v }));
        const edge = inferRateEdge(target, [...columnAbove, ...composites], rowRates);
        if (edge) {
          node.rate = edge.rate.value;
          node.rateCell = edge.rate.row >= 0 ? edge.rate.address : null;
          node.rounding = edge.rounding;
          node.baseCells.push(edge.base.address);
          // 값-추론 요율 간선은 기초값이 이미 확정돼 있다(edge.base) — 조합 기초(가짜 주소, 예:
          // "재료비+노무비")는 lookupCell로 조회할 수 없으므로 여기서 직접 채운다.
          node.baseAmount = edge.base.value;
          if (edge.base.row >= 0) {
            const lc = pickRowLabelCell(idx, edge.base.row, edge.base.col);
            node.baseLabel = lc ? cellLabel(lc).trim() : null;
          } else {
            node.baseLabel = edge.base.address; // 조합 기초 주소 문자열 그대로(예: '재료비+노무비')
          }
          node.status = 'VALUE_INFERRED';
          matched = true;
        }
      }

      // ④' 인접 동일값 (재료비=직접재료비 단일 자식 등, 3행 이내)
      if (!matched) {
        const near = (idx.numByCol.get(col) ?? []).find(
          (c) => c.row > row && c.row <= row + 3 && c.value === amount,
        );
        if (near) {
          const child = buildNode(near.address, depth + 1);
          if (child) node.children.push(child);
          node.status = 'VALUE_INFERRED';
          matched = true;
        }
      }

      // ⑤' 연속 구간 합
      if (!matched) {
        const rangeHit = inferSumRange(target, idx.numByCol.get(col) ?? []);
        if (rangeHit) {
          for (const nc of rangeHit) {
            const child = buildNode(nc.address, depth + 1);
            if (child) node.children.push(child);
          }
          node.status = 'VALUE_INFERRED';
        }
      }
    }

    // 산출기초 금액/라벨: 요율이 있고 기초 참조가 하나 이상이면 채운다. VALUE_INFERRED 요율 간선(③')은
    // 이미 위에서 직접 채웠으므로(조합 기초는 lookupCell로 조회 불가) node.baseAmount == null 가드로 건너뜀.
    if (node.rate != null && node.baseCells.length > 0 && node.baseAmount == null) {
      let sum = 0;
      let any = false;
      for (const ref of node.baseCells) {
        const bc = lookupCell(ref, summarySheet, wb);
        const v = bc ? parseKoreanNumber(bc.cachedValue) : null;
        if (v != null) {
          sum += v;
          any = true;
        }
      }
      node.baseAmount = any ? sum : null;
      const localBaseRefs = node.baseCells.filter((r) => !r.includes('!'));
      if (localBaseRefs.length === 1) {
        const { row: bRow, col: bCol } = decodeAddr(localBaseRefs[0]);
        const lc = pickRowLabelCell(idx, bRow, bCol);
        node.baseLabel = lc ? cellLabel(lc).trim() : null;
      }
    }

    // 검산: 자식 2개 이상 + 요율 간선 아님 → 부모=자식합
    if (node.children.length >= 2 && node.rate == null && amount != null) {
      const sum = node.children.reduce((s, c) => s + (c.amount ?? 0), 0);
      if (sum !== amount && !matchesWithRounding(sum, amount)) {
        node.status = 'MISMATCH';
        node.note = `부모-자식 합 불일치 (차이 ${Math.round(amount - sum).toLocaleString('en-US')})`;
      }
    }
    return node;
  };

  // ③ 루트: 총계 → 총원가 → 부가세 앵커
  const roots: CostNode[] = [];
  const rootCell = labeled.get('총계') ?? labeled.get('총원가') ?? null;
  const rootAddr = rootCell?.address ?? anchor?.amountCell ?? null;
  if (rootAddr) {
    const root = buildNode(rootAddr, 0);
    if (root) roots.push(root);
  } else {
    issues.push('총계·총원가·부가가치세 행을 찾지 못했습니다.');
  }

  // ⑤ 미연결 CANONICAL 행 추가 (누락 방지)
  for (const [row, cells] of idx.byRow) {
    for (const cell of cells) {
      const canon = cell.dataType === 'STRING' ? matchCanonicalLabel(cellLabel(cell)) : null;
      if (!canon) continue;
      const amtAddr = `${numToCol(amountColNum)}${row}`;
      if (seen.has(amtAddr)) continue; // 같은 행에 라벨이 여러 번 나와도 행당 1개만
      const amtCell = idx.byAddr.get(amtAddr);
      const amtValue = amtCell ? parseKoreanNumber(amtCell.cachedValue) : null;
      if (amtValue != null) {
        const extra = buildNode(amtAddr, 0);
        if (extra) {
          extra.note = extra.note ? `${extra.note} · 트리 미연결` : '트리 미연결';
          roots.push(extra);
        }
        continue;
      }
      // 금액을 못 읽은 CANONICAL 행: 조용히 버리지 않고 UNRESOLVED로 표면화한다(스펙: "못 읽은 항목은
      // 전부 상태로 표면화"). amount가 null이므로 costTreeToDetectedItems는 이 노드를 건너뛴다 —
      // 검증/골든 집계에는 영향이 없고, 인식결과 화면에서 "이 항목은 못 읽었다"가 보이게만 한다.
      // 금액 열 셀이 STRING이면(설명문·머리글 겹침) 금액 자리로 볼 수 없으므로 제외한다.
      if (amtCell?.dataType === 'STRING') continue;
      seen.add(amtAddr);
      const label = cellLabel(cell).trim();
      roots.push({
        id: `${summarySheet}!${amtAddr}`,
        label, labelCell: cell.address, canonicalName: canon, depth: 0,
        amountCell: amtAddr, amount: null,
        rateCell: null, rate: null, baseCells: [], baseAmount: null, baseLabel: null, children: [],
        status: 'UNRESOLVED', rounding: null,
        note: `금액 셀 미발견 (금액 열 ${numToCol(amountColNum)})`,
      });
    }
  }

  return { summarySheet, amountColumn: numToCol(amountColNum), roots, issues };
}

export function flattenTree(tree: CostTree): CostNode[] {
  const out: CostNode[] = [];
  const visited = new Set<string>();
  const walk = (n: CostNode) => {
    if (!visited.has(n.id)) {
      visited.add(n.id);
      out.push(n);
    }
    n.children.forEach(walk);
  };
  tree.roots.forEach(walk);
  return out;
}

/**
 * 트리 결과를 채택하는 표준항목 개수 하한. 이 값 미만이면 detectItems가 레거시 라벨 스캔으로
 * 폴백하고(validation.ts), 인식결과 화면은 "표는 참고용" 경고를 띄운다(CostTreeView.tsx).
 * 두 곳이 같은 상수를 봐야 화면과 검증의 분기점이 어긋나지 않는다.
 */
export const TREE_ITEM_MIN = 3;

/** CANONICAL_ITEMS에 등록된 항목만 기존 DetectedItem 계약으로 투영 (STRUCTURAL 골격은 제외) */
export function costTreeToDetectedItems(tree: CostTree): DetectedItem[] {
  if (!tree.summarySheet) return [];
  const out: DetectedItem[] = [];
  const emitted = new Set<string>();
  for (const node of flattenTree(tree)) {
    if (!node.canonicalName || node.amount == null) continue;
    const def = getCanonicalDef(node.canonicalName);
    if (!def || emitted.has(node.amountCell)) continue;
    emitted.add(node.amountCell);
    out.push({
      canonicalName: def.canonical,
      originalLabel: node.label,
      category: def.category,
      sheetName: tree.summarySheet,
      labelCell: node.labelCell ?? node.amountCell,
      amountCell: node.amountCell,
      rateCell: node.rateCell,
      amountValue: node.amount,
      rateValue: node.rate,
      requiresReference: def.requiresReference,
    });
  }
  return out;
}
