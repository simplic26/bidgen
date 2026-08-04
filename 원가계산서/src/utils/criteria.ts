import type { IRCell, LaborRate, ProcurementType, RateCriterion, WorkbookIR } from '../types';
import { buildWorkbookIR } from './excel';
import { matchCanonicalLabel, parseKoreanNumber, parseRate } from './validation';

// --- A1 좌표 헬퍼 (xlsx 비의존) ---
function decodeCol(letters: string): number {
  let c = 0;
  for (const ch of letters) c = c * 26 + (ch.charCodeAt(0) - 64);
  return c - 1;
}

function decodeCell(addr: string): { c: number; r: number } {
  const m = /^\$?([A-Z]{1,3})\$?(\d+)$/.exec(addr);
  if (!m) return { c: 0, r: 0 };
  return { c: decodeCol(m[1]), r: parseInt(m[2], 10) - 1 };
}

function mergedEndCol(cell: IRCell): number | null {
  if (!cell.mergedRange) return null;
  const [, end] = cell.mergedRange.split(':');
  return end ? decodeCell(end).c : null;
}

function groupByRow(cells: IRCell[]): Map<number, IRCell[]> {
  const byRow = new Map<number, IRCell[]>();
  for (const cell of cells) {
    const { r } = decodeCell(cell.address);
    const arr = byRow.get(r) ?? [];
    arr.push(cell);
    byRow.set(r, arr);
  }
  for (const arr of byRow.values()) arr.sort((a, b) => decodeCell(a.address).c - decodeCell(b.address).c);
  return byRow;
}

/**
 * 제비율 Excel IR을 훑어 표준 원가항목의 기준요율과 셀 근거를 추출한다. (PRD §8.2 / §8.5 RATE_EXCEL)
 * 항목명 별칭이 있는 행에서 같은 행의 %서식 또는 0<x<1 숫자 셀을 요율로 인식한다.
 */
export function buildRateCriteria(ir: WorkbookIR): RateCriterion[] {
  const out: RateCriterion[] = [];
  const seen = new Set<string>();

  for (const sheet of ir.sheets) {
    const byRow = groupByRow(sheet.cells);
    for (const cell of sheet.cells) {
      if (cell.dataType !== 'STRING') continue;
      const label = typeof cell.cachedValue === 'string' ? cell.cachedValue : cell.displayValue;
      const canonical = matchCanonicalLabel(label);
      if (!canonical || canonical === '부가가치세' || seen.has(canonical)) continue;

      const { r, c: labelCol } = decodeCell(cell.address);
      const rowCells = (byRow.get(r) ?? []).filter((rc) => decodeCell(rc.address).c > labelCol);

      let rateCell: IRCell | null = null;
      for (const rc of rowCells) {
        const num = parseKoreanNumber(rc.cachedValue);
        if (num == null) continue;
        const isPercentFormat = rc.numberFormat?.includes('%') ?? false;
        const looksRate = isPercentFormat || (num > 0 && num < 1);
        if (looksRate) {
          rateCell = rc;
          break;
        }
      }
      if (!rateCell) continue;

      const rate = parseRate(rateCell.cachedValue, rateCell.numberFormat);
      if (rate == null || rate <= 0 || rate >= 1) continue;

      seen.add(canonical);
      out.push({
        canonicalName: canonical,
        rate,
        sheetName: sheet.sheetName,
        cell: rateCell.address,
        displayValue: rateCell.displayValue || `${(rate * 100).toFixed(3)}%`,
        confidence: rateCell.numberFormat?.includes('%') ? 0.95 : 0.8,
      });
    }
  }
  return out;
}

// 간접공사비 적용기준(조달청형 매트릭스)의 산출기초식 헤더 → 표준 원가항목 매핑.
// 원가계산서 실제 수식으로 교차검증됨: (재+노+경)→일반관리비, (노+경+일)→이윤, (재+직노)+기초액→산업안전보건관리비 등.
const BASIS_RULES: Array<{ test: (s: string) => boolean; item: string; confidence: number }> = [
  { test: (s) => s.includes('기초액'), item: '산업안전보건관리비', confidence: 0.85 }, // (재+직노)×율+기초액
  { test: (s) => s.includes('재+노+경') || s.includes('재,노,경'), item: '일반관리비', confidence: 0.9 },
  { test: (s) => s.includes('노+경+일'), item: '이윤', confidence: 0.9 },
  { test: (s) => s.includes('직접공사비'), item: '안전관리비', confidence: 0.6 },
  { test: (s) => s.includes('재+노') && !s.includes('재+노+경'), item: '기타경비', confidence: 0.7 },
  { test: (s) => s.includes('직노'), item: '간접노무비', confidence: 0.8 }, // (직노)×율
];

function normHeader(s: string): string {
  return s.replace(/\s+/g, '').toLowerCase();
}

function matchBasis(text: string): { item: string; confidence: number } | null {
  const s = normHeader(text);
  if (!s.includes('율')) return null; // 요율 헤더만 (예: "(직노) x 율")
  for (const rule of BASIS_RULES) if (rule.test(s)) return { item: rule.item, confidence: rule.confidence };
  return null;
}

/**
 * 간접공사비 적용기준 매트릭스(산출기초식 헤더 × 공사규모·기간 구간)에서 요율을 추출한다.
 * 값이 %서식이 아니라 평범한 숫자(17.5 = 17.5%)인 형식을 처리하며, 최소 구간(첫 데이터행) 요율을 채택하고 구간 조건을 함께 기록한다.
 */
function buildRateCriteriaMatrix(ir: WorkbookIR): RateCriterion[] {
  const out: RateCriterion[] = [];
  const seen = new Set<string>();

  for (const sheet of ir.sheets) {
    const byRow = groupByRow(sheet.cells);
    const rowIdxs = [...byRow.keys()].sort((a, b) => a - b);

    // 헤더 행 탐지: 산출기초식(…×율) 셀이 2개 이상인 행
    let headerRow = -1;
    const headerCells: Array<{ col: number; endCol: number; item: string; confidence: number }> = [];
    for (const r of rowIdxs) {
      const matches: Array<{ col: number; endCol: number; item: string; confidence: number }> = [];
      for (const cell of byRow.get(r)!) {
        if (cell.dataType !== 'STRING') continue;
        const m = matchBasis(String(cell.cachedValue ?? cell.displayValue));
        if (m) {
          const col = decodeCell(cell.address).c;
          matches.push({ col, endCol: mergedEndCol(cell) ?? col, ...m });
        }
      }
      if (matches.length >= 2) {
        headerRow = r;
        headerCells.push(...matches);
        break;
      }
    }
    if (headerRow < 0) continue;

    // 헤더 그룹의 열 범위(현재 헤더 ~ 다음 헤더 직전) 계산 — 값이 헤더 열이 아닌 하위 열에 있는 경우 대비
    const sortedHeaders = [...headerCells].sort((a, b) => a.col - b.col);
    const spanEnd = (header: { col: number; endCol: number }) => {
      const next = sortedHeaders.find((candidate) => candidate.col > header.col);
      const nextLimit = next ? next.col - 1 : null;
      if (header.endCol > header.col) return nextLimit == null ? header.endCol : Math.min(header.endCol, nextLimit);
      return nextLimit != null ? Math.min(nextLimit, header.col + 10) : header.col + 10;
    };

    // 각 헤더 그룹에서 첫 숫자 셀(왼→오, 위→아래) = 최소 구간 요율
    for (const h of headerCells) {
      if (seen.has(h.item)) continue;
      let rateCell: IRCell | null = null;
      let dataRow = -1;
      const end = spanEnd(h);
      let rateColIdx = h.col;
      // 요율 후보: 0<값<100 (기초액 등 큰 숫자는 요율이 아니므로 제외)
      outer: for (let col = h.col; col <= end; col += 1) {
        for (const r of rowIdxs) {
          if (r <= headerRow) continue;
          const cell = byRow.get(r)!.find((c) => decodeCell(c.address).c === col);
          const num = cell ? parseKoreanNumber(cell.cachedValue) : null;
          if (num != null && num > 0 && num < 100) {
            rateCell = cell!;
            dataRow = r;
            rateColIdx = col;
            break outer;
          }
        }
      }
      if (!rateCell) continue;
      const raw = parseKoreanNumber(rateCell.cachedValue);
      if (raw == null || raw <= 0) continue;
      const rate = raw > 1 ? raw / 100 : raw; // 17.5 → 0.175, 0.175 → 0.175

      // 구간 조건: 데이터행 왼쪽의 텍스트 셀들(공사규모·기간)
      const condParts = byRow
        .get(dataRow)!
        .filter((c) => c.dataType === 'STRING' && decodeCell(c.address).c < rateColIdx)
        .map((c) => String(c.cachedValue ?? c.displayValue).replace(/\s+/g, ' ').trim())
        .filter(Boolean);
      const condition = condParts.join(' / ') || '최소 구간';

      seen.add(h.item);
      out.push({
        canonicalName: h.item,
        rate,
        sheetName: sheet.sheetName,
        cell: rateCell.address,
        displayValue: `${(rate * 100).toFixed(2)}%`,
        confidence: h.confidence,
        condition,
      });
    }
  }
  return out;
}

/**
 * 노임단가 Excel IR을 훑어 직종별 단가와 셀 근거를 추출한다. (PRD §8.3 / §8.5 LABOR_RATE_EXCEL)
 * '직종명(한글 문자) + 단가(1000 이상 숫자)' 행을 인식한다.
 */
export function buildLaborRates(ir: WorkbookIR, opts?: { maxRows?: number }): LaborRate[] {
  const maxRows = opts?.maxRows ?? 500;
  const out: LaborRate[] = [];
  const seen = new Set<string>();

  for (const sheet of ir.sheets) {
    const byRow = groupByRow(sheet.cells);
    for (const [, cells] of byRow) {
      if (out.length >= maxRows) break;

      const labelCell = cells.find(
        (rc) => rc.dataType === 'STRING' && /[가-힣]/.test(String(rc.cachedValue ?? rc.displayValue)),
      );
      if (!labelCell) continue;
      const name = String(labelCell.cachedValue ?? labelCell.displayValue).trim();
      // 직종명 휴리스틱: 길이 1~12, 숫자 미포함 (헤더/설명행 배제)
      if (name.length === 0 || name.length > 12 || /\d/.test(name)) continue;

      const labelCol = decodeCell(labelCell.address).c;
      let priceCell: IRCell | null = null;
      for (const rc of cells) {
        if (decodeCell(rc.address).c <= labelCol) continue;
        const num = parseKoreanNumber(rc.cachedValue);
        if (num != null && num >= 1000) {
          priceCell = rc; // 라벨 오른쪽 첫 1000 이상 숫자 = 단가 열
          break;
        }
      }
      if (!priceCell) continue;

      const unitPrice = parseKoreanNumber(priceCell.cachedValue);
      if (unitPrice == null) continue;
      if (seen.has(name)) continue;

      seen.add(name);
      out.push({
        occupationName: name,
        unitPrice,
        sheetName: sheet.sheetName,
        cell: priceCell.address,
        displayValue: priceCell.displayValue || unitPrice.toLocaleString('en-US'),
        confidence: 0.7,
      });
    }
  }
  return out;
}

/** 기준자료 Excel 파일을 브라우저에서 파싱해 IR과 추출 결과를 반환한다. */
export async function parseRateWorkbook(
  file: File,
  procurementType: ProcurementType,
): Promise<{ ir: WorkbookIR; criteria: RateCriterion[] }> {
  const ir = await buildWorkbookIR(file, procurementType);
  // 간접공사비 매트릭스형(산출기초식 헤더)을 우선 시도, 항목이 잡히면 그대로 사용. 아니면 라벨-스캔 폴백.
  const matrix = buildRateCriteriaMatrix(ir);
  return { ir, criteria: matrix.length >= 2 ? matrix : buildRateCriteria(ir) };
}

export async function parseLaborWorkbook(
  file: File,
  procurementType: ProcurementType,
): Promise<{ ir: WorkbookIR; laborRates: LaborRate[] }> {
  const ir = await buildWorkbookIR(file, procurementType);
  return { ir, laborRates: buildLaborRates(ir) };
}
