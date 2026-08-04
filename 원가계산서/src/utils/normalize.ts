import type { NormalizedCostRow, ResolutionStatus, ValidationResult, ValidationType } from '../types';

// 검증정책 매핑 (PRD §6.4 검증정책 열)
const POLICY_BY_TYPE: Record<ValidationType, string> = {
  RATE: 'RATE_CHECK',
  LABOR: 'LABOR_RATE_CHECK',
  ARITHMETIC: 'INTERNAL_CONSISTENCY',
  TOTAL: 'INTERNAL_CONSISTENCY',
  FORMULA: 'INTERNAL_CONSISTENCY',
  REFERENCE: 'INTERNAL_CONSISTENCY',
  BASE: 'REFERENCE_REQUIRED',
  CONDITION: 'MANUAL_REVIEW',
  STANDARD: 'STANDARD_PRODUCTION_CHECK',
};

// 비용성격 매핑 (원본구간=category, 비용성격=검증 성격)
const NATURE_BY_TYPE: Record<ValidationType, string> = {
  RATE: '제비율성 경비',
  LABOR: '노임성 직접비',
  BASE: '별도 직접비',
  STANDARD: '표준품셈 연계',
  ARITHMETIC: '산술 검증',
  TOTAL: '집계 검증',
  FORMULA: '수식 검증',
  REFERENCE: '참조 검증',
  CONDITION: '적용조건 검증',
};

// 산출기초 해결상태 추론 (PRD FR-037)
function resolveStatus(result: ValidationResult): ResolutionStatus {
  const { formula, referencedCells } = result.excel;
  const hasFormula = !!formula && formula.trim().startsWith('=');
  const crossSheet = (formula ?? '').includes('!') || referencedCells.some((ref) => ref.includes('!'));
  if (crossSheet) return 'CROSS_SHEET_TRACE';
  if (hasFormula) return referencedCells.length > 0 ? 'FORMULA_TRACE' : 'FORMULA_UNRESOLVED';
  if (referencedCells.length > 0) return 'SAME_ROW';
  return 'UNRESOLVED';
}

// 수식 안에 숨은 상수(보정값·절사단위) 추출 — 셀참조/함수명 제거 후 남는 숫자
function extractConstants(formula: string | undefined): string {
  if (!formula) return '';
  const stripped = formula
    .replace(/[A-Za-z가-힣_']+!/g, '') // 시트 접두사
    .replace(/\b[A-Z]{1,3}\d+(?::[A-Z]{1,3}\d+)?\b/g, '') // 셀/범위 참조
    .replace(/\b(SUM|ROUND|ROUNDDOWN|INT)\b/gi, ''); // 함수명
  const nums = stripped.match(/-?\d+(?:\.\d+)?/g);
  return nums ? nums.join(', ') : '';
}

/**
 * 정규화 원가계산서 행(normalizeCostStatement 산출물)에 검증결과 상태를 주석으로 입힌다. (PRD §9.1 → §6.4)
 * 행의 존재는 정규화가 결정하고, 검증은 각 행의 상태·계산값·차액만 채운다(역산 아님).
 * 계산셀(calculationCell)로 결과를 매칭하며, 같은 셀에 여러 결과가 있으면 비-정상 결과를 우선한다.
 */
export function annotateRows(rows: NormalizedCostRow[], results: ValidationResult[]): NormalizedCostRow[] {
  const byCell = new Map<string, ValidationResult>();
  for (const r of results) {
    const key = `${r.excel.sheetName}!${r.excel.cell}`;
    const existing = byCell.get(key);
    if (!existing || (existing.status === 'OK' && r.status !== 'OK')) byCell.set(key, r);
  }
  return rows.map((row) => {
    const res = byCell.get(row.calculationCell);
    if (!res) return row;
    return {
      ...row,
      rowId: res.resultId, // 결과 리스트와 선택 동기화
      status: res.status,
      calculatedAmount: res.expected.finalAmount || row.calculatedAmount,
      difference: res.difference || row.difference,
      rate: res.expected.rate ?? row.rate,
      note: res.summary || row.note,
    };
  });
}

/**
 * (딥링크 데모 전용) 검증결과에서 정규화 표 행을 파생한다. 실제 업로드 경로는 normalizeCostStatement + annotateRows를 사용한다.
 */
export function buildNormalizedRows(results: ValidationResult[]): NormalizedCostRow[] {
  return results.map((r) => {
    const baseLabel =
      r.excel.referencedCells.find((ref) => /[가-힣]/.test(ref))?.replace(/^[^ ]+\s+/, '') ?? '';
    return {
      rowId: r.resultId,
      status: r.status,
      sourceSection: r.item.category || '기타',
      costNature: NATURE_BY_TYPE[r.validationType] ?? '기타',
      validationPolicy: POLICY_BY_TYPE[r.validationType] ?? 'MANUAL_REVIEW',
      canonicalName: r.item.canonicalName,
      originalName: r.item.originalName,
      baseLabel,
      baseAmount: r.expected.baseAmount || '',
      rate: r.expected.rate ?? r.excel.inputRate ?? '',
      rateSource: r.excel.rateCell ? `${r.excel.sheetName}!${r.excel.rateCell}` : '',
      amount: r.excel.inputValue || '',
      calculatedAmount: r.expected.finalAmount || '',
      difference: r.difference || '0',
      calculationCell: `${r.excel.sheetName}!${r.excel.cell}`,
      resolutionStatus: resolveStatus(r),
      formulaConstants: extractConstants(r.excel.formula),
      tracePath: r.excel.referencedCells.join(' → '),
      note: r.summary,
    };
  });
}
