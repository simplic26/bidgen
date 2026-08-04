import type { ReferenceFiles, ResultFilters, ValidationResult, ValidationStatus } from '../types';

// 입력된 기준자료 파일 개수 (제비율/노임단가/표준품셈 중 채워진 슬롯)
export function countReferenceFiles(rf: ReferenceFiles): number {
  return [rf.rateFile, rf.laborFile, rf.standardPdf].filter(Boolean).length;
}

type Evidence = ValidationResult['evidence'];

// 판단근거 위치 한 줄 표기 (PRD §13.1): 기준자료 Excel은 시트!셀·표시값, 표준품셈은 페이지
export function evidenceLocation(evidence: Evidence): string {
  if (
    (evidence.documentType === 'RATE_EXCEL' || evidence.documentType === 'LABOR_RATE_EXCEL') &&
    evidence.sheetName &&
    evidence.cell
  ) {
    const loc = `${evidence.sheetName}!${evidence.cell}`;
    return evidence.displayValue ? `${loc} · ${evidence.displayValue}` : loc;
  }
  if (evidence.documentType === 'STANDARD_PDF' && evidence.page != null) {
    return `${evidence.documentTitle} · ${evidence.page}페이지`;
  }
  return evidence.documentTitle;
}

export const statusLabels: Record<ValidationStatus, string> = {
  ERROR: '오류',
  NEEDS_REVIEW: '확인 필요',
  WARNING: '주의',
  OK: '정상',
  UNAVAILABLE: '검증 불가',
};

export const validationTypeLabels = {
  ARITHMETIC: '산술',
  TOTAL: '합계',
  RATE: '요율',
  BASE: '산출기초',
  CONDITION: '적용조건',
  FORMULA: '수식',
  REFERENCE: '참조',
  LABOR: '노임단가',
  STANDARD: '표준품셈',
} as const;

const statusOrder: Record<ValidationStatus, number> = {
  ERROR: 0,
  NEEDS_REVIEW: 1,
  WARNING: 2,
  UNAVAILABLE: 3,
  OK: 4,
};

export function getStatusCounts(results: ValidationResult[]) {
  return results.reduce(
    (counts, result) => {
      counts[result.status] += 1;
      return counts;
    },
    { ERROR: 0, NEEDS_REVIEW: 0, WARNING: 0, OK: 0, UNAVAILABLE: 0 },
  );
}

export function sortResultsByPriority(results: ValidationResult[]) {
  return [...results].sort((a, b) => {
    const statusDiff = statusOrder[a.status] - statusOrder[b.status];
    if (statusDiff !== 0) return statusDiff;
    return a.resultId.localeCompare(b.resultId);
  });
}

export function filterResults(results: ValidationResult[], filters: ResultFilters) {
  return sortResultsByPriority(results).filter((result) => {
    const matchesStatus = filters.status === 'ALL' || result.status === filters.status;
    const matchesType = filters.validationType === 'ALL' || result.validationType === filters.validationType;
    const matchesSheet = filters.sheetName === 'ALL' || result.excel.sheetName === filters.sheetName;
    return matchesStatus && matchesType && matchesSheet;
  });
}

export function getReviewResults(results: ValidationResult[]) {
  return sortResultsByPriority(results).filter((result) => result.status !== 'OK');
}

export function getAdjacentResult(results: ValidationResult[], selectedId: string, direction: 'previous' | 'next') {
  const reviewResults = getReviewResults(results);
  const index = reviewResults.findIndex((result) => result.resultId === selectedId);
  if (index === -1) return reviewResults[0];
  const nextIndex = direction === 'next' ? index + 1 : index - 1;
  return reviewResults[(nextIndex + reviewResults.length) % reviewResults.length];
}

export function formatConfidence(confidence: number) {
  if (confidence >= 0.9) return '높음';
  if (confidence >= 0.7) return '보통';
  return '낮음';
}

export function uniqueSheets(results: ValidationResult[]) {
  return Array.from(new Set(results.map((result) => result.excel.sheetName)));
}
