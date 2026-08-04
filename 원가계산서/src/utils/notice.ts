import type { FileMeta, InsuranceKey, NoticePrefill, ProcurementType, WorkbookIR } from '../types';
import { detectItems, normalizeLabel, parseKoreanNumber } from './validation';

// ---------------------------------------------------------------------------
// 공고문 생성 화면 자동 채움: 검증 플로우에서 확보한 데이터로 폼 초기값을 유도한다.
// ---------------------------------------------------------------------------

const TOTAL_KEYS = ['합계', '총계', '총액', '도급액', '총공사비'];
const NET_COST_KEYS = ['순공사원가', '순공사비'];

const INSURANCE_KEYS: InsuranceKey[] = [
  '국민건강보험료',
  '국민연금보험료',
  '노인장기요양보험료',
  '퇴직공제부금비',
  '산업안전보건관리비',
];

function parseAddress(address: string): { r: number; c: number } | null {
  const m = address.match(/^([A-Z]+)(\d+)$/);
  if (!m) return null;
  let c = 0;
  for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64);
  return { r: Number(m[2]), c };
}

/** 원가계산서 IR에서 라벨 키워드에 해당하는 금액을 찾습니다. 라벨 오른쪽 가장 오른쪽 숫자 셀 채택. */
export function findAmountByLabel(ir: WorkbookIR, keys: string[]): number | null {
  const normKeys = keys.map((k) => normalizeLabel(k));
  for (const sheet of ir.sheets) {
    // 기준자료성 시트는 제외 (제비율표·노임단가표)
    if (sheet.sheetRole === 'RATE_STANDARD' || sheet.sheetRole === 'WAGE_RATE') continue;

    const byRow = new Map<number, Array<{ c: number; value: number }>>();
    const labels: Array<{ r: number; c: number }> = [];

    for (const cell of sheet.cells) {
      const pos = parseAddress(cell.address);
      if (!pos) continue;
      if (cell.dataType === 'STRING') {
        const norm = normalizeLabel(String(cell.cachedValue ?? cell.displayValue));
        if (normKeys.some((k) => norm.includes(k))) labels.push(pos);
        continue;
      }
      const num = parseKoreanNumber(cell.cachedValue);
      if (num == null) continue;
      const arr = byRow.get(pos.r) ?? [];
      arr.push({ c: pos.c, value: num });
      byRow.set(pos.r, arr);
    }

    for (const label of labels) {
      const candidates = (byRow.get(label.r) ?? []).filter((cell) => cell.c > label.c);
      if (candidates.length === 0) continue;
      candidates.sort((a, b) => a.c - b.c);
      const amount = candidates[candidates.length - 1].value;
      if (amount > 0) return amount;
    }
  }
  return null;
}

/** 원가계산서 IR에서 합계(총액) 금액을 찾습니다. */
export function findTotalAmount(ir: WorkbookIR): number | null {
  return findAmountByLabel(ir, TOTAL_KEYS);
}

/** 업로드 파일명에서 공사명 후보를 유도합니다. (확장자·안내 접미어 제거) */
function projectNameFromFile(fileName: string): string {
  return fileName
    .replace(/\.(xlsx|xls)$/i, '')
    .replace(/[_-]?(최종|검토|수정|v?\d+)$/i, '')
    .trim();
}

/** 검증 플로우 상태에서 공고문 폼 자동 채움 값을 만듭니다. 업로드가 없으면 유형만 전달. */
export function buildNoticePrefill(
  procurementType: ProcurementType,
  excelFile: FileMeta | null,
  workbookIR: WorkbookIR | null,
): NoticePrefill {
  const prefill: NoticePrefill = { procurementType };
  if (!excelFile) return prefill;

  prefill.sourceFileName = excelFile.name;
  const name = projectNameFromFile(excelFile.name);
  if (name) prefill.projectName = name;

  if (!workbookIR) return prefill;

  // 항목 탐지 1회: 보험료 5종 + 부가세 + 일반관리비 + 이윤 (canonicalName → 첫 탐지 금액)
  const amounts = new Map<string, number>();
  for (const item of detectItems(workbookIR)) {
    if (item.amountValue == null || item.amountValue <= 0) continue;
    if (!amounts.has(item.canonicalName)) amounts.set(item.canonicalName, item.amountValue);
  }

  const insurances: NoticePrefill['insurances'] = {};
  for (const key of INSURANCE_KEYS) {
    const value = amounts.get(key);
    if (value != null) insurances[key] = value;
  }
  if (Object.keys(insurances).length > 0) prefill.insurances = insurances;

  const vat = amounts.get('부가가치세');
  if (vat != null) prefill.vatAmount = vat;

  // 추정가격 제안: 합계는 부가세 포함이므로, 부가세가 탐지되면 차감해 부가세 별도 금액으로 보정
  const total = findTotalAmount(workbookIR);
  if (total != null) prefill.totalAmount = vat != null && vat < total ? total - vat : total;

  // 순공사원가: 라벨 직접 탐지 → 폴백(합계−일반관리비−이윤−부가세, 전부 탐지된 경우만)
  const netCost = findAmountByLabel(workbookIR, NET_COST_KEYS);
  if (netCost != null) {
    prefill.netConstructionCost = netCost;
  } else if (total != null && vat != null) {
    const admin = amounts.get('일반관리비');
    const profit = amounts.get('이윤');
    if (admin != null && profit != null) {
      const fallback = total - admin - profit - vat;
      if (fallback > 0) prefill.netConstructionCost = fallback;
    }
  }

  return prefill;
}
