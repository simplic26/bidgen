import type { DerivedKey, InsuranceKey, NoticeForm, NoticeOverrides, NoticePrefill, ValueSource } from '../types';

// ---------------------------------------------------------------------------
// 공고문 자동 유도 엔진 (순수 함수 — DOM 비의존, node 검증 스크립트에서 재사용)
//
// 유도값은 폼 상태에 저장하지 않는다. 매 렌더마다 resolveNoticeValues로 재계산하고,
// 사용자가 수정하면 NoticeOverrides에 기록 → 유효값 = overrides[key] ?? derived[key].
// "↺ 자동값" 리셋은 overrides에서 키를 지우는 것.
// ---------------------------------------------------------------------------

export const INSURANCE_KEYS: InsuranceKey[] = [
  '국민건강보험료',
  '국민연금보험료',
  '노인장기요양보험료',
  '퇴직공제부금비',
  '산업안전보건관리비',
];

const DEFAULT_AGENCY = '한국지역난방공사';

function toNumber(raw: string): number | null {
  const digits = raw.replace(/[^\d]/g, '');
  if (!digits) return null;
  return Number(digits);
}

/** 예비가격기초금액 = round(추정가격 × 1.1) — 부가세 포함. 추정가격이 숫자가 아니면 ''. */
export function deriveBasePrice(estimatedPrice: string): string {
  const est = toNumber(estimatedPrice);
  if (est == null) return '';
  return String(Math.round(est * 1.1));
}

/** 추정금액 = round(추정가격×1.1) + round(사급자재비×1.1). 사급자재비 미입력은 0으로 간주. */
export function deriveEstimatedTotal(estimatedPrice: string, ownerMaterialCost: string): string {
  const est = toNumber(estimatedPrice);
  if (est == null) return '';
  const material = toNumber(ownerMaterialCost) ?? 0;
  return String(Math.round(est * 1.1) + Math.round(material * 1.1));
}

/** A값 = 보험료 5종 합계. 5종이 전부 숫자일 때만 계산, 아니면 ''. */
export function deriveAValue(insurances: Record<InsuranceKey, string>): string {
  let sum = 0;
  for (const key of INSURANCE_KEYS) {
    const value = toNumber(insurances[key] ?? '');
    if (value == null) return '';
    sum += value;
  }
  return String(sum);
}

/** 사업소 선택값으로 유도되는 문자열들. 사업소 미선택 시 빈 값. */
export function deriveOfficeStrings(office: string): { 개찰장소: string; 공고명의: string; 문의부서1: string } {
  if (!office) return { 개찰장소: '', 공고명의: '', 문의부서1: '' };
  return {
    개찰장소: `${office} 고객지원부 입찰집행관 PC`,
    공고명의: `${DEFAULT_AGENCY} ${office}장`,
    문의부서1: `${office} 고객지원부`,
  };
}

export interface ResolvedNotice {
  values: Record<DerivedKey, string>;
  sources: Record<DerivedKey, ValueSource>;
}

/**
 * 유도값 일괄 계산. 2-pass:
 *  1) 규칙 유도(auto) + 원가계산서 연동(linked) 기본값 계산
 *  2) 보험료 5종에 오버라이드 반영 후 그 "유효값"으로 A값 계산 → 나머지 오버라이드 적용(edited)
 * 따라서 사용자가 보험료 하나를 고치면 A값이 자동으로 따라온다 (A값 자체를 고치지 않는 한).
 */
export function resolveNoticeValues(form: NoticeForm, prefill: NoticePrefill, overrides: NoticeOverrides): ResolvedNotice {
  const office = deriveOfficeStrings(form.office);

  // 1-pass: 기본 유도값 + 출처
  const base: Record<DerivedKey, string> = {
    예비가격기초금액: deriveBasePrice(form.estimatedPrice),
    추정금액: deriveEstimatedTotal(form.estimatedPrice, form.ownerMaterialCost),
    A값: '', // 2-pass에서 계산
    개찰장소: office.개찰장소,
    공고명의: office.공고명의,
    문의부서1: office.문의부서1,
    발주기관: DEFAULT_AGENCY,
    국민건강보험료: prefill.insurances?.국민건강보험료 != null ? String(prefill.insurances.국민건강보험료) : '',
    국민연금보험료: prefill.insurances?.국민연금보험료 != null ? String(prefill.insurances.국민연금보험료) : '',
    노인장기요양보험료: prefill.insurances?.노인장기요양보험료 != null ? String(prefill.insurances.노인장기요양보험료) : '',
    퇴직공제부금비: prefill.insurances?.퇴직공제부금비 != null ? String(prefill.insurances.퇴직공제부금비) : '',
    산업안전보건관리비: prefill.insurances?.산업안전보건관리비 != null ? String(prefill.insurances.산업안전보건관리비) : '',
    순공사원가: prefill.netConstructionCost != null ? String(prefill.netConstructionCost) : '',
  };

  const sources = {} as Record<DerivedKey, ValueSource>;
  for (const key of Object.keys(base) as DerivedKey[]) {
    const linked =
      (INSURANCE_KEYS as readonly string[]).includes(key) ? base[key] !== '' : key === '순공사원가' && base[key] !== '';
    sources[key] = linked ? 'linked' : 'auto';
  }

  // 2-pass: 보험료 유효값(오버라이드 우선)으로 A값 계산 → 전체 오버라이드 적용
  const effectiveInsurances = {} as Record<InsuranceKey, string>;
  for (const key of INSURANCE_KEYS) {
    effectiveInsurances[key] = overrides[key] ?? base[key];
  }
  base.A값 = deriveAValue(effectiveInsurances);

  const values = { ...base };
  for (const [key, value] of Object.entries(overrides) as Array<[DerivedKey, string]>) {
    if (value == null) continue;
    values[key] = value;
    sources[key] = 'edited';
  }

  return { values, sources };
}
