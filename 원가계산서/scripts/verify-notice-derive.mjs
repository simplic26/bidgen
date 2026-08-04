// 공고문 자동 유도 규칙 단위 검증 (견본: 판교지사 약품주입설비 개선공사 공고)
//   node scripts/verify-notice-derive.mjs
// Node 23.6+ 타입 스트리핑으로 src의 .ts를 직접 임포트한다.
import {
  deriveAValue,
  deriveBasePrice,
  deriveEstimatedTotal,
  deriveOfficeStrings,
  resolveNoticeValues,
} from '../src/utils/noticeDerive.ts';

let failed = 0;
function check(name, actual, expected) {
  const ok = JSON.stringify(actual) === JSON.stringify(expected);
  console.log(`${ok ? '✓' : '✗'} ${name}${ok ? '' : ` — 기대 ${JSON.stringify(expected)}, 실제 ${JSON.stringify(actual)}`}`);
  if (!ok) failed += 1;
}

// 금액 유도 (견본 검증값)
check("예비가격기초금액 = 296,380,000 × 1.1", deriveBasePrice('296380000'), '326018000');
check('예비가격기초금액: 콤마 입력 허용', deriveBasePrice('296,380,000'), '326018000');
check('예비가격기초금액: 빈 입력 → 빈 값', deriveBasePrice(''), '');
check('추정금액 (사급자재비 0)', deriveEstimatedTotal('296380000', '0'), '326018000');
check('추정금액 (사급자재비 1,000,000)', deriveEstimatedTotal('296380000', '1000000'), '327118000');

// A값 = 보험료 5종 합 (견본: 15,715,065)
const SAMPLE_INSURANCE = {
  국민건강보험료: '2723659',
  국민연금보험료: '3598715',
  노인장기요양보험료: '357888',
  퇴직공제부금비: '1742535',
  산업안전보건관리비: '7292268',
};
check('A값 = 보험료 5종 합', deriveAValue(SAMPLE_INSURANCE), '15715065');
check('A값: 하나라도 빈 값이면 미계산', deriveAValue({ ...SAMPLE_INSURANCE, 국민연금보험료: '' }), '');

// 사업소 문자열 조합 (견본: 판교지사)
check('사업소 유도 문자열', deriveOfficeStrings('판교지사'), {
  개찰장소: '판교지사 고객지원부 입찰집행관 PC',
  공고명의: '한국지역난방공사 판교지사장',
  문의부서1: '판교지사 고객지원부',
});
check('사업소 미선택 → 빈 값', deriveOfficeStrings(''), { 개찰장소: '', 공고명의: '', 문의부서1: '' });

// resolveNoticeValues: 오버라이드 캐스케이드
const form = {
  office: '판교지사',
  estimatedPrice: '296380000',
  ownerMaterialCost: '0',
};
const prefill = {
  procurementType: 'CONSTRUCTION',
  insurances: {
    국민건강보험료: 2723659,
    국민연금보험료: 3598715,
    노인장기요양보험료: 357888,
    퇴직공제부금비: 1742535,
    산업안전보건관리비: 7292268,
  },
  netConstructionCost: 253590813,
};

const base = resolveNoticeValues(form, prefill, {});
check('연동: 보험료 → A값 자동합계', base.values.A값, '15715065');
check('연동: 보험료 출처 = linked', base.sources.국민건강보험료, 'linked');
check('연동: 순공사원가', base.values.순공사원가, '253590813');
check('유도: 예비가격기초금액 출처 = auto', base.sources.예비가격기초금액, 'auto');

// 보험료 하나 수정 → A값이 유효값 기준으로 재계산돼야 함
const edited = resolveNoticeValues(form, prefill, { 국민건강보험료: '3000000' });
check('오버라이드: 보험료 수정 시 A값 재계산', edited.values.A값, String(3000000 + 3598715 + 357888 + 1742535 + 7292268));
check('오버라이드: 수정 필드 출처 = edited', edited.sources.국민건강보험료, 'edited');

// A값 자체를 오버라이드하면 그 값이 우선
const aOverride = resolveNoticeValues(form, prefill, { A값: '99999999' });
check('오버라이드: A값 직접 수정 우선', aOverride.values.A값, '99999999');

console.log(failed ? `\n실패 ${failed}건` : '\n유도 규칙 검증 전부 통과');
process.exit(failed ? 1 : 0);
