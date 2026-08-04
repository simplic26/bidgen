// 공유 라벨 매칭(matchCanonicalLabel / matchStructuralLabel) 회귀 테스트.
// 이 함수들은 costTree(원가계산서 파싱), criteria(제비율 기준요율 추출), validation/notice가 모두 쓰므로
// 여기서 깨지면 앱 전체가 조용히 오작동한다. 픽스처 문자열은 전부 실제 레퍼런스 파일에서 채취했다.
import assert from 'node:assert/strict';
import test from 'node:test';
import { createServer } from 'vite';

async function loadCostItems() {
  const server = await createServer({
    configFile: 'vite.config.mjs',
    configLoader: 'native',
    server: { middlewareMode: true },
  });
  try {
    const mod = await server.ssrLoadModule('/src/utils/costItems.ts');
    return { mod, close: () => server.close() };
  } catch (error) {
    await server.close();
    throw error;
  }
}

// 실제 파일 라벨 → 기대 canonical. 출처를 주석으로 남긴다.
const CANONICAL_CASES = [
  // '2. 추정가격내역서(간단).xlsx' 시트 '제비율 적용기준' A42:A51 — 대괄호 안 산출식 부기
  ['간접노무비\r\n[(직노)X요율]', '간접노무비'],
  ['산업재해보상보험료\r\n[(노)X요율]', '산재보험료'],
  ['고용보험료\r\n[(노)X요율]', '고용보험료'],
  ['건강보험료\r\n[(직노)X요율]', '국민건강보험료'],
  ['연금보험료\r\n[(직노)X요율]', '국민연금보험료'],
  ['노인장기요양보험료\r\n[(건강보험료)X요율]', '노인장기요양보험료'],
  ['산업안전보건관리비\r\n[(재+직노)X요율]', '산업안전보건관리비'],
  ['기타경비\r\n[(재+노)X요율]', '기타경비'],
  ['일반관리비\r\n[(재+노+경)X요율]', '일반관리비'],
  ['이      윤\r\n[(노+경+일)X요율]', '이윤'], // 자간 벌린 라벨 + 줄바꿈 + 대괄호 부기
  // 괄호 부기 (ref-05/ref-06/ref-08 파일)
  ['국민건강보험료(실적정산)', '국민건강보험료'],
  ['장기요양보험료(실적정산)', '노인장기요양보험료'],
  ['부가가치세(VAT)', '부가가치세'],
  ['간접노무비 (복리후생비 등)', '간접노무비'],
  // 번호·항목기호 prefix
  ['1) 산재보험료', '산재보험료'],
  ['1. 산재보험료', '산재보험료'],
  ['(2) 고용보험료', '고용보험료'],
  ['나. 간접노무비', '간접노무비'],
  ['라. 건설근로자퇴직공제부금비', '퇴직공제부금비'], // ref-07 B51
  ['① 이윤', '이윤'],
  ['ㅇ 일반관리비', '일반관리비'],
  ['■ 기타경비', '기타경비'],
  ['▶ 부가가치세', '부가가치세'],
  ['- 고용보험료', '고용보험료'],
  ['※ 환경보전비', '환경보전비'],
  // 양끝 구두점 / 숫자 꼬리
  ['일반관리비 :', '일반관리비'],
  ['이윤 15%', '이윤'],
  // 줄바꿈으로 쪼개진 이름 (ref-07 A42)
  ['건설근로자\n퇴직공제부금비', '퇴직공제부금비'],
  ['건설근로자퇴직공제부금비', '퇴직공제부금비'],
];

// 산출식·설명문·잡담 — 항목 라벨이 아니므로 매칭되면 안 된다.
const CANONICAL_NON_CASES = [
  '이윤절사',
  '부가가치세포함',
  '공급가액+부가가치세',
  '(재료비+노무비)×',
  '( 노무비 + 경비 + 일반관리비 ) ×',
  '조달청 기타경비 요율', // 간단 파일 A36 — 여기 요율을 기타경비 기준요율로 잡으면 안 된다
  '조달청 기타경비 요율에서 수도광열비 제외한 요율',
  '※ 기타경비 : 복리후생비, 소모품비, 여비교통통신비, 세금과공과, 도서인쇄비',
  '2025년 유지보수 공사 기타경비 적용기준',
  '안전관리비 대상액', // 간단 파일 D48 — 요율 셀이 아니라 대상액 열 헤더
  '근로자재해보장책임보험\r\n[(노)X요율]', // 표준 원가항목이 아니다
  '전기공사공제조합 요율적용',
  '총공사비\n(부가세포함) :', // ref-07 총괄표 P49 — 부가세 앵커 오탐의 원인이던 잡담 문구
  '이윤율',
  '',
  '   ',
];

const STRUCTURAL_CASES = [
  ['순 공 사 원 가', '순공사원가'],
  ['합 계', '총계'],
  ['재료비', '재료비'],
  ['직접노무비', '직접노무비'],
  ['노무비 (직접+간접)', '노무비'],
  ['가. 경비', '경비'],
];

const STRUCTURAL_NON_CASES = [
  '(재료비+노무비+경비)×',
  '재료비계상',
  '노무비단가',
  '총공사비',
];

test('matchCanonicalLabel: 실제 파일의 부기·번호 붙은 라벨을 항목명으로 인식한다', async () => {
  const { mod, close } = await loadCostItems();
  try {
    for (const [label, expected] of CANONICAL_CASES) {
      assert.equal(
        mod.matchCanonicalLabel(label),
        expected,
        `${JSON.stringify(label)} → ${expected} 이어야 함 (실제 ${mod.matchCanonicalLabel(label)}, strip=${JSON.stringify(mod.stripLabelDecorations(label))})`,
      );
    }
  } finally {
    await close();
  }
});

test('matchCanonicalLabel: 산출식·설명문·잡담은 항목명으로 인식하지 않는다', async () => {
  const { mod, close } = await loadCostItems();
  try {
    for (const label of CANONICAL_NON_CASES) {
      assert.equal(
        mod.matchCanonicalLabel(label),
        null,
        `${JSON.stringify(label)} 은 매칭되면 안 됨 (실제 ${mod.matchCanonicalLabel(label)}, strip=${JSON.stringify(mod.stripLabelDecorations(label))})`,
      );
    }
  } finally {
    await close();
  }
});

test('matchStructuralLabel: 트리 골격 라벨도 같은 규칙으로 판정한다', async () => {
  const { mod, close } = await loadCostItems();
  try {
    for (const [label, expected] of STRUCTURAL_CASES) {
      assert.equal(mod.matchStructuralLabel(label), expected, `${JSON.stringify(label)} → ${expected}`);
    }
    for (const label of STRUCTURAL_NON_CASES) {
      assert.equal(mod.matchStructuralLabel(label), null, `${JSON.stringify(label)} 은 매칭되면 안 됨`);
    }
  } finally {
    await close();
  }
});

test('stripLabelDecorations: 중첩 괄호·선행 번호·양끝 구두점을 제거한다', async () => {
  const { mod, close } = await loadCostItems();
  try {
    assert.equal(mod.stripLabelDecorations('이      윤\r\n[(노+경+일)X요율]'), '이윤');
    assert.equal(mod.stripLabelDecorations('1) 산재보험료'), '산재보험료');
    assert.equal(mod.stripLabelDecorations('국민건강보험료(실적정산)'), '국민건강보험료');
    assert.equal(mod.stripLabelDecorations('일반관리비 :'), '일반관리비');
    assert.equal(mod.stripLabelDecorations('(재료비+노무비)×'), '');
    // 이름 자체는 절대 훼손하지 않는다
    assert.equal(mod.stripLabelDecorations('간접노무비'), '간접노무비');
    assert.equal(mod.stripLabelDecorations('안전관리비 대상액'), '안전관리비대상액');
  } finally {
    await close();
  }
});
