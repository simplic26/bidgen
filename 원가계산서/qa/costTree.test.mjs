import assert from 'node:assert/strict';
import test from 'node:test';
import { createServer } from 'vite';

async function loadCostTree() {
  const server = await createServer({
    configFile: 'vite.config.mjs',
    configLoader: 'native',
    server: { middlewareMode: true },
  });
  try {
    const mod = await server.ssrLoadModule('/src/utils/costTree.ts');
    return { mod, close: () => server.close() };
  } catch (error) {
    await server.close();
    throw error;
  }
}

// 합성 IR 헬퍼. cells 항목: { a: 주소, v: 값, f?: 수식(= 제외), refs?: 참조배열, z?: 서식 }
function sheetIR(sheetName, cells, role = 'OTHER') {
  return {
    sheetName,
    sheetRole: role,
    rowCount: 50, columnCount: 20,
    cellCount: cells.length,
    formulaCount: cells.filter((c) => c.f).length,
    mergeCount: 0,
    cells: cells.map((c) => ({
      address: c.a,
      dataType: c.f ? 'FORMULA' : typeof c.v === 'number' ? 'NUMBER' : 'STRING',
      rawValue: c.f ? `=${c.f}` : String(c.v ?? ''),
      cachedValue: c.v ?? null,
      displayValue: String(c.v ?? ''),
      numberFormat: c.z ?? null,
      mergedRange: null,
      hidden: false,
      references: c.refs ?? [],
    })),
  };
}

function workbookIR(sheets) {
  return {
    schemaVersion: '1.0', fileName: 'test.xlsx', procurementType: 'CONSTRUCTION',
    generatedAt: '2026-08-04T00:00:00Z', sheets,
    totals: { sheetCount: sheets.length, cellCount: 0, formulaCount: 0, mergeCount: 0 },
  };
}

// 요약시트 표준 픽스처 (간단 파일 축약형, 수식 살아있음). 검산 일치 숫자:
// 재료비 450000 = 직재 450000 / 노무비 829538 = 직노 721338 + 간노 INT(721338×0.15)=108200
// 경비 29531 = 산재 INT(829538×0.0356)=29531 / 일반관리비 INT((450000+829538+29531)×0.06)=78544
// 이윤 INT((829538+29531+78544)×0.15)=140641 / 총원가 1528254 / 부가세 INT(×0.1)=152825 / 합계 1681079
function summarySheetFixture(name = '원가계산서') {
  return sheetIR(name, [
    { a: 'A6', v: '재료비' }, { a: 'F6', v: 450000, f: 'F7', refs: ['F7'] },
    { a: 'B7', v: '직접재료비' }, { a: 'F7', v: 450000 },
    { a: 'A8', v: '노무비' }, { a: 'F8', v: 829538, f: 'F9+F10', refs: ['F9', 'F10'] },
    { a: 'B9', v: '직접노무비' }, { a: 'F9', v: 721338 },
    { a: 'B10', v: '간접노무비' }, { a: 'E10', v: 0.15, z: '0.00%' },
    { a: 'F10', v: 108200, f: 'INT(F9*E10)', refs: ['F9', 'E10'] },
    { a: 'A11', v: '경비' }, { a: 'F11', v: 29531, f: 'F12', refs: ['F12'] },
    { a: 'B12', v: '산재보험료' }, { a: 'E12', v: 0.0356, z: '0.00%' },
    { a: 'F12', v: 29531, f: 'INT(F8*E12)', refs: ['F8', 'E12'] },
    { a: 'A13', v: '일반관리비' }, { a: 'E13', v: 0.06, z: '0.00%' },
    { a: 'F13', v: 78544, f: 'INT((F6+F8+F11)*E13)', refs: ['F6', 'F8', 'F11', 'E13'] },
    { a: 'A14', v: '이윤' }, { a: 'E14', v: 0.15, z: '0.00%' },
    { a: 'F14', v: 140641, f: 'INT((F8+F11+F13)*E14)', refs: ['F8', 'F11', 'F13', 'E14'] },
    { a: 'A15', v: '총원가' }, { a: 'F15', v: 1528254, f: 'F6+F8+F11+F13+F14', refs: ['F6', 'F8', 'F11', 'F13', 'F14'] },
    { a: 'A16', v: '부가가치세' }, { a: 'F16', v: 152825, f: 'INT(F15*0.1)', refs: ['F15'] },
    { a: 'A17', v: '합계' }, { a: 'F17', v: 1681079, f: 'F15+F16', refs: ['F15', 'F16'] },
  ]);
}

// 요율 기준표 픽스처: 라벨은 겹치지만 원가 트리가 없는 시트 (ref-07 '기타경비요율' 오탐 재현)
function rateTableFixture(name = '기타경비요율') {
  return sheetIR(name, [
    { a: 'A4', v: '항 목' }, { a: 'B4', v: '공사규모' }, { a: 'C4', v: '공사기간' },
    { a: 'G4', v: '항목별 비율(%)' },
    { a: 'A5', v: '간접노무비' }, { a: 'F5', v: 0.15, z: '0.00%' },
    { a: 'A6', v: '기타경비' }, { a: 'F6', v: 0.05, z: '0.00%' },
    { a: 'A7', v: '일반관리비' }, { a: 'F7', v: 0.06, z: '0.00%' },
  ]);
}

// 라벨-금액 사이에 다른 canonical 항목명(기초금액 설명, 예: "이윤" 행의 근거란에 나오는 "일반관리비")이
// 끼는 레이아웃 재현(실제 파일 4종에서 발견된 버그 패턴). N5는 dominantNumericColumn이 N열을 고르도록
// 하는 패딩용 숫자 셀(J열보다 N열 후보가 많아지도록)일 뿐 라벨과 무관하다.
function labelHijackFixture(name = '원가계산서') {
  return sheetIR(
    name,
    [
      { a: 'N5', v: 999999 },
      { a: 'A20', v: '이      윤' },
      { a: 'C20', v: '( 노무비' },
      { a: 'I20', v: '일반관리비' },
      { a: 'J20', v: 1000000 },
      { a: 'L20', v: 0.1, z: '0.0%' },
      { a: 'N20', v: 100000, f: 'INT(J20*L20)', refs: ['J20', 'L20'] },
    ],
    'COST_SUMMARY',
  );
}

// ref-07 '기타경비요율' 시트 재현(강화판): 조달청 제비율표라서 canonical 라벨은 잔뜩 있지만
// 원가 트리는 없다. 리치니스 무제한 가산 시절엔 이런 시트가 진짜 원가계산서를 이길 수 있었다.
// 실제 시트(4,312셀, core 4 / detail 6)보다 라벨을 더 많이 실어 "최악의 적"으로 만든다.
function realRateTableFixture(name = '기타경비요율') {
  const labels = [
    '■ 기타 요율   :', '조 달 청 발 표 요 율', '공 사 기 간', '공 사 규 모', '항목별 비율(%)', '적용기준',
    '간 접 노 무 비', '기 타 경 비', '일 반 관 리 비', '이      윤', '산재보험료', '고용보험료',
    '국민건강보험료', '노인장기요양보험료', '산업안전보건관리비', '환경보전비', '순 공 사 원 가',
  ];
  const cells = labels.map((v, i) => ({ a: `A${i + 5}`, v }));
  labels.forEach((_, i) => cells.push({ a: `F${i + 5}`, v: 0.05 + i / 1000, z: '0.00%' }));
  const sheet = sheetIR(name, cells);
  sheet.cellCount = 4312; // 실제 시트 셀 수 — 대형표 패널티가 걸린다
  return sheet;
}

test('findSummarySheet: 원가 트리 없는 실제 요율 기준표 단독 워크북은 통과선을 넘지 못한다', async () => {
  const { mod, close } = await loadCostTree();
  try {
    const ir = workbookIR([realRateTableFixture()]);
    const [scored] = mod.scoreSummarySheets(ir);
    assert.equal(mod.findSummarySheet(ir), null, `요율표가 요약시트로 뽑히면 안 됨: ${JSON.stringify(scored)}`);
    // 진짜 원가계산서와의 격차도 확인 — 라벨만 겹치는 표는 크게 낮아야 한다.
    const both = mod.scoreSummarySheets(workbookIR([realRateTableFixture(), summarySheetFixture()]));
    const byName = new Map(both.map((s) => [s.sheetName, s.score]));
    assert.ok(
      byName.get('원가계산서') - byName.get('기타경비요율') >= 8,
      `격차가 너무 작다: ${JSON.stringify(both)}`,
    );
  } finally {
    await close();
  }
});

// 같은 행 기초금액 셀의 canonicalName 억제는 "값이 다를 때"만 걸려야 한다.
// F12 = INT(G12)로 G12(29,531.7 원시값)를 절사해 적은 passthrough — 값이 절사 오차 범위에서 같으므로
// G12의 라벨(산재보험료)은 유지돼야 한다(ref-09 파일의 D열=SUM(E열) 이중 표기와 같은 구조).
function sameRowPassthroughFixture() {
  const sheet = summarySheetFixture();
  sheet.cells = sheet.cells.map((c) =>
    c.address === 'F12' ? { ...c, rawValue: '=INT(G12)', references: ['G12'] } : c,
  );
  sheet.cells.push({
    address: 'G12', dataType: 'NUMBER', rawValue: '29531.7', cachedValue: 29531.7,
    displayValue: '29531.7', numberFormat: null, mergedRange: null, hidden: false, references: [],
  });
  return sheet;
}

test('buildNode: 같은 행 기초금액이 절사 오차 범위에서 같은 값이면 canonicalName을 유지한다', async () => {
  const { mod, close } = await loadCostTree();
  try {
    const tree = mod.parseCostTree(workbookIR([sameRowPassthroughFixture()]));
    const items = mod.costTreeToDetectedItems(tree).filter((i) => i.canonicalName === '산재보험료');
    // F12(절사값)와 G12(원시값) 둘 다 정당한 산재보험료 표기 — 억제되면 안 된다.
    assert.equal(items.length, 2, `F12·G12 둘 다 남아야 함: ${JSON.stringify(items)}`);
    assert.deepEqual(items.map((i) => i.amountCell).sort(), ['F12', 'G12']);
    assert.deepEqual(items.map((i) => i.amountValue).sort((a, b) => a - b), [29531, 29531.7]);
  } finally {
    await close();
  }
});

test('buildNode: 라벨-금액 사이 기초 항목명(다른 canonical)이 있어도 진짜 라벨을 고른다(leftmost canonical 우선)', async () => {
  const { mod, close } = await loadCostTree();
  try {
    const tree = mod.parseCostTree(workbookIR([labelHijackFixture()]));
    assert.equal(tree.summarySheet, '원가계산서');
    const items = mod.costTreeToDetectedItems(tree);
    const byName = new Map(items.map((i) => [i.canonicalName, i]));
    assert.ok(byName.has('이윤'), `이윤 항목이 검출되어야 함: ${JSON.stringify(items)}`);
    assert.equal(byName.get('이윤').amountValue, 100000);
    assert.equal(byName.get('이윤').rateValue, 0.1);
    // 근거란의 "일반관리비"(기초금액 설명, I20)나 그 기초값(J20=1,000,000)이 별도 canonical 항목으로
    // 오탐되면 안 된다 — 같은 행 + 값이 부모와 다른 기초금액 셀은 canonicalName이 제거되어야 한다.
    assert.ok(!byName.has('일반관리비'), `일반관리비가 오탐되면 안 됨: ${JSON.stringify(items)}`);
  } finally {
    await close();
  }
});

// 거울 셀(display mirror) 픽스처: 실제 파일(2. 추정가격내역서_최종(복잡).xlsx, 추정가격내역서(견적)
// 시트 행8) 재현. 간접노무비 행의 기초금액 칸(E8)이 별도 계산 없이 직접노무비 행(O7)의 금액을
// 그대로 재참조하는 "표시용 거울 셀"(순수 단일참조 수식 =O7)이라, 예전 파서는 이를 별도 자식
// 노드로 만들어 라벨이 중복 노출됐다. summarySheetFixture는 다른 테스트가 공유하므로 손대지 않고
// 별도 전용 픽스처로 둔다. 금액 열은 O(15) — 부가세 앵커(O13, INT(O12*0.1))로 확정된다.
function mirrorCellFixture(name = '원가계산서') {
  return sheetIR(
    name,
    [
      { a: 'A6', v: '재료비' }, { a: 'O6', v: 450000 },
      { a: 'A7', v: '직접노무비' }, { a: 'O7', v: 167250000 },
      { a: 'A8', v: '간접노무비' },
      { a: 'E8', v: 167250000, f: 'O7', refs: ['O7'] }, // 거울 셀: 순수 단일참조 =O7
      { a: 'M8', v: 0.15, z: '0.00%' },
      { a: 'O8', v: 25088000, f: 'ROUND(E8*M8,-3)', refs: ['E8', 'M8'] },
      { a: 'A9', v: '경비' }, { a: 'O9', v: 29531 },
      { a: 'A10', v: '일반관리비' }, { a: 'O10', v: 78544 },
      { a: 'A11', v: '이윤' }, { a: 'O11', v: 140641 },
      { a: 'A12', v: '총원가' }, { a: 'O12', v: 891134 },
      { a: 'A13', v: '부가가치세' }, { a: 'O13', v: 89113, f: 'INT(O12*0.1)', refs: ['O12'] },
    ],
    'COST_SUMMARY',
  );
}

test('buildNode: 같은 행 거울 셀(순수 단일참조 수식)은 자식 노드를 만들지 않고 baseCells에 원본을 기록한다', async () => {
  const { mod, close } = await loadCostTree();
  try {
    const tree = mod.parseCostTree(workbookIR([mirrorCellFixture()]));
    assert.equal(tree.summarySheet, '원가계산서');
    const flat = mod.flattenTree(tree);

    // 거울 셀(E8) 자체가 별도 자식 노드로 만들어지면 안 된다 — "간접노무비" 라벨 중복 노출 버그 재현.
    assert.ok(!flat.some((n) => n.amountCell === 'E8'), `E8이 노드로 남으면 안 됨: ${JSON.stringify(flat.map((n) => n.amountCell))}`);

    const indirect = flat.find((n) => n.amountCell === 'O8');
    assert.ok(indirect, `O8 노드가 있어야 함: ${JSON.stringify(flat.map((n) => n.amountCell))}`);
    assert.ok(!indirect.children.some((c) => c.amountCell === 'E8'), 'O8의 자식에 E8이 있으면 안 됨');
    assert.deepEqual(indirect.baseCells, ['O7'], `baseCells는 거울 체인의 최종 타깃이어야 함: ${JSON.stringify(indirect.baseCells)}`);
    assert.equal(indirect.rate, 0.15);
    assert.equal(indirect.baseAmount, 167250000, 'baseAmount는 O7의 값이어야 함');
    assert.equal(indirect.baseLabel, '직접노무비', 'baseLabel은 O7이 속한 행(직접노무비)의 라벨이어야 함');
  } finally {
    await close();
  }
});

test('findSummarySheet: 요약시트를 선정하고 요율 기준표는 배제한다', async () => {
  const { mod, close } = await loadCostTree();
  try {
    const ir = workbookIR([rateTableFixture(), summarySheetFixture()]);
    assert.equal(mod.findSummarySheet(ir), '원가계산서');
  } finally {
    await close();
  }
});

test('findSummarySheet: 요약시트가 없으면 null', async () => {
  const { mod, close } = await loadCostTree();
  try {
    const ir = workbookIR([rateTableFixture()]);
    assert.equal(mod.findSummarySheet(ir), null);
  } finally {
    await close();
  }
});

test('decomposeFormula: INT(base×rate) 꼴에서 요율·기초·절사를 분리한다', async () => {
  const { mod, close } = await loadCostTree();
  try {
    const shape = mod.decomposeFormula('=INT(F9*E10)', ['F9', 'E10'], (ref) => ref === 'E10');
    assert.deepEqual(shape.rateRefs, ['E10']);
    assert.deepEqual(shape.baseRefs, ['F9']);
    assert.equal(shape.rounding, 'INT');

    const r = mod.decomposeFormula('=ROUND(E8*M8,-3)', ['E8', 'M8'], (ref) => ref === 'M8');
    assert.equal(r.rounding, 'ROUND');
    assert.equal(r.roundDigits, -3);

    const range = mod.decomposeFormula('=SUM(F6:F8)', ['F6:F8'], () => false);
    assert.deepEqual(range.baseRefs, ['F6', 'F7', 'F8']);

    const vat = mod.decomposeFormula('=INT(F15*0.1)', ['F15'], () => false);
    assert.equal(vat.literalRate, 0.1);

    const cross = mod.decomposeFormula("='내역서'!F26", ["'내역서'!F26"], () => false);
    assert.deepEqual(cross.baseRefs, ['내역서!F26']);
  } finally {
    await close();
  }
});

// 실제 IR 추출 결과 재현: REFERENCE_PATTERN은 콜론을 버려서 SUM(D13:D25) 같은 범위 수식도
// references 배열에는 양끝 셀(D13, D25)만 남는다. decomposeFormula는 수식 본문을 재스캔해
// 범위를 복원해야 한다 — 그러지 못하면 13개 자식이 2개(D13, D25)로 잘못 줄어든다(ref-05 설계내역서
// 소계 행 재현 버그).
test('decomposeFormula: 범위 수식(SUM(D13:D25))은 IR의 bare 양끝 참조만으로도 전체 범위로 전개된다', async () => {
  const { mod, close } = await loadCostTree();
  try {
    const shape = mod.decomposeFormula('=SUM(D13:D25)', ['D13', 'D25'], () => false);
    const expected = Array.from({ length: 13 }, (_, i) => `D${13 + i}`);
    assert.deepEqual(shape.baseRefs, expected);
    assert.equal(shape.baseRefs.length, 13);

    // 시트간 범위: '내역서'!F6:F8 — bare 양끝도 시트 접두어를 달고 오는 경우.
    const cross = mod.decomposeFormula("=SUM('내역서'!F6:F8)", ["'내역서'!F6", "'내역서'!F8"], () => false);
    assert.deepEqual(cross.baseRefs, ['내역서!F6', '내역서!F7', '내역서!F8']);
  } finally {
    await close();
  }
});

test('inferSumRange: 같은 열 연속 구간 합을 위/아래 방향으로 찾는다', async () => {
  const { mod, close } = await loadCostTree();
  try {
    const col = [
      { address: 'F6', row: 6, col: 6, value: 450000 },
      { address: 'F8', row: 8, col: 6, value: 829538 },
      { address: 'F11', row: 11, col: 6, value: 29531 },
      { address: 'F13', row: 13, col: 6, value: 78544 },
      { address: 'F14', row: 14, col: 6, value: 140641 },
    ];
    const target = { address: 'F15', row: 15, col: 6, value: 1528254 }; // 위쪽 5개 합
    const found = mod.inferSumRange(target, col);
    assert.deepEqual(found.map((c) => c.address).sort(), ['F11', 'F13', 'F14', 'F6', 'F8']);
    // 아래 방향: target이 구간 위에 있는 경우
    const below = mod.inferSumRange(
      { address: 'F5', row: 5, col: 6, value: 1279538 },
      [{ address: 'F6', row: 6, col: 6, value: 450000 }, { address: 'F8', row: 8, col: 6, value: 829538 }],
    );
    assert.deepEqual(below.map((c) => c.address).sort(), ['F6', 'F8']);
    // 일치하는 구간이 없으면 null
    assert.equal(mod.inferSumRange({ ...target, value: 999 }, col), null);
  } finally {
    await close();
  }
});

test('matchesWithRounding: 천단위 절사 창은 실제 함수 정의역으로 제한된다 (금액 0 제외)', async () => {
  const { mod, close } = await loadCostTree();
  try {
    // (a) actual === 0 은 어떤 값과도 매칭되면 안 된다 — 0 % 1000 === 0 이라 옛 조건에서는
    //     "금액 0 행 = base × rate" 라는 거짓 요율 간선이 붙었다.
    assert.equal(mod.matchesWithRounding(800, 0), null);
    assert.equal(mod.matchesWithRounding(-800, 0), null);
    assert.equal(mod.matchesWithRounding(0, 0), 'EXACT'); // 0 === 0 은 EXACT 경로
    // (b) 아래 방향은 1000 미만까지(절사): expected - actual = 999
    assert.equal(mod.matchesWithRounding(1_234_999, 1_234_000), 'ROUNDDOWN');
    assert.equal(mod.matchesWithRounding(1_235_000, 1_234_000), null); // 차이 1000 = 다른 값
    // (c) 위 방향은 ROUND(x,-3) 올림 몫(−500)까지만
    assert.equal(mod.matchesWithRounding(1_233_500, 1_234_000), 'ROUNDDOWN');
    assert.equal(mod.matchesWithRounding(1_233_499, 1_234_000), null); // 차이 −501
    // (d) 천단위가 아닌 actual은 이 분기에 들어오지 않는다
    assert.equal(mod.matchesWithRounding(1_234_800, 1_234_500), null);
  } finally {
    await close();
  }
});

test('inferRateEdge: base×rate(절사 포함) 관계를 찾는다', async () => {
  const { mod, close } = await loadCostTree();
  try {
    const bases = [
      { address: 'F7', row: 7, col: 6, value: 450000 },
      { address: 'F9', row: 9, col: 6, value: 721338 },
    ];
    const rates = [{ address: 'E10', row: 10, col: 5, value: 0.15 }];
    // INT(721338×0.15)=108200
    const hit = mod.inferRateEdge({ address: 'F10', row: 10, col: 6, value: 108200 }, bases, rates);
    assert.equal(hit.base.address, 'F9');
    assert.equal(hit.rate.address, 'E10');
    assert.equal(hit.rounding, 'INT');
    // 어떤 조합으로도 안 맞으면 null
    assert.equal(mod.inferRateEdge({ address: 'F10', row: 10, col: 6, value: 123456 }, bases, rates), null);
  } finally {
    await close();
  }
});

test('parseCostTree: 수식 트리 복원 (요율·기초·검산·오탐 차단)', async () => {
  const { mod, close } = await loadCostTree();
  try {
    const ir = workbookIR([rateTableFixture(), summarySheetFixture()]);
    const tree = mod.parseCostTree(ir);
    assert.equal(tree.summarySheet, '원가계산서');
    assert.equal(tree.amountColumn, 'F');
    const flat = mod.flattenTree(tree);
    const byName = new Map(flat.map((n) => [n.canonicalName, n]));
    assert.equal(byName.get('총계').amount, 1681079);
    assert.equal(byName.get('부가가치세').rate, 0.1);
    const ind = byName.get('간접노무비');
    assert.equal(ind.amount, 108200);
    assert.equal(ind.rate, 0.15);
    assert.equal(ind.rateCell, 'E10');
    assert.equal(ind.status, 'FORMULA_VERIFIED');
    // 기준표 시트의 '간접노무비'(기타경비요율!A5)는 트리에 없어야 함 (오탐 차단)
    assert.ok(flat.every((n) => n.id.startsWith('원가계산서!')));
  } finally {
    await close();
  }
});

test('parseCostTree: 수식 없는(값 붙여넣기) 파일은 값-검산으로 복원', async () => {
  const { mod, close } = await loadCostTree();
  try {
    // 픽스처에서 수식 제거 → 순수 값 (요율 셀 0.15/0.0356/0.06은 숫자로 남음)
    const dead = summarySheetFixture();
    dead.cells = dead.cells.map((c) => ({
      ...c,
      dataType: c.dataType === 'FORMULA' ? 'NUMBER' : c.dataType,
      rawValue: String(c.cachedValue ?? ''),
      references: [],
    }));
    dead.formulaCount = 0;
    const tree = mod.parseCostTree(workbookIR([dead]));
    assert.equal(tree.amountColumn, 'F');
    const flat = mod.flattenTree(tree);
    const byName = new Map(flat.map((n) => [n.canonicalName, n]));
    assert.equal(byName.get('총계').amount, 1681079);
    assert.equal(byName.get('총계').status, 'VALUE_INFERRED');   // 총원가+부가세 골격 합
    assert.equal(byName.get('총원가').status, 'VALUE_INFERRED'); // 재+노+경+일관+이윤 골격 합
    const ind = byName.get('간접노무비');
    assert.equal(ind.status, 'VALUE_INFERRED'); // INT(직노×0.15) 요율 간선
    assert.equal(ind.rate, 0.15);
    assert.equal(byName.get('일반관리비').status, 'VALUE_INFERRED'); // 조합 기초 (재+노+경)×0.06
    assert.equal(byName.get('부가가치세').rate, 0.1); // 총원가×10% 특례
  } finally {
    await close();
  }
});

test('parseCostTree: 금액 셀이 비어있는 CANONICAL 행도 UNRESOLVED 루트로 표면화한다', async () => {
  const { mod, close } = await loadCostTree();
  try {
    // 라벨은 있으나 금액 열(F)에 셀이 없는 행 — 조용히 버리면 "못 읽었다"는 사실이 화면에서 사라진다.
    const sheet = summarySheetFixture();
    sheet.cells.push({
      address: 'A20', dataType: 'STRING', rawValue: '환경보전비', cachedValue: '환경보전비',
      displayValue: '환경보전비', numberFormat: null, mergedRange: null, hidden: false, references: [],
    });
    const tree = mod.parseCostTree(workbookIR([sheet]));
    assert.equal(tree.amountColumn, 'F');
    const node = tree.roots.find((n) => n.canonicalName === '환경보전비');
    assert.ok(node, `환경보전비 행이 표면화되어야 함: ${JSON.stringify(tree.roots.map((n) => n.canonicalName))}`);
    assert.equal(node.amount, null);
    assert.equal(node.status, 'UNRESOLVED');
    assert.equal(node.amountCell, 'F20');
    assert.equal(node.labelCell, 'A20');
    assert.match(node.note, /금액 셀 미발견 \(금액 열 F\)/);
    // 금액이 null이므로 DetectedItem 계약에는 새지 않는다 — 검증·골든 집계는 영향받지 않는다.
    const items = mod.costTreeToDetectedItems(tree);
    assert.ok(!items.some((i) => i.canonicalName === '환경보전비'), JSON.stringify(items));
  } finally {
    await close();
  }
});

test('parseCostTree: 요약시트 미발견 시 빈 트리 + issue (throw 금지)', async () => {
  const { mod, close } = await loadCostTree();
  try {
    const tree = mod.parseCostTree(workbookIR([rateTableFixture()]));
    assert.equal(tree.summarySheet, null);
    assert.equal(tree.roots.length, 0);
    assert.ok(tree.issues.length >= 1);
  } finally {
    await close();
  }
});

test('costTreeToDetectedItems: CANONICAL 항목만 DetectedItem으로 변환', async () => {
  const { mod, close } = await loadCostTree();
  try {
    const tree = mod.parseCostTree(workbookIR([summarySheetFixture()]));
    const items = mod.costTreeToDetectedItems(tree);
    const names = items.map((i) => i.canonicalName);
    assert.ok(names.includes('간접노무비'));
    assert.ok(names.includes('산재보험료'));
    assert.ok(names.includes('부가가치세'));
    assert.ok(names.includes('일반관리비'));
    assert.ok(names.includes('이윤'));
    assert.ok(!names.includes('총계'));   // STRUCTURAL 골격은 제외
    assert.ok(!names.includes('재료비'));
    const ind = items.find((i) => i.canonicalName === '간접노무비');
    assert.equal(ind.sheetName, '원가계산서');
    assert.equal(ind.amountCell, 'F10');
    assert.equal(ind.rateCell, 'E10');
    assert.equal(ind.rateValue, 0.15);
    assert.equal(ind.requiresReference, true);
  } finally {
    await close();
  }
});
