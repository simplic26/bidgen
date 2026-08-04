import assert from 'node:assert/strict';
import test from 'node:test';
import { createServer } from 'vite';

async function loadModules() {
  const server = await createServer({
    configFile: 'vite.config.mjs',
    configLoader: 'native',
    server: { middlewareMode: true },
  });
  try {
    const validation = await server.ssrLoadModule('/src/utils/validation.ts');
    const costTree = await server.ssrLoadModule('/src/utils/costTree.ts');
    return { validation, costTree, close: () => server.close() };
  } catch (error) {
    await server.close();
    throw error;
  }
}

// 합성 IR 헬퍼 (costTree.test.mjs와 동일 규약). cells 항목: { a: 주소, v: 값, z?: 서식 }
function sheetIR(sheetName, cells, role = 'OTHER') {
  return {
    sheetName,
    sheetRole: role,
    rowCount: 50, columnCount: 20,
    cellCount: cells.length,
    formulaCount: 0,
    mergeCount: 0,
    cells: cells.map((c) => ({
      address: c.a,
      dataType: typeof c.v === 'number' ? 'NUMBER' : 'STRING',
      rawValue: String(c.v ?? ''),
      cachedValue: c.v ?? null,
      displayValue: String(c.v ?? ''),
      numberFormat: c.z ?? null,
      mergedRange: null,
      hidden: false,
      references: [],
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

// 요약시트 통과선(11점)을 못 넘는 시트: 법정 골격(재료비·노무비·경비·일반관리비·이윤)도 없고
// 부가세 앵커도 없어 findSummarySheet는 null을 준다 → 트리는 비고 항목 0개.
// 반면 라벨-금액이 한 행에 인접해 있어 레거시 라벨 스캔은 항목을 찾아낸다.
function looseLabelSheet(name = '내역서') {
  return sheetIR(name, [
    { a: 'A5', v: '간접노무비' }, { a: 'B5', v: 100000 },
    { a: 'A6', v: '산재보험료' }, { a: 'B6', v: 5000 },
    { a: 'A7', v: '기타경비' }, { a: 'B7', v: 3000 },
  ]);
}

test('detectItems: 트리 표준항목이 하한 미만이면 레거시 라벨 스캔 결과를 돌려준다', async () => {
  const { validation, costTree, close } = await loadModules();
  try {
    const ir = workbookIR([looseLabelSheet()]);

    // 전제: 트리 경로는 이 워크북에서 항목을 못 뽑는다 (요약시트 미발견 → 하한 미만)
    const tree = costTree.parseCostTree(ir);
    assert.equal(tree.summarySheet, null);
    const treeItems = costTree.costTreeToDetectedItems(tree);
    assert.ok(treeItems.length < costTree.TREE_ITEM_MIN, `트리 항목이 하한 미만이어야 함: ${treeItems.length}`);

    // 폴백 결과: 레거시 스캔이 찾은 3개 항목이 그대로 나와야 한다.
    const items = validation.detectItems(ir);
    assert.equal(items.length, 3, JSON.stringify(items));
    assert.deepEqual(
      items.map((i) => i.canonicalName).sort(),
      ['간접노무비', '기타경비', '산재보험료'],
    );
    const ind = items.find((i) => i.canonicalName === '간접노무비');
    assert.equal(ind.sheetName, '내역서');
    assert.equal(ind.labelCell, 'A5');   // 레거시는 라벨 셀 주소를 그대로 쓴다
    assert.equal(ind.amountCell, 'B5');
    assert.equal(ind.amountValue, 100000);
  } finally {
    await close();
  }
});

test('detectItems: 트리가 하한 이상 항목을 뽑으면 트리 결과를 쓴다(폴백 아님)', async () => {
  const { validation, costTree, close } = await loadModules();
  try {
    // 요약시트로 인정되는 시트 + 라벨만 흩어진 보조 시트를 함께 둔다.
    // 트리 경로가 채택되면 결과는 전부 요약시트 소속이어야 한다(보조 시트 라벨이 섞이면 레거시 폴백).
    const summary = sheetIR('원가계산서', [
      { a: 'A6', v: '재료비' }, { a: 'F6', v: 450000 },
      { a: 'A8', v: '노무비' }, { a: 'F8', v: 829538 },
      { a: 'B10', v: '간접노무비' }, { a: 'E10', v: 0.15, z: '0.00%' }, { a: 'F10', v: 108200 },
      { a: 'A11', v: '경비' }, { a: 'F11', v: 29531 },
      { a: 'B12', v: '산재보험료' }, { a: 'E12', v: 0.0356, z: '0.00%' }, { a: 'F12', v: 29531 },
      { a: 'A13', v: '일반관리비' }, { a: 'E13', v: 0.06, z: '0.00%' }, { a: 'F13', v: 78544 },
      { a: 'A14', v: '이윤' }, { a: 'E14', v: 0.15, z: '0.00%' }, { a: 'F14', v: 140641 },
      { a: 'A15', v: '총원가' }, { a: 'F15', v: 1528254 },
      { a: 'A16', v: '부가가치세' }, { a: 'F16', v: 152825 },
      { a: 'A17', v: '합계' }, { a: 'F17', v: 1681079 },
    ], 'COST_SUMMARY');
    const ir = workbookIR([summary, looseLabelSheet()]);

    const treeItems = costTree.costTreeToDetectedItems(costTree.parseCostTree(ir));
    assert.ok(treeItems.length >= costTree.TREE_ITEM_MIN, `트리 항목이 하한 이상이어야 함: ${treeItems.length}`);

    const items = validation.detectItems(ir);
    assert.ok(items.every((i) => i.sheetName === '원가계산서'), JSON.stringify(items.map((i) => [i.canonicalName, i.sheetName])));
    assert.deepEqual(items.map((i) => i.amountCell), treeItems.map((i) => i.amountCell));
  } finally {
    await close();
  }
});
