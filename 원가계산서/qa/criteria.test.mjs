import assert from 'node:assert/strict';
import { existsSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import test from 'node:test';
import { createServer } from 'vite';
import * as XLSX from 'xlsx';

async function loadCriteriaModule() {
  const server = await createServer({
    configFile: 'vite.config.mjs',
    configLoader: 'native',
    server: { middlewareMode: true },
  });

  try {
    const mod = await server.ssrLoadModule('/src/utils/criteria.ts');
    return { mod, close: () => server.close() };
  } catch (error) {
    await server.close();
    throw error;
  }
}

function workbookFileFromSheet(sheetName, worksheet) {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
  const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });
  return new File([buffer], 'rate-standard.xlsx', {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}

test('extracts a rate from the full merged header span in wide rate tables', async () => {
  const { mod, close } = await loadCriteriaModule();
  try {
    const rows = [
      ['(재+직노) x 율 + 기초액', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '(직노) x 율'],
      [],
      [],
      ['구분', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '요율', '', '', '', '', '건축'],
      ['건축공사', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 3.11, '', '', '', '', 17.5],
    ];
    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    worksheet['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 2, c: 20 } },
      { s: { r: 0, c: 22 }, e: { r: 2, c: 24 } },
    ];

    const file = workbookFileFromSheet('건축제비율', worksheet);
    const { criteria } = await mod.parseRateWorkbook(file, 'CONSTRUCTION');

    const safety = criteria.find((criterion) => criterion.canonicalName === '산업안전보건관리비');
    assert.ok(safety, `expected 산업안전보건관리비 in ${JSON.stringify(criteria)}`);
    assert.equal(safety.cell, 'R5');
    assert.equal(safety.rate, 0.0311);
  } finally {
    await close();
  }
});

// 실제 제비율 시트 회귀: 공유 라벨 매칭이 죽으면 기준요율 추출이 조용히 0건이 된다.
// (라벨 길이 예산(MAX_LABEL_SLACK) 방식이 "산업재해보상보험료 [(노)X요율]" 류 라벨을 전부
//  거부해 이 시트에서 11건 → 0건이 됐던 사고의 회귀 테스트다.)
const REAL_FILE = resolve('..', '레퍼런스 엄선', '2. 추정가격내역서(간단).xlsx');
const REAL_SHEET = '제비율 적용기준';
const REAL_EXPECTED = [
  ['간접노무비', 0.126, 'F42'],
  ['산재보험료', 0.0356, 'F43'],
  ['고용보험료', 0.0101, 'F44'],
  ['국민건강보험료', 0.03545, 'F45'],
  ['국민연금보험료', 0.045, 'F46'],
  ['노인장기요양보험료', 0.1295, 'F47'],
  ['산업안전보건관리비', 0.0364, 'F48'],
  ['기타경비', 0.05002, 'F49'],
  ['일반관리비', 0.06, 'F50'],
  ['이윤', 0.15, 'F51'],
];

test('buildRateCriteria: 실제 제비율 시트에서 표준 항목 기준요율을 전부 추출한다', { skip: existsSync(REAL_FILE) ? false : `레퍼런스 파일 없음: ${REAL_FILE}` }, async () => {
  const server = await createServer({ configFile: 'vite.config.mjs', configLoader: 'native', server: { middlewareMode: true } });
  try {
    const excel = await server.ssrLoadModule('/src/utils/excel.ts');
    const mod = await server.ssrLoadModule('/src/utils/criteria.ts');
    const wb = XLSX.read(readFileSync(REAL_FILE), { cellNF: true, cellStyles: true, cellFormula: true, cellDates: true });
    const ir = excel.workbookToIR(wb, { fileName: REAL_FILE, procurementType: 'CONSTRUCTION' });
    const sheet = ir.sheets.find((s) => s.sheetName === REAL_SHEET);
    assert.ok(sheet, `시트 ${REAL_SHEET} 를 찾지 못했습니다`);

    const criteria = mod.buildRateCriteria({ ...ir, sheets: [sheet] });
    const byName = new Map(criteria.map((c) => [c.canonicalName, c]));
    for (const [name, rate, cell] of REAL_EXPECTED) {
      const hit = byName.get(name);
      assert.ok(hit, `${name} 기준요율이 추출돼야 함 (실제: ${JSON.stringify(criteria.map((c) => c.canonicalName))})`);
      assert.equal(hit.cell, cell, `${name} 근거 셀`);
      assert.ok(Math.abs(hit.rate - rate) < 1e-9, `${name} 요율 기대 ${rate} ↔ 실제 ${hit.rate}`);
    }
    // 잡담 문구("조달청 기타경비 요율", "※ 기타경비 : …", "안전관리비 대상액")가 항목을 추가로 만들면 안 된다.
    assert.equal(criteria.length, REAL_EXPECTED.length, `추출 항목 수: ${JSON.stringify(criteria.map((c) => `${c.canonicalName}@${c.cell}`))}`);
  } finally {
    await server.close();
  }
});
