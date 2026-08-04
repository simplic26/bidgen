import assert from 'node:assert/strict';
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
