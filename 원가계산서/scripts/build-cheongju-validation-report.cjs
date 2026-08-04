const fs = require('node:fs');
const path = require('node:path');
const XLSX = require('xlsx');

const outputDir = 'C:/Users/user/.codex/visualizations/2026/07/22/019f8730-deb9-7402-b177-88629938069c';
const normalizedJsonPath = path.join(outputDir, 'standard-cost-statements.json');
const htmlOutput = path.join(outputDir, 'cheongju-validation-report.html');
const jsonOutput = path.join(outputDir, 'cheongju-validation-report.json');

const files = {
  cost: 'C:/Users/user/Documents/AI Pioneer/7\uC8FC\uCC28/\uB808\uD37C\uB7F0\uC2A4 \uC5C4\uC120/8. \uC124\uACC4\uB0B4\uC5ED\uC11C(\uCCAD\uC8FC\uC9C0\uC0AC).xlsx',
  labor: 'C:/Users/user/Documents/AI Pioneer/7\uC8FC\uCC28/\uB178\uC784\uB2E8\uAC00\uD45C.xlsx',
  rate: 'C:/Users/user/Documents/AI Pioneer/7\uC8FC\uCC28/\uD1A0\uBAA9\uACF5\uC0AC \uAC04\uC811\uACF5\uC0AC\uBE44 \uC801\uC6A9\uAE30\uC900(260413).xlsx',
};

function readWorkbook(file) {
  return XLSX.readFile(file, { cellFormula: true, cellNF: true, cellDates: true });
}

function display(ws, addr) {
  const cell = ws[addr];
  if (!cell) return '';
  return cell.w ?? (cell.v == null ? '' : String(cell.v));
}

function formula(ws, addr) {
  const cell = ws[addr];
  return cell?.f ? `=${cell.f}` : '';
}

function cellValue(ws, addr) {
  const cell = ws[addr];
  if (!cell) return null;
  if (typeof cell.v === 'number') return cell.v;
  const text = String(cell.w ?? cell.v ?? '').trim();
  if (!text || text === '-') return null;
  const cleaned = text.replace(/[,\s()%원]/g, '');
  const value = Number(cleaned);
  return Number.isFinite(value) ? value : null;
}

function rateNumberFromText(text) {
  const match = String(text || '').match(/([-+]?\d+(?:\.\d+)?)/);
  if (!match) return null;
  return Number(match[1]) / 100;
}

function rateFromCell(ws, addr) {
  const n = cellValue(ws, addr);
  if (n == null) return null;
  return n > 1 ? n / 100 : n;
}

function criteriaRateFromCell(ws, addr) {
  const n = cellValue(ws, addr);
  return n == null ? null : n / 100;
}

function source(fileName, sheetName, addr) {
  return `${fileName} / ${sheetName}!${addr}`;
}

function normalizeName(text) {
  return String(text || '')
    .replace(/\s+/g, '')
    .replace(/[()[\]{}·,._\-]/g, '')
    .toLowerCase();
}

function statusForDiff(diff, tolerance = 1) {
  return Math.abs(diff) <= tolerance ? '정상' : '오류';
}

function roundDown(value, digits = 0) {
  const factor = 10 ** -digits;
  if (digits >= 0) return Math.trunc(value * 10 ** digits) / 10 ** digits;
  return Math.trunc(value / factor) * factor;
}

function formatNumber(value, digits = 0) {
  if (value == null || Number.isNaN(Number(value))) return '';
  return Number(value).toLocaleString('ko-KR', {
    maximumFractionDigits: digits,
    minimumFractionDigits: digits > 0 ? Math.min(digits, 2) : 0,
  });
}

function formatRate(value) {
  if (value == null || Number.isNaN(Number(value))) return '';
  return `${(Number(value) * 100).toLocaleString('ko-KR', { maximumFractionDigits: 4 })}%`;
}

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>"']/g, (ch) => (
    { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]
  ));
}

function loadCheongjuStatement() {
  const normalized = JSON.parse(fs.readFileSync(normalizedJsonPath, 'utf8'));
  const workbook = normalized.workbooks.find((item) => item.fileName.includes('청주'));
  if (!workbook?.statements?.length) {
    throw new Error('청주 원가계산서 정규화 결과를 찾지 못했습니다. build-standard-cost-report.cjs를 먼저 실행하세요.');
  }
  return { workbook, statement: workbook.statements[0] };
}

function parseLaborCriteria(wb) {
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const range = XLSX.utils.decode_range(ws['!ref']);
  const rows = [];
  const byName = new Map();
  for (let row = range.s.r + 1; row <= range.e.r; row += 1) {
    const jobNo = display(ws, XLSX.utils.encode_cell({ r: row, c: 0 })).trim();
    const name = display(ws, XLSX.utils.encode_cell({ r: row, c: 1 })).trim();
    const amountAddr = XLSX.utils.encode_cell({ r: row, c: 2 });
    const amount = cellValue(ws, amountAddr);
    if (!name) continue;
    const item = {
      jobNo,
      name,
      amount,
      source: source(path.basename(files.labor), sheetName, amountAddr),
      sourceCell: amountAddr,
    };
    rows.push(item);
    if (amount != null) byName.set(normalizeName(name), item);
  }
  return { sheetName, rows, byName, numericCount: rows.filter((item) => item.amount != null).length };
}

function parseLaborUsages(costWb, laborCriteria) {
  const sheetName = costWb.SheetNames.find((name) => name.includes('일위대가') && name !== '일위대가목록') || costWb.SheetNames[6];
  const ws = costWb.Sheets[sheetName];
  const range = XLSX.utils.decode_range(ws['!ref']);
  const checks = [];
  for (let row = range.s.r + 1; row <= range.e.r; row += 1) {
    const name = display(ws, XLSX.utils.encode_cell({ r: row, c: 0 })).trim();
    const unit = display(ws, XLSX.utils.encode_cell({ r: row, c: 2 })).trim();
    const qty = cellValue(ws, XLSX.utils.encode_cell({ r: row, c: 3 }));
    const unitPriceAddr = XLSX.utils.encode_cell({ r: row, c: 6 });
    const inputRate = cellValue(ws, unitPriceAddr);
    const amount = cellValue(ws, XLSX.utils.encode_cell({ r: row, c: 7 }));
    const note = display(ws, XLSX.utils.encode_cell({ r: row, c: 12 })).trim();
    const criteria = laborCriteria.byName.get(normalizeName(name));
    if (!criteria || inputRate == null || inputRate <= 0) continue;
    const diff = inputRate - criteria.amount;
    checks.push({
      category: '노임단가',
      status: statusForDiff(diff, 0.5),
      sourceRow: row + 1,
      item: name,
      unit,
      quantity: qty,
      inputValue: inputRate,
      expectedValue: criteria.amount,
      difference: diff,
      amount,
      sourceCell: `${sheetName}!${unitPriceAddr}`,
      criteriaSource: criteria.source,
      note: note || '일위대가 노무비 단가와 업로드 노임단가표 비교',
    });
  }
  return { sheetName, checks };
}

function rateCriteria(rateWb) {
  const sheetName = rateWb.SheetNames[0];
  const ws = rateWb.Sheets[sheetName];
  const fileName = path.basename(files.rate);
  const directSmallCivil = {
    label: '10억 미만·6개월 이하·토목 기준',
    indirectLabor: criteriaRateFromCell(ws, 'AC15'),
    otherExpenses: criteriaRateFromCell(ws, 'AL15'),
  };
  return {
    sheetName,
    title: display(ws, 'AL2') || display(ws, 'B2'),
    effectiveDate: display(ws, 'B2'),
    entries: {
      INDIRECT_LABOR_COST: {
        expectedRate: null,
        referenceRate: directSmallCivil.indirectLabor,
        sourceCell: 'AC15',
        sourceText: `${directSmallCivil.label}: ${formatRate(directSmallCivil.indirectLabor)}`,
        note: '업로드 제비율표는 토목·조경·산업환경설비 기준입니다. 청주 원가계산서는 건축 유지보수 성격이라 건축 간접노무비 기준이 필요합니다.',
      },
      WORKERS_COMP_INSURANCE: {
        expectedRate: rateNumberFromText(display(ws, 'BZ51')),
        sourceCell: 'BZ51',
        sourceText: display(ws, 'BZ51'),
      },
      EMPLOYMENT_INSURANCE: {
        expectedRate: criteriaRateFromCell(ws, 'AH79'),
        sourceCell: 'AH79',
        sourceText: `${display(ws, 'B79')} / ${display(ws, 'AH79')}%`,
      },
      HEALTH_INSURANCE: {
        expectedRate: rateNumberFromText(display(ws, 'AU51')),
        sourceCell: 'AU51',
        sourceText: display(ws, 'AU51'),
      },
      NATIONAL_PENSION: {
        expectedRate: rateNumberFromText(display(ws, 'BR51')),
        sourceCell: 'BR51',
        sourceText: display(ws, 'BR51'),
      },
      LONG_TERM_CARE_INSURANCE: {
        expectedRate: rateNumberFromText(display(ws, 'BD51')),
        sourceCell: 'BD51',
        sourceText: display(ws, 'BD51'),
      },
      RETIREMENT_MUTUAL_AID: {
        expectedRate: rateNumberFromText(display(ws, 'AU110')),
        sourceCell: 'AU110',
        sourceText: display(ws, 'AU110'),
      },
      OCCUPATIONAL_SAFETY_HEALTH: {
        expectedRate: criteriaRateFromCell(ws, 'ES19'),
        sourceCell: 'ES19',
        sourceText: `${display(ws, 'DK19')} / ${display(ws, 'EB19')} / ${display(ws, 'ES19')}%`,
      },
      OTHER_EXPENSES: {
        expectedRate: null,
        referenceRate: directSmallCivil.otherExpenses,
        sourceCell: 'AL15',
        sourceText: `${directSmallCivil.label}: ${formatRate(directSmallCivil.otherExpenses)}`,
        note: '청주 원가계산서의 기타경비율 4.424%는 업로드 토목 제비율표의 토목 기준과 직접 매칭되지 않습니다. 건축/유지보수 기타경비 적용기준 또는 원본의 붙임 기준표가 필요합니다.',
      },
      ENVIRONMENT_PRESERVATION_COST: {
        expectedRate: null,
        referenceRate: criteriaRateFromCell(ws, 'BU88'),
        sourceCell: 'BU88',
        sourceText: `${display(ws, 'AU88')} / ${display(ws, 'AZ88')} / ${display(ws, 'BU88')}%, ${display(ws, 'AZ94')} / ${display(ws, 'BU94')}%`,
        note: '환경보전비는 업로드 기준표에서 정확한 공종 분류 선택이 필요합니다. 청주 원본의 0.30%는 전문·개보수 성격 기준으로 보이며, 업로드 기준표만으로 자동 확정하지 않았습니다.',
      },
      GENERAL_ADMIN_COST: {
        expectedRate: criteriaRateFromCell(ws, 'BD15'),
        sourceCell: 'BD15',
        sourceText: `${display(ws, 'AU15')} / ${display(ws, 'BD15')}%`,
      },
      PROFIT: {
        expectedRate: criteriaRateFromCell(ws, 'BS15'),
        sourceCell: 'BS15',
        sourceText: `${display(ws, 'BS3')} / ${display(ws, 'BS15')}%`,
      },
      WORKERS_ACCIDENT_LIABILITY_INSURANCE: {
        expectedRate: null,
        sourceCell: '',
        sourceText: '업로드 제비율표에서 근로자재해보장책임보험료 기준을 자동 식별하지 못했습니다.',
        note: '원가계산서에는 노무비 x 0.125% 또는 20,000원 중 큰 금액으로 계산되어 있습니다.',
      },
    },
    fileName,
  };
}

function buildRateChecks(statement, criteria) {
  const rowByCode = new Map(statement.rows.map((row) => [row.canonicalCode, row]));
  const targetCodes = [
    'INDIRECT_LABOR_COST',
    'WORKERS_COMP_INSURANCE',
    'EMPLOYMENT_INSURANCE',
    'HEALTH_INSURANCE',
    'NATIONAL_PENSION',
    'LONG_TERM_CARE_INSURANCE',
    'RETIREMENT_MUTUAL_AID',
    'OCCUPATIONAL_SAFETY_HEALTH',
    'OTHER_EXPENSES',
    'ENVIRONMENT_PRESERVATION_COST',
    'GENERAL_ADMIN_COST',
    'PROFIT',
    'WORKERS_ACCIDENT_LIABILITY_INSURANCE',
  ];
  return targetCodes.map((code) => {
    const row = rowByCode.get(code);
    const criteriaEntry = criteria.entries[code] || {};
    const inputRate = row?.rate?.value ?? null;
    const expectedRate = criteriaEntry.expectedRate ?? null;
    const diff = inputRate != null && expectedRate != null ? inputRate - expectedRate : null;
    let status = '검증불가';
    let message = criteriaEntry.note || '';
    if (row && inputRate == null && expectedRate != null) {
      status = '검증불가';
      message = '원가계산서에서 요율 셀을 찾지 못했습니다.';
    } else if (row && inputRate != null && expectedRate != null) {
      const same = Math.abs(diff) <= 0.0000001;
      if (same) status = '정상';
      else if ((row.amount?.value ?? 0) === 0 || /미적용/.test(row.note || '')) {
        status = '확인 필요';
        message = '요율은 업로드 기준과 다르지만 원가계산서에서 미적용 또는 0원 처리되어 금액 영향은 제한적입니다.';
      } else {
        status = '오류';
        message = '원가계산서 요율과 업로드 기준자료 요율이 다릅니다.';
      }
    }
    return {
      category: '제비율',
      status,
      code,
      item: row?.canonicalName || code,
      originalName: row?.originalName || '',
      section: row?.section || '',
      sourceRow: row?.sourceRow || '',
      inputRate,
      expectedRate,
      referenceRate: criteriaEntry.referenceRate ?? null,
      difference: diff,
      inputAmount: row?.amount?.value ?? null,
      sourceCell: row?.rate?.address ? `${statement.sheetName}!${row.rate.address}` : '',
      criteriaSource: criteriaEntry.sourceCell ? source(criteria.fileName, criteria.sheetName, criteriaEntry.sourceCell) : '',
      criteriaText: criteriaEntry.sourceText || '',
      note: message,
    };
  });
}

function buildArithmeticChecks(costWb) {
  const sheetName = costWb.SheetNames.find((name) => name === '원가계산서') || costWb.SheetNames[1];
  const ws = costWb.Sheets[sheetName];
  const v = (addr) => cellValue(ws, addr) ?? 0;
  const r = (addr) => rateFromCell(ws, addr) ?? 0;
  const checks = [];
  function add(label, addr, expected, formulaText, tolerance = 1) {
    const actual = v(addr);
    const diff = actual - expected;
    checks.push({
      category: '산식',
      status: statusForDiff(diff, tolerance),
      item: label,
      sourceCell: `${sheetName}!${addr}`,
      inputValue: actual,
      expectedValue: expected,
      difference: diff,
      formula: formula(ws, addr) || formulaText,
      note: formulaText,
    });
  }
  add('재료비 소계', 'E7', v('E4'), 'E4');
  add('간접노무비', 'E9', roundDown(v('E8') * r('H9')), 'ROUNDDOWN(직접노무비 x 15%, 0)');
  add('노무비 소계', 'E10', v('E8') + v('E9'), '직접노무비 + 간접노무비');
  add('산재보험료', 'E12', Math.trunc(v('E10') * r('H12')), 'INT(노무비 x 3.56%)');
  add('고용보험료', 'E13', Math.trunc(v('E10') * r('H13')), 'INT(노무비 x 1.01%)');
  add('퇴직공제부금비', 'E17', Math.trunc(v('E8') * r('H17')), 'INT(직접노무비 x 2.30%)');
  add('산업안전보건관리비', 'E18', Math.trunc(v('E7') + v('E8')) * r('H18'), 'INT(재료비+직접노무비) x 3.11%', 0.01);
  add('기타경비', 'E19', Math.trunc((v('E7') + v('E10')) * r('H19')), 'INT((재료비+노무비) x 4.424%)');
  add('환경보전비', 'E20', Math.trunc((v('E7') + v('E8') + v('E11')) * r('H20')), 'INT((재료비+직접노무비+기계경비) x 0.30%)');
  add('경비 소계', 'E21', v('E11') + v('E12') + v('E13') + v('E14') + v('E15') + v('E16') + v('E17') + v('E18') + v('E19') + v('E20'), 'SUM(E11:E20)', 0.01);
  add('순공사원가 계', 'E22', v('E7') + v('E10') + v('E21'), '재료비 소계 + 노무비 소계 + 경비 소계', 0.01);
  add('일반관리비', 'E23', roundDown(v('E22') * r('H23')), 'ROUNDDOWN(계 x 8%, 0)');
  add('이윤', 'E24', roundDown((v('E10') + v('E21') + v('E23')) * r('H24')), 'ROUNDDOWN((노무비+경비+일반관리비) x 15%, 0)');
  add('근로자재해보장책임보험', 'E25', roundDown(Math.max(v('E10') * r('H25'), 20000)), 'ROUNDDOWN(MAX(노무비 x 0.125%, 20,000), 0)');
  add('공급가액', 'E27', roundDown(v('E22') + v('E23') + v('E24') + v('E25') + v('E26'), -4), 'ROUNDDOWN(SUM(E22:E26), -4)');
  add('부가가치세', 'E28', Math.round(v('E27') * r('H28') * 10) / 10, 'ROUND(공급가액 x 10%, 1)');
  add('도급액', 'E29', v('E27') + v('E28'), '공급가액 + 부가가치세');

  const totalDiff = v('D30') - v('E29');
  checks.push({
    category: '산식',
    status: Math.abs(totalDiff) <= 1 ? '정상' : '확인 필요',
    item: '총공사비 표시값',
    sourceCell: `${sheetName}!D30`,
    inputValue: v('D30'),
    expectedValue: v('E29'),
    difference: totalDiff,
    formula: formula(ws, 'D30'),
    note: 'D30 총공사비가 도급액에서 10,786원을 차감하는 별도 수식입니다. 의도된 조정인지 확인이 필요합니다.',
  });
  return { sheetName, checks };
}

function buildSectionChecks(statement) {
  const rows = statement.rows;
  const laborIdx = rows.findIndex((row) => row.canonicalCode === 'LABOR_SUBTOTAL');
  const expenseIdx = rows.findIndex((row) => row.canonicalCode === 'EXPENSE_SUBTOTAL');
  const checks = [];
  const expense = '경비';
  const between = laborIdx >= 0 && expenseIdx >= 0 ? rows.slice(laborIdx + 1, expenseIdx + 1) : [];
  const after = expenseIdx >= 0 ? rows.slice(expenseIdx + 1) : [];
  const betweenBad = between.filter((row) => row.section !== expense);
  const afterBad = after.filter((row) => row.section === expense);
  checks.push({
    category: '정규화',
    status: betweenBad.length ? '오류' : '정상',
    item: '노무비 소계~경비 소계 구간',
    inputValue: between.length,
    expectedValue: between.length,
    difference: betweenBad.length,
    sourceCell: `${statement.sheetName}!${between[0]?.sourceRow || ''}:${between.at(-1)?.sourceRow || ''}`,
    note: betweenBad.length ? `${betweenBad.length}개 행이 경비로 분류되지 않았습니다.` : '구간 내 모든 행이 경비로 분류되었습니다.',
  });
  checks.push({
    category: '정규화',
    status: afterBad.length ? '오류' : '정상',
    item: '경비 소계 이후 구간',
    inputValue: after.length,
    expectedValue: after.length,
    difference: afterBad.length,
    sourceCell: `${statement.sheetName}!${after[0]?.sourceRow || ''}:끝`,
    note: afterBad.length ? `${afterBad.length}개 행이 경비를 잘못 상속했습니다.` : '경비 소계 이후 경비 상속이 없습니다.',
  });
  return checks;
}

function countStatuses(checks) {
  return checks.reduce((acc, check) => {
    acc[check.status] = (acc[check.status] || 0) + 1;
    return acc;
  }, {});
}

function rowClass(status) {
  if (status === '오류') return 'danger';
  if (status === '확인 필요' || status === '검증불가') return 'warn';
  return 'ok';
}

function renderCheckRows(checks, mode) {
  return checks.map((check) => {
    const input = mode === 'rate' ? formatRate(check.inputRate) : formatNumber(check.inputValue, 2);
    const expected = mode === 'rate'
      ? (check.expectedRate != null ? formatRate(check.expectedRate) : check.referenceRate != null ? `참고 ${formatRate(check.referenceRate)}` : '')
      : formatNumber(check.expectedValue, 2);
    const diff = mode === 'rate'
      ? (check.difference != null ? formatRate(check.difference) : '')
      : formatNumber(check.difference, 2);
    const item = check.item || check.originalName;
    return `<tr class="${rowClass(check.status)}">
      <td><span class="pill">${escapeHtml(check.status)}</span></td>
      <td>${escapeHtml(check.category)}</td>
      <td>${escapeHtml(item)}</td>
      <td class="num">${escapeHtml(input)}</td>
      <td class="num">${escapeHtml(expected)}</td>
      <td class="num">${escapeHtml(diff)}</td>
      <td>${escapeHtml(check.sourceCell || '')}</td>
      <td>${escapeHtml(check.criteriaSource || check.formula || '')}</td>
      <td>${escapeHtml(check.note || check.criteriaText || '')}</td>
    </tr>`;
  }).join('');
}

function renderLaborRows(checks) {
  return checks.map((check) => `<tr class="${rowClass(check.status)}">
    <td><span class="pill">${escapeHtml(check.status)}</span></td>
    <td>${escapeHtml(check.item)}</td>
    <td>${escapeHtml(check.unit)}</td>
    <td class="num">${formatNumber(check.quantity, 3)}</td>
    <td class="num">${formatNumber(check.inputValue)}</td>
    <td class="num">${formatNumber(check.expectedValue)}</td>
    <td class="num">${formatNumber(check.difference)}</td>
    <td>${escapeHtml(check.sourceCell)}</td>
    <td>${escapeHtml(check.criteriaSource)}</td>
  </tr>`).join('');
}

function renderNormalizedRows(rows) {
  return rows.map((row) => `<tr>
    <td>${escapeHtml(row.sourceRow)}</td>
    <td>${escapeHtml(row.section)}</td>
    <td>${escapeHtml(row.canonicalName)}</td>
    <td>${escapeHtml(row.originalName)}</td>
    <td class="num">${formatNumber(row.amount?.value, 2)}</td>
    <td>${escapeHtml(row.rate ? formatRate(row.rate.value) : '')}</td>
    <td>${escapeHtml(row.calculationCell?.source || row.amount?.address || '')}</td>
    <td>${escapeHtml(row.resolutionStatus || '')}</td>
    <td>${escapeHtml(row.note || '')}</td>
  </tr>`).join('');
}

function html(payload) {
  const { summary, criteriaSummary, rateChecks, laborChecks, arithmeticChecks, sectionChecks, normalizedRows } = payload;
  const statusCounts = summary.statusCounts;
  const metrics = [
    ['전체 상태', summary.overallStatus],
    ['오류', statusCounts['오류'] || 0],
    ['확인 필요', statusCounts['확인 필요'] || 0],
    ['검증불가', statusCounts['검증불가'] || 0],
    ['정상', statusCounts['정상'] || 0],
    ['노임 불일치', laborChecks.filter((item) => item.status === '오류').length],
  ];
  return `<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>청주 원가계산서 검증 리포트</title>
  <style>
    :root { color-scheme: light; --ink:#1f2937; --muted:#667085; --line:#d9dee7; --bg:#f6f8fb; --card:#fff; --danger:#b42318; --warn:#b54708; --ok:#027a48; --blue:#175cd3; }
    * { box-sizing: border-box; }
    body { margin:0; font-family: Arial, "Malgun Gothic", sans-serif; color:var(--ink); background:var(--bg); }
    header { padding:24px 28px 18px; background:#14213d; color:white; }
    header h1 { margin:0 0 8px; font-size:24px; }
    header p { margin:0; color:#d6e0f5; font-size:14px; }
    main { padding:22px 28px 40px; display:grid; gap:18px; }
    section { background:var(--card); border:1px solid var(--line); border-radius:8px; overflow:hidden; }
    section h2 { margin:0; padding:14px 16px; font-size:17px; border-bottom:1px solid var(--line); background:#fbfcff; }
    .content { padding:16px; }
    .metrics { display:grid; grid-template-columns: repeat(6, minmax(120px, 1fr)); gap:10px; }
    .metric { border:1px solid var(--line); border-radius:8px; padding:12px; background:#fff; }
    .metric span { display:block; color:var(--muted); font-size:12px; margin-bottom:6px; }
    .metric strong { font-size:20px; }
    .notice { padding:12px 14px; border-left:4px solid var(--warn); background:#fff7ed; margin-top:14px; line-height:1.55; }
    .grid2 { display:grid; grid-template-columns: 1fr 1fr; gap:12px; }
    .kv { display:grid; grid-template-columns: 160px 1fr; gap:6px 12px; font-size:14px; }
    .kv b { color:#344054; }
    table { width:100%; border-collapse:collapse; font-size:13px; }
    th, td { border-bottom:1px solid #e8ecf2; padding:8px 9px; vertical-align:top; }
    th { text-align:left; color:#475467; background:#f8fafc; position:sticky; top:0; z-index:1; }
    .table-wrap { max-height:560px; overflow:auto; }
    .num { text-align:right; white-space:nowrap; font-variant-numeric: tabular-nums; }
    .pill { display:inline-block; min-width:58px; text-align:center; padding:3px 7px; border-radius:999px; font-size:12px; font-weight:700; }
    tr.ok .pill { background:#ecfdf3; color:var(--ok); }
    tr.warn .pill { background:#fff7ed; color:var(--warn); }
    tr.danger .pill { background:#fef3f2; color:var(--danger); }
    .cards { display:grid; grid-template-columns: repeat(2, minmax(240px, 1fr)); gap:10px; }
    .card { border:1px solid var(--line); border-radius:8px; padding:12px; background:#fff; }
    .card h3 { margin:0 0 8px; font-size:14px; color:#344054; }
    .card p { margin:0; line-height:1.55; color:#475467; font-size:13px; }
    .small { color:var(--muted); font-size:12px; line-height:1.5; }
    @media (max-width: 960px) { .metrics, .grid2, .cards { grid-template-columns:1fr; } }
  </style>
</head>
<body>
  <header>
    <h1>청주지사 원가계산서 검증 리포트</h1>
    <p>원가계산서 Excel + 제비율 Excel + 노임단가 Excel 기준 테스트 결과</p>
  </header>
  <main>
    <section>
      <h2>검증결과 요약</h2>
      <div class="content">
        <div class="metrics">
          ${metrics.map(([label, value]) => `<div class="metric"><span>${escapeHtml(label)}</span><strong>${escapeHtml(value)}</strong></div>`).join('')}
        </div>
        <div class="notice">
          <b>핵심 확인사항</b><br>
          1. 일위대가 노임단가가 업로드한 노임단가표와 불일치합니다. 도장공, 보통인부, 특별인부 단가가 모두 낮게 입력되어 있습니다.<br>
          2. 업로드한 제비율표는 2026년 토목·조경·산업환경설비 기준이고, 청주 원본에는 2025년 건축제비율 시트가 포함되어 있어 일부 항목은 기준자료 불일치 가능성이 있습니다.<br>
          3. 건강보험료·연금보험료·노인장기요양보험료 요율은 업로드 기준과 다르지만 원가계산서에서는 1개월 미만 공사 미적용으로 0원 처리되어 있습니다.<br>
          4. 총공사비 D30은 도급액에서 10,786원을 차감하는 별도 수식이므로 의도된 조정인지 확인이 필요합니다.
        </div>
      </div>
    </section>

    <section>
      <h2>입력 파일 및 기준자료 인식</h2>
      <div class="content grid2">
        <div class="kv">
          <b>원가계산서</b><span>${escapeHtml(criteriaSummary.costFile)}</span>
          <b>원가 시트</b><span>${escapeHtml(criteriaSummary.costSheet)}</span>
          <b>정규화 행</b><span>${escapeHtml(summary.normalizedRowCount)}개</span>
          <b>노임단가표</b><span>${escapeHtml(criteriaSummary.laborFile)} · ${escapeHtml(criteriaSummary.laborCount)}개 직종</span>
        </div>
        <div class="kv">
          <b>제비율표</b><span>${escapeHtml(criteriaSummary.rateFile)}</span>
          <b>제비율 시트</b><span>${escapeHtml(criteriaSummary.rateSheet)}</span>
          <b>적용시기</b><span>${escapeHtml(criteriaSummary.effectiveDate)}</span>
          <b>문서 제목</b><span>${escapeHtml(criteriaSummary.rateTitle)}</span>
        </div>
      </div>
    </section>

    <section>
      <h2>제비율 검증</h2>
      <div class="table-wrap">
        <table>
          <thead><tr><th>상태</th><th>구분</th><th>항목</th><th>입력요율</th><th>기준요율</th><th>차이</th><th>원가 위치</th><th>기준 위치</th><th>판단 메모</th></tr></thead>
          <tbody>${renderCheckRows(rateChecks, 'rate')}</tbody>
        </table>
      </div>
    </section>

    <section>
      <h2>노임단가 검증</h2>
      <div class="table-wrap">
        <table>
          <thead><tr><th>상태</th><th>직종</th><th>단위</th><th>수량</th><th>입력단가</th><th>기준단가</th><th>차이</th><th>원가 위치</th><th>기준 위치</th></tr></thead>
          <tbody>${renderLaborRows(laborChecks)}</tbody>
        </table>
      </div>
    </section>

    <section>
      <h2>산식 및 합계 검증</h2>
      <div class="table-wrap">
        <table>
          <thead><tr><th>상태</th><th>구분</th><th>항목</th><th>입력값</th><th>계산값</th><th>차이</th><th>원가 위치</th><th>수식</th><th>판단 메모</th></tr></thead>
          <tbody>${renderCheckRows([...sectionChecks, ...arithmeticChecks], 'amount')}</tbody>
        </table>
      </div>
    </section>

    <section>
      <h2>정규화 원가계산서 표</h2>
      <div class="table-wrap">
        <table>
          <thead><tr><th>원본행</th><th>원본구간</th><th>표준항목</th><th>원본항목</th><th>금액</th><th>요율</th><th>계산셀</th><th>추적상태</th><th>비고</th></tr></thead>
          <tbody>${renderNormalizedRows(normalizedRows)}</tbody>
        </table>
      </div>
      <div class="content small">정규화 기준: 노무비 소계 다음부터 경비 소계까지는 경비로 분류하고, 경비 소계 이후는 경비를 상속하지 않습니다.</div>
    </section>
  </main>
</body>
</html>`;
}

const { workbook: normalizedWorkbook, statement } = loadCheongjuStatement();
const costWb = readWorkbook(files.cost);
const laborWb = readWorkbook(files.labor);
const rateWb = readWorkbook(files.rate);

const laborCriteria = parseLaborCriteria(laborWb);
const laborUsages = parseLaborUsages(costWb, laborCriteria);
const criteria = rateCriteria(rateWb);
const rateChecks = buildRateChecks(statement, criteria);
const arithmetic = buildArithmeticChecks(costWb);
const sectionChecks = buildSectionChecks(statement);
const allChecks = [...rateChecks, ...laborUsages.checks, ...arithmetic.checks, ...sectionChecks];
const statusCounts = countStatuses(allChecks);
const overallStatus = statusCounts['오류'] ? '오류' : statusCounts['확인 필요'] || statusCounts['검증불가'] ? '확인 필요' : '정상';

const payload = {
  generatedAt: new Date().toISOString(),
  summary: {
    overallStatus,
    statusCounts,
    normalizedRowCount: statement.rows.length,
    rateCheckCount: rateChecks.length,
    laborCheckCount: laborUsages.checks.length,
    arithmeticCheckCount: arithmetic.checks.length,
  },
  criteriaSummary: {
    costFile: normalizedWorkbook.fileName,
    costSheet: statement.sheetName,
    laborFile: path.basename(files.labor),
    laborSheet: laborCriteria.sheetName,
    laborCount: laborCriteria.numericCount,
    rateFile: criteria.fileName,
    rateSheet: criteria.sheetName,
    rateTitle: criteria.title,
    effectiveDate: criteria.effectiveDate,
  },
  rateChecks,
  laborChecks: laborUsages.checks,
  arithmeticChecks: arithmetic.checks,
  sectionChecks,
  normalizedRows: statement.rows,
};

fs.mkdirSync(outputDir, { recursive: true });
fs.writeFileSync(jsonOutput, `${JSON.stringify(payload, null, 2)}\n`, 'utf8');
fs.writeFileSync(htmlOutput, html(payload), 'utf8');

console.log(JSON.stringify({
  htmlOutput,
  jsonOutput,
  overallStatus,
  statusCounts,
  rateChecks: rateChecks.length,
  laborChecks: laborUsages.checks.length,
  arithmeticChecks: arithmetic.checks.length,
}, null, 2));
