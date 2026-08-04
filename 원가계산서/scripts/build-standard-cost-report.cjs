const fs = require('node:fs');
const path = require('node:path');
const XLSX = require('xlsx');

const base = 'C:/Users/user/Documents/AI Pioneer/7\uC8FC\uCC28/\uB808\uD37C\uB7F0\uC2A4 \uC5C4\uC120';
const outputDir = 'C:/Users/user/.codex/visualizations/2026/07/22/019f8730-deb9-7402-b177-88629938069c';
const htmlOutput = path.join(outputDir, 'standard-cost-statement-review.html');
const jsonOutput = path.join(outputDir, 'standard-cost-statements.json');

const inputFiles = [
  {
    id: 'pyeongtaek-2026',
    path: `${base}/3. 2026\uB144 \uD3C9\uD0DD\uC9C0\uC0AC \uC5F4\uC218\uC1A1\uAD00\uACF5\uC0AC \uC124\uACC4\uB0B4\uC5ED\uC11C.xls`,
  },
  {
    id: 'estimate-2',
    path: `${base}/2. \uCD94\uC815\uAC00\uACA9\uB0B4\uC5ED\uC11C.xlsx`,
  },
  {
    id: 'cheongju-8',
    path: `${base}/8. \uC124\uACC4\uB0B4\uC5ED\uC11C(\uCCAD\uC8FC\uC9C0\uC0AC).xlsx`,
  },
];

const REF_RE = /(?:(?:'[^']+'|[A-Za-z0-9_\uAC00-\uD7A3 ()_.]+)!)?\$?[A-Z]{1,3}\$?[0-9]+/g;

const canonicalItems = [
  ['DIRECT_MATERIAL_COST', '직접재료비', ['직접재료비', '직접재료']],
  ['INDIRECT_MATERIAL_COST', '간접재료비', ['간접재료비', '간접재료']],
  ['MATERIAL_BYPRODUCT_DEDUCTION', '작업설·부산물', ['작업설부산물', '작업설,부산물', '부산물']],
  ['MATERIAL_SUBTOTAL', '재료비 소계', ['재료비소계', '재료비계', '소계'], '재료비'],
  ['DIRECT_LABOR_COST', '직접노무비', ['직접노무비', '직접노무']],
  ['INDIRECT_LABOR_COST', '간접노무비', ['간접노무비', '간접노무']],
  ['LABOR_SUBTOTAL', '노무비 소계', ['노무비소계', '노무비계', '소계'], '노무비'],
  ['ELECTRIC_POWER_COST', '전력비', ['전력비']],
  ['TRANSPORTATION_COST', '운반비', ['운반비']],
  ['MACHINE_COST', '기계경비', ['기계경비']],
  ['QUALITY_CONTROL_COST', '품질관리비', ['품질관리비']],
  ['WORKERS_COMP_INSURANCE', '산재보험료', ['산재보험료', '산업재해보상보험료']],
  ['EMPLOYMENT_INSURANCE', '고용보험료', ['고용보험료']],
  ['NATIONAL_PENSION', '연금보험료', ['연금보험료', '국민연금보험료', '국민연금']],
  ['HEALTH_INSURANCE', '건강보험료', ['건강보험료', '국민건강보험료']],
  ['LONG_TERM_CARE_INSURANCE', '노인장기요양보험료', ['노인장기요양보험료', '노인장기요양보험']],
  ['RETIREMENT_MUTUAL_AID', '건설근로자퇴직공제부금비', ['건설근로자퇴직공제부금비', '퇴직공제부금비', '퇴직공제부금']],
  ['OCCUPATIONAL_SAFETY_HEALTH', '산업안전보건관리비', ['산업안전보건관리비', '안전보건관리비']],
  ['SAFETY_MANAGEMENT_COST', '안전관리비', ['안전관리비']],
  ['ENVIRONMENT_PRESERVATION_COST', '환경보전비', ['환경보전비']],
  ['SERVICE_FEE', '지급수수료', ['지급수수료']],
  ['PERFORMANCE_BOND_FEE', '공사이행보증수수료', ['공사이행보증수수료']],
  ['SUBCONTRACT_PAYMENT_BOND_FEE', '하도급대금지급보증서발급수수료', ['하도급대금지급보증서발급수수료']],
  ['EQUIPMENT_RENTAL_PAYMENT_BOND_FEE', '건설기계대여대금지급보증서발급수수료', ['건설기계대여대금지급보증서발급수수료']],
  ['MATERIAL_MANAGEMENT_COST', '자재관리비', ['자재관리비']],
  ['RENTAL_COST', '지급임차료', ['지급임차료']],
  ['OTHER_EXPENSES', '기타경비', ['기타경비']],
  ['ASBESTOS_LEVY', '석면분담금', ['석면분담금']],
  ['WAGE_CLAIM_CHARGE', '임금채권부담금', ['임금채권부담금']],
  ['WASTE_DISPOSAL_COST', '폐기물처리비', ['폐기물처리비']],
  ['EQUIPMENT_MATERIAL_MANAGEMENT_COST', '기자재관리비', ['기자재관리비']],
  ['COMPENSATION_COST', '보상비', ['보상비']],
  ['R_AND_D_COST', '연구개발비', ['연구개발비']],
  ['EXPENSE_SUBTOTAL', '경비 소계', ['경비소계', '소계'], '경비'],
  ['GENERAL_ADMIN_COST', '일반관리비', ['일반관리비', '일반관리']],
  ['PROFIT', '이윤', ['이윤']],
  ['SHOP_DRAWING_COST', '시공상세도작성비', ['시공상세도작성비', '시공상세도도작성비']],
  ['STRUCTURAL_REVIEW_SERVICE_COST', '구조안정성검토용역비', ['구조안정성검토용역비', '열수송관전용교구조안정성검토용역비']],
  ['PUBLIC_SURVEYING_COST', '공공측량', ['공공측량']],
  ['TEMPORARY_WORK_COST', '가설비', ['가설비']],
  ['TRAFFIC_MEASURE_SERVICE_COST', '교통소통대책수립용역비', ['교통소통대책수립용역비']],
  ['WORK_COST', '공사원가', ['공사원가']],
  ['CONSTRUCTION_DAMAGE_INSURANCE', '공사손해보험료', ['공사손해보험료']],
  ['TOTAL_COST', '총원가', ['총원가']],
  ['SUPPLY_AMOUNT', '공급가액', ['공급가액']],
  ['VAT', '부가가치세', ['부가가치세', '부가세']],
  ['CONTRACT_AMOUNT', '도급액', ['도급액']],
  ['GRAND_TOTAL_CONSTRUCTION_COST', '총공사비', ['총공사비']],
  ['WORKERS_ACCIDENT_LIABILITY_INSURANCE', '근로자재해보장책임보험료', ['근로자재해보장책임보험료', '근로자재해보장책임보험']],
  ['ELECTRICAL_DAMAGE_LIABILITY_INSURANCE', '전기공사 손해배상보험료', ['전기공사손해배상보험료', '전기공사손해배상보험']],
  ['CONSTRUCTION_COST_SUBTOTAL', '공사비 소계', ['공사비소계']],
];

function normalizeText(text) {
  return String(text || '')
    .replace(/\s+/g, '')
    .replace(/[\[\](){},.·ㆍ:：]/g, '')
    .toLowerCase();
}

function cleanLabel(text) {
  return String(text || '')
    .replace(/\s+/g, ' ')
    .replace(/^[\s\-–—·ㆍo○]+/, '')
    .replace(/[\s:：]+$/, '')
    .trim();
}

function display(ws, addr) {
  const cell = ws[addr];
  if (!cell) return '';
  return cell.w || (cell.v == null ? '' : String(cell.v));
}

function raw(ws, addr) {
  const cell = ws[addr];
  if (!cell) return '';
  return cell.f ? `=${cell.f}` : display(ws, addr);
}

function isNumericCell(cell) {
  if (!cell) return false;
  if (cell.t === 'n' || cell.t === 'd') return true;
  const text = cell.w || (cell.v == null ? '' : String(cell.v));
  return /^\(?[-+]?\d[\d,]*(?:\.\d+)?\)?%?$/.test(text.trim());
}

function toNumber(cell) {
  if (!cell) return null;
  if (typeof cell.v === 'number') return cell.v;
  const text = (cell.w || cell.v || '').toString().trim();
  if (!text || text === '-') return null;
  const negative = /^\(.*\)$/.test(text);
  const cleaned = text.replace(/[(),\s원%]/g, '');
  const value = Number(cleaned);
  if (!Number.isFinite(value)) return null;
  return negative ? -value : value;
}

function looksLikeRate(cell) {
  if (!cell) return false;
  const text = display({ A1: cell }, 'A1').trim();
  if (String(cell.z || '').includes('%')) return true;
  return /^[-+]?\d[\d,]*(?:\.\d+)?%$/.test(text);
}

function rateValue(cell) {
  if (!cell) return null;
  const n = toNumber(cell);
  if (n == null) return null;
  if (looksLikeRate(cell)) return n;
  return n > 1 ? n / 100 : n;
}

function columnName(index) {
  return XLSX.utils.encode_col(index);
}

function address(row, col) {
  return `${columnName(col)}${row + 1}`;
}

function getRange(ws) {
  return ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : null;
}

function rowCells(ws, rowIndex, startCol, endCol) {
  const cells = [];
  for (let col = startCol; col <= endCol; col += 1) {
    const addr = address(rowIndex, col);
    const cell = ws[addr];
    if (!cell) continue;
    const text = display(ws, addr);
    const formula = cell.f ? `=${cell.f}` : '';
    cells.push({ addr, row: rowIndex, col, cell, text, formula });
  }
  return cells;
}

function findCostSheetNames(workbook) {
  return workbook.SheetNames.filter((name) => {
    const compact = name.replace(/\s+/g, '');
    if (!workbook.Sheets[name]?.['!ref']) return false;
    if (/제비율|기타경비|요율/.test(compact)) return false;
    return compact.includes('원가계산서') || compact === '원가계산';
  });
}

function findHeaderRows(ws, range) {
  const rows = [];
  for (let row = range.s.r; row <= Math.min(range.e.r, range.s.r + 12); row += 1) {
    const text = rowCells(ws, row, range.s.c, range.e.c)
      .map((item) => item.text)
      .join(' ');
    if (/금\s*액|합\s*계|적\s*용\s*율|비\s*목|구\s*분|구\s*성/.test(text)) rows.push(row);
  }
  return rows;
}

function inferAmountColumns(ws, range, headerRows) {
  const score = new Map();
  const headerCols = new Set();
  for (const row of headerRows) {
    for (const item of rowCells(ws, row, range.s.c, range.e.c)) {
      if (/금\s*액|합\s*계|원가계산서/.test(item.text)) {
        score.set(item.col, (score.get(item.col) || 0) + 10);
        headerCols.add(item.col);
      }
    }
  }
  for (let row = range.s.r; row <= range.e.r; row += 1) {
    for (const item of rowCells(ws, row, range.s.c, range.e.c)) {
      if (!isNumericCell(item.cell) || looksLikeRate(item.cell)) continue;
      if (item.col < range.s.c + 2) continue;
      score.set(item.col, (score.get(item.col) || 0) + 1);
    }
  }
  const entries = Array.from(score.entries())
    .sort((a, b) => b[1] - a[1] || a[0] - b[0])
    .map(([col]) => col);
  if (headerCols.size > 0) return entries.filter((col) => headerCols.has(col));
  return entries.slice(0, 3);
}

function inferRateColumns(ws, range, headerRows) {
  const score = new Map();
  const headerCols = new Set();
  for (const row of headerRows) {
    for (const item of rowCells(ws, row, range.s.c, range.e.c)) {
      if (/적\s*용\s*율|요\s*율|율/.test(item.text)) {
        score.set(item.col, (score.get(item.col) || 0) + 10);
        headerCols.add(item.col);
      }
    }
  }
  for (let row = range.s.r; row <= range.e.r; row += 1) {
    for (const item of rowCells(ws, row, range.s.c, range.e.c)) {
      if (looksLikeRate(item.cell)) score.set(item.col, (score.get(item.col) || 0) + 1);
    }
  }
  const entries = Array.from(score.entries())
    .sort((a, b) => b[1] - a[1] || a[0] - b[0])
    .map(([col]) => col);
  if (headerCols.size > 0) return entries.filter((col) => headerCols.has(col));
  return entries;
}

function isHeaderOrTitleRow(text) {
  const compact = normalizeText(text);
  if (!compact) return true;
  if (/원가계산서/.test(compact)) return true;
  if (/공사명|계약명/.test(compact)) return true;
  if (/비목.*구분|구분.*구성내역|금액.*비고/.test(compact)) return true;
  if (/비목/.test(compact) && /합계|금액/.test(compact) && /비고/.test(compact)) return true;
  return false;
}

function isWeakLabel(text) {
  const value = cleanLabel(text);
  if (!value) return true;
  const compact = normalizeText(value);
  if (/^[순공사원가계재료비노무경]{1,3}$/.test(compact)) return true;
  if (/^[()×x+=\-]+$/.test(compact)) return true;
  if (/^\d+$/.test(compact)) return true;
  if (/^실적정산$/.test(compact)) return true;
  if (/^만단위절사$|^천원단위절사$|^천원단위절삭$/.test(compact)) return false;
  return false;
}

function isWeakBaseLabel(text) {
  const value = cleanLabel(text);
  if (!value) return true;
  const compact = normalizeText(value);
  if (!compact) return true;
  if (/^[순공사원가계재료비노무경]{1}$/.test(compact)) return true;
  if (/^[()×x+=\-]+$/.test(compact)) return true;
  if (/^\d+$/.test(compact)) return true;
  if (/^실적정산$/.test(compact)) return true;
  return false;
}

function canonicalFor(label, currentSection) {
  const compact = normalizeText(label);
  if (!compact) return null;
  for (const [code, name, aliases, sectionHint] of canonicalItems) {
    if (sectionHint && currentSection !== sectionHint) continue;
    if (aliases.some((alias) => compact.includes(normalizeText(alias)))) {
      return { code, name, status: compact === normalizeText(name) ? 'EXACT' : 'ALIAS' };
    }
  }
  for (const [code, name, aliases] of canonicalItems) {
    if (aliases.some((alias) => compact.includes(normalizeText(alias)))) {
      return { code, name, status: compact === normalizeText(name) ? 'EXACT' : 'ALIAS' };
    }
  }
  if (/소계|합계|계$/.test(compact)) return { code: 'SUBTOTAL', name: `${currentSection || '구간'} 소계`, status: 'SUBTOTAL' };
  return null;
}

function isTopLevelIndependentRow(labelItem, rangeStartCol) {
  return Boolean(labelItem && labelItem.col === rangeStartCol);
}

function isSectionClosingRow(rowModel) {
  if (!rowModel) return false;
  if (rowModel.canonicalCode === 'SUBTOTAL' || /_SUBTOTAL$/.test(rowModel.canonicalCode || '')) return true;
  return /소계|합계/.test(normalizeText(rowModel.originalName));
}

function inheritableSectionAfter(rowModel) {
  if (!rowModel) return '';
  if (rowModel.canonicalCode === 'MATERIAL_SUBTOTAL') return '노무비';
  if (rowModel.canonicalCode === 'LABOR_SUBTOTAL') return '경비';
  if (isSectionClosingRow(rowModel)) return '';
  if (['재료비', '노무비', '경비', '별도비목'].includes(rowModel.section)) return rowModel.section;
  return '';
}

function sectionFromLabel(label, currentSection, context = {}) {
  const compact = normalizeText(label);
  if (/직접재료비|간접재료비|재료비소계|재료비$/.test(compact)) return '재료비';
  if (/직접노무비|간접노무비|노무비소계|노무비$/.test(compact)) return '노무비';
  if (/일반관리비/.test(compact)) return '일반관리비';
  if (/이윤/.test(compact)) return '이윤';
  if (/공급가액|부가가치세|도급액|총공사비|총원가|공사원가|공사비소계/.test(compact)) return '집계';
  if (/^경비$|경비소계|기타경비/.test(compact)) return '경비';
  if (/보험료|관리비|보전비|수수료|부금비|처리비|전력비|운반비/.test(compact)) {
    return currentSection === '경비' ? '경비' : '별도비목';
  }
  if (context.isTopLevelIndependent && !currentSection) return '별도비목';
  return currentSection || '';
}

function chooseLabel(rowItems, amountCols, rateCols, currentSection, amountItem) {
  const candidates = rowItems
    .filter((item) => !isNumericCell(item.cell))
    .filter((item) => !amountCols.includes(item.col))
    .filter((item) => !rateCols.includes(item.col))
    .filter((item) => !amountItem || item.col <= amountItem.col)
    .map((item) => ({ ...item, cleaned: cleanLabel(item.text) }))
    .filter((item) => !isWeakLabel(item.cleaned));

  const aggregateCodes = new Set([
    'WORK_COST',
    'TOTAL_COST',
    'SUPPLY_AMOUNT',
    'VAT',
    'CONTRACT_AMOUNT',
    'GRAND_TOTAL_CONSTRUCTION_COST',
    'CONSTRUCTION_COST_SUBTOTAL',
    'SUBTOTAL',
  ]);
  const canonicalCandidates = candidates
    .map((item) => ({ item, canonical: canonicalFor(item.cleaned, currentSection) }))
    .filter((entry) => entry.canonical)
    .sort((a, b) => {
      const aAggregate = aggregateCodes.has(a.canonical.code) ? 1 : 0;
      const bAggregate = aggregateCodes.has(b.canonical.code) ? 1 : 0;
      if (aAggregate !== bAggregate) return aAggregate - bAggregate;
      return a.item.col - b.item.col;
    });
  if (canonicalCandidates.length) return canonicalCandidates[0].item;
  return candidates.sort((a, b) => b.cleaned.length - a.cleaned.length || a.col - b.col)[0] || null;
}

function chooseAmount(rowItems, amountCols) {
  for (const col of amountCols) {
    const item = rowItems.find((candidate) => candidate.col === col && candidate.text !== '');
    if (item) return item;
  }
  const numeric = rowItems
    .filter((item) => isNumericCell(item.cell) && !looksLikeRate(item.cell))
    .sort((a, b) => b.col - a.col);
  return numeric[0] || null;
}

function chooseRate(rowItems, rateCols) {
  for (const col of rateCols) {
    const item = rowItems.find((candidate) => candidate.col === col && looksLikeRate(candidate.cell));
    if (item) return item;
  }
  return rowItems.find((item) => looksLikeRate(item.cell)) || null;
}

function extractReferences(formula) {
  if (!formula) return [];
  return parseFormulaReferences(formula, '')
    .map((ref) => formatReference(ref))
    .filter((ref, index, all) => ref && all.indexOf(ref) === index);
}

function formulaBody(formula) {
  return String(formula || '').replace(/^=/, '');
}

function normalizeAddress(value) {
  return String(value || '').replace(/\$/g, '');
}

function formatReference(ref) {
  if (!ref) return '';
  return ref.sheetName ? `${ref.sheetName}!${ref.address}` : ref.address;
}

function parseFormulaReferences(formula, currentSheetName) {
  const refs = [];
  const addr = '\\$?[A-Z]{1,3}\\$?\\d+(?::\\$?[A-Z]{1,3}\\$?\\d+)?';
  let masked = formulaBody(formula);

  masked = masked.replace(new RegExp(`'((?:[^']|'')+)'!\\s*(${addr})`, 'g'), (match, sheetName, addressText) => {
    refs.push({
      sheetName: sheetName.replace(/''/g, "'"),
      address: normalizeAddress(addressText),
      explicitSheet: true,
    });
    return ' '.repeat(match.length);
  });

  masked = masked.replace(
    new RegExp(`([A-Za-z0-9_\\uAC00-\\uD7A3][A-Za-z0-9_\\uAC00-\\uD7A3 .()_\\-]{0,90})!\\s*(${addr})`, 'g'),
    (match, sheetName, addressText) => {
      refs.push({
        sheetName: sheetName.trim(),
        address: normalizeAddress(addressText),
        explicitSheet: true,
      });
      return ' '.repeat(match.length);
    },
  );

  const localMatches = masked.match(new RegExp(addr, 'g')) || [];
  for (const addressText of localMatches) {
    refs.push({
      sheetName: currentSheetName,
      address: normalizeAddress(addressText),
      explicitSheet: false,
    });
  }

  return refs;
}

function expandReferenceAddress(addressText, maxCells = 40) {
  const normalized = normalizeAddress(addressText);
  if (!normalized.includes(':')) return [normalized];
  const range = XLSX.utils.decode_range(normalized);
  const count = (range.e.r - range.s.r + 1) * (range.e.c - range.s.c + 1);
  if (count > maxCells) return [normalized];
  const addresses = [];
  for (let row = range.s.r; row <= range.e.r; row += 1) {
    for (let col = range.s.c; col <= range.e.c; col += 1) {
      addresses.push(address(row, col));
    }
  }
  return addresses;
}

function cellEntry(workbook, sheetName, addressText, depth = 0, pathParts = []) {
  const ws = workbook.Sheets[sheetName];
  const normalized = normalizeAddress(addressText);
  const cell = ws?.[normalized];
  return {
    sheetName,
    address: normalized,
    source: `${sheetName}!${normalized}`,
    depth,
    display: cell ? display(ws, normalized) : '',
    value: cell ? toNumber(cell) : null,
    formula: cell?.f ? `=${cell.f}` : null,
    isNumeric: Boolean(cell && isNumericCell(cell)),
    isRate: Boolean(cell && looksLikeRate(cell)),
    path: pathParts,
  };
}

function labelsNearCell(workbook, sheetName, addressText) {
  const ws = workbook.Sheets[sheetName];
  if (!ws || !addressText || addressText.includes(':')) return '';
  const decoded = XLSX.utils.decode_cell(normalizeAddress(addressText));
  const labels = [];
  for (let col = Math.max(0, decoded.c - 6); col < decoded.c; col += 1) {
    const addr = address(decoded.r, col);
    const cell = ws[addr];
    if (!cell || isNumericCell(cell) || looksLikeRate(cell)) continue;
    const label = cleanLabel(display(ws, addr));
    if (!isWeakBaseLabel(label)) labels.push(label);
  }
  return labels.join(' + ');
}

function stripFormulaReferences(formula, currentSheetName) {
  let stripped = formulaBody(formula);
  for (const ref of parseFormulaReferences(formula, currentSheetName)) {
    const escapedSheet = ref.explicitSheet
      ? `(?:'${ref.sheetName.replace(/'/g, "''")}'|${ref.sheetName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')})!\\s*`
      : '';
    const escapedAddress = ref.address.replace(/\$/g, '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(':', ':\\$?');
    stripped = stripped.replace(new RegExp(`${escapedSheet}\\$?${escapedAddress.replace(/[A-Z]+/g, (m) => m.replace(/([A-Z])/g, '\\$?$1'))}`, 'g'), '');
  }
  return stripped.replace(/"(?:""|[^"])*"/g, '');
}

function numericLiterals(formula, currentSheetName) {
  const stripped = stripFormulaReferences(formula, currentSheetName);
  const matches = stripped.match(/(?<![A-Za-z0-9_.])[-+]?\d+(?:\.\d+)?%?(?![A-Za-z0-9_.])/g) || [];
  return matches
    .filter((value) => !['0', '1'].includes(value))
    .map((value) => ({ display: value, value: value.endsWith('%') ? Number(value.slice(0, -1)) / 100 : Number(value) }))
    .filter((item) => Number.isFinite(item.value));
}

function formulaTrace(workbook, sheetName, formula, maxDepth = 3, maxNodes = 90) {
  if (!formula) return [];
  const queue = parseFormulaReferences(formula, sheetName).map((ref) => ({
    ref,
    depth: 1,
    path: [`${sheetName}!{${formulaBody(formula)}}`, formatReference(ref)],
  }));
  const nodes = [];
  const visited = new Set();

  while (queue.length && nodes.length < maxNodes) {
    const current = queue.shift();
    const sheet = current.ref.sheetName || sheetName;
    for (const addr of expandReferenceAddress(current.ref.address)) {
      if (nodes.length >= maxNodes) break;
      const key = `${sheet}!${addr}`;
      if (visited.has(key)) continue;
      visited.add(key);
      const node = cellEntry(workbook, sheet, addr, current.depth, current.path);
      nodes.push(node);
      const ws = workbook.Sheets[sheet];
      const cell = ws?.[addr];
      if (cell?.f && current.depth < maxDepth) {
        for (const next of parseFormulaReferences(cell.f, sheet)) {
          queue.push({
            ref: next,
            depth: current.depth + 1,
            path: current.path.concat(formatReference(next)),
          });
        }
      }
    }
  }

  return nodes;
}

function uniqueCellRefs(refs, workbook, currentSheetName) {
  const seen = new Set();
  const cells = [];
  for (const ref of refs) {
    const sheet = ref.sheetName || currentSheetName;
    for (const addr of expandReferenceAddress(ref.address)) {
      const key = `${sheet}!${addr}`;
      if (seen.has(key)) continue;
      seen.add(key);
      const ws = workbook.Sheets[sheet];
      cells.push({ sheetName: sheet, address: addr, cell: ws?.[addr] || null });
    }
  }
  return cells;
}

function calculationCell(workbook, sheetName, amountItem) {
  if (!amountItem) return null;
  const amountCell = cellEntry(workbook, sheetName, amountItem.addr);
  if (!amountItem.formula) return amountCell;
  const refs = uniqueCellRefs(parseFormulaReferences(amountItem.formula, sheetName), workbook, sheetName);
  const formulaRefs = refs.filter((ref) => ref.cell?.f);
  if (refs.length === 1 && formulaRefs.length === 1) {
    return cellEntry(workbook, formulaRefs[0].sheetName, formulaRefs[0].address);
  }
  return amountCell;
}

function directFormulaInputs(workbook, calcCell) {
  if (!calcCell?.formula) return { baseAmounts: [], rateRefs: [], literals: [] };
  const refs = uniqueCellRefs(parseFormulaReferences(calcCell.formula, calcCell.sheetName), workbook, calcCell.sheetName);
  const baseAmounts = [];
  const rateRefs = [];
  const seenBases = new Set();
  const seenRates = new Set();

  for (const ref of refs) {
    if (!ref.cell || !isNumericCell(ref.cell)) continue;
    const key = `${ref.sheetName}!${ref.address}`;
    const entry = cellEntry(workbook, ref.sheetName, ref.address);
    entry.label = labelsNearCell(workbook, ref.sheetName, ref.address);
    if (looksLikeRate(ref.cell)) {
      if (!seenRates.has(key)) rateRefs.push(entry);
      seenRates.add(key);
    } else {
      if (!seenBases.has(key)) baseAmounts.push(entry);
      seenBases.add(key);
    }
  }

  return {
    baseAmounts,
    rateRefs,
    literals: numericLiterals(calcCell.formula, calcCell.sheetName),
  };
}

function rowTraceModel(workbook, statement, amountItem, rateItem, base) {
  const amountTrace = amountItem?.formula ? formulaTrace(workbook, statement.sheetName, amountItem.formula) : [];
  const calc = calculationCell(workbook, statement.sheetName, amountItem);
  const inputs = directFormulaInputs(workbook, calc);
  const rateTrace = rateItem?.formula ? formulaTrace(workbook, statement.sheetName, rateItem.formula, 2, 30) : [];
  const explicitRate = rateItem
    ? {
        address: rateItem.addr,
        sheetName: statement.sheetName,
        source: `${statement.sheetName}!${rateItem.addr}`,
        display: rateItem.text,
        value: rateValue(rateItem.cell),
        formula: rateItem.formula || null,
      }
    : null;
  const resolvedRate = inputs.rateRefs[0] || rateTrace.find((item) => item.isRate) || explicitRate;
  const literals = [
    ...(amountItem?.formula ? numericLiterals(amountItem.formula, statement.sheetName) : []),
    ...inputs.literals,
  ].filter((item, index, all) => all.findIndex((candidate) => candidate.display === item.display) === index);
  const sameRowBaseSources = new Set(base.baseAmounts.map((item) => `${statement.sheetName}!${item.address}`));
  const traceBaseAmounts = inputs.baseAmounts.filter((item) => {
    if (!amountItem) return true;
    return item.source !== `${statement.sheetName}!${amountItem.addr}` && !sameRowBaseSources.has(item.source);
  });
  const hasSameRowBase = base.baseAmounts.length > 0;
  const hasTraceBase = traceBaseAmounts.length > 0;
  const crossSheet = amountTrace.some((item) => item.sheetName && item.sheetName !== statement.sheetName);
  const resolutionStatus = hasSameRowBase
    ? 'SAME_ROW'
    : hasTraceBase && crossSheet
      ? 'CROSS_SHEET_TRACE'
      : hasTraceBase
        ? 'FORMULA_TRACE'
        : crossSheet
          ? 'CROSS_SHEET_UNRESOLVED'
          : amountItem?.formula
            ? 'FORMULA_UNRESOLVED'
            : 'UNRESOLVED';

  return {
    calculationCell: calc,
    formulaTrace: amountTrace,
    traceBaseAmounts,
    resolvedRate,
    literalAdjustments: literals,
    resolutionStatus,
  };
}

function baseInfo(rowItems, labelItem, amountItem, rateItem, amountCols, rateCols) {
  const used = new Set([labelItem?.addr, amountItem?.addr, rateItem?.addr].filter(Boolean));
  const maxBaseCol = rateItem?.col ?? amountItem?.col ?? Infinity;
  const textParts = [];
  const numberParts = [];
  for (const item of rowItems) {
    if (used.has(item.addr)) continue;
    if (item.col > maxBaseCol) continue;
    const candidateLabel = cleanLabel(item.text);
    const candidateCompact = normalizeText(candidateLabel);
    if (labelItem && item.col < labelItem.col && /^(재료비|노무비|경비|순공사원가|공사원가)$/.test(candidateCompact)) continue;
    if (amountCols.includes(item.col) || rateCols.includes(item.col)) continue;
    if (looksLikeRate(item.cell)) continue;
    if (isNumericCell(item.cell)) {
      const n = toNumber(item.cell);
      if (n != null) numberParts.push({ address: item.addr, value: n, display: item.text, formula: item.formula || null });
      continue;
    }
    const label = candidateLabel;
    if (!isWeakBaseLabel(label)) textParts.push(label);
  }
  return {
    baseLabel: textParts.join(' + '),
    baseAmounts: numberParts,
  };
}

function normalizedRow({ workbook, statement, rowIndex, labelItem, amountItem, rateItem, currentSection, rowItems, amountCols, rateCols, rangeStartCol }) {
  const originalName = labelItem.cleaned;
  const canonical = canonicalFor(originalName, currentSection);
  let section = sectionFromLabel(originalName, currentSection, {
    isTopLevelIndependent: isTopLevelIndependentRow(labelItem, rangeStartCol),
  });
  if (!section && canonical && amountItem) section = '별도비목';
  const base = baseInfo(rowItems, labelItem, amountItem, rateItem, amountCols, rateCols);
  const traced = rowTraceModel(workbook, statement, amountItem, rateItem, base);
  const formula = amountItem?.formula || null;
  const refs = extractReferences(formula || '');
  const mappingStatus = canonical ? canonical.status : 'UNMAPPED';
  const confidence = canonical ? (mappingStatus === 'EXACT' ? 0.98 : mappingStatus === 'SUBTOTAL' ? 0.85 : 0.92) : 0.42;
  return {
    rowId: `${statement.statementId}-r${String(rowIndex + 1).padStart(3, '0')}`,
    sourceRow: rowIndex + 1,
    level: /소계|합계|공사원가|총원가|공급가액|도급액|총공사비/.test(normalizeText(originalName)) ? 1 : 2,
    section,
    canonicalCode: canonical?.code || 'UNMAPPED',
    canonicalName: canonical?.name || '',
    originalName,
    baseLabel: base.baseLabel,
    baseAmounts: base.baseAmounts,
    traceBaseAmounts: traced.traceBaseAmounts,
    rate: rateItem
      ? {
          address: rateItem.addr,
          display: rateItem.text,
          value: rateValue(rateItem.cell),
          formula: rateItem.formula || null,
      }
      : null,
    resolvedRate: traced.resolvedRate,
    amount: amountItem
      ? {
          address: amountItem.addr,
          display: amountItem.text,
          value: toNumber(amountItem.cell),
          formula,
          references: refs,
        }
      : null,
    calculationCell: traced.calculationCell,
    formulaTrace: traced.formulaTrace,
    literalAdjustments: traced.literalAdjustments,
    resolutionStatus: traced.resolutionStatus,
    note: rowItems
      .filter((item) => item.col > (amountItem?.col || 0))
      .map((item) => cleanLabel(item.text))
      .filter(Boolean)
      .filter((item) => item !== originalName)
      .join(' / '),
    source: {
      fileName: statement.fileName,
      sheetName: statement.sheetName,
      labelCell: labelItem.addr,
      amountCell: amountItem?.addr || null,
      rateCell: rateItem?.addr || null,
    },
    mappingStatus,
    confidence,
  };
}

function normalizeSheet(file, workbook, sheetName) {
  const ws = workbook.Sheets[sheetName];
  const range = getRange(ws);
  const statement = {
    statementId: `${file.id}-${normalizeText(sheetName) || 'sheet'}`,
    fileName: path.basename(file.path),
    filePath: file.path,
    sheetName,
    detectedRange: ws['!ref'] || '',
    standardColumns: [
      'section',
      'canonicalCode',
      'canonicalName',
      'originalName',
      'baseLabel',
      'baseAmounts',
      'traceBaseAmounts',
      'rate',
      'resolvedRate',
      'amount',
      'calculationCell',
      'formula',
      'formulaTrace',
      'literalAdjustments',
      'resolutionStatus',
      'sourceCell',
      'note',
      'mappingStatus',
      'confidence',
    ],
    rows: [],
    diagnostics: {
      skippedHeaderRows: [],
      amountColumns: [],
      rateColumns: [],
      unmappedCount: 0,
    },
  };
  if (!range) return statement;

  const headerRows = findHeaderRows(ws, range);
  const amountCols = inferAmountColumns(ws, range, headerRows);
  const rateCols = inferRateColumns(ws, range, headerRows).filter((col) => !amountCols.includes(col));
  statement.diagnostics.amountColumns = amountCols.map(columnName);
  statement.diagnostics.rateColumns = rateCols.map(columnName);

  let currentSection = '';
  for (let row = range.s.r; row <= range.e.r; row += 1) {
    const items = rowCells(ws, row, range.s.c, range.e.c);
    const line = items.map((item) => item.text).join(' ');
    if (isHeaderOrTitleRow(line)) {
      if (line.trim()) statement.diagnostics.skippedHeaderRows.push(row + 1);
      continue;
    }
    const amountItem = chooseAmount(items, amountCols);
    const rateItem = chooseRate(items, rateCols);
    const labelItem = chooseLabel(items, amountCols, rateCols, currentSection, amountItem);
    if (!labelItem) continue;

    const canonical = canonicalFor(labelItem.cleaned, currentSection);
    if (!amountItem && !canonical && !rateItem) continue;

    const rowModel = normalizedRow({
      workbook,
      statement,
      rowIndex: row,
      labelItem,
      amountItem,
      rateItem,
      currentSection,
      rowItems: items,
      amountCols,
      rateCols,
      rangeStartCol: range.s.c,
    });
    currentSection = inheritableSectionAfter(rowModel);
    statement.rows.push(rowModel);
  }
  statement.diagnostics.unmappedCount = statement.rows.filter((row) => row.mappingStatus === 'UNMAPPED').length;
  statement.diagnostics.referenceSheets = Object.entries(
    statement.rows.reduce((acc, row) => {
      for (const item of row.formulaTrace || []) {
        if (!item.sheetName || item.sheetName === statement.sheetName) continue;
        acc[item.sheetName] = (acc[item.sheetName] || 0) + 1;
      }
      return acc;
    }, {}),
  )
    .sort((a, b) => b[1] - a[1])
    .map(([sheetName, count]) => ({ sheetName, count }));
  statement.summary = {
    rowCount: statement.rows.length,
    mappedCount: statement.rows.filter((row) => row.mappingStatus !== 'UNMAPPED').length,
    unmappedCount: statement.diagnostics.unmappedCount,
    formulaAmountCount: statement.rows.filter((row) => row.amount?.formula).length,
    tracedBaseCount: statement.rows.filter((row) => (row.baseAmounts?.length || row.traceBaseAmounts?.length)).length,
    crossSheetTraceCount: statement.rows.filter((row) => (row.formulaTrace || []).some((item) => item.sheetName && item.sheetName !== statement.sheetName)).length,
    literalAdjustmentCount: statement.rows.filter((row) => row.literalAdjustments?.length).length,
    rateCount: statement.rows.filter((row) => row.rate).length,
    amountCount: statement.rows.filter((row) => row.amount).length,
  };
  return statement;
}

function normalizeWorkbook(file) {
  const workbook = XLSX.readFile(file.path, {
    cellFormula: true,
    cellNF: true,
    cellStyles: true,
    cellDates: true,
  });
  const costSheets = findCostSheetNames(workbook);
  return {
    fileId: file.id,
    fileName: path.basename(file.path),
    filePath: file.path,
    costSheetNames: costSheets,
    statements: costSheets.map((sheetName) => normalizeSheet(file, workbook, sheetName)),
  };
}

function esc(value) {
  return String(value ?? '').replace(/[&<>"']/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
}

function money(value) {
  if (value == null || !Number.isFinite(Number(value))) return '';
  return Number(value).toLocaleString('ko-KR');
}

function pct(value) {
  if (value == null || !Number.isFinite(Number(value))) return '';
  return `${(Number(value) * 100).toLocaleString('ko-KR', { maximumFractionDigits: 4 })}%`;
}

function jsonForScript(data) {
  return JSON.stringify(data).replace(/</g, '\\u003c');
}

function html(data) {
  const scriptJson = jsonForScript(data);
  return `<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>표준 원가계산서 정규화 리포트</title>
  <style>
    :root {
      --bg: #f4f6f8;
      --panel: #ffffff;
      --ink: #172033;
      --muted: #667085;
      --line: #d8dee8;
      --head: #eef3f7;
      --accent: #0f766e;
      --soft: #dff4ef;
      --warn: #fff7d6;
      --bad: #fee2e2;
      --ok: #dcfce7;
      --unknown: #f1f5f9;
      font-family: "Segoe UI", "Malgun Gothic", Arial, sans-serif;
    }
    * { box-sizing: border-box; }
    body { margin: 0; background: var(--bg); color: var(--ink); }
    .shell { min-height: 100vh; display: grid; grid-template-columns: 330px minmax(0, 1fr); }
    aside { background: #202938; color: white; padding: 20px 16px; overflow: auto; }
    main { padding: 18px; min-width: 0; display: grid; gap: 14px; grid-template-rows: auto auto minmax(0, 1fr); }
    h1 { margin: 0 0 8px; font-size: 20px; }
    h2 { margin: 18px 0 10px; font-size: 15px; }
    p { margin: 0; color: var(--muted); line-height: 1.5; }
    aside p { color: #cbd5e1; }
    .metrics { display: grid; grid-template-columns: repeat(2, 1fr); gap: 8px; margin: 16px 0; }
    .metric { border: 1px solid rgba(255,255,255,.16); background: rgba(255,255,255,.06); border-radius: 8px; padding: 10px; }
    .metric span { display: block; font-size: 12px; color: #cbd5e1; }
    .metric strong { font-size: 18px; }
    .nav { display: grid; gap: 8px; }
    .nav button {
      border: 1px solid rgba(255,255,255,.14); background: rgba(255,255,255,.05); color: white; border-radius: 8px;
      padding: 10px; text-align: left; cursor: pointer; font: inherit;
    }
    .nav button.active { background: var(--accent); border-color: #5eead4; }
    .nav strong { display: block; font-size: 13px; overflow-wrap: anywhere; }
    .nav span { display: block; color: #cbd5e1; margin-top: 4px; font-size: 12px; }
    .toolbar, .cards, .table-card { background: var(--panel); border: 1px solid var(--line); border-radius: 8px; }
    .toolbar { padding: 12px; display: flex; gap: 10px; flex-wrap: wrap; align-items: center; }
    .toolbar label { font-size: 13px; color: var(--muted); display: flex; gap: 6px; align-items: center; }
    select, input { border: 1px solid var(--line); border-radius: 6px; min-height: 36px; padding: 7px 10px; background: white; color: var(--ink); }
    input[type="search"] { min-width: 280px; }
    .cards { padding: 12px; display: grid; grid-template-columns: repeat(6, minmax(120px, 1fr)); gap: 10px; }
    .card { border: 1px solid var(--line); border-radius: 8px; padding: 10px; background: #fbfcfd; }
    .card span { display: block; font-size: 12px; color: var(--muted); }
    .card strong { font-size: 17px; }
    .table-card { overflow: auto; }
    table { border-collapse: separate; border-spacing: 0; min-width: 1780px; width: 100%; font-size: 12px; }
    th, td { border-right: 1px solid var(--line); border-bottom: 1px solid var(--line); padding: 7px 8px; vertical-align: top; }
    th { position: sticky; top: 0; z-index: 2; background: var(--head); color: #475467; text-align: left; }
    td.num { text-align: right; font-variant-numeric: tabular-nums; }
    td.small { color: var(--muted); }
    tr:hover td { background: #f8fbff; }
    .status { display: inline-flex; border-radius: 999px; padding: 3px 7px; font-size: 11px; border: 1px solid var(--line); }
    .resolve { display: inline-flex; border-radius: 999px; padding: 3px 7px; font-size: 11px; border: 1px solid var(--line); background: var(--unknown); white-space: nowrap; }
    .SAME_ROW { background: var(--ok); }
    .FORMULA_TRACE, .CROSS_SHEET_TRACE { background: var(--soft); }
    .CROSS_SHEET_UNRESOLVED, .FORMULA_UNRESOLVED { background: var(--warn); }
    .UNRESOLVED { background: var(--bad); }
    .EXACT { background: var(--ok); }
    .ALIAS { background: var(--soft); }
    .SUBTOTAL { background: var(--warn); }
    .UNMAPPED { background: var(--bad); }
    code { font-family: Consolas, "SFMono-Regular", monospace; font-size: 11px; }
    .source { color: #155e75; }
    .formula { max-width: 260px; overflow-wrap: anywhere; }
    .trace { max-width: 280px; overflow-wrap: anywhere; }
    .empty { padding: 40px; text-align: center; color: var(--muted); }
    @media (max-width: 1100px) {
      .shell { grid-template-columns: 1fr; }
      aside { max-height: 40vh; }
      .cards { grid-template-columns: repeat(2, minmax(120px, 1fr)); }
    }
  </style>
</head>
<body>
  <div class="shell">
    <aside>
      <h1>표준 원가계산서 정규화</h1>
      <p>원본 시트 모양과 관계없이 동일한 컬럼 구조로 추출한 검수 리포트입니다.</p>
      <div class="metrics" id="globalMetrics"></div>
      <h2>원가계산서</h2>
      <div class="nav" id="statementNav"></div>
    </aside>
    <main>
      <section class="toolbar">
        <label>원가계산서 <select id="statementSelect"></select></label>
        <label>검색 <input type="search" id="search" placeholder="항목명, 코드, 셀, 수식 검색" /></label>
        <label>매핑상태 <select id="statusFilter"><option value="ALL">전체</option><option value="EXACT">EXACT</option><option value="ALIAS">ALIAS</option><option value="SUBTOTAL">SUBTOTAL</option><option value="UNMAPPED">UNMAPPED</option></select></label>
      </section>
      <section class="cards" id="statementCards"></section>
      <section class="table-card" id="tableWrap"></section>
    </main>
  </div>
  <script>
    const DATA = ${scriptJson};
    const statements = DATA.workbooks.flatMap((wb) => wb.statements.map((statement) => ({ workbook: wb, statement })));
    const state = { index: 0, search: '', status: 'ALL' };
    const $ = (id) => document.getElementById(id);
    const esc = (value) => String(value ?? '').replace(/[&<>"']/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
    const fmt = (value) => value == null || Number.isNaN(Number(value)) ? '' : Number(value).toLocaleString('ko-KR');
    const pct = (value) => value == null || Number.isNaN(Number(value)) ? '' : (Number(value) * 100).toLocaleString('ko-KR', { maximumFractionDigits: 4 }) + '%';
    function renderGlobal() {
      const totals = statements.reduce((acc, item) => {
        acc.rows += item.statement.summary.rowCount;
        acc.mapped += item.statement.summary.mappedCount;
        acc.unmapped += item.statement.summary.unmappedCount;
        acc.formulas += item.statement.summary.formulaAmountCount;
        acc.traced += item.statement.summary.tracedBaseCount;
        acc.cross += item.statement.summary.crossSheetTraceCount;
        acc.literals += item.statement.summary.literalAdjustmentCount;
        return acc;
      }, { rows: 0, mapped: 0, unmapped: 0, formulas: 0, traced: 0, cross: 0, literals: 0 });
      $('globalMetrics').innerHTML = [
        ['파일', DATA.workbooks.length],
        ['원가시트', statements.length],
        ['표준행', totals.rows],
        ['매핑', totals.mapped],
        ['미매핑', totals.unmapped],
        ['수식금액', totals.formulas],
        ['기초확인', totals.traced],
        ['타시트추적', totals.cross],
        ['수식상수', totals.literals],
      ].map(([label, value]) => '<div class="metric"><span>' + label + '</span><strong>' + fmt(value) + '</strong></div>').join('');
    }
    function renderNav() {
      $('statementNav').innerHTML = statements.map((item, index) => {
        const active = index === state.index ? ' active' : '';
        return '<button class="' + active + '" data-index="' + index + '"><strong>' + esc(item.statement.fileName) + '</strong><span>' + esc(item.statement.sheetName) + ' · 행 ' + fmt(item.statement.summary.rowCount) + ' · 매핑 ' + fmt(item.statement.summary.mappedCount) + '</span></button>';
      }).join('');
      $('statementSelect').innerHTML = statements.map((item, index) => '<option value="' + index + '">' + esc(item.statement.fileName + ' / ' + item.statement.sheetName) + '</option>').join('');
      $('statementSelect').value = String(state.index);
      document.querySelectorAll('[data-index]').forEach((button) => button.addEventListener('click', () => {
        state.index = Number(button.dataset.index);
        renderAll();
      }));
    }
    function active() { return statements[state.index].statement; }
    function renderCards() {
      const s = active();
      const cards = [
        ['파일', s.fileName],
        ['시트', s.sheetName],
        ['원본범위', s.detectedRange],
        ['표준행', s.summary.rowCount],
        ['금액열', s.diagnostics.amountColumns.join(', ') || '-'],
        ['요율열', s.diagnostics.rateColumns.join(', ') || '-'],
        ['참조시트', (s.diagnostics.referenceSheets || []).map((item) => item.sheetName).join(', ') || '-'],
      ];
      $('statementCards').innerHTML = cards.map(([label, value]) => '<div class="card"><span>' + esc(label) + '</span><strong>' + esc(value) + '</strong></div>').join('');
    }
    function rowMatches(row) {
      if (state.status !== 'ALL' && row.mappingStatus !== state.status) return false;
      if (!state.search) return true;
      const hay = [
        row.section, row.canonicalCode, row.canonicalName, row.originalName, row.baseLabel,
        row.amount?.address, row.amount?.display, row.amount?.formula, row.rate?.display, row.note,
        row.calculationCell?.source, row.resolvedRate?.source, row.resolutionStatus,
        ...(row.traceBaseAmounts || []).map((item) => item.source + ' ' + item.display + ' ' + (item.label || '')),
        ...(row.formulaTrace || []).map((item) => item.source + ' ' + (item.formula || '')),
      ].join(' ').toLowerCase();
      return hay.includes(state.search.toLowerCase());
    }
    function renderTable() {
      const rows = active().rows.filter(rowMatches);
      if (!rows.length) {
        $('tableWrap').innerHTML = '<div class="empty">조건에 맞는 행이 없습니다.</div>';
        return;
      }
      const body = rows.map((row) => {
        const sameRowBases = (row.baseAmounts || []).map((b) => esc(b.display) + ' <code>[' + esc(b.address) + ']</code>');
        const tracedBases = (row.traceBaseAmounts || []).map((b) => {
          const label = b.label ? esc(b.label) + '<br>' : '';
          return label + esc(b.display) + ' <code>[' + esc(b.source) + ']</code>';
        });
        const baseAmounts = sameRowBases.concat(tracedBases).join('<br>');
        const refs = (row.amount?.references || []).join(', ');
        const calc = row.calculationCell ? row.calculationCell.source + (row.calculationCell.source === (row.source.sheetName + '!' + row.source.amountCell) ? '' : ' ← ' + row.source.sheetName + '!' + row.source.amountCell) : '';
        const rateSource = row.resolvedRate ? (row.resolvedRate.display || pct(row.resolvedRate.value) || '') + ' [' + row.resolvedRate.source + ']' : '';
        const literals = (row.literalAdjustments || []).map((item) => item.display).join(', ');
        const trace = (row.formulaTrace || []).slice(0, 5).map((item) => item.source + (item.formula ? ' ' + item.formula : '')).join('<br>');
        return '<tr>' +
          '<td class="num">' + row.sourceRow + '</td>' +
          '<td>' + esc(row.section) + '</td>' +
          '<td><code>' + esc(row.canonicalCode) + '</code></td>' +
          '<td>' + esc(row.canonicalName) + '</td>' +
          '<td>' + esc(row.originalName) + '</td>' +
          '<td>' + esc(row.baseLabel) + '</td>' +
          '<td class="num">' + baseAmounts + '</td>' +
          '<td class="num">' + esc(row.rate?.display || pct(row.rate?.value) || '') + '</td>' +
          '<td class="num"><strong>' + esc(row.amount?.display || '') + '</strong></td>' +
          '<td class="source">' + esc(calc) + '</td>' +
          '<td class="source">' + esc(rateSource) + '</td>' +
          '<td><span class="resolve ' + row.resolutionStatus + '">' + esc(row.resolutionStatus) + '</span></td>' +
          '<td class="small">' + esc(literals) + '</td>' +
          '<td class="formula"><code>' + esc(row.amount?.formula || '') + '</code></td>' +
          '<td class="trace"><code>' + trace + '</code></td>' +
          '<td class="source">' + esc(row.source.sheetName + '!' + (row.source.amountCell || row.source.labelCell)) + '</td>' +
          '<td class="small">' + esc(refs) + '</td>' +
          '<td>' + esc(row.note) + '</td>' +
          '<td><span class="status ' + row.mappingStatus + '">' + row.mappingStatus + '</span></td>' +
          '<td class="num">' + Math.round(row.confidence * 100) + '%</td>' +
        '</tr>';
      }).join('');
      $('tableWrap').innerHTML = '<table><thead><tr>' +
        '<th>원본행</th><th>구간</th><th>표준코드</th><th>표준항목</th><th>원본항목</th><th>산출기초명</th><th>산출기초값</th><th>요율</th><th>금액</th><th>계산셀</th><th>요율출처</th><th>해결상태</th><th>수식상수</th><th>금액수식</th><th>추적경로</th><th>원본위치</th><th>참조셀</th><th>비고</th><th>매핑</th><th>신뢰도</th>' +
        '</tr></thead><tbody>' + body + '</tbody></table>';
    }
    function renderAll() {
      renderGlobal();
      renderNav();
      renderCards();
      renderTable();
    }
    $('statementSelect').addEventListener('change', (event) => { state.index = Number(event.target.value); renderAll(); });
    $('search').addEventListener('input', (event) => { state.search = event.target.value.trim(); renderTable(); });
    $('statusFilter').addEventListener('change', (event) => { state.status = event.target.value; renderTable(); });
    renderAll();
  </script>
</body>
</html>`;
}

const workbooks = inputFiles.map(normalizeWorkbook);
const payload = {
  schemaVersion: 'StandardCostStatement.v0.1',
  generatedAt: new Date().toISOString(),
  standardRowColumns: [
    'sourceRow',
    'section',
    'canonicalCode',
    'canonicalName',
    'originalName',
    'baseLabel',
    'baseAmounts',
    'traceBaseAmounts',
    'rate',
    'resolvedRate',
    'amount',
    'calculationCell',
    'formula',
    'formulaTrace',
    'literalAdjustments',
    'resolutionStatus',
    'source',
    'references',
    'note',
    'mappingStatus',
    'confidence',
  ],
  workbooks,
};

fs.mkdirSync(outputDir, { recursive: true });
fs.writeFileSync(jsonOutput, `${JSON.stringify(payload, null, 2)}\n`, 'utf8');
fs.writeFileSync(htmlOutput, html(payload), 'utf8');

const statementCount = workbooks.reduce((sum, wb) => sum + wb.statements.length, 0);
const rowCount = workbooks.reduce((sum, wb) => sum + wb.statements.reduce((s, st) => s + st.summary.rowCount, 0), 0);
const mappedCount = workbooks.reduce((sum, wb) => sum + wb.statements.reduce((s, st) => s + st.summary.mappedCount, 0), 0);
const unmappedCount = workbooks.reduce((sum, wb) => sum + wb.statements.reduce((s, st) => s + st.summary.unmappedCount, 0), 0);

console.log(
  JSON.stringify(
    {
      htmlOutput,
      jsonOutput,
      workbookCount: workbooks.length,
      statementCount,
      rowCount,
      mappedCount,
      unmappedCount,
    },
    null,
    2,
  ),
);
