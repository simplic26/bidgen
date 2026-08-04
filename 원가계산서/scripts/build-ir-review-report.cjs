const fs = require('node:fs');
const path = require('node:path');
const XLSX = require('xlsx');

const base = 'C:/Users/user/Documents/AI Pioneer/7\uC8FC\uCC28/\uB808\uD37C\uB7F0\uC2A4 \uC5C4\uC120';
const inputFiles = [
  `${base}/3. 2026\uB144 \uD3C9\uD0DD\uC9C0\uC0AC \uC5F4\uC218\uC1A1\uAD00\uACF5\uC0AC \uC124\uACC4\uB0B4\uC5ED\uC11C.xls`,
  `${base}/2. \uCD94\uC815\uAC00\uACA9\uB0B4\uC5ED\uC11C.xlsx`,
  `${base}/8. \uC124\uACC4\uB0B4\uC5ED\uC11C(\uCCAD\uC8FC\uC9C0\uC0AC).xlsx`,
];

const defaultOutput =
  'C:/Users/user/.codex/visualizations/2026/07/22/019f8730-deb9-7402-b177-88629938069c/cost-ir-review.html';
const outputPath = process.argv[2] || defaultOutput;

const REFERENCE_PATTERN = /(?:(?:'[^']+'|[A-Za-z0-9_\uAC00-\uD7A3]+)!)?\$?[A-Z]{1,3}\$?[0-9]+/g;

function classifySheetRole(sheetName) {
  if (/\uD45C\uC9C0|\uBAA9\uCC28|\uC694\uC57D/.test(sheetName)) return 'COVER_SUMMARY';
  if (/\uC6D0\uAC00|\uACC4\uC0B0|\uCD94\uC815\uAC00\uACA9/.test(sheetName)) return 'COST_SUMMARY';
  if (/\uC77C\uC704\uB300\uAC00/.test(sheetName)) return 'UNIT_PRICE';
  if (/\uB0B4\uC5ED/.test(sheetName)) return 'CONSTRUCTION_ITEMS';
  if (/\uC218\uB7C9|\uBB3C\uB7C9/.test(sheetName)) return 'QUANTITY';
  if (/\uB178\uC784|\uC784\uAE08/.test(sheetName)) return 'WAGE_RATE';
  if (/\uC81C\uBE44\uC728|\uC694\uC728\uAE30\uC900/.test(sheetName)) return 'RATE_STANDARD';
  if (/\uAC00\uACA9\uC870\uC0AC|\uB2E8\uAC00/.test(sheetName)) return 'PRICE_SURVEY';
  return 'OTHER';
}

function dataType(cell) {
  if (cell.f) return 'FORMULA';
  if (cell.t === 'n') return 'NUMBER';
  if (cell.t === 's' || cell.t === 'str') return 'STRING';
  if (cell.t === 'b') return 'BOOLEAN';
  if (cell.t === 'd') return 'DATE';
  if (cell.t === 'e') return 'ERROR';
  return 'BLANK';
}

function cellValueForJson(value) {
  if (value == null) return null;
  if (value instanceof Date) return value.toISOString();
  if (typeof value === 'number' || typeof value === 'string' || typeof value === 'boolean') return value;
  return String(value);
}

function extractReferences(formula) {
  const matches = formula.match(REFERENCE_PATTERN);
  return matches ? Array.from(new Set(matches)) : [];
}

function encodeCell(row, col) {
  return XLSX.utils.encode_cell({ r: row, c: col });
}

function convertWorkbook(filePath) {
  const workbook = XLSX.readFile(filePath, {
    cellFormula: true,
    cellNF: true,
    cellStyles: true,
    cellDates: true,
  });
  const sheets = workbook.SheetNames.map((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    const ref = worksheet['!ref'] || null;
    const bounds = ref ? XLSX.utils.decode_range(ref) : null;
    const merges = worksheet['!merges'] || [];
    const rows = worksheet['!rows'] || [];
    const cols = worksheet['!cols'] || [];
    const mergeAnchors = new Map();
    const mergeCovers = new Set();

    for (const merge of merges) {
      const anchor = encodeCell(merge.s.r, merge.s.c);
      mergeAnchors.set(anchor, {
        range: XLSX.utils.encode_range(merge),
        rowSpan: merge.e.r - merge.s.r + 1,
        colSpan: merge.e.c - merge.s.c + 1,
      });
      for (let r = merge.s.r; r <= merge.e.r; r += 1) {
        for (let c = merge.s.c; c <= merge.e.c; c += 1) {
          const address = encodeCell(r, c);
          if (address !== anchor) mergeCovers.add(address);
        }
      }
    }

    let formulaCount = 0;
    const cells = {};
    for (const key of Object.keys(worksheet)) {
      if (key.charAt(0) === '!') continue;
      const cell = worksheet[key];
      const decoded = XLSX.utils.decode_cell(key);
      const kind = dataType(cell);
      if (kind === 'FORMULA') formulaCount += 1;
      const merge = mergeAnchors.get(key) || null;
      cells[key] = {
        address: key,
        row: decoded.r,
        col: decoded.c,
        dataType: kind,
        rawValue: cell.f ? `=${cell.f}` : cellValueForJson(cell.v),
        cachedValue: cellValueForJson(cell.v),
        displayValue: cell.w || (cell.v == null ? '' : String(cell.v)),
        numberFormat: cell.z || null,
        mergedRange: merge ? merge.range : null,
        rowSpan: merge ? merge.rowSpan : 1,
        colSpan: merge ? merge.colSpan : 1,
        hidden: Boolean(rows[decoded.r]?.hidden || cols[decoded.c]?.hidden),
        references: cell.f ? extractReferences(cell.f) : [],
      };
    }

    return {
      sheetName,
      sheetRole: classifySheetRole(sheetName),
      ref,
      startRow: bounds ? bounds.s.r : 0,
      startCol: bounds ? bounds.s.c : 0,
      rowCount: bounds ? bounds.e.r - bounds.s.r + 1 : 0,
      columnCount: bounds ? bounds.e.c - bounds.s.c + 1 : 0,
      cellCount: Object.keys(cells).length,
      formulaCount,
      mergeCount: merges.length,
      cells,
      mergeCovers: Array.from(mergeCovers),
    };
  });

  return {
    fileName: path.basename(filePath),
    filePath,
    sheetCount: sheets.length,
    totals: {
      cellCount: sheets.reduce((sum, sheet) => sum + sheet.cellCount, 0),
      formulaCount: sheets.reduce((sum, sheet) => sum + sheet.formulaCount, 0),
      mergeCount: sheets.reduce((sum, sheet) => sum + sheet.mergeCount, 0),
    },
    sheets,
  };
}

function jsonForScript(data) {
  return JSON.stringify(data).replace(/</g, '\\u003c');
}

function htmlDocument(data) {
  const generatedAt = new Date().toISOString();
  const json = jsonForScript(data);
  return `<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>원가계산서 IR 표 검수 리포트</title>
  <style>
    :root {
      color-scheme: light;
      --bg: #f6f7f9;
      --panel: #ffffff;
      --ink: #1d2433;
      --muted: #667085;
      --line: #d6dbe4;
      --head: #eef2f7;
      --accent: #0f766e;
      --accent-soft: #d9f3ee;
      --formula: #fff4d6;
      --merged: #e9e7ff;
      --hidden: #f3f4f6;
      --error: #fee2e2;
      --select: #dbeafe;
      font-family: "Segoe UI", "Malgun Gothic", Arial, sans-serif;
    }
    * { box-sizing: border-box; }
    body { margin: 0; background: var(--bg); color: var(--ink); }
    .shell { min-height: 100vh; display: grid; grid-template-columns: 320px minmax(0, 1fr) 360px; }
    aside, main, .detail { min-width: 0; }
    aside { background: #202938; color: #fff; padding: 20px 16px; overflow: auto; }
    h1 { font-size: 20px; margin: 0 0 8px; letter-spacing: 0; }
    h2 { font-size: 15px; margin: 20px 0 10px; letter-spacing: 0; }
    p { margin: 0; color: var(--muted); line-height: 1.5; }
    aside p { color: #c9d3e1; }
    .summary-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 8px; margin-top: 16px; }
    .metric { border: 1px solid rgba(255,255,255,.16); border-radius: 8px; padding: 10px; background: rgba(255,255,255,.06); }
    .metric span { display: block; color: #c9d3e1; font-size: 12px; }
    .metric strong { font-size: 18px; }
    .workbook-list, .sheet-list { display: grid; gap: 8px; margin-top: 12px; }
    button { font: inherit; }
    .nav-button {
      width: 100%; text-align: left; border: 1px solid rgba(255,255,255,.14); border-radius: 8px;
      background: rgba(255,255,255,.05); color: #fff; padding: 10px; cursor: pointer;
    }
    .nav-button.is-active { background: #0f766e; border-color: #5eead4; }
    .nav-button strong { display: block; font-size: 13px; overflow-wrap: anywhere; }
    .nav-button span { display: block; color: #c9d3e1; font-size: 12px; margin-top: 4px; }
    main { padding: 18px; overflow: hidden; display: grid; grid-template-rows: auto auto minmax(0, 1fr); gap: 12px; }
    .toolbar, .sheet-meta, .detail { background: var(--panel); border: 1px solid var(--line); border-radius: 8px; }
    .toolbar { padding: 12px; display: flex; flex-wrap: wrap; gap: 10px; align-items: center; }
    .toolbar label { color: var(--muted); font-size: 13px; display: flex; align-items: center; gap: 6px; }
    select, input {
      border: 1px solid var(--line); border-radius: 6px; padding: 8px 10px; min-height: 36px; background: #fff; color: var(--ink);
    }
    input[type="search"] { min-width: 260px; }
    .segmented { display: inline-flex; border: 1px solid var(--line); border-radius: 7px; overflow: hidden; background: #fff; }
    .segmented button { border: 0; background: transparent; padding: 8px 12px; cursor: pointer; }
    .segmented button.is-active { background: var(--accent); color: #fff; }
    .sheet-meta { padding: 12px; display: flex; flex-wrap: wrap; gap: 10px; align-items: center; }
    .pill { border: 1px solid var(--line); border-radius: 999px; padding: 5px 9px; font-size: 12px; color: var(--muted); background: #fff; }
    .pill.strong { color: #0f766e; background: var(--accent-soft); border-color: #99ded1; }
    .grid-wrap { background: var(--panel); border: 1px solid var(--line); border-radius: 8px; overflow: auto; }
    table.sheet-grid { border-collapse: separate; border-spacing: 0; min-width: 100%; font-size: 12px; }
    .sheet-grid th, .sheet-grid td { border-right: 1px solid var(--line); border-bottom: 1px solid var(--line); height: 28px; padding: 4px 7px; white-space: nowrap; max-width: 280px; overflow: hidden; text-overflow: ellipsis; }
    .sheet-grid thead th { position: sticky; top: 0; z-index: 3; background: var(--head); text-align: center; color: #445066; font-weight: 600; }
    .row-head { position: sticky; left: 0; z-index: 2; background: var(--head); text-align: right; color: #445066; width: 52px; }
    .corner { left: 0; z-index: 4 !important; }
    .cell { background: #fff; cursor: pointer; }
    .cell.is-formula { background: var(--formula); }
    .cell.is-merged { background: var(--merged); }
    .cell.is-hidden { background: var(--hidden); color: #777; }
    .cell.is-error { background: var(--error); }
    .cell.is-match { outline: 2px solid #2563eb; outline-offset: -2px; }
    .cell.is-selected { background: var(--select); outline: 2px solid #1d4ed8; outline-offset: -2px; }
    .detail { border-radius: 0; border-width: 0 0 0 1px; padding: 18px; overflow: auto; }
    .detail h2 { margin-top: 0; font-size: 18px; }
    .kv { display: grid; grid-template-columns: 120px minmax(0, 1fr); border-top: 1px solid var(--line); }
    .kv div { padding: 9px 0; border-bottom: 1px solid var(--line); min-width: 0; overflow-wrap: anywhere; }
    .kv div:nth-child(odd) { color: var(--muted); padding-right: 10px; }
    code { font-family: Consolas, "SFMono-Regular", monospace; font-size: 12px; }
    .legend { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 12px; }
    .legend span { font-size: 12px; border-radius: 6px; padding: 5px 7px; border: 1px solid var(--line); }
    .legend .formula { background: var(--formula); }
    .legend .merged { background: var(--merged); }
    .legend .hidden { background: var(--hidden); }
    .empty { padding: 40px; text-align: center; color: var(--muted); }
    @media (max-width: 1100px) {
      .shell { grid-template-columns: 280px minmax(0, 1fr); }
      .detail { grid-column: 1 / -1; border: 1px solid var(--line); border-radius: 8px; margin: 0 18px 18px; }
    }
  </style>
</head>
<body>
  <div class="shell">
    <aside>
      <h1>IR 표 검수 리포트</h1>
      <p>생성시각 ${generatedAt}. Excel을 고정 JSON 구조로 변환한 뒤, 시트별 격자로 다시 보여주는 검수 화면입니다.</p>
      <div class="summary-grid" id="globalSummary"></div>
      <h2>파일</h2>
      <div class="workbook-list" id="workbookList"></div>
      <h2>시트</h2>
      <div class="sheet-list" id="sheetList"></div>
    </aside>
    <main>
      <section class="toolbar">
        <label>시트 <select id="sheetSelect"></select></label>
        <label>검색 <input id="searchInput" type="search" placeholder="셀 값, 수식, 주소 검색" /></label>
        <div class="segmented" aria-label="표시 모드">
          <button id="valueMode" class="is-active" type="button">표시값</button>
          <button id="formulaMode" type="button">수식</button>
        </div>
      </section>
      <section class="sheet-meta" id="sheetMeta"></section>
      <section class="grid-wrap" id="gridWrap"></section>
    </main>
    <section class="detail">
      <h2>셀 상세</h2>
      <p id="detailHelp">셀을 클릭하면 변환된 JSON 필드가 여기 표시됩니다.</p>
      <div class="kv" id="cellDetail"></div>
      <div class="legend">
        <span class="formula">수식 셀</span>
        <span class="merged">병합 앵커</span>
        <span class="hidden">숨김 행/열</span>
      </div>
    </section>
  </div>
  <script>
    const DATA = ${json};
    const state = { workbook: 0, sheet: 0, mode: 'value', search: '', selected: '' };
    const $ = (id) => document.getElementById(id);
    const fmt = (n) => Number(n || 0).toLocaleString('ko-KR');
    function colName(index) {
      let s = '';
      let n = index + 1;
      while (n > 0) {
        const r = (n - 1) % 26;
        s = String.fromCharCode(65 + r) + s;
        n = Math.floor((n - 1) / 26);
      }
      return s;
    }
    function esc(value) {
      return String(value ?? '').replace(/[&<>"']/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
    }
    function activeWorkbook() { return DATA.workbooks[state.workbook]; }
    function activeSheet() { return activeWorkbook().sheets[state.sheet]; }
    function renderSummary() {
      const totals = DATA.workbooks.reduce((acc, wb) => {
        acc.sheets += wb.sheetCount;
        acc.cells += wb.totals.cellCount;
        acc.formulas += wb.totals.formulaCount;
        acc.merges += wb.totals.mergeCount;
        return acc;
      }, { sheets: 0, cells: 0, formulas: 0, merges: 0 });
      $('globalSummary').innerHTML = [
        ['파일', DATA.workbooks.length],
        ['시트', fmt(totals.sheets)],
        ['셀', fmt(totals.cells)],
        ['수식', fmt(totals.formulas)],
        ['병합', fmt(totals.merges)],
        ['버전', DATA.schemaVersion],
      ].map(([label, value]) => '<div class="metric"><span>' + label + '</span><strong>' + value + '</strong></div>').join('');
    }
    function renderWorkbookList() {
      $('workbookList').innerHTML = DATA.workbooks.map((wb, index) => {
        const active = index === state.workbook ? ' is-active' : '';
        return '<button class="nav-button' + active + '" type="button" data-workbook="' + index + '"><strong>' + esc(wb.fileName) + '</strong><span>' + fmt(wb.sheetCount) + '시트 · ' + fmt(wb.totals.cellCount) + '셀 · ' + fmt(wb.totals.formulaCount) + '수식</span></button>';
      }).join('');
      document.querySelectorAll('[data-workbook]').forEach((button) => {
        button.addEventListener('click', () => {
          state.workbook = Number(button.dataset.workbook);
          state.sheet = 0;
          state.selected = '';
          renderAll();
        });
      });
    }
    function renderSheetControls() {
      const wb = activeWorkbook();
      $('sheetList').innerHTML = wb.sheets.map((sheet, index) => {
        const active = index === state.sheet ? ' is-active' : '';
        return '<button class="nav-button' + active + '" type="button" data-sheet="' + index + '"><strong>' + esc(sheet.sheetName) + '</strong><span>' + esc(sheet.sheetRole) + ' · ' + fmt(sheet.rowCount) + '×' + fmt(sheet.columnCount) + ' · 수식 ' + fmt(sheet.formulaCount) + '</span></button>';
      }).join('');
      $('sheetSelect').innerHTML = wb.sheets.map((sheet, index) => '<option value="' + index + '">' + esc(sheet.sheetName) + '</option>').join('');
      $('sheetSelect').value = String(state.sheet);
      document.querySelectorAll('[data-sheet]').forEach((button) => {
        button.addEventListener('click', () => {
          state.sheet = Number(button.dataset.sheet);
          state.selected = '';
          renderAll();
        });
      });
    }
    function renderSheetMeta() {
      const sheet = activeSheet();
      $('sheetMeta').innerHTML = [
        ['파일', activeWorkbook().fileName, 'strong'],
        ['시트', sheet.sheetName, 'strong'],
        ['역할', sheet.sheetRole, ''],
        ['범위', sheet.ref || '사용영역 없음', ''],
        ['행×열', fmt(sheet.rowCount) + '×' + fmt(sheet.columnCount), ''],
        ['셀', fmt(sheet.cellCount), ''],
        ['수식', fmt(sheet.formulaCount), ''],
        ['병합', fmt(sheet.mergeCount), ''],
      ].map(([k, v, cls]) => '<span class="pill ' + cls + '">' + esc(k) + ': ' + esc(v) + '</span>').join('');
    }
    function cellText(cell) {
      if (!cell) return '';
      if (state.mode === 'formula') return cell.dataType === 'FORMULA' ? cell.rawValue : cell.displayValue;
      return cell.displayValue;
    }
    function matchesSearch(address, cell) {
      if (!state.search) return false;
      const needle = state.search.toLowerCase();
      const hay = [address, cell?.displayValue, cell?.rawValue, cell?.cachedValue, cell?.numberFormat, ...(cell?.references || [])].join(' ').toLowerCase();
      return hay.includes(needle);
    }
    function renderGrid() {
      const sheet = activeSheet();
      if (!sheet.rowCount || !sheet.columnCount) {
        $('gridWrap').innerHTML = '<div class="empty">이 시트는 사용영역이 비어 있습니다.</div>';
        renderDetail(null);
        return;
      }
      const covers = new Set(sheet.mergeCovers);
      let html = '<table class="sheet-grid"><thead><tr><th class="corner row-head"></th>';
      for (let c = sheet.startCol; c < sheet.startCol + sheet.columnCount; c += 1) {
        html += '<th>' + colName(c) + '</th>';
      }
      html += '</tr></thead><tbody>';
      for (let r = sheet.startRow; r < sheet.startRow + sheet.rowCount; r += 1) {
        html += '<tr><th class="row-head">' + (r + 1) + '</th>';
        for (let c = sheet.startCol; c < sheet.startCol + sheet.columnCount; c += 1) {
          const address = colName(c) + (r + 1);
          if (covers.has(address)) continue;
          const cell = sheet.cells[address];
          const classes = ['cell'];
          if (cell?.dataType === 'FORMULA') classes.push('is-formula');
          if (cell?.mergedRange) classes.push('is-merged');
          if (cell?.hidden) classes.push('is-hidden');
          if (cell?.dataType === 'ERROR') classes.push('is-error');
          if (matchesSearch(address, cell)) classes.push('is-match');
          if (state.selected === address) classes.push('is-selected');
          const span = cell ? ' rowspan="' + cell.rowSpan + '" colspan="' + cell.colSpan + '"' : '';
          const title = cell ? esc(address + ' ' + (cell.rawValue ?? cell.displayValue ?? '')) : address;
          html += '<td class="' + classes.join(' ') + '" data-address="' + address + '"' + span + ' title="' + title + '">' + esc(cellText(cell)) + '</td>';
        }
        html += '</tr>';
      }
      html += '</tbody></table>';
      $('gridWrap').innerHTML = html;
      document.querySelectorAll('[data-address]').forEach((td) => {
        td.addEventListener('click', () => {
          state.selected = td.dataset.address;
          renderGrid();
          renderDetail(activeSheet().cells[state.selected] || { address: state.selected });
        });
      });
      const firstMatch = document.querySelector('.is-match');
      if (firstMatch) firstMatch.scrollIntoView({ block: 'center', inline: 'center' });
      renderDetail(state.selected ? sheet.cells[state.selected] || { address: state.selected } : null);
    }
    function renderDetail(cell) {
      if (!cell) {
        $('detailHelp').style.display = '';
        $('cellDetail').innerHTML = '';
        return;
      }
      $('detailHelp').style.display = 'none';
      const rows = [
        ['address', cell.address],
        ['dataType', cell.dataType || 'BLANK'],
        ['displayValue', cell.displayValue || ''],
        ['rawValue', cell.rawValue ?? ''],
        ['cachedValue', cell.cachedValue ?? ''],
        ['numberFormat', cell.numberFormat || ''],
        ['mergedRange', cell.mergedRange || ''],
        ['hidden', cell.hidden ? 'true' : 'false'],
        ['references', (cell.references || []).join(', ')],
      ];
      $('cellDetail').innerHTML = rows.map(([k, v]) => '<div>' + esc(k) + '</div><div><code>' + esc(v) + '</code></div>').join('');
    }
    function renderAll() {
      renderSummary();
      renderWorkbookList();
      renderSheetControls();
      renderSheetMeta();
      renderGrid();
    }
    $('sheetSelect').addEventListener('change', (event) => {
      state.sheet = Number(event.target.value);
      state.selected = '';
      renderAll();
    });
    $('searchInput').addEventListener('input', (event) => {
      state.search = event.target.value.trim();
      renderGrid();
    });
    $('valueMode').addEventListener('click', () => {
      state.mode = 'value';
      $('valueMode').classList.add('is-active');
      $('formulaMode').classList.remove('is-active');
      renderGrid();
    });
    $('formulaMode').addEventListener('click', () => {
      state.mode = 'formula';
      $('formulaMode').classList.add('is-active');
      $('valueMode').classList.remove('is-active');
      renderGrid();
    });
    renderAll();
  </script>
</body>
</html>`;
}

const workbooks = inputFiles.map(convertWorkbook);
const payload = {
  schemaVersion: 'WorkbookIR review v1',
  generatedAt: new Date().toISOString(),
  workbooks,
};

fs.mkdirSync(path.dirname(outputPath), { recursive: true });
fs.writeFileSync(outputPath, htmlDocument(payload), 'utf8');

console.log(
  JSON.stringify(
    {
      outputPath,
      workbookCount: workbooks.length,
      sheetCount: workbooks.reduce((sum, wb) => sum + wb.sheetCount, 0),
      cellCount: workbooks.reduce((sum, wb) => sum + wb.totals.cellCount, 0),
      formulaCount: workbooks.reduce((sum, wb) => sum + wb.totals.formulaCount, 0),
      mergeCount: workbooks.reduce((sum, wb) => sum + wb.totals.mergeCount, 0),
    },
    null,
    2,
  ),
);
