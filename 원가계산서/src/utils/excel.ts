import * as XLSX from 'xlsx';
import type { IRCell, IRCellType, IRSheet, ProcurementType, RecognitionSheet, SheetRole, WorkbookIR } from '../types';

const SCHEMA_VERSION = '1.0';

/** 수식에서 참조 셀(로컬 + 시트간)을 추출합니다. 예: "'가격조사서'!C5", "공사내역서!D3", "F8" */
const REFERENCE_PATTERN = /(?:(?:'[^']+'|[A-Za-z0-9_가-힣]+)!)?\$?[A-Z]{1,3}\$?[0-9]+/g;

function extractReferences(formula: string): string[] {
  const matches = formula.match(REFERENCE_PATTERN);
  if (!matches) return [];
  return Array.from(new Set(matches));
}

/** SheetJS 셀 타입(t) + 수식 여부를 IR 데이터 타입으로 매핑합니다. */
function toDataType(cell: XLSX.CellObject): IRCellType {
  if (cell.f) return 'FORMULA';
  switch (cell.t) {
    case 'n':
      return 'NUMBER';
    case 's':
      return 'STRING';
    case 'b':
      return 'BOOLEAN';
    case 'd':
      return 'DATE';
    case 'e':
      return 'ERROR';
    default:
      return 'BLANK';
  }
}

/** 시트명 키워드로 고정 역할(enum)을 추정합니다. */
export function classifySheetRole(sheetName: string): SheetRole {
  if (/표지|목차|요약/.test(sheetName)) return 'COVER_SUMMARY';
  if (/원가|계산|추정가격/.test(sheetName)) return 'COST_SUMMARY';
  if (/일위대가/.test(sheetName)) return 'UNIT_PRICE';
  if (/내역/.test(sheetName)) return 'CONSTRUCTION_ITEMS';
  if (/수량|물량/.test(sheetName)) return 'QUANTITY';
  if (/노임|임금/.test(sheetName)) return 'WAGE_RATE';
  if (/제비율|요율기준/.test(sheetName)) return 'RATE_STANDARD';
  if (/가격조사|단가/.test(sheetName)) return 'PRICE_SURVEY';
  return 'OTHER';
}

const ROLE_DISPLAY: Record<SheetRole, { role: string; description: string; status: RecognitionSheet['status'] }> = {
  COVER_SUMMARY: { role: '표지·요약', description: '문서 표지·목차·요약', status: '분석 완료' },
  COST_SUMMARY: { role: '원가계산서', description: '재료비·노무비·경비·제비율·총액', status: '자동 인식' },
  CONSTRUCTION_ITEMS: { role: '세부내역', description: '품명·규격·수량·단가·금액', status: '자동 인식' },
  QUANTITY: { role: '수량', description: '품명·규격·수량', status: '자동 인식' },
  UNIT_PRICE: { role: '산출근거', description: '단위당 재료·노무·경비', status: '자동 인식' },
  PRICE_SURVEY: { role: '단가 기준', description: '조사단가와 적용단가 열 구분 필요', status: '확인 필요' },
  WAGE_RATE: { role: '노임단가', description: '직종명·단가·기준일', status: '자동 인식' },
  RATE_STANDARD: { role: '제비율 기준', description: '요율·산출기초·금액구간', status: '자동 인식' },
  OTHER: { role: '보조자료', description: '참조 범위와 중간 계산값', status: '분석 완료' },
};

/** IR 시트를 인식 화면 표시용(RecognitionSheet)으로 변환합니다. */
export function irSheetToRecognition(sheet: IRSheet): RecognitionSheet {
  const display = ROLE_DISPLAY[sheet.sheetRole];
  return { sheetName: sheet.sheetName, status: display.status, role: display.role, description: display.description };
}

/** 이미 파싱된 XLSX 워크북을 고정 JSON 스키마(Workbook IR)로 정규화합니다. (순수 함수 · 테스트 가능) */
export function workbookToIR(
  workbook: XLSX.WorkBook,
  meta: { fileName: string; procurementType: ProcurementType; generatedAt?: string },
): WorkbookIR {
  const sheets: IRSheet[] = workbook.SheetNames.map((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    const cells: IRCell[] = [];
    let formulaCount = 0;

    const merges = worksheet?.['!merges'] ?? [];
    const mergeAnchors = new Map<string, string>();
    for (const range of merges) {
      mergeAnchors.set(XLSX.utils.encode_cell(range.s), XLSX.utils.encode_range(range));
    }

    const rows = worksheet?.['!rows'] ?? [];
    const cols = worksheet?.['!cols'] ?? [];
    const ref = worksheet?.['!ref'];
    const bounds = ref ? XLSX.utils.decode_range(ref) : null;

    if (worksheet) {
      for (const key of Object.keys(worksheet)) {
        if (key.charCodeAt(0) === 33) continue; // '!' 로 시작하는 메타 키 제외
        const cell = worksheet[key] as XLSX.CellObject;
        if (!cell || typeof cell !== 'object') continue;

        const dataType = toDataType(cell);
        if (dataType === 'FORMULA') formulaCount += 1;

        const { c, r } = XLSX.utils.decode_cell(key);
        const rowHidden = Boolean(rows[r]?.hidden);
        const colHidden = Boolean(cols[c]?.hidden);

        cells.push({
          address: key,
          dataType,
          rawValue: cell.f ? `=${cell.f}` : cell.v == null ? '' : String(cell.v),
          cachedValue: (cell.v ?? null) as IRCell['cachedValue'],
          displayValue: cell.w ?? (cell.v == null ? '' : String(cell.v)),
          numberFormat: (cell.z as string | undefined) ?? null,
          mergedRange: mergeAnchors.get(key) ?? null,
          hidden: rowHidden || colHidden,
          references: dataType === 'FORMULA' && cell.f ? extractReferences(cell.f) : [],
        });
      }
    }

    cells.sort((a, b) => {
      const pa = XLSX.utils.decode_cell(a.address);
      const pb = XLSX.utils.decode_cell(b.address);
      return pa.r - pb.r || pa.c - pb.c;
    });

    return {
      sheetName,
      sheetRole: classifySheetRole(sheetName),
      rowCount: bounds ? bounds.e.r - bounds.s.r + 1 : 0,
      columnCount: bounds ? bounds.e.c - bounds.s.c + 1 : 0,
      cellCount: cells.length,
      formulaCount,
      mergeCount: merges.length,
      cells,
    };
  });

  return {
    schemaVersion: SCHEMA_VERSION,
    fileName: meta.fileName,
    procurementType: meta.procurementType,
    generatedAt: meta.generatedAt ?? new Date().toISOString(),
    sheets,
    totals: {
      sheetCount: sheets.length,
      cellCount: sheets.reduce((sum, s) => sum + s.cellCount, 0),
      formulaCount: sheets.reduce((sum, s) => sum + s.formulaCount, 0),
      mergeCount: sheets.reduce((sum, s) => sum + s.mergeCount, 0),
    },
  };
}

/** 업로드한 Excel 파일을 브라우저에서 직접 읽어 고정 JSON 스키마(Workbook IR)로 정규화합니다. */
export async function buildWorkbookIR(file: File, procurementType: ProcurementType): Promise<WorkbookIR> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, {
    type: 'array',
    cellNF: true,
    cellStyles: true,
    cellFormula: true,
    cellDates: true,
  });
  return workbookToIR(workbook, { fileName: file.name, procurementType });
}
