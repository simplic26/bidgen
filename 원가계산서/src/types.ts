export type ProcurementType = 'CONSTRUCTION' | 'SERVICE' | 'GOODS';

export type ValidationStatus = 'ERROR' | 'NEEDS_REVIEW' | 'WARNING' | 'OK' | 'UNAVAILABLE';

export type ValidationType =
  | 'ARITHMETIC'
  | 'TOTAL'
  | 'RATE'
  | 'BASE'
  | 'CONDITION'
  | 'FORMULA'
  | 'REFERENCE'
  | 'LABOR'
  | 'STANDARD';

export interface FileMeta {
  name: string;
  sizeLabel: string;
  detail: string;
}

// 기준자료 입력 (PRD §8.1): 제비율·노임단가는 Excel 필수, 표준품셈은 PDF 선택
export interface ReferenceFiles {
  rateFile: FileMeta | null; // 제비율 Excel (.xlsx) 필수
  laborFile: FileMeta | null; // 노임단가 Excel (.xlsx) 필수
  standardPdf: FileMeta | null; // 표준품셈 PDF (.pdf) 선택
}

// 판단근거 문서 유형 (PRD §13.1)
export type EvidenceDocType = 'RATE_EXCEL' | 'LABOR_RATE_EXCEL' | 'STANDARD_PDF' | 'BUILTIN';

export interface ValidationResult {
  resultId: string;
  status: ValidationStatus;
  severity: 'HIGH' | 'MEDIUM' | 'LOW' | 'INFO';
  validationType: ValidationType;
  procurementType: ProcurementType;
  item: {
    canonicalName: string;
    originalName: string;
    category: string;
  };
  excel: {
    sheetName: string;
    cell: string;
    rateCell?: string;
    inputValue: string;
    inputRate?: string;
    formula: string;
    referencedCells: string[];
    previewRows: Array<Record<string, string>>;
  };
  expected: {
    baseAmount: string;
    rate?: string;
    rawAmount: string;
    roundingMethod: string;
    finalAmount: string;
  };
  difference: string;
  summary: string;
  reason: string;
  evidence: {
    documentTitle: string;
    documentType: EvidenceDocType;
    // 기준자료 Excel 근거 (제비율/노임단가)
    sheetName?: string;
    cell?: string;
    displayValue?: string;
    // 선택 표준품셈 PDF 근거
    page?: number;
    tableTitle: string;
    quote: string;
    confidence: number;
    appliedCondition: string;
  };
  recommendedAction: string;
}

// --- 정규화 원가계산서 리포트 (PRD §6.4 / FR-034) ---

export type ResolutionStatus =
  | 'SAME_ROW'
  | 'FORMULA_TRACE'
  | 'CROSS_SHEET_TRACE'
  | 'CROSS_SHEET_UNRESOLVED'
  | 'FORMULA_UNRESOLVED'
  | 'UNRESOLVED';

export interface NormalizedCostRow {
  rowId: string;
  status: ValidationStatus;
  sourceSection: string; // 원본구간
  costNature: string; // 비용성격
  validationPolicy: string; // 검증정책 (RATE_CHECK 등)
  canonicalName: string; // 표준항목
  originalName: string; // 원본항목
  baseLabel: string; // 산출기초명
  baseAmount: string; // 산출기초값
  rate: string; // 요율
  rateSource: string; // 요율출처
  amount: string; // 금액
  calculatedAmount: string; // 계산값
  difference: string; // 차이
  calculationCell: string; // 계산셀
  resolutionStatus: ResolutionStatus; // 해결상태
  formulaConstants: string; // 수식상수
  tracePath: string; // 추적경로
  note: string; // 비고
}

export interface RecognitionSheet {
  sheetName: string;
  status: '자동 인식' | '확인 필요' | '분석 완료';
  role: string;
  description: string;
}

export interface RecognitionCriterion {
  title: string;
  count: number;
  status: '추출 완료' | '확인 필요' | '검증 불가' | '미수행';
  description: string;
}

export interface RecognitionSummary {
  sheets: RecognitionSheet[];
  criteria: RecognitionCriterion[];
  reviewPrompt: string;
}

export interface ProgressStep {
  label: string;
  detail: string;
}

// --- Workbook IR: 양식과 무관한 고정 JSON 스키마 (PRD §9.2 기반) ---

export type SheetRole =
  | 'COVER_SUMMARY'
  | 'COST_SUMMARY'
  | 'CONSTRUCTION_ITEMS'
  | 'QUANTITY'
  | 'UNIT_PRICE'
  | 'PRICE_SURVEY'
  | 'WAGE_RATE'
  | 'RATE_STANDARD'
  | 'OTHER';

export type IRCellType = 'FORMULA' | 'NUMBER' | 'STRING' | 'BOOLEAN' | 'DATE' | 'ERROR' | 'BLANK';

export interface IRCell {
  address: string; // "F12"
  dataType: IRCellType;
  rawValue: string; // 수식이면 "=INT(F8*F10)", 아니면 리터럴
  cachedValue: string | number | boolean | null; // 마지막 저장 계산값
  displayValue: string; // 표시 텍스트
  numberFormat: string | null;
  mergedRange: string | null; // 병합 앵커 셀이면 "A1:B1", 아니면 null
  hidden: boolean; // 행 또는 열 숨김
  references: string[]; // 수식에서 추출 (로컬 + 시트간)
}

export interface IRSheet {
  sheetName: string;
  sheetRole: SheetRole;
  rowCount: number;
  columnCount: number;
  cellCount: number;
  formulaCount: number;
  mergeCount: number;
  cells: IRCell[];
}

export interface WorkbookIR {
  schemaVersion: string; // "1.0"
  fileName: string;
  procurementType: ProcurementType;
  generatedAt: string; // ISO
  sheets: IRSheet[];
  totals: { sheetCount: number; cellCount: number; formulaCount: number; mergeCount: number };
}

// --- 검증 엔진 설정 ---

export type ValidationMode = 'ARITHMETIC_ONLY' | 'ARITHMETIC_AND_RATE';

export type RoundingMethod = 'ROUND_WON' | 'FLOOR_WON' | 'FLOOR_TEN' | 'NONE';

// 기준자료 Excel에서 추출한 요율의 출처 (셀 근거)
export interface RateSource {
  documentTitle: string; // 파일명
  sheetName: string;
  cell: string;
  displayValue: string;
}

export interface ReferenceRate {
  canonicalName: string;
  rate: number | null; // 분수(0.032). null = 미입력
  roundingMethod?: RoundingMethod;
  source?: RateSource; // 제비율 Excel에서 파싱한 경우의 셀 근거
}

// 제비율 Excel에서 파싱한 기준요율 (PRD §8.5 RATE_EXCEL)
export interface RateCriterion {
  canonicalName: string;
  rate: number; // 분수(0.0356)
  sheetName: string;
  cell: string;
  displayValue: string;
  confidence: number;
  condition?: string; // 간접공사비 매트릭스의 적용 구간(직접공사비)·기간 조건
}

// 노임단가 Excel에서 파싱한 직종별 단가 (PRD §8.5 LABOR_RATE_EXCEL)
export interface LaborRate {
  occupationName: string;
  unitPrice: number;
  sheetName: string;
  cell: string;
  displayValue: string;
  confidence: number;
}

export interface ValidationConfig {
  mode: ValidationMode;
  referenceRates: Record<string, ReferenceRate>; // canonicalName 키
  laborRates?: LaborRate[]; // 노임단가 Excel 파싱 결과 (V-CON-007)
  laborDocumentTitle?: string; // 노임단가 파일명 (근거 표기용)
  maxCellsScanned?: number; // 기본 20000
  defaultRounding?: RoundingMethod; // 기본 ROUND_WON
}

export interface DetectedItem {
  canonicalName: string;
  originalLabel: string;
  category: string;
  sheetName: string;
  labelCell: string;
  amountCell: string | null;
  rateCell: string | null;
  amountValue: number | null;
  rateValue: number | null; // 분수
  requiresReference: boolean;
}

export interface ResultFilters {
  status: ValidationStatus | 'ALL';
  validationType: ValidationType | 'ALL';
  sheetName: string;
}

// --- 공고문(hwpx) 생성 ---

// 공고문 생성 폼 입력값. 핵심 10개(필수) + 추가 입력(선택). 자동 유도값은 폼에 저장하지
// 않고 noticeDerive.ts의 resolveNoticeValues가 매번 계산한다 (NoticeOverrides로 수정).
export interface NoticeForm {
  // 핵심 입력 (필수 10)
  office: string; // 사업소 (콤보박스)
  contractType: string; // 계약종류 (콤보박스)
  contractDetail: string; // 세부계약종류 (계약종류에 연동)
  orderMonth: string; // 발주년월 (YYYY-MM — 출력 시 YYYY/MM)
  bidMethod: string; // 입찰방법 (콤보박스 — 예: 제한경쟁(전자입찰))
  agreementTarget: string; // 협정대상여부 (비대상/대상)
  projectName: string; // 공사명 (파일명 자동 제안)
  estimatedPrice: string; // 추정가격 (부가세 별도 — 원가계산서 자동 제안)
  noticeDate: string; // 공고일 (YYYY-MM-DD, 기본 오늘)
  period: string; // 공사기간 (예: 착공일로부터 3개월)
  procurementType: ProcurementType; // 발주유형 (공사/용역/물품)
  // 추가 입력 (선택 — 생성 비차단). 일정 4개는 자동 제안 없이 직접 입력(사용자 결정).
  bidStart: string; // 가격입찰서 접수 개시 일시
  bidDeadline: string; // 가격입찰서 접수 마감 일시
  openingDate: string; // 개찰 일시
  depositDeadline: string; // 입찰보증금 납부기한
  description: string; // 공사내용
  constructionClass: string; // 공사구분 (기본: 전문공사)
  jointContract: string; // 공동계약 (기본: 불허)
  qualification: string; // 입찰참가자격 (요구 면허/업종)
  ownerMaterialCost: string; // 사급자재비 (부가세 별도, 기본: 0)
  noticeNumber: string; // 입찰공고번호
  contactName: string; // 담당자
  contactPhone: string; // 연락처(문의전화1)
  inquiryDept2: string; // 공사내용 문의 부서
  inquiryPhone2: string; // 공사내용 문의 전화
  reportPhone: string; // 비리신고 전화
  reportEmail: string; // 비리신고 이메일
  objectionContact: string; // 이의제기 담당자
  objectionPhone: string; // 이의제기 전화
  objectionEmail: string; // 이의제기 이메일
}

// 자동 유도되는 값의 키 (수정 가능 — NoticeOverrides에 기록되면 수동값이 우선)
export type DerivedKey =
  | '예비가격기초금액'
  | '추정금액'
  | 'A값'
  | '개찰장소'
  | '공고명의'
  | '문의부서1'
  | '발주기관'
  | '국민건강보험료'
  | '국민연금보험료'
  | '노인장기요양보험료'
  | '퇴직공제부금비'
  | '산업안전보건관리비'
  | '순공사원가';

export type NoticeOverrides = Partial<Record<DerivedKey, string>>;

// 필드 값의 출처: manual=직접 입력, auto=규칙 유도, linked=원가계산서 연동, edited=유도값을 수정함
export type ValueSource = 'manual' | 'auto' | 'linked' | 'edited';

export type InsuranceKey =
  | '국민건강보험료'
  | '국민연금보험료'
  | '노인장기요양보험료'
  | '퇴직공제부금비'
  | '산업안전보건관리비';

// 원가계산서 검증 플로우에서 넘어온 자동 채움 값
export interface NoticePrefill {
  procurementType: ProcurementType;
  projectName?: string; // 업로드 파일명에서 유도
  totalAmount?: number; // 추정가격 제안값 (부가세 차감 보정 후, 원)
  vatAmount?: number; // 탐지된 부가가치세(원)
  insurances?: Partial<Record<InsuranceKey, number>>; // 보험료 5종 (탐지된 것만)
  netConstructionCost?: number; // 순공사원가 (탐지/폴백 계산 성공 시)
  sourceFileName?: string; // 자동입력 안내 배지용
}
