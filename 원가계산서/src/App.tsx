import { useMemo, useState } from 'react';
import { recognitionSummary, progressSteps, sampleResults } from './data/sampleResults';
import { DashboardScreen } from './screens/DashboardScreen';
import { DetailScreen } from './screens/DetailScreen';
import { ProgressScreen } from './screens/ProgressScreen';
import { RecognitionScreen } from './screens/RecognitionScreen';
import { NoticeScreen } from './screens/NoticeScreen';
import { UploadScreen } from './screens/UploadScreen';
import type {
  FileMeta,
  LaborRate,
  ProcurementType,
  RateCriterion,
  RecognitionSummary,
  ReferenceFiles,
  ReferenceRate,
  ResultFilters,
  ValidationMode,
  ValidationResult,
  WorkbookIR,
} from './types';
import { countReferenceFiles, filterResults, getAdjacentResult, getReviewResults, sortResultsByPriority } from './utils/results';
import { irSheetToRecognition } from './utils/excel';
import { buildReferenceRows, normalizeCostStatement, parseRate, runValidation } from './utils/validation';
import { annotateRows, buildNormalizedRows } from './utils/normalize';
import { buildNoticePrefill } from './utils/notice';
import { parseCostTree } from './utils/costTree';

type Step = 'upload' | 'recognition' | 'progress' | 'dashboard' | 'detail' | 'notice';

interface AppProps {
  initialStep?: Step;
}

const fallbackExcel: FileMeta = {
  name: '추정가격내역서_최종.xlsx',
  sizeLabel: '2.4MB',
  detail: '18개 시트 · 912개 수식 · 시연 메타데이터',
};

const emptyReferences: ReferenceFiles = { rateFile: null, laborFile: null, standardPdf: null };

const fallbackReferences: ReferenceFiles = {
  rateFile: { name: '2026년 공사원가계산 제비율 적용기준.xlsx', sizeLabel: '0.3MB', detail: '제비율 28개 · 시트·셀 근거' },
  laborFile: { name: '2026년 상반기 노임단가.xlsx', sizeLabel: '0.2MB', detail: '직종 142개 · 기준일 2026-01-01' },
  standardPdf: { name: '조달청 시설공사 표준품셈.pdf', sizeLabel: '8.4MB', detail: '317개 항목 (선택)' },
};

export default function App({ initialStep = 'upload' }: AppProps) {
  const [step, setStep] = useState<Step>(initialStep);
  const [procurementType, setProcurementType] = useState<ProcurementType>('CONSTRUCTION');
  const [excelFile, setExcelFile] = useState<FileMeta | null>(initialStep === 'upload' ? null : fallbackExcel);
  const [workbookIR, setWorkbookIR] = useState<WorkbookIR | null>(null);
  const [referenceFiles, setReferenceFiles] = useState<ReferenceFiles>(
    initialStep === 'upload' ? emptyReferences : fallbackReferences,
  );
  const [rateCriteria, setRateCriteria] = useState<RateCriterion[]>([]);
  const [laborRates, setLaborRates] = useState<LaborRate[]>([]);
  const [mode, setMode] = useState<ValidationMode>('ARITHMETIC_ONLY');
  const [rateInputs, setRateInputs] = useState<Record<string, string>>({});
  const [results, setResults] = useState<ValidationResult[]>([]);
  const [selectedResultId, setSelectedResultId] = useState('');
  const [filters, setFilters] = useState<ResultFilters>({
    status: 'ALL',
    validationType: 'ALL',
    sheetName: 'ALL',
  });

  // 실제 결과가 있으면 그것을, 없으면(딥링크 데모) 샘플을 사용
  const orderedResults = useMemo(
    () => (results.length ? sortResultsByPriority(results) : sortResultsByPriority(sampleResults)),
    [results],
  );
  const visibleResults = useMemo(() => filterResults(orderedResults, filters), [filters, orderedResults]);

  // 정규화 원가계산서: 실제 업로드는 normalizeCostStatement(정규화 선행) + 검증결과 주석,
  // 딥링크 데모(업로드 없음)는 샘플 결과에서 파생
  const normalizedRows = useMemo(() => {
    if (workbookIR && results.length) return annotateRows(normalizeCostStatement(workbookIR), results);
    return buildNormalizedRows(orderedResults);
  }, [workbookIR, results, orderedResults]);
  const selectedResult =
    orderedResults.find((result) => result.resultId === selectedResultId) ?? visibleResults[0] ?? orderedResults[0];

  const summary = useMemo<RecognitionSummary>(() => {
    if (!workbookIR || workbookIR.sheets.length === 0) return recognitionSummary;
    return { ...recognitionSummary, sheets: workbookIR.sheets.map(irSheetToRecognition) };
  }, [workbookIR]);

  const referenceRows = useMemo<ReferenceRate[]>(() => (workbookIR ? buildReferenceRows(workbookIR) : []), [workbookIR]);

  // 파싱 확인용 원가 트리 — 인식 화면에서 검증 실행 전 육안 확인
  const costTree = useMemo(() => (workbookIR ? parseCostTree(workbookIR) : null), [workbookIR]);

  // 공고문 생성 화면 자동 채움: 업로드된 원가계산서에서 총액·파일명·발주유형 유도
  const noticePrefill = useMemo(
    () => buildNoticePrefill(procurementType, excelFile, workbookIR),
    [procurementType, excelFile, workbookIR],
  );

  const steps = useMemo(() => {
    if (!workbookIR) return progressSteps;
    const { sheetCount, formulaCount, mergeCount } = workbookIR.totals;
    const structureDetail = `${sheetCount}개 시트, ${formulaCount.toLocaleString()}개 수식, ${mergeCount.toLocaleString()}개 병합범위`;
    return progressSteps.map((s) => (s.label === 'Excel 시트 및 셀 구조 분석' ? { ...s, detail: structureDetail } : s));
  }, [workbookIR]);

  function handleExcelChange(meta: FileMeta, ir: WorkbookIR | null) {
    setExcelFile(meta);
    setWorkbookIR(ir);
    setRateInputs({});
  }

  function handleRateExcel(meta: FileMeta, criteria: RateCriterion[]) {
    setReferenceFiles((prev) => ({ ...prev, rateFile: meta }));
    setRateCriteria(criteria);
    if (criteria.length > 0) setMode('ARITHMETIC_AND_RATE'); // 요율 근거가 있으면 요율검증 자동 활성
  }

  function handleLaborExcel(meta: FileMeta, rates: LaborRate[]) {
    setReferenceFiles((prev) => ({ ...prev, laborFile: meta }));
    setLaborRates(rates);
  }

  function handleStandardPdf(meta: FileMeta) {
    setReferenceFiles((prev) => ({ ...prev, standardPdf: meta }));
  }

  function handleRun() {
    if (workbookIR) {
      const referenceRates: Record<string, ReferenceRate> = {};
      // 1) 탐지 항목 기본 요율
      for (const row of referenceRows) {
        referenceRates[row.canonicalName] = { canonicalName: row.canonicalName, rate: row.rate };
      }
      // 2) 제비율 Excel 파싱값 (셀 근거 포함) 우선 적용
      for (const c of rateCriteria) {
        referenceRates[c.canonicalName] = {
          canonicalName: c.canonicalName,
          rate: c.rate,
          source: {
            documentTitle: referenceFiles.rateFile?.name ?? '제비율.xlsx',
            sheetName: c.sheetName,
            cell: c.cell,
            displayValue: c.displayValue,
          },
        };
      }
      // 3) 사용자 수기 입력이 있으면 최종 override (셀 근거 없이 사용자 입력 근거)
      for (const [name, raw] of Object.entries(rateInputs)) {
        if (raw != null && raw.trim() !== '') {
          referenceRates[name] = { canonicalName: name, rate: parseRate(raw, null) };
        }
      }
      const real = sortResultsByPriority(
        runValidation(workbookIR, {
          mode,
          referenceRates,
          laborRates,
          laborDocumentTitle: referenceFiles.laborFile?.name,
        }),
      );
      setResults(real);
      setSelectedResultId(getReviewResults(real)[0]?.resultId ?? real[0]?.resultId ?? '');
    }
    setStep('progress');
  }

  function applyFilters(nextFilters: ResultFilters) {
    setFilters(nextFilters);
    const nextVisible = filterResults(orderedResults, nextFilters);
    if (nextVisible.length > 0 && !nextVisible.some((result) => result.resultId === selectedResultId)) {
      setSelectedResultId(nextVisible[0].resultId);
    }
  }

  if (step === 'upload') {
    return (
      <UploadScreen
        excelFile={excelFile}
        onExcelChange={handleExcelChange}
        onRateExcel={handleRateExcel}
        onLaborExcel={handleLaborExcel}
        onStandardPdf={handleStandardPdf}
        onProcurementTypeChange={setProcurementType}
        onSkipToNotice={() => setStep('notice')}
        onStart={() => setStep('recognition')}
        procurementType={procurementType}
        referenceFiles={referenceFiles}
      />
    );
  }

  if (step === 'notice') {
    return <NoticeScreen prefill={noticePrefill} onBack={() => setStep('upload')} />;
  }

  if (step === 'recognition') {
    return (
      <RecognitionScreen
        excelFile={excelFile ?? fallbackExcel}
        mode={mode}
        onBack={() => setStep('upload')}
        onModeChange={setMode}
        onRateInput={(name, value) => setRateInputs((prev) => ({ ...prev, [name]: value }))}
        onRun={handleRun}
        referenceFiles={referenceFiles}
        rateInputs={rateInputs}
        referenceRows={referenceRows}
        rateCriteria={rateCriteria}
        laborRates={laborRates}
        summary={summary}
        workbookIR={workbookIR}
        costTree={costTree}
      />
    );
  }

  if (step === 'progress') {
    return <ProgressScreen excelFile={excelFile ?? fallbackExcel} onDone={() => setStep('dashboard')} steps={steps} />;
  }

  if (step === 'detail') {
    return (
      <DetailScreen
        onBack={() => setStep('dashboard')}
        onNavigate={(direction) => {
          const adjacent = getAdjacentResult(orderedResults, selectedResult.resultId, direction);
          if (adjacent) setSelectedResultId(adjacent.resultId);
        }}
        result={selectedResult}
      />
    );
  }

  return (
    <DashboardScreen
      excelFile={excelFile ?? fallbackExcel}
      filters={filters}
      onFilterChange={applyFilters}
      onOpenDetail={() => setStep('detail')}
      onRestart={() => setStep('upload')}
      onSelectResult={setSelectedResultId}
      normalizedRows={normalizedRows}
      referenceFiles={countReferenceFiles(referenceFiles) ? referenceFiles : fallbackReferences}
      results={orderedResults}
      selectedResult={selectedResult}
    />
  );
}
