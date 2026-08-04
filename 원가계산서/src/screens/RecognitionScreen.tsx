import { useState } from 'react';
import { ArrowLeft, ArrowRight, CheckCircle2, FileSearch, FileText, TriangleAlert } from '../components/icons';
import type {
  FileMeta,
  LaborRate,
  RateCriterion,
  RecognitionCriterion,
  RecognitionSummary,
  ReferenceFiles,
  ReferenceRate,
  ValidationMode,
  WorkbookIR,
} from '../types';
import { Button, Panel } from '../components/ui';
import { countReferenceFiles } from '../utils/results';
import { downloadJson } from '../utils/download';

interface RecognitionScreenProps {
  excelFile: FileMeta | null;
  referenceFiles: ReferenceFiles;
  summary: RecognitionSummary;
  workbookIR: WorkbookIR | null;
  mode: ValidationMode;
  onModeChange: (mode: ValidationMode) => void;
  referenceRows: ReferenceRate[];
  rateCriteria: RateCriterion[];
  laborRates: LaborRate[];
  rateInputs: Record<string, string>;
  onRateInput: (canonicalName: string, value: string) => void;
  onBack: () => void;
  onRun: () => void;
}

const MODE_OPTIONS: Array<{ value: ValidationMode; label: string; description: string }> = [
  { value: 'ARITHMETIC_ONLY', label: '산술/합계/수식 검증만', description: '기준자료 없이 파일 내 값만으로 검증' },
  { value: 'ARITHMETIC_AND_RATE', label: '산술 + 요율 검증', description: '아래 기준요율을 입력해 제비율·보험료까지 검증' },
];

export function RecognitionScreen({
  excelFile,
  referenceFiles,
  summary,
  workbookIR,
  mode,
  onModeChange,
  referenceRows,
  rateCriteria,
  laborRates,
  rateInputs,
  onRateInput,
  onBack,
  onRun,
}: RecognitionScreenProps) {
  const [showJson, setShowJson] = useState(false);

  // 파싱된 요율을 canonicalName으로 조회 (요율 입력란 셀 근거 힌트용)
  const criterionByName = new Map(rateCriteria.map((c) => [c.canonicalName, c]));

  // 실제 업로드 파일이 있으면 파싱 결과로, 없으면(딥링크 데모) 정적 요약으로 기준자료 인식 카드 구성
  const criteria: RecognitionCriterion[] = workbookIR
    ? [
        {
          title: '제비율 기준 Excel',
          count: rateCriteria.length,
          status: rateCriteria.length > 0 ? '추출 완료' : '미수행',
          description:
            rateCriteria.length > 0
              ? rateCriteria
                  .slice(0, 3)
                  .map((c) => `${c.canonicalName} ${c.displayValue} ← ${c.sheetName}!${c.cell}`)
                  .join(', ')
              : '제비율 Excel 미업로드 · 요율검증은 수기 입력으로 진행',
        },
        {
          title: '노임단가 Excel',
          count: laborRates.length,
          status: laborRates.length > 0 ? '추출 완료' : '미수행',
          description:
            laborRates.length > 0
              ? `${laborRates.slice(0, 4).map((l) => l.occupationName).join(', ')} 등 직종별 단가`
              : '노임단가 Excel 미업로드',
        },
        {
          title: '표준품셈 PDF',
          count: 0,
          status: referenceFiles.standardPdf ? '확인 필요' : '미수행',
          description: referenceFiles.standardPdf
            ? '선택 입력 · 대표항목 페이지 근거 (2일 시연형 제한)'
            : '표준품셈 PDF 미업로드 (선택 항목)',
        },
      ]
    : summary.criteria.map((criterion) => {
        if (criterion.title.includes('표준품셈') && !referenceFiles.standardPdf) {
          return { ...criterion, status: '미수행', description: '표준품셈 PDF가 입력되지 않아 미수행입니다. (선택 항목)' };
        }
        return criterion;
      });

  return (
    <main className="app-shell">
      <aside className="workflow-rail">
        <span>2 / 5 단계</span>
        <strong>문서 인식</strong>
      </aside>

      <section className="workspace">
        <header className="topbar">
          <div>
            <h1>문서 인식 결과</h1>
            <p>{excelFile?.name ?? '원가계산서'} · 기준자료 {countReferenceFiles(referenceFiles) || 3}개</p>
          </div>
        </header>

        <Panel>
          <div className="section-heading">
            <span>Excel 구조 인식</span>
            <p>시트 역할과 주요 내용을 자동 분류했습니다.</p>
          </div>
          <div className="data-table">
            <div className="table-row table-head">
              <span>시트명</span>
              <span>인식 결과</span>
              <span>주요 내용</span>
            </div>
            {summary.sheets.map((sheet) => (
              <div className="table-row" key={sheet.sheetName}>
                <strong>{sheet.sheetName}</strong>
                <span className={sheet.status === '확인 필요' ? 'text-review' : 'text-ok'}>
                  {sheet.status === '확인 필요' ? <TriangleAlert size={15} /> : <CheckCircle2 size={15} />}
                  {sheet.status}
                </span>
                <span>
                  {sheet.role} · {sheet.description}
                </span>
              </div>
            ))}
          </div>
        </Panel>

        {workbookIR ? (
          <Panel>
            <div className="section-heading">
              <span>구조 JSON (Workbook IR)</span>
              <p>
                양식과 무관한 고정 스키마 v{workbookIR.schemaVersion} · 시트 {workbookIR.totals.sheetCount}개 · 셀{' '}
                {workbookIR.totals.cellCount.toLocaleString()}개 · 수식 {workbookIR.totals.formulaCount.toLocaleString()}개 · 병합{' '}
                {workbookIR.totals.mergeCount.toLocaleString()}개
              </p>
            </div>
            <div className="json-actions">
              <Button icon={<FileSearch size={16} />} onClick={() => setShowJson((v) => !v)}>
                {showJson ? '구조 JSON 닫기' : '구조 JSON 보기'}
              </Button>
              <Button
                icon={<FileText size={16} />}
                onClick={() => downloadJson(`${workbookIR.fileName.replace(/\.[^.]+$/, '')}.ir.json`, workbookIR)}
              >
                JSON 다운로드
              </Button>
            </div>
            {showJson ? <pre className="json-view">{JSON.stringify(workbookIR, null, 2)}</pre> : null}
          </Panel>
        ) : null}

        <Panel>
          <div className="section-heading">
            <span>검증 모드</span>
            <p>기준요율 없이 산술만 검증하거나, 요율까지 함께 검증할 수 있습니다.</p>
          </div>
          <div className="procurement-grid">
            {MODE_OPTIONS.map((option) => (
              <button
                className={`procurement-option ${mode === option.value ? 'is-selected' : ''}`}
                key={option.value}
                onClick={() => onModeChange(option.value)}
                type="button"
              >
                <strong>{option.label}</strong>
                <span>{option.description}</span>
              </button>
            ))}
          </div>
        </Panel>

        {mode === 'ARITHMETIC_AND_RATE' ? (
          <Panel>
            <div className="section-heading">
              <span>기준요율 입력</span>
              <p>
                제비율 Excel에서 파싱된 요율(셀 근거 표시)은 비워두면 자동 적용됩니다. 값을 입력하면 수기 요율로 대체합니다.
              </p>
            </div>
            {workbookIR && referenceRows.length > 0 ? (
              <div className="rate-table">
                {referenceRows.map((row) => {
                  const parsed = criterionByName.get(row.canonicalName);
                  const fallback = row.rate != null ? (row.rate * 100).toString() : '예: 3.56';
                  return (
                    <label className="rate-row" key={row.canonicalName}>
                      <span>
                        {row.canonicalName}
                        {parsed ? (
                          <em className="rate-source">
                            제비율 Excel {parsed.sheetName}!{parsed.cell} · {parsed.displayValue} (파싱됨)
                          </em>
                        ) : null}
                      </span>
                      <span className="rate-input">
                        <input
                          aria-label={`${row.canonicalName} 기준요율(%)`}
                          inputMode="decimal"
                          onChange={(event) => onRateInput(row.canonicalName, event.target.value)}
                          placeholder={parsed ? (parsed.rate * 100).toString() : fallback}
                          type="text"
                          value={rateInputs[row.canonicalName] ?? ''}
                        />
                        <em>%</em>
                      </span>
                    </label>
                  );
                })}
              </div>
            ) : (
              <div className="review-note">
                <TriangleAlert size={18} />
                <span>{workbookIR ? '탐지된 제비율·보험료 항목이 없습니다. 산술 검증만 진행됩니다.' : 'Excel을 업로드하면 기준요율을 입력할 수 있습니다.'}</span>
              </div>
            )}
          </Panel>
        ) : null}

        <Panel>
          <div className="section-heading">
            <span>기준자료 인식</span>
            <p>추출된 기준 종류와 확인 필요 항목입니다.</p>
          </div>
          <div className="criteria-grid">
            {criteria.map((criterion) => (
              <div className={`criterion ${criterion.status === '미수행' ? 'is-skipped' : ''}`} key={criterion.title}>
                <strong>{criterion.title}</strong>
                <span>{criterion.status === '미수행' ? '미수행' : `${criterion.count}개 기준`}</span>
                <p>{criterion.description}</p>
              </div>
            ))}
          </div>
          <div className="review-note">
            <TriangleAlert size={18} />
            <span>{summary.reviewPrompt}</span>
          </div>
        </Panel>

        <div className="footer-actions">
          <Button icon={<ArrowLeft size={16} />} onClick={onBack}>
            파일 다시 선택
          </Button>
          <Button icon={<ArrowRight size={16} />} onClick={onRun} variant="primary">
            검증 실행
          </Button>
        </div>
      </section>
    </main>
  );
}
