import { ArrowRight, FileSearch, RotateCcw } from '../components/icons';
import type { NormalizedCostRow, ReferenceFiles, ResultFilters, ValidationResult, ValidationStatus, ValidationType } from '../types';
import { Button, Panel, StatusBadge, SummaryCard } from '../components/ui';
import {
  countReferenceFiles,
  evidenceLocation,
  filterResults,
  formatConfidence,
  getStatusCounts,
  sortResultsByPriority,
  statusLabels,
  uniqueSheets,
  validationTypeLabels,
} from '../utils/results';

interface DashboardScreenProps {
  excelFile: { name: string } | null;
  referenceFiles: ReferenceFiles;
  results: ValidationResult[];
  normalizedRows: NormalizedCostRow[];
  filters: ResultFilters;
  selectedResult: ValidationResult;
  onFilterChange: (filters: ResultFilters) => void;
  onSelectResult: (id: string) => void;
  onOpenDetail: () => void;
  onRestart: () => void;
}

const statusFilters: Array<ValidationStatus | 'ALL'> = ['ALL', 'ERROR', 'NEEDS_REVIEW', 'WARNING', 'OK', 'UNAVAILABLE'];
const typeFilters: Array<ValidationType | 'ALL'> = ['ALL', 'ARITHMETIC', 'TOTAL', 'RATE', 'BASE', 'CONDITION', 'FORMULA'];

export function DashboardScreen({
  excelFile,
  referenceFiles,
  results,
  normalizedRows,
  filters,
  selectedResult,
  onFilterChange,
  onSelectResult,
  onOpenDetail,
  onRestart,
}: DashboardScreenProps) {
  const counts = getStatusCounts(results);
  const filtered = filterResults(results, filters);
  const sheets = uniqueSheets(results);

  // 핵심 확인사항 3~5개: 오류 → 확인 필요 순 상위 (PRD §6.3)
  const keyFindings = sortResultsByPriority(results)
    .filter((r) => r.status === 'ERROR' || r.status === 'NEEDS_REVIEW')
    .slice(0, 5);

  // 정규화 원가계산서 리포트 (PRD §6.4 / FR-034) — normalizeCostStatement 산출물(상태 필터 적용)
  const visibleRows =
    filters.status === 'ALL' ? normalizedRows : normalizedRows.filter((row) => row.status === filters.status);
  const overall = counts.ERROR > 0 ? '오류' : counts.NEEDS_REVIEW > 0 ? '확인 필요' : '정상';

  return (
    <main className="app-shell dashboard-layout">
      <aside className="workflow-rail">
        <span>4 / 5 단계</span>
        <strong>검증결과</strong>
      </aside>

      <section className="workspace dashboard-workspace">
        <header className="topbar dashboard-topbar">
          <div>
            <h1>검증결과</h1>
            <p>
              {excelFile?.name ?? '추정가격내역서_최종.xlsx'} · 기준자료 {countReferenceFiles(referenceFiles) || 3}개 · 공사
            </p>
          </div>
          <Button icon={<RotateCcw size={16} />} onClick={onRestart}>
            파일 다시 검증
          </Button>
        </header>

        <div className="summary-grid">
          <SummaryCard label="종합 결과" tone={overall === '정상' ? 'ok' : 'review'} value={overall} />
          <SummaryCard label="오류" tone="error" value={`${counts.ERROR}건`} />
          <SummaryCard label="확인 필요" tone="review" value={`${counts.NEEDS_REVIEW}건`} />
          <SummaryCard label="주의" tone="warning" value={`${counts.WARNING}건`} />
          <SummaryCard label="정상" tone="ok" value={`${counts.OK}건`} />
          <SummaryCard label="검증 불가" tone="neutral" value={`${counts.UNAVAILABLE}건`} />
        </div>

        {keyFindings.length > 0 ? (
          <Panel className="key-findings">
            <h2>핵심 확인사항</h2>
            <ol>
              {keyFindings.map((finding) => (
                <li key={finding.resultId}>
                  <StatusBadge status={finding.status} />
                  <button type="button" onClick={() => onSelectResult(finding.resultId)}>
                    <strong>{finding.item.canonicalName}</strong>
                    <span>{finding.summary}</span>
                  </button>
                  <small>
                    {finding.excel.sheetName}!{finding.excel.cell}
                  </small>
                </li>
              ))}
            </ol>
          </Panel>
        ) : null}

        <div className="dashboard-grid">
          <Panel className="filter-panel">
            <h2>검증 항목</h2>
            <div className="filter-group">
              {statusFilters.map((status) => {
                const label = status === 'ALL' ? `전체 ${results.length}` : `${statusLabels[status]} ${counts[status]}`;
                return (
                  <button
                    className={filters.status === status ? 'is-selected' : ''}
                    key={status}
                    onClick={() => onFilterChange({ ...filters, status })}
                    type="button"
                  >
                    {label}
                  </button>
                );
              })}
            </div>

            <h2>검증 종류</h2>
            <div className="check-list">
              {typeFilters.map((type) => (
                <button
                  className={filters.validationType === type ? 'is-selected' : ''}
                  key={type}
                  onClick={() => onFilterChange({ ...filters, validationType: type })}
                  type="button"
                >
                  {type === 'ALL' ? '전체' : validationTypeLabels[type]}
                </button>
              ))}
            </div>

            <label className="select-label">
              시트
              <select
                onChange={(event) => onFilterChange({ ...filters, sheetName: event.target.value })}
                value={filters.sheetName}
              >
                <option value="ALL">전체</option>
                {sheets.map((sheet) => (
                  <option key={sheet} value={sheet}>
                    {sheet}
                  </option>
                ))}
              </select>
            </label>
          </Panel>

          <Panel className="result-list-panel">
            <div className="panel-heading-row">
              <h2>검증결과</h2>
              <span>{filtered.length}개 표시</span>
            </div>
            <div className="result-list">
              {filtered.map((result) => (
                <button
                  className={`result-row ${selectedResult.resultId === result.resultId ? 'is-selected' : ''}`}
                  key={result.resultId}
                  onClick={() => onSelectResult(result.resultId)}
                  type="button"
                >
                  <StatusBadge status={result.status} />
                  <div>
                    <strong>{result.item.canonicalName}</strong>
                    <span>{result.summary}</span>
                  </div>
                  <small>
                    {result.excel.sheetName}!{result.excel.cell}
                  </small>
                </button>
              ))}
            </div>
          </Panel>

          <Panel className="evidence-panel">
            <div className="panel-heading-row">
              <h2>선택 항목 판단근거</h2>
              <FileSearch size={18} />
            </div>
            <StatusBadge status={selectedResult.status} />
            <h3>{selectedResult.item.canonicalName}</h3>
            <p>{selectedResult.reason}</p>

            <dl className="evidence-metrics">
              <div>
                <dt>Excel 입력값</dt>
                <dd>{selectedResult.excel.inputValue}원</dd>
              </div>
              <div>
                <dt>시스템 계산값</dt>
                <dd>{selectedResult.expected.finalAmount}원</dd>
              </div>
              <div>
                <dt>차이</dt>
                <dd>{selectedResult.difference}원</dd>
              </div>
              <div>
                <dt>적용 기준</dt>
                <dd>
                  {selectedResult.expected.baseAmount} × {selectedResult.expected.rate ?? '기준식'}
                </dd>
              </div>
            </dl>

            <div className="quote-box">
              <strong>기준자료 판단근거</strong>
              <p>{selectedResult.evidence.quote}</p>
              <span>
                {evidenceLocation(selectedResult.evidence)} · 신뢰도 {formatConfidence(selectedResult.evidence.confidence)}
              </span>
            </div>

            <Button icon={<ArrowRight size={16} />} onClick={onOpenDetail} variant="primary">
              상세 근거 보기
            </Button>
          </Panel>
        </div>

        <Panel className="normalized-report">
          <div className="panel-heading-row">
            <h2>정규화 원가계산서 리포트</h2>
            <span>{visibleRows.length}개 행 · 원본 양식과 무관한 표준 컬럼</span>
          </div>
          <div className="normalized-table-scroll">
            <table className="normalized-table">
              <thead>
                <tr>
                  <th>상태</th>
                  <th>원본구간</th>
                  <th>비용성격</th>
                  <th>검증정책</th>
                  <th>표준항목</th>
                  <th>원본항목</th>
                  <th className="num">산출기초값</th>
                  <th className="num">요율</th>
                  <th>요율출처</th>
                  <th className="num">금액</th>
                  <th className="num">계산값</th>
                  <th className="num">차이</th>
                  <th>계산셀</th>
                  <th>해결상태</th>
                </tr>
              </thead>
              <tbody>
                {visibleRows.map((row) => (
                  <tr
                    key={row.rowId}
                    className={`${selectedResult.resultId === row.rowId ? 'is-selected' : ''} status-row-${row.status.toLowerCase()}`}
                    onClick={() => onSelectResult(row.rowId)}
                  >
                    <td>
                      <StatusBadge status={row.status} />
                    </td>
                    <td>{row.sourceSection}</td>
                    <td>{row.costNature}</td>
                    <td>
                      <code>{row.validationPolicy}</code>
                    </td>
                    <td>{row.canonicalName}</td>
                    <td>{row.originalName}</td>
                    <td className="num">{row.baseAmount}</td>
                    <td className="num">{row.rate}</td>
                    <td>{row.rateSource}</td>
                    <td className="num">{row.amount}</td>
                    <td className="num">{row.calculatedAmount}</td>
                    <td className={`num ${row.difference !== '0' && row.difference !== '' ? 'diff' : ''}`}>{row.difference}</td>
                    <td>{row.calculationCell}</td>
                    <td>
                      <code>{row.resolutionStatus}</code>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Panel>
      </section>
    </main>
  );
}
