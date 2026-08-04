import { ArrowLeft, ArrowRight, CheckSquare, FileSpreadsheet, FileText } from '../components/icons';
import type { ValidationResult } from '../types';
import { Button, Panel, StatusBadge } from '../components/ui';
import { formatConfidence } from '../utils/results';

interface DetailScreenProps {
  result: ValidationResult;
  onBack: () => void;
  onNavigate: (direction: 'previous' | 'next') => void;
}

export function DetailScreen({ result, onBack, onNavigate }: DetailScreenProps) {
  return (
    <main className="app-shell detail-layout">
      <aside className="workflow-rail">
        <span>5 / 5 단계</span>
        <strong>상세 근거</strong>
      </aside>

      <section className="workspace">
        <header className="topbar">
          <Button icon={<ArrowLeft size={16} />} onClick={onBack}>
            검증결과
          </Button>
          <div>
            <h1>{result.item.canonicalName} 상세</h1>
            <p>{result.excel.sheetName}!{result.excel.cell}</p>
          </div>
        </header>

        <div className="detail-grid">
          <Panel>
            <div className="panel-heading-row">
              <h2>Excel 원가계산서</h2>
              <FileSpreadsheet size={18} />
            </div>
            <dl className="detail-list">
              <div>
                <dt>시트</dt>
                <dd>{result.excel.sheetName}</dd>
              </div>
              <div>
                <dt>셀 위치</dt>
                <dd>{result.excel.cell}</dd>
              </div>
              <div>
                <dt>원본 수식</dt>
                <dd>{result.excel.formula}</dd>
              </div>
              <div>
                <dt>참조 셀</dt>
                <dd>{result.excel.referencedCells.join(', ')}</dd>
              </div>
            </dl>
            <div className="mini-sheet">
              {result.excel.previewRows.map((row, index) => (
                <div className={index === 1 ? 'is-target' : ''} key={`${row.A}-${index}`}>
                  {Object.entries(row).map(([key, value]) => (
                    <span key={key}>{value}</span>
                  ))}
                </div>
              ))}
            </div>
          </Panel>

          <Panel>
            <div className="panel-heading-row">
              <h2>기준자료 판단근거</h2>
              {result.evidence.documentType === 'STANDARD_PDF' ? <FileText size={18} /> : <FileSpreadsheet size={18} />}
            </div>
            <dl className="detail-list">
              <div>
                <dt>기준 문서</dt>
                <dd>{result.evidence.documentTitle}</dd>
              </div>
              {result.evidence.documentType === 'RATE_EXCEL' || result.evidence.documentType === 'LABOR_RATE_EXCEL' ? (
                <>
                  <div>
                    <dt>시트</dt>
                    <dd>{result.evidence.sheetName ?? '-'}</dd>
                  </div>
                  <div>
                    <dt>셀 주소</dt>
                    <dd>{result.evidence.cell ?? '-'}</dd>
                  </div>
                  <div>
                    <dt>기준 표시값</dt>
                    <dd>{result.evidence.displayValue ?? '-'}</dd>
                  </div>
                </>
              ) : result.evidence.documentType === 'STANDARD_PDF' ? (
                <>
                  <div>
                    <dt>페이지</dt>
                    <dd>{result.evidence.page != null ? `${result.evidence.page}페이지` : '-'}</dd>
                  </div>
                  <div>
                    <dt>표 또는 조항</dt>
                    <dd>{result.evidence.tableTitle}</dd>
                  </div>
                </>
              ) : (
                <div>
                  <dt>근거 유형</dt>
                  <dd>{result.evidence.tableTitle}</dd>
                </div>
              )}
              <div>
                <dt>적용 대상</dt>
                <dd>{result.evidence.appliedCondition}</dd>
              </div>
            </dl>
            <blockquote>{result.evidence.quote}</blockquote>
            <span className="confidence">추출 신뢰도 {formatConfidence(result.evidence.confidence)}</span>
          </Panel>
        </div>

        <Panel className="calculation-panel">
          <h2>시스템 계산과정</h2>
          <div className="formula-strip">
            <strong>{result.expected.baseAmount}원</strong>
            <span>×</span>
            <strong>{result.expected.rate ?? '기준식'}</strong>
            <span>=</span>
            <strong>{result.expected.finalAmount}원</strong>
          </div>
          <dl className="calculation-summary">
            <div>
              <dt>Excel 입력값</dt>
              <dd>{result.excel.inputValue}원</dd>
            </div>
            <div>
              <dt>시스템 기대값</dt>
              <dd>{result.expected.finalAmount}원</dd>
            </div>
            <div>
              <dt>차액</dt>
              <dd>{result.difference}원</dd>
            </div>
            <div>
              <dt>끝수처리</dt>
              <dd>{result.expected.roundingMethod}</dd>
            </div>
          </dl>

          <div className="judgment">
            <StatusBadge status={result.status} />
            <strong>{result.summary}</strong>
            <p>{result.reason}</p>
          </div>

          <div className="evidence-bullets">
            <strong>판단근거</strong>
            <ol>
              <li>Excel 입력값은 {result.excel.inputValue}원입니다.</li>
              <li>기준자료의 적용값은 {result.expected.rate ?? result.expected.finalAmount}입니다.</li>
              <li>시스템 계산값은 {result.expected.finalAmount}원입니다.</li>
              <li>차액은 {result.difference}원입니다.</li>
            </ol>
          </div>

          <div className="recommendation">
            <CheckSquare size={18} />
            <span>{result.recommendedAction}</span>
          </div>

          <label className="memo-box">
            사용자 판단
            <textarea placeholder="예외 적용 항목으로 확인한 경우 메모를 남깁니다." />
          </label>

          <div className="footer-actions">
            <Button onClick={() => onNavigate('previous')}>이전 항목</Button>
            <Button icon={<ArrowRight size={16} />} onClick={() => onNavigate('next')} variant="primary">
              다음 오류 항목
            </Button>
          </div>
        </Panel>
      </section>
    </main>
  );
}
