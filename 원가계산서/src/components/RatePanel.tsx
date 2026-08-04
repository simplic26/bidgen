import { useMemo, useState } from 'react';
import type { FileMeta, RateCriterion } from '../types';
import { Button, Panel, SummaryCard } from './ui';
import { Download, FileSearch, FileSpreadsheet, LoaderCircle, RotateCcw } from './icons';
import { buildCriteriaFromRecords, downloadRateExcel, refreshRates, type RefreshResult } from '../utils/rateApi';

type Gongjong = '토목' | '건축';
type Status = 'idle' | 'loading' | 'done' | 'error';

interface RatePanelProps {
  rateFileMeta: FileMeta | null;
  /** 최신화 완료 시 파생된 기준요율을 검증 파이프라인에 전달 (기존 onRateExcel 대체). */
  onRateResult: (meta: FileMeta, criteria: RateCriterion[]) => void;
}

function todayISO(): string {
  const now = new Date();
  const tz = now.getTimezoneOffset() * 60000;
  return new Date(now.getTime() - tz).toISOString().slice(0, 10);
}

function nowStamp(): string {
  return todayISO().replace(/-/g, '') + '_' + new Date().toTimeString().slice(0, 8).replace(/:/g, '');
}

export function RatePanel({ rateFileMeta, onRateResult }: RatePanelProps) {
  const [years, setYears] = useState(3);
  const [refDate, setRefDate] = useState(todayISO());
  const [status, setStatus] = useState<Status>('idle');
  const [error, setError] = useState('');
  const [result, setResult] = useState<RefreshResult | null>(null);
  const [gongjong, setGongjong] = useState<Gongjong>('토목');
  const [showLog, setShowLog] = useState(false);
  const [showPreview, setShowPreview] = useState(false);
  const [downloading, setDownloading] = useState(false);

  const previewRows = useMemo(
    () => (result ? result.preview.filter((r) => r.공종 === gongjong) : []),
    [result, gongjong],
  );

  function feedCriteria(next: RefreshResult, g: Gongjong) {
    const criteria = buildCriteriaFromRecords(next.records, g);
    const meta: FileMeta = {
      name: `조달청 제비율 자동수집 (${next.files.length}개 파일)`,
      sizeLabel: `${next.count.toLocaleString()}행`,
      detail: `${g} 기준 · 요율 ${criteria.length}개 · 게시물 ${next.posts}건 검색`,
    };
    onRateResult(meta, criteria);
  }

  async function handleRefresh() {
    setStatus('loading');
    setError('');
    setShowPreview(false);
    try {
      const next = await refreshRates(years, refDate);
      setResult(next);
      setStatus('done');
      feedCriteria(next, gongjong);
    } catch (e) {
      setStatus('error');
      setError(e instanceof Error ? e.message : String(e));
    }
  }

  function handleGongjong(next: Gongjong) {
    setGongjong(next);
    if (result) feedCriteria(result, next);
  }

  async function handleDownload() {
    if (!result) return;
    setDownloading(true);
    setError('');
    try {
      await downloadRateExcel(result.records, `제비율_정규화_${nowStamp()}.xlsx`);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setDownloading(false);
    }
  }

  const loading = status === 'loading';

  return (
    <Panel className="rate-panel">
      <div className="section-heading">
        <span>제비율 기준 (조달청 자동수집)</span>
        <p>업로드 대신, 조달청 시설공사 게시판의 「간접공사비(제비율) 적용기준」을 직접 최신화·정규화합니다.</p>
      </div>

      {/* 1) 제비율 최신화 하기 (크롤링) */}
      <div className="rate-controls">
        <label className="rate-field">
          <span>조회기준일</span>
          <input type="date" value={refDate} onChange={(e) => setRefDate(e.currentTarget.value)} />
        </label>
        <label className="rate-field">
          <span>수집기간(년)</span>
          <input
            type="number"
            min={1}
            max={10}
            value={years}
            onChange={(e) => setYears(Math.max(1, Math.min(10, Number(e.currentTarget.value) || 3)))}
          />
        </label>
        <div className="rate-seg" role="group" aria-label="공종 선택">
          {(['토목', '건축'] as Gongjong[]).map((g) => (
            <button
              key={g}
              type="button"
              className={gongjong === g ? 'is-active' : ''}
              onClick={() => handleGongjong(g)}
            >
              {g}
            </button>
          ))}
        </div>
        <Button
          variant="primary"
          onClick={handleRefresh}
          disabled={loading}
          icon={loading ? <LoaderCircle className="spin" size={16} /> : <RotateCcw size={16} />}
        >
          {loading ? '수집 중…' : '제비율 최신화 하기'}
        </Button>
      </div>

      {loading ? <p className="rate-hint">조달청 접속·수집·정규화 중입니다. 최대 1분까지 걸릴 수 있습니다…</p> : null}
      {status === 'error' ? <p className="rate-error">⚠ {error}</p> : null}

      {result ? (
        <>
          <div className="rate-kpis">
            <SummaryCard label="검색 게시물" value={result.posts} />
            <SummaryCard label="수집 파일" value={result.files.length} tone="ok" />
            <SummaryCard label="정규화 행" value={result.count.toLocaleString()} tone="review" />
            <SummaryCard label="오류" value={result.errors.length} tone={result.errors.length ? 'error' : 'neutral'} />
          </div>

          {rateFileMeta ? (
            <div className="upload-meta rate-meta">
              <FileSpreadsheet size={16} /> {rateFileMeta.name} · {rateFileMeta.detail}
            </div>
          ) : null}

          {/* 2) 제비율 조회하기 / 3) 제비율 다운로드 */}
          <div className="rate-actions">
            <Button
              variant="secondary"
              onClick={() => setShowPreview((v) => !v)}
              icon={<FileSearch size={16} />}
            >
              {showPreview ? '미리보기 닫기' : '제비율 조회하기'}
            </Button>
            <Button
              variant="primary"
              onClick={handleDownload}
              disabled={downloading || result.count === 0}
              icon={downloading ? <LoaderCircle className="spin" size={16} /> : <Download size={16} />}
            >
              {downloading ? '생성 중…' : '제비율 다운로드 (엑셀)'}
            </Button>
            <button type="button" className="rate-loglink" onClick={() => setShowLog((v) => !v)}>
              {showLog ? '수집 로그 숨기기' : '수집 로그 보기'}
            </button>
          </div>

          {showLog ? <pre className="rate-log">{result.logs.join('\n')}</pre> : null}

          {showPreview ? (
            <div className="rate-preview">
              <div className="rate-preview-head">
                {gongjong} · 최대 {previewRows.length.toLocaleString()}행 미리보기 (전체 {result.count.toLocaleString()}행)
              </div>
              <div className="rate-table-scroll">
                <table className="rate-table">
                  <thead>
                    <tr>{result.columns.map((c) => <th key={c}>{c}</th>)}</tr>
                  </thead>
                  <tbody>
                    {previewRows.map((row, i) => (
                      <tr key={i}>
                        {result.columns.map((c) => {
                          const v = (row as Record<string, unknown>)[c];
                          return <td key={c}>{v == null ? '' : String(v)}</td>;
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          ) : null}
        </>
      ) : null}
    </Panel>
  );
}
