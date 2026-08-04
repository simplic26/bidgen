import type { RateCriterion } from '../types';
import { matchCanonicalLabel } from './validation';

/** 조달청 제비율 정규화 레코드 (parser.py 출력 스키마와 동일 · 한글 키). */
export interface NormalizedRecord {
  공종: string;
  적용기준일: string;
  비목: string;
  산정기준: string;
  세부구분: string;
  공사규모: string;
  공사기간: string;
  '요율(%)': number | null;
  '기초액(천원)': number | null;
  비고: string;
}

export interface RateFileMeta {
  file: string;
  공종: string;
  적용기준일: string;
  recognized: boolean;
}

/** POST /api/refresh 응답 (crawler + parser 결과). */
export interface RefreshResult {
  ok: boolean;
  error?: string;
  trace?: string;
  downloaded: string[];
  posts: number;
  files: RateFileMeta[];
  count: number;
  columns: string[];
  records: NormalizedRecord[];
  preview: NormalizedRecord[];
  errors: Array<[string, string]>;
  logs: string[];
}

/** 제비율 최신화(크롤링) + 정규화 실행. */
export async function refreshRates(years: number, refDate: string): Promise<RefreshResult> {
  const res = await fetch('/api/refresh', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ years, ref_date: refDate }),
  });
  let data: RefreshResult;
  try {
    data = (await res.json()) as RefreshResult;
  } catch {
    throw new Error(`서버 응답을 해석할 수 없습니다 (HTTP ${res.status}).`);
  }
  if (!res.ok || !data.ok) throw new Error(data.error || `수집 실패 (HTTP ${res.status})`);
  return data;
}

/** 정규화 레코드 → 표준 엑셀 다운로드 (POST /api/export → xlsx Blob). */
export async function downloadRateExcel(records: NormalizedRecord[], fileName: string): Promise<void> {
  const res = await fetch('/api/export', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ records }),
  });
  if (!res.ok) {
    let msg = `다운로드 실패 (HTTP ${res.status})`;
    try {
      const j = await res.json();
      if (j?.error) msg = j.error;
    } catch {
      /* 바이너리 응답이 아닌 오류 */
    }
    throw new Error(msg);
  }
  const blob = await res.blob();
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

/**
 * 정규화 레코드 → 검증용 기준요율(RateCriterion[]) 파생.
 * 선택 공종의 가장 최신 적용기준일을 기준으로, 표준 원가항목별 대표(최초) 요율을 채택한다.
 * (원가계산서 자동검증의 요율검증 입력으로 사용됨 — 기존 '제비율 Excel 업로드'를 대체)
 */
export function buildCriteriaFromRecords(records: NormalizedRecord[], gongjong: '토목' | '건축'): RateCriterion[] {
  const scoped = records.filter((r) => r.공종 === gongjong);
  if (scoped.length === 0) return [];
  const dates = scoped.map((r) => r.적용기준일).sort();
  const latest = dates[dates.length - 1];
  const rows = scoped.filter((r) => r.적용기준일 === latest);

  const out: RateCriterion[] = [];
  const seen = new Set<string>();
  for (const r of rows) {
    const canonical = matchCanonicalLabel(r.비목);
    if (!canonical || seen.has(canonical)) continue;
    const pct = typeof r['요율(%)'] === 'number' ? r['요율(%)'] : null;
    if (pct == null) continue;
    const rate = pct > 1 ? pct / 100 : pct; // 15 → 0.15
    if (!(rate > 0 && rate < 1)) continue;
    seen.add(canonical);
    out.push({
      canonicalName: canonical,
      rate,
      sheetName: `${gongjong} 제비율 (${latest})`,
      cell: '-',
      displayValue: `${pct}%`,
      confidence: 0.9,
      condition: [r.세부구분, r.공사규모, r.공사기간].filter(Boolean).join(' / ') || undefined,
    });
  }
  return out;
}
