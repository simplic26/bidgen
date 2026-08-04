import { Fragment } from 'react';
import type { CostNode, CostTree } from '../utils/costTree';
import { TREE_ITEM_MIN, costTreeToDetectedItems, flattenTree, matchesWithRounding } from '../utils/costTree';
import { Panel } from './ui';

const STATUS_BADGE: Record<CostNode['status'], { text: string; cls: string }> = {
  FORMULA_VERIFIED: { text: '✓ 수식확인', cls: 'tree-badge-ok' },
  VALUE_INFERRED: { text: '✓ 값검산', cls: 'tree-badge-inferred' },
  MISMATCH: { text: '⚠ 확인 필요', cls: 'tree-badge-warn' },
  UNRESOLVED: { text: '? 미해결', cls: 'tree-badge-unresolved' },
};

// 절사 종류 → 화면 표기. INT/TRUNC/ROUNDDOWN은 모두 "절사"로 통일해 보여준다(사용자 관점에서
// 셀 값이 산출값보다 작게 잘린다는 사실이 같기 때문 — 함수명 차이는 근거 셀에서 확인 가능).
const ROUNDING_LABEL: Record<'INT' | 'TRUNC' | 'ROUNDDOWN' | 'ROUND' | 'ROUNDUP', string> = {
  INT: '절사', TRUNC: '절사', ROUNDDOWN: '절사', ROUND: '반올림', ROUNDUP: '올림',
};

function fmtAmount(n: number | null): string {
  return n == null ? '—' : Math.round(n).toLocaleString('en-US');
}

function fmtPct(rate: number): string {
  return `${(rate * 100).toFixed(3)}%`;
}

/** 요율 산출식 breakdown 라인: "└ 기초라벨(셀) 기초금액 × 요율 = 원시값 → 절사 = 최종금액 ✓/⚠" */
function calcBreakdownLine(n: CostNode): string | null {
  if (n.rate == null || n.baseAmount == null || n.amount == null) return null;
  const raw = n.baseAmount * n.rate;
  const baseLabel = n.baseLabel ?? '산출기초';
  const firstCell = n.baseCells[0] ?? '';
  const roundingName = n.rounding ? ROUNDING_LABEL[n.rounding] : null;
  const ok = Boolean(matchesWithRounding(raw, n.amount));
  const tail = ok ? '✓' : `⚠ 차이 ${fmtAmount(raw - n.amount)}`;
  const eq = roundingName
    ? `= ${fmtAmount(raw)} → ${roundingName} = ${fmtAmount(n.amount)}`
    : `= ${fmtAmount(n.amount)}`;
  return `└ ${baseLabel}(${firstCell}) ${fmtAmount(n.baseAmount)} × ${fmtPct(n.rate)} ${eq} ${tail}`;
}

/** 요율 없는 합산 노드의 breakdown 라인: "└ 자식1 + 자식2 + … = 합계 ✓/⚠" */
function calcChildrenSumLine(n: CostNode): string | null {
  if (n.rate != null) return null;
  const withAmount = n.children.filter((c) => c.amount != null);
  if (withAmount.length < 2) return null;
  const sum = withAmount.reduce((s, c) => s + (c.amount as number), 0);
  const labels = withAmount.map((c) => c.canonicalName ?? c.amountCell);
  const joined = labels.slice(0, 5).join(' + ') + (labels.length > 5 ? ' + …' : '');
  const ok = n.amount != null && Boolean(matchesWithRounding(sum, n.amount));
  const tail = ok ? '✓' : `⚠ 차이 ${fmtAmount(sum - (n.amount ?? 0))}`;
  return `└ ${joined} = ${fmtAmount(sum)} ${tail}`;
}

function calcLineFor(n: CostNode): string | null {
  return calcBreakdownLine(n) ?? calcChildrenSumLine(n);
}

export function CostTreeView({ tree }: { tree: CostTree }) {
  if (!tree.summarySheet) {
    return (
      <Panel className="tree-section">
        <p className="tree-empty">원가계산서 요약시트를 찾지 못했습니다. {tree.issues.join(' ')}</p>
      </Panel>
    );
  }
  const rows = flattenTree(tree);
  // 검증(detectItems)은 트리에서 뽑은 표준항목이 TREE_ITEM_MIN 미만이면 레거시 라벨 스캔으로 폴백한다.
  // 그 경우 이 표(트리)와 실제 검증 대상이 서로 다른 산출물이 되므로, 사용자가 "이 표를 확인했다"고
  // 오해하지 않도록 화면에서 분기를 드러낸다.
  const usesLegacyFallback = costTreeToDetectedItems(tree).length < TREE_ITEM_MIN;
  return (
    <Panel className="tree-section">
      <div className="section-heading">
        <span>원가계산서 인식 결과</span>
        <p className="tree-source">
          {tree.summarySheet} 시트 · 금액 열 {tree.amountColumn}
        </p>
      </div>
      {usesLegacyFallback && (
        <p className="tree-fallback-warning">
          ⚠ 트리에서 인식된 표준항목이 3개 미만이라, 검증은 레거시 라벨 스캔 방식으로 수행됩니다. 아래 표는 참고용입니다.
        </p>
      )}
      {tree.issues.length > 0 && <p className="tree-issues">⚠ {tree.issues.join(' / ')}</p>}
      <div className="tree-table-wrap">
        <table className="tree-table">
          <thead>
            <tr>
              <th scope="col">항목</th>
              <th scope="col">금액</th>
              <th scope="col">요율</th>
              <th scope="col">산출기초</th>
              <th scope="col">근거 셀</th>
              <th scope="col">상태</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((n) => {
              const badge = STATUS_BADGE[n.status];
              const calcLine = calcLineFor(n);
              const rowCls = [n.status === 'MISMATCH' ? 'tree-row-warn' : '', calcLine ? 'tree-row-has-calc' : '']
                .filter(Boolean)
                .join(' ');
              return (
                <Fragment key={n.id}>
                  <tr className={rowCls}>
                    <td style={{ paddingLeft: `${n.depth * 16 + 8}px` }}>{n.label || n.canonicalName || '(라벨 없음)'}</td>
                    <td className="tree-num">{fmtAmount(n.amount)}</td>
                    <td className="tree-num">{n.rate != null ? fmtPct(n.rate) : ''}</td>
                    <td>{n.baseCells.join(', ')}</td>
                    <td>
                      {n.id}
                      {n.rateCell ? ` · 요율 ${n.rateCell}` : ''}
                    </td>
                    <td>
                      <span className={`tree-badge ${badge.cls}`}>{badge.text}</span>
                      {n.note && <span className="tree-note"> {n.note}</span>}
                    </td>
                  </tr>
                  {calcLine && (
                    <tr>
                      <td colSpan={6} className="tree-calc-line" style={{ paddingLeft: `${n.depth * 16 + 24}px` }}>
                        {calcLine}
                      </td>
                    </tr>
                  )}
                </Fragment>
              );
            })}
          </tbody>
        </table>
      </div>
    </Panel>
  );
}
