import { CheckCircle2, Circle, LoaderCircle } from '../components/icons';
import { useEffect, useState } from 'react';
import type { FileMeta, ProgressStep } from '../types';
import { Panel } from '../components/ui';

interface ProgressScreenProps {
  excelFile: FileMeta | null;
  steps: ProgressStep[];
  onDone: () => void;
}

export function ProgressScreen({ excelFile, steps, onDone }: ProgressScreenProps) {
  const [activeIndex, setActiveIndex] = useState(0);

  useEffect(() => {
    if (activeIndex >= steps.length) {
      const doneTimer = window.setTimeout(onDone, 350);
      return () => window.clearTimeout(doneTimer);
    }

    const timer = window.setTimeout(() => setActiveIndex((index) => index + 1), 220);
    return () => window.clearTimeout(timer);
  }, [activeIndex, onDone, steps.length]);

  const progress = Math.min(100, Math.round((activeIndex / steps.length) * 100));
  const current = steps[Math.min(activeIndex, steps.length - 1)];

  return (
    <main className="app-shell progress-layout">
      <aside className="workflow-rail">
        <span>3 / 5 단계</span>
        <strong>자동검증</strong>
      </aside>

      <section className="workspace centered-workspace">
        <Panel className="progress-panel">
          <LoaderCircle className="spin" size={28} />
          <h1>원가계산서를 검증하고 있습니다</h1>
          <p>{excelFile?.name ?? '추정가격내역서_최종.xlsx'}</p>
          <div className="progress-track">
            <div style={{ width: `${progress}%` }} />
          </div>
          <strong className="progress-percent">{progress}%</strong>

          <div className="progress-steps">
            {steps.map((step, index) => {
              const isDone = index < activeIndex;
              const isActive = index === activeIndex;
              return (
                <div className={`progress-step ${isActive ? 'is-active' : ''}`} key={step.label}>
                  {isDone ? <CheckCircle2 size={18} /> : <Circle size={18} />}
                  <div>
                    <strong>{step.label}</strong>
                    <span>{step.detail}</span>
                  </div>
                </div>
              );
            })}
          </div>

          <div className="current-check">
            <span>현재 확인 중</span>
            <strong>{current?.label ?? '판단근거 생성'}</strong>
            <p>추정가격내역서!H32 · 산업안전보건관리비</p>
          </div>
        </Panel>
      </section>
    </main>
  );
}
