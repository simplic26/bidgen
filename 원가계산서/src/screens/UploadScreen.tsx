import { FileSpreadsheet, FileText, Play, ShieldCheck } from '../components/icons';
import type { FileMeta, LaborRate, ProcurementType, RateCriterion, ReferenceFiles } from '../types';
import { Button, Panel, UploadPanel } from '../components/ui';
import { buildWorkbookIR } from '../utils/excel';
import { parseLaborWorkbook, parseRateWorkbook } from '../utils/criteria';
import type { WorkbookIR } from '../types';

const procurementOptions: Array<{ value: ProcurementType; label: string; description: string }> = [
  { value: 'CONSTRUCTION', label: '공사', description: '요율, 노임단가, 표준품셈까지 상세 검증' },
  { value: 'SERVICE', label: '용역', description: '인건비, 경비, 합계 중심 공통 검증' },
  { value: 'GOODS', label: '물품', description: '수량, 단가, 부가세, 총액 검증' },
];

interface UploadScreenProps {
  procurementType: ProcurementType;
  excelFile: FileMeta | null;
  referenceFiles: ReferenceFiles;
  onProcurementTypeChange: (type: ProcurementType) => void;
  onExcelChange: (file: FileMeta, ir: WorkbookIR | null) => void;
  onRateExcel: (meta: FileMeta, criteria: RateCriterion[]) => void;
  onLaborExcel: (meta: FileMeta, rates: LaborRate[]) => void;
  onStandardPdf: (meta: FileMeta) => void;
  onSkipToNotice: () => void;
  onStart: () => void;
}

function fileToMeta(file: File, detail: string): FileMeta {
  return {
    name: file.name,
    sizeLabel: `${(file.size / 1024 / 1024 || 0.1).toFixed(1)}MB`,
    detail,
  };
}

export function UploadScreen({
  procurementType,
  excelFile,
  referenceFiles,
  onProcurementTypeChange,
  onExcelChange,
  onRateExcel,
  onLaborExcel,
  onStandardPdf,
  onSkipToNotice,
  onStart,
}: UploadScreenProps) {
  return (
    <main className="app-shell upload-layout">
      <aside className="workflow-rail">
        <div className="brand-mark">
          <ShieldCheck size={22} />
        </div>
        <div>
          <span>1 / 5 단계</span>
          <strong>파일 업로드</strong>
        </div>
      </aside>

      <section className="workspace">
        <header className="topbar">
          <div>
            <h1>원가계산서 자동검증</h1>
            <p>Excel 계산 구조와 기준자료 근거를 연결해 검증결과를 확인합니다.</p>
          </div>
          <div className="topbar-actions">
            <Button onClick={onSkipToNotice} variant="ghost">
              공고문 생성 바로가기
            </Button>
            <Button variant="ghost">도움말</Button>
          </div>
        </header>

        <Panel>
          <div className="section-heading">
            <span>발주 유형</span>
            <p>공사는 시연에서 가장 깊게 검증합니다.</p>
          </div>
          <div className="procurement-grid">
            {procurementOptions.map((option) => (
              <button
                className={`procurement-option ${procurementType === option.value ? 'is-selected' : ''}`}
                key={option.value}
                onClick={() => onProcurementTypeChange(option.value)}
                type="button"
              >
                <strong>{option.label}</strong>
                <span>{option.description}</span>
              </button>
            ))}
          </div>
        </Panel>

        <div className="upload-grid">
          <Panel>
            <UploadPanel
              accept=".xlsx,.xls"
              action="Excel 파일 선택"
              description="원가계산서 Excel 파일을 여기에 놓거나 선택하세요."
              label="Excel 파일 선택"
              onChange={async (event) => {
                const file = event.currentTarget.files?.[0];
                if (!file) return;
                try {
                  const ir = await buildWorkbookIR(file, procurementType);
                  const detail = `${ir.totals.sheetCount}개 시트 · ${ir.totals.formulaCount.toLocaleString()}개 수식`;
                  onExcelChange(fileToMeta(file, detail), ir);
                } catch {
                  onExcelChange(fileToMeta(file, '시트 정보를 읽지 못했습니다'), null);
                }
              }}
              title="검증할 원가계산서"
              meta={
                excelFile ? (
                  <span>
                    <FileSpreadsheet size={16} /> {excelFile.name} · {excelFile.detail}
                  </span>
                ) : null
              }
            />
          </Panel>

          <Panel>
            <UploadPanel
              accept=".xlsx,.xls"
              action="제비율 Excel 선택"
              description="제비율 적용기준 Excel(.xlsx)을 등록하세요. 요율·산출기초 근거로 사용됩니다."
              label="제비율 Excel 선택"
              onChange={async (event) => {
                const file = event.currentTarget.files?.[0];
                if (!file) return;
                try {
                  const { criteria } = await parseRateWorkbook(file, procurementType);
                  onRateExcel(fileToMeta(file, `제비율 ${criteria.length}개 추출 · 셀 근거 포함`), criteria);
                } catch {
                  onRateExcel(fileToMeta(file, '요율 추출 실패 · 수기 입력 필요'), []);
                }
              }}
              title={
                <>
                  제비율 기준 Excel <em className="req-tag">필수</em>
                </>
              }
              meta={
                referenceFiles.rateFile ? (
                  <span>
                    <FileSpreadsheet size={16} /> {referenceFiles.rateFile.name} · {referenceFiles.rateFile.detail}
                  </span>
                ) : null
              }
            />
          </Panel>

          <Panel>
            <UploadPanel
              accept=".xlsx,.xls"
              action="노임단가 Excel 선택"
              description="노임단가표 Excel(.xlsx)을 등록하세요. 직종별 단가 비교에 사용됩니다."
              label="노임단가 Excel 선택"
              onChange={async (event) => {
                const file = event.currentTarget.files?.[0];
                if (!file) return;
                try {
                  const { laborRates } = await parseLaborWorkbook(file, procurementType);
                  onLaborExcel(fileToMeta(file, `직종 ${laborRates.length}개 추출 · 셀 근거 포함`), laborRates);
                } catch {
                  onLaborExcel(fileToMeta(file, '단가 추출 실패'), []);
                }
              }}
              title={
                <>
                  노임단가 기준 Excel <em className="req-tag">필수</em>
                </>
              }
              meta={
                referenceFiles.laborFile ? (
                  <span>
                    <FileSpreadsheet size={16} /> {referenceFiles.laborFile.name} · {referenceFiles.laborFile.detail}
                  </span>
                ) : null
              }
            />
          </Panel>

          <Panel>
            <UploadPanel
              accept=".pdf"
              action="표준품셈 PDF 선택"
              description="표준품셈·가격산정지침 PDF(.pdf)는 선택 입력입니다. 없으면 표준품셈 검증은 미수행됩니다."
              label="표준품셈 PDF 선택"
              onChange={(event) => {
                const file = event.currentTarget.files?.[0];
                if (!file) return;
                onStandardPdf(fileToMeta(file, '표준품셈 · 페이지 근거 (선택)'));
              }}
              title={
                <>
                  표준품셈 PDF <em className="opt-tag">선택</em>
                </>
              }
              meta={
                referenceFiles.standardPdf ? (
                  <span>
                    <FileText size={16} /> {referenceFiles.standardPdf.name} · {referenceFiles.standardPdf.detail}
                  </span>
                ) : null
              }
            />
          </Panel>
        </div>

        <div className="footer-actions">
          <p>제비율·노임단가 Excel이 없어도 산술·수식·합계 검증은 진행할 수 있습니다.</p>
          <Button disabled={!excelFile} icon={<Play size={16} />} onClick={onStart} variant="primary">
            자동검증 시작
          </Button>
        </div>
      </section>
    </main>
  );
}
