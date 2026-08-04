import { useMemo, useState } from 'react';
import type { ReactNode } from 'react';
import { ArrowLeft, FileSpreadsheet, FileText } from '../components/icons';
import { Button, Panel } from '../components/ui';
import { AGREEMENT_TARGETS, BID_METHODS, CONTRACT_TYPES, CONTRACT_TYPE_DETAILS, OFFICES } from '../data/noticeOptions';
import type { DerivedKey, NoticeForm, NoticeOverrides, NoticePrefill, ProcurementType, ValueSource } from '../types';
import { buildPlaceholderMap, downloadHwpx, fillHwpxTemplate, loadNoticeTemplate } from '../utils/hwpx';
import { INSURANCE_KEYS, resolveNoticeValues } from '../utils/noticeDerive';

const CONSTRUCTION_CLASSES = ['전문공사', '종합공사'];
const JOINT_CONTRACT_OPTIONS = ['불허', '허용'];

interface NoticeScreenProps {
  prefill: NoticePrefill;
  onBack: () => void;
}

function todayISO(): string {
  const now = new Date();
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const dd = String(now.getDate()).padStart(2, '0');
  return `${now.getFullYear()}-${mm}-${dd}`;
}

function initialForm(prefill: NoticePrefill): NoticeForm {
  return {
    office: '',
    contractType: '',
    contractDetail: '',
    orderMonth: '',
    bidMethod: '',
    agreementTarget: '',
    projectName: prefill.projectName ?? '',
    estimatedPrice: prefill.totalAmount != null ? String(prefill.totalAmount) : '',
    noticeDate: todayISO(),
    period: '',
    procurementType: prefill.procurementType,
    bidStart: '',
    bidDeadline: '',
    openingDate: '',
    depositDeadline: '',
    description: '',
    constructionClass: '전문공사',
    jointContract: '불허',
    qualification: '',
    ownerMaterialCost: '0',
    noticeNumber: '',
    contactName: '',
    contactPhone: '',
    inquiryDept2: '',
    inquiryPhone2: '',
    reportPhone: '',
    reportEmail: '',
    objectionContact: '',
    objectionPhone: '',
    objectionEmail: '',
  };
}

interface FieldProps {
  label: string;
  hint?: string;
  children: ReactNode;
}

function Field({ label, hint, children }: FieldProps) {
  return (
    <label className="notice-field">
      <span>
        {label}
        {hint ? <em>{hint}</em> : null}
      </span>
      {children}
    </label>
  );
}

const SOURCE_LABELS: Record<ValueSource, string> = {
  manual: '',
  auto: '자동',
  linked: '원가연동',
  edited: '수정됨',
};

interface DerivedFieldProps {
  label: string;
  value: string;
  source: ValueSource;
  onChange: (value: string) => void;
  onReset: () => void;
  placeholder?: string;
}

function DerivedField({ label, value, source, onChange, onReset, placeholder }: DerivedFieldProps) {
  return (
    <label className="notice-field">
      <span>
        {label}
        <em className={`source-badge source-${source}`}>{SOURCE_LABELS[source]}</em>
        {source === 'edited' ? (
          <button className="derived-reset" onClick={onReset} title="자동값으로 되돌리기" type="button">
            ↺ 자동값
          </button>
        ) : null}
      </span>
      <input onChange={(event) => onChange(event.currentTarget.value)} placeholder={placeholder} type="text" value={value} />
    </label>
  );
}

interface CollapsiblePanelProps {
  title: string;
  description: string;
  badge?: string;
  children: ReactNode;
}

function CollapsiblePanel({ title, description, badge, children }: CollapsiblePanelProps) {
  const [open, setOpen] = useState(false);
  return (
    <Panel>
      <button className="collapsible-header" onClick={() => setOpen((prev) => !prev)} type="button">
        <div className="section-heading">
          <span>
            {title} {badge ? <em className="panel-badge">{badge}</em> : null}
          </span>
          <p>{description}</p>
        </div>
        <span className="collapse-chevron">{open ? '▲ 접기' : '▼ 펼치기'}</span>
      </button>
      {open ? children : null}
    </Panel>
  );
}

export function NoticeScreen({ prefill, onBack }: NoticeScreenProps) {
  const [form, setForm] = useState<NoticeForm>(() => initialForm(prefill));
  const [overrides, setOverrides] = useState<NoticeOverrides>({});
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const [generatedName, setGeneratedName] = useState('');

  // 유도값은 상태에 저장하지 않고 매번 재계산 — 핵심 입력이 바뀌면 즉시 반영된다
  const { values: derived, sources } = useMemo(
    () => resolveNoticeValues(form, prefill, overrides),
    [form, prefill, overrides],
  );

  const missingRequired = [
    form.office === '' ? '사업소' : null,
    form.contractType === '' ? '계약종류' : null,
    form.contractDetail === '' ? '세부계약종류' : null,
    form.orderMonth === '' ? '발주년월' : null,
    form.bidMethod === '' ? '입찰방법' : null,
    form.agreementTarget === '' ? '협정대상여부' : null,
    form.projectName.trim() === '' ? '공사명' : null,
    form.estimatedPrice.trim() === '' ? '추정가격' : null,
    form.noticeDate === '' ? '공고일' : null,
    form.period.trim() === '' ? '공사기간' : null,
  ].filter((name): name is string => name !== null);
  const requiredReady = missingRequired.length === 0;

  const contractDetailOptions = CONTRACT_TYPE_DETAILS[form.contractType] ?? [];

  const derivedKeys = Object.keys(derived) as DerivedKey[];
  const autoCount = derivedKeys.filter((key) => sources[key] !== 'edited' && derived[key] !== '').length;
  const editedCount = derivedKeys.filter((key) => sources[key] === 'edited').length;

  function update<K extends keyof NoticeForm>(key: K, value: NoticeForm[K]) {
    setForm((prev) => ({ ...prev, [key]: value }));
  }

  function updateContractType(value: string) {
    setForm((prev) => ({ ...prev, contractType: value, contractDetail: '' }));
  }

  function overrideDerived(key: DerivedKey, value: string) {
    setOverrides((prev) => ({ ...prev, [key]: value }));
  }

  function resetDerived(key: DerivedKey) {
    setOverrides((prev) => {
      const next = { ...prev };
      delete next[key];
      return next;
    });
  }

  function derivedFieldProps(key: DerivedKey) {
    return {
      value: derived[key],
      source: sources[key],
      onChange: (value: string) => overrideDerived(key, value),
      onReset: () => resetDerived(key),
    };
  }

  async function handleGenerate() {
    setBusy(true);
    setError('');
    setGeneratedName('');
    try {
      const template = await loadNoticeTemplate();
      const data = fillHwpxTemplate(template, buildPlaceholderMap(form, derived));
      const fileName = `입찰공고문_${form.projectName.trim() || '무제'}.hwpx`;
      downloadHwpx(fileName, data);
      setGeneratedName(fileName);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : '공고문 생성 중 오류가 발생했습니다.');
    } finally {
      setBusy(false);
    }
  }

  return (
    <main className="app-shell upload-layout">
      <aside className="workflow-rail">
        <div className="brand-mark">
          <FileText size={22} />
        </div>
        <div>
          <span>부가 기능</span>
          <strong>공고문 생성</strong>
        </div>
      </aside>

      <section className="workspace">
        <header className="topbar">
          <div>
            <h1>입찰공고문 생성</h1>
            <p>핵심 정보만 입력하면 나머지는 자동으로 채워 hwpx 공고문을 만들어 드립니다.</p>
          </div>
          <Button icon={<ArrowLeft size={16} />} onClick={onBack} variant="ghost">
            돌아가기
          </Button>
        </header>

        {prefill.sourceFileName ? (
          <div className="notice-prefill-banner">
            <FileSpreadsheet size={16} />
            <span>
              업로드한 원가계산서에서 자동 입력됨 · {prefill.sourceFileName}
              {prefill.totalAmount != null ? ` · 추정가격 ${prefill.totalAmount.toLocaleString('ko-KR')}원` : ''}
              {prefill.insurances ? ` · 보험료 ${Object.keys(prefill.insurances).length}종 추출` : ''}
            </span>
          </div>
        ) : null}

        <Panel>
          <div className="section-heading">
            <span>핵심 입력</span>
            <p>이 10개만 채우면 생성할 수 있습니다. 나머지는 아래에서 자동 생성되며 수정할 수 있습니다.</p>
          </div>
          <div className="notice-form-grid">
            <Field label="사업소" hint="필수">
              <select onChange={(event) => update('office', event.currentTarget.value)} value={form.office}>
                <option disabled value="">
                  선택하세요
                </option>
                {OFFICES.map((office) => (
                  <option key={office} value={office}>
                    {office}
                  </option>
                ))}
              </select>
            </Field>
            <Field label="계약종류" hint="필수">
              <select onChange={(event) => updateContractType(event.currentTarget.value)} value={form.contractType}>
                <option disabled value="">
                  선택하세요
                </option>
                {CONTRACT_TYPES.map((type) => (
                  <option key={type} value={type}>
                    {type}
                  </option>
                ))}
              </select>
            </Field>
            <Field label="세부계약종류" hint={form.contractType ? '필수' : '필수 · 계약종류 먼저 선택'}>
              <select
                disabled={form.contractType === ''}
                onChange={(event) => update('contractDetail', event.currentTarget.value)}
                value={form.contractDetail}
              >
                <option disabled value="">
                  {form.contractType ? '선택하세요' : '계약종류를 먼저 선택하세요'}
                </option>
                {contractDetailOptions.map((detail) => (
                  <option key={detail} value={detail}>
                    {detail}
                  </option>
                ))}
              </select>
            </Field>
            <Field label="발주년월" hint="필수">
              <input onChange={(event) => update('orderMonth', event.currentTarget.value)} type="month" value={form.orderMonth} />
            </Field>
            <Field label="입찰방법" hint="필수">
              <select onChange={(event) => update('bidMethod', event.currentTarget.value)} value={form.bidMethod}>
                <option disabled value="">
                  선택하세요
                </option>
                {BID_METHODS.map((method) => (
                  <option key={method} value={method}>
                    {method}
                  </option>
                ))}
              </select>
            </Field>
            <Field label="협정대상여부" hint="필수">
              <select onChange={(event) => update('agreementTarget', event.currentTarget.value)} value={form.agreementTarget}>
                <option disabled value="">
                  선택하세요
                </option>
                {AGREEMENT_TARGETS.map((target) => (
                  <option key={target} value={target}>
                    {target}
                  </option>
                ))}
              </select>
            </Field>
            <Field label="공사명" hint="필수">
              <input
                onChange={(event) => update('projectName', event.currentTarget.value)}
                placeholder="판교지사 약품주입설비 개선공사"
                type="text"
                value={form.projectName}
              />
            </Field>
            <Field label="추정가격" hint="필수 · 부가세 별도">
              <input
                inputMode="numeric"
                onChange={(event) => update('estimatedPrice', event.currentTarget.value)}
                placeholder="296380000"
                type="text"
                value={form.estimatedPrice}
              />
            </Field>
            <Field label="공고일" hint="필수 · 기본 오늘">
              <input onChange={(event) => update('noticeDate', event.currentTarget.value)} type="date" value={form.noticeDate} />
            </Field>
            <Field label="공사기간" hint="필수">
              <input
                onChange={(event) => update('period', event.currentTarget.value)}
                placeholder="착공일로부터 3개월"
                type="text"
                value={form.period}
              />
            </Field>
          </div>
        </Panel>

        <CollapsiblePanel
          description="일정·공사조건·담당 연락처 등. 비워 두면 공고문에서 해당 자리만 비어 나옵니다."
          title="추가 입력 (선택)"
        >
          <div className="notice-form-grid">
            <Field label="가격입찰서 접수 개시 일시">
              <input
                onChange={(event) => update('bidStart', event.currentTarget.value)}
                placeholder="2026. 06. 22.(월) 09:00"
                type="text"
                value={form.bidStart}
              />
            </Field>
            <Field label="가격입찰서 접수 마감 일시">
              <input
                onChange={(event) => update('bidDeadline', event.currentTarget.value)}
                placeholder="2026. 07. 03.(금) 13:00"
                type="text"
                value={form.bidDeadline}
              />
            </Field>
            <Field label="개찰 일시">
              <input
                onChange={(event) => update('openingDate', event.currentTarget.value)}
                placeholder="2026. 07. 03.(금) 14:00 이후"
                type="text"
                value={form.openingDate}
              />
            </Field>
            <Field label="입찰보증금 납부기한">
              <input
                onChange={(event) => update('depositDeadline', event.currentTarget.value)}
                placeholder="2026. 07. 02.(목) 18:00"
                type="text"
                value={form.depositDeadline}
              />
            </Field>
            <Field label="공사내용">
              <input
                onChange={(event) => update('description', event.currentTarget.value)}
                placeholder="약품주입설비 개선"
                type="text"
                value={form.description}
              />
            </Field>
            <Field label="공사구분">
              <select onChange={(event) => update('constructionClass', event.currentTarget.value)} value={form.constructionClass}>
                {CONSTRUCTION_CLASSES.map((cls) => (
                  <option key={cls} value={cls}>
                    {cls}
                  </option>
                ))}
              </select>
            </Field>
            <Field label="공동계약">
              <select onChange={(event) => update('jointContract', event.currentTarget.value)} value={form.jointContract}>
                {JOINT_CONTRACT_OPTIONS.map((option) => (
                  <option key={option} value={option}>
                    {option}
                  </option>
                ))}
              </select>
            </Field>
            <Field label="입찰참가자격 (요구 면허/업종)">
              <input
                onChange={(event) => update('qualification', event.currentTarget.value)}
                placeholder="기계설비·가스공사업 면허(주력분야 : 기계설비공사)"
                type="text"
                value={form.qualification}
              />
            </Field>
            <Field label="사급자재비" hint="부가세 별도">
              <input
                inputMode="numeric"
                onChange={(event) => update('ownerMaterialCost', event.currentTarget.value)}
                type="text"
                value={form.ownerMaterialCost}
              />
            </Field>
            <Field label="입찰공고번호">
              <input
                onChange={(event) => update('noticeNumber', event.currentTarget.value)}
                placeholder="제309100-00000호"
                type="text"
                value={form.noticeNumber}
              />
            </Field>
            <Field label="담당자">
              <input
                onChange={(event) => update('contactName', event.currentTarget.value)}
                placeholder="계약팀 홍길동"
                type="text"
                value={form.contactName}
              />
            </Field>
            <Field label="연락처 (입찰서류 문의 전화)">
              <input
                onChange={(event) => update('contactPhone', event.currentTarget.value)}
                placeholder="031)000-0000"
                type="text"
                value={form.contactPhone}
              />
            </Field>
            <Field label="공사내용 문의 부서">
              <input
                onChange={(event) => update('inquiryDept2', event.currentTarget.value)}
                placeholder="판교지사 기계부"
                type="text"
                value={form.inquiryDept2}
              />
            </Field>
            <Field label="공사내용 문의 전화">
              <input
                onChange={(event) => update('inquiryPhone2', event.currentTarget.value)}
                placeholder="031)000-0000"
                type="text"
                value={form.inquiryPhone2}
              />
            </Field>
            <Field label="비리신고 전화">
              <input
                onChange={(event) => update('reportPhone', event.currentTarget.value)}
                placeholder="031)000-0000"
                type="text"
                value={form.reportPhone}
              />
            </Field>
            <Field label="비리신고 이메일">
              <input
                onChange={(event) => update('reportEmail', event.currentTarget.value)}
                placeholder="report@example.co.kr"
                type="text"
                value={form.reportEmail}
              />
            </Field>
            <Field label="이의제기 담당자">
              <input
                onChange={(event) => update('objectionContact', event.currentTarget.value)}
                placeholder="고객지원부 ○○○ 부장"
                type="text"
                value={form.objectionContact}
              />
            </Field>
            <Field label="이의제기 전화">
              <input
                onChange={(event) => update('objectionPhone', event.currentTarget.value)}
                placeholder="031)000-0000"
                type="text"
                value={form.objectionPhone}
              />
            </Field>
            <Field label="이의제기 이메일">
              <input
                onChange={(event) => update('objectionEmail', event.currentTarget.value)}
                placeholder="contact@example.co.kr"
                type="text"
                value={form.objectionEmail}
              />
            </Field>
          </div>
        </CollapsiblePanel>

        <CollapsiblePanel
          badge={`자동 ${autoCount} · 수정 ${editedCount}`}
          description="핵심 입력에서 자동으로 계산·구성된 값입니다. 필요하면 직접 수정할 수 있습니다."
          title="자동 생성 값 확인"
        >
          <div className="derived-group-title">금액 (추정가격·사급자재비 기준 자동계산)</div>
          <div className="notice-form-grid">
            <DerivedField label="예비가격기초금액 (부가세 포함)" placeholder="추정가격 입력 시 자동계산" {...derivedFieldProps('예비가격기초금액')} />
            <DerivedField label="추정금액" placeholder="추정가격 입력 시 자동계산" {...derivedFieldProps('추정금액')} />
            <DerivedField label="적격심사 A값" placeholder="보험료 5종 입력 시 합계 자동계산" {...derivedFieldProps('A값')} />
            <DerivedField label="순공사원가" placeholder="원가계산서에서 추출 시 자동" {...derivedFieldProps('순공사원가')} />
          </div>
          <div className="derived-group-title">장소·명의 (사업소 기준 자동구성)</div>
          <div className="notice-form-grid">
            <DerivedField label="개찰 장소" placeholder="사업소 선택 시 자동구성" {...derivedFieldProps('개찰장소')} />
            <DerivedField label="공고 명의" placeholder="사업소 선택 시 자동구성" {...derivedFieldProps('공고명의')} />
            <DerivedField label="입찰서류 문의 부서" placeholder="사업소 선택 시 자동구성" {...derivedFieldProps('문의부서1')} />
            <DerivedField label="발주기관" {...derivedFieldProps('발주기관')} />
          </div>
          <div className="derived-group-title">보험료 5종 (원가계산서 연동 — 수정하면 A값에 반영)</div>
          <div className="notice-form-grid">
            {INSURANCE_KEYS.map((key) => (
              <DerivedField key={key} label={key} placeholder="원가계산서에서 추출 시 자동" {...derivedFieldProps(key)} />
            ))}
          </div>
        </CollapsiblePanel>

        <div className="footer-actions">
          <p>
            {error ? (
              <span className="notice-error">{error}</span>
            ) : generatedName ? (
              <span className="notice-success">{generatedName} 다운로드 완료 · 한글에서 열어 확인하세요.</span>
            ) : !requiredReady ? (
              <span className="notice-missing">
                필수 항목을 입력하면 생성 버튼이 활성화됩니다: {missingRequired.join(', ')}
              </span>
            ) : (
              '생성된 공고문은 hwpx 파일로 다운로드되며 한글(Hancom Office)에서 열 수 있습니다.'
            )}
          </p>
          <Button disabled={!requiredReady || busy} icon={<FileText size={16} />} onClick={handleGenerate} variant="primary">
            {busy ? '생성 중…' : '공고문 생성 (hwpx)'}
          </Button>
        </div>
      </section>
    </main>
  );
}
