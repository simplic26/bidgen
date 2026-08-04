import { unzipSync, zipSync } from 'fflate';
import type { DerivedKey, NoticeForm } from '../types';

// ---------------------------------------------------------------------------
// hwpx 템플릿 치환 엔진
// hwpx = ZIP으로 압축된 OWPML XML 묶음. 템플릿의 {{자리표시자}}를 폼 값으로 치환한다.
// downloadHwpx를 제외한 모든 함수는 DOM 비의존 → node 검증 스크립트에서 재사용 가능.
// ---------------------------------------------------------------------------

export const NOTICE_TEMPLATE_URL = '/templates/notice-template.hwpx';

// 앱이 치환할 수 있는 자리표시자 키 (검증 스크립트와 동기화 유지)
// 템플릿에 일부만 있어도 됨 — 없는 키는 건너뛰고, 목록에 없는 마커만 오류가 된다.
export const NOTICE_PLACEHOLDER_KEYS = [
  // 발주 정보
  '사업소',
  '계약종류',
  '세부계약종류',
  '발주년월',
  '입찰방법',
  '입찰종류',
  '협정대상여부',
  // 공고 개요
  '입찰공고번호',
  '공고일',
  '공사명',
  '발주기관',
  '발주유형',
  '공고명의',
  // 금액 (유도 포함)
  '추정가격',
  '사급자재비',
  '추정금액',
  '예비가격기초금액',
  'A값',
  '순공사원가',
  // 공사 내용/조건
  '공사기간',
  '공사내용',
  '공사구분',
  '참가자격',
  '공동계약',
  // 일정 (직접 입력)
  '입찰개시일시',
  '입찰마감일시',
  '개찰일시',
  '개찰장소',
  '보증금납부기한',
  // 원가계산서 연동 보험료 5종
  '국민건강보험료',
  '국민연금보험료',
  '노인장기요양보험료',
  '퇴직공제부금비',
  '산업안전보건관리비',
  // 담당/문의/신고
  '담당자',
  '연락처',
  '문의부서1',
  '문의전화1',
  '문의부서2',
  '문의전화2',
  '신고전화',
  '신고이메일',
  '이의제기담당자',
  '이의제기전화',
  '이의제기이메일',
] as const;

const PROCUREMENT_LABELS = { CONSTRUCTION: '공사', SERVICE: '용역', GOODS: '물품' } as const;

/** 배포 서버에서 공고문 hwpx 템플릿을 내려받습니다. */
export async function loadNoticeTemplate(url: string = NOTICE_TEMPLATE_URL): Promise<Uint8Array> {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`공고문 템플릿을 불러오지 못했습니다 (HTTP ${res.status})`);
  return new Uint8Array(await res.arrayBuffer());
}

/** XML 텍스트 노드에 안전하게 삽입되도록 이스케이프합니다. (& 를 먼저 치환) */
export function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/** 숫자 문자열을 한국식 천단위 콤마 + "원" 표기로 바꿉니다. 숫자가 아니면 원문 그대로. */
export function formatKoreanAmount(raw: string): string {
  const digits = raw.replace(/[^\d]/g, '');
  if (!digits) return raw.trim();
  return `${Number(digits).toLocaleString('ko-KR')}원`;
}

/** YYYY-MM-DD를 "YYYY년 M월 D일"로 바꿉니다. 형식이 다르면 원문 그대로. */
export function formatKoreanDate(raw: string): string {
  const m = raw.trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return raw.trim();
  return `${m[1]}년 ${Number(m[2])}월 ${Number(m[3])}일`;
}

/** month input 값(YYYY-MM)을 발주년월 양식 "YYYY/MM"으로 바꿉니다. 형식이 다르면 원문 그대로. */
export function formatOrderMonth(raw: string): string {
  const m = raw.trim().match(/^(\d{4})-(\d{2})$/);
  if (!m) return raw.trim();
  return `${m[1]}/${m[2]}`;
}

/** 숫자 문자열을 천단위 콤마 표기(원 접미어 없음)로 바꿉니다. 숫자가 아니면 원문 그대로. */
export function formatNumberComma(raw: string): string {
  const digits = raw.replace(/[^\d]/g, '');
  if (!digits) return raw.trim();
  return Number(digits).toLocaleString('ko-KR');
}

/**
 * 폼 값 + 유도 결과(resolveNoticeValues의 values)를 자리표시자 키 → 치환 문자열 맵으로 변환합니다.
 * 보험료 5종은 견본 표기(콤마만, '원' 없음), A값·순공사원가 등 금액류는 "1,234원" 표기.
 */
export function buildPlaceholderMap(form: NoticeForm, resolved: Record<DerivedKey, string>): Record<string, string> {
  return {
    사업소: form.office.trim(),
    계약종류: form.contractType.trim(),
    세부계약종류: form.contractDetail.trim(),
    발주년월: formatOrderMonth(form.orderMonth),
    입찰방법: form.bidMethod.trim(),
    // 템플릿의 "({{입찰종류}}, 전자입찰)" 같은 문맥용 — 입찰방법에서 괄호 부분 제거 (제한경쟁(전자입찰) → 제한경쟁)
    입찰종류: form.bidMethod.replace(/\([^)]*\)\s*$/, '').trim(),
    협정대상여부: form.agreementTarget.trim(),
    입찰공고번호: form.noticeNumber.trim(),
    공고일: formatKoreanDate(form.noticeDate),
    공사명: form.projectName.trim(),
    발주기관: resolved.발주기관.trim(),
    발주유형: PROCUREMENT_LABELS[form.procurementType],
    공고명의: resolved.공고명의.trim(),
    추정가격: formatKoreanAmount(form.estimatedPrice),
    사급자재비: formatKoreanAmount(form.ownerMaterialCost),
    추정금액: formatKoreanAmount(resolved.추정금액),
    예비가격기초금액: formatKoreanAmount(resolved.예비가격기초금액),
    A값: formatKoreanAmount(resolved.A값),
    순공사원가: formatKoreanAmount(resolved.순공사원가),
    공사기간: form.period.trim(),
    공사내용: form.description.trim(),
    공사구분: form.constructionClass.trim(),
    참가자격: form.qualification.trim(),
    공동계약: form.jointContract.trim(),
    입찰개시일시: form.bidStart.trim(),
    입찰마감일시: form.bidDeadline.trim(),
    개찰일시: form.openingDate.trim(),
    개찰장소: resolved.개찰장소.trim(),
    보증금납부기한: form.depositDeadline.trim(),
    국민건강보험료: formatNumberComma(resolved.국민건강보험료),
    국민연금보험료: formatNumberComma(resolved.국민연금보험료),
    노인장기요양보험료: formatNumberComma(resolved.노인장기요양보험료),
    퇴직공제부금비: formatNumberComma(resolved.퇴직공제부금비),
    산업안전보건관리비: formatNumberComma(resolved.산업안전보건관리비),
    담당자: form.contactName.trim(),
    연락처: form.contactPhone.trim(),
    문의부서1: resolved.문의부서1.trim(),
    문의전화1: form.contactPhone.trim(),
    문의부서2: form.inquiryDept2.trim(),
    문의전화2: form.inquiryPhone2.trim(),
    신고전화: form.reportPhone.trim(),
    신고이메일: form.reportEmail.trim(),
    이의제기담당자: form.objectionContact.trim(),
    이의제기전화: form.objectionPhone.trim(),
    이의제기이메일: form.objectionEmail.trim(),
  };
}

/** XML 문자열에 남아 있는 {{...}} 마커 목록을 반환합니다. */
export function findUnreplacedPlaceholders(xml: string): string[] {
  return xml.match(/\{\{[^}]{1,40}\}\}/g) ?? [];
}

/**
 * 템플릿 hwpx의 Contents/section*.xml에서 {{키}}를 값으로 치환해 새 hwpx 바이트를 만듭니다.
 * - 값은 escapeXml로 이스케이프됩니다.
 * - 치환 후 잔여 마커가 있으면 throw (템플릿 준비 오류를 조기에 노출).
 * - mimetype 엔트리는 첫 번째·무압축으로 유지합니다.
 */
export function fillHwpxTemplate(template: Uint8Array, values: Record<string, string>): Uint8Array {
  const entries = unzipSync(template);
  const decoder = new TextDecoder('utf-8');
  const encoder = new TextEncoder();
  const leftovers: string[] = [];

  for (const name of Object.keys(entries)) {
    if (!/^Contents\/section\d+\.xml$/.test(name)) continue;
    let xml = decoder.decode(entries[name]);
    for (const [key, value] of Object.entries(values)) {
      xml = xml.split(`{{${key}}}`).join(escapeXml(value)); // replaceAll 대체 (tsconfig target ES2020)
    }
    leftovers.push(...findUnreplacedPlaceholders(xml));
    entries[name] = encoder.encode(xml);
  }

  if (leftovers.length > 0) {
    throw new Error(`치환되지 않은 자리표시자가 있습니다: ${[...new Set(leftovers)].join(', ')}`);
  }

  // mimetype을 첫 엔트리·무압축(level 0)으로 기록 (OWPML 관례)
  const ordered: Parameters<typeof zipSync>[0] = {};
  if (entries['mimetype']) ordered['mimetype'] = [entries['mimetype'], { level: 0 }];
  for (const [name, data] of Object.entries(entries)) {
    if (name === 'mimetype') continue;
    ordered[name] = data;
  }
  return zipSync(ordered);
}

/** 생성된 hwpx 바이트를 파일로 다운로드합니다. (download.ts의 anchor 패턴 재사용) */
export function downloadHwpx(fileName: string, data: Uint8Array): void {
  const blob = new Blob([data as BlobPart], { type: 'application/hwp+zip' });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}
