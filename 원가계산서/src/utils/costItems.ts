// ---------------------------------------------------------------------------
// 숫자·요율 파싱
// ---------------------------------------------------------------------------

export function parseKoreanNumber(raw: string | number | boolean | null): number | null {
  if (raw == null || raw === '') return null;
  if (typeof raw === 'number') return Number.isFinite(raw) ? raw : null;
  if (typeof raw === 'boolean') return null;
  let s = String(raw).trim().replace(/[,\s원]/g, '');
  let negative = false;
  if (/^\(.*\)$/.test(s)) {
    negative = true;
    s = s.slice(1, -1);
  }
  if (s.endsWith('%')) s = s.slice(0, -1);
  const n = Number(s);
  if (!Number.isFinite(n)) return null;
  return negative ? -n : n;
}

export function parseRate(raw: string | number | boolean | null, numberFormat: string | null): number | null {
  const n = parseKoreanNumber(raw);
  if (n == null) return null;
  if (numberFormat && numberFormat.includes('%')) return n; // SheetJS는 %서식 값을 분수로 저장
  if (typeof raw === 'string' && raw.includes('%')) return n / 100;
  if (n > 1) return n / 100; // 3.2 → 0.032
  return n; // 이미 분수
}

export function normalizeLabel(s: string): string {
  return s.replace(/\s+/g, '').toLowerCase();
}

// ---------------------------------------------------------------------------
// 표준 원가항목 사전
// ---------------------------------------------------------------------------

export interface CanonicalDef {
  canonical: string;
  category: string;
  aliases: string[]; // 정규화된 형태
  requiresReference: boolean;
  defaultRate?: number;
}

export const CANONICAL_ITEMS: CanonicalDef[] = [
  { canonical: '산재보험료', category: '제보험료', aliases: ['산업재해보상보험료', '산재보험료', '산재보험'], requiresReference: true },
  { canonical: '고용보험료', category: '제보험료', aliases: ['고용보험료', '고용보험'], requiresReference: true },
  { canonical: '국민건강보험료', category: '제보험료', aliases: ['국민건강보험료', '건강보험료'], requiresReference: true },
  { canonical: '노인장기요양보험료', category: '제보험료', aliases: ['노인장기요양보험료', '장기요양보험료'], requiresReference: true },
  // 아래 2개는 공고문 자동 채움용 탐지 목적 — 기준요율 검증 대상은 아니라 requiresReference: false (NEEDS_REVIEW 노이즈 방지)
  { canonical: '국민연금보험료', category: '제보험료', aliases: ['국민연금보험료', '연금보험료', '국민연금'], requiresReference: false },
  { canonical: '퇴직공제부금비', category: '제경비', aliases: ['건설근로자퇴직공제부금비', '건설근로자퇴직공제부금', '퇴직공제부금비', '퇴직공제부금'], requiresReference: false },
  { canonical: '산업안전보건관리비', category: '제경비', aliases: ['산업안전보건관리비', '안전보건관리비'], requiresReference: true },
  { canonical: '간접노무비', category: '노무비', aliases: ['간접노무비'], requiresReference: true },
  { canonical: '기타경비', category: '제경비', aliases: ['기타경비'], requiresReference: true },
  { canonical: '일반관리비', category: '일반관리비', aliases: ['일반관리비'], requiresReference: true },
  { canonical: '이윤', category: '이윤', aliases: ['이윤'], requiresReference: true },
  { canonical: '환경보전비', category: '제경비', aliases: ['환경보전비'], requiresReference: true },
  { canonical: '안전관리비', category: '제경비', aliases: ['안전관리비'], requiresReference: true },
  { canonical: '부가가치세', category: '부가세', aliases: ['부가가치세', '부가세', 'vat'], requiresReference: false, defaultRate: 0.1 },
];

// 긴 별칭 우선 매칭 (부분포함 오탐 최소화)
const ALIAS_INDEX: Array<{ alias: string; def: CanonicalDef }> = CANONICAL_ITEMS.flatMap((def) =>
  def.aliases.map((alias) => ({ alias, def })),
).sort((a, b) => b.alias.length - a.alias.length);

// ---------------------------------------------------------------------------
// 라벨 부기(annotation) 제거 + 경계 매칭
//
// 실제 파일의 항목 라벨은 이름 뒤/앞에 다양한 부기가 붙는다:
//   "산업재해보상보험료 [(노)X요율]"  "이      윤\n[(노+경+일)X요율]"
//   "국민건강보험료(실적정산)"        "부가가치세(VAT)"
//   "1) 산재보험료"                   "나. 간접노무비"    "① 이윤"
// 반대로 아래는 항목 라벨이 아니라 산출식·설명문·잡담이므로 매칭돼선 안 된다:
//   "이윤절사"  "부가가치세포함"  "공급가액+부가가치세"  "(재료비+노무비)×"
//   "조달청 기타경비 요율"  "※ 기타경비 : 복리후생비, 소모품비 …"  "안전관리비 대상액"
// 그래서 (1) 괄호류 부기와 선행 번호·불릿을 제거하고, (2) 남은 텍스트가 별칭과 "같거나"
// "별칭 뒤에 이름 문자(한글·라틴)가 전혀 남지 않는" 경우에만 매칭한다.
// (길이 여유(slack) 예산 방식은 "산업재해보상보험료 [(노)X요율]" 같은 정당한 라벨까지
//  통째로 죽여 제비율 기준요율 추출이 0건이 되는 문제가 있어 폐기했다.)
// ---------------------------------------------------------------------------

// 짝이 맞는 괄호류 한 겹. 안쪽부터 반복 적용하면 중첩 부기도 지워진다.
// 예: "이윤[(노+경+일)x요율]" → "이윤[x요율]" → "이윤"
const BRACKET_PAIR_RES: RegExp[] = [
  /\([^()]*\)/g,
  /\[[^[\]]*\]/g,
  /\{[^{}]*\}/g,
  /（[^（）]*）/g,
  /〔[^〔〕]*〕/g,
  /【[^【】]*】/g,
  /「[^「」]*」/g,
  /『[^『』]*』/g,
  /<[^<>]*>/g,
  /〈[^〈〉]*〉/g,
  /《[^《》]*》/g,
];

// 선행 번호·항목기호: "1)" "1." "가." "나)" "①" "ㅇ" "■" "▶" "※" "*" "-" 등.
// 한글 항목기호는 실제 열거에 쓰이는 음절만 허용한다([가-하] 같은 범위는 사실상 모든
// 한글 음절을 삼켜 "간.", "경)" 류 오제거를 부른다).
const LEADING_MARKER_RE = /^(?:[①-⑳]|[❶-❿]|\d{1,2}\s*[.)]|[가나다라마바사아자차카타파하]\s*[.)]|[ㅇㅁ○●◯■□▣▶▷◆◇※★*·•∙+\-–—~=]+)/;

// 양끝에 남은 구두점(":" "," "." "·" 등)
const EDGE_NOISE_RE = /^[^0-9a-z가-힣]+|[^0-9a-z가-힣]+$/g;

// "이름 문자" — 이게 별칭 뒤에 남아 있으면 다른 낱말이 이어진 것이므로 항목 라벨이 아니다.
const NAME_CHAR_RE = /[a-z가-힣]/;

/** 라벨에서 괄호 부기·선행 번호·양끝 구두점을 제거한 정규화 문자열을 돌려준다. */
export function stripLabelDecorations(label: string): string {
  let s = normalizeLabel(label);
  for (let i = 0; i < 5; i += 1) {
    const before = s;
    for (const re of BRACKET_PAIR_RES) s = s.replace(re, '');
    if (s === before) break;
  }
  for (let i = 0; i < 4; i += 1) {
    const next = s.replace(LEADING_MARKER_RE, '');
    if (next === s) break;
    s = next;
  }
  return s.replace(EDGE_NOISE_RE, '');
}

/** 별칭 사전 조회: 완전일치 우선 → 별칭으로 시작하고 뒤에 이름 문자가 없는 경우 */
function lookupAlias<T>(norm: string, index: Array<{ alias: string; def: T }>): T | null {
  if (!norm) return null;
  for (const { alias, def } of index) {
    if (norm === alias) return def;
  }
  for (const { alias, def } of index) {
    if (!norm.startsWith(alias)) continue;
    if (!NAME_CHAR_RE.test(norm.slice(alias.length))) return def; // "이윤15%" 등 숫자·기호 꼬리만 허용
  }
  return null;
}

export function matchCanonical(label: string): CanonicalDef | null {
  return lookupAlias(stripLabelDecorations(label), ALIAS_INDEX);
}

/** 라벨을 표준 원가항목명으로 매핑합니다. 매칭 실패 시 null. (기준자료 파서에서 재사용) */
export function matchCanonicalLabel(label: string): string | null {
  return matchCanonical(label)?.canonical ?? null;
}

/** canonical 이름으로 사전 정의를 조회합니다. (어댑터에서 category·requiresReference 참조용) */
export function getCanonicalDef(canonical: string): CanonicalDef | null {
  return CANONICAL_ITEMS.find((d) => d.canonical === canonical) ?? null;
}

// 트리 골격 항목 (요율 검증 대상이 아니라 CANONICAL_ITEMS와 분리 — DetectedItem으로는 내보내지 않음)
export interface StructuralDef {
  canonical: string;
  aliases: string[];
}

export const STRUCTURAL_ITEMS: StructuralDef[] = [
  { canonical: '직접재료비', aliases: ['직접재료비'] },
  { canonical: '간접재료비', aliases: ['간접재료비'] },
  { canonical: '재료비', aliases: ['재료비'] },
  { canonical: '직접노무비', aliases: ['직접노무비'] },
  { canonical: '노무비', aliases: ['노무비'] },
  { canonical: '경비', aliases: ['경비'] },
  { canonical: '순공사원가', aliases: ['순공사원가', '순용역원가', '순공사비', '순원가'] },
  { canonical: '총원가', aliases: ['총원가', '총공사원가', '총용역원가'] },
  { canonical: '총계', aliases: ['총계', '총합계', '합계'] },
];

// 긴 별칭 우선 (부분포함 오탐 최소화 — ALIAS_INDEX와 동일 원칙)
const STRUCTURAL_INDEX: Array<{ alias: string; def: StructuralDef }> = STRUCTURAL_ITEMS.flatMap((def) =>
  def.aliases.map((alias) => ({ alias, def })),
).sort((a, b) => b.alias.length - a.alias.length);

/** 트리 골격 라벨 매칭. CANONICAL 매칭 실패 시 폴백으로 사용. */
export function matchStructuralLabel(label: string): string | null {
  return lookupAlias(stripLabelDecorations(label), STRUCTURAL_INDEX)?.canonical ?? null;
}
