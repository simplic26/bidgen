# 세 문서 정규화 검수실 설계

> 작성일: 2026-07-22
> 대상 앱: 원가계산서 자동검증 React/Vite MVP
> 대표 제비율표: `public/_ref_rate.xlsx`의 저장소 이력본
> 대표 노임단가표: `public/_ref_labor.xlsx`의 저장소 이력본

## 1. 목적

원가계산서, 제비율표, 노임단가표를 업로드한 직후 세 문서의 정규화 결과를 검증 실행 전에 사용자가 검수할 수 있는 `정규화 검수실`을 추가한다. 원가계산서는 기존 정규화 흐름을 확장하고, 노임단가표는 모든 직종 행을 보존하며, 제비율표는 병합 헤더와 다중 조건을 구조적으로 해석해 전체 적용 규칙을 행 단위로 펼친다.

이 기능의 핵심은 파서가 맞았다고 단정하는 것이 아니라 다음 사실을 사용자가 검증할 수 있게 만드는 것이다.

- 어떤 원본 셀과 병합범위에서 각 정규화 필드가 생성됐는가
- 어떤 조건 조합에 어떤 요율·금액·계산식이 연결됐는가
- 원본 중 무엇이 규칙, 주석, 규정 또는 장식으로 분류됐는가
- 어떤 셀이 미분류, 충돌, 낮은 신뢰도 또는 원본 오류 상태인가
- 현재 원가계산서에서 추출한 공사조건과 일치하는 규칙은 무엇인가

## 2. 범위

### 포함

- 업로드와 검증 실행 사이의 정규화 검수 단계
- 원가계산서·제비율표·노임단가표 3개 탭
- 원가계산서의 모든 시트에서 공사종류·금액·기간 후보 추출
- 제비율표 전체 규칙 정규화와 관련규정 연결
- 노임단가표의 유효 단가와 미공표 행 정규화
- 원본 셀·병합범위·표시값·숫자형식 근거
- 검색, 필터, 상태 요약, 행 상세 근거 패널
- 치명적 오류와 진행 가능한 경고의 분리
- 대표 파일 기반 단위·통합·브라우저 회귀 테스트

### 제외

- 정규화 결과 직접 수정
- 사용자 보정값 저장 또는 영구 보관
- AI/LLM 기반 숫자 판정
- 임의 기관의 모든 제비율표 양식 지원
- 원본 Excel 파일 수정

향후 직접 수정 기능은 원본 정규화 결과를 변경하지 않고 별도 `NormalizationOverride` 레이어로 추가한다.

## 3. 사용자 흐름

```text
파일 업로드
  → 세 문서 파싱 및 정규화
  → 정규화 검수실
      ├─ 자동 추출 공사조건 확인
      ├─ 원가계산서 탭 검수
      ├─ 제비율표 탭 전체 규칙 검수
      └─ 노임단가표 탭 검수
  → 치명적 오류가 없으면 검증 실행
  → 진행 화면
  → 검증 결과 대시보드
```

현재 `RecognitionScreen`이 담당하는 문서 인식 단계는 `NormalizationAuditScreen`으로 대체한다. 검증 모드 선택은 검수실 하단의 검증 설정 영역에 유지한다. 파싱된 정규화 값은 읽기 전용이며 임의 요율 수정 UI는 이번 범위에서 제공하지 않는다.

## 4. 아키텍처

### 4.1 처리 경계

1. `workbookToIR`은 세 문서의 원본 구조를 동일한 `WorkbookIR`로 보존한다.
2. 문서별 정규화기는 `WorkbookIR`을 도메인 행과 파싱 이슈로 변환한다.
3. `extractProjectContext`는 원가계산서 전체 시트에서 적용조건 후보를 만든다.
4. `matchApplicableRateRules`는 공사조건과 전체 제비율 규칙을 비교한다.
5. 화면은 원본 IR, 정규화 행, 커버리지, 이슈를 함께 표시한다.
6. 검증 엔진은 화면에서 임의로 재해석하지 않고 선택된 정규화 규칙만 사용한다.

### 4.2 공통 결과 계약

```ts
type NormalizationKind = 'COST_STATEMENT' | 'RATE_TABLE' | 'LABOR_TABLE';
type NormalizationSeverity = 'INFO' | 'WARNING' | 'FATAL';

interface CellEvidence {
  sheetName: string;
  cell: string;
  mergedRange: string | null;
  rawValue: string;
  displayValue: string;
  numberFormat: string | null;
}

interface NormalizationIssue {
  issueId: string;
  code:
    | 'UNCLASSIFIED_SOURCE_CELL'
    | 'LOW_CONFIDENCE'
    | 'DUPLICATE_RULE'
    | 'CONDITION_CONFLICT'
    | 'SOURCE_FORMULA_ERROR'
    | 'MISSING_REQUIRED_VALUE'
    | 'NO_NORMALIZED_ROWS';
  severity: NormalizationSeverity;
  message: string;
  evidence: CellEvidence[];
}

interface NormalizationCoverage {
  nonEmptySourceCells: number;
  classifiedSourceCells: number;
  unclassifiedSourceCells: number;
  normalizedRows: number;
  warningCount: number;
  fatalCount: number;
}

interface NormalizedDocument<Row> {
  kind: NormalizationKind;
  file: FileMeta;
  ir: WorkbookIR;
  rows: Row[];
  issues: NormalizationIssue[];
  coverage: NormalizationCoverage;
}
```

커버리지는 `classifiedSourceCells / nonEmptySourceCells`를 기준으로 계산한다. 빈 셀은 분모에서 제외한다. 장식 셀도 `DECORATION`으로 명시 분류해야 하며, 분류되지 않은 비어 있지 않은 셀은 숨기지 않는다.

## 5. 원가계산서 정규화와 공사조건 추출

### 5.1 기존 정규화 확장

기존 `normalizeCostStatement(ir)`가 만드는 `NormalizedCostRow`를 유지한다. 각 행에는 계산 셀 외에도 항목 라벨, 요율 셀, 산출기초 참조 셀의 `CellEvidence[]`를 추가한다. 결과 검증 전에도 행이 표시되어야 하므로 `results.length`에 의존하지 않는다.

### 5.2 공사조건 후보

```ts
type ProjectAmountBasis =
  | 'ESTIMATED_PRICE'
  | 'DIRECT_CONSTRUCTION_COST'
  | 'TOTAL_CONSTRUCTION_COST'
  | 'SUPPLY_AMOUNT'
  | 'SAFETY_MANAGEMENT_BASE';

interface ContextCandidate<T> {
  value: T;
  confidence: number;
  evidence: CellEvidence[];
  reason: string;
}

interface ProjectContext {
  constructionType: ContextCandidate<string> | null;
  amount: ContextCandidate<{ basis: ProjectAmountBasis; value: number }> | null;
  durationDays: ContextCandidate<number> | null;
  bidDate: ContextCandidate<string> | null;
  alternatives: Array<ContextCandidate<string | number>>;
  issues: NormalizationIssue[];
}
```

추출기는 특정 시트에 한정하지 않고 모든 시트를 검색한다.

- 금액 별칭: `추정가격`, `도급액`, `공사금액`, `총공사금액`, `직접공사비`, `공급가액`, `산업안전보건관리비 대상액`
- 기간 별칭: `공사기간`, `공기`, `착공일`, `준공일`, `계약기간`
- 공사종류 별칭: `건축`, `토목`, `중건설`, `특수건설`, `산업설비`, `전기`, `통신`, `소방`, `전문`

라벨과 같은 행 또는 인접 셀의 값을 후보로 수집하고, 수식 셀이면 참조 경로를 따라 실제 값과 근거 셀을 보존한다. 착공일과 준공일이 있으면 양 끝 날짜를 포함한 달력 일수로 기간을 계산한다. 후보가 여러 개면 라벨 정확도, 표지·요약 시트 우선순위, 수식 연결성, 값 유효성을 점수화한다. 최고 후보끼리 충돌하면 자동 확정하지 않고 사용자 입력란을 표시한다.

## 6. 제비율표 정규화

### 6.1 대표 파일 구조

대표 파일의 주 시트는 `건축제비율(1.1)`이며 121행, Workbook IR 기준 164열, 병합범위 378개다. `관련규정 링크` 시트는 항목, 소관부처, 링크를 제공한다.

파서는 상단의 대분류 병합 헤더와 각 섹션의 조건 축을 먼저 찾고, 숫자 셀에서 거꾸로 조건 헤더를 추적한다. 단순히 첫 숫자 셀을 선택하지 않는다.

### 6.2 규칙 데이터 계약

```ts
type RateValueKind =
  | 'PERCENT_RATE'
  | 'FIXED_AMOUNT'
  | 'RATE_PLUS_BASE_AMOUNT'
  | 'FORMULA'
  | 'NOT_APPLICABLE'
  | 'SOURCE_ERROR';

type AmountBasis =
  | 'DIRECT_LABOR'
  | 'LABOR'
  | 'MATERIAL_PLUS_LABOR'
  | 'MATERIAL_PLUS_DIRECT_LABOR'
  | 'MATERIAL_PLUS_DIRECT_LABOR_PLUS_OWNER_SUPPLIED'
  | 'DIRECT_CONSTRUCTION_COST'
  | 'ESTIMATED_PRICE'
  | 'TOTAL_CONSTRUCTION_COST'
  | 'SAFETY_MANAGEMENT_BASE'
  | 'OTHER';

interface NumericRange {
  min: number | null;
  minInclusive: boolean;
  max: number | null;
  maxInclusive: boolean;
  unit: 'KRW' | 'DAY' | 'YEAR';
  sourceText: string;
}

interface RateConditions {
  constructionTypes: string[];
  amountBasis: ProjectAmountBasis | null;
  amountRange: NumericRange | null;
  durationRange: NumericRange | null;
  procurementGrade: string | null;
  contractorClass: 'GENERAL' | 'SPECIALTY' | null;
  bidMethods: string[];
  includesOwnerSuppliedMaterial: boolean | null;
  requiredConditions: string[];
  exclusions: string[];
}

interface RegulationReference {
  canonicalName: string;
  ministry: string;
  title: string;
  url: string;
  evidence: CellEvidence[];
}

interface NormalizedRateRule {
  ruleId: string;
  canonicalName: string;
  originalSectionTitle: string;
  effectiveFrom: string | null;
  basis: AmountBasis;
  valueKind: RateValueKind;
  rate: number | null;
  fixedAmount: number | null;
  formula: string | null;
  roundingRule: string | null;
  conditions: RateConditions;
  notes: string[];
  regulations: RegulationReference[];
  confidence: number;
  status: 'NORMAL' | 'NEEDS_REVIEW' | 'SOURCE_ERROR';
  evidence: CellEvidence[];
}
```

### 6.3 구조 기반 파싱 알고리즘

1. 모든 병합범위를 인덱스화하고 각 좌표가 어떤 병합 헤더에 포함되는지 조회할 수 있게 한다.
2. `[간접노무비]`, `[기타경비]`, `[일반관리비]` 같은 대분류 병합 헤더의 열 범위를 섹션으로 분할한다.
3. 각 값 셀에서 위쪽과 왼쪽의 병합 헤더를 추적해 공사종류, 금액구간, 기간구간, 업종, 등급, 계약방식 조건을 수집한다.
4. `10억 미만`, `50억 이상 - 100억 미만`, `36개월 초과`를 경계 포함 여부가 명확한 `NumericRange`로 변환한다.
5. `%` 숫자형식, 일반 숫자, 텍스트식 요율을 구분하고 내부 요율은 소수로 통일한다.
6. `요율 + 기초액`, 초과금액식, 공기(년) 곱셈식은 원문 계산식과 파싱된 숫자를 함께 보존한다.
7. `-`, 빈 값, `#REF!`를 삭제하지 않고 각각 적용 제외, 누락 또는 원본 오류로 분류한다.
8. 섹션 인접 주석에서 적용조건, 제외조건, 고시번호, 반올림 규칙을 연결한다.
9. `관련규정 링크` 시트를 표준항목명으로 매핑해 규칙에 연결한다.
10. 동일한 조건 집합과 항목에 규칙이 중복되면 자동으로 하나를 버리지 않고 `DUPLICATE_RULE`을 생성한다.

### 6.4 반드시 보존할 조건 축

- 공사규모와 공사기간
- 건축·산업설비·토목·중건설·특수건설 등 공사종류
- 종합건설업·전문건설업
- 조달청 1~7등급과 등급 미만
- 추정가격·직접공사비·총공사금액·안전관리비 대상액 등 금액 기준
- 도급자관급 포함·미포함
- 종심제·종평제·턴키·대안공사
- 법정 적용대상과 적용 제외 공사
- 적용시기, 고시번호, 관련규정, 반올림 규칙

### 6.5 적용 규칙 선택

`matchApplicableRateRules(context, rules)`는 다음 순서로 후보를 좁힌다.

1. 기준일이 `effectiveFrom` 이후인지 확인
2. 공사종류 일치
3. 규칙이 요구하는 금액 기준의 값과 구간 일치
4. 공사기간 구간 일치
5. 업종·등급·계약방식·도급자관급 조건 일치
6. 제외조건에 해당하지 않는지 확인

항목별 최종 후보가 정확히 1개면 적용한다. 2개 이상이면 `CONDITION_CONFLICT`, 0개면 검증결과를 `UNAVAILABLE`로 만든다. 낮은 신뢰도 규칙은 오류 확정에 사용하지 않고 확인 필요로 처리한다.

## 7. 노임단가표 정규화

```ts
interface NormalizedLaborRate {
  rowId: string;
  occupationCode: string;
  occupationName: string;
  unitPrice: number | null;
  publicationStatus: 'PUBLISHED' | 'NOT_PUBLISHED' | 'INVALID';
  markers: string[];
  notes: string[];
  confidence: number;
  evidence: CellEvidence[];
}
```

대표 파일의 `직종번호·직종명·노임단가` 헤더를 구조적으로 찾는다. 행 번호를 고정하지 않는다. 직종번호의 `*`, `**` 표식은 제거하지 않고 `markers`에 별도로 보존한다. 단가가 `-`인 행도 삭제하지 않고 `NOT_PUBLISHED`로 표시한다. 동일 직종명 또는 코드가 중복되면 확인 필요 이슈를 만든다.

## 8. 정규화 검수실 UI

### 8.1 공통 구성

- 상단 단계 표시: `파일 업로드 → 정규화 검수 → 검증 실행`
- 공사조건 요약 카드: 값, 신뢰도, 원본 셀, 대안 후보
- 탭: `원가계산서`, `제비율표`, `노임단가표`
- 탭 배지: 전체 행, 정상, 확인 필요, 미분류, 원본 오류
- 공통 도구: 검색, 상태 필터, 원본 시트 필터, 문제 행만 보기
- 표 행 선택 시 우측 근거 패널 표시
- 하단 고정 영역: 이전, 검증 모드, 검증 실행

### 8.2 제비율표 전용 필터

- 표준항목
- 공사종류
- 금액 기준과 금액구간
- 기간구간
- 값 유형
- 업종·등급·계약방식
- 적용 가능·조건 충돌·적용 규칙 없음

### 8.3 근거 패널

근거 패널은 정규화 필드별로 원본 시트, 값 셀, 병합 헤더 범위, 원본값, 숫자형식, 연결 규정, 파싱 이유를 보여준다. 긴 주석은 요약과 원문을 분리한다. 모든 상태는 색상과 텍스트를 함께 사용한다.

## 9. 오류 및 진행 정책

검증 실행을 차단하는 치명적 오류:

- 필수 파일 읽기 실패
- 원가계산서·제비율표·노임단가표 중 하나의 정규화 행이 0개
- 정규화 데이터 계약을 만족하지 못해 화면 또는 검증 엔진이 소비할 수 없음

경고 후 진행 가능한 오류:

- 미분류 원본 셀
- 낮은 신뢰도
- 중복 규칙
- 조건 충돌
- 일부 원본 수식 오류
- 미공표 노임단가
- 일부 공사조건 자동 추출 실패

공사조건을 찾지 못했거나 후보가 충돌할 때만 해당 입력란을 노출한다. 사용자가 입력한 값은 원본 추출값과 별도로 표시하고 세션 동안만 사용한다.

## 10. 향후 직접 수정 확장

이번 UI는 읽기 전용이다. 향후 기능은 다음 계약을 추가해 구현한다.

```ts
interface NormalizationOverride {
  targetRowId: string;
  fieldPath: string;
  originalValue: unknown;
  correctedValue: unknown;
  reason: string;
  createdAt: string;
}
```

검증 엔진은 `원본 정규화값 → 사용자 보정값` 순서로 적용하되 두 값을 모두 감사 가능하게 보존한다. 이번 구현에서는 타입 확장 지점만 고려하고 보정 UI나 상태는 만들지 않는다.

## 11. 성능과 React 구현 원칙

- 세 파일의 독립 파싱은 `Promise.all`로 병렬 실행한다.
- 정규화는 업로드 이벤트에서 한 번 수행하고 렌더 중 재계산하지 않는다.
- 검색·필터 결과는 원시 행 배열과 원시 필터 값에 대해 `useMemo`로 계산한다.
- 500행 안팎의 대표 표는 추가 가상화 의존성 없이 렌더한다.
- 상세 근거 패널은 선택된 행만 계산한다.
- SheetJS는 기존 직접 import를 유지하고 새 대형 UI 의존성을 추가하지 않는다.

## 12. 테스트 전략

### 12.1 파서 단위 테스트

- 병합 헤더가 병합범위 밖으로 전파되지 않는다.
- `AC15`는 간접노무비 17.5%와 `10억 미만`, `6개월 이하`, `건축` 조건을 갖는다.
- `ER19`는 산업안전보건관리비 3.11%와 해당 공사종류·금액구간만 가지며 인접 조건이 섞이지 않는다.
- `P98`, `P101`, `P104`, `P107`, `P110`의 공사이행보증수수료 계산식을 원문 그대로 보존한다.
- 고정 기초액과 요율이 함께 있는 산업안전보건관리비 규칙을 `RATE_PLUS_BASE_AMOUNT`로 만든다.
- `AH81`의 `#REF!`를 `SOURCE_ERROR`와 `SOURCE_FORMULA_ERROR`로 노출한다.
- 관련규정 링크의 소관부처·URL을 표준항목에 연결한다.
- 비어 있지 않은 미분류 셀을 커버리지에서 숨기지 않는다.

### 12.2 노임단가 테스트

- 작업반장 `C2 = 215,907`을 보존한다.
- `*1014`는 코드 `1014`와 표식 `*`를 모두 보존한다.
- 단가가 `-`인 직종은 삭제하지 않고 `NOT_PUBLISHED`로 만든다.
- 유효 숫자 단가와 원본 셀 주소가 일치한다.

### 12.3 공사조건 테스트

- 모든 시트를 검색한다.
- 표지·요약 시트의 명확한 라벨을 우선한다.
- 착공일·준공일로 기간을 계산한다.
- 충돌하는 금액 후보를 자동 확정하지 않는다.
- 수식 참조의 원본 경로를 근거에 포함한다.

### 12.4 통합·브라우저 테스트

- 세 문서 업로드 후 정규화 검수실로 이동한다.
- 세 탭의 행 수와 상태 배지가 표시된다.
- 제비율 필터와 검색이 함께 동작한다.
- 행 선택 시 올바른 셀과 병합범위가 근거 패널에 표시된다.
- 치명적 오류가 있을 때만 검증 실행 버튼이 비활성화된다.
- 경고만 있으면 경고 문구와 함께 검증 실행이 가능하다.
- 데스크톱과 모바일에서 탭, 표 가로 스크롤, 근거 패널을 사용할 수 있다.

## 13. 수용 기준

- 세 문서 정규화 결과를 검증 전에 탭으로 볼 수 있다.
- 제비율표의 전체 조건 조합이 독립 규칙 행으로 유지된다.
- 제비율 조건이 인접 섹션이나 다른 공사종류와 섞이지 않는다.
- 요율, 고정금액, 복합식, 적용 제외, 원본 오류를 구분한다.
- 공사조건을 원가계산서 모든 시트에서 자동 추출하고 근거를 표시한다.
- 자동 추출에 실패하거나 충돌한 조건만 사용자에게 입력받는다.
- 관련규정 시트가 항목별 근거로 연결된다.
- 노임단가의 미공표 행과 코드 표식을 보존한다.
- 모든 정규화 필드에서 원본 셀 또는 병합범위를 역추적할 수 있다.
- 미분류 셀, 중복, 충돌, 원본 오류가 커버리지와 이슈 목록에 나타난다.
- 파일 읽기 실패나 행 0개만 검증을 차단하고 나머지는 경고 후 진행할 수 있다.
- 대표 파일의 핵심 셀 회귀 테스트와 전체 빌드가 통과한다.
