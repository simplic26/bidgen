// 9개 레퍼런스 파일을 파싱해 qa/golden/*.json 정답지와 대조. 사용: node scripts/golden-test.mjs [--draft]
// --draft: 파서 출력을 qa/golden/에 초안 JSON으로 저장 (confirmed:false — 사람 검수 전용, 채점에는 confirmed만 사용해도 됨)
// 골든은 sourceSha256(레퍼런스 파일 내용 해시)로 원본을 찾는다 — 레퍼런스는 사내 자료라 파일명조차
// 저장소(공개)에 남기면 안 되기 때문. 실명 파일명은 콘솔 로그에만 나온다.
import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { createServer } from 'vite';
import * as XLSX from 'xlsx';

const REF_DIR = resolve('..', '레퍼런스 엄선');
const GOLDEN_DIR = resolve('qa', 'golden');
const draft = process.argv.includes('--draft');

// 레퍼런스 엑셀은 저장소에 커밋되지 않는다(사내 자료). 없는 환경에서는 ENOENT로 죽지 말고
// SKIP으로 알리고 성공 종료한다 — 이 하네스는 레퍼런스가 있는 환경에서만 의미가 있다.
if (!existsSync(REF_DIR)) {
  console.log('SKIP: 레퍼런스 폴더 없음 (' + REF_DIR + ')');
  process.exit(0);
}

// 레퍼런스 폴더의 엑셀을 내용 해시로 색인 — 골든의 sourceSha256 → 실제 경로
const refByHash = new Map();
for (const f of readdirSync(REF_DIR).filter((n) => /\.xlsx?$/i.test(n))) {
  const p = join(REF_DIR, f);
  refByHash.set(createHash('sha256').update(readFileSync(p)).digest('hex'), p);
}

const server = await createServer({ configFile: 'vite.config.mjs', configLoader: 'native', server: { middlewareMode: true } });
try {
  const excel = await server.ssrLoadModule('/src/utils/excel.ts');
  const ct = await server.ssrLoadModule('/src/utils/costTree.ts');

  const parseFile = (path) => {
    const wb = XLSX.read(readFileSync(path), { cellNF: true, cellStyles: true, cellFormula: true, cellDates: true });
    const ir = excel.workbookToIR(wb, { fileName: path, procurementType: 'CONSTRUCTION' });
    return ct.parseCostTree(ir);
  };

  if (draft) {
    mkdirSync(GOLDEN_DIR, { recursive: true });
    for (const [hash, p] of refByHash) {
      const tree = parseFile(p);
      const items = ct.costTreeToDetectedItems(tree).map((i) => ({
        canonical: i.canonicalName,
        amount: i.amountValue,
        ...(i.rateValue != null ? { rate: i.rateValue } : {}),
      }));
      // 초안 파일명에도 실명을 남기지 않는다 — 해시 앞 12자리로 이름 짓고, 실명은 콘솔에만.
      const goldenName = `draft-${hash.slice(0, 12)}.json`;
      writeFileSync(join(GOLDEN_DIR, goldenName), JSON.stringify({
        sourceSha256: hash, summarySheet: tree.summarySheet, items, confirmed: false, note: '',
      }, null, 2) + '\n');
      console.log(`[draft] ${goldenName} (${p}): ${items.length} items, sheet=${tree.summarySheet}, issues=[${tree.issues.join(' / ')}]`);
    }
  } else {
    let totalExpected = 0, totalFound = 0, totalFalse = 0;
    const failures = [];
    const skipped = [];
    const scored = [];
    // 채점된 confirmed 골든 개수 하한. 골든이 조용히 confirmed:false로 뒤집히거나 파일이 사라져
    // "재현율 0/0 통과" 같은 무의미한 성공이 나오는 것을 막는다.
    const MIN_SCORED_GOLDENS = 8;
    for (const g of readdirSync(GOLDEN_DIR).filter((n) => n.endsWith('.json'))) {
      const golden = JSON.parse(readFileSync(join(GOLDEN_DIR, g), 'utf8'));
      // confirmed:false 골든은 채점 대상이 아니다(예: 견적비교표처럼 검증 대상 자체가 아닌 파일) — SKIPPED로 로그만 남기고
      // 재현율/오탐 집계에서 제외한다. plan의 수용 기준 "confirmed 골든에서 재현율 100%"와 하네스를 일치시킨다.
      if (golden.confirmed === false) {
        skipped.push(g);
        continue;
      }
      scored.push(g);
      const srcPath = refByHash.get(golden.sourceSha256);
      if (!srcPath) {
        failures.push(`${g} :: sourceSha256에 해당하는 레퍼런스 파일 없음 (${String(golden.sourceSha256).slice(0, 12)}…)`);
        continue;
      }
      const tree = parseFile(srcPath);
      const detected = ct.costTreeToDetectedItems(tree);
      // canonicalName이 같은 항목이 여러 개 나올 수 있다(예: ref-09 파일의 D/E 이중 열 — 같은 canonical이 부모/자식
      // 셀 2건으로 정당하게 존재). Map 1개로는 중복을 못 다루므로 "기대 항목마다 미소비 실제 항목 하나를 소비"한다.
      //
      // 순서 의존을 없애기 위해 2패스로 나눈다. first-fit 한 번만 돌리면, rate 없는 기대 항목이 먼저 나와
      // rate 있는 실제 항목을 먹어버리고 뒤에 오는 rate 지정 기대 항목이 미탐지로 떨어질 수 있다(골든 항목
      // 나열 순서가 결과를 바꿈).
      //   1패스: rate를 지정한 기대 항목 — canonical+amount+rate 전부 일치해야 소비.
      //   2패스: `"rate": null`을 **명시**한 기대 항목 — "이 항목에는 셀 근거가 되는 요율이 없다"는
      //          적극적 주장이다. canonical+amount가 맞고 rateValue가 null인 실제 항목만 소비한다.
      //          파서가 요율을 만들어내면(요율 없는 행에 값-추론으로 엉뚱한 간선을 붙이면) FAIL이어야 한다.
      //   3패스: rate 키가 아예 없는 기대 항목 — 요율 미검증(주장 없음). canonical+amount만 일치하면 되고,
      //          rate 없는 실제 항목을 우선 소비하되 남은 게 없으면 rate 있는 실제 항목도 쓴다.
      // 2패스를 3패스보다 먼저 두는 이유: 같은 canonical+amount에서 "명시적 null"이 rate 없는 실제 항목을
      // 먼저 차지해야 한다(주장 없는 쪽이 먹어버리면 명시 주장이 억울하게 FAIL 난다).
      const hasExplicitNullRate = (exp) =>
        Object.prototype.hasOwnProperty.call(exp, 'rate') && exp.rate === null;
      const entries = detected.map((item, idx) => ({ item, idx }));
      const consumedIdx = new Set();
      const withRate = golden.items.filter((exp) => exp.rate != null);
      const withNullRate = golden.items.filter(hasExplicitNullRate);
      const withoutRate = golden.items.filter((exp) => exp.rate == null && !hasExplicitNullRate(exp));
      totalExpected += golden.items.length;

      const consume = (exp, pick) => {
        const hit = pick(entries.filter(({ idx }) => !consumedIdx.has(idx)));
        if (hit) {
          consumedIdx.add(hit.idx);
          totalFound += 1;
          return true;
        }
        const near = detected.find((d) => d.canonicalName === exp.canonical);
        const wantRate = hasExplicitNullRate(exp) ? '요율없음(명시)' : (exp.rate ?? '-');
        failures.push(
          `${g} :: ${exp.canonical} 기대 ${exp.amount}/${wantRate} ↔ 실제 ${near?.amountValue ?? '미탐지'}/${near?.rateValue ?? '-'}`,
        );
        return false;
      };

      for (const exp of withRate) {
        consume(exp, (free) => free.find(({ item }) =>
          item.canonicalName === exp.canonical &&
          item.amountValue === exp.amount &&
          item.rateValue != null && Math.abs(item.rateValue - exp.rate) < 1e-6));
      }
      for (const exp of withNullRate) {
        consume(exp, (free) => free.find(({ item }) =>
          item.canonicalName === exp.canonical &&
          item.amountValue === exp.amount &&
          item.rateValue == null));
      }
      for (const exp of withoutRate) {
        consume(exp, (free) => {
          const same = free.filter(({ item }) =>
            item.canonicalName === exp.canonical && item.amountValue === exp.amount);
          return same.find(({ item }) => item.rateValue == null) ?? same[0];
        });
      }
      for (const { item, idx } of entries) {
        if (consumedIdx.has(idx)) continue;
        totalFalse += 1;
        failures.push(`${g} :: 오탐 ${item.canonicalName}=${item.amountValue} (${item.amountCell})`);
      }
      if (golden.summarySheet && tree.summarySheet !== golden.summarySheet) {
        failures.push(`${g} :: 요약시트 기대 ${golden.summarySheet} ↔ 실제 ${tree.summarySheet}`);
      }
    }
    if (scored.length < MIN_SCORED_GOLDENS) {
      failures.push(`채점된 confirmed 골든 ${scored.length}개 — 최소 ${MIN_SCORED_GOLDENS}개여야 함 (채점: ${scored.join(', ') || '없음'})`);
    }
    console.log(`재현율 ${totalFound}/${totalExpected}, 오탐 ${totalFalse}건, 채점 파일 ${scored.length}개`);
    if (skipped.length) console.log(`SKIPPED (confirmed:false, 채점 제외): ${skipped.join(', ')}`);
    for (const f of failures) console.log('  FAIL ' + f);
    process.exitCode = failures.length ? 1 : 0;
  }
} finally {
  await server.close();
}
