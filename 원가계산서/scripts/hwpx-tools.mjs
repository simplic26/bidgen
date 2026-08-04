// 공고문 hwpx 템플릿 도구
//   node scripts/hwpx-tools.mjs make-skeleton   개발용 스켈레톤 템플릿 생성 → public/templates/notice-template.hwpx
//   node scripts/hwpx-tools.mjs verify          템플릿의 자리표시자 무결성 검사 (누락 = 미삽입 or run 분할)
//   node scripts/hwpx-tools.mjs fill-test       더미 값 치환 후 잔여 마커·zip 무결성 검사
//
// 자리표시자 키는 src/utils/hwpx.ts의 NOTICE_PLACEHOLDER_KEYS와 동기화 유지할 것.
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const TEMPLATE_PATH = resolve(ROOT, 'public/templates/notice-template.hwpx');

const PLACEHOLDER_KEYS = [
  '사업소',
  '계약종류',
  '세부계약종류',
  '발주년월',
  '입찰방법',
  '입찰종류',
  '협정대상여부',
  '입찰공고번호',
  '공고일',
  '공사명',
  '발주기관',
  '발주유형',
  '공고명의',
  '추정가격',
  '사급자재비',
  '추정금액',
  '예비가격기초금액',
  'A값',
  '순공사원가',
  '공사기간',
  '공사내용',
  '공사구분',
  '참가자격',
  '공동계약',
  '입찰개시일시',
  '입찰마감일시',
  '개찰일시',
  '개찰장소',
  '보증금납부기한',
  '국민건강보험료',
  '국민연금보험료',
  '노인장기요양보험료',
  '퇴직공제부금비',
  '산업안전보건관리비',
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
];

// ---------------------------------------------------------------------------
// 스켈레톤 hwpx (최소 OWPML) — 실제 공고문 템플릿이 준비되기 전 개발용 대체물.
// 최종 인수는 한글에서 변환한 실제 템플릿으로 해야 한다.
// ---------------------------------------------------------------------------

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const VERSION_XML =
  XML_DECL +
  '<hv:HCFVersion xmlns:hv="http://www.hancom.co.kr/hwpml/2011/version" tagetApplication="WORDPROCESSOR" major="5" minor="1" micro="1" buildNumber="0" os="1" xmlVersion="1.4" application="Hancom Office Hangul" appVersion="9, 1, 1, 5656"/>';

const CONTAINER_XML =
  XML_DECL +
  '<ocf:container xmlns:ocf="urn:oasis:names:tc:opendocument:xmlns:container">' +
  '<ocf:rootfiles>' +
  '<ocf:rootfile full-path="Contents/content.hpf" media-type="application/hwpml-package+xml"/>' +
  '</ocf:rootfiles>' +
  '</ocf:container>';

const MANIFEST_XML =
  XML_DECL +
  '<odf:manifest xmlns:odf="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">' +
  '<odf:file-entry odf:full-path="/" odf:media-type="application/hwp+zip"/>' +
  '<odf:file-entry odf:full-path="Contents/header.xml" odf:media-type="application/xml"/>' +
  '<odf:file-entry odf:full-path="Contents/section0.xml" odf:media-type="application/xml"/>' +
  '</odf:manifest>';

const CONTENT_HPF =
  XML_DECL +
  '<opf:package xmlns:opf="http://www.idpf.org/2007/opf/" version="" unique-identifier="" id="">' +
  '<opf:metadata><opf:title>입찰공고문</opf:title><opf:language>ko</opf:language></opf:metadata>' +
  '<opf:manifest>' +
  '<opf:item id="header" href="Contents/header.xml" media-type="application/xml"/>' +
  '<opf:item id="section0" href="Contents/section0.xml" media-type="application/xml"/>' +
  '<opf:item id="settings" href="settings.xml" media-type="application/xml"/>' +
  '</opf:manifest>' +
  '<opf:spine><opf:itemref idref="header" linear="yes"/><opf:itemref idref="section0" linear="yes"/></opf:spine>' +
  '</opf:package>';

const SETTINGS_XML =
  XML_DECL +
  '<ha:HWPApplicationSetting xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0">' +
  '<ha:CaretPosition listIDRef="0" paraIDRef="0" pos="0"/>' +
  '</ha:HWPApplicationSetting>';

const FONT_LANGS = ['HANGUL', 'LATIN', 'HANJA', 'JAPANESE', 'OTHER', 'SYMBOL', 'USER'];

const HEADER_XML =
  XML_DECL +
  '<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" version="1.4" secCnt="1">' +
  '<hh:beginNum page="1" footnote="1" endnote="1" pic="1" tbl="1" equation="1"/>' +
  '<hh:refList>' +
  `<hh:fontfaces itemCnt="${FONT_LANGS.length}">` +
  FONT_LANGS.map(
    (lang) =>
      `<hh:fontface lang="${lang}" fontCnt="1"><hh:font id="0" face="함초롬바탕" type="TTF" isEmbedded="0"/></hh:fontface>`,
  ).join('') +
  '</hh:fontfaces>' +
  '<hh:borderFills itemCnt="1">' +
  '<hh:borderFill id="1" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">' +
  '<hh:slash type="NONE" Crooked="0" isCounter="0"/><hh:backSlash type="NONE" Crooked="0" isCounter="0"/>' +
  '<hh:leftBorder type="NONE" width="0.1 mm" color="#000000"/><hh:rightBorder type="NONE" width="0.1 mm" color="#000000"/>' +
  '<hh:topBorder type="NONE" width="0.1 mm" color="#000000"/><hh:bottomBorder type="NONE" width="0.1 mm" color="#000000"/>' +
  '<hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/>' +
  '</hh:borderFill>' +
  '</hh:borderFills>' +
  '<hh:charProperties itemCnt="1">' +
  '<hh:charPr id="0" height="1000" textColor="#000000" shadeColor="none" useFontSpace="0" useKerning="0" symMark="NONE" borderFillIDRef="1">' +
  '<hh:fontRef hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>' +
  '<hh:ratio hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>' +
  '<hh:spacing hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>' +
  '<hh:relSz hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>' +
  '<hh:offset hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>' +
  '</hh:charPr>' +
  '</hh:charProperties>' +
  '<hh:paraProperties itemCnt="1">' +
  '<hh:paraPr id="0" tabPrIDRef="0" condense="0" fontLineHeight="0" snapToGrid="1" suppressLineNumbers="0" checked="0">' +
  '<hh:align horizontal="JUSTIFY" vertical="BASELINE"/>' +
  '<hh:heading type="NONE" idRef="0" level="0"/>' +
  '<hh:breakSetting breakLatinWord="KEEP_WORD" breakNonLatinWord="BREAK_WORD" widowOrphan="0" keepWithNext="0" keepLines="0" pageBreakBefore="0" lineWrap="BREAK"/>' +
  '<hh:margin><hh:intent value="0" unit="HWPUNIT"/><hh:left value="0" unit="HWPUNIT"/><hh:right value="0" unit="HWPUNIT"/><hh:prev value="0" unit="HWPUNIT"/><hh:next value="0" unit="HWPUNIT"/></hh:margin>' +
  '<hh:lineSpacing type="PERCENT" value="160" unit="HWPUNIT"/>' +
  '<hh:border borderFillIDRef="1" offsetLeft="0" offsetRight="0" offsetTop="0" offsetBottom="0" connect="0" ignoreMargin="0"/>' +
  '<hh:autoSpacing eAsianEng="0" eAsianNum="0"/>' +
  '</hh:paraPr>' +
  '</hh:paraProperties>' +
  '<hh:styles itemCnt="1">' +
  '<hh:style id="0" type="PARA" name="바탕글" engName="Normal" paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="0" langID="1042" lockForm="0"/>' +
  '</hh:styles>' +
  '</hh:refList>' +
  '</hh:head>';

function esc(text) {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function para(text, id) {
  return (
    `<hp:p id="${id}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">` +
    `<hp:run charPrIDRef="0"><hp:t>${esc(text)}</hp:t></hp:run>` +
    '</hp:p>'
  );
}

function buildSectionXml() {
  const lines = [
    '입 찰 공 고',
    '',
    `입찰공고번호 : {{입찰공고번호}}    공고일 : {{공고일}}`,
    '',
    '1. 입찰에 부치는 사항',
    '  가. 공사명 : {{공사명}}',
    '  나. 발주기관 : {{발주기관}}',
    '  다. 발주유형 : {{발주유형}}',
    '  라. 추정가격 : {{추정가격}} (부가가치세 별도)',
    '  마. 예비가격기초금액 : {{예비가격기초금액}} (부가가치세 포함)',
    '  바. 공사기간 : {{공사기간}}',
    '  사. 공사내용 : {{공사내용}}',
    '',
    '2. 입찰 일정',
    '  가. 입찰서 제출마감 : {{입찰마감일시}}',
    '  나. 개찰일시 : {{개찰일시}}',
    '',
    '3. 문의처',
    '  담당자 : {{담당자}}  연락처 : {{연락처}}',
    '',
    '위와 같이 공고합니다.',
  ];
  // 첫 문단에 최소 secPr(용지 설정) 포함
  const firstRun =
    '<hp:run charPrIDRef="0"><hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="1134" tabStop="8000" tabStopVal="4000" tabStopUnit="HWPUNIT" outlineShapeIDRef="1" memoShapeIDRef="0" textVerticalWidthHead="0" masterPageCnt="0">' +
    '<hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0" strtnum="0"/>' +
    '<hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0"/>' +
    '<hp:visibility hideFirstHeader="0" hideFirstFooter="0" hideFirstMasterPage="0" border="SHOW_ALL" fill="SHOW_ALL" hideFirstPageNum="0" hideFirstEmptyLine="0" showLineNumber="0"/>' +
    '<hp:lineNumberShape restartType="0" countBy="0" distance="0" startNumber="0"/>' +
    '<hp:pagePr landscape="WIDELY" width="59528" height="84188" gutterType="LEFT_ONLY">' +
    '<hp:margin header="4252" footer="4252" gutter="0" left="8504" right="8504" top="5668" bottom="4252"/>' +
    '</hp:pagePr>' +
    '</hp:secPr>' +
    `<hp:t>${esc(lines[0])}</hp:t></hp:run>`;
  const first = `<hp:p id="1" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">${firstRun}</hp:p>`;
  const rest = lines.slice(1).map((line, i) => para(line, i + 2));
  return (
    XML_DECL +
    '<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">' +
    first +
    rest.join('') +
    '</hs:sec>'
  );
}

function makeSkeleton() {
  const entries = {
    mimetype: [strToU8('application/hwp+zip'), { level: 0 }],
    'version.xml': strToU8(VERSION_XML),
    'META-INF/container.xml': strToU8(CONTAINER_XML),
    'META-INF/manifest.xml': strToU8(MANIFEST_XML),
    'Contents/content.hpf': strToU8(CONTENT_HPF),
    'Contents/header.xml': strToU8(HEADER_XML),
    'Contents/section0.xml': strToU8(buildSectionXml()),
    'settings.xml': strToU8(SETTINGS_XML),
  };
  const zipped = zipSync(entries);
  mkdirSync(dirname(TEMPLATE_PATH), { recursive: true });
  writeFileSync(TEMPLATE_PATH, zipped);
  console.log(`스켈레톤 템플릿 생성 완료: ${TEMPLATE_PATH} (${zipped.length.toLocaleString()} bytes)`);
  console.log('주의: 개발용 대체물입니다. 최종 인수는 한글에서 변환한 실제 공고문 템플릿으로 진행하세요.');
}

// ---------------------------------------------------------------------------
// verify / fill-test
// ---------------------------------------------------------------------------

function loadTemplate() {
  const buf = readFileSync(TEMPLATE_PATH);
  return unzipSync(new Uint8Array(buf));
}

function sectionNames(entries) {
  return Object.keys(entries).filter((name) => /^Contents\/section\d+\.xml$/.test(name));
}

function verify() {
  const entries = loadTemplate();
  let failed = false;

  const mimetype = entries['mimetype'] ? strFromU8(entries['mimetype']) : null;
  if (mimetype !== 'application/hwp+zip') {
    console.error(`✗ mimetype 이상: ${JSON.stringify(mimetype)}`);
    failed = true;
  } else {
    console.log('✓ mimetype: application/hwp+zip');
  }

  const sections = sectionNames(entries);
  if (sections.length === 0) {
    console.error('✗ Contents/section*.xml 없음');
    process.exit(1);
  }
  const xml = sections.map((name) => strFromU8(entries[name])).join('\n');

  // 템플릿에 일부 키만 있어도 됨 (없는 키는 앱이 건너뜀) → 누락은 안내만.
  // 단, 앱 키 목록에 없는 {{마커}}는 생성 시 오류를 던지므로 검증 실패로 처리.
  const present = [];
  const absent = [];
  for (const key of PLACEHOLDER_KEYS) {
    (xml.includes(`{{${key}}}`) ? present : absent).push(key);
  }
  for (const key of present) console.log(`✓ {{${key}}}`);
  if (absent.length) console.log(`- 템플릿에 없음(선택): ${absent.map((k) => `{{${k}}}`).join(', ')}`);
  if (present.length === 0) {
    console.error('✗ 앱이 치환할 수 있는 자리표시자가 하나도 없음');
    failed = true;
  }

  const known = new Set(PLACEHOLDER_KEYS.map((k) => `{{${k}}}`));
  const unknown = (xml.match(/\{\{[^}]{1,40}\}\}/g) ?? []).filter((m) => !known.has(m));
  if (unknown.length) {
    console.error(
      `✗ 앱 키 목록에 없는 자리표시자 (생성 시 오류 발생 — 오타이거나 run 분할, 또는 hwpx.ts NOTICE_PLACEHOLDER_KEYS에 키 추가 필요): ${[...new Set(unknown)].join(', ')}`,
    );
    failed = true;
  }

  if (failed) process.exit(1);
  console.log(`\n검증 통과: ${sections.length}개 섹션, 자리표시자 ${present.length}/${PLACEHOLDER_KEYS.length}개 사용 중`);
}

function fillTest() {
  const raw = new Uint8Array(readFileSync(TEMPLATE_PATH));
  const entries = unzipSync(raw);
  const dummy = Object.fromEntries(PLACEHOLDER_KEYS.map((key, i) => [key, `테스트값${i + 1} & <검증>`]));
  const usedKeys = new Set(); // 템플릿에 실제 존재해서 치환이 일어난 키

  for (const name of sectionNames(entries)) {
    let xml = strFromU8(entries[name]);
    for (const [key, value] of Object.entries(dummy)) {
      if (!xml.includes(`{{${key}}}`)) continue;
      usedKeys.add(key);
      const escaped = value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
      xml = xml.replaceAll(`{{${key}}}`, escaped);
    }
    const leftovers = xml.match(/\{\{[^}]{1,40}\}\}/g) ?? [];
    if (leftovers.length) {
      console.error(`✗ ${name} 잔여 마커: ${leftovers.join(', ')}`);
      process.exit(1);
    }
    entries[name] = strToU8(xml);
  }

  if (usedKeys.size === 0) {
    console.error('✗ 템플릿에 치환 가능한 자리표시자가 없음');
    process.exit(1);
  }

  const ordered = { mimetype: [entries['mimetype'], { level: 0 }] };
  for (const [name, data] of Object.entries(entries)) {
    if (name !== 'mimetype') ordered[name] = data;
  }
  const output = zipSync(ordered);

  // 재해석 무결성 + mimetype이 첫 로컬 헤더인지(오프셋 0 파일명) 확인
  const reparsed = unzipSync(output);
  if (!reparsed['Contents/section0.xml']) {
    console.error('✗ 재압축 결과에 section0.xml 없음');
    process.exit(1);
  }
  const nameLen = output[26] | (output[27] << 8);
  const firstName = strFromU8(output.subarray(30, 30 + nameLen));
  if (firstName !== 'mimetype') {
    console.error(`✗ 첫 zip 엔트리가 mimetype이 아님: ${firstName}`);
    process.exit(1);
  }
  // 템플릿에 실제 있던 키들의 더미값이 이스케이프된 형태로 들어갔는지 확인
  const filled = sectionNames(reparsed)
    .map((name) => strFromU8(reparsed[name]))
    .join('\n');
  for (const key of usedKeys) {
    const idx = PLACEHOLDER_KEYS.indexOf(key);
    const expected = `테스트값${idx + 1} &amp; &lt;검증&gt;`;
    if (!filled.includes(expected)) {
      console.error(`✗ {{${key}}} 치환·이스케이프 결과가 기대와 다름 (기대 조각: ${expected})`);
      process.exit(1);
    }
  }
  console.log(
    `fill-test 통과: 치환 ${usedKeys.size}개 키, 잔여 마커 0개, mimetype 첫 엔트리, XML 이스케이프 정상 (출력 ${output.length.toLocaleString()} bytes)`,
  );
}

const command = process.argv[2];
if (command === 'make-skeleton') makeSkeleton();
else if (command === 'verify') verify();
else if (command === 'fill-test') fillTest();
else {
  console.error('사용법: node scripts/hwpx-tools.mjs <make-skeleton|verify|fill-test>');
  process.exit(1);
}
