import assert from 'node:assert/strict';
import { existsSync } from 'node:fs';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { chromium } from 'playwright';

const reportPath = resolve('docs_0722_status/6조_프로젝트_진행상황_260722.html');
assert.ok(existsSync(reportPath), `보고서가 없습니다: ${reportPath}`);

const browser = await chromium.launch({ headless: true });
try {
  const page = await browser.newPage({ viewport: { width: 1440, height: 900 } });
  await page.goto(pathToFileURL(reportPath).href, { waitUntil: 'domcontentloaded' });

  const bodyText = await page.locator('body').innerText();
  for (const phrase of [
    '원가계산서를 어떻게 파싱했는가',
    'Workbook IR',
    '사용자가 별도로 첨부한 외부 제비율표',
    '사용자가 별도로 첨부한 외부 노임단가표',
    'AI 없이',
    '공고문',
    '관련 법령',
  ]) assert.ok(bodyText.includes(phrase), `필수 문구 누락: ${phrase}`);

  assert.equal(await page.locator('.sidebar .nav a').count(), 10);
  await page.locator('#parsing').scrollIntoViewIfNeeded();
  await page.waitForTimeout(300);
  assert.equal(await page.locator('.nav a.active').getAttribute('href'), '#parsing');
  assert.equal(
    await page.locator('.nav a').first().evaluate((node) => getComputedStyle(node).backgroundColor),
    'rgba(0, 0, 0, 0)',
  );
  assert.equal(await page.locator('.shot img').count(), 4);
  const images = await page.locator('.shot img').evaluateAll((nodes) =>
    nodes.map((image) => ({ complete: image.complete, width: image.naturalWidth })),
  );
  assert.ok(images.every((image) => image.complete && image.width > 0), JSON.stringify(images));

  await page.locator('.shot img').first().focus();
  await page.keyboard.press('Enter');
  assert.equal(await page.locator('#lightbox').evaluate((dialog) => dialog.open), true);
  await page.keyboard.press('Escape');
  assert.equal(await page.locator('#lightbox').evaluate((dialog) => dialog.open), false);

  await page.emulateMedia({ media: 'print' });
  assert.equal(await page.locator('.sidebar').evaluate((node) => getComputedStyle(node).display), 'none');

  await page.setViewportSize({ width: 390, height: 844 });
  await page.emulateMedia({ media: 'screen' });
  const overflow = await page.evaluate(() => document.documentElement.scrollWidth - innerWidth);
  assert.ok(overflow <= 1, `모바일 가로 넘침: ${overflow}px`);
  assert.equal(await page.locator('[data-mobile-menu]').count(), 1);

  await page.locator('[data-mobile-menu]').click();
  assert.equal(await page.locator('.sidebar').evaluate((node) => getComputedStyle(node).display), 'block');
  await page.emulateMedia({ media: 'print' });
  assert.equal(await page.locator('.sidebar').evaluate((node) => getComputedStyle(node).display), 'none');
  assert.equal(await page.locator('.report-tools').evaluate((node) => getComputedStyle(node).display), 'none');

  const externalResources = await page.locator('link[href^="http"],script[src^="http"],img[src^="http"]').count();
  assert.equal(externalResources, 0);
} finally {
  await browser.close();
}

console.log('project report QA passed');
