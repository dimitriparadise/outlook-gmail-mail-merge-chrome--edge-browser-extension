// Captures the tutorial screenshots in docs/images (English) and docs/images/zh (Chinese)
// from the real extension popup, loaded unpacked in Chromium, using the files in examples/.
//
// Usage:
//   ./scripts/build_release.sh
//   npm install --no-save playwright && npx playwright install chromium
//   node scripts/capture_docs_screenshots.mjs
// Set CHROMIUM_PATH to use a different Chromium build (it must allow --load-extension).
import { createHash } from 'node:crypto';
import { mkdirSync, mkdtempSync, realpathSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const { chromium } = await import('playwright').catch(() => import('playwright-core'));

const repo = fileURLToPath(new URL('..', import.meta.url));
const extDir = path.join(repo, 'dist', 'mail-merge-draft-helper');

// Unpacked extension IDs are derived from the folder path.
function extensionId(dir) {
  const hex = createHash('sha256').update(realpathSync(dir)).digest('hex').slice(0, 32);
  return [...hex].map(ch => String.fromCharCode(97 + parseInt(ch, 16))).join('');
}

// Language and output folder of the run in progress; used by the screenshot helpers.
let lang = 'en';
let outDir = '';

const HIGHLIGHT = 'outline: 3px solid #ef4444 !important; outline-offset: 3px !important;';

async function highlight(page, selector) {
  await page.evaluate(([sel, css]) => {
    document.querySelectorAll('[data-shot-hl]').forEach(el => { el.style.cssText = el.dataset.shotHl; delete el.dataset.shotHl; });
    if (!sel) return;
    for (const el of document.querySelectorAll(sel)) {
      el.dataset.shotHl = el.style.cssText;
      el.style.cssText += css;
    }
  }, [selector, HIGHLIGHT]);
}

async function shotOf(page, file, selectors, pad = 12) {
  await page.evaluate(() => window.scrollTo(0, 0));
  const boxes = [];
  for (const sel of selectors) {
    const box = await page.locator(sel).first().boundingBox();
    if (box) boxes.push(box);
  }
  const x = Math.max(0, Math.min(...boxes.map(b => b.x)) - pad);
  const y = Math.max(0, Math.min(...boxes.map(b => b.y)) - pad);
  const right = Math.min(460, Math.max(...boxes.map(b => b.x + b.width)) + pad);
  const bottom = Math.max(...boxes.map(b => b.y + b.height)) + pad;
  await page.screenshot({ path: path.join(outDir, file), clip: { x, y, width: right - x, height: bottom - y }, fullPage: true });
  console.log('saved', lang, file);
}

const TEMPLATE = {
  en: { subject: "You're invited, {{First Name}}!", body: "Hi {{First Name}},\n\nI'd like to invite you to our workshop next week.\n\nDate: Tuesday, 12:30-2:00 PM\nLocation: Room 101\n\nHope to see you there!" },
  zh: { subject: '{{First Name}}，邀请你参加工作坊', body: '{{First Name}} 你好，\n\n诚邀你参加下周的工作坊。\n\n时间：周二 12:30-14:00\n地点：101 教室\n\n期待你的到来！' }
};

async function capture(runLang, runDir) {
  lang = runLang;
  outDir = runDir;
  const userDataDir = mkdtempSync(path.join(tmpdir(), `mail-merge-docs-${lang}-`));
  mkdirSync(outDir, { recursive: true });
  const id = extensionId(extDir);
  const ctx = await chromium.launchPersistentContext(userDataDir, {
    executablePath: process.env.CHROMIUM_PATH || undefined,
    headless: true,
    viewport: { width: 1100, height: 700 },
    deviceScaleFactor: 2,
    args: [`--disable-extensions-except=${extDir}`, `--load-extension=${extDir}`, `--lang=${lang === 'zh' ? 'zh-CN' : 'en-US'}`]
  });

  if (lang === 'en') {
    const ext = await ctx.newPage();
    await ext.goto('chrome://extensions');
    await ext.locator('#devMode').click();
    await ext.waitForTimeout(500);
    for (const sel of ['#devMode', '#loadUnpacked']) {
      await ext.locator(sel).evaluate((el, css) => { el.style.cssText += css; }, HIGHLIGHT);
    }
    await ext.screenshot({ path: path.join(outDir, '01-load-unpacked.png'), clip: { x: 0, y: 0, width: 1100, height: 425 } });
    console.log('saved en 01-load-unpacked.png');
    await ext.close();
  }

  const page = await ctx.newPage();
  await page.setViewportSize({ width: 460, height: 900 });
  await page.goto(`chrome-extension://${id}/popup.html`);
  await page.waitForTimeout(300);
  await page.selectOption('#uiLanguage', lang);
  await page.waitForTimeout(200);

  // Step 1 of a fresh popup, with the step bar highlighted.
  await highlight(page, '.step-bar');
  await page.screenshot({ path: path.join(outDir, '02-popup-overview.png'), fullPage: true });
  console.log('saved', lang, '02-popup-overview.png');
  await highlight(page, null);

  // Plan bar and the status after importing a plan.
  await page.setInputFiles('#planFile', path.join(repo, 'examples', 'sample-plan.json'));
  await page.waitForTimeout(300);
  await highlight(page, '#importPlanBtn, #exportPlanBtn');
  await shotOf(page, '04-import-plan-bar.png', ['.app-header', '.plan-bar']);
  await highlight(page, '#status');
  await shotOf(page, '04-import-plan-status.png', ['#status']);
  await highlight(page, null);

  // Excel import: the column picker, then the loaded list.
  await page.setInputFiles('#csvFile', path.join(repo, 'examples', 'sample-roster.xlsx'));
  await page.waitForTimeout(400);
  await highlight(page, '.mapping-columns, .mapping-apply');
  await shotOf(page, '03-excel-columns.png', ['main > section.card:nth-of-type(2)']);
  await highlight(page, null);
  await page.click('.mapping-apply');
  await page.waitForTimeout(300);
  await highlight(page, '#csvSummary, #wizardNext');
  await shotOf(page, '03-excel-loaded.png', ['main > section.card:nth-of-type(2)', '.wizard-nav']);
  await highlight(page, null);

  // Step 2: the template.
  await page.click('#wizardNext');
  await page.waitForTimeout(200);
  await page.fill('#subject', TEMPLATE[lang].subject);
  await page.fill('#body', TEMPLATE[lang].body);
  await highlight(page, '.variable-panel');
  await shotOf(page, '05-template.png', ['main > section.card:nth-of-type(3)']);
  await highlight(page, null);

  // Typing {{ shows column suggestions.
  await page.focus('#body');
  await page.keyboard.press('ControlOrMeta+End');
  await page.keyboard.type('\n\n{{La');
  await page.waitForTimeout(200);
  await highlight(page, '#variableSuggest');
  await shotOf(page, '05-autocomplete.png', ['#body', '#variableSuggest']);
  await highlight(page, null);
  await page.keyboard.press('Escape');
  await page.fill('#body', TEMPLATE[lang].body);

  // A misspelled variable is flagged right away.
  await page.fill('#subject', TEMPLATE[lang].subject.replace('{{First Name}}', '{{Frist Name}}'));
  await page.dispatchEvent('#subject', 'input');
  await page.waitForTimeout(200);
  await highlight(page, '#templateWarnings');
  await shotOf(page, '05-warning.png', ['label[for="subject"]', '#templateWarnings']);
  await highlight(page, null);
  await page.fill('#subject', TEMPLATE[lang].subject);
  await page.dispatchEvent('#subject', 'input');

  // Step 3: generate, review notes, preview.
  await page.click('#wizardNext');
  await page.waitForTimeout(400);
  await highlight(page, '.step-tab.done, .step-tab.active');
  await shotOf(page, '02-step-bar.png', ['.step-bar']);
  await highlight(page, '#reviewPanel');
  await shotOf(page, '07-review.png', ['#status', '#reviewPanel']);
  await highlight(page, '#previewControls');
  await shotOf(page, '07-preview.png', ['#previewControls', '#preview']);
  await highlight(page, null);

  // Step 4: settings, test send, batches.
  await page.click('#wizardNext');
  await page.waitForTimeout(200);
  await page.fill('#batchSize', '3');
  await page.dispatchEvent('#batchSize', 'change');
  await highlight(page, '.segmented-control, #batchSize');
  await shotOf(page, '06-settings.png', ['main > section.card:nth-of-type(4)']);
  await highlight(page, null);
  await page.fill('#testEmail', 'me@example.com');
  await highlight(page, '#sendTestBtn, #openBatchBtn');
  await shotOf(page, '08-send.png', ['#sendControls']);
  await highlight(page, null);

  // Auto-Send wording.
  await page.check('#autoSend');
  await page.waitForTimeout(200);
  await highlight(page, '.setting-callout');
  await shotOf(page, '09-auto-send.png', ['.setting-callout']);
  await highlight(page, '#sendTestBtn, #openBatchBtn');
  await shotOf(page, '09-auto-send-buttons.png', ['#sendControls .send-box:first-child', '#openBatchBtn']);
  await highlight(page, null);
  await page.uncheck('#autoSend');

  // Help links: under an error, and in the footer.
  await page.fill('#testEmail', '');
  await page.click('#sendTestBtn');
  await page.waitForTimeout(200);
  await highlight(page, '#statusHelp');
  await shotOf(page, '10-help-error.png', ['#status', '#statusHelp']);
  await highlight(page, '#helpLink, #wizardToggle');
  await shotOf(page, '10-help-footer.png', ['.footer-note']);
  await highlight(page, null);

  await ctx.close();
  rmSync(userDataDir, { recursive: true, force: true });
}

rmSync(path.join(repo, 'docs', 'images'), { recursive: true, force: true });
await capture('en', path.join(repo, 'docs', 'images'));
await capture('zh', path.join(repo, 'docs', 'images', 'zh'));
