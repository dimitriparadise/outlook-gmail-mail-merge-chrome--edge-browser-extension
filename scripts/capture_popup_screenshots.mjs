import { chromium } from 'playwright';
import { mkdir } from 'node:fs/promises';
import { resolve } from 'node:path';

const root = resolve(new URL('..', import.meta.url).pathname);
const outputDir = resolve(root, 'store-assets/raw-popup');
const popupUrl = `file://${resolve(root, 'popup.html')}`;

async function setupPage(browser) {
  const page = await browser.newPage({ viewport: { width: 460, height: 860 }, deviceScaleFactor: 2 });
  await page.addInitScript(() => {
    const state = {};
    window.chrome = {
      storage: {
        local: {
          async get(key) {
            if (typeof key === 'string') return { [key]: state[key] };
            return { ...state };
          },
          async set(next) {
            Object.assign(state, next);
          },
          async remove(key) {
            delete state[key];
          }
        }
      },
      tabs: {
        create(_options, callback) {
          callback?.();
        }
      },
      runtime: {}
    };
  });
  await page.goto(popupUrl);
  await page.waitForLoadState('load');
  return page;
}

async function fillAndGenerate(page) {
  await page.locator('#csvText').fill(`Name,Email,Course,Section,DueDate,CcEmail,BccEmail
John,john@example.com,ISOM 210,A,Friday,ta@example.com,archive@example.com
Jane,jane@example.com,ISOM 340,B,Monday,ta@example.com,archive@example.com`);
  await page.locator('#composeMode').selectOption('gmail');
  await page.locator('#subject').fill('Reminder for {{Course}}');
  await page.locator('#cc').fill('{{CcEmail}}');
  await page.locator('#bcc').fill('{{BccEmail}}');
  await page.locator('#body').fill('Hi {{Name}},\\n\\nThis is a quick reminder about {{Course}} section {{Section}}, due {{DueDate}}.');
  await page.locator('#loadBtn').click();
  await page.locator('#preview').waitFor({ state: 'visible' });
}

async function main() {
  await mkdir(outputDir, { recursive: true });
  const browser = await chromium.launch({ headless: true });

  try {
    const initial = await setupPage(browser);
    await initial.screenshot({ path: resolve(outputDir, '01-initial-popup.png'), fullPage: true });
    await initial.close();

    const preview = await setupPage(browser);
    await fillAndGenerate(preview);
    await preview.screenshot({ path: resolve(outputDir, '02-generated-preview.png'), fullPage: true });

    await preview.locator('#previewNextBtn').click();
    await preview.screenshot({ path: resolve(outputDir, '03-second-preview.png'), fullPage: true });

    await preview.locator('#autoSend').check();
    await preview.screenshot({ path: resolve(outputDir, '04-auto-send-labels.png'), fullPage: true });
    await preview.close();
  } finally {
    await browser.close();
  }
}

main().catch(error => {
  console.error(error);
  process.exit(1);
});
