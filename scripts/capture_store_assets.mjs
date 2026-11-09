// Builds Chrome Web Store screenshots (1280x800) and small promo tiles (440x280) in English,
// Chinese, and Spanish from the real extension popup, loaded unpacked in Chromium.
//
// Usage:
//   ./scripts/build_release.sh
//   npm install --no-save playwright && npx playwright install chromium
//   node scripts/capture_store_assets.mjs
// Set CHROMIUM_PATH to use a different Chromium build (it must allow --load-extension).
import { createHash } from 'node:crypto';
import { mkdirSync, mkdtempSync, readFileSync, realpathSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

const { chromium } = await import('playwright').catch(() => import('playwright-core'));

const root = fileURLToPath(new URL('..', import.meta.url));
const extDir = join(root, 'dist', 'mail-merge-draft-helper');
const screenshotDir = join(root, 'store-assets', 'screenshots');
const promoDir = join(root, 'store-assets', 'promo');
const LANGS = ['en', 'zh', 'es'];

// Unpacked extension IDs are derived from the folder path.
function extensionId(dir) {
  const hex = createHash('sha256').update(realpathSync(dir)).digest('hex').slice(0, 32);
  return [...hex].map(ch => String.fromCharCode(97 + parseInt(ch, 16))).join('');
}

const TEMPLATE = {
  en: { subject: "You're invited, {{First Name}}!", body: "Hi {{First Name|there}},\n\nI'd like to invite you to our workshop next week.\n\nDate: Tuesday, 12:30-2:00 PM\nLocation: Room 101\n\nHope to see you there!" },
  zh: { subject: '{{First Name}}，邀请你参加工作坊', body: '{{First Name|同学}} 你好，\n\n诚邀你参加下周的工作坊。\n\n时间：周二 12:30-14:00\n地点：101 教室\n\n期待你的到来！' },
  es: { subject: '{{First Name}}, te invitamos al taller', body: 'Hola {{First Name|estudiante}}:\n\nTe invito a nuestro taller de la próxima semana.\n\nFecha: martes, 12:30-14:00\nLugar: aula 101\n\n¡Espero verte allí!' }
};

const COPY = {
  en: {
    tagline: 'Outlook & Gmail mail merge from Excel. No sign\u2011in.',
    shots: [
      { file: '01-overview', title: 'Mail merge for Outlook and Gmail', sub: 'Personalized emails from an Excel or CSV list. No sign\u2011in, no mailbox access.', bullets: ['Preview every email before it goes out', 'Review notes catch nicknames and odd addresses', 'Send in batches, with your progress saved'], raw: 'step3' },
      { file: '02-excel', title: 'Start from your Excel list', sub: 'Pick the email column and only the columns your message needs.', bullets: ['Reads .xlsx and CSV right in the browser', 'Grades, IDs, and other columns stay out', 'Or try it first with sample data'], raw: 'step1' },
      { file: '03-template', title: 'Write once, personalize for everyone', sub: 'Insert columns as variables and catch mistakes as you type.', bullets: ['Type {{ to pick a column', 'Defaults like {{First Name|there}}', 'Misspelled variables flagged instantly'], raw: 'step2' },
      { file: '04-send', title: 'Test first, then send in batches', sub: 'Drafts open in Outlook Web or Gmail, 50 at a time or one by one.', bullets: ['Send yourself a [TEST] copy first', 'The batch counter remembers where you stopped', 'Optional Auto-Send (experimental)'], raw: 'step4' },
      { file: '05-private', title: 'Private by design', sub: 'Everything happens in your browser. Your list is never uploaded.', bullets: ['No account to create', 'Nothing to authorize in Microsoft 365 or Google', 'Interface in English, 中文, and Español'], raw: 'languages' }
    ]
  },
  zh: {
    tagline: '从 Excel 群发个性化邮件，支持 Outlook 和 Gmail，无需登录。',
    shots: [
      { file: '01-overview', title: 'Outlook 和 Gmail 的邮件合并工具', sub: '从 Excel 或 CSV 名单生成个性化邮件。无需登录，不读取你的邮箱。', bullets: ['发出前逐封预览', '自动标出昵称、异常邮箱等问题', '分批发送，自动记住进度'], raw: 'step3' },
      { file: '02-excel', title: '直接用你的 Excel 名单', sub: '选好邮箱列，只载入邮件需要的列。', bullets: ['在浏览器里直接读取 .xlsx 和 CSV', '成绩、学号等其他列默认不载入', '也可以先用示例数据试一试'], raw: 'step1' },
      { file: '03-template', title: '写一次，人人个性化', sub: '把列当作变量插入，写错当场提示。', bullets: ['输入 {{ 选择列名', '支持默认值，如 {{First Name|同学}}', '变量写错立即提示'], raw: 'step2' },
      { file: '04-send', title: '先测试，再分批发送', sub: '草稿在 Outlook 网页版或 Gmail 中打开，每批 50 封或逐封打开。', bullets: ['先给自己发一封 [TEST] 测试', '分批计数，记住发到哪里', '可选的自动发送（实验功能）'], raw: 'step4' },
      { file: '05-private', title: '为隐私而设计', sub: '一切都在你的浏览器里完成，名单不会上传。', bullets: ['无需注册账号', '无需授权 Microsoft 365 或 Google', '界面支持 English、中文、Español'], raw: 'languages' }
    ]
  },
  es: {
    tagline: 'Combinación de correo para Outlook y Gmail desde Excel, sin iniciar sesión.',
    shots: [
      { file: '01-overview', title: 'Combinación de correo para Outlook y Gmail', sub: 'Correos personalizados desde una lista de Excel o CSV. Sin iniciar sesión ni acceso a tu buzón.', bullets: ['Revisa cada correo antes de enviarlo', 'Avisos de apodos y direcciones poco comunes', 'Envía por lotes y conserva tu progreso'], raw: 'step3' },
      { file: '02-excel', title: 'Empieza con tu lista de Excel', sub: 'Elige la columna de correo y solo las columnas que necesita tu mensaje.', bullets: ['Lee .xlsx y CSV en el navegador', 'Notas, identificadores y demás quedan fuera', 'O pruébalo antes con datos de ejemplo'], raw: 'step1' },
      { file: '03-template', title: 'Escribe una vez, personaliza para todos', sub: 'Inserta columnas como variables y detecta errores mientras escribes.', bullets: ['Escribe {{ para elegir una columna', 'Valores por defecto: {{First Name|estudiante}}', 'Variables mal escritas, señaladas al instante'], raw: 'step2' },
      { file: '04-send', title: 'Primero una prueba, luego por lotes', sub: 'Los borradores se abren en Outlook Web o Gmail, de 50 en 50 o uno a uno.', bullets: ['Envíate primero una copia [TEST]', 'El contador recuerda dónde te quedaste', 'Envío automático opcional (experimental)'], raw: 'step4' },
      { file: '05-private', title: 'Privado por diseño', sub: 'Todo ocurre en tu navegador. Tu lista nunca se sube.', bullets: ['Sin crear cuentas', 'Nada que autorizar en Microsoft 365 ni Google', 'Interfaz en English, 中文 y Español'], raw: 'languages' }
    ]
  }
};

async function scrollTo(page, selector, offset) {
  await page.evaluate(([sel, off]) => {
    const el = document.querySelector(sel);
    window.scrollTo(0, el ? el.getBoundingClientRect().top + window.scrollY + off : 0);
  }, [selector, offset]);
  await page.waitForTimeout(150);
}

// Real popup views at Chrome's popup size (460 x 600), one per step.
async function capturePopup(context, id, lang) {
  const page = await context.newPage();
  await page.setViewportSize({ width: 460, height: 600 });
  await page.goto(`chrome-extension://${id}/popup.html`);
  // Start every language from a fresh popup, not the state the previous language left behind.
  await page.evaluate(() => chrome.storage.local.clear());
  await page.reload();
  await page.waitForTimeout(300);
  await page.selectOption('#uiLanguage', lang);
  const shots = {};
  const snap = async name => { shots[name] = await page.screenshot(); };

  await snap('top');

  await page.setInputFiles('#csvFile', join(root, 'examples', 'sample-roster.xlsx'));
  await page.waitForTimeout(400);
  await scrollTo(page, '#mappingPanel', -120);
  await snap('step1');
  await page.click('.mapping-apply');
  await page.click('#wizardNext');
  await page.waitForTimeout(200);

  await page.fill('#subject', TEMPLATE[lang].subject);
  await page.fill('#body', TEMPLATE[lang].body);
  await page.focus('#body');
  await page.keyboard.press('ControlOrMeta+End');
  await page.keyboard.type('\n{{');
  await page.waitForTimeout(200);
  await scrollTo(page, 'label[for="subject"]', -12);
  await snap('step2');
  await page.keyboard.press('Escape');
  await page.fill('#body', TEMPLATE[lang].body);

  await page.click('#wizardNext');
  await page.waitForTimeout(400);
  await scrollTo(page, 'body', 0);
  await snap('step3');

  await page.click('#wizardNext');
  await page.fill('#testEmail', 'me@example.com');
  await page.evaluate(() => {
    const box = document.querySelector('#sendControls').getBoundingClientRect();
    window.scrollTo(0, box.bottom + window.scrollY - 600 + 16);
  });
  await page.waitForTimeout(150);
  await snap('step4');

  await page.close();
  return shots;
}

const dataUrl = buffer => `data:image/png;base64,${buffer.toString('base64')}`;
const escapeHtml = text => String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const icon = dataUrl(readFileSync(join(root, 'icons', 'icon-128.png')));
const FONT = "-apple-system, BlinkMacSystemFont, 'Segoe UI', 'PingFang SC', 'Hiragino Sans GB', 'Microsoft YaHei', Roboto, Helvetica, Arial, sans-serif";

function screenshotHtml(copy, images) {
  const popups = images.map((src, i) => `<img class="popup${images.length > 1 ? ` pair pair-${i}` : ''}" src="${src}">`).join('');
  return `<!doctype html><html><head><meta charset="utf-8"><style>
    * { box-sizing: border-box; }
    body { margin: 0; width: 1280px; height: 800px; overflow: hidden; font-family: ${FONT}; color: #0f172a;
      background: radial-gradient(circle at 85% 20%, #dbeafe 0, transparent 45%), linear-gradient(135deg, #eff6ff 0%, #f8fafc 60%, #ffffff 100%);
      display: flex; align-items: center; padding: 0 72px 0 88px; gap: 56px; }
    .text { flex: 1; }
    .brand { display: flex; align-items: center; gap: 12px; font-size: 20px; font-weight: 700; color: #1d4ed8; margin-bottom: 30px; }
    .brand img { width: 40px; height: 40px; }
    h1 { font-size: 46px; line-height: 1.15; margin: 0 0 18px; letter-spacing: -0.01em; }
    .sub { font-size: 23px; line-height: 1.45; color: #475569; margin: 0 0 34px; }
    ul { list-style: none; padding: 0; margin: 0; }
    li { display: flex; align-items: flex-start; gap: 14px; font-size: 22px; line-height: 1.4; margin-bottom: 18px; }
    li::before { content: '✓'; flex: none; width: 30px; height: 30px; border-radius: 999px; background: #dcfce7; color: #166534;
      font-weight: 800; font-size: 17px; display: flex; align-items: center; justify-content: center; margin-top: 1px; }
    .shots { position: relative; width: 520px; height: 700px; flex: none; }
    .popup { position: absolute; top: 11px; left: 0; width: 520px; border-radius: 16px; border: 1px solid #cbd5e1;
      box-shadow: 0 30px 60px -20px rgba(15, 23, 42, 0.35), 0 8px 20px -8px rgba(15, 23, 42, 0.2); }
    .pair { width: 400px; }
    .pair-0 { top: 0; left: 0; transform: rotate(-3deg); }
    .pair-1 { top: 120px; left: 130px; transform: rotate(2deg); }
  </style></head><body>
    <div class="text">
      <div class="brand"><img src="${icon}">Mail Merge Draft Helper</div>
      <h1>${escapeHtml(copy.title)}</h1>
      <p class="sub">${escapeHtml(copy.sub)}</p>
      <ul>${copy.bullets.map(b => `<li>${escapeHtml(b)}</li>`).join('')}</ul>
    </div>
    <div class="shots">${popups}</div>
  </body></html>`;
}

function promoHtml(tagline) {
  return `<!doctype html><html><head><meta charset="utf-8"><style>
    body { margin: 0; width: 440px; height: 280px; overflow: hidden; font-family: ${FONT}; color: #ffffff;
      background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 60%, #1e40af 100%); padding: 40px 36px; box-sizing: border-box; }
    .row { display: flex; align-items: center; gap: 16px; }
    img { width: 72px; height: 72px; border-radius: 16px; background: #ffffff; padding: 6px; box-sizing: border-box; }
    .name { font-size: 29px; font-weight: 800; line-height: 1.1; }
    p { margin: 30px 0 0; font-size: 19px; line-height: 1.45; color: #dbeafe; }
  </style></head><body>
    <div class="row"><img src="${icon}"><div class="name">Mail Merge<br>Draft Helper</div></div>
    <p>${escapeHtml(tagline)}</p>
  </body></html>`;
}

const profile = mkdtempSync(join(tmpdir(), 'mail-merge-store-'));
const context = await chromium.launchPersistentContext(profile, {
  executablePath: process.env.CHROMIUM_PATH || undefined,
  headless: true,
  viewport: { width: 460, height: 600 },
  deviceScaleFactor: 2,
  args: [`--disable-extensions-except=${extDir}`, `--load-extension=${extDir}`]
});

try {
  const id = extensionId(extDir);
  const raw = {};
  for (const lang of LANGS) raw[lang] = await capturePopup(context, id, lang);

  // A fresh page per image: reusing one page across viewport sizes can leave stale tiles in the capture.
  const renderPng = async (html, width, height, file) => {
    const page = await context.newPage();
    await page.setViewportSize({ width, height });
    await page.setContent(html, { waitUntil: 'load' });
    // deviceScaleFactor is 2 for crisp popups; scale back to the exact store size.
    await page.screenshot({ path: file, scale: 'css' });
    await page.close();
  };

  for (const lang of LANGS) {
    const outDir = join(screenshotDir, lang);
    mkdirSync(outDir, { recursive: true });
    for (const shot of COPY[lang].shots) {
      const images = shot.raw === 'languages'
        ? LANGS.filter(other => other !== lang).map(other => dataUrl(raw[other].top))
        : [dataUrl(raw[lang][shot.raw])];
      await renderPng(screenshotHtml(shot, images), 1280, 800, join(outDir, `${shot.file}.png`));
      console.log('saved', lang, `${shot.file}.png`);
    }

    const promoOut = join(promoDir, lang);
    mkdirSync(promoOut, { recursive: true });
    await renderPng(promoHtml(COPY[lang].tagline), 440, 280, join(promoOut, 'small-promo-tile-440x280.png'));
    console.log('saved', lang, 'small-promo-tile-440x280.png');
  }
} finally {
  await context.close();
  rmSync(profile, { recursive: true, force: true });
}
