import assert from 'node:assert/strict';
import fs from 'node:fs';
import test from 'node:test';
import vm from 'node:vm';
import { makeXlsx } from './xlsx-fixture.mjs';

function createElement(id, initial = {}) {
  const listeners = new Map();
  const classes = new Set(initial.classes || []);
  return {
    id,
    value: initial.value ?? '',
    checked: initial.checked ?? false,
    textContent: '',
    innerHTML: '',
    className: '',
    files: [],
    max: '',
    disabled: false,
    style: {},
    classList: {
      add(name) { classes.add(name); },
      remove(name) { classes.delete(name); },
      toggle(name, force) {
        const on = force === undefined ? !classes.has(name) : Boolean(force);
        if (on) classes.add(name);
        else classes.delete(name);
        return on;
      },
      contains(name) { return classes.has(name); }
    },
    addEventListener(type, handler) {
      if (!listeners.has(type)) listeners.set(type, []);
      listeners.get(type).push(handler);
    },
    async dispatchEvent(event) {
      for (const handler of listeners.get(event.type) || []) {
        await handler(event);
      }
    },
    async click() {
      await this.dispatchEvent({ type: 'click' });
    },
    appendChild() {},
    focus() {},
    setSelectionRange() {},
    get selectionStart() {
      return this.value.length;
    },
    get selectionEnd() {
      return this.value.length;
    }
  };
}

function loadPopup(state) {
  const elements = new Map();
  const defaults = {
    ccBccContainer: { classes: ['hidden'] },
    csvText: { value: '' },
    composeMode: { value: 'outlook' },
    autoSend: { checked: false },
    subject: { value: 'Reminder for {{Course}}' },
    cc: { value: '' },
    bcc: { value: '' },
    body: { value: 'Hi {{Name}},\n\nThis is a quick reminder about {{Course}}.' }
  };
  const ids = [
    'csvText',
    'composeMode',
    'autoSend',
    'subject',
    'cc',
    'bcc',
    'body',
    'variableButtons',
    'sendControls',
    'previewControls',
    'rangeStart',
    'rangeEnd',
    'status',
    'preview',
    'previewCounter',
    'previewPrevBtn',
    'previewNextBtn',
    'loadBtn',
    'openNextBtn',
    'openRangeBtn',
    'copyRecipientHint',
    'composeModeHint',
    'csvFile'
  ];
  for (const id of ids) elements.set(id, createElement(id, defaults[id]));

  const stored = state === undefined ? {} : { outlookMailMergeState: state };
  const saved = [];
  const openedTabs = [];
  const createdElements = [];
  const blobs = [];
  const context = {
    console,
    // Run timers immediately so batch delays do not slow the tests down.
    setTimeout(fn) {
      if (typeof fn === 'function') fn();
      return 0;
    },
    TextDecoder,
    DecompressionStream,
    Response,
    // The exported plan's JSON is readable as blob.text; real Blob behavior is kept for xlsx.js.
    Blob: class TestBlob extends Blob {
      constructor(parts, options = {}) {
        super(parts, options);
        this.text = parts.every(part => typeof part === 'string') ? parts.join('') : undefined;
      }
    },
    URL: {
      createObjectURL(blob) {
        blobs.push(blob);
        return 'blob:test';
      },
      revokeObjectURL() {}
    },
    Event: class Event {
      constructor(type, options = {}) {
        this.type = type;
        this.bubbles = Boolean(options.bubbles);
      }
    },
    document: {
      getElementById(id) {
        if (!elements.has(id)) elements.set(id, createElement(id));
        return elements.get(id);
      },
      createElement(tag) {
        const element = createElement(tag);
        createdElements.push(element);
        return element;
      }
    },
    chrome: {
      storage: {
        local: {
          async get() {
            return stored;
          },
          async set(next) {
            saved.push(next);
            Object.assign(stored, next);
          }
        }
      },
      tabs: {
        create(options, callback) {
          openedTabs.push(options);
          callback();
        }
      },
      runtime: {}
    }
  };
  vm.createContext(context);
  // popup.html loads these three scripts in this order.
  const script = ['../i18n.js', '../xlsx.js', '../popup.js']
    .map(file => fs.readFileSync(new URL(file, import.meta.url), 'utf8'))
    .join('\n');
  vm.runInContext(`${script}
globalThis.__popupTestApi = {
  parseCSV,
  applyTemplate,
  validateMessages,
  findMissingTemplateVariables,
  syncDraftActionLabels,
  restoreState,
  makeMessages,
  composeUrl,
  parsePlan,
  buildPlan,
  rowsToCSV,
  reviewRecipients,
  nextBatchRange,
  mappingToCSV,
  guessEmailColumn,
  handleRecipientFile,
  setUiLanguage,
  applyLanguageSelection,
  t,
  UI_STRINGS,
  state: () => ({ rows, currentIndex, generated, dismissedNotes, mapping })
};`, context);
  return { context, elements, saved, openedTabs, createdElements, blobs, api: context.__popupTestApi };
}

test('restoreState preserves intentionally blank subject and body', async () => {
  const { elements, api } = loadPopup({
    csvText: '',
    composeMode: 'outlook',
    autoSend: false,
    subject: '',
    cc: '',
    bcc: '',
    body: '',
    rows: [],
    csvHeaders: [],
    generated: [],
    currentIndex: 0,
    previewIndex: 0
  });

  await api.restoreState();

  assert.equal(elements.get('subject').value, '');
  assert.equal(elements.get('body').value, '');
});

test('parseCSV supports quoted fields containing newlines', () => {
  const { api } = loadPopup();

  const rows = api.parseCSV('Name,Email,Notes\nJane,jane@example.com,"Line one\nLine two"');

  assert.equal(rows.length, 1);
  assert.equal(rows[0].Notes, 'Line one\nLine two');
});

test('parseCSV explains that wrapped text still needs a real line break', () => {
  const { api } = loadPopup();

  assert.throws(
    () => api.parseCSV('Name, Email, Course, Section, DueDate, John, john@example.com, ISOM 210, A, Friday'),
    /real line break.*press Enter/i
  );
});

test('findMissingTemplateVariables reports variables that are not CSV headers', () => {
  const { api } = loadPopup();

  const missing = api.findMissingTemplateVariables(
    ['Name', 'Email', 'Course'],
    ['Hi {{Name}}', 'Reminder for {{Coursee}}']
  );

  assert.deepEqual(Array.from(missing), ['Coursee']);
});

test('validateMessages reports invalid, duplicate, and overlapping recipients', () => {
  const { api } = loadPopup();
  const messages = [
    {
      email: 'student@example.com',
      cc: 'student@example.com, bad-email',
      bcc: '',
      subject: 'Hello',
      body: 'Body'
    },
    {
      email: 'student@example.com',
      cc: '',
      bcc: '',
      subject: 'Hello',
      body: 'Body'
    }
  ];

  const issues = api.validateMessages(messages);

  assert(issues.some(issue => issue.includes('Invalid email address: bad-email')));
  assert(issues.some(issue => issue.includes('Duplicate To recipient: student@example.com')));
  assert(issues.some(issue => issue.includes('also appears in CC/BCC')));
});

test('composeUrl blocks Outlook auto-send when CC or BCC is present', () => {
  const { elements, api } = loadPopup();
  elements.get('composeMode').value = 'outlook';
  elements.get('autoSend').checked = true;

  assert.throws(
    () => api.composeUrl({
      email: 'student@example.com',
      cc: 'ta@example.com',
      bcc: '',
      subject: 'Hello',
      body: 'Body'
    }),
    /Auto-Send with Outlook Mode cannot be used with CC or BCC/
  );
});

test('syncDraftActionLabels changes button text when Auto-Send is enabled', () => {
  const { elements, api } = loadPopup();

  elements.get('autoSend').checked = true;
  api.syncDraftActionLabels();

  assert.equal(elements.get('openNextBtn').textContent, 'Open & Send Preview Email');
  assert.equal(elements.get('openRangeBtn').textContent, 'Open & Send Selected Emails');
});

test('syncDraftActionLabels uses draft wording when Auto-Send is disabled', () => {
  const { elements, api } = loadPopup();

  elements.get('autoSend').checked = false;
  api.syncDraftActionLabels();

  assert.equal(elements.get('openNextBtn').textContent, 'Open Preview Draft');
  assert.equal(elements.get('openRangeBtn').textContent, 'Open Selected Drafts');
});

test('single open button opens the currently previewed draft', async () => {
  const generated = [
    { email: 'first@example.com', cc: '', bcc: '', subject: 'First', body: 'Body', row: {} },
    { email: 'second@example.com', cc: '', bcc: '', subject: 'Second', body: 'Body', row: {} }
  ];
  const { elements, openedTabs, api } = loadPopup({
    csvText: 'Name,Email\nFirst,first@example.com\nSecond,second@example.com',
    composeMode: 'gmail',
    autoSend: false,
    templatePreset: '',
    subject: '{{Name}}',
    cc: '',
    bcc: '',
    body: 'Body',
    rows: [
      { Name: 'First', Email: 'first@example.com' },
      { Name: 'Second', Email: 'second@example.com' }
    ],
    csvHeaders: ['Name', 'Email'],
    generated,
    currentIndex: 0,
    previewIndex: 1
  });

  await api.restoreState();
  await elements.get('openNextBtn').click();

  assert.equal(openedTabs.length, 1);
  assert.match(openedTabs[0].url, /to=second%40example\.com/);
  assert.doesNotMatch(openedTabs[0].url, /first%40example\.com/);
});

test('parseCSV recognizes alternative email column headers and Chinese variables', () => {
  const { api } = loadPopup();
  const rows = api.parseCSV('姓名,邮箱,课程\n张三,zhangsan@example.com,计算机科学');

  assert.equal(rows.length, 1);
  assert.equal(rows[0].Email, 'zhangsan@example.com');
  assert.equal(rows[0]['姓名'], '张三');

  const rendered = api.applyTemplate('你好 {{姓名}}，欢迎参加 {{课程}}', rows[0]);
  assert.equal(rendered, '你好 张三，欢迎参加 计算机科学');
});

function planFile(name, plan) {
  const text = typeof plan === 'string' ? plan : JSON.stringify(plan);
  return { name, async text() { return text; } };
}

async function importPlan(elements, file) {
  elements.get('planFile').files = [file];
  await elements.get('planFile').dispatchEvent({ type: 'change' });
}

test('export plan downloads the CSV, templates, and settings without Auto-Send', async () => {
  const { elements, createdElements, blobs } = loadPopup();
  elements.get('csvText').value = 'Name,Email\nAda,ada@example.com';
  elements.get('subject').value = 'Hello {{Name}}';
  elements.get('cc').value = 'ta@example.com';
  elements.get('body').value = 'Hi {{Name}},\n\nSee you there.';
  elements.get('composeMode').value = 'gmail';
  elements.get('batchDelay').value = '2000';
  elements.get('autoSend').checked = true;

  await elements.get('exportPlanBtn').click();

  assert.equal(blobs.length, 1);
  const plan = JSON.parse(blobs[0].text);
  assert.equal(plan.type, 'mail-merge-plan');
  assert.equal(plan.version, 1);
  assert.equal(plan.csv, 'Name,Email\nAda,ada@example.com');
  assert.equal(plan.subject, 'Hello {{Name}}');
  assert.equal(plan.cc, 'ta@example.com');
  assert.equal(plan.body, 'Hi {{Name}},\n\nSee you there.');
  assert.deepEqual(plan.settings, { composeMode: 'gmail', batchDelay: 2000, batchSize: 50 });
  assert.equal(JSON.stringify(plan).includes('autoSend'), false);

  const link = createdElements.find(element => element.download);
  assert.match(link.download, /^mail-merge-plan-\d{4}-\d{2}-\d{2}\.json$/);
  assert.equal(link.href, 'blob:test');
  assert.match(elements.get('status').textContent, /keep it somewhere private/);
});

test('export plan asks for a CSV first', async () => {
  const { elements, blobs } = loadPopup();
  elements.get('csvText').value = '   ';

  await elements.get('exportPlanBtn').click();

  assert.equal(blobs.length, 0);
  assert.equal(elements.get('status').className, 'error');
  assert.match(elements.get('status').textContent, /Load a recipient CSV/);
});

test('import plan fills every field, turns Auto-Send off, and loads the recipients', async () => {
  const { elements, context } = loadPopup();
  elements.get('autoSend').checked = true;
  elements.get('templatePreset').value = 'courseReminder';

  await importPlan(elements, planFile('event.json', {
    type: 'mail-merge-plan',
    version: 1,
    name: 'Event invite',
    csv: ['FirstName,Email', 'Ada,ada@example.com', 'Grace,grace@example.com'],
    subject: 'Join us, {{FirstName}}',
    bcc: 'me@example.com',
    body: ['Hi {{FirstName}},', '', 'See you there!'],
    settings: { composeMode: 'gmail', batchDelay: 2000, autoSend: true }
  }));

  assert.equal(elements.get('csvText').value, 'FirstName,Email\nAda,ada@example.com\nGrace,grace@example.com');
  assert.equal(elements.get('subject').value, 'Join us, {{FirstName}}');
  assert.equal(elements.get('cc').value, '');
  assert.equal(elements.get('bcc').value, 'me@example.com');
  assert.equal(elements.get('body').value, 'Hi {{FirstName}},\n\nSee you there!');
  assert.equal(elements.get('composeMode').value, 'gmail');
  assert.equal(elements.get('batchDelay').value, '2000');
  assert.equal(elements.get('templatePreset').value, '');
  assert.equal(elements.get('autoSend').checked, false);
  assert.equal(elements.get('status').className, 'ok');
  assert.match(elements.get('status').textContent, /"Event invite" imported: 2 recipients\. Auto-Send is off/);
  const stored = (await context.chrome.storage.local.get()).outlookMailMergeState;
  assert.equal(stored.rows.length, 2);
  assert.equal(stored.autoSend, false);
});

test('import plan replaces a previously chosen CSV file when generating drafts', async () => {
  const { elements } = loadPopup();
  const csvFile = elements.get('csvFile');
  csvFile.files = [{ name: 'old.csv', async text() { return 'Name,Email\nOld,old@example.com'; } }];
  // Like a real file input, clearing the value clears the chosen file.
  Object.defineProperty(csvFile, 'value', {
    get() { return this.files.length ? 'C:\\fakepath\\old.csv' : ''; },
    set(next) { if (next === '') this.files = []; }
  });

  await importPlan(elements, planFile('plan.json', {
    type: 'mail-merge-plan',
    version: 1,
    csv: 'Name,Email\nNew,new@example.com',
    subject: 'Hi {{Name}}',
    body: 'Body'
  }));
  await elements.get('loadBtn').click();

  assert.match(elements.get('preview').innerHTML, /new@example\.com/);
  assert.doesNotMatch(elements.get('preview').innerHTML, /old@example\.com/);
});

test('import plan rejects files that are not valid plans and keeps the current setup', async () => {
  const { elements } = loadPopup();
  elements.get('subject').value = 'Keep me';

  for (const [file, message] of [
    [planFile('broken.json', '{not json'), /not valid JSON/],
    [planFile('other.json', { hello: 'world' }), /not a Mail Merge plan/],
    [planFile('future.json', { type: 'mail-merge-plan', version: 2, csv: 'Email\na@example.com' }), /Unsupported plan version/],
    [planFile('empty.json', { type: 'mail-merge-plan', version: 1, csv: '', body: 'Hi' }), /no recipient CSV/],
    [planFile('noemail.json', { type: 'mail-merge-plan', version: 1, csv: 'Name\nAda', body: 'Hi' }), /no rows with an email/]
  ]) {
    await importPlan(elements, file);
    assert.equal(elements.get('status').className, 'error');
    assert.match(elements.get('status').textContent, message);
    assert.equal(elements.get('subject').value, 'Keep me');
  }
});

test('parsePlan falls back to safe defaults for unknown settings', () => {
  const { api } = loadPopup();

  const plan = api.parsePlan(JSON.stringify({
    type: 'mail-merge-plan',
    version: 1,
    csv: 'Email\na@example.com',
    settings: { composeMode: 'yahoo', batchDelay: 10 }
  }));

  assert.equal(plan.composeMode, 'outlook');
  assert.equal(plan.batchDelay, '1000');
  assert.equal(plan.subject, '');
  assert.equal(plan.body, '');
});

// --- 0.2.0: languages, batches, test send, review notes, Excel import, help links ---

const plain = value => JSON.parse(JSON.stringify(value));
const FIVE = 'Name,Email\nA,a@example.com\nB,b@example.com\nC,c@example.com\nD,d@example.com\nE,e@example.com';

async function generate(elements, csv, { batchSize = '2', subject = 'Hi {{Name}}', body = 'Body' } = {}) {
  elements.get('csvText').value = csv;
  elements.get('batchSize').value = batchSize;
  elements.get('subject').value = subject;
  elements.get('body').value = body;
  await elements.get('loadBtn').click();
}

test('every language has the same message keys and placeholders as English', () => {
  const { api } = loadPopup();
  const placeholders = text => [...text.matchAll(/\{(\w+)\}/g)].map(match => match[1]).sort().join(',');
  const english = api.UI_STRINGS.en;

  for (const lang of ['zh', 'es']) {
    const strings = api.UI_STRINGS[lang];
    assert.deepEqual(Object.keys(strings).sort(), Object.keys(english).sort(), `${lang} keys`);
    for (const key of Object.keys(english)) {
      assert.equal(placeholders(strings[key]), placeholders(english[key]), `${lang}.${key} placeholders`);
    }
  }
});

test('switching the language relabels the buttons and is remembered', async () => {
  const { elements, context } = loadPopup();
  elements.get('uiLanguage').value = 'es';
  await elements.get('uiLanguage').dispatchEvent({ type: 'change' });

  assert.equal(elements.get('openNextBtn').textContent, 'Abrir borrador de la vista previa');
  const stored = (await context.chrome.storage.local.get()).outlookMailMergeState;
  assert.equal(stored.uiLanguage, 'es');

  const reopened = loadPopup(stored);
  await reopened.api.restoreState();
  assert.equal(reopened.elements.get('openNextBtn').textContent, 'Abrir borrador de la vista previa');
  assert.equal(reopened.elements.get('uiLanguage').value, 'es');

  reopened.api.applyLanguageSelection('zh');
  assert.equal(reopened.elements.get('openRangeBtn').textContent, '打开所选草稿');
  assert.throws(() => reopened.api.parsePlan('{bad'), /不是有效的 JSON/);
});

test('t falls back to English and fills placeholders', () => {
  const { api } = loadPopup();
  api.setUiLanguage('fr');
  assert.equal(api.t('rowLabel', { n: 3 }), 'Row 3');
  api.setUiLanguage('zh');
  assert.equal(api.t('rowLabel', { n: 3 }), '第 3 行');
  assert.equal(api.t('insertVariable', { name: 'FirstName' }), '插入 {{FirstName}}');
  assert.equal(api.t('noSuchKey'), 'noSuchKey');
});

test('batches open the next group of drafts and track progress', async () => {
  const { elements, openedTabs } = loadPopup();
  await generate(elements, FIVE);

  assert.equal(elements.get('openBatchBtn').textContent, 'Open Batch 1 of 3 (drafts 1–2)');
  assert.equal(elements.get('batchProgress').textContent, '0 of 5 opened');

  await elements.get('openBatchBtn').click();
  assert.deepEqual(openedTabs.map(tab => decodeURIComponent(/to=([^&]+)/.exec(tab.url)[1])), ['a@example.com', 'b@example.com']);
  assert.equal(elements.get('openBatchBtn').textContent, 'Open Batch 2 of 3 (drafts 3–4)');
  assert.equal(elements.get('batchProgress').textContent, '2 of 5 opened');
  assert.match(elements.get('lastBatchNote').textContent, /^Last opened: 1–2 at /);

  await elements.get('openBatchBtn').click();
  await elements.get('openBatchBtn').click();
  assert.equal(openedTabs.length, 5);
  assert.equal(elements.get('openBatchBtn').textContent, 'All batches opened');
  assert.equal(elements.get('openBatchBtn').disabled, true);
});

test('regenerating after a template edit keeps progress; editing the list starts over', async () => {
  const { elements, openedTabs } = loadPopup();
  await generate(elements, FIVE);
  await elements.get('openBatchBtn').click();

  elements.get('subject').value = 'Hello {{Name}}';
  await elements.get('subject').dispatchEvent({ type: 'input' });
  await elements.get('loadBtn').click();
  assert.equal(elements.get('openBatchBtn').textContent, 'Open Batch 2 of 3 (drafts 3–4)');
  assert.match(elements.get('status').textContent, /Progress kept: 2 of 5 already opened/);

  await elements.get('openBatchBtn').click();
  assert.match(openedTabs.at(-1).url, /subject=Hello%20D/);

  elements.get('csvText').value = FIVE.replace('A,a@', 'Ann,a@');
  await elements.get('csvText').dispatchEvent({ type: 'input' });
  await elements.get('loadBtn').click();
  assert.equal(elements.get('openBatchBtn').textContent, 'Open Batch 1 of 3 (drafts 1–2)');
});

test('Auto-Send changes the batch button to send wording', async () => {
  const { elements } = loadPopup();
  await generate(elements, FIVE);
  elements.get('autoSend').checked = true;
  await elements.get('autoSend').dispatchEvent({ type: 'change' });

  assert.equal(elements.get('openBatchBtn').textContent, 'Send Batch 1 of 3 (emails 1–2)');
  assert.equal(elements.get('sendTestBtn').textContent, 'Send Test to Me');
});

test('send test opens the previewed message for my own address only', async () => {
  const { elements, openedTabs, api } = loadPopup();
  await generate(elements, FIVE);

  await elements.get('sendTestBtn').click();
  assert.equal(openedTabs.length, 0);
  assert.match(elements.get('status').textContent, /Enter your own email address/);
  assert.equal(elements.get('statusHelp').classList.contains('hidden'), false);

  elements.get('testEmail').value = 'me@example.com';
  await elements.get('sendTestBtn').click();
  assert.equal(openedTabs.length, 1);
  assert.match(openedTabs[0].url, /to=me%40example\.com/);
  assert.match(openedTabs[0].url, /subject=%5BTEST%5D%20Hi%20A/);
  assert.doesNotMatch(openedTabs[0].url, /a%40example\.com/);
  assert.equal(api.state().currentIndex, 0);
  assert.match(elements.get('status').textContent, /Opened a test copy of draft 1, addressed to me@example\.com/);
  assert.equal(elements.get('statusHelp').classList.contains('hidden'), true);
});

const REVIEW_CSV = [
  'FirstName,Email',
  'Ada,ada.lovelace@school.edu',
  'Jhoune-Lynn,lynn.duplan@school.edu',
  'Si on,sion.chen@school.edu',
  'Grace,grace@gmail.com',
  'Alan,alan.turing@school.edu',
  ',empty@school.edu'
].join('\n');

test('reviewRecipients flags name mismatches, multi-word names, odd domains, and empty values', () => {
  const { api } = loadPopup();
  const rows = api.parseCSV(REVIEW_CSV);

  const notes = plain(api.reviewRecipients(rows, ['FirstName', 'Email'], ['Hi {{FirstName}}']))
    .map(note => `${note.index}:${note.key.split(':')[0]}`);

  assert.deepEqual(notes, ['1:mismatch', '2:multiword', '3:domain', '5:empty']);
});

test('reviewRecipients skips the name check when addresses do not contain names', () => {
  const { api } = loadPopup();
  const rows = api.parseCSV('FirstName,Email\nAda,u1001@school.edu\nGrace,u1002@school.edu\nAlan,u1003@school.edu');

  assert.equal(api.reviewRecipients(rows, ['FirstName', 'Email'], ['Hi {{FirstName}}']).length, 0);
});

test('review notes can be fixed inline or dismissed', async () => {
  const { elements, api } = loadPopup();
  await generate(elements, REVIEW_CSV, { subject: 'Hi {{FirstName}}' });
  const panel = elements.get('reviewPanel');

  assert.match(panel.innerHTML, /4 recipients worth a quick look/);
  assert.match(elements.get('status').textContent, /4 recipients are worth a quick look below/);

  await panel.dispatchEvent({ type: 'change', target: { dataset: { row: '1', field: 'FirstName' }, value: 'Lynn' } });
  assert.match(elements.get('csvText').value, /^Lynn,lynn\.duplan@school\.edu$/m);
  assert.match(elements.get('preview').innerHTML, /Hi Ada/);
  assert.match(panel.innerHTML, /3 recipients worth a quick look/);

  const domainNote = plain(api.reviewRecipients(api.state().rows, ['FirstName', 'Email'], ['Hi {{FirstName}}'])).find(note => note.email === 'grace@gmail.com');
  await panel.dispatchEvent({ type: 'click', target: { dataset: { dismiss: domainNote.key } } });
  assert.match(panel.innerHTML, /2 recipients worth a quick look/);
  assert.doesNotMatch(panel.innerHTML, /grace@gmail\.com/);
});

function xlsxFile(name, sheets) {
  const buffer = makeXlsx(sheets);
  return { name, async arrayBuffer() { return buffer; } };
}

const ROSTER = [{
  name: 'all students',
  rows: [
    ['Student', 'Last Name', 'First Name', 'Suffolk Email Address', 'Cumulative GPA'],
    ['Ada Lovelace (U1)', 'Lovelace', 'Ada', 'ada@school.edu', 3.9],
    ['No Email (U2)', 'Email', 'No', '', 2.5],
    ['Grace Hopper (U3)', 'Hopper', 'Grace', 'grace@school.edu', 3.7]
  ]
}];

test('Excel import picks the email and name columns and leaves the rest out', async () => {
  const { elements, api } = loadPopup();
  await api.handleRecipientFile(xlsxFile('roster.xlsx', ROSTER));

  const panel = elements.get('mappingPanel');
  assert.equal(panel.classList.contains('hidden'), false);
  assert.match(panel.innerHTML, /Choose columns from &quot;roster\.xlsx&quot;/);
  assert.match(panel.innerHTML, /Cumulative GPA/);
  assert.match(panel.innerHTML, /2 of 3 rows have an email address/);
  assert.deepEqual(plain([...api.state().mapping.columns]), [1, 2]);

  await panel.dispatchEvent({ type: 'click', target: { dataset: { mapping: 'apply' } } });

  assert.equal(elements.get('csvText').value, 'Last Name,First Name,Email\nLovelace,Ada,ada@school.edu\nHopper,Grace,grace@school.edu');
  assert.equal(api.state().rows.length, 2);
  assert.equal(api.state().mapping, null);
  assert.match(elements.get('status').textContent, /Loaded 2 recipients from "roster\.xlsx"/);
});

test('Excel import can include an extra column on request', async () => {
  const { elements, api } = loadPopup();
  await api.handleRecipientFile(xlsxFile('roster.xlsx', ROSTER));
  const panel = elements.get('mappingPanel');

  await panel.dispatchEvent({ type: 'change', target: { dataset: { mapping: 'column' }, value: '0', checked: true } });
  await panel.dispatchEvent({ type: 'change', target: { dataset: { mapping: 'column' }, value: '1', checked: false } });
  await panel.dispatchEvent({ type: 'click', target: { dataset: { mapping: 'apply' } } });

  assert.equal(elements.get('csvText').value.split('\n')[0], 'Student,First Name,Email');
  assert.match(elements.get('csvText').value, /^Ada Lovelace \(U1\),Ada,ada@school\.edu$/m);
});

test('Excel import explains old .xls files and unreadable files', async () => {
  const { elements, api } = loadPopup();

  await api.handleRecipientFile({ name: 'roster.xls', async arrayBuffer() { return new ArrayBuffer(8); } });
  assert.match(elements.get('status').textContent, /Old \.xls files are not supported/);

  await api.handleRecipientFile({ name: 'broken.xlsx', async arrayBuffer() { return new TextEncoder().encode('not a zip').buffer; } });
  assert.match(elements.get('status').textContent, /Could not read this Excel file/);
  assert.equal(elements.get('statusHelp').classList.contains('hidden'), false);
});

test('guessEmailColumn falls back to the column whose values look like email addresses', () => {
  const { api } = loadPopup();
  const headers = ['Name', 'Contact', 'Notes'];
  const dataRows = [['Ada', 'ada@example.com', 'x'], ['Grace', 'grace@example.com', 'y@']];

  assert.equal(api.guessEmailColumn(headers, dataRows), 1);
  assert.equal(api.guessEmailColumn(['Name', 'Correo'], [['Ada', '']]), 1);
});

test('mappingToCSV and rowsToCSV quote values that need it', () => {
  const { api } = loadPopup();

  assert.equal(
    api.mappingToCSV(['Name', 'Note', 'Mail'], [['Lovelace, Ada', 'Says "hi"', 'ada@example.com']], 2, new Set([0, 1])),
    'Name,Note,Email\n"Lovelace, Ada","Says ""hi""",ada@example.com'
  );
  assert.equal(api.rowsToCSV(['A', 'B'], [{ A: 'line\nbreak', B: 'plain' }]), 'A,B\n"line\nbreak",plain');
});

test('plans carry the batch size', () => {
  const { api } = loadPopup();
  const base = { type: 'mail-merge-plan', version: 1, csv: 'Email\na@example.com' };

  assert.equal(api.parsePlan(JSON.stringify({ ...base, settings: { batchSize: 40 } })).batchSize, 40);
  assert.equal(api.parsePlan(JSON.stringify({ ...base, settings: { batchSize: 0 } })).batchSize, 50);
  assert.equal(api.parsePlan(JSON.stringify(base)).batchSize, 50);
});
