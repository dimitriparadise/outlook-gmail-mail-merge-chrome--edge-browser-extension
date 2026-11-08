import assert from 'node:assert/strict';
import fs from 'node:fs';
import test from 'node:test';
import vm from 'node:vm';

function createElement(id, initial = {}) {
  const listeners = new Map();
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
    classList: {
      add() {},
      remove() {},
      toggle() {}
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
    setTimeout() {},
    Blob: class Blob {
      constructor(parts, options = {}) {
        this.text = parts.join('');
        this.type = options.type;
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
  const script = fs.readFileSync(new URL('../popup.js', import.meta.url), 'utf8');
  vm.runInContext(`${script}
globalThis.__popupTestApi = {
  parseCSV,
  applyTemplate,
  validateMessages: typeof validateMessages === 'function' ? validateMessages : undefined,
  findMissingTemplateVariables: typeof findMissingTemplateVariables === 'function' ? findMissingTemplateVariables : undefined,
  syncDraftActionLabels,
  restoreState,
  makeMessages,
  composeUrl,
  parsePlan: typeof parsePlan === 'function' ? parsePlan : undefined,
  buildPlan: typeof buildPlan === 'function' ? buildPlan : undefined
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
  assert.deepEqual(plan.settings, { composeMode: 'gmail', batchDelay: 2000 });
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
